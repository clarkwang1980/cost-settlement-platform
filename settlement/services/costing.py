from __future__ import annotations

import uuid
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP

from django.db import transaction
from django.db.models import Sum
from django.utils import timezone

from settlement.models import (
    AccountingPeriod,
    ExceptionItem,
    InventoryTransaction,
    InventoryValuation,
    MatchAllocation,
    MatchingMode,
    OpeningInventoryBalance,
    OpeningWip,
    OrderCostResult,
    ProductionOrder,
    ReconciliationResult,
    SettlementRun,
    WorkshopSettlement,
    WorkshopCostPool,
)
from settlement.services.rollforward import sync_carry_forward
from settlement.services.matching import (
    auto_match_all,
    is_order_material,
    is_production_receipt,
    is_purchase_receipt,
    is_sales_issue,
)
from settlement.services.sales_uninvoiced import (
    build_goods_issued_rollforward,
    build_sales_uninvoiced_rows,
    summarize_current_sales_cost,
    summarize_sales_uninvoiced,
)


ZERO = Decimal("0")
CENT = Decimal("0.01")
PRICE_PRECISION = Decimal("0.00000001")
CONVERGENCE_TOLERANCE = Decimal("0.005")
MAX_ITERATIONS = 50


def money(value) -> Decimal:
    return Decimal(value or ZERO).quantize(CENT, rounding=ROUND_HALF_UP)


def price(value) -> Decimal:
    return Decimal(value or ZERO).quantize(
        PRICE_PRECISION, rounding=ROUND_HALF_UP
    )


def clamp_rate(value) -> Decimal:
    value = Decimal(value or ZERO)
    return min(max(value, ZERO), Decimal("1"))


def is_opening_transaction(txn: InventoryTransaction) -> bool:
    return "期初" in txn.business_type


def _fixed_receipt_amount(txn: InventoryTransaction) -> bool:
    if is_opening_transaction(txn):
        return True
    if is_purchase_receipt(txn):
        return txn.amount is not None and bool(txn.price_source)
    if is_production_receipt(txn):
        return txn.amount is not None
    return txn.amount is not None and txn.price_source not in {
        "",
        "月末加权平均",
        "生产入库成本",
    }


def _valuation_pass(period: AccountingPeriod, previous_prices=None):
    previous_prices = previous_prices or {}
    grouped = defaultdict(list)
    for txn in period.inventory_transactions.select_related("material").all():
        grouped[txn.material_code].append(txn)
    automatic_openings = {
        item.material_code: item
        for item in OpeningInventoryBalance.objects.filter(period=period)
    }
    for material_code in automatic_openings:
        grouped.setdefault(material_code, [])

    summaries = {}
    prices = {}
    for material_code, txns in grouped.items():
        opening_qty = ZERO
        opening_amount = ZERO
        has_manual_opening = False
        receipt_qty = ZERO
        fixed_receipt_amount = ZERO
        floating_receipts = []
        issue_qty = ZERO

        for txn in txns:
            qty = abs(txn.quantity)
            if is_opening_transaction(txn):
                has_manual_opening = True
                opening_qty += qty
                opening_amount += Decimal(txn.amount or ZERO)
            elif txn.direction == InventoryTransaction.Direction.IN:
                receipt_qty += qty
                if _fixed_receipt_amount(txn):
                    fixed_receipt_amount += Decimal(txn.amount or ZERO)
                else:
                    floating_receipts.append(txn)
            else:
                issue_qty += qty
        automatic_opening = automatic_openings.get(material_code)
        if automatic_opening and not has_manual_opening:
            opening_qty = Decimal(automatic_opening.quantity or ZERO)
            opening_amount = Decimal(automatic_opening.amount or ZERO)

        prior_price = Decimal(previous_prices.get(material_code, ZERO))
        floating_qty = sum((abs(txn.quantity) for txn in floating_receipts), ZERO)
        numerator = opening_amount + fixed_receipt_amount + floating_qty * prior_price
        denominator = opening_qty + receipt_qty
        weighted_price = numerator / denominator if denominator else ZERO

        # Returns and other unpriced receipts are valued at the same monthly average.
        # One additional algebraic step removes their self-referential effect.
        if floating_qty and denominator > floating_qty:
            weighted_price = (
                opening_amount + fixed_receipt_amount
            ) / (denominator - floating_qty)
        weighted_price = price(weighted_price)
        prices[material_code] = weighted_price

        for txn in txns:
            qty = abs(txn.quantity)
            if txn.direction == InventoryTransaction.Direction.OUT:
                txn.unit_price = weighted_price
                txn.amount = money(qty * weighted_price)
                txn.price_source = "月末加权平均"
                txn.save(update_fields=["unit_price", "amount", "price_source"])
            elif (
                not is_opening_transaction(txn)
                and not _fixed_receipt_amount(txn)
                and not is_production_receipt(txn)
            ):
                txn.unit_price = weighted_price
                txn.amount = money(qty * weighted_price)
                txn.price_source = "月末加权平均"
                txn.save(update_fields=["unit_price", "amount", "price_source"])

        total_receipt_amount = sum(
            (
                Decimal(txn.amount or ZERO)
                for txn in txns
                if txn.direction == InventoryTransaction.Direction.IN
                and not is_opening_transaction(txn)
            ),
            ZERO,
        )
        total_issue_amount = sum(
            (
                Decimal(txn.amount or ZERO)
                for txn in txns
                if txn.direction == InventoryTransaction.Direction.OUT
            ),
            ZERO,
        )
        ending_qty = opening_qty + receipt_qty - issue_qty
        ending_amount = money(ending_qty * weighted_price)
        material_name = next(
            (txn.material_name for txn in txns if txn.material_name),
            automatic_opening.material_name if automatic_opening else "",
        )
        summaries[material_code] = {
            "material_name": material_name,
            "opening_quantity": opening_qty,
            "opening_amount": money(opening_amount),
            "receipt_quantity": receipt_qty,
            "receipt_amount": money(total_receipt_amount),
            "issue_quantity": issue_qty,
            "issue_amount": money(total_issue_amount),
            "weighted_average_price": weighted_price,
            "ending_quantity": ending_qty,
            "ending_amount": ending_amount,
        }
    return prices, summaries


def _opening_wip_map(period):
    return {item.order_no: item for item in period.opening_wip.all()}


def _order_current_material(order: ProductionOrder) -> Decimal:
    total = ZERO
    allocations = MatchAllocation.objects.filter(
        period=order.period,
        kind=MatchAllocation.Kind.ORDER_MATERIAL,
        production_order=order,
    ).select_related("inventory_transaction")
    for allocation in allocations:
        txn = allocation.inventory_transaction
        if not txn.quantity:
            continue
        allocated_amount = Decimal(txn.amount or ZERO) * (
            allocation.matched_quantity / abs(txn.quantity)
        )
        if txn.direction == InventoryTransaction.Direction.OUT:
            total += allocated_amount
        else:
            total -= allocated_amount
    return money(total)


def _material_finished_rate(order: ProductionOrder) -> Decimal:
    if order.status == ProductionOrder.Status.COMPLETED:
        return Decimal("1")
    if order.material_completion_rate is not None:
        return clamp_rate(order.material_completion_rate)
    denominator = order.finished_quantity + order.wip_quantity
    if denominator <= 0:
        return ZERO
    return clamp_rate(order.finished_quantity / denominator)


def _calculate_order_costs(period: AccountingPeriod, workshops=None):
    opening_map = _opening_wip_map(period)
    orders_by_workshop = defaultdict(list)
    current_material = {}
    orders = period.production_orders.exclude(
        match_mode=MatchingMode.IGNORED
    )
    if workshops is not None:
        orders = orders.filter(workshop__in=set(workshops))
    for order in orders:
        orders_by_workshop[order.workshop].append(order)
        current_material[order.id] = _order_current_material(order)

    results_to_replace = OrderCostResult.objects.filter(period=period)
    if workshops is not None:
        results_to_replace = results_to_replace.filter(
            order__workshop__in=set(workshops)
        )
    results_to_replace.delete()
    results = []
    for workshop, orders in orders_by_workshop.items():
        pool = period.workshop_cost_pools.filter(workshop=workshop).first()
        opening_for_order = {
            order.id: opening_map.get(order.order_no) for order in orders
        }
        material_base = {}
        for order in orders:
            opening = opening_for_order[order.id]
            material_base[order.id] = money(
                current_material[order.id] + (opening.material if opening else ZERO)
            )

        total_material_base = sum(material_base.values(), ZERO)
        cost_rates = {}
        for field in (
            "water_power",
            "steam",
            "natural_gas",
            "depreciation",
            "consumables",
        ):
            pool_amount = Decimal(getattr(pool, field, ZERO) if pool else ZERO)
            cost_rates[field] = (
                pool_amount / total_material_base if total_material_base else ZERO
            )

        total_equivalent_hours = ZERO
        hours = {}
        for order in orders:
            finished_hours = order.finished_quantity * order.standard_hours
            wip_hours = (
                order.wip_quantity
                * order.standard_hours
                * clamp_rate(order.wip_conversion_rate)
            )
            hours[order.id] = (finished_hours, wip_hours)
            total_equivalent_hours += finished_hours + wip_hours
        wage_rate = (
            Decimal(pool.wages) / total_equivalent_hours
            if pool and total_equivalent_hours
            else ZERO
        )

        for order in orders:
            opening = opening_for_order[order.id]
            direct_material = material_base[order.id]
            allocated = {}
            for field, rate in cost_rates.items():
                opening_amount = Decimal(getattr(opening, field, ZERO) if opening else ZERO)
                allocated[field] = money(direct_material * rate + opening_amount)

            finished_hours, wip_hours = hours[order.id]
            opening_wages = Decimal(opening.wages if opening else ZERO)
            hour_total = finished_hours + wip_hours
            finished_hour_share = finished_hours / hour_total if hour_total else ZERO
            current_finished_wages = finished_hours * wage_rate
            current_wip_wages = wip_hours * wage_rate
            finished_wages = money(
                current_finished_wages + opening_wages * finished_hour_share
            )
            wip_wages = money(
                current_wip_wages + opening_wages * (Decimal("1") - finished_hour_share)
            )
            wages = finished_wages + wip_wages

            material_rate = _material_finished_rate(order)
            finished_material = money(direct_material * material_rate)
            wip_material = money(direct_material - finished_material)
            overhead_total = sum(allocated.values(), ZERO)
            finished_overhead = money(overhead_total * material_rate)
            wip_overhead = money(overhead_total - finished_overhead)
            finished_cost = money(
                finished_material + finished_overhead + finished_wages
            )
            wip_cost = money(wip_material + wip_overhead + wip_wages)
            total_cost = money(finished_cost + wip_cost)
            unit_cost = (
                price(finished_cost / order.finished_quantity)
                if order.finished_quantity
                else ZERO
            )
            result = OrderCostResult.objects.create(
                period=period,
                order=order,
                direct_material=direct_material,
                water_power=allocated["water_power"],
                steam=allocated["steam"],
                natural_gas=allocated["natural_gas"],
                depreciation=allocated["depreciation"],
                consumables=allocated["consumables"],
                wages=wages,
                total_cost=total_cost,
                finished_material=finished_material,
                finished_overhead=finished_overhead,
                finished_wages=finished_wages,
                finished_cost=finished_cost,
                wip_material=wip_material,
                wip_overhead=wip_overhead,
                wip_wages=wip_wages,
                wip_cost=wip_cost,
                finished_unit_cost=unit_cost,
            )
            results.append(result)
    return results


def _update_production_receipts(period: AccountingPeriod, workshops=None):
    changes = []
    results = {
        item.order_id: item
        for item in period.order_cost_results.select_related("order").all()
    }
    allocations_by_order = defaultdict(list)
    allocations = period.match_allocations.filter(
        kind=MatchAllocation.Kind.PRODUCTION_RECEIPT
    )
    if workshops is not None:
        allocations = allocations.filter(
            production_order__workshop__in=set(workshops)
        )
    for allocation in allocations.select_related(
        "inventory_transaction", "production_order"
    ):
        allocations_by_order[allocation.production_order_id].append(allocation)

    for order_id, allocations in allocations_by_order.items():
        result = results.get(order_id)
        if not result:
            continue
        total_receipt_qty = sum((item.matched_quantity for item in allocations), ZERO)
        unit_cost = (
            price(result.finished_cost / total_receipt_qty)
            if total_receipt_qty
            else ZERO
        )
        for allocation in allocations:
            txn = allocation.inventory_transaction
            old_amount = Decimal(txn.amount or ZERO)
            allocated_amount = money(allocation.matched_quantity * unit_cost)
            allocation.unit_price = unit_cost
            allocation.amount = allocated_amount
            allocation.save(update_fields=["unit_price", "amount"])
            txn.unit_price = unit_cost
            txn.amount = money(txn.quantity * unit_cost)
            txn.price_source = "生产入库成本"
            txn.save(update_fields=["unit_price", "amount", "price_source"])
            changes.append(abs(txn.amount - old_amount))
    return max(changes, default=ZERO)


def _persist_valuations(period, summaries, iterations, converged):
    InventoryValuation.objects.filter(period=period).delete()
    InventoryValuation.objects.bulk_create(
        [
            InventoryValuation(
                period=period,
                material_code=material_code,
                material_name=data["material_name"],
                opening_quantity=data["opening_quantity"],
                opening_amount=data["opening_amount"],
                receipt_quantity=data["receipt_quantity"],
                receipt_amount=data["receipt_amount"],
                issue_quantity=data["issue_quantity"],
                issue_amount=data["issue_amount"],
                weighted_average_price=data["weighted_average_price"],
                ending_quantity=data["ending_quantity"],
                ending_amount=data["ending_amount"],
                iterations=iterations,
                converged=converged,
            )
            for material_code, data in summaries.items()
        ]
    )


def _add_exception(period, category, severity, message, reference=""):
    ExceptionItem.objects.create(
        period=period,
        category=category,
        severity=severity,
        message=message,
        reference=reference,
    )


def _build_exceptions(period, converged):
    period.exceptions.all().delete()
    if not converged:
        _add_exception(
            period,
            "循环计算",
            ExceptionItem.Severity.BLOCKING,
            f"生产入库成本在{MAX_ITERATIONS}次循环后仍未收敛。",
        )
    for txn in period.inventory_transactions.all():
        if (
            txn.match_mode != MatchingMode.IGNORED
            and is_purchase_receipt(txn)
            and (
                txn.amount is None
                or txn.unit_price is None
                or txn.unit_price == 0
            )
        ):
            _add_exception(
                period,
                "采购入库",
                ExceptionItem.Severity.BLOCKING,
                "采购入库未匹配合同价格或发票价格。",
                f"{txn.document_no}/{txn.material_code}",
            )
        if (
            txn.match_mode != MatchingMode.IGNORED
            and txn.match_status
            != InventoryTransaction.MatchStatus.NOT_REQUIRED
            and is_order_material(txn)
            and not txn.allocations.filter(
                kind=MatchAllocation.Kind.ORDER_MATERIAL
            ).exists()
        ):
            _add_exception(
                period,
                "生产领退料",
                ExceptionItem.Severity.BLOCKING,
                "生产领退料未匹配到生产订单。",
                f"{txn.document_no}/{txn.material_code}",
            )
        if (
            txn.match_mode != MatchingMode.IGNORED
            and txn.match_status
            != InventoryTransaction.MatchStatus.NOT_REQUIRED
            and is_production_receipt(txn)
            and not txn.allocations.filter(
                kind=MatchAllocation.Kind.PRODUCTION_RECEIPT
            ).exists()
        ):
            _add_exception(
                period,
                "生产入库",
                ExceptionItem.Severity.BLOCKING,
                "生产入库未匹配到生产订单。",
                f"{txn.document_no}/{txn.material_code}",
            )
        if (
            txn.direction == InventoryTransaction.Direction.OUT
            and txn.quantity
            and Decimal(txn.amount or ZERO) == 0
        ):
            _add_exception(
                period,
                "库存计价",
                ExceptionItem.Severity.BLOCKING,
                "出库数量不为零但出库金额为零。",
                f"{txn.document_no}/{txn.material_code}",
            )
    for invoice in period.purchase_invoices.filter(
        remaining_quantity__gt=0
    ).exclude(match_mode=MatchingMode.IGNORED):
        _add_exception(
            period,
            "采购发票",
            ExceptionItem.Severity.WARNING,
            f"采购发票仍有{invoice.remaining_quantity}未匹配入库。",
            invoice.invoice_no,
        )
    for invoice in period.sales_invoices.filter(
        remaining_quantity__gt=0
    ).exclude(match_mode=MatchingMode.IGNORED):
        _add_exception(
            period,
            "销售发票",
            ExceptionItem.Severity.WARNING,
            f"销售发票仍有{invoice.remaining_quantity}未匹配销售出库。",
            invoice.invoice_no,
        )
    uninvoiced_rows = build_sales_uninvoiced_rows(period)
    if uninvoiced_rows:
        uninvoiced = summarize_sales_uninvoiced(
            uninvoiced_rows,
            period,
        )
        _add_exception(
            period,
            "已出库未开票",
            ExceptionItem.Severity.INFO,
            (
                f"月末尚有{uninvoiced['issue_count']}笔销售出库未全部开票，"
                f"未开票数量{uninvoiced['outstanding_quantity']}，"
                f"发出商品成本{money(uninvoiced['outstanding_cost'])}。"
            ),
            f"{period.month:%Y-%m}",
        )
    for order in period.production_orders.exclude(
        match_mode=MatchingMode.IGNORED
    ):
        if not order.cost_result.direct_material and not period.opening_wip.filter(
            order_no=order.order_no
        ).exists():
            _add_exception(
                period,
                "订单成本",
                ExceptionItem.Severity.WARNING,
                "生产订单没有材料成本。",
                order.order_no,
            )
        receipt_qty = (
            period.match_allocations.filter(
                kind=MatchAllocation.Kind.PRODUCTION_RECEIPT,
                production_order=order,
            ).aggregate(total=Sum("matched_quantity"))["total"]
            or ZERO
        )
        if order.cost_result.finished_cost and receipt_qty == 0:
            _add_exception(
                period,
                "生产入库",
                ExceptionItem.Severity.BLOCKING,
                "订单已形成完工成本，但没有匹配的生产入库记录。",
                order.order_no,
            )
        if not period.workshop_cost_pools.filter(workshop=order.workshop).exists():
            _add_exception(
                period,
                "车间费用",
                ExceptionItem.Severity.WARNING,
                "该车间没有导入本月费用池，制造费用按零计算。",
                order.workshop,
            )
    for valuation in period.inventory_valuations.filter(ending_quantity__lt=0):
        _add_exception(
            period,
            "库存数量",
            ExceptionItem.Severity.WARNING,
            f"期末库存数量为负数：{valuation.ending_quantity}",
            valuation.material_code,
        )


def _reconciliation_status(difference, tolerance=Decimal("0.10")):
    return (
        ReconciliationResult.Status.PASS
        if abs(difference) <= tolerance
        else ReconciliationResult.Status.FAIL
    )


def _build_reconciliations(period):
    ReconciliationResult.objects.filter(period=period).delete()

    opening = (
        period.inventory_valuations.aggregate(total=Sum("opening_amount"))["total"]
        or ZERO
    )
    receipts = sum(
        (
            Decimal(txn.amount or ZERO)
            for txn in period.inventory_transactions.filter(
                direction=InventoryTransaction.Direction.IN
            )
            if not is_opening_transaction(txn)
        ),
        ZERO,
    )
    issues = sum(
        (
            Decimal(txn.amount or ZERO)
            for txn in period.inventory_transactions.filter(
                direction=InventoryTransaction.Direction.OUT
            )
        ),
        ZERO,
    )
    ending = (
        period.inventory_valuations.aggregate(total=Sum("ending_amount"))["total"]
        or ZERO
    )
    left = money(opening + receipts)
    right = money(issues + ending)
    difference = money(left - right)
    ReconciliationResult.objects.create(
        period=period,
        check_code="inventory_amount",
        check_name="库存金额：期初＋入库＝出库＋期末",
        left_amount=left,
        right_amount=right,
        difference=difference,
        status=_reconciliation_status(difference),
    )

    material_flow = ZERO
    for txn in period.inventory_transactions.all():
        if not is_order_material(txn):
            continue
        signed = Decimal(txn.amount or ZERO)
        material_flow += (
            signed
            if txn.direction == InventoryTransaction.Direction.OUT
            else -signed
        )
    opening_material = (
        period.opening_wip.aggregate(total=Sum("material"))["total"] or ZERO
    )
    order_current_material = (
        period.order_cost_results.aggregate(total=Sum("direct_material"))["total"]
        or ZERO
    ) - opening_material
    left = money(material_flow)
    right = money(order_current_material)
    difference = money(left - right)
    ReconciliationResult.objects.create(
        period=period,
        check_code="order_material",
        check_name="生产领退料金额＝订单本期材料成本",
        left_amount=left,
        right_amount=right,
        difference=difference,
        status=_reconciliation_status(difference),
    )

    opening_total = sum((item.total for item in period.opening_wip.all()), ZERO)
    pool_total = ZERO
    for pool in period.workshop_cost_pools.all():
        pool_total += (
            pool.water_power
            + pool.steam
            + pool.natural_gas
            + pool.depreciation
            + pool.consumables
            + pool.wages
        )
    finished_total = (
        period.order_cost_results.aggregate(total=Sum("finished_cost"))["total"]
        or ZERO
    )
    wip_total = (
        period.order_cost_results.aggregate(total=Sum("wip_cost"))["total"] or ZERO
    )
    left = money(opening_total + material_flow + pool_total)
    right = money(finished_total + wip_total)
    difference = money(left - right)
    ReconciliationResult.objects.create(
        period=period,
        check_code="order_cost",
        check_name="期初在制＋本期投入＝完工成本＋期末在制",
        left_amount=left,
        right_amount=right,
        difference=difference,
        status=_reconciliation_status(difference),
    )

    production_receipt_amount = sum(
        (
            Decimal(txn.amount or ZERO)
            for txn in period.inventory_transactions.all()
            if is_production_receipt(txn)
        ),
        ZERO,
    )
    left = money(finished_total)
    right = money(production_receipt_amount)
    difference = money(left - right)
    ReconciliationResult.objects.create(
        period=period,
        check_code="production_receipt",
        check_name="订单完工成本＝生产入库金额",
        left_amount=left,
        right_amount=right,
        difference=difference,
        status=_reconciliation_status(difference),
    )

    sales = summarize_current_sales_cost(period)
    right = money(
        sales["invoiced_cost"]
        + sales["goods_issued"]
        + sales["sample_cost"]
        + sales["ignored_cost"]
    )
    difference = money(sales["sales_cost"] - right)
    ReconciliationResult.objects.create(
        period=period,
        check_code="sales_cost",
        check_name="本月销售出库成本＝已开票成本＋未开票成本＋样品＋忽略",
        left_amount=sales["sales_cost"],
        right_amount=right,
        difference=difference,
        status=_reconciliation_status(difference),
        details={
            "invoiced_cost": str(sales["invoiced_cost"]),
            "goods_issued": str(sales["goods_issued"]),
            "sample_cost": str(sales["sample_cost"]),
            "ignored_cost": str(sales["ignored_cost"]),
        },
    )

    goods_issued = build_goods_issued_rollforward(period)
    left = money(
        goods_issued["opening_amount"]
        + goods_issued["current_issue_amount"]
    )
    right = money(
        goods_issued["invoice_cleared_amount"]
        + goods_issued["ending_amount"]
    )
    difference = money(left - right)
    ReconciliationResult.objects.create(
        period=period,
        check_code="goods_issued_rollforward",
        check_name="期初发出商品＋本月出库＝本月开票结转＋期末发出商品",
        left_amount=left,
        right_amount=right,
        difference=difference,
        status=_reconciliation_status(difference),
        details={
            "opening_amount": str(goods_issued["opening_amount"]),
            "current_issue_amount": str(
                goods_issued["current_issue_amount"]
            ),
            "invoice_cleared_amount": str(
                goods_issued["invoice_cleared_amount"]
            ),
            "ending_amount": str(goods_issued["ending_amount"]),
        },
    )


def ensure_workshop_settlements(period: AccountingPeriod):
    workshops = set(
        period.production_orders.exclude(workshop="")
        .exclude(match_mode=MatchingMode.IGNORED)
        .values_list("workshop", flat=True)
    )
    workshops.update(
        period.workshop_cost_pools.exclude(workshop="").values_list(
            "workshop", flat=True
        )
    )
    existing = {
        item.workshop: item
        for item in period.workshop_settlements.all()
    }
    next_sequence = max(
        (item.sequence for item in existing.values()), default=0
    )
    for workshop in sorted(workshops):
        if workshop in existing:
            continue
        next_sequence += 1
        existing[workshop] = WorkshopSettlement.objects.create(
            period=period,
            workshop=workshop,
            sequence=next_sequence,
        )
    stale = set(existing) - workshops
    period.workshop_settlements.filter(
        workshop__in=stale,
        status=WorkshopSettlement.Status.PENDING,
    ).delete()
    return period.workshop_settlements.order_by("sequence", "workshop")


def update_workshop_sequences(period: AccountingPeriod, sequence_by_id):
    if period.status == AccountingPeriod.Status.LOCKED:
        raise ValueError("该期间已锁账，不能调整车间结转顺序。")
    settlements = list(ensure_workshop_settlements(period))
    if any(
        item.status == WorkshopSettlement.Status.SETTLED
        for item in settlements
    ):
        raise ValueError("已有车间完成结转；如需改序，请先重置车间结转。")
    values = []
    for item in settlements:
        try:
            value = int(sequence_by_id[item.id])
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(f"{item.workshop}的结转顺序无效。") from exc
        if value <= 0:
            raise ValueError("结转顺序必须是大于0的整数。")
        values.append(value)
        item.sequence = value
    if len(set(values)) != len(values):
        raise ValueError("不同车间不能使用相同结转顺序。")
    WorkshopSettlement.objects.bulk_update(settlements, ["sequence"])
    return settlements


def _clear_pending_production_receipts(period):
    pending = set(
        period.workshop_settlements.exclude(
            status=WorkshopSettlement.Status.SETTLED
        ).values_list("workshop", flat=True)
    )
    if not pending:
        return
    allocations = period.match_allocations.filter(
        kind=MatchAllocation.Kind.PRODUCTION_RECEIPT,
        production_order__workshop__in=pending,
    ).select_related("inventory_transaction")
    transaction_ids = []
    for allocation in allocations:
        transaction_ids.append(allocation.inventory_transaction_id)
        allocation.unit_price = None
        allocation.amount = None
    if allocations:
        MatchAllocation.objects.bulk_update(
            list(allocations), ["unit_price", "amount"]
        )
    period.inventory_transactions.filter(id__in=transaction_ids).update(
        unit_price=None,
        amount=None,
        price_source="",
    )


def reset_workshop_settlements(period: AccountingPeriod):
    if period.status == AccountingPeriod.Status.LOCKED:
        raise ValueError("该期间已锁账，不能重置车间结转。")
    ensure_workshop_settlements(period)
    period.workshop_settlements.update(
        status=WorkshopSettlement.Status.PENDING,
        iterations=0,
        converged=False,
        settled_at=None,
        settled_by=None,
        order_cost=ZERO,
        production_receipt_amount=ZERO,
        error_message="",
    )
    OrderCostResult.objects.filter(period=period).delete()
    InventoryValuation.objects.filter(period=period).delete()
    ReconciliationResult.objects.filter(period=period).delete()
    period.exceptions.all().delete()
    _clear_pending_production_receipts(period)
    period.status = AccountingPeriod.Status.OPEN
    period.last_run_at = None
    period.save(update_fields=["status", "last_run_at"])


def _prepare_period(period):
    carry_summary = sync_carry_forward(period)
    auto_match_all(period)
    ensure_workshop_settlements(period)
    _clear_pending_production_receipts(period)
    return carry_summary


def _settle_workshop_core(period, settlement, user=None):
    previous_prices = {}
    summaries = {}
    converged = False
    iteration = 0
    workshop_filter = [settlement.workshop]
    for iteration in range(1, MAX_ITERATIONS + 1):
        previous_prices, summaries = _valuation_pass(
            period, previous_prices
        )
        _calculate_order_costs(period, workshop_filter)
        max_change = _update_production_receipts(
            period, workshop_filter
        )
        if max_change <= CONVERGENCE_TOLERANCE:
            converged = True
            break

    previous_prices, summaries = _valuation_pass(period, previous_prices)
    _calculate_order_costs(period, workshop_filter)
    _update_production_receipts(period, workshop_filter)
    previous_prices, summaries = _valuation_pass(period, previous_prices)

    settlement.status = (
        WorkshopSettlement.Status.SETTLED
        if converged
        else WorkshopSettlement.Status.FAILED
    )
    settlement.iterations = iteration
    settlement.converged = converged
    settlement.settled_at = timezone.now()
    settlement.settled_by = user
    settlement.order_cost = money(
        period.order_cost_results.filter(
            order__workshop=settlement.workshop
        ).aggregate(total=Sum("total_cost"))["total"]
        or ZERO
    )
    settlement.production_receipt_amount = money(
        period.match_allocations.filter(
            kind=MatchAllocation.Kind.PRODUCTION_RECEIPT,
            production_order__workshop=settlement.workshop,
        ).aggregate(total=Sum("amount"))["total"]
        or ZERO
    )
    settlement.error_message = (
        "" if converged else f"{MAX_ITERATIONS}次循环后仍未收敛。"
    )
    settlement.save(
        update_fields=[
            "status",
            "iterations",
            "converged",
            "settled_at",
            "settled_by",
            "order_cost",
            "production_receipt_amount",
            "error_message",
        ]
    )
    if not converged:
        raise ValueError(
            f"{settlement.workshop}生产入库成本在"
            f"{MAX_ITERATIONS}次循环后仍未收敛。"
        )
    return {
        "iterations": iteration,
        "summaries": summaries,
    }


def _finalize_period_state(period):
    _prices, summaries = _valuation_pass(period, {})
    settlements = list(period.workshop_settlements.all())
    all_settled = all(
        item.status == WorkshopSettlement.Status.SETTLED
        and item.converged
        for item in settlements
    )
    total_iterations = sum((item.iterations for item in settlements), 0)
    _persist_valuations(
        period,
        summaries,
        total_iterations,
        all_settled,
    )
    if all_settled:
        _build_exceptions(period, True)
        _build_reconciliations(period)
        period.status = AccountingPeriod.Status.CALCULATED
        period.last_run_at = timezone.now()
        period.save(update_fields=["status", "last_run_at"])
    else:
        ReconciliationResult.objects.filter(period=period).delete()
        period.exceptions.all().delete()
        period.status = AccountingPeriod.Status.OPEN
        period.last_run_at = timezone.now()
        period.save(update_fields=["status", "last_run_at"])
    return {
        "all_settled": all_settled,
        "total_iterations": total_iterations,
    }


def settle_workshop(period, settlement_id, user=None):
    run = SettlementRun.objects.create(
        period=period,
        run_id=str(uuid.uuid4()),
        status=SettlementRun.Status.RUNNING,
        summary={"mode": "workshop"},
    )
    try:
        with transaction.atomic():
            period = AccountingPeriod.objects.select_for_update().get(pk=period.pk)
            if period.status == AccountingPeriod.Status.LOCKED:
                raise ValueError("该期间已锁账，不能继续结转。")
            carry_summary = _prepare_period(period)
            settlement = period.workshop_settlements.select_for_update().get(
                id=settlement_id
            )
            if settlement.status == WorkshopSettlement.Status.SETTLED:
                raise ValueError(f"{settlement.workshop}已经完成结转。")
            predecessors = period.workshop_settlements.filter(
                sequence__lt=settlement.sequence
            ).exclude(status=WorkshopSettlement.Status.SETTLED)
            if predecessors.exists():
                names = "、".join(
                    predecessors.values_list("workshop", flat=True)
                )
                raise ValueError(f"请先完成前序车间：{names}。")
            result = _settle_workshop_core(period, settlement, user)
            final_state = _finalize_period_state(period)
            run.status = SettlementRun.Status.SUCCESS
            run.iterations = result["iterations"]
            run.converged = settlement.converged
            run.finished_at = timezone.now()
            run.summary = {
                "mode": "workshop",
                "workshop": settlement.workshop,
                "all_settled": final_state["all_settled"],
                "carry_forward": carry_summary,
            }
            run.save(
                update_fields=[
                    "status",
                    "iterations",
                    "converged",
                    "finished_at",
                    "summary",
                ]
            )
            return run
    except Exception as exc:
        run.status = SettlementRun.Status.FAILED
        run.error_message = str(exc)
        run.finished_at = timezone.now()
        run.save(update_fields=["status", "error_message", "finished_at"])
        raise


def run_settlement(period: AccountingPeriod, user=None):
    run = SettlementRun.objects.create(
        period=period,
        run_id=str(uuid.uuid4()),
        status=SettlementRun.Status.RUNNING,
    )
    try:
        with transaction.atomic():
            period = AccountingPeriod.objects.select_for_update().get(pk=period.pk)
            if period.status == AccountingPeriod.Status.LOCKED:
                raise ValueError("该期间已锁账，不能重新结算。")
            carry_summary = _prepare_period(period)
            total_iterations = 0
            for settlement in period.workshop_settlements.exclude(
                status=WorkshopSettlement.Status.SETTLED
            ).order_by("sequence", "workshop"):
                result = _settle_workshop_core(period, settlement, user)
                total_iterations += result["iterations"]
            final_state = _finalize_period_state(period)
            converged = final_state["all_settled"]
            total_iterations = final_state["total_iterations"]

            run.status = SettlementRun.Status.SUCCESS
            run.iterations = total_iterations
            run.converged = converged
            run.finished_at = timezone.now()
            run.summary = {
                "inventory_materials": period.inventory_valuations.count(),
                "orders": period.order_cost_results.count(),
                "blocking_exceptions": period.exceptions.filter(
                    severity=ExceptionItem.Severity.BLOCKING,
                    status=ExceptionItem.Status.OPEN,
                ).count(),
                "failed_reconciliations": period.reconciliations.filter(
                    status=ReconciliationResult.Status.FAIL
                ).count(),
                "carry_forward": carry_summary,
                "workshop_sequence": list(
                    period.workshop_settlements.order_by(
                        "sequence", "workshop"
                    ).values_list("workshop", flat=True)
                ),
            }
            run.save(
                update_fields=[
                    "status",
                    "iterations",
                    "converged",
                    "finished_at",
                    "summary",
                ]
            )
            period.status = AccountingPeriod.Status.CALCULATED
            period.last_run_at = timezone.now()
            period.save(update_fields=["status", "last_run_at"])
            return run
    except Exception as exc:
        run.status = SettlementRun.Status.FAILED
        run.error_message = str(exc)
        run.finished_at = timezone.now()
        run.save(update_fields=["status", "error_message", "finished_at"])
        raise


def can_lock(period: AccountingPeriod):
    if period.workshop_settlements.exclude(
        status=WorkshopSettlement.Status.SETTLED
    ).exists():
        return False, "仍有车间尚未完成独立结转。"
    latest_run = period.runs.first()
    if not latest_run or latest_run.status != SettlementRun.Status.SUCCESS:
        return False, "尚未完成成功结算。"
    if not latest_run.converged:
        return False, "循环计算尚未收敛。"
    if period.exceptions.filter(
        severity=ExceptionItem.Severity.BLOCKING,
        status=ExceptionItem.Status.OPEN,
    ).exists():
        return False, "仍有未处理的阻断异常。"
    if period.reconciliations.filter(
        status=ReconciliationResult.Status.FAIL
    ).exists():
        return False, "仍有未通过的勾稽项目。"
    return True, ""
