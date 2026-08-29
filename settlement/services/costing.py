from __future__ import annotations

import uuid
from collections import defaultdict
from datetime import timedelta
from decimal import Decimal, ROUND_HALF_UP

from django.db import transaction
from django.db.models import Sum
from django.utils import timezone

from settlement.models import (
    AccountingPeriod,
    BomLine,
    EndingInventoryBalance,
    ExceptionItem,
    InventoryTransaction,
    InventoryValuation,
    MatchAllocation,
    MatchingMode,
    Material,
    OpeningInventoryBalance,
    OpeningWip,
    ReconciliationResult,
    SettlementRun,
    WorkOrderCostResult,
    WorkshopSettlement,
    WorkshopCostPool,
)
from settlement.services.rollforward import sync_carry_forward
from settlement.services.matching import (
    auto_match_all,
    is_purchase_receipt,
    is_sales_issue,
)
from settlement.services.sales_uninvoiced import (
    build_goods_issued_rollforward,
    build_sales_uninvoiced_rows,
    summarize_current_sales_cost,
    summarize_sales_uninvoiced,
)
from settlement.services.workorder_costing import (
    FINISHED_RECEIPT_TYPES,
    calculate_work_order_costs,
    clear_baogong_receipt_prices,
    is_baogong_receipt,
    price_baogong_receipts,
)
from settlement.services.imports import WORKSHOP_ORDER, WORKSHOP_LABELS


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
    if is_baogong_receipt(txn):
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
    # 期初可能按仓库/批次多行：按物料汇总数量金额
    automatic_openings = defaultdict(
        lambda: {"quantity": ZERO, "amount": ZERO, "material_name": ""}
    )
    for item in OpeningInventoryBalance.objects.filter(period=period):
        row = automatic_openings[item.material_code]
        row["quantity"] += Decimal(item.quantity or ZERO)
        row["amount"] += Decimal(item.amount or ZERO)
        if item.material_name and not row["material_name"]:
            row["material_name"] = item.material_name
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
            # 被忽略（内部流转/无需匹配）的流水不参与月末加权平均计价
            if txn.match_mode == MatchingMode.IGNORED:
                continue
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
            opening_qty = Decimal(automatic_opening["quantity"] or ZERO)
            opening_amount = Decimal(automatic_opening["amount"] or ZERO)

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
            if txn.match_mode == MatchingMode.IGNORED:
                continue
            qty = abs(txn.quantity)
            if txn.direction == InventoryTransaction.Direction.OUT:
                txn.unit_price = weighted_price
                txn.amount = money(qty * weighted_price)
                txn.price_source = "月末加权平均"
                txn.save(update_fields=["unit_price", "amount", "price_source"])
            elif (
                not is_opening_transaction(txn)
                and not _fixed_receipt_amount(txn)
                and not is_baogong_receipt(txn)
            ):
                txn.unit_price = weighted_price
                txn.amount = money(qty * weighted_price)
                txn.price_source = "月末加权平均"
                txn.save(update_fields=["unit_price", "amount", "price_source"])

        total_receipt_amount = sum(
            (
                Decimal(txn.amount or ZERO)
                for txn in txns
                if txn.match_mode != MatchingMode.IGNORED
                and txn.direction == InventoryTransaction.Direction.IN
                and not is_opening_transaction(txn)
            ),
            ZERO,
        )
        total_issue_amount = sum(
            (
                Decimal(txn.amount or ZERO)
                for txn in txns
                if txn.match_mode != MatchingMode.IGNORED
                and txn.direction == InventoryTransaction.Direction.OUT
            ),
            ZERO,
        )
        ending_qty = opening_qty + receipt_qty - issue_qty
        ending_amount = money(ending_qty * weighted_price)
        material_name = next(
            (txn.material_name for txn in txns if txn.material_name),
            automatic_opening["material_name"] if automatic_opening else "",
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
    # 期末存货以用户上传报表为准（数量不变），单价金额由系统按月末加权价回填
    _fill_ending_prices(period, summaries)
    # 系统自动期末：对"无人工期末"的物料按 期初+入库−出库(FIFO批次) 生成
    # （浆粕成品等未上传期末的物料由此自动算出期末数；有人工期末的物料以人工为准）
    _persist_ending_by_batch(period, summaries)


def _fill_ending_prices(period, summaries):
    """用户上传的期末存货：数量保持不变，回填单价金额 = 月末加权价 × 数量。"""
    for ending in period.ending_inventory_balances.filter(
        is_system_generated=False
    ):
        if ending.material_code not in summaries:
            continue
        price = summaries[ending.material_code]["weighted_average_price"]
        ending.unit_price = price
        ending.amount = money(Decimal(ending.quantity or ZERO) * price)
        ending.save(update_fields=["unit_price", "amount"])


def _persist_ending_by_batch(period, summaries):
    """按 物料+仓库+批次 生成期末存货（批次级）。

    规则：
    1. 期初各仓库/批次行 = 期初库存，顺序按期初行（即批次入库先后）
    2. 本月入库流水：能匹配到同仓库同批次的，累加到该批次；
       匹配不到或没有批次号的，新建一行（新批次/无批次）
    3. 本月出库流水：按先进先出（FIFO）从最早批次扣减
    4. 期末行 = 剩余数量非零的批次行；金额 = 数量 × 物料月末加权价
    """
    EndingInventoryBalance.objects.filter(
        period=period, is_system_generated=True
    ).delete()
    manual_codes = set(
        EndingInventoryBalance.objects.filter(
            period=period, is_system_generated=False
        ).values_list("material_code", flat=True)
    )
    month_end = (
        (period.month.replace(day=28) + timedelta(days=4)).replace(day=1)
        - timedelta(days=1)
    )

    # 期初批次行（按 id 顺序 = 上传顺序 = 先进先出顺序）
    opening_rows = {}
    for opening in OpeningInventoryBalance.objects.filter(period=period).order_by(
        "id"
    ):
        code = opening.material_code
        key = (opening.warehouse or "", opening.batch_no or "")
        opening_rows.setdefault(code, []).append(
            {
                "warehouse": opening.warehouse,
                "batch_no": opening.batch_no,
                "category_name": opening.category_name or "",
                "uom": opening.uom or "",
                "qty": Decimal(opening.quantity or ZERO),
            }
        )

    # 本月入库/出库流水（排除被忽略的内部流转）
    txns = list(
        period.inventory_transactions.exclude(
            match_mode=MatchingMode.IGNORED
        ).order_by("txn_date", "id")
    )

    # 期初类型流水（业务类型含"期初"的入库）也作为初始批次行
    for txn in txns:
        if (
            not is_opening_transaction(txn)
            or txn.direction != InventoryTransaction.Direction.IN
        ):
            continue
        code = txn.material_code
        key = (txn.warehouse or "", txn.batch_no or "")
        rows = opening_rows.setdefault(code, [])
        matched = None
        for row in rows:
            if (
                (row["warehouse"] or "") == key[0]
                and (row["batch_no"] or "") == key[1]
            ):
                matched = row
                break
        if matched:
            matched["qty"] += Decimal(txn.quantity or ZERO)
        else:
            rows.append(
                {
                    "warehouse": txn.warehouse or None,
                    "batch_no": txn.batch_no or None,
                    "category_name": txn.category_name or "",
                    "uom": txn.uom or "",
                    "qty": Decimal(txn.quantity or ZERO),
                }
            )

    ending_by_code = {}
    codes = set(opening_rows) | {
        t.material_code for t in txns
    }
    for code in sorted(codes):
        if code in manual_codes:
            continue
        batches = list(opening_rows.get(code, []))
        if not batches:
            # 无期初的物料：本月入库形成新批次行
            batches = [
                {
                    "warehouse": None,
                    "batch_no": None,
                    "category_name": "",
                    "uom": "",
                    "qty": ZERO,
                }
            ]
        # 仓库/批次索引，用于入库匹配
        index = {}
        for row in batches:
            index.setdefault(
                (row["warehouse"] or "", row["batch_no"] or ""), []
            ).append(row)

        for txn in txns:
            if txn.material_code != code:
                continue
            qty = abs(Decimal(txn.quantity or ZERO))
            if is_opening_transaction(txn):
                continue
            if txn.direction == InventoryTransaction.Direction.IN:
                key = (txn.warehouse or "", txn.batch_no or "")
                pool = index.get(key)
                if pool:
                    row = pool[-1]
                else:
                    row = {
                        "warehouse": txn.warehouse or None,
                        "batch_no": txn.batch_no or None,
                        "category_name": txn.category_name or "",
                        "uom": txn.uom or "",
                        "qty": ZERO,
                    }
                    batches.append(row)
                    index.setdefault(key, []).append(row)
                row["qty"] += qty
            else:
                # 出库：FIFO 从最早批次扣减
                remaining = qty
                for row in batches:
                    if remaining <= 0:
                        break
                    take = min(row["qty"], remaining)
                    row["qty"] -= take
                    remaining -= take
                if remaining > 0:
                    # 出库超出全部批次：记为最早批次负数（负库存）
                    if batches:
                        batches[0]["qty"] -= remaining

        weighted_price = ZERO
        if code in summaries:
            weighted_price = summaries[code]["weighted_average_price"]
        for row in batches:
            if row["qty"] == ZERO:
                continue
            ending_by_code.setdefault(code, []).append(
                {
                    "material_name": (
                        summaries[code]["material_name"]
                        if code in summaries
                        else ""
                    ),
                    "category_name": row["category_name"],
                    "warehouse": row["warehouse"],
                    "batch_no": row["batch_no"],
                    "uom": row["uom"],
                    "qty": row["qty"],
                    "amount": money(row["qty"] * weighted_price),
                    "unit_price": weighted_price,
                }
            )

    rows_to_create = []
    for code, rows in ending_by_code.items():
        for row in rows:
            rows_to_create.append(
                EndingInventoryBalance(
                    period=period,
                    material_code=code,
                    material_name=row["material_name"],
                    ending_date=month_end,
                    category_name=row["category_name"],
                    warehouse=row["warehouse"],
                    batch_no=row["batch_no"],
                    uom=row["uom"],
                    quantity=row["qty"],
                    amount=row["amount"],
                    unit_price=row["unit_price"],
                    is_system_generated=True,
                )
            )
    EndingInventoryBalance.objects.bulk_create(rows_to_create)


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
            and txn.direction == InventoryTransaction.Direction.OUT
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
    for result in period.work_order_cost_results.select_related("work_order").all():
        order = result.work_order
        if not result.direct_material and not period.opening_wip.filter(
            order_no=order.order_no
        ).exists():
            _add_exception(
                period,
                "工单成本",
                ExceptionItem.Severity.WARNING,
                "生产工单没有材料成本。",
                order.order_no,
            )
        if (
            result.finished_quantity
            and not result.wip_quantity
            and not order.plan_quantity
        ):
            _add_exception(
                period,
                "工单成本",
                ExceptionItem.Severity.WARNING,
                "工单未填计划数量，本期按全部完工处理，在产品为0。",
                order.order_no,
            )
        if not period.workshop_cost_pools.filter(
            workshop=order.production_dept_code
        ).exists():
            _add_exception(
                period,
                "车间费用",
                ExceptionItem.Severity.WARNING,
                "该车间没有导入本月费用池，制造费用按零计算。",
                order.production_dept_code,
            )
    # 投料物料无月末加权价
    priced_codes = set(
        period.inventory_valuations.values_list("material_code", flat=True)
    )
    for code in (
        period.material_issues.exclude(material_code="")
        .values_list("material_code", flat=True)
        .distinct()
    ):
        valuation = (
            period.inventory_valuations.filter(material_code=code).first()
        )
        if code not in priced_codes or (
            valuation is not None and valuation.weighted_average_price == 0
        ):
            _add_exception(
                period,
                "投料计价",
                ExceptionItem.Severity.WARNING,
                "投料物料没有月末加权平均价，材料成本按0计算。",
                code,
            )
    # 报工入库未按完工成本计价
    for txn in period.inventory_transactions.filter(
        business_type__in=FINISHED_RECEIPT_TYPES
    ).exclude(match_mode=MatchingMode.IGNORED):
        if txn.amount is None:
            _add_exception(
                period,
                "完工入库",
                ExceptionItem.Severity.WARNING,
                "报工入库未按工单完工成本计价（工单无完工成本）。",
                f"{txn.document_no}/{txn.material_code}",
            )
    # 报工记录数量与报工入库数量差异
    report_qty = defaultdict(lambda: ZERO)
    for row in (
        period.work_reports.exclude(work_order_no="")
        .values("work_order_no")
        .annotate(total=Sum("quantity"))
    ):
        report_qty[row["work_order_no"]] = Decimal(row["total"] or ZERO)
    receipt_qty = defaultdict(lambda: ZERO)
    for txn in period.inventory_transactions.filter(business_type__in=FINISHED_RECEIPT_TYPES):
        receipt_qty[txn.document_no] += Decimal(txn.quantity or ZERO)
    for order_no, reported in report_qty.items():
        received = receipt_qty.get(order_no, ZERO)
        if abs(reported - received) > Decimal("0.001"):
            _add_exception(
                period,
                "完工入库",
                ExceptionItem.Severity.INFO,
                f"报工数量{reported}与报工入库数量{received}不一致。",
                order_no,
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
    # 期末金额：用户上传了期末存货则以上传报表为准，否则用系统计算值
    uploaded_ending = period.ending_inventory_balances.filter(
        is_system_generated=False
    )
    if uploaded_ending.exists():
        ending = (
            uploaded_ending.aggregate(total=Sum("amount"))["total"] or ZERO
        )
    else:
        ending = (
            period.inventory_valuations.aggregate(
                total=Sum("ending_amount")
            )["total"]
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
        if (txn.business_type or "").strip() not in {
            "投料",
            "投料回撤",
            "投料撤回",
        }:
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
        period.work_order_cost_results.aggregate(total=Sum("direct_material"))[
            "total"
        ]
        or ZERO
    ) - opening_material
    left = money(material_flow)
    right = money(order_current_material)
    difference = money(left - right)
    ReconciliationResult.objects.create(
        period=period,
        check_code="work_order_material",
        check_name="投料出库金额＝工单本期材料成本",
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
        period.work_order_cost_results.aggregate(total=Sum("finished_cost"))[
            "total"
        ]
        or ZERO
    )
    wip_total = (
        period.work_order_cost_results.aggregate(total=Sum("wip_cost"))["total"]
        or ZERO
    )
    left = money(opening_total + material_flow + pool_total)
    right = money(finished_total + wip_total)
    difference = money(left - right)
    ReconciliationResult.objects.create(
        period=period,
        check_code="work_order_cost",
        check_name="期初在制＋本期投入＝完工成本＋期末在产品",
        left_amount=left,
        right_amount=right,
        difference=difference,
        status=_reconciliation_status(difference),
    )

    baogong_receipt_amount = sum(
        (
            Decimal(txn.amount or ZERO)
            for txn in period.inventory_transactions.all()
            if (txn.business_type or "").strip() in FINISHED_RECEIPT_TYPES
        ),
        ZERO,
    )
    left = money(finished_total)
    right = money(baogong_receipt_amount)
    difference = money(left - right)
    ReconciliationResult.objects.create(
        period=period,
        check_code="work_order_receipt",
        check_name="工单完工成本＝报工入库金额",
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
        period.work_orders.exclude(production_dept_code="")
        .values_list("production_dept_code", flat=True)
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
    # 默认结转顺序：WORKSHOP_ORDER（染色车间→纺纱→包覆→针织→机织），
    # 不在列表中的车间按字母序追加到末尾
    ordered = [code for code in WORKSHOP_ORDER if code in workshops]
    ordered += sorted(workshops - set(WORKSHOP_ORDER))
    next_sequence = max(
        (item.sequence for item in existing.values()), default=0
    )
    for workshop in ordered:
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
            raise ValueError(f"{WORKSHOP_LABELS.get(item.workshop, item.workshop)}的结转顺序无效。") from exc
        if value <= 0:
            raise ValueError("结转顺序必须是大于0的整数。")
        values.append(value)
        item.sequence = value
    if len(set(values)) != len(values):
        raise ValueError("不同车间不能使用相同结转顺序。")
    WorkshopSettlement.objects.bulk_update(settlements, ["sequence"])
    return settlements


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
    WorkOrderCostResult.objects.filter(period=period).delete()
    InventoryValuation.objects.filter(period=period).delete()
    EndingInventoryBalance.objects.filter(
        period=period, is_system_generated=True
    ).delete()
    ReconciliationResult.objects.filter(period=period).delete()
    period.exceptions.all().delete()
    clear_baogong_receipt_prices(period)
    period.status = AccountingPeriod.Status.OPEN
    period.last_run_at = None
    period.save(update_fields=["status", "last_run_at"])


def _prepare_period(period):
    carry_summary = sync_carry_forward(period)
    _normalize_priced_production_flows(period)
    auto_match_all(period)
    ensure_workshop_settlements(period)
    clear_baogong_receipt_prices(period)
    # 清空后用现有工单成本结果立即回填全部报工入库，
    # 保证已结转车间的报工入库始终带完工成本，不被后续单车间结转清空
    price_baogong_receipts(period, None)
    _fill_opening_prices(period)
    return carry_summary


def _fill_opening_prices(period):
    """期初存货只填数量时，用当月采购合同不含税价回填单价、金额。

    金额 = 数量 × 合同不含税价；没有合同价的期初行保持 0（由异常清单提示）。
    """
    contract_prices = {}
    for contract in period.purchase_contracts.order_by(
        "contract_date", "id"
    ):
        contract_prices.setdefault(
            contract.material_code, contract.unit_price_excl_tax
        )
    filled = 0
    for opening in period.opening_inventory_balances.filter(
        is_system_generated=False
    ):
        if (opening.unit_price or ZERO) > 0 or (opening.amount or ZERO) > 0:
            continue
        price = contract_prices.get(opening.material_code)
        if price is None:
            continue
        opening.unit_price = price
        opening.amount = money(Decimal(opening.quantity or ZERO) * price)
        opening.save(update_fields=["unit_price", "amount"])
        filled += 1
    return filled


def _normalize_priced_production_flows(period: AccountingPeriod):
    """历史导入的 投料/报工 流水曾按内部流转"忽略"处理；
    新口径下它们要参与存货计价，结算前自动修正匹配标志与方向。"""
    from settlement.models import InventoryTransaction

    txns = period.inventory_transactions.filter(
        business_type__in=["投料", "投料回撤", "投料撤回", "报工"],
        match_mode=MatchingMode.IGNORED,
    )
    updated = 0
    for txn in txns:
        txn.match_mode = MatchingMode.AUTO
        txn.match_status = InventoryTransaction.MatchStatus.UNMATCHED
        if txn.business_type in {"投料回撤", "投料撤回"}:
            txn.direction = InventoryTransaction.Direction.IN
        txn.save(
            update_fields=["match_mode", "match_status", "direction"]
            if txn.business_type in {"投料回撤", "投料撤回"}
            else ["match_mode", "match_status"]
        )
        updated += 1
    return updated


def _settle_workshop_core(period, settlement, user=None):
    # 粗纱链处理已改为计算层（build_roving_plan 内置于 calculate_work_order_costs），
    # 不再自动生成任何库存流水/投料记录，原始数据保持干净。
    # 浆粕车间：特殊车间（无工单/无投料/无报工），成本归集 = 原料领料(用户上传浆粕流水) + 车间费用，
    # 按产量分摊到 4 种成品，回填生产入库单价（见 pulp_costing）
    if settlement.workshop == "浆粕车间":
        from settlement.services.pulp_costing import (
            collect_pulp_costs,
            price_pulp_receipts,
        )
        costs = collect_pulp_costs(period)
        max_change = price_pulp_receipts(period)
        converged = max_change <= CONVERGENCE_TOLERANCE
        iteration = 1
        for _ in range(2, 4):
            max_change = price_pulp_receipts(period)
            iteration += 1
            converged = max_change <= CONVERGENCE_TOLERANCE
            if max_change <= CONVERGENCE_TOLERANCE:
                break
        settlement.status = (
            WorkshopSettlement.Status.SETTLED
            if converged
            else WorkshopSettlement.Status.FAILED
        )
        settlement.iterations = iteration
        settlement.converged = converged
        settlement.settled_at = timezone.now()
        settlement.settled_by = user
        settlement.order_cost = money(sum(
            (Decimal(c["total"]) for c in costs["products"].values()), ZERO
        ))
        settlement.production_receipt_amount = money(sum(
            (
                Decimal(t.amount or ZERO)
                for t in period.inventory_transactions.filter(business_type="生产入库")
                if (t.material_name or "") and (
                    "浆粕" in t.material_name or "微纤" in t.material_name
                    or str(t.material_code or "").upper().startswith("JP-")
                )
            ),
            ZERO,
        ))
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
        return {
            "iterations": iteration,
            "converged": converged,
            "workshop": settlement.workshop,
            "summary": {"workshop": settlement.workshop, "all_settled": False},
        }

    previous_prices = {}
    summaries = {}
    converged = False
    iteration = 0
    workshop_filter = [settlement.workshop]
    for iteration in range(1, MAX_ITERATIONS + 1):
        previous_prices, summaries = _valuation_pass(
            period, previous_prices
        )
        calculate_work_order_costs(period, workshop_filter, previous_prices)
        max_change = price_baogong_receipts(
            period, workshop_filter
        )
        if max_change <= CONVERGENCE_TOLERANCE:
            converged = True
            break

    previous_prices, summaries = _valuation_pass(period, previous_prices)
    calculate_work_order_costs(period, workshop_filter, previous_prices)
    price_baogong_receipts(period, workshop_filter)
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
        period.work_order_cost_results.filter(
            work_order__production_dept_code=settlement.workshop
        ).aggregate(total=Sum("total_cost"))["total"]
        or ZERO
    )
    settlement.production_receipt_amount = money(
        period.inventory_transactions.filter(
            business_type__in=FINISHED_RECEIPT_TYPES,
            document_no__in=[
                item.work_order.order_no
                for item in period.work_order_cost_results.select_related(
                    "work_order"
                ).filter(
                    work_order__production_dept_code=settlement.workshop
                )
            ],
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
            f"{WORKSHOP_LABELS.get(settlement.workshop, settlement.workshop)}完工入库成本在"
            f"{MAX_ITERATIONS}次循环后仍未收敛。"
        )
    return {
        "iterations": iteration,
        "summaries": summaries,
    }


def _global_cost_finalize(period: AccountingPeriod):
    """全部车间结转后，用最终月末加权价重算全部工单成本与报工入库，
    确保工单成本结果与库存流水按同一价格口径，勾稽才能闭合。

    返回 (最终价格, 估值汇总)。
    """
    previous_prices = {}
    summaries = {}
    converged = False
    for _ in range(MAX_ITERATIONS):
        previous_prices, summaries = _valuation_pass(
            period, previous_prices
        )
        calculate_work_order_costs(period, None, previous_prices)
        max_change = price_baogong_receipts(period, None)
        if max_change <= CONVERGENCE_TOLERANCE:
            converged = True
            break
    # 终值对齐：以最终价格重算工单成本（与最终流水金额同价）
    previous_prices, summaries = _valuation_pass(
        period, previous_prices
    )
    calculate_work_order_costs(period, None, previous_prices)
    price_baogong_receipts(period, None)
    previous_prices, summaries = _valuation_pass(
        period, previous_prices
    )
    calculate_work_order_costs(period, None, previous_prices)

    # 同步车间结转汇总数字
    for settlement in period.workshop_settlements.all():
        if settlement.workshop == "浆粕车间":
            continue  # 浆粕无工单：成本已在浆粕分支回填，不覆盖
        settlement.order_cost = money(
            period.work_order_cost_results.filter(
                work_order__production_dept_code=settlement.workshop
            ).aggregate(total=Sum("total_cost"))["total"]
            or ZERO
        )
        settlement.production_receipt_amount = money(
            period.inventory_transactions.filter(
                business_type__in=FINISHED_RECEIPT_TYPES,
                document_no__in=[
                    item.work_order.order_no
                    for item in period.work_order_cost_results.select_related(
                        "work_order"
                    ).filter(
                        work_order__production_dept_code=settlement.workshop
                    )
                ],
            ).aggregate(total=Sum("amount"))["total"]
            or ZERO
        )
        settlement.save(
            update_fields=["order_cost", "production_receipt_amount"]
        )
    return previous_prices, summaries, converged


def _finalize_period_state(period):
    settlements = list(period.workshop_settlements.all())
    all_settled = all(
        item.status == WorkshopSettlement.Status.SETTLED
        and item.converged
        for item in settlements
    )
    total_iterations = sum((item.iterations for item in settlements), 0)
    if all_settled:
        _prices, summaries, _converged = _global_cost_finalize(period)
    else:
        _prices, summaries = _valuation_pass(period, {})
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
                raise ValueError(f"{WORKSHOP_LABELS.get(settlement.workshop, settlement.workshop)}已经完成结转。")
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


def _fill_bom_prices(period: AccountingPeriod):
    """回填BOM单价：子项单价 = 子项物料本月月末加权平均价；
    标准成品单价 = 按工单汇总 Σ(子项用量 × 子项单价)。"""
    lines = list(period.bom_lines.all())
    if not lines:
        return
    prices = {
        row["material_code"]: row["weighted_average_price"]
        for row in period.inventory_valuations.values(
            "material_code", "weighted_average_price"
        )
    }
    standard_by_order = defaultdict(Decimal)
    for line in lines:
        unit_price = Decimal(prices.get(line.sub_item_code, ZERO) or ZERO)
        line.unit_price = unit_price
        standard_by_order[line.work_order_no] += line.qty * unit_price
    for line in lines:
        line.standard_price = money(standard_by_order[line.work_order_no])
    BomLine.objects.bulk_update(lines, ["unit_price", "standard_price"])


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

            # 回填BOM单价：子项单价=子项物料月末加权价；标准成品单价=Σ(子项用量×子项单价)
            _fill_bom_prices(period)

            run.status = SettlementRun.Status.SUCCESS
            run.iterations = total_iterations
            run.converged = converged
            run.finished_at = timezone.now()
            run.summary = {
                "inventory_materials": period.inventory_valuations.count(),
                "orders": period.work_order_cost_results.count(),
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
