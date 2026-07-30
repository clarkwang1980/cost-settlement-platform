from __future__ import annotations

from datetime import timedelta
from decimal import Decimal, ROUND_HALF_UP

from django.db import transaction

from settlement.models import (
    AccountingPeriod,
    InventoryValuation,
    MatchAllocation,
    MatchingMode,
    OpeningInventoryBalance,
    OpeningWip,
    OrderCostResult,
    ProductionOrder,
    PurchaseContractLine,
    ReconciliationResult,
    SettlementRun,
    WorkshopSettlement,
)


ZERO = Decimal("0")
CENT = Decimal("0.01")
OVERHEAD_FIELDS = (
    "water_power",
    "steam",
    "natural_gas",
    "depreciation",
    "consumables",
)


def money(value) -> Decimal:
    return Decimal(value or ZERO).quantize(CENT, rounding=ROUND_HALF_UP)


def previous_month_start(month):
    return (month.replace(day=1) - timedelta(days=1)).replace(day=1)


def previous_period_for(period: AccountingPeriod):
    previous_month = previous_month_start(period.month)
    previous = AccountingPeriod.objects.filter(month=previous_month).first()
    if previous:
        return previous
    if AccountingPeriod.objects.filter(month__lt=period.month).exists():
        raise ValueError(
            f"{period}之前已有核算期间，但缺少{previous_month:%Y年%m月}。"
            "请补建缺失月份，避免跨期断档。"
        )
    return None


def carry_forward_context(period: AccountingPeriod):
    try:
        previous = previous_period_for(period)
        error = ""
    except ValueError as exc:
        previous = None
        error = str(exc)
    return {
        "previous": previous,
        "error": error,
        "is_first_period": previous is None and not error,
        "ready": bool(
            previous and previous.status == AccountingPeriod.Status.LOCKED
        ),
        "opening_inventory_count": period.opening_inventory_balances.count(),
        "opening_wip_count": period.opening_wip.filter(
            is_system_generated=True
        ).count(),
        "carried_order_count": period.production_orders.filter(
            is_system_generated=True
        ).count(),
        "carried_contract_count": period.purchase_contracts.filter(
            is_system_generated=True
        ).count(),
    }


def _manual_opening_inventory_codes(period):
    return set(
        period.inventory_transactions.filter(
            business_type__contains="期初"
        ).values_list("material_code", flat=True)
    )


def _carry_inventory(period, previous):
    valuations = list(previous.inventory_valuations.all())
    active_codes = set()
    for valuation in valuations:
        if valuation.ending_quantity == ZERO and valuation.ending_amount == ZERO:
            continue
        active_codes.add(valuation.material_code)
        OpeningInventoryBalance.objects.update_or_create(
            period=period,
            material_code=valuation.material_code,
            defaults={
                "source_period": previous,
                "material_name": valuation.material_name,
                "quantity": valuation.ending_quantity,
                "amount": valuation.ending_amount,
                "unit_price": valuation.weighted_average_price,
                "is_system_generated": True,
            },
        )
    period.opening_inventory_balances.exclude(material_code__in=active_codes).delete()
    return {
        "count": len(active_codes),
        "manual_overrides": len(
            _manual_opening_inventory_codes(period) & active_codes
        ),
    }


def _wip_components(result):
    overhead_total = sum(
        (Decimal(getattr(result, field) or ZERO) for field in OVERHEAD_FIELDS),
        ZERO,
    )
    wip_overhead = Decimal(result.wip_overhead or ZERO)
    ratio = wip_overhead / overhead_total if overhead_total else ZERO
    components = {
        field: money(Decimal(getattr(result, field) or ZERO) * ratio)
        for field in OVERHEAD_FIELDS
    }
    residual = money(wip_overhead - sum(components.values(), ZERO))
    components["consumables"] = money(components["consumables"] + residual)
    return components


def _carry_wip_and_orders(period, previous):
    prior_results = [
        item
        for item in previous.order_cost_results.select_related("order").all()
        if Decimal(item.wip_cost or ZERO) != ZERO
    ]
    manual_wip_orders = set(
        period.opening_wip.filter(
            is_system_generated=False
        ).values_list("order_no", flat=True)
    )
    active_wip_orders = set()
    for result in prior_results:
        order_no = result.order.order_no
        active_wip_orders.add(order_no)
        if order_no in manual_wip_orders:
            continue
        components = _wip_components(result)
        OpeningWip.objects.update_or_create(
            period=period,
            order_no=order_no,
            defaults={
                "batch": None,
                "source_period": previous,
                "is_system_generated": True,
                "material": result.wip_material,
                "water_power": components["water_power"],
                "steam": components["steam"],
                "natural_gas": components["natural_gas"],
                "depreciation": components["depreciation"],
                "consumables": components["consumables"],
                "wages": result.wip_wages,
            },
        )
    period.opening_wip.filter(is_system_generated=True).exclude(
        order_no__in=active_wip_orders
    ).delete()

    prior_orders = {}
    for result in prior_results:
        prior_orders[result.order.order_no] = result.order
    for order in previous.production_orders.filter(
        status=ProductionOrder.Status.OPEN
    ).exclude(match_mode=MatchingMode.IGNORED):
        if order.wip_quantity != ZERO:
            prior_orders.setdefault(order.order_no, order)

    manual_orders = set(
        period.production_orders.filter(
            is_system_generated=False
        ).values_list("order_no", flat=True)
    )
    active_order_nos = set(prior_orders)
    for order_no, prior_order in prior_orders.items():
        if order_no in manual_orders:
            continue
        ProductionOrder.objects.update_or_create(
            period=period,
            order_no=order_no,
            defaults={
                "batch": None,
                "carried_from": prior_order,
                "is_system_generated": True,
                "workshop": prior_order.workshop,
                "product_code": prior_order.product_code,
                "product_name": prior_order.product_name,
                "plan_quantity": prior_order.wip_quantity,
                "finished_quantity": ZERO,
                "wip_quantity": prior_order.wip_quantity,
                "status": ProductionOrder.Status.OPEN,
                "standard_hours": prior_order.standard_hours,
                "material_completion_rate": ZERO,
                "wip_conversion_rate": prior_order.wip_conversion_rate,
                "production_receipt_document": "",
                "match_mode": MatchingMode.AUTO,
                "ignore_reason": "",
            },
        )
    period.production_orders.filter(is_system_generated=True).exclude(
        order_no__in=active_order_nos
    ).delete()
    return {
        "wip_count": len(active_wip_orders),
        "wip_manual_overrides": len(manual_wip_orders & active_wip_orders),
        "order_count": len(active_order_nos),
        "order_manual_overrides": len(manual_orders & active_order_nos),
    }


def _carry_contracts(period, previous):
    prior_contracts = list(
        previous.purchase_contracts.filter(remaining_quantity__gt=ZERO)
    )
    active_source_ids = set()
    for source in prior_contracts:
        active_source_ids.add(source.id)
        quantity = source.remaining_quantity
        PurchaseContractLine.objects.update_or_create(
            period=period,
            carried_from=source,
            defaults={
                "batch": None,
                "is_system_generated": True,
                "contract_no": source.contract_no,
                "contract_date": source.contract_date,
                "supplier_code": source.supplier_code,
                "supplier_name": source.supplier_name,
                "material_code": source.material_code,
                "material_name": source.material_name,
                "quantity": quantity,
                "remaining_quantity": quantity,
                "unit_price_excl_tax": source.unit_price_excl_tax,
            },
        )
    period.purchase_contracts.filter(is_system_generated=True).exclude(
        carried_from_id__in=active_source_ids
    ).delete()
    return len(active_source_ids)


@transaction.atomic
def sync_carry_forward(period: AccountingPeriod):
    previous = previous_period_for(period)
    if previous is None:
        return {
            "source_period": None,
            "is_first_period": True,
            "inventory_count": 0,
            "wip_count": 0,
            "order_count": 0,
            "contract_count": 0,
        }
    if previous.status != AccountingPeriod.Status.LOCKED:
        raise ValueError(
            f"上期{previous}尚未锁账。请先完成上期勾稽并锁账，"
            "再执行本期结算。"
        )
    if not previous.inventory_valuations.exists():
        raise ValueError(f"上期{previous}没有库存结算结果，不能自动结转。")

    inventory = _carry_inventory(period, previous)
    wip = _carry_wip_and_orders(period, previous)
    contract_count = _carry_contracts(period, previous)
    return {
        "source_period": previous.month.isoformat(),
        "is_first_period": False,
        "inventory_count": inventory["count"],
        "inventory_manual_overrides": inventory["manual_overrides"],
        "wip_count": wip["wip_count"],
        "wip_manual_overrides": wip["wip_manual_overrides"],
        "order_count": wip["order_count"],
        "order_manual_overrides": wip["order_manual_overrides"],
        "contract_count": contract_count,
    }


@transaction.atomic
def invalidate_downstream_periods(source_period: AccountingPeriod):
    later_periods = list(
        AccountingPeriod.objects.filter(
            month__gt=source_period.month
        ).order_by("-month")
    )
    locked = [
        item
        for item in later_periods
        if item.status == AccountingPeriod.Status.LOCKED
    ]
    if locked:
        names = "、".join(str(item) for item in reversed(locked))
        raise ValueError(
            f"后续期间{names}已经锁账。请从最后一个月份开始倒序重开。"
        )

    for period in later_periods:
        MatchAllocation.objects.filter(period=period).delete()
        InventoryValuation.objects.filter(period=period).delete()
        OrderCostResult.objects.filter(period=period).delete()
        ReconciliationResult.objects.filter(period=period).delete()
        SettlementRun.objects.filter(period=period).delete()
        period.exceptions.all().delete()
        WorkshopSettlement.objects.filter(period=period).update(
            status=WorkshopSettlement.Status.PENDING,
            iterations=0,
            converged=False,
            settled_at=None,
            settled_by=None,
            order_cost=ZERO,
            production_receipt_amount=ZERO,
            error_message="",
        )
        OpeningInventoryBalance.objects.filter(period=period).delete()
        OpeningWip.objects.filter(
            period=period, is_system_generated=True
        ).delete()
        ProductionOrder.objects.filter(
            period=period, is_system_generated=True
        ).delete()
        PurchaseContractLine.objects.filter(
            period=period, is_system_generated=True
        ).delete()
        period.status = AccountingPeriod.Status.OPEN
        period.last_run_at = None
        period.save(update_fields=["status", "last_run_at"])
    if later_periods:
        from settlement.services.matching import (
            refresh_purchase_invoice_metadata,
            refresh_sales_invoice_metadata,
        )

        refresh_purchase_invoice_metadata(source_period)
        refresh_sales_invoice_metadata(source_period)
    return len(later_periods)
