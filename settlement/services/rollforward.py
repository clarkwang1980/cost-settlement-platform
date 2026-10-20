from __future__ import annotations

from datetime import timedelta
from decimal import Decimal, ROUND_HALF_UP

from django.db import transaction

from settlement.models import (
    AccountingPeriod,
    EndingInventoryBalance,
    InventoryValuation,
    MatchAllocation,
    MatchingMode,
    Material,
    OpeningInventoryBalance,
    OpeningWip,
    PurchaseContractLine,
    ReconciliationResult,
    SettlementRun,
    WorkOrderCostResult,
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
        "carried_order_count": period.work_orders.filter(
            period=period
        ).count(),
        "carried_contract_count": period.purchase_contracts.filter(
            is_system_generated=True
        ).count(),
    }


def _manual_opening_inventory_codes(period):
    codes = set(
        period.opening_inventory_balances.filter(
            is_system_generated=False
        ).values_list("material_code", flat=True)
    )
    codes.update(
        period.inventory_transactions.filter(
            business_type__contains="期初"
        ).values_list("material_code", flat=True)
    )
    return codes


def _carry_inventory(period, previous):
    valuations = list(previous.inventory_valuations.all())
    active_keys = set()
    manual_codes = _manual_opening_inventory_codes(period)
    material_uoms = {
        code: uom
        for code, uom in Material.objects.filter(
            code__in=[item.material_code for item in valuations]
        ).values_list("code", "uom")
    }
    # 由上期期末承接：优先用人工上传的期末存货，否则用系统生成的期末
    previous_endings = list(
        previous.ending_inventory_balances.filter(is_system_generated=False)
    )
    if not previous_endings:
        previous_endings = list(
            previous.ending_inventory_balances.filter(
                is_system_generated=True
            )
        )
    if previous_endings:
        for ending in previous_endings:
            if ending.quantity == ZERO and ending.amount == ZERO:
                continue
            if ending.material_code in manual_codes:
                continue
            key = (
                ending.material_code,
                ending.warehouse or "",
                ending.batch_no or "",
            )
            active_keys.add(key)
            OpeningInventoryBalance.objects.update_or_create(
                period=period,
                material_code=ending.material_code,
                warehouse=ending.warehouse or None,
                batch_no=ending.batch_no or None,
                defaults={
                    "source_period": previous,
                    "material_name": ending.material_name,
                    "category_name": ending.category_name or "",
                    "uom": ending.uom or material_uoms.get(
                        ending.material_code, ""
                    ),
                    "quantity": ending.quantity,
                    "amount": ending.amount,
                    "unit_price": ending.unit_price,
                    "is_system_generated": True,
                },
            )
        period.opening_inventory_balances.filter(
            is_system_generated=True
        ).exclude(
            material_code__in={key[0] for key in active_keys}
        ).delete()
        return {
            "count": len(active_keys),
            "manual_overrides": len(
                manual_codes & {key[0] for key in active_keys}
            ),
        }
    for valuation in valuations:
        if valuation.ending_quantity == ZERO and valuation.ending_amount == ZERO:
            continue
        if valuation.material_code in manual_codes:
            continue
        active_keys.add((valuation.material_code, "", ""))
        OpeningInventoryBalance.objects.update_or_create(
            period=period,
            material_code=valuation.material_code,
            warehouse=None,
            batch_no=None,
            defaults={
                "source_period": previous,
                "material_name": valuation.material_name,
                "uom": material_uoms.get(valuation.material_code, ""),
                "quantity": valuation.ending_quantity,
                "amount": valuation.ending_amount,
                "unit_price": valuation.weighted_average_price,
                "is_system_generated": True,
            },
        )
    period.opening_inventory_balances.filter(
        is_system_generated=True
    ).exclude(material_code__in={key[0] for key in active_keys}).delete()
    return {
        "count": len(active_keys),
        "manual_overrides": len(
            manual_codes & {key[0] for key in active_keys}
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
    manual_wip_orders = set(
        period.opening_wip.filter(
            is_system_generated=False
        ).values_list("order_no", flat=True)
    )
    active_wip_orders = set()

    # 工单成本结果（生产工单+投料+报工）的在产品承接
    for result in previous.work_order_cost_results.select_related(
        "work_order"
    ).all():
        if Decimal(result.wip_cost or ZERO) == ZERO:
            continue
        order_no = result.work_order.order_no
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
                "quantity": result.wip_quantity,
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
    return {
        "wip_count": len(active_wip_orders),
        "wip_manual_overrides": len(manual_wip_orders & active_wip_orders),
        "order_count": len(active_wip_orders),
        "order_manual_overrides": len(manual_wip_orders & active_wip_orders),
    }


def _carry_contracts(period, previous):
    prior_contracts = list(
        previous.purchase_contracts.filter(remaining_quantity__gt=ZERO).exclude(
            # 忽略匹配的合同不往后续期间承接（否则每月都会把它带出来）
            match_mode=MatchingMode.IGNORED
        )
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
                # 用户 2026-09-19：暂估归属按合同的「收货期间」，这两个字段必须跟着承接 ——
                # 匹配时用的是**本期这一行**（`_candidate_contracts` 只看 period.purchase_contracts），
                # 不承接的话副本的收货期间为空 → 规则在承接后的月份失效、回落成入库流水月份。
                "receipt_period": source.receipt_period,
                "received_quantity": source.received_quantity,
                "invoice_date": source.invoice_date,
                "amount_excl_tax": source.amount_excl_tax,
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
        EndingInventoryBalance.objects.filter(
            period=period, is_system_generated=True
        ).delete()
        WorkOrderCostResult.objects.filter(period=period).delete()
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
        OpeningInventoryBalance.objects.filter(
            period=period, is_system_generated=True
        ).delete()
        OpeningWip.objects.filter(
            period=period, is_system_generated=True
        ).delete()
        PurchaseContractLine.objects.filter(
            period=period, is_system_generated=True
        ).delete()
        period.status = AccountingPeriod.Status.OPEN
        period.last_run_at = None
        period.save(update_fields=["status", "last_run_at"])
    if later_periods:
        from settlement.services.matching import refresh_purchase_invoice_metadata

        refresh_purchase_invoice_metadata(source_period)
    return len(later_periods)
