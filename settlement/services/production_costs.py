from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP

from settlement.models import (
    MatchAllocation,
    MatchingMode,
    WorkshopSettlement,
)


ZERO = Decimal("0")
CENT = Decimal("0.01")
FEE_FIELDS = (
    ("water_power", "水电"),
    ("steam", "蒸汽"),
    ("natural_gas", "天然气"),
    ("depreciation", "折旧"),
    ("consumables", "机物料"),
    ("wages", "工资"),
)


def money(value):
    return Decimal(value or ZERO).quantize(CENT, rounding=ROUND_HALF_UP)


def _current_material_by_order(period, order_ids):
    amounts = defaultdict(lambda: ZERO)
    allocations = MatchAllocation.objects.filter(
        period=period,
        kind=MatchAllocation.Kind.ORDER_MATERIAL,
        production_order_id__in=order_ids,
    ).select_related("inventory_transaction")
    for allocation in allocations:
        txn = allocation.inventory_transaction
        if not txn.quantity:
            continue
        allocated_amount = Decimal(txn.amount or ZERO) * (
            allocation.matched_quantity / abs(txn.quantity)
        )
        if txn.direction == txn.Direction.OUT:
            amounts[allocation.production_order_id] += allocated_amount
        else:
            amounts[allocation.production_order_id] -= allocated_amount
    return {order_id: money(value) for order_id, value in amounts.items()}


def build_production_cost_summary(period, selected_workshop=""):
    orders = list(
        period.production_orders.exclude(
            match_mode=MatchingMode.IGNORED
        ).order_by("workshop", "order_no")
    )
    pools = {
        item.workshop: item
        for item in period.workshop_cost_pools.order_by("workshop")
    }
    settlements = {
        item.workshop: item
        for item in period.workshop_settlements.order_by("sequence", "workshop")
    }
    openings = {item.order_no: item for item in period.opening_wip.all()}
    results = {
        item.order_id: item
        for item in period.order_cost_results.select_related("order")
    }
    material_by_order = _current_material_by_order(
        period, [order.id for order in orders]
    )

    orders_by_workshop = defaultdict(list)
    for order in orders:
        orders_by_workshop[order.workshop].append(order)

    workshops = sorted(
        set(orders_by_workshop) | set(pools) | set(settlements)
    )
    rows = []
    all_order_rows = []
    for workshop in workshops:
        pool = pools.get(workshop)
        settlement = settlements.get(workshop)
        workshop_orders = orders_by_workshop.get(workshop, [])
        fee_breakdown = [
            {
                "key": field,
                "label": label,
                "amount": money(getattr(pool, field, ZERO) if pool else ZERO),
            }
            for field, label in FEE_FIELDS
        ]
        current_expenses = money(
            sum((item["amount"] for item in fee_breakdown), ZERO)
        )

        order_rows = []
        for order in workshop_orders:
            opening = openings.get(order.order_no)
            result = results.get(order.id)
            opening_wip = money(opening.total if opening else ZERO)
            current_material = material_by_order.get(order.id, ZERO)
            current_allocated_fees = None
            if result:
                current_allocated_fees = money(
                    sum(
                        (
                            Decimal(getattr(result, field, ZERO))
                            - Decimal(
                                getattr(opening, field, ZERO)
                                if opening
                                else ZERO
                            )
                            for field, _label in FEE_FIELDS
                        ),
                        ZERO,
                    )
                )
            order_row = {
                "order": order,
                "opening_wip": opening_wip,
                "current_material": current_material,
                "current_allocated_fees": current_allocated_fees,
                "total_cost": result.total_cost if result else None,
                "finished_cost": result.finished_cost if result else None,
                "ending_wip": result.wip_cost if result else None,
                "is_calculated": result is not None,
            }
            order_rows.append(order_row)
            all_order_rows.append(order_row)

        opening_wip = money(
            sum((item["opening_wip"] for item in order_rows), ZERO)
        )
        current_material = money(
            sum((item["current_material"] for item in order_rows), ZERO)
        )
        current_input = money(current_material + current_expenses)
        cost_to_account = money(opening_wip + current_input)
        finished_cost = money(
            sum(
                (
                    Decimal(item["finished_cost"])
                    for item in order_rows
                    if item["finished_cost"] is not None
                ),
                ZERO,
            )
        )
        ending_wip = money(
            sum(
                (
                    Decimal(item["ending_wip"])
                    for item in order_rows
                    if item["ending_wip"] is not None
                ),
                ZERO,
            )
        )
        accounted_output = money(finished_cost + ending_wip)
        is_calculated = bool(
            settlement
            and settlement.status == WorkshopSettlement.Status.SETTLED
            and all(item["is_calculated"] for item in order_rows)
        )
        rows.append(
            {
                "workshop": workshop,
                "settlement": settlement,
                "status_label": (
                    settlement.get_status_display() if settlement else "待结转"
                ),
                "is_calculated": is_calculated,
                "order_count": len(order_rows),
                "finished_quantity": sum(
                    (item["order"].finished_quantity for item in order_rows),
                    ZERO,
                ),
                "wip_quantity": sum(
                    (item["order"].wip_quantity for item in order_rows), ZERO
                ),
                "opening_wip": opening_wip,
                "current_material": current_material,
                "fee_breakdown": fee_breakdown,
                "current_expenses": current_expenses,
                "current_input": current_input,
                "cost_to_account": cost_to_account,
                "finished_cost": finished_cost,
                "ending_wip": ending_wip,
                "accounted_output": accounted_output,
                "difference": money(cost_to_account - accounted_output),
                "order_rows": order_rows,
            }
        )

    selected_rows = [
        row
        for row in rows
        if not selected_workshop or row["workshop"] == selected_workshop
    ]
    selected_order_rows = [
        item
        for row in selected_rows
        for item in row["order_rows"]
    ]
    totals = {}
    for field in (
        "opening_wip",
        "current_material",
        "current_expenses",
        "current_input",
        "cost_to_account",
        "finished_cost",
        "ending_wip",
        "accounted_output",
        "difference",
    ):
        totals[field] = money(
            sum((Decimal(row[field]) for row in selected_rows), ZERO)
        )
    totals.update(
        {
            "workshop_count": len(selected_rows),
            "order_count": sum(row["order_count"] for row in selected_rows),
            "pending_workshop_count": sum(
                1 for row in selected_rows if not row["is_calculated"]
            ),
            "all_calculated": bool(selected_rows)
            and all(row["is_calculated"] for row in selected_rows),
            "fee_breakdown": [
                {
                    "key": field,
                    "label": label,
                    "amount": money(
                        sum(
                            (
                                next(
                                    item["amount"]
                                    for item in row["fee_breakdown"]
                                    if item["key"] == field
                                )
                                for row in selected_rows
                            ),
                            ZERO,
                        )
                    ),
                }
                for field, label in FEE_FIELDS
            ],
        }
    )
    return {
        "workshop_choices": workshops,
        "selected_workshop": selected_workshop,
        "scope_label": selected_workshop or "全部车间",
        "rows": selected_rows,
        "order_rows": selected_order_rows,
        "totals": totals,
    }
