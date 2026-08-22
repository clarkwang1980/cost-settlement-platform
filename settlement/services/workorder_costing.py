"""生产工单成本核算（基于 生产工单 + 投料记录 + 报工记录）。

口径：
- 成本对象 = 工单（工单主档 ∪ 投料记录工单 ∪ 报工记录工单 ∪ 期初在产品工单）
- 车间 = 工单主档.生产部门编号；主档缺失时按 报工记录.工单部门 → 投料记录.工序名称 兜底映射
- 直接材料 = Σ(投料记录.实际投料数量 × 物料月末加权平均价)
- 完工数量 = Σ(报工记录.报工数量)；在产品 = max(计划数量 - 完工数量, 0)，计划缺失视为 0（全完工）
- 制造费用（水电/蒸汽/天然气/折旧/机物料/染料助剂）按材料成本口径分摊：
  分配率 = (本期车间费用 + 期初在产品费用) ÷ 车间材料成本总额；工单费用 = 工单直接材料 × 分配率
- 工资按折合数量（完工 + 0.5×在产品）分摊
- 完工/在产品按材料成本比例拆分（公斤产品单位材料成本一致时即数量比例）；完工单位成本 = 完工成本 / 完工数量
- 库存流水"报工"入库按单据号=工单编号关联工单，单价 = 该工单完工单位成本
"""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP

from django.db.models import Sum

from settlement.models import (
    AccountingPeriod,
    InventoryTransaction,
    OpeningWip,
    WorkOrder,
    WorkOrderCostResult,
    WorkReport,
)
from settlement.services.imports import normalize_workshop

ZERO = Decimal("0")
CENT = Decimal("0.01")
PRICE_PRECISION = Decimal("0.00000001")

FEE_FIELDS = (
    ("water_power", "水电"),
    ("steam", "蒸汽"),
    ("natural_gas", "天然气"),
    ("depreciation", "折旧"),
    ("consumables", "机物料"),
    ("dye_auxiliary", "染料助剂"),
    ("wages", "工资"),
)
# 制造费用类（按材料成本口径分摊）；工资单独按折合数量分摊
MANUFACTURING_FIELDS = (
    "water_power",
    "steam",
    "natural_gas",
    "depreciation",
    "consumables",
    "dye_auxiliary",
)
WIP_CONVERSION_RATE = Decimal("0.5")


def money(value) -> Decimal:
    return Decimal(value or ZERO).quantize(CENT, rounding=ROUND_HALF_UP)


def price(value) -> Decimal:
    return Decimal(value or ZERO).quantize(PRICE_PRECISION, rounding=ROUND_HALF_UP)


def clamp_rate(value) -> Decimal:
    value = Decimal(value or ZERO)
    return min(max(value, ZERO), Decimal("1"))


def is_baogong_receipt(txn: InventoryTransaction) -> bool:
    """库存流水"报工"入库（完工入库），按工单完工成本计价。"""
    return txn.direction == InventoryTransaction.Direction.IN and (
        txn.business_type or ""
    ).strip() == "报工"


def resolve_workshop(period, order_no: str) -> str:
    """工单车间：主档部门编码优先，其次报工记录工单部门，最后投料记录工序名称。"""
    wo = WorkOrder.objects.filter(period=period, order_no=order_no).first()
    if wo and wo.production_dept_code:
        return wo.production_dept_code
    dept = (
        WorkReport.objects.filter(period=period, work_order_no=order_no)
        .exclude(work_order_department="")
        .values_list("work_order_department", flat=True)
        .first()
    )
    if dept:
        normalized = normalize_workshop(dept)
        if normalized:
            return normalized
    return ""


def _order_product(period, order_no: str):
    report = (
        WorkReport.objects.filter(period=period, work_order_no=order_no)
        .exclude(material_code="")
        .first()
    )
    if report:
        return report.material_code, report.material_name
    return "", ""


def _collect_work_orders(period: AccountingPeriod):
    """返回需要核算的工单列表，主档缺失的自动补建。"""
    masters = {wo.order_no: wo for wo in period.work_orders.all()}
    issue_nos = set(
        period.material_issues.exclude(work_order_no="").values_list(
            "work_order_no", flat=True
        )
    )
    report_nos = set(
        period.work_reports.exclude(work_order_no="").values_list(
            "work_order_no", flat=True
        )
    )
    # 期初在产品：仅纳入有成本的工单（全零占位行不参与核算）
    opening_nos = set(
        period.opening_wip.exclude(
            material=ZERO,
            water_power=ZERO,
            steam=ZERO,
            natural_gas=ZERO,
            depreciation=ZERO,
            consumables=ZERO,
            dye_auxiliary=ZERO,
            wages=ZERO,
        ).values_list("order_no", flat=True)
    )
    all_nos = set(masters) | issue_nos | report_nos | opening_nos
    orders = []
    for order_no in sorted(all_nos):
        wo = masters.get(order_no)
        if wo is None:
            wo, _created = WorkOrder.objects.update_or_create(
                period=period,
                order_no=order_no,
                defaults={
                    "order_type": "自动补建",
                    "remark": "由投料/报工台账自动补建",
                    "production_dept_code": resolve_workshop(period, order_no),
                },
            )
        orders.append(wo)
    return orders


def _material_by_order(period: AccountingPeriod, prices: dict):
    """按 工单 × 物料 汇总实际投料数量，乘以物料月末加权价。"""
    amounts = defaultdict(lambda: ZERO)
    issues = (
        period.material_issues.exclude(work_order_no="")
        .exclude(material_code="")
        .values("work_order_no", "material_code", "actual_issue_quantity")
    )
    for item in issues:
        material_price = Decimal(prices.get(item["material_code"], ZERO) or ZERO)
        amount = Decimal(item["actual_issue_quantity"] or ZERO) * material_price
        amounts[item["work_order_no"]] += amount
    return {order_no: money(value) for order_no, value in amounts.items()}


def _finished_by_order(period: AccountingPeriod):
    """按工单汇总报工数量（完工数量）。

    浆粕车间无报工记录：完工数量以库存流水"报工"入库数量为准
    （用户每月上传浆粕成品入库表，业务类型=报工，单据编号自动关联当月浆粕工单）。
    """
    rows = (
        period.work_reports.exclude(work_order_no="")
        .values("work_order_no")
        .annotate(total=Sum("quantity"))
    )
    finished = {
        row["work_order_no"]: Decimal(row["total"] or ZERO)
        for row in rows
    }
    for wo in period.work_orders.filter(production_dept_code="浆粕车间"):
        received = (
            period.inventory_transactions.filter(
                business_type="报工", document_no=wo.order_no
            ).aggregate(total=Sum("quantity"))["total"]
            or ZERO
        )
        if received:
            finished[wo.order_no] = max(
                finished.get(wo.order_no, ZERO), Decimal(received)
            )
    return finished


def calculate_work_order_costs(period: AccountingPeriod, workshops=None, prices=None):
    """计算并持久化工单成本结果，返回结果列表。"""
    prices = prices or {}
    opening_map = {item.order_no: item for item in period.opening_wip.all()}
    material_by_order = _material_by_order(period, prices)
    finished_by_order = _finished_by_order(period)

    orders = _collect_work_orders(period)
    if workshops is not None:
        workshop_set = set(workshops)
        orders = [wo for wo in orders if wo.production_dept_code in workshop_set]

    pools = {
        item.workshop: item
        for item in period.workshop_cost_pools.all()
    }

    results_to_replace = WorkOrderCostResult.objects.filter(period=period)
    if workshops is not None:
        results_to_replace = results_to_replace.filter(
            work_order__production_dept_code__in=set(workshops)
        )
    results_to_replace.delete()

    created = []
    for wo in orders:
        opening = opening_map.get(wo.order_no)
        opening_material = Decimal(opening.material if opening else ZERO)
        direct_material = money(
            Decimal(material_by_order.get(wo.order_no, ZERO))
            + opening_material
        )
        finished_qty = finished_by_order.get(wo.order_no, ZERO)
        plan_qty = Decimal(wo.plan_quantity or ZERO)
        wip_qty = max(plan_qty - finished_qty, ZERO)
        if (
            direct_material == ZERO
            and finished_qty == ZERO
            and wip_qty == ZERO
        ):
            continue

        product_code, product_name = _order_product(period, wo.order_no)
        pool = pools.get(wo.production_dept_code)
        allocated = {}
        for field, _label in FEE_FIELDS:
            allocated[field] = ZERO

        # 制造费用（水电/蒸汽/天然气/折旧/机物料/染料助剂）按材料成本口径分摊：
        #   分配率 = (本期车间费用 + 期初在产品费用合计) ÷ 车间材料成本总额
        #            （材料成本总额 = Σ工单直接材料 = 完工材料 + 在产品材料）
        #   工单费用 = 工单直接材料 × 分配率（完工/在产品按材料成本比例拆分）
        ws = wo.production_dept_code
        workshop_material_total = ZERO
        workshop_opening_fee = defaultdict(Decimal)
        for other in orders:
            if other.production_dept_code != ws:
                continue
            opening_other = opening_map.get(other.order_no)
            workshop_material_total += Decimal(
                material_by_order.get(other.order_no, ZERO)
            ) + Decimal(opening_other.material if opening_other else ZERO)
            for field in MANUFACTURING_FIELDS:
                workshop_opening_fee[field] += Decimal(
                    getattr(opening_other, field, ZERO) if opening_other else ZERO
                )
        for field in MANUFACTURING_FIELDS:
            pool_amount = Decimal(getattr(pool, field, ZERO) if pool else ZERO)
            total_fee = workshop_opening_fee[field] + pool_amount
            rate = (
                (total_fee / workshop_material_total)
                if workshop_material_total
                else ZERO
            )
            allocated[field] = money(Decimal(direct_material) * rate)

        # 工资按折合数量（完工 + 0.5×在产品）占比分摊
        workshop_equivalent = ZERO
        equivalent_by_order = {}
        for other in orders:
            if other.production_dept_code != wo.production_dept_code:
                continue
            other_finished = finished_by_order.get(other.order_no, ZERO)
            other_plan = Decimal(other.plan_quantity or ZERO)
            other_wip = max(other_plan - other_finished, ZERO)
            equivalent = other_finished + other_wip * WIP_CONVERSION_RATE
            equivalent_by_order[other.order_no] = equivalent
            workshop_equivalent += equivalent
        pool_wages = Decimal(pool.wages if pool else ZERO)
        my_equivalent = equivalent_by_order.get(wo.order_no, ZERO)
        wage_rate = (
            pool_wages / workshop_equivalent if workshop_equivalent else ZERO
        )
        allocated["wages"] = money(
            my_equivalent * wage_rate
            + Decimal(getattr(opening, "wages", ZERO) if opening else ZERO)
        )

        # 完工/在产品拆分：
        #   材料、制造费用按材料成本比例（公斤产品单位材料成本一致时即数量比例）
        #   工资按折合工时（完工 + 0.5×在产品）占比拆分（工资口径暂保持）
        if wip_qty <= ZERO:
            finished_share = Decimal("1")
        elif finished_qty <= ZERO:
            finished_share = ZERO
        else:
            finished_share = clamp_rate(
                finished_qty / (finished_qty + wip_qty)
            )
        finished_material = money(direct_material * finished_share)
        wip_material = money(direct_material - finished_material)
        overhead_total = sum(
            (
                allocated[field]
                for field, _label in FEE_FIELDS
                if field != "wages"
            ),
            ZERO,
        )
        finished_overhead = money(overhead_total * finished_share)
        wip_overhead = money(overhead_total - finished_overhead)
        wages_total = allocated["wages"]
        finished_hours = finished_qty
        wip_hours = wip_qty * WIP_CONVERSION_RATE
        hour_total = finished_hours + wip_hours
        finished_hour_share = (
            finished_hours / hour_total if hour_total else ZERO
        )
        finished_wages = money(wages_total * finished_hour_share)
        wip_wages = money(wages_total - finished_wages)
        finished_cost = money(finished_material + finished_overhead + finished_wages)
        wip_cost = money(wip_material + wip_overhead + wip_wages)
        total_cost = money(finished_cost + wip_cost)
        unit_cost = (
            price(finished_cost / finished_qty)
            if finished_qty
            else ZERO
        )
        result = WorkOrderCostResult.objects.create(
            period=period,
            work_order=wo,
            product_code=product_code,
            product_name=product_name,
            plan_quantity=plan_qty,
            finished_quantity=finished_qty,
            wip_quantity=wip_qty,
            direct_material=direct_material,
            water_power=allocated["water_power"],
            steam=allocated["steam"],
            natural_gas=allocated["natural_gas"],
            depreciation=allocated["depreciation"],
            consumables=allocated["consumables"],
            dye_auxiliary=allocated["dye_auxiliary"],
            wages=allocated["wages"],
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
        created.append(result)
    return created


def price_baogong_receipts(period: AccountingPeriod, workshops=None) -> Decimal:
    """库存流水"报工"入库按工单完工单位成本计价，返回最大变动金额（用于收敛判断）。"""
    results = {
        item.work_order.order_no: item
        for item in period.work_order_cost_results.select_related("work_order").all()
    }
    txns = period.inventory_transactions.filter(business_type="报工")
    if workshops is not None:
        workshop_set = set(workshops)
        txns = [
            txn
            for txn in txns
            if results.get(txn.document_no)
            and results[txn.document_no].work_order.production_dept_code
            in workshop_set
        ]
    changes = []
    for txn in txns:
        result = results.get(txn.document_no)
        if not result or not result.finished_quantity:
            continue
        unit_cost = price(result.finished_cost / result.finished_quantity)
        old_amount = Decimal(txn.amount or ZERO)
        txn.unit_price = unit_cost
        txn.amount = money(Decimal(txn.quantity or ZERO) * unit_cost)
        txn.price_source = "完工入库成本"
        txn.save(update_fields=["unit_price", "amount", "price_source"])
        changes.append(abs(Decimal(txn.amount) - old_amount))
    return max(changes, default=ZERO)


def clear_baogong_receipt_prices(period: AccountingPeriod, workshops=None):
    """重置时清空"报工"入库的计价。"""
    txns = period.inventory_transactions.filter(business_type="报工")
    if workshops is not None:
        workshop_set = set(workshops)
        results = {
            item.work_order.order_no: item
            for item in period.work_order_cost_results.select_related("work_order").all()
        }
        txns = [
            txn
            for txn in txns
            if results.get(txn.document_no)
            and results[txn.document_no].work_order.production_dept_code
            in workshop_set
        ]
    for txn in txns:
        txn.unit_price = None
        txn.amount = None
        txn.price_source = ""
    if txns:
        InventoryTransaction.objects.bulk_update(
            txns, ["unit_price", "amount", "price_source"]
        )


def build_work_order_cost_summary(period: AccountingPeriod, selected_workshop=""):
    """本月生产成本与费用（工单口径）。"""
    from settlement.models import WorkshopSettlement

    results = list(
        period.work_order_cost_results.select_related("work_order").order_by(
            "work_order__production_dept_code", "work_order__order_no"
        )
    )
    if selected_workshop:
        results = [
            item
            for item in results
            if item.work_order.production_dept_code == selected_workshop
        ]
    pools = {
        item.workshop: item
        for item in period.workshop_cost_pools.order_by("workshop")
    }
    settlements = {
        item.workshop: item
        for item in period.workshop_settlements.order_by("sequence", "workshop")
    }
    openings = {item.order_no: item for item in period.opening_wip.all()}

    by_workshop = defaultdict(list)
    for item in results:
        by_workshop[item.work_order.production_dept_code].append(item)

    workshops = sorted(
        set(by_workshop) | set(pools) | set(settlements)
    )
    rows = []
    all_order_rows = []
    for workshop in workshops:
        pool = pools.get(workshop)
        settlement = settlements.get(workshop)
        workshop_results = by_workshop.get(workshop, [])
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
        for result in workshop_results:
            order = result.work_order
            opening = openings.get(order.order_no)
            opening_wip = money(opening.total if opening else ZERO)
            current_material = money(
                result.direct_material
                - (Decimal(opening.material if opening else ZERO))
            )
            current_allocated_fees = money(
                sum(
                    (
                        Decimal(getattr(result, field, ZERO))
                        - Decimal(
                            getattr(opening, field, ZERO) if opening else ZERO
                        )
                        for field, _label in FEE_FIELDS
                    ),
                    ZERO,
                )
            )
            order_row = {
                "order": order,
                "result": result,
                "opening_wip": opening_wip,
                "current_material": current_material,
                "current_allocated_fees": current_allocated_fees,
                "total_cost": result.total_cost,
                "finished_cost": result.finished_cost,
                "ending_wip": result.wip_cost,
                "is_calculated": settlement
                and settlement.status
                == WorkshopSettlement.Status.SETTLED,
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
            sum((item["finished_cost"] for item in order_rows), ZERO)
        )
        ending_wip = money(
            sum((item["ending_wip"] for item in order_rows), ZERO)
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
                    (item["result"].finished_quantity for item in order_rows),
                    ZERO,
                ),
                "wip_quantity": sum(
                    (item["result"].wip_quantity for item in order_rows),
                    ZERO,
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
            sum((Decimal(row[field]) for row in rows), ZERO)
        )
    totals.update(
        {
            "workshop_count": len(rows),
            "order_count": sum(row["order_count"] for row in rows),
            "pending_workshop_count": sum(
                1 for row in rows if not row["is_calculated"]
            ),
            "all_calculated": bool(rows)
            and all(row["is_calculated"] for row in rows),
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
                                for row in rows
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
        "rows": rows,
        "order_rows": all_order_rows,
        "totals": totals,
    }
