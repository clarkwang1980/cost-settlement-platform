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

from django.db.models import Max, Sum

from settlement.models import (
    AccountingPeriod,
    BomLine,
    ImportBatch,
    InventoryTransaction,
    MatchingMode,
    MaterialIssue,
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

# 视同完工：工单有完工时间且完工率达标的工单，尾料按损耗处理，在产品清零、成本并入完工
# （避免"已结束工单在产品永远挂账结转"）
# 完工时间口径与工单成本页「实际完工时间」列一致：工单明细实际完工时间 → 报工记录最大报工时间兜底；
# 或完成状态为已完工（超时完成/按时完成/已完成）
# （染色/纺纱/包覆/针织/机织适用；浆粕不适用）
MIN_FINISH_RATE_PCT = Decimal("70")     # 完工率阈值（%）：完工÷计划×100
FINISHED_WORKORDER_STATUSES = ("超时完成", "按时完成", "已完成")  # 工单明细中视为已完工的完成状态
SMALL_ORDER_WORKSHOPS = ("RanZheng", "tefang", "BaoFu", "ZhenZhi", "JiZhi")


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


def build_product_resolver(period: AccountingPeriod):
    """构建"工单 -> 产品编码/名称"解析器（含 BOM 主产出索引，显示层兜底用）。

    返回 callable: resolve(work_order_obj, product_code, product_name) -> (code, name)
    匹配顺序：
      1) 传入的 product_code/product_name（结算时来自报工记录，优先保留）
      2) BOM 主产出物料编码/名称（该工单 BOM 第一条主产出，编码/名称最完整）
      3) 工单标识码 identifier 前缀提取（如 14.1661-BF20260413 -> 14.1661）
      4) 工单编号本身（最后兜底，仅编码）
    """
    bom_product = {}
    for line in period.bom_lines.exclude(material_code="").order_by("id"):
        bom_product.setdefault(
            line.work_order_no, (line.material_code, line.material_name)
        )

    import re as _re

    def resolve(work_order, product_code="", product_name=""):
        order_no = work_order.order_no
        code = product_code or ""
        name = product_name or ""
        bom = bom_product.get(order_no)
        if not code:
            if bom and bom[0]:
                code = bom[0]
            else:
                ident = work_order.identifier or order_no
                m = _re.match(r"^([A-Za-z0-9]+(?:\.[0-9]+)?)", ident)
                code = m.group(1) if m else ident
        if not name:
            if bom and bom[1]:
                name = bom[1]
        return code, name

    return resolve


def _collect_work_orders(period: AccountingPeriod):
    """返回需要核算的工单列表。

    只核算"有业务"的工单：投料记录 / 报工记录 / 报工入库流水 / 期初在产品
    （有主档但本期间无任何业务的工单不核算——计划时间与核算月份无关，
     成本按业务发生的期间归属）。主档仅用于补充计划数量、车间等属性。

    主档缺失的台账工单不允许自动补建（用户要求）：直接报错中断，
    提示先导入工单主档或核对台账单号。
    """
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
    # 报工入库流水（浆粕/粗纱等中间品，单据号=工单号）
    receipt_nos = set(
        period.inventory_transactions.filter(business_type="报工")
        .exclude(document_no="")
        .values_list("document_no", flat=True)
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
    all_nos = issue_nos | report_nos | receipt_nos | opening_nos
    missing = sorted(n for n in all_nos if n not in masters)
    if missing:
        shown = "、".join(missing[:20])
        suffix = f" 等共 {len(missing)} 个" if len(missing) > 20 else f"（共 {len(missing)} 个）"
        raise ValueError(
            "以下工单在工单主档中不存在，无法核算（系统不允许自动补建）："
            f"{shown}{suffix}。请先在『工单主档』导入这些工单，"
            "或核对投料/报工/期初在产品台账中的工单编号。"
        )
    return [masters[n] for n in sorted(all_nos)]


def _material_by_order(period: AccountingPeriod, prices: dict, roving_plan=None):
    """按 工单 × 物料 汇总实际投料数量，乘以物料月末加权价；
    叠加计算层粗纱补投（roving_plan.auto_issues，粗纱无月末加权价时用粗纱计算单价）；
    再对有报工、无投料、无期初、非粗纱链的工单按 BOM 单耗×报工数量 反推补投材料成本（仅计算层）。

    单价统一口径（与页面/导出一致）：月末加权价优先 → 历史单价（入库价优先、出库价兜底）；
    两者都没有则该物料金额 0。
    """
    prices = dict(prices)
    for _hp in period.historical_prices.all():
        if _hp.receipt_price is not None and not prices.get(_hp.material_code):
            prices[_hp.material_code] = Decimal(_hp.receipt_price)
        elif _hp.issue_price is not None and not prices.get(_hp.material_code):
            prices[_hp.material_code] = Decimal(_hp.issue_price)
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
    if roving_plan:
        roving_prices = roving_plan.get("roving_prices", {})
        for order_no, auto_list in roving_plan.get("auto_issues", {}).items():
            for material_code, _name, qty, _uom in auto_list:
                material_price = Decimal(prices.get(material_code, ZERO) or ZERO)
                if not material_price:
                    material_price = Decimal(roving_prices.get(material_code, ZERO) or ZERO)
                amounts[order_no] += Decimal(qty) * material_price
    _backfill_material_by_bom(period, prices, amounts)
    return {order_no: money(value) for order_no, value in amounts.items()}


# 有报工、无投料、无期初在产品、非粗纱链的工单（染色/纺纱/包覆/针织/机织）：
# 材料成本按 BOM 反推补投（仅计算层，不落库存/投料台账）
# - 染色/纺纱/包覆/针织：BOM 单耗 × 报工数量 × 物料价
# - 机织：BOM 需求数 × 完工比例 × 物料价（完工比例 = min(报工/计划, 1)，超产封顶；
#         需求数 = BOM模板"需求数"列，缺失时用 计划×单耗 现算，与 ERP 申请数量口径一致）
BOM_BACKFILL_WORKSHOPS = ("RanZheng", "tefang", "BaoFu", "ZhenZhi", "JiZhi")
# 机织用"需求数×完工比例"口径；其余车间维持"单耗×报工"
JIZHI_BACKFILL = ("JiZhi",)
# 机织完工比例超产判定阈值（报工/计划 > 120% 视为异常，按需求数全额，不乘超额）
JIZHI_OVERRUN_PCT = Decimal("120")


def _backfill_material_by_bom(period, prices, amounts):
    reported = set(
        period.work_reports.exclude(work_order_no="")
        .values_list("work_order_no", flat=True)
    )
    if not reported:
        return
    issued = set(
        period.material_issues.exclude(work_order_no="")
        .values_list("work_order_no", flat=True)
    )
    opening = set(
        period.opening_wip.exclude(material=ZERO)
        .values_list("order_no", flat=True)
    )
    roving = set(
        period.bom_lines.filter(sub_item_code__startswith="BC-")
        .values_list("work_order_no", flat=True)
    )
    roving |= set(
        period.material_issues.filter(material_code__startswith="BC-")
        .exclude(work_order_no="")
        .values_list("work_order_no", flat=True)
    )
    roving |= set(
        period.issue_progress.filter(material_code__startswith="BC-")
        .exclude(work_order_no="")
        .values_list("work_order_no", flat=True)
    )
    dept_map = {
        o.order_no: o.production_dept_code
        for o in period.work_orders.all()
    }
    missing = sorted(
        n for n in (reported - issued - opening - roving)
        if dept_map.get(n) in BOM_BACKFILL_WORKSHOPS
    )
    if not missing:
        return
    # 报工数量合计（与完工数量口径一致）
    finish_qty = defaultdict(Decimal)
    for row in period.work_reports.exclude(work_order_no="").values(
        "work_order_no", "quantity"
    ):
        finish_qty[row["work_order_no"]] += Decimal(row["quantity"] or ZERO)
    # 计划数量（机织完工比例）
    plan_qty = {
        o.order_no: Decimal(o.plan_quantity or ZERO)
        for o in period.work_orders.all()
    }
    # BOM 行按工单分组
    bom_by_order = defaultdict(list)
    for line in period.bom_lines.filter(work_order_no__in=missing):
        bom_by_order[line.work_order_no].append(line)
    # 历史单价兜底（月末加权价缺失时）：入库价优先、出库价兜底
    hist_price = {}
    for hp in period.historical_prices.all():
        if hp.receipt_price is not None:
            hist_price[hp.material_code] = Decimal(hp.receipt_price)
        elif hp.issue_price is not None:
            hist_price[hp.material_code] = Decimal(hp.issue_price)
    for order_no in missing:
        qty = finish_qty.get(order_no, ZERO)
        if qty <= ZERO:
            continue
        ws = dept_map.get(order_no)
        if ws in JIZHI_BACKFILL:
            # 机织：BOM 需求数 × 完工比例
            plan = plan_qty.get(order_no, ZERO)
            if plan > ZERO:
                ratio = qty / plan
                if ratio > Decimal("1"):
                    ratio = Decimal("1")  # 超产封顶（≤120% 按需求数全额，超 120% 也按全额并提示）
                if qty > plan * JIZHI_OVERRUN_PCT / Decimal("100"):
                    ratio = Decimal("1")
            else:
                ratio = Decimal("1")  # 计划缺失视为全完工（与系统口径一致）
            for line in bom_by_order.get(order_no, []):
                sub_qty = _jizhi_line_need(line, plan) * ratio
                if sub_qty <= ZERO:
                    continue
                mat_price = Decimal(prices.get(line.sub_item_code, ZERO) or ZERO)
                if not mat_price:
                    mat_price = hist_price.get(line.sub_item_code, ZERO)
                if not mat_price:
                    continue
                amounts[order_no] += sub_qty * mat_price
        else:
            # 染色/纺纱/包覆/针织：BOM 单耗 × 报工数量（维持原口径）
            for line in bom_by_order.get(order_no, []):
                sub_qty = Decimal(line.qty) * qty
                if sub_qty <= ZERO:
                    continue
                mat_price = Decimal(prices.get(line.sub_item_code, ZERO) or ZERO)
                if not mat_price:
                    mat_price = hist_price.get(line.sub_item_code, ZERO)
                if not mat_price:
                    continue
                amounts[order_no] += sub_qty * mat_price


def _jizhi_line_need(line, plan_qty):
    """机织 BOM 行的需求数：优先取模板"需求数"列；缺失时用 计划×单耗×(1+损耗率) 现算。"""
    if line.required_qty is not None and line.required_qty > ZERO:
        return Decimal(line.required_qty)
    base = Decimal(line.qty) * (plan_qty if plan_qty > ZERO else Decimal("1"))
    if line.loss_rate is not None and line.loss_rate > ZERO:
        base = base * (Decimal("1") + Decimal(line.loss_rate) / Decimal("100"))
    return base


def _roving_output_code(period: AccountingPeriod, order_no: str) -> str:
    """粗纱工单的产出物料：优先取BOM成品物料编号，其次从工单号前缀提取。"""
    code = (
        BomLine.objects.filter(period=period, work_order_no=order_no)
        .exclude(material_code="")
        .values_list("material_code", flat=True)
        .first()
    )
    if code:
        return code
    prefix = order_no
    for sep in ("-20", "-19"):
        idx = prefix.find(sep)
        if idx > 0:
            prefix = prefix[:idx]
            break
    return prefix


def build_roving_plan(period: AccountingPeriod, prices=None):
    """粗纱链计算层计划（仅计算，不写任何库存流水/投料记录）：

    1) BC-开头粗纱工单：无报工记录且无"报工"入库流水的，
       计算层视为完工（完工数量=计划数量），并按BOM自动补原料投料；
    2) 有BOM粗纱子项、未实际投该粗纱且已有报工数量的下游工单：
       计算层补粗纱投料（数量=BOM单耗×报工数量）。
    3) 粗纱计算单价 = 粗纱工单BOM原料成本÷计划数量（供下游补投粗纱计价）。

    返回 {"finished": {工单号: 数量}, "auto_issues": {工单号: [(物料, 名称, 数量, 单位)...]},
           "roving_prices": {粗纱物料编码: 单价}}。
    """
    from decimal import Decimal as _D

    prices = prices or {}
    plan = {"finished": {}, "auto_issues": defaultdict(list), "roving_prices": {}}

    reported = set(
        period.work_reports.exclude(work_order_no="")
        .values_list("work_order_no", flat=True)
    )
    receipt_orders = set(
        period.inventory_transactions.filter(business_type="报工")
        .exclude(document_no="")
        .values_list("document_no", flat=True)
    )
    bc_orders = list(period.work_orders.filter(order_no__startswith="BC-"))
    bom_subs = defaultdict(list)
    for line in period.bom_lines.filter(
        work_order_no__in=[w.order_no for w in bc_orders]
    ):
        bom_subs[line.work_order_no].append(line)

    for wo in bc_orders:
        if wo.order_no in reported or wo.order_no in receipt_orders:
            continue  # 已有真实报工/报工入库，以真实数据为准
        plan_qty = _D(wo.plan_quantity or ZERO)
        if plan_qty <= ZERO:
            continue
        material_code = _roving_output_code(period, wo.order_no)
        if not material_code:
            continue
        plan["finished"][wo.order_no] = plan_qty
        material_cost = ZERO
        for line in bom_subs.get(wo.order_no, []):
            qty = _D(line.qty) * plan_qty
            if qty <= ZERO:
                continue
            plan["auto_issues"][wo.order_no].append(
                (line.sub_item_code, line.sub_item_name, qty, line.uom or "kg")
            )
            material_cost += qty * _D(prices.get(line.sub_item_code, ZERO) or ZERO)
        plan["roving_prices"][material_code] = (
            material_cost / plan_qty
        ).quantize(Decimal("0.0001"))

    # 下游工单计算层补粗纱投料
    for line in period.bom_lines.filter(sub_item_code__startswith="BC-"):
        order_no = line.work_order_no
        roving_code = line.sub_item_code
        if not WorkOrder.objects.filter(period=period, order_no=order_no).exists():
            continue
        if MaterialIssue.objects.filter(
            period=period, work_order_no=order_no, material_code=roving_code
        ).exists():
            continue  # 已有真实投料，不补
        finished = (
            period.work_reports.filter(work_order_no=order_no)
            .aggregate(total=Sum("quantity"))["total"]
            or ZERO
        )
        if not finished:
            continue
        plan["auto_issues"][order_no].append(
            (roving_code, line.sub_item_name, _D(line.qty) * _D(finished), line.uom or "kg")
        )
    return plan


def _finished_by_order(period: AccountingPeriod, roving_plan=None):
    """按工单汇总报工数量（完工数量）。

    浆粕车间无报工记录：完工数量以库存流水"报工"入库数量为准
    （用户每月上传浆粕成品入库表，业务类型=报工，单据编号自动关联当月浆粕工单）。
    粗纱（BC-开头）等中间品工单：结算前由 prepare_roving_chain 自动生成
    "报工"入库流水，完工数量同样取报工入库数量（无报工记录时兜底）。
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
    # 兜底：无报工记录、但有"报工"入库流水的工单（浆粕/粗纱等中间品），
    # 完工数量 = 报工入库数量合计（按单据编号=工单号关联）
    receipt_totals = (
        period.inventory_transactions.filter(business_type="报工")
        .exclude(document_no="")
        .values("document_no")
        .annotate(total=Sum("quantity"))
    )
    for row in receipt_totals:
        order_no = row["document_no"]
        if order_no in finished:
            continue
        received = Decimal(row["total"] or ZERO)
        if received:
            finished[order_no] = received
    # 计算层粗纱自动完工（无报工、无入库流水的 BC 工单，按计划数量）
    if roving_plan:
        for order_no, qty in roving_plan.get("finished", {}).items():
            if order_no not in finished and qty:
                finished[order_no] = qty
    return finished


def calculate_work_order_costs(period: AccountingPeriod, workshops=None, prices=None):
    """计算并持久化工单成本结果，返回结果列表。"""
    prices = prices or {}
    opening_map = {item.order_no: item for item in period.opening_wip.all()}
    # 工单明细（用户上传）：用于视同完工判定（实际完工时间 / 完成状态）
    work_order_detail_map = {
        d.order_no: d for d in period.work_order_details.all()
    }
    # 各工单实际完工时间兜底：报工记录最大报工时间（与工单成本页「实际完工时间」列口径一致）
    report_time_agg = {
        row["work_order_no"]: row["max_time"]
        for row in period.work_reports.exclude(work_order_no="")
        .values("work_order_no")
        .annotate(max_time=Max("report_time"))
    }
    # 粗纱链计算层计划（仅计算不落库：自动完工/补投/粗纱计算单价）
    roving_plan = build_roving_plan(period, prices)
    material_by_order = _material_by_order(period, prices, roving_plan)
    finished_by_order = _finished_by_order(period, roving_plan)

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
    # 每工单实际投料数量合计（用于纺纱"不用粗纱"工单的在产品口径：在产品=投料kg-完工kg）
    issue_qty_by_order = defaultdict(Decimal)
    for row in period.material_issues.exclude(work_order_no="").values(
        "work_order_no", "actual_issue_quantity"
    ):
        issue_qty_by_order[row["work_order_no"]] += Decimal(
            row["actual_issue_quantity"] or ZERO
        )
    # 用粗纱的工单（BOM子项或投料含 BC-）：粗纱自动补投=按完工反推，无在产品
    roving_orders = set(
        period.bom_lines.filter(sub_item_code__startswith="BC-")
        .values_list("work_order_no", flat=True)
    ) | set(
        period.material_issues.filter(material_code__startswith="BC-")
        .exclude(work_order_no="")
        .values_list("work_order_no", flat=True)
    )

    def _wip_qty(wo_obj, finished_qty_val, plan_qty_val):
        """纺纱车间在产品口径（用户规则）：
        粗纱工单(BC-)/用粗纱工单 → 0（粗纱无在产品）；
        不用粗纱工单 → max(投料kg-完工kg, 0)；
        其他车间 → max(计划-完工, 0)。
        """
        ws = wo_obj.production_dept_code
        order_no = wo_obj.order_no
        if ws == "tefang":
            if str(order_no).upper().startswith("BC-") or order_no in roving_orders:
                return ZERO
            if issue_qty_by_order.get(order_no, ZERO) > ZERO:
                return max(issue_qty_by_order[order_no] - finished_qty_val, ZERO)
        return max(plan_qty_val - finished_qty_val, ZERO)

    for wo in orders:
        opening = opening_map.get(wo.order_no)
        opening_material = Decimal(opening.material if opening else ZERO)
        direct_material = money(
            Decimal(material_by_order.get(wo.order_no, ZERO))
            + opening_material
        )
        finished_qty = finished_by_order.get(wo.order_no, ZERO)
        plan_qty = Decimal(wo.plan_quantity or ZERO)
        ws = wo.production_dept_code
        wip_qty = _wip_qty(wo, finished_qty, plan_qty)
        # 视同完工：① 工单有完工时间（工单明细实际完工时间，或报工记录最大报工时间兜底，与页面口径一致）
        #          ② 完成状态为已完工（超时完成/按时完成/已完成）；完工率达标 → 尾料按损耗处理，
        #             在产品清零、成本并入完工（避免"已结束工单在产品永远挂账结转"）
        _detail = work_order_detail_map.get(wo.order_no)
        _detail_finished = (
            report_time_agg.get(wo.order_no) is not None
            or (
                _detail is not None
                and (
                    bool(_detail.actual_finish)
                    or (_detail.completion_status in FINISHED_WORKORDER_STATUSES)
                )
            )
        )
        if (
            ws in SMALL_ORDER_WORKSHOPS
            and _detail_finished
            and plan_qty > ZERO
            and finished_qty > ZERO
            and wip_qty > ZERO
            and (finished_qty / plan_qty * Decimal("100")) >= MIN_FINISH_RATE_PCT
        ):
            wip_qty = ZERO
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
            other_wip = _wip_qty(other, other_finished, other_plan)
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
