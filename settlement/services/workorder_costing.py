"""生产工单成本核算（基于 生产工单 + 投料记录 + 报工记录）。

口径：
- 成本对象 = 工单（工单主档 ∪ 投料记录工单 ∪ 报工记录工单 ∪ 期初在产品工单）
- 车间 = 工单主档.生产部门编号；主档缺失时按 报工记录.工单部门 → 投料记录.工序名称 兜底映射
- 直接材料 = Σ(投料记录.实际投料数量 × 物料月末加权平均价)
- 完工数量 = Σ(报工记录.报工数量)；在产品 = max(计划数量 - 完工数量, 0)，计划缺失视为 0（全完工）
  机织车间（JiZhi）另有重量法在产品，见 `_wip_qty`（2026-09-29 修：投料取累计口径兜底、
  上限改用累计完工，避免跨月工单在产品被算成 0）
- 制造费用（水电/蒸汽/天然气/折旧/机物料/染料助剂）按材料成本口径分摊：
  分配率 = (本期车间费用 + 期初在产品费用) ÷ 车间材料成本总额；工单费用 = 工单直接材料 × 分配率
- 工资按折合数量（完工 + 0.5×在产品）分摊
- 完工/在产品按材料成本比例拆分（公斤产品单位材料成本一致时即数量比例）；完工单位成本 = 完工成本 / 完工数量
- 库存流水"报工"入库按单据号=工单编号关联工单，单价 = 该工单完工单位成本
"""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

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
from settlement.services.imports import WORKSHOP_LABELS, normalize_workshop

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
    ("processing_fee", "加工费"),
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
    "processing_fee",
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
# 视同完工**只在"没有明显超投料"时清零**（用户 2026-09-30）：
#   报工 < 投料时，超出的部分是**留给在产的料**，不是尾料损耗。实测 13.554-HF20260725：
#   投料 772.4、报工 400.62（应有 19,915.79），材料 39,082.81 ＝ 应有 1.96 倍 →
#   清零会把超投的料全压进报工量（单位成本 120.08 vs 标准 62.61）。
#   判据用**金额**（材料 vs BOM 应有），所以机织"报工米 vs 投料kg"的量纲不同也不影响判断。
WIP_KEEP_OVERRUN_PCT = Decimal("10")    # 材料超出"应有(报工量)"这个比例以上 → 不清零、保留在产品


def money(value) -> Decimal:
    return Decimal(value or ZERO).quantize(CENT, rounding=ROUND_HALF_UP)


def price(value) -> Decimal:
    return Decimal(value or ZERO).quantize(PRICE_PRECISION, rounding=ROUND_HALF_UP)


def clamp_rate(value) -> Decimal:
    value = Decimal(value or ZERO)
    return min(max(value, ZERO), Decimal("1"))


# 完工入库业务类型：报工（染色/纺纱等 ERP 习惯）+ 生产入库（浆粕等）
# 两者都按工单完工成本计价、作为完工数量兜底
FINISHED_RECEIPT_TYPES = ("报工", "生产入库")


def is_baogong_receipt(txn: InventoryTransaction) -> bool:
    """库存流水完工入库（报工/生产入库），按工单完工成本计价。"""
    return txn.direction == InventoryTransaction.Direction.IN and (
        (txn.business_type or "").strip() in FINISHED_RECEIPT_TYPES
    )


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
        period.inventory_transactions.filter(business_type__in=FINISHED_RECEIPT_TYPES)
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
            processing_fee=ZERO,
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


def load_manual_adjustments(period):
    """读人工调整：工单号 → {issue_quantity / issue_amount / report_quantity}。

    口径（用户 2026-09-24，方案 A「覆盖」）：填了就以人工值为准，留空＝用系统算出的值。
    见 `WorkOrderManualAdjustment`；不改动原始上传台账，所以核算层的每个取数口都要叠加它。
    """
    from settlement.models import WorkOrderManualAdjustment

    if period is None:
        return {}
    return {
        row.order_no: {
            "issue_quantity": row.issue_quantity,
            "issue_amount": row.issue_amount,
            "report_quantity": row.report_quantity,
        }
        for row in period.work_order_adjustments.all()
    }


def _adjusted_report_qty(order_no, system_qty, adjustments):
    """报工数量：有人工值时以人工值为准（填 0 表示视为没有报工）。"""
    override = (adjustments or {}).get(order_no, {}).get("report_quantity")
    if override is None:
        return system_qty
    return Decimal(override)


def _code_variant_price(prices, code, name="", name_by_code=None):
    """取价时的**物料编码容错**（用户 2026-09-30 实测工单 18.441-1358059）：

    投料记录写的是 `18.660`、而「单价表 / 历史单价 / 物料主档」里是 `18.66`（**同名同料**），
    精确匹配取不到价 → 那笔投料金额记 0、整张单材料成本＝0。
    这里只容忍「小数部分多写/少写尾零」这一种差异，并且**要求物料名称一致**才采用
    （避免把"只差尾零但其实是两个料"的编码误并）。

    prices: {物料编码: 价}；name: 待查编码的物料名称；name_by_code: {编码: 名称}（可缺省）。
    """
    text = str(code or "").strip()
    if "." not in text:
        return None
    head, _, tail = text.partition(".")
    trimmed = tail.rstrip("0")
    if trimmed == tail:
        return None  # 本身没有多余尾零，交给精确匹配
    for candidate in ((f"{head}.{trimmed}" if trimmed else head), head):
        value = prices.get(candidate)
        if not value:
            continue
        other_name = (name_by_code or {}).get(candidate) or ""
        if name and other_name and str(name).strip() != str(other_name).strip():
            continue  # 名称不一致 → 不当成同一个料
        return Decimal(value)
    return None


def _material_by_order(
    period: AccountingPeriod,
    prices: dict,
    roving_plan=None,
    adjustments=None,
    backfilled_out=None,
):
    """按 工单 × 物料 汇总实际投料数量，乘以物料月末加权价；
    叠加计算层粗纱补投（roving_plan.auto_issues，粗纱无月末加权价时用粗纱计算单价）；
    再对有报工、无投料、无期初、非粗纱链的工单按 BOM 单耗×报工数量 反推补投材料成本（仅计算层）。

    单价统一口径（与页面/导出一致）：月末加权价优先 → 历史单价（入库价优先、出库价兜底）；
    两者都没有则该物料金额 0。
    backfilled_out：传入 set 时，把"材料是 BOM 反推补进来的"工单记进去
    （结算层据此把那笔料**全部**算给本期完工，见 calculate_work_order_costs）。
    """
    prices = dict(prices)
    for _hp in period.historical_prices.all():
        # 投料=出库动作，历史单价兜底用出库单价优先、入库单价兜底
        if _hp.issue_price is not None and not prices.get(_hp.material_code):
            prices[_hp.material_code] = Decimal(_hp.issue_price)
        elif _hp.receipt_price is not None and not prices.get(_hp.material_code):
            prices[_hp.material_code] = Decimal(_hp.receipt_price)
    amounts = defaultdict(lambda: ZERO)
    raw_qty = defaultdict(lambda: ZERO)
    # 编码容错要用的物料名称（取价来源里带的名称）
    name_by_code = {}
    for row in period.inventory_valuations.values("material_code", "material_name"):
        if row["material_name"]:
            name_by_code.setdefault(row["material_code"], row["material_name"])
    for row in period.historical_prices.values("material_code", "material_name"):
        if row["material_name"]:
            name_by_code.setdefault(row["material_code"], row["material_name"])
    issues = (
        period.material_issues.exclude(work_order_no="")
        .exclude(material_code="")
        .values("work_order_no", "material_code", "material_name", "actual_issue_quantity")
    )
    for item in issues:
        material_price = Decimal(prices.get(item["material_code"], ZERO) or ZERO)
        if not material_price:
            fallback = _code_variant_price(
                prices, item["material_code"], item["material_name"], name_by_code
            )
            if fallback:
                material_price = fallback
        quantity = Decimal(item["actual_issue_quantity"] or ZERO)
        amounts[item["work_order_no"]] += quantity * material_price
        raw_qty[item["work_order_no"]] += quantity
    # 人工调整（投料）：填了就覆盖"实领"这一块（粗纱补投/BOM 反推照旧在后面叠加）
    adjustments = load_manual_adjustments(period) if adjustments is None else adjustments
    for order_no, adjustment in adjustments.items():
        manual_amount = adjustment.get("issue_amount")
        manual_qty = adjustment.get("issue_quantity")
        if manual_amount is not None:
            amounts[order_no] = Decimal(manual_amount)
        elif manual_qty is not None:
            manual_qty = Decimal(manual_qty)
            system_qty = raw_qty.get(order_no, ZERO)
            if system_qty > ZERO:
                # 只填了数量：按该工单原投料单价折算金额
                amounts[order_no] = manual_qty * (amounts[order_no] / system_qty)
            else:
                # 原来没有投料 → 取不到单价，必须同时填金额（页面上有提示）
                amounts[order_no] = ZERO
    if roving_plan:
        roving_prices = roving_plan.get("roving_prices", {})
        for order_no, auto_list in roving_plan.get("auto_issues", {}).items():
            for material_code, _name, qty, _uom in auto_list:
                material_price = Decimal(prices.get(material_code, ZERO) or ZERO)
                if not material_price:
                    material_price = Decimal(roving_prices.get(material_code, ZERO) or ZERO)
                amounts[order_no] += Decimal(qty) * material_price
    _backfill_material_by_bom(
        period, prices, amounts, adjustments, backfilled_out=backfilled_out
    )
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

# 情形③「有投料-报工超投料量」的**差额反推**（用户 2026-09-30 拍板落地）：
#   触发：有投料 + 有报工、报工数量 > 投料数量（同量纲车间 ISSUE_REPORT_COMPARABLE）、
#         **无期初结转材料、也无计算层补投**（这两种已经有料在成本里，再反推会重复计料）；
#   反推额 ＝ max(BOM应有（单耗 × 报工数量 × 价） − 本期实领金额, 0)；
#   缺口率 ＝ 缺口 ÷ BOM应有：阈值 100% ＝ **缺多少补多少，一律自动补料**
#   （用户 2026-09-30：「自动补料，不然低于标准成本太多」——原来卡在 50%，像 14.614C-BF20260701
#     缺口 91% 就不补，材料只算到应有的 8.9%、单位成本比标准低 65%；现在不再卡阈值，
#     只保留「无期初、无补投」这个前提，避免重复计料）。
BOM_GAP_BACKFILL_MAX_PCT = Decimal("100")
# 机织重量法在产品：在产品kg = 期初数量kg + 投料kg − 报工米×米重×(1+损耗率)
# 米重 = BOM 单耗之和（kg/米）；损耗率用户拍板 5%~10%，取中值 7.5%（可配置）
# 2026-09-29 修（用户实测工单 32.29G-HF20250815G/HF20261259G 在产品被算成 0）：
#   ① 「投料kg」原来只取本期投料台账（MaterialIssue），台账里没有这张单时投料＝0，
#      重量法必然算出 ≤0 → 在产品 0。实际投料量多数记在「投料进度表」（ERP 累计投料）
#      与「工单明细」的投料数量里 → 本期台账为空时按累计口径兜底：
#      在产品kg = 累计投料kg − 累计完工米 × 米重 × (1+损耗率)。
#   ② 米数法上限原来是「计划 − 本期报工」，跨月工单会把往月已完工的米数又当成在产品
#      → 改为「计划 − 累计完工」（累计完工＝工单明细入库数量，缺失退回本期报工）。
# 2026-09-30 修（用户实测工单 32.29G-HF20260508 在产品仍为 0）：
#   ③ 期初在产品的「数量」列在这批数据里几乎全是空的（177 行只有金额）→ 重量法从 0kg 起步，
#      上期结转的料等于不存在 → 按「期初材料金额 ÷ 本期投料加权单价」折算期初 kg。
#   ④ 工单明细「入库数量」常说不通（超过计划，或按米重折算所需料超过账上累计投料）→
#      这种单不能用它当上限，退回「本期报工」，否则未完工米数会被当成已完工、在产品被压成 0。
JIZHI_WIP_LOSS_RATE = Decimal("0.075")

# bulk_update 限批（2026-09-30 实测）：本机 SQLite（WAL）在事务里执行"超大单条 SQL"会报
#   `OperationalError: unable to open database file` —— 49 行（SQL≈11.7KB）就失败，
#   加 batch_size=20 后 689 行 / 9,944 行都能过。所以全站 bulk_update 统一限批，
#   顺带也是 MySQL/PostgreSQL 上的好习惯（避免单条 SQL 过大）。
BULK_UPDATE_BATCH = 20


def bom_kg_per_meter(period: AccountingPeriod) -> dict:
    """工单编号 → BOM 米重（kg/米 ＝ Σ 子项用量「分子÷分母」）。机织重量法用；没 BOM 的不在表里。"""
    kg_per_m = defaultdict(Decimal)
    for line in period.bom_lines.all():
        denominator = line.qty_denominator or Decimal("1")
        if denominator > 0:
            kg_per_m[line.work_order_no] += line.qty_numerator / denominator
    return kg_per_m


def equivalent_quantity(workshop, order_no, finished_qty, wip_qty, kg_per_m=None):
    """折合产量（用户 2026-09-30 定稿口径，页面/结算/导出共用）：

    * **机织**（有 BOM 米重）：在产品以 **kg（在产纱线重量）** 计，完工也换成 kg
      （完工米 × 米重）→ 折合产量 ＝ 完工kg ＋ 0.5 × 在产品kg；
    * 其它车间（含没有米重的机织单）：完工数量 ＋ 0.5 × 在产品数量（各自单位：件/kg/米）。

    工资＝车间工资池 ×（本工单折合产量 ÷ 车间折合产量合计）。
    """
    finished = Decimal(finished_qty or 0)
    wip = Decimal(wip_qty or 0)
    if workshop == "JiZhi" and kg_per_m:
        kg_m = Decimal(kg_per_m.get(order_no) or 0)
        if kg_m > ZERO:
            finished = finished * kg_m
    return finished + wip * WIP_CONVERSION_RATE


def _backfill_material_by_bom(
    period, prices, amounts, adjustments=None, backfilled_out=None
):
    adjustments = load_manual_adjustments(period) if adjustments is None else adjustments
    reported = set(
        period.work_reports.exclude(work_order_no="")
        .values_list("work_order_no", flat=True)
    )
    # 系统报工数量（按工单合计）：人工调整要能"覆盖"系统值，所以这里必须传系统值。
    # ⚠ 2026-09-30 修 bug：原来把 system_qty 传成 ZERO → **没有人工作调整的工单全被滤掉**
    #   （reported 变空集 → BOM 反推整体不生效，报了工没投料的单材料成本一直是 0）。
    #   实测：311 张有报工的工单被滤成 0 张；修好后 reported 恢复 311 张。
    system_report_qty = defaultdict(Decimal)
    for row in period.work_reports.exclude(work_order_no="").values(
        "work_order_no", "quantity"
    ):
        system_report_qty[row["work_order_no"]] += Decimal(row["quantity"] or ZERO)
    # 人工把报工数量改成 0 的工单视为"没报工"；人工填了 >0 的视为"有报工"
    reported = {
        n
        for n in reported
        if _adjusted_report_qty(n, system_report_qty.get(n, ZERO), adjustments) != ZERO
    }
    reported |= {
        n
        for n, a in adjustments.items()
        if a.get("report_quantity") is not None and Decimal(a["report_quantity"]) > ZERO
    }
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
    # 注意：这里**不能** early-return（2026-09-30）——下面还有「情形③ 差额反推」，
    # 而情形③ 的工单是"有投料"的，本来就不在 missing 里。
    # 报工数量合计（与完工数量口径一致）；人工填了报工数量的以人工值为准。
    # 复用上面那张 system_report_qty，保证"判定有报工"与"取报工数量"是同一个数。
    finish_qty = defaultdict(Decimal, system_report_qty)
    for order_no in adjustments:
        manual = adjustments[order_no].get("report_quantity")
        if manual is not None:
            finish_qty[order_no] = Decimal(manual)
    # 计划数量（机织完工比例）
    plan_qty = {
        o.order_no: Decimal(o.plan_quantity or ZERO)
        for o in period.work_orders.all()
    }
    # BOM 行按工单分组（全部工单：情形④ 反推与情形③ 差额反推都要用）
    bom_by_order = defaultdict(list)
    for line in period.bom_lines.all():
        bom_by_order[line.work_order_no].append(line)
    # 历史单价兜底（月末加权价缺失时）：投料=出库动作，出库单价优先、入库单价兜底
    hist_price = {}
    for hp in period.historical_prices.all():
        if hp.issue_price is not None:
            hist_price[hp.material_code] = Decimal(hp.issue_price)
        elif hp.receipt_price is not None:
            hist_price[hp.material_code] = Decimal(hp.receipt_price)
    for order_no in missing:
        qty = finish_qty.get(order_no, ZERO)
        if qty <= ZERO:
            continue
        ws = dept_map.get(order_no)
        before_amount = amounts.get(order_no, ZERO)
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
        # 记下"这笔材料是反推补的"：结算层会把反推进来的料**全部**算给本期完工
        # （这些工单没有投料台账、也没有期初材料，料本来就是为了这些报工量才补的）
        if backfilled_out is not None and amounts.get(order_no, ZERO) > before_amount:
            backfilled_out.add(order_no)

    # ---- 情形③：有投料、报工超投料量 → **差额反推**补足材料成本（用户 2026-09-30）----
    #   只对「无期初结转材料 + 无计算层补投 + 同量纲车间（报工数与投料数可比）」的单做，
    #   且缺口率 ≤ BOM_GAP_BACKFILL_MAX_PCT 才自动反推（超过阈值只标记，见页面「情形」提示）。
    issue_qty_by_order = defaultdict(Decimal)
    for row in period.material_issues.exclude(work_order_no="").values(
        "work_order_no", "actual_issue_quantity"
    ):
        issue_qty_by_order[row["work_order_no"]] += Decimal(
            row["actual_issue_quantity"] or ZERO
        )
    for order_no, adjustment in adjustments.items():
        if adjustment.get("issue_quantity") is not None:
            issue_qty_by_order[order_no] = Decimal(adjustment["issue_quantity"])

    missing_set = set(missing)
    for order_no in sorted(set(issue_qty_by_order) - opening - roving - missing_set):
        ws = dept_map.get(order_no)
        if ws not in ISSUE_REPORT_COMPARABLE:
            continue  # 机织的量纲不同（报工米 vs 投料kg），不在这一类
        qty = finish_qty.get(order_no, ZERO)
        issued_qty = issue_qty_by_order.get(order_no, ZERO)
        if qty <= ZERO or issued_qty <= ZERO or qty <= issued_qty:
            continue  # 不是"报工超投料量"
        need = ZERO
        for line in bom_by_order.get(order_no, []):
            sub_qty = Decimal(line.qty) * qty
            if sub_qty <= ZERO:
                continue
            mat_price = Decimal(prices.get(line.sub_item_code, ZERO) or ZERO)
            if not mat_price:
                mat_price = hist_price.get(line.sub_item_code, ZERO)
            if not mat_price:
                continue
            need += sub_qty * mat_price
        if need <= ZERO:
            continue
        have = amounts.get(order_no, ZERO)
        gap = need - have
        if gap <= ZERO:
            continue
        if gap / need * Decimal("100") > BOM_GAP_BACKFILL_MAX_PCT:
            continue  # 缺口太大 → 不自动反推，页面标"需人工确认"
        amounts[order_no] = have + gap
        if backfilled_out is not None:
            backfilled_out.add(order_no)


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


def _roving_output_qty(period: AccountingPeriod, order_no: str):
    """粗纱单的产出量（用户 2026-09-30 定稿）：**按实际投料量折算**。

    粗纱是"料变料"，投入≈产出（实测这几张单 BOM 单耗合计正好＝1.0）；原来的"按计划数量"
    会把单位成本算高（BC-16.0419ST-20260754：投料 3,314 而计划 2,800 → 单位成本 49.43 vs 31.98）。
    取数顺序：工单明细「投料数量」→ 投料台账合计 → 投料进度表；都没有才返回 None（由调用方退回计划）。
    """
    detail = period.work_order_details.filter(order_no=order_no).first()
    if detail is not None and Decimal(detail.issue_quantity or 0) > ZERO:
        return Decimal(detail.issue_quantity)
    total = (
        period.material_issues.filter(work_order_no=order_no)
        .aggregate(total=Sum("actual_issue_quantity"))["total"]
        or ZERO
    )
    if Decimal(total) > ZERO:
        return Decimal(total)
    total = (
        period.issue_progress.filter(work_order_no=order_no)
        .aggregate(total=Sum("actual_issue_quantity"))["total"]
        or ZERO
    )
    return Decimal(total) if Decimal(total) > ZERO else None


def build_roving_plan(period: AccountingPeriod, prices=None, adjustments=None):
    """粗纱链计算层计划（仅计算，不写任何库存流水/投料记录）：

    1) BC-开头粗纱工单：**用户 2026-10-16：不再"视为完工"**（完工数量一律＝实际报工，
       不拿计划数量当完工）；没有报工就不算完工。材料仍按 BOM 自动补原料投料（反推只用于材料）；
    2) 有BOM粗纱子项、未实际投该粗纱且已有报工数量的下游工单：
       计算层补粗纱投料（数量=BOM单耗×报工数量）。
    3) 粗纱计算单价 = 粗纱工单BOM原料成本÷计划数量（供下游补投粗纱计价）。

    返回 {"finished": {工单号: 数量}, "auto_issues": {工单号: [(物料, 名称, 数量, 单位)...]},
           "roving_prices": {粗纱物料编码: 单价}}。
    """
    from decimal import Decimal as _D

    prices = prices or {}
    adjustments = load_manual_adjustments(period) if adjustments is None else adjustments
    plan = {"finished": {}, "auto_issues": defaultdict(list), "roving_prices": {}}

    reported = {
        n
        for n in (
            period.work_reports.exclude(work_order_no="")
            .values_list("work_order_no", flat=True)
        )
        if _adjusted_report_qty(n, ZERO, adjustments) != ZERO
    }
    # 人工填了报工（>0）的工单也算"有报工"（这样粗纱工单不会被按计划自动完工）
    reported |= {
        n
        for n, a in adjustments.items()
        if a.get("report_quantity") is not None and Decimal(a["report_quantity"]) > ZERO
    }
    receipt_orders = set(
        period.inventory_transactions.filter(business_type__in=FINISHED_RECEIPT_TYPES)
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
        # 产出量（用户 2026-09-30）：**实际投料量 ÷ BOM 单耗合计**（单耗合计＝1 时即投料量）；
        # 没有投料记录才退回计划数量。
        qty_sum = sum((_D(line.qty or 0) for line in bom_subs.get(wo.order_no, [])), ZERO)
        issued_qty = _roving_output_qty(period, wo.order_no)
        output_qty = (
            issued_qty / qty_sum
            if issued_qty and qty_sum > ZERO
            else plan_qty
        )
        if output_qty <= ZERO:
            continue
        # 用户 2026-10-16：不再把反推产出量当完工数量——完工数量一律=实际报工；
        # 反推只用于材料成本 / 粗纱单价（roving_prices 不变，供下游补投计价）。
        # plan["finished"][wo.order_no] = output_qty
        material_cost = ZERO
        # 这单自己已经实际投过料 → **不再按 BOM 补投**，否则会和台账重复计一遍料
        # （用户 2026-09-30 实测 BC-16.0419ST-20260754：台账 3,314kg＝57,445.01，
        #   粗纱链又补了 3,314kg → 材料翻倍成 114,890.02 ✗）
        already_issued = bool(issued_qty and issued_qty > ZERO)
        for line in bom_subs.get(wo.order_no, []):
            qty = _D(line.qty) * output_qty
            if qty <= ZERO:
                continue
            if not already_issued:
                plan["auto_issues"][wo.order_no].append(
                    (line.sub_item_code, line.sub_item_name, qty, line.uom or "kg")
                )
            material_cost += qty * _D(prices.get(line.sub_item_code, ZERO) or ZERO)
        plan["roving_prices"][material_code] = (
            material_cost / output_qty
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
        finished = _adjusted_report_qty(
            order_no,
            (
                period.work_reports.filter(work_order_no=order_no)
                .aggregate(total=Sum("quantity"))["total"]
                or ZERO
            ),
            adjustments,
        )
        if not finished:
            continue
        plan["auto_issues"][order_no].append(
            (roving_code, line.sub_item_name, _D(line.qty) * _D(finished), line.uom or "kg")
        )
    return plan


def _finished_by_order(period: AccountingPeriod, roving_plan=None, adjustments=None):
    """按工单汇总报工数量（完工数量）。

    浆粕车间无报工记录：完工数量以库存流水"报工"入库数量为准
    （用户每月上传浆粕成品流水（业务类型=生产入库），单据编号自动关联当月浆粕工单）。
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
        period.inventory_transactions.filter(business_type__in=FINISHED_RECEIPT_TYPES)
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
    # 人工调整（报工）：放在最后覆盖，避免被"入库流水兜底 / 粗纱按计划自动完工"重新盖掉
    adjustments = load_manual_adjustments(period) if adjustments is None else adjustments
    for order_no, adjustment in adjustments.items():
        manual = adjustment.get("report_quantity")
        if manual is not None:
            finished[order_no] = Decimal(manual)
    return finished


def attach_finish_progress(items, details=None):
    """给工单成本行挂上「本期完工 / 累计完工 / 本期完工率 / 累计完工率」。

    用户 2026-09-24 定稿口径（原来是「本期=累计」的复制写法，两对列完全重复，已改）：

      本期完工     ＝ 本月报工数量（＝工单成本结果的 finished_quantity，成本口径的完工数量）
      累计完工     ＝ 「工单明细」（数据导入 → 工单明细上传）里的**入库数量**，即 ERP 的跨月累计；
                     该工单没有工单明细、或明细里没填这一列时，退回本月报工数量
      本期完工率   ＝ 本期完工 ÷ 计划数量 × 100
      累计完工率   ＝ 累计完工 ÷ 计划数量 × 100

    与平台其它地方的完工率同口径（视同完工判定 MIN_FINISH_RATE_PCT、机织 BOM 反推
    min(报工/计划, 1) 用的都是「完工÷计划」）。
    计划数量为 0（或空）时两个比率留空（页面显示 —）；完工 > 计划（超产）原样显示、不封顶。
    """
    details = details or {}
    hundred = Decimal("100")
    for item in items:
        order_no = item.work_order.order_no
        plan_qty = Decimal(item.plan_quantity or 0)
        current = Decimal(item.finished_quantity or 0)
        detail = details.get(order_no)
        receipt = getattr(detail, "receipt_quantity", None) if detail is not None else None
        accum = Decimal(receipt) if receipt is not None else current
        item.current_finished = current
        item.accum_finished = accum
        item.current_finish_rate_pct = (
            (current / plan_qty * hundred) if plan_qty else None
        )
        item.finish_rate_pct = (accum / plan_qty * hundred) if plan_qty else None
    return items


def finish_rate_text(value):
    """完工率的导出文本：两位小数、半进位（与页面 `floatformat:2` 完全一致）、0/空 → —。

    页面显示走 Django 的 floatformat（四舍五入），导出若直接用 f"{value:.2f}" 会按
    ROUND_HALF_EVEN 处理（129.565 → 129.56 而页面是 129.57）。这里统一成同一口径，
    保证页面 / 工单成本导出 / 结算工作簿三处对同一张工单显示同一个数。
    """
    if value is None or value == "":
        return "—"
    try:
        number = Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError):
        return "—"
    if number == 0:
        return "—"
    rounded = number.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return "—" if rounded == 0 else "%s%%" % rounded


# ---------------------------------------------------------------------------
# 「投料报工情形速查」（用户 2026-09-24）
#   把《投料报工情形_核算方法对照手册》里的几种情况，落成"每张工单一行"的表：
#   一眼看出这张单走了哪条核算路径、材料成本由什么构成、要不要人工介入。
#   判定用的口径（投料计价、BOM 应有、米重、粗纱链）都与核算层同源，不另起一套规则。
# ---------------------------------------------------------------------------
# 「投料报工情形速查」的情形（用户 2026-09-24 第五次确认）：
#   ① 粗纱-有工单无报工：粗纱本来就不会报工，其成本按 BOM 反推成下游纱线的材料成本（无需处理）
#   ② 在产长期挂账：**在产品 > 0 且计划完工已过 3 个月以上**（用户 2026-09-27 改口径：
#      不再要求"本月有投料"——旧口径下这张表恒为 0，真正挂着的单全落在「正常」里）
#   ③ 有投料-报工超投料量（材料成本该按实际报工数量反推，按实际投料算会偏低）
#   ④ 无投料-有报工（材料成本按 BOM 反推）
#   ⚠️ **「无投料」＝从来没有投过料**（用户 2026-09-27）：本期投料台账、以前月份的投料
#      （投料进度表／工单明细的投料数量）、期初在产品的材料，三者全为 0 才算；
#      只要以前投过料，就不算「无投料-有报工」，归「— 正常」。
#   ⚠️ **原⑤「有投料-部分报工（超3个月未报完）」已并入 ②**（用户 2026-09-27：5 类 → 4 类）：
#      两者本来就是"计划完工超 3 个月"的同一大类，⑤ 只是其中"本月还在投料报文"的细分
#      （7 月实测两条交集 0 张、⑤ 恒为 0）；合并后 ② 同时吃「在产品挂着」和「本月投料但没报完」两种触发。
# 其余归「— 正常」。不再单列：投料不足、投料偏多/疑串料、机织在产品归零、无投料-无报工
#   （上月已投料、本月没投料没报工的单，在产品继续挂着，属正常）。
CASE_ROVING_PLAN = "① 粗纱-有工单无报工（成本按BOM反推）"
# 用户 2026-09-27：② 名称恢复成「有投料-无报工」，括号里保持上一轮定下的文案不变；
# 触发条件＝计划完工已过 3 个月，且「在产品 > 0」或「本月投料了但没报完」（原⑤已并入）
CASE_NO_REPORT = "② 有投料-无报工（在产长期挂账超计划完工3个月以上）"
CASE_REPORT_EXCEED = "③ 有投料-报工超投料量"
CASE_REPORT_ONLY = "④ 无投料-有报工（按BOM反推）"
# 原⑤「有投料-部分报工（超3个月未报完）」已并入 ②（用户 2026-09-27），不再单独列一类
# 明细表里"没有需要关注的点"的行
CASE_NORMAL = "— 正常"
# ⑥ BOM 子项缺单价的工单（用户 2026-09-30：整理出来好补单价；缺价那一项只能算 0 → 成本偏低）
CASE_NO_PRICE = "⑥ BOM 子项缺单价的工单"
# 小结里"其余"那一行（用户 2026-09-27：改成「正常投料生产工单」，并加序号 ⑤——
# 序号接在 4 类情形 ①~④ 之后，所以小结是 5 行 + 合计）
CASE_OTHER = "⑤ 正常投料生产工单"
CASE_CHOICES = (
    CASE_ROVING_PLAN, CASE_NO_REPORT, CASE_REPORT_EXCEED,
    CASE_REPORT_ONLY,
)
# 长期在产的判定阈值（月）：计划完工时间（没有则计划开工）距核算月 ≥ 该值才算"长挂"
LONG_PENDING_MONTHS = 3
# 投料与报工"量纲可比"的车间（机织是米、浆粕无工单，不能直接比数量）
ISSUE_REPORT_COMPARABLE = ("RanZheng", "tefang", "BaoFu", "ZhenZhi")
# 每种情形的「建议处理方法」（用户 2026-09-24：一句话，说清"怎么算/怎么办"）
CASE_SUGGESTIONS = {
    CASE_ROVING_PLAN: "不用处理：粗纱本来就不报工",
    CASE_NO_REPORT: (
        # 原⑤已并入本条（用户 2026-09-27）：所以「在产挂着」和「本月投料但没报完」两种情况都看这条
        "要人工核查：在产挂了 3 个月以上（本月还在投料、但没报完的也算）"
        "——料很可能被别的单用了，到「工单投料」页改到真正用料的工单"
    ),
    CASE_REPORT_EXCEED: (
        "已经自动补料：报工比投料多，缺的按 BOM 补进完工（不用人工处理）"
    ),
    CASE_REPORT_ONLY: "从来没投过料（本月、以前、上月结转都没有），料只能按报工量倒推",
    CASE_NO_PRICE: (
        "这些单的 BOM 有物料没单价（月末加权价、历史单价都没有）"
        " → 到「数据导入 → 历史单价」补价，补完重新结转即可"
    ),
    CASE_NORMAL: "不用处理",
    CASE_OTHER: "正常单，不用处理",
}

# 每种情形「在产品 / 产成品成本怎么算」（用户 2026-09-30：**放进「建议处理方法」一列**，不单开列；
#   大白话、只讲要点，文案必须与结算层实际算法一致，逐条用 2026-07 结转后的数据核对过）
_CASE_BASIS_NORMAL = (
    "料＝实际投料（本月没投料但有报工的按 BOM 倒推）＋上月结转，"
    "再按完工/在产分：机织看耗料公斤比、其它看数量比；费用跟料走，"
    "工资按折合产量（完工＋半个在产品）分\n"
    "完工数量＝报工；在产品数量：机织＝剩余纱线（kg）、纺纱＝投料−完工、其它＝计划−完工"
)
CASE_COST_BASIS = {
    CASE_ROVING_PLAN: (
        "完工按计划数量算，在产品＝0\n"
        "料按 BOM 反推出来，跟下游纱线/成品一起结转"
    ),
    CASE_NO_REPORT: (
        "本月没报工 → 料（本月投料＋上月结转）全留在在产品\n"
        "本月有报工的，按报工那部分算到产成品"
    ),
    CASE_REPORT_EXCEED: (
        "产成品＝报工数量\n"
        "在产品＝0（料还不够报工用，没有料留给在产品）"
    ),
    CASE_REPORT_ONLY: (
        "产成品：按 BOM 单耗×报工数量算料钱，全部算产成品\n"
        "在产品＝0（没投过料，没料可留）"
    ),
    CASE_NO_PRICE: (
        "缺价的那一项，实领和按 BOM 反推都只能算 0\n"
        "所以完工成本、在产品成本都会偏低；补上单价、重新结转就恢复正常"
    ),
    CASE_NORMAL: _CASE_BASIS_NORMAL,
    CASE_OTHER: _CASE_BASIS_NORMAL,
}


def _order_pending_months(order, period):
    """工单"挂了多久"（月）：优先用计划完工时间，没有则用计划开工时间；都没有返回 None。

    用户 2026-09-24：判断"长期在产/长期未报工"用这个月数，
    当月（以及 3 个月以内）的在产属正常，不算问题。

    注意：工单主档的 created_at 是**文件导入时间**（不是 ERP 的业务时间），不能用；
    planned_start / planned_finish 才是 ERP 的日期。
    """
    from datetime import date as _date

    month = period.month
    month = _date(month.year, month.month, 1)
    for value in (getattr(order, "planned_finish", None), getattr(order, "planned_start", None)):
        if not value:
            continue
        stamp = value.date() if hasattr(value, "date") else value
        return (month.year - stamp.year) * 12 + (month.month - stamp.month)
    return None


def _issue_prices(period):
    """投料计价口径（与 _material_by_order 完全一致）：月末加权价 → 历史出库价 → 历史入库价 → 0。"""
    prices = {
        item.material_code: Decimal(item.weighted_average_price or 0)
        for item in period.inventory_valuations.all()
    }
    for hp in period.historical_prices.all():
        if hp.issue_price is not None and not prices.get(hp.material_code):
            prices[hp.material_code] = Decimal(hp.issue_price)
        elif hp.receipt_price is not None and not prices.get(hp.material_code):
            prices[hp.material_code] = Decimal(hp.receipt_price)
    return prices


def issue_unit_price(issue_qty_by_pair, material_price, order_no):
    """工单的投料单价＝**按实际投料**的加权平均单价（保留 2 位小数）。

    ＝ Σ(投料数量 × 该物料月末加权价) ÷ Σ投料数量（只用有价的物料）。
    本单没有投料、或投料都没有价格时返回 None（由调用方回退到 BOM 口径）。
    """
    qty_total = ZERO
    amt_total = ZERO
    for (wo_no, mat_code), qty in issue_qty_by_pair.items():
        if wo_no != order_no:
            continue
        price = material_price.get(mat_code, ZERO)
        if price <= ZERO:
            continue
        qty_total += qty
        amt_total += qty * price
    if qty_total <= ZERO:
        return None
    return (amt_total / qty_total).quantize(CENT, rounding=ROUND_HALF_UP)


def bom_unit_price(period, order_no, material_price, bom_lines=None):
    """按 **BOM 表**折算的每单位料钱（用户 2026-09-27 定；页面按 2 位小数显示）。

    用途：本单本期（或历史上）没有任何投料记录时，「投料单价」原来显示为 0（页面显示 —），
    现在改成**按 BOM 表的子项用量折算**，让页面/导出的这一列总有参考值：
    ＝ Σ(BOM 子项用量 × 该子项月末加权价) ÷ Σ(BOM 子项用量)。

    取价口径与投料完全一致（月末加权价优先、历史价兜底，见 `_issue_prices`）；
    子项用量取 BOM 的「子项用量（分子÷分母）」；取不到价的子项跳过。
    整单一个都没取到价时返回 None（调用方保留 0／显示 —）。
    """
    if bom_lines is None:
        bom_lines = period.bom_lines.filter(work_order_no=order_no)
    qty_total = ZERO
    amt_total = ZERO
    for line in bom_lines:
        price = material_price.get(line.sub_item_code, ZERO)
        if price <= ZERO:
            continue
        qty = Decimal(getattr(line, "qty", ZERO) or ZERO)
        if qty <= ZERO:
            continue
        qty_total += qty
        amt_total += qty * price
    if qty_total <= ZERO:
        return None
    # 返回**不四舍五入**的值：页面按 2 位小数显示，悬停提示里能看到精确值
    # （实测 TDE025-BF20260720：BOM 折算价 12.9991，页面显示 13.00）
    return amt_total / qty_total


def unit_cost_diff_reason(
    item,
    standard_unit_material,
    missing_codes,
    threshold_pct=Decimal("30"),
    short_items=None,
    over_items=None,
    has_issue_ledger=True,
):
    """单位成本差异原因（用户 2026-09-30）：**只有 |差异率| ≥ 30% 才给**，一句话说清真实原因。

    用户要的是"为什么差"，不是把成本补平 —— 例如 `31.11-YT-1456013` 单位成本明显低于标准，
    真实原因是 BOM 两种料只投了一种，就写「少投：11.716…」。
    用户追加：**把少投/超投的是哪种原料、差多少数量也写出来**（简写）。

    入参：
      item                 —— 已挂好 standard_unit_cost / unit_cost_diff_pct 等字段的结果行
      standard_unit_material —— 标准料钱（Σ 单耗×价，元/单位），用于折算"应有材料"
      missing_codes        —— 该单 BOM 里**一点都没投**的子项 [(编码, 名称), …]
      short_items          —— 投了但不够的子项 [(编码, 名称, 差数量, 单位), …]
      over_items           —— 投多了的子项 [(编码, 名称, 多数量, 单位), …]
      has_issue_ledger     —— 该单有没有投料台账（没有时材料来自期初结转/补投，原因里写明）
    返回：原因字符串（不需要说明时返回 ""）。
    """
    pct = getattr(item, "unit_cost_diff_pct", None)
    if pct is None or abs(Decimal(pct)) < threshold_pct:
        return ""
    finish = Decimal(getattr(item, "finished_quantity", None) or 0)
    plan = Decimal(getattr(item, "plan_quantity", None) or 0)
    material = Decimal(getattr(item, "direct_material", None) or 0)
    wip = Decimal(getattr(item, "wip_quantity", None) or 0)
    accum = Decimal(getattr(item, "accum_finished", None) or 0)
    unit_material = Decimal(standard_unit_material or 0)
    need = unit_material * finish if unit_material > 0 else ZERO

    def _fmt(value):
        return "{:,.2f}".format(Decimal(value))

    def _qty(value):
        return "{:,.2f}".format(Decimal(value)).rstrip("0").rstrip(".")

    def _items(items, limit=2):
        # 用户 2026-09-30：数量用「（）」包起来，物料和数量视觉上分得清（如 11.716（139.49千克））
        parts = [
            "%s（%s%s）" % (code, _qty(qty), (uom or ""))
            for code, _name, qty, uom in (items or [])[:limit]
        ]
        more = " 等 %d 种" % len(items) if items and len(items) > limit else ""
        return "、".join(parts) + more

    if Decimal(pct) < 0:  # ---------- 低于标准 ----------
        if missing_codes:
            shown = "、".join(
                "%s（%s）" % (code, (name or "")[:12]) for code, name in missing_codes[:2]
            )
            more = "等 %d 种" % len(missing_codes) if len(missing_codes) > 2 else ""
            body = "少投料：%s%s" % (shown, more)
            if short_items:
                body += "；%s" % _items(short_items)
            return body
        if short_items:
            return "投料不足：%s" % _items(short_items)
        if material <= ZERO:
            return "无材料成本（未投料或 BOM 子项缺单价）"
        if need > ZERO and material < need * Decimal("0.99"):
            return "投料不足：应有 %s／账上 %s" % (_fmt(need), _fmt(material))
        if accum > ZERO and finish > ZERO and accum > finish:
            return "报工少于入库：完工按报工 %s（入库 %s）" % (_fmt(finish), _fmt(accum))
        if plan > ZERO and finish > plan:
            return "超产 %s%%：费用摊薄" % _fmt((finish / plan - 1) * 100)
        return "材料/费用低于标准"
    # ---------- 高于标准 ----------
    if over_items:
        return "投料超 BOM：%s" % _items(over_items)
    if need > ZERO and material > need * Decimal("1.01"):
        amount = _fmt(material - need)
        if not has_issue_ledger:
            return "投料超 BOM：多 %s（本单没有投料台账，料来自期初结转/计算层补投）" % amount
        return "投料超 BOM：多 %s（超投的料不在 BOM 子项内）" % amount
    if wip <= ZERO and plan > finish > ZERO:
        return "在产品清零：尾料全进完工"
    if unit_material > ZERO and finish > ZERO and material / finish > unit_material * Decimal("1.1"):
        return "单位材料高于标准"
    return "工资/制造费用高于标准"


def build_issue_report_rows(
    period,
    workshop_filter="",
    case_filter="",
    order_filter="",
    product_filter="",
    include_roving_only=False,
):
    """「投料报工情形速查」逐单一行（页面/导出共用，用户 2026-09-24）。

    每行的两列判定：

    * **投料/报工组合**（纯客观）：有投料+有报工 / 有投料+无报工 /
      无投料+有报工 / 无投料+无报工；
    * **情形判定**（可操作分类，取最需要关注的一条，与手册 ①~⑦ 对应）：
      ① 粗纱-无报工按计划完工 → ⑦ 机织-投料不足（在产品归零） →
      ② 有投料-无报工 → ⑥ 有投料-报工超投料量 → ⑤ 无投料-有报工（BOM 反推） →
      ④ 有投料-投料不足 / 投料偏多 → ③ 有投料-有报工（正常） → — 无投料-无报工。

    「材料来源」把材料成本拆成：实领（投料×单价）＋ 上期结转 ＋ 计算层补投（差额，
    即粗纱链补投＋BOM 反推，「提示」里标明是哪一种）。

    **统计口径（用户 2026-09-24 明确）**：默认只列「**当月发生核算成本的工单**」＝本期有
    工单成本结果的工单（与「工单成本」页行数一致）。工单号以 BC- 开头、本期没有报工因而
    "按计划自动完工"的粗纱工单**不产生成本行**（它的材料只用于推算粗纱单价、不进任何成本），
    默认不列入；需要看时把 `include_roving_only=True` 打开（页面上的勾选框）。
    """
    prices = _issue_prices(period)
    adjustments = load_manual_adjustments(period)
    # 与核算层同一套物料编码容错（2026-09-30）：投料记 18.660、价源里是 18.66 时也要取到价，
    # 否则页面上的「实领金额 / 缺口率」会和结算层的材料成本对不上（18.441-1358059 就是这种情况）。
    _name_by_code = {}
    for row in period.inventory_valuations.values("material_code", "material_name"):
        if row["material_name"]:
            _name_by_code.setdefault(row["material_code"], row["material_name"])
    for row in period.historical_prices.values("material_code", "material_name"):
        if row["material_name"]:
            _name_by_code.setdefault(row["material_code"], row["material_name"])

    def price_of(code, name=""):
        value = prices.get(code)
        if value:
            return value
        return _code_variant_price(prices, code, name, _name_by_code) or ZERO

    # 本月工单的 BOM 子项里**取不到价**的（用户 2026-09-30：整理出来好补单价）：
    # 月末加权价、历史出/入库价都没有 → 按 BOM 反推时这项算 0，成本天然偏低。
    unpriced_by_order = defaultdict(list)
    for _line in period.bom_lines.all():
        if not _line.work_order_no:
            continue
        if price_of(_line.sub_item_code, _line.sub_item_name):
            continue
        unpriced_by_order[_line.work_order_no].append(_line.sub_item_code)

    issue_qty = defaultdict(Decimal)
    issue_amount = defaultdict(Decimal)
    for row in (
        period.material_issues.exclude(work_order_no="")
        .exclude(material_code="")
        .values("work_order_no", "material_code", "material_name", "actual_issue_quantity")
    ):
        qty = Decimal(row["actual_issue_quantity"] or ZERO)
        order_no = row["work_order_no"]
        issue_qty[order_no] += qty
        issue_amount[order_no] += qty * price_of(row["material_code"], row["material_name"])
    # 人工调整（投料）：与核算层同一口径——填了就以人工值为准（覆盖实领）
    for order_no, adjustment in adjustments.items():
        manual_amount = adjustment.get("issue_amount")
        manual_qty = adjustment.get("issue_quantity")
        if manual_amount is not None:
            issue_amount[order_no] = Decimal(manual_amount)
            if manual_qty is not None:
                issue_qty[order_no] = Decimal(manual_qty)
        elif manual_qty is not None:
            manual_qty = Decimal(manual_qty)
            system_qty = issue_qty.get(order_no, ZERO)
            system_amount = issue_amount.get(order_no, ZERO)
            issue_qty[order_no] = manual_qty
            issue_amount[order_no] = (
                manual_qty * (system_amount / system_qty) if system_qty > ZERO else ZERO
            )

    report_qty = defaultdict(Decimal)
    for row in (
        period.work_reports.exclude(work_order_no="")
        .values("work_order_no")
        .annotate(total=Sum("quantity"))
    ):
        report_qty[row["work_order_no"]] = Decimal(row["total"] or ZERO)
    # 记下"系统值"（页面上做占位提示用，如"系统 1,000.00"）
    raw_issue_qty = dict(issue_qty)
    raw_issue_amount = dict(issue_amount)
    raw_report_qty = dict(report_qty)
    # 累计投料／以前月份的投料（用户 2026-09-27）：
    # 判「无投料-有报工」必须**从来没有投过料**，所以除了本期投料台账，还要看
    #   ① 投料进度表（ERP 的累计投料进度）
    #   ② 工单明细上传的「投料数量」（ERP 侧该工单的投料量）
    #   ③ 期初在产品的材料金额（下面 opening_material）——以前投过、还在产的证据
    # 用户 2026-10-16：累计投料＝**投料台账跨期累计**（本期＋以前各期），
    #   不再用「投料进度表」——它常没同步（会出现"累计 123 < 本期 998.87"的矛盾），
    #   而且判「从来没有投过料」也依赖它，会把情形 ①②④ 算错。
    accum_issue = defaultdict(Decimal)
    for _ap in AccountingPeriod.objects.filter(month__lte=period.month):
        for row in (
            _ap.material_issues.exclude(work_order_no="")
            .values("work_order_no", "actual_issue_quantity")
        ):
            accum_issue[row["work_order_no"]] += Decimal(
                row["actual_issue_quantity"] or ZERO
            )
    detail_issue_qty = {
        d.order_no: Decimal(d.issue_quantity or ZERO)
        for d in period.work_order_details.all()
    }
    _report_all_roving = all_roving_bom_orders(period)
    # 人工调整（报工）：填了就以人工值为准（填 0 ＝视为没有报工）
    for order_no, adjustment in adjustments.items():
        manual = adjustment.get("report_quantity")
        if manual is not None:
            report_qty[order_no] = Decimal(manual)

    # 报工入库流水（完工=0 时用来判断"到底有没有产出"）
    receipt_orders = set(
        period.inventory_transactions.filter(business_type__in=FINISHED_RECEIPT_TYPES)
        .exclude(document_no="")
        .values_list("document_no", flat=True)
    )

    # BOM：按工单分组、米重（分子÷分母之和）、是否含粗纱子项
    bom_by_order = defaultdict(list)
    kg_per_m = defaultdict(Decimal)
    for line in period.bom_lines.all():
        bom_by_order[line.work_order_no].append(line)
        denominator = line.qty_denominator or Decimal("1")
        if denominator > ZERO:
            kg_per_m[line.work_order_no] += line.qty_numerator / denominator
    roving_orders = {
        order_no
        for order_no, lines in bom_by_order.items()
        if any(str(line.sub_item_code).startswith("BC-") for line in lines)
    }
    roving_orders |= set(
        period.material_issues.filter(material_code__startswith="BC-")
        .exclude(work_order_no="")
        .values_list("work_order_no", flat=True)
    )

    opening_material = {
        item.order_no: Decimal(item.material or ZERO)
        for item in period.opening_wip.all()
    }
    results = {
        r.work_order.order_no: r
        for r in period.work_order_cost_results.select_related("work_order")
    }
    masters = {w.order_no: w for w in period.work_orders.all()}

    def bom_need_amount(order_no, workshop, finished_qty, plan_qty):
        """BOM 应有材料金额（与核算层同一口径）：机织＝需求数×完工比例，其余＝单耗×完工数量。"""
        total = ZERO
        if workshop == "JiZhi":
            ratio = (
                clamp_rate(finished_qty / plan_qty) if plan_qty > ZERO else Decimal("1")
            )
            for line in bom_by_order.get(order_no, []):
                total += _jizhi_line_need(line, plan_qty) * ratio * price_of(line.sub_item_code)
            return total
        for line in bom_by_order.get(order_no, []):
            total += Decimal(line.qty) * finished_qty * price_of(line.sub_item_code)
        return total

    rows = []

    def _row(order_no, workshop, order, result=None):
        adjustment = adjustments.get(order_no) or {}
        has_manual = any(
            adjustment.get(key) is not None
            for key in ("issue_quantity", "issue_amount", "report_quantity")
        )
        plan_qty = Decimal(
            (result.plan_quantity if result is not None else order.plan_quantity) or ZERO
        )
        finished_qty = Decimal(result.finished_quantity or ZERO) if result is not None else ZERO
        wip_qty = Decimal(result.wip_quantity or ZERO) if result is not None else ZERO
        direct_material = Decimal(result.direct_material or ZERO) if result is not None else None
        issue_quantity = issue_qty.get(order_no, ZERO)
        issue_amount_value = issue_amount.get(order_no, ZERO)
        report_quantity = report_qty.get(order_no, ZERO)
        has_issue = order_no in issue_qty
        has_report = order_no in report_qty
        # 以前投过料？（进度表 或 工单明细的投料数量 或 期初在产品材料 > 0）
        prior_issue_qty = accum_issue.get(order_no, ZERO)
        prior_detail_qty = detail_issue_qty.get(order_no, ZERO)
        prior_opening_material = opening_material.get(order_no, ZERO)
        issued_before = (
            prior_issue_qty > ZERO
            or prior_detail_qty > ZERO
            or prior_opening_material > ZERO
        )
        # 从来没有投过料（本期没有 + 以前也没有）
        never_issued = not has_issue and not issued_before
        is_roving_order = str(order_no).startswith("BC-")
        roving_link = order_no in roving_orders
        kg_m = kg_per_m.get(order_no, ZERO)
        needed_kg = (
            finished_qty * kg_m * (Decimal("1") + JIZHI_WIP_LOSS_RATE) if kg_m > ZERO else ZERO
        )
        need_amount = bom_need_amount(order_no, workshop, report_quantity, plan_qty)

        # ---- 投料/报工组合（纯客观） ----
        if has_issue and has_report:
            combo = "有投料+有报工"
        elif has_issue:
            combo = "有投料+无报工"
        elif has_report:
            combo = "无投料+有报工"
        else:
            combo = "无投料+无报工"

        # ---- 情形判定（用户 2026-09-27：已合并成 4 类；2026-09-30 加第 6 类"缺单价"） ----
        # "长期在产"＝计划完工时间（没有则计划开工）距核算月 ≥ LONG_PENDING_MONTHS
        pending_months = _order_pending_months(order, period)
        long_pending = pending_months is not None and pending_months >= LONG_PENDING_MONTHS
        if unpriced_by_order.get(order_no):
            # ⑥ BOM 子项缺单价的工单（用户 2026-09-30）：优先标出来，好去补价。
            # 取价口径下这一项只能算 0 → 材料/在产品天然偏低，所以单独成一类。
            case = CASE_NO_PRICE
        elif is_roving_order and not has_report:
            # 粗纱（BC-）本来就不会报工：其成本按 BOM 反推成下游纱线的材料成本，不是异常
            case = CASE_ROVING_PLAN
        elif (
            has_issue and has_report
            and workshop in ISSUE_REPORT_COMPARABLE
            and report_quantity > issue_quantity
        ):
            # ③ 有投料-报工超投料量（材料成本该按实际报工数量反推）
            case = CASE_REPORT_EXCEED
        elif long_pending and (
            # ② 有投料-无报工（在产长期挂账超计划完工3个月以上）——
            # 用户 2026-09-27 改口径 + 合并原⑤：计划完工已过 3 个月，且满足下面任一：
            #   ① 在产品还挂着（不再要求本月有投料；旧口径在真实数据里恒为 0）
            #   ② 本月投料了但没报完（原⑤「有投料-部分报工」的触发，现已并入本条）
            wip_qty > ZERO
            or (
                has_issue and has_report
                and plan_qty > ZERO and report_quantity < plan_qty
            )
        ):
            case = CASE_NO_REPORT
        elif never_issued and has_report:
            # ④「无投料-有报工」＝**纯没有投料**（用户 2026-09-27）：
            # 本期、以前、期初都没有投料记录的单才归这一类；
            # 以前投过料（进度表/工单明细/期初在产品有材料）的归「— 正常」。
            case = CASE_REPORT_ONLY
        else:
            case = CASE_NORMAL

        # ---- 材料来源拆分 ----
        from_issue = issue_amount_value
        from_opening = opening_material.get(order_no, ZERO)
        from_calc = (
            max(direct_material - from_issue - from_opening, ZERO)
            if direct_material is not None
            else ZERO
        )

        # ---- 提示（附加标记） ----
        flags = []
        if plan_qty > ZERO and report_quantity > plan_qty:
            flags.append("报工超计划（超产）")
        if from_opening > ZERO:
            flags.append("含上期结转 %s" % money(from_opening))
        if from_calc > ZERO:
            flags.append(
                "计算层补投 %s（%s）"
                % (
                    money(from_calc),
                    "粗纱链" if (roving_link or is_roving_order) else "BOM 反推",
                )
            )
        if finished_qty > ZERO and direct_material == ZERO:
            flags.append("报了工但材料成本＝0")
        if (
            workshop == "JiZhi"
            and result is not None
            and wip_qty == ZERO
            and plan_qty > finished_qty
        ):
            flags.append("在产品＝0（剩余米成本已进完工）")
        if case == CASE_REPORT_EXCEED and issue_quantity > ZERO:
            flags.append(
                "报工÷投料 = %s 倍" % (report_quantity / issue_quantity).quantize(Decimal("0.01"))
            )
        if case == CASE_REPORT_EXCEED and need_amount > ZERO:
            # 差额反推口径（用户 2026-09-30）：缺口＝BOM应有−本期实领；
            # 只在「无期初结转、无补投」且缺口率 ≤ BOM_GAP_BACKFILL_MAX_PCT 时才自动反推。
            # ⚠ 这里不能用 from_calc 判断"有没有补投"：差额反推补进来的钱本身就算在 from_calc 里，
            #   用它判断会把"已反推"误标成"有补投不反推"（2026-09-30 修）。粗纱链才是真的另有补投。
            gap = need_amount - issue_amount_value
            if gap > ZERO:
                gap_pct = gap / need_amount * Decimal("100")
                if from_opening > ZERO or roving_link or is_roving_order:
                    flags.append(
                        "缺口 %s（%.1f%%）：有期初结转/粗纱补投，不做差额反推"
                        % (money(gap), gap_pct)
                    )
                elif gap_pct <= BOM_GAP_BACKFILL_MAX_PCT:
                    flags.append(
                        "缺口 %s（%.1f%%）：已按 BOM 差额反推补料"
                        % (money(gap), gap_pct)
                    )
                else:
                    flags.append(
                        "缺口 %s（%.1f%%）：缺口率超 %.0f%%，未自动反推，需人工确认"
                        % (money(gap), gap_pct, BOM_GAP_BACKFILL_MAX_PCT)
                    )
        if has_report and not has_issue and issued_before:
            # 用户 2026-09-27：以前投过料、本月没投料有报工 → 不算「无投料-有报工」
            flags.append(
                "以前投过料（累计投料 %s），本月无投料" % money(
                    prior_issue_qty or prior_detail_qty or prior_opening_material
                )
            )
        if case == CASE_NO_REPORT and pending_months is not None:
            # ② 现在同时覆盖「在产品挂着」和「本月投料了但没报完」（原⑤已并入）
            if wip_qty > ZERO:
                flags.append("计划完工已过 %d 个月，在产品仍挂着" % pending_months)
            else:
                flags.append("计划完工已过 %d 个月未报完" % pending_months)
        if result is None:
            flags.append("本期无投料/报工/入库，未生成成本结果")
        _unpriced = sorted(set(unpriced_by_order.get(order_no, [])))
        if _unpriced:
            # 用户 2026-09-30：BOM 子项没单价的要在明细里标出来，好去补价
            flags.append(
                "BOM 缺单价 %d 个：%s" % (len(_unpriced), "、".join(_unpriced[:5]))
            )
        if has_manual:
            flags.insert(0, "已人工调整（投料以人工值为准）")

        rows.append(
            {
                "workshop": workshop,
                "workshop_label": WORKSHOP_LABELS.get(workshop, workshop or "—"),
                "order_no": order_no,
                "product_code": (result.product_code if result is not None else "") or "",
                "product_name": (result.product_name if result is not None else "") or "",
                "combo": combo,
                "case": case,
                "suggestion": CASE_SUGGESTIONS.get(case, ""),
                "plan_qty": plan_qty,
                "issue_qty": issue_quantity,
                "issue_amount": issue_amount_value,
                "report_qty": report_quantity,
                # 人工调整：可编辑列的回填值 + 系统值（占位提示）+ 是否有调整
                "adjustment": adjustment,
                "has_manual": has_manual,
                "manual_issue_qty": adjustment.get("issue_quantity"),
                "manual_issue_amount": adjustment.get("issue_amount"),
                "manual_report_qty": adjustment.get("report_quantity"),
                "raw_issue_qty": raw_issue_qty.get(order_no, ZERO),
                "raw_issue_amount": raw_issue_amount.get(order_no, ZERO),
                "raw_report_qty": raw_report_qty.get(order_no, ZERO),
                "finished_qty": finished_qty,
                "wip_qty": wip_qty,
                "need_amount": need_amount,
                "diff_amount": (issue_amount_value - need_amount) if need_amount > ZERO else None,
                "from_issue": from_issue,
                "from_opening": from_opening,
                "from_calc": from_calc,
                "direct_material": direct_material,
                "finished_cost": (
                    Decimal(result.finished_cost or ZERO) if result is not None else None
                ),
                "wip_cost": Decimal(result.wip_cost or ZERO) if result is not None else None,
                # BOM 子项取不到价的物料（用户 2026-09-30：整理出来好补单价）
                "unpriced_codes": sorted(set(unpriced_by_order.get(order_no, []))),
                "flags": "；".join(flags),
            }
        )

    for order_no, result in results.items():
        workshop = result.work_order.production_dept_code
        _row(order_no, workshop, result.work_order, result)

    # 粗纱工单（BC- 开头）：本期没有报工 → 计算层按计划自动完工，**不产生成本行**。
    # 默认不列入（用户 2026-09-24：本表只要"当月发生核算成本的工单"）；勾选后才补上。
    roving_only = [
        order_no
        for order_no, master in masters.items()
        if order_no not in results and str(order_no).startswith("BC-")
    ]
    if include_roving_only:
        for order_no in roving_only:
            _row(order_no, masters[order_no].production_dept_code, masters[order_no], None)

    if workshop_filter:
        rows = [r for r in rows if r["workshop"] == workshop_filter]
    if case_filter:
        rows = [r for r in rows if r["case"] == case_filter]
    if order_filter:
        needle = order_filter.lower()
        rows = [r for r in rows if needle in r["order_no"].lower()]
    if product_filter:
        needle = product_filter.lower()
        rows = [
            r for r in rows
            if needle in ("%s %s" % (r["product_code"], r["product_name"])).lower()
        ]

    rows.sort(
        key=lambda r: (
            CASE_CHOICES.index(r["case"]) if r["case"] in CASE_CHOICES else 99,
            -(r["direct_material"] or ZERO),
            r["order_no"],
        )
    )
    totals = summarize_issue_rows(rows)
    totals["roving_only_count"] = len(roving_only)
    totals["include_roving_only"] = bool(include_roving_only)
    return rows, totals, summarize_issue_by_case(rows)


def summarize_issue_rows(rows):
    """按当前显示的行汇总（金额列相加；「应有」按有值的行相加）。"""
    return {
        "count": len(rows),
        "issue_amount": sum((r["issue_amount"] for r in rows), ZERO),
        "need_amount": sum((r["need_amount"] for r in rows), ZERO),
        "direct_material": sum((r["direct_material"] or ZERO for r in rows), ZERO),
        "finished_cost": sum((r["finished_cost"] or ZERO for r in rows), ZERO),
        "wip_cost": sum((r["wip_cost"] or ZERO for r in rows), ZERO),
        "from_issue": sum((r["from_issue"] for r in rows), ZERO),
        "from_opening": sum((r["from_opening"] for r in rows), ZERO),
        "from_calc": sum((r["from_calc"] for r in rows), ZERO),
    }


def summarize_issue_by_case(rows):
    """按情形统计（页面小结 + 下拉里显示条数）。"""
    by_case = defaultdict(lambda: {"count": 0, "direct_material": ZERO})
    for row in rows:
        bucket = by_case[row["case"]]
        bucket["count"] += 1
        bucket["direct_material"] += row["direct_material"] or ZERO
    return by_case


#: 物料名 → 物料分类（纱线/面料/染色纤维/原料/坯布…）的关键词规则。
#: 与「工单成本」页「类别」列、完工入库结转凭证的「库存商品—分类」共用同一套。
ORDER_CATEGORY_KEYWORDS = (
    ("面料", ("面料", "坯布", "织布", "摇粒绒", "螺纹", "绒布", "布", "绒")),
    ("纱线", ("纱线", "纱")),
    ("染色纤维", (
        "纤维", "长丝", "丝", "腈纶", "粘胶", "涤纶", "莱赛尔",
        "芳纶", "氨纶", "短纤", "棉纱",
    )),
    ("原料", ("原料", "棉", "麻")),
)

#: 车间兜底分类（产品编码/物料名都判不出来时按车间取）
WORKSHOP_DEFAULT_CATEGORY = {
    "RanZheng": "染色纤维",
    "tefang": "纱线",
    "BaoFu": "纱线",
    "ZhenZhi": "面料",
    "JiZhi": "面料",
    "浆粕车间": "浆粕",   # ＝ PULP_WORKSHOP（常量定义在本文件靠后，这里用字面量）
}


def category_from_name(name):
    """物料名称 → 物料分类（判不出返回空串）。"""
    if not name:
        return ""
    for category, keywords in ORDER_CATEGORY_KEYWORDS:
        if any(keyword in name for keyword in keywords):
            return category
    return ""


def build_order_category_resolver(period, work_order_details=None):
    """返回「工单 → 物料分类」的解析函数（工单成本页「类别」列与完工入库凭证共用）。

    规则（与原来「工单成本」页内联的那段完全一致）：
      产品编码的库存流水分类名 → 产品编码（结算结果里的）→ BOM 主产出物料编码
      → BOM 主产出物料名关键词 → 工单明细（上传）的类别 → 车间默认类别 → 工单类型。
    """
    category_by_code = {}
    for code, category in (
        period.inventory_transactions.exclude(category_name="")
        .order_by("material_code")
        .values_list("material_code", "category_name")
    ):
        category_by_code.setdefault(code, category)
    bom_product = {}
    for line in period.bom_lines.exclude(material_code="").order_by("id"):
        bom_product.setdefault(
            line.work_order_no, (line.material_code, line.material_name)
        )
    details = (
        work_order_details
        if work_order_details is not None
        else {item.order_no: item for item in period.work_order_details.all()}
    )

    def resolve(order_no, product_code="", workshop="", order_type=""):
        bom = bom_product.get(order_no)
        detail = details.get(order_no)
        return (
            category_by_code.get(product_code)
            or (category_by_code.get(bom[0]) if bom and bom[0] else "")
            or category_from_name(bom[1] if bom else "")
            or (detail.category if detail is not None else "")
            or WORKSHOP_DEFAULT_CATEGORY.get(workshop or "", "")
            or order_type
            or ""
        )

    return resolve


def all_roving_bom_orders(period) -> set:
    """BOM 子项**全部是粗纱（BC-）**的工单（至少 1 行 BOM）。

    用户 2026-09-30 口径：
      * 这些工单**没有在产品**（粗纱不报工、不挂在产品）；
      * 「是否用粗纱」这一列也只在**100% 全粗纱**时显示"是"——
        混料的（粗纱＋其它料，如 RAJ031A-HF20260603）显示"否"。
    """
    codes_by_order = defaultdict(set)
    for line in period.bom_lines.all():
        if line.work_order_no:
            codes_by_order[line.work_order_no].add(str(line.sub_item_code))
    return {
        order_no
        for order_no, codes in codes_by_order.items()
        if codes and all(code.startswith("BC-") for code in codes)
    }


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
    # 人工调整（投料/报工）：读一次，传给下面每个取数口（用户 2026-09-24）
    adjustments = load_manual_adjustments(period)
    # 粗纱链计算层计划（仅计算不落库：自动完工/补投/粗纱计算单价）
    roving_plan = build_roving_plan(period, prices, adjustments)
    # 材料是"BOM 反推补进来"的工单：结算层把反推额全部算给本期完工（口径见下方拆分）
    backfilled_orders = set()
    material_by_order = _material_by_order(
        period, prices, roving_plan, adjustments, backfilled_orders
    )
    finished_by_order = _finished_by_order(period, roving_plan, adjustments)

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
    # 人工调整（投料数量）：纺纱"投料−完工"在产品口径要用人工值
    for order_no, adjustment in adjustments.items():
        manual = adjustment.get("issue_quantity")
        if manual is not None:
            issue_qty_by_order[order_no] = Decimal(manual)
    # 用粗纱的工单（BOM子项或投料含 BC-）：粗纱自动补投=按完工反推，无在产品
    roving_orders = set(
        period.bom_lines.filter(sub_item_code__startswith="BC-")
        .values_list("work_order_no", flat=True)
    ) | set(
        period.material_issues.filter(material_code__startswith="BC-")
        .exclude(work_order_no="")
        .values_list("work_order_no", flat=True)
    )
    # BOM 子项**全部是粗纱（BC-）**的工单（用户 2026-09-30 口径）：
    #   这种单没有在产品；**混料**的（粗纱＋其它料，如 RAJ031A-HF20260603）照常算在产品。
    #   原来写成"只要用了粗纱就没有在产品"，把混料单也清零了 → 13.4 万料全压进 32.8 的完工量。
    # BOM 子项**全部是粗纱（BC-）**的工单（用户 2026-09-30 口径）：这种单没有在产品；
    #   混料的（粗纱＋其它料，如 RAJ031A-HF20260603）照常算在产品；「是否用粗纱」也只在
    #   100% 全粗纱时显示"是"（见 all_roving_bom_orders()）。
    all_roving_bom_orders_set = all_roving_bom_orders(period)
    # BOM 行按工单 + 取价表（用于"完工材料按完工入库反推"，用户 2026-09-30）
    _bom_lines_by_order = defaultdict(list)
    for _bom_line in period.bom_lines.all():
        _bom_lines_by_order[_bom_line.work_order_no].append(_bom_line)
    _need_prices = dict(prices or {})
    for _hp in period.historical_prices.all():
        if _hp.issue_price is not None and not _need_prices.get(_hp.material_code):
            _need_prices[_hp.material_code] = Decimal(_hp.issue_price)
        elif _hp.receipt_price is not None and not _need_prices.get(_hp.material_code):
            _need_prices[_hp.material_code] = Decimal(_hp.receipt_price)

    def _bom_need_for_finished(order_no, finished_qty):
        """完工部分的 BOM 应有材料 ＝ Σ(BOM 单耗 × 完工数量 × 价)。

        用于"完工材料按完工入库反推"（用户 2026-09-30）：上月有在产、本月没投料但有报工的单，
        完工该承担的料是按**完工（入库）数量**推出来的，而不是把账上材料按数量比例摊一点。
        """
        total = ZERO
        for _line in _bom_lines_by_order.get(order_no, []):
            _price = _need_prices.get(_line.sub_item_code)
            if not _price:
                continue
            total += Decimal(_line.qty) * finished_qty * Decimal(_price)
        return total

    # 机织反推工单（无投料、无期初、非粗纱链、有报工）：
    # 材料成本 = BOM需求数 × 完工比例 × 单价，该金额即为本期完工材料成本，
    # 无投料/无期初无"投入"台账依据，不确认在产品 → 在产品=0、材料全部进完工
    # （避免"反推金额×完工比例 后再按数量比例拆一次"造成完工材料二次打折）
    _jz_issue_set = set(
        period.material_issues.exclude(work_order_no="")
        .values_list("work_order_no", flat=True)
    )
    # 人工调整：填了投料（数量或金额）算"有投料"；都把 0 填上算"无投料"
    for order_no, adjustment in adjustments.items():
        manual_qty = adjustment.get("issue_quantity")
        manual_amount = adjustment.get("issue_amount")
        has_manual = (
            manual_qty is not None and Decimal(manual_qty) > ZERO
        ) or (
            manual_amount is not None and Decimal(manual_amount) > ZERO
        )
        cleared = (manual_qty is None or Decimal(manual_qty) == ZERO) and (
            manual_amount is None or Decimal(manual_amount) == ZERO
        ) and (manual_qty is not None or manual_amount is not None)
        if has_manual:
            _jz_issue_set.add(order_no)
        elif cleared:
            _jz_issue_set.discard(order_no)
    _jz_open_set = set(
        period.opening_wip.exclude(material=ZERO)
        .values_list("order_no", flat=True)
    )
    _jz_rov_set = set(roving_orders)
    _jz_rov_set |= set(
        period.issue_progress.filter(material_code__startswith="BC-")
        .exclude(work_order_no="")
        .values_list("work_order_no", flat=True)
    )
    _jz_report_set = set(
        period.work_reports.exclude(work_order_no="")
        .values_list("work_order_no", flat=True)
    )
    # 人工调整：报工数量填 0 视为没报工；填了 >0 视为有报工
    _jz_report_set = {
        n for n in _jz_report_set if _adjusted_report_qty(n, ZERO, adjustments) != ZERO
    }
    _jz_report_set |= {
        n
        for n, a in adjustments.items()
        if a.get("report_quantity") is not None and Decimal(a["report_quantity"]) > ZERO
    }
    jizhi_backfill_orders = {
        wo.order_no
        for wo in orders
        if wo.production_dept_code == "JiZhi"
        and wo.order_no in _jz_report_set
        and wo.order_no not in _jz_issue_set
        and wo.order_no not in _jz_open_set
        and wo.order_no not in _jz_rov_set
    }
    # 机织重量法在产品所需的映射：
    # 米重(kg/米) = BOM 单耗之和（分子÷分母），缺失为 0（退回米数法）
    jz_kg_per_m = defaultdict(Decimal)
    for _ln in period.bom_lines.all():
        _d = _ln.qty_denominator or Decimal("1")
        if _d > ZERO:
            jz_kg_per_m[_ln.work_order_no] += _ln.qty_numerator / _d
    # 期初在产品数量（kg）——机织在产品重量法的起点。
    # ⚠️ 2026-09-30 修：这批数据 177 行期初在产品的「数量」列几乎全是空的（只有材料金额，
    #    合计 1,153 万元）。若按 0 kg 起步，上期结转的料等于不存在 → 重量法必然把在产品
    #    算成 0、成本全额压到本期完工（实测 32.29G-HF20260508：期初材料 533,156.37 元，
    #    本期投料台账 4,731.75kg，重量法 0＋4731.75−4974.19≤0 → 在产品 0）。
    #    数量为空时按「期初材料金额 ÷ 该工单本期投料加权单价」折算 kg
    #    （533,156.37 ÷ 75.7798 = 7,035.60kg，与 ERP 的「累计投料−本期投料＝7,035.60」一致）。
    _hist_first = {}
    for _hp in period.historical_prices.all():
        if _hp.issue_price is not None:
            _hist_first.setdefault(_hp.material_code, Decimal(_hp.issue_price))
        elif _hp.receipt_price is not None:
            _hist_first.setdefault(_hp.material_code, Decimal(_hp.receipt_price))
    _ledger_amount = defaultdict(Decimal)
    for _row in (
        period.material_issues.exclude(work_order_no="")
        .exclude(material_code="")
        .values("work_order_no", "material_code", "actual_issue_quantity")
    ):
        _unit = Decimal(prices.get(_row["material_code"], ZERO) or ZERO)
        if not _unit:
            _unit = _hist_first.get(_row["material_code"], ZERO)
        _ledger_amount[_row["work_order_no"]] += (
            Decimal(_row["actual_issue_quantity"] or ZERO) * _unit
        )

    def _opening_kg(order_no):
        """期初在产品数量（kg）：有数量用数量；只有金额时用金额÷本月投料加权单价折算。"""
        item = opening_map.get(order_no)
        if item is None:
            return ZERO
        quantity = Decimal(item.quantity or ZERO)
        material = Decimal(item.material or ZERO)
        if quantity > ZERO or material <= ZERO:
            return quantity
        ledger_qty = issue_qty_by_order.get(order_no, ZERO)
        ledger_amount = _ledger_amount.get(order_no, ZERO)
        if ledger_qty <= ZERO or ledger_amount <= ZERO:
            return ZERO
        return material / (ledger_amount / ledger_qty)

    jz_opening_qty = {
        item.order_no: _opening_kg(item.order_no)
        for item in period.opening_wip.all()
    }
    # 机织重量法的投料兜底来源（2026-09-29）：
    # 投料记录台账（MaterialIssue）只覆盖一部分工单；机织工单的实际投料量多数记在
    # 「投料进度表」（ERP 累计投料）与「工单明细」的投料数量里。若只看台账，
    # 台账里没有的工单投料＝0 → 重量法在产品必然算成 0（实测 32.29G-HF20250815G/HF20261259G）。
    jz_accum_issue = defaultdict(Decimal)
    for row in (
        period.issue_progress.exclude(work_order_no="")
        .values("work_order_no", "actual_issue_quantity")
    ):
        jz_accum_issue[row["work_order_no"]] += Decimal(
            row["actual_issue_quantity"] or ZERO
        )
    jz_detail_issue = {
        d.order_no: Decimal(d.issue_quantity or ZERO)
        for d in period.work_order_details.all()
    }
    # 累计完工（米/件）＝ 工单明细的入库数量（就是工单成本页「累计完工」那一列）
    jz_accum_finished = {
        d.order_no: Decimal(d.receipt_quantity or ZERO)
        for d in period.work_order_details.all()
    }

    def _jz_accum_issued(order_no):
        """机织工单的「累计投料kg」：投料进度表（ERP 累计投料）与工单明细投料数量取大。"""
        return max(
            jz_accum_issue.get(order_no, ZERO),
            jz_detail_issue.get(order_no, ZERO),
        )

    def _wip_qty(wo_obj, finished_qty_val, plan_qty_val):
        """在产品数量口径：
        纺纱车间：**BOM 100% 只有粗纱（BC-）的工单 → 0**（粗纱不报工、不挂在产品）；
                 其它纺纱工单（含"粗纱＋其它料"的混料单）：max(投料kg-完工kg, 0)；
                 本月台账没有这张单的投料 → 累计口径（投料进度表/工单明细的累计投料 − 累计完工），
                 再兜底 max(计划−完工, 0)；
        机织车间（重量法，用户定稿）：在产品＝**在产纱线的重量（kg）**
                 ＝ 期初数量kg + 投料kg − 报工米×米重×(1+损耗率)，**不再折回米**
                 （用户 2026-09-30：不要算米数，要算在产的纱线重量）；
                 本期投料台账（MaterialIssue）为空时，说明这张单的投料没进台账，
                 按**累计口径**兜底：在产品kg = 累计投料kg − 累计完工米×米重×(1+损耗率)
                 （累计投料＝投料进度表/工单明细，累计完工＝工单明细入库数量）；
                 米重缺失（BOM 无单耗）→ 退回 max(计划-完工, 0)（米，并注明口径）；
        其他车间 → max(计划-完工, 0)。
        """
        ws = wo_obj.production_dept_code
        order_no = wo_obj.order_no
        loss = Decimal("1") + JIZHI_WIP_LOSS_RATE
        if ws == "tefang":
            # 用户 2026-09-30 口径：**BOM 100% 只有粗纱的工单没有在产品**（真·粗纱工单 BC- 也一样），
            # 其它工单一律照常算在产品（混料单 RAJ031A-HF20260603 之前被误清零就是这个问题）。
            if str(order_no).upper().startswith("BC-") or order_no in all_roving_bom_orders_set:
                return ZERO
            issued_now = issue_qty_by_order.get(order_no, ZERO)
            if issued_now > ZERO:
                return max(issued_now - finished_qty_val, ZERO)
            # 本月台账没有这张单的投料 → 累计口径兜底（与机织同一套映射），
            # 避免"本月没投料 → 在产品必然 0"（跨月工单的料在上月投、本月只报工）。
            accum_issued = _jz_accum_issued(order_no)
            if accum_issued > ZERO:
                accum_finished = jz_accum_finished.get(order_no) or finished_qty_val
                return max(accum_issued - accum_finished, ZERO)
            return max(plan_qty_val - finished_qty_val, ZERO)
        if ws == "JiZhi":
            kg_m = jz_kg_per_m.get(order_no, ZERO)
            if kg_m > ZERO:
                opening_kg = jz_opening_qty.get(order_no, ZERO)
                issued_now = issue_qty_by_order.get(order_no, ZERO)
                if issued_now > ZERO:
                    # 本期口径（原有，行为不变；单位由米改为 kg）
                    consumed_kg = finished_qty_val * kg_m * loss
                    return max(opening_kg + issued_now - consumed_kg, ZERO)
                accum_issued = _jz_accum_issued(order_no)
                if accum_issued > ZERO:
                    # 累计口径兜底：投料与完工都用累计值，不再叠加期初数量（累计投料已含往月）
                    accum_finished = (
                        jz_accum_finished.get(order_no) or finished_qty_val
                    )
                    consumed_kg = accum_finished * kg_m * loss
                    return max(accum_issued - consumed_kg, ZERO)
                # 本期、累计都没有投料数据 → 维持原口径（只有期初数量能形成在产品）
                consumed_kg = finished_qty_val * kg_m * loss
                return max(opening_kg - consumed_kg, ZERO)
            return max(plan_qty_val - finished_qty_val, ZERO)
        return max(plan_qty_val - finished_qty_val, ZERO)

    def _jz_accum_finished_ok(order_no, accum_finished, plan_qty_val):
        """工单明细「入库数量」能不能当累计完工用（2026-09-30）。

        这批数据里它常常说不通：12 张机织单的入库数量按米重折算所需料都超过账上累计投料
        （如 32.29G-HF20260508：33,220.30 米需要 15,427.51kg，账上只投了 11,767.35kg），
        还有多张超过计划数量（32.29G-HF20260508 是计划的 132.88%）。
        说不通时不能拿它当上限依据（否则未完工的米数会被当成已完工，在产品被压成 0）。
        """
        if plan_qty_val > ZERO and accum_finished > plan_qty_val:
            return False
        kg_m = jz_kg_per_m.get(order_no, ZERO)
        accum_issued = _jz_accum_issued(order_no)
        if kg_m > ZERO and accum_issued > ZERO:
            need_kg = accum_finished * kg_m * (Decimal("1") + JIZHI_WIP_LOSS_RATE)
            if need_kg > accum_issued:
                return False
        return True

    def _jz_cap_finished(order_no, finished_qty_val, plan_qty_val):
        """机织米数法上限用的「已完工数量」——与重量法同一口径：
        本期有投料台账 → 本期报工；本期台账为空（走累计口径）→ 累计完工（工单明细入库数量），
        但该数说不通时（超过计划、或所需料超过账上累计投料）退回本期报工。
        """
        if issue_qty_by_order.get(order_no, ZERO) > ZERO:
            return finished_qty_val
        if _jz_accum_issued(order_no) > ZERO:
            accum_finished = jz_accum_finished.get(order_no) or finished_qty_val
            if _jz_accum_finished_ok(order_no, accum_finished, plan_qty_val):
                return accum_finished
        return finished_qty_val

    _wip_cache = {}

    def _final_wip_qty(wo_obj):
        """最终在产品数量：原始口径 → 机织米数法上限 → 机织反推工单清零 → 视同完工清零。

        完工/在产品拆分（材料、制造费用、工资）与工资的折合数量分摊共用这一个值
        （原来工资那边把 _wip_qty 重算了一遍，既没加上限也没算视同完工，两边口径不一致）。
        """
        order_no = wo_obj.order_no
        if order_no in _wip_cache:
            return _wip_cache[order_no]
        finished_qty_val = finished_by_order.get(order_no, ZERO)
        plan_qty_val = Decimal(wo_obj.plan_quantity or ZERO)
        wip = _wip_qty(wo_obj, finished_qty_val, plan_qty_val)
        if (
            wo_obj.production_dept_code == "JiZhi"
            and order_no not in jizhi_backfill_orders
        ):
            # 机织重量法在产品 ≤ 米数法上限（计划 − 已完工）：
            # 投料过剩/米重偏小时以计划为界。已完工口径与重量法一致（本期台账→本期报工、
            # 累计口径→累计完工）：跨月工单不能再拿本期报工当上限，否则往月已完工的米数
            # 会被当成在产品（2026-09-29 修）。上限原来是米，2026-09-30 起在产品以 kg 计
            # → 有米重时把上限也乘米重换成 kg（没有米重才留在米口径）。
            cap_finished = _jz_cap_finished(order_no, finished_qty_val, plan_qty_val)
            cap = max(plan_qty_val - cap_finished, ZERO)
            _kg_m = jz_kg_per_m.get(order_no, ZERO)
            if _kg_m > ZERO:
                cap = cap * _kg_m
            wip = min(wip, cap)
        if order_no in jizhi_backfill_orders:
            # 机织反推工单：无投料无期初，反推金额=本期完工材料成本，无在产品
            wip = ZERO
        if order_no in backfilled_orders:
            # 材料是"按报工量反推补进来"的（情形④ 无投料-有报工 / 情形③ 差额反推）：
            # 那笔料刚好覆盖**报工产出**，没有料留给在产 → **不确认在产品**（在产品＝0）
            # 用户 2026-09-30 实测：④ 有 2 张挂出 2.60 / 11.10 的在产品、却一点材料都没有，
            # 口径上说不通（"无投料-有报工怎么可能有在产品"），故清零。
            wip = ZERO
        # 视同完工：① 工单有完工时间（工单明细实际完工时间，或报工记录最大报工时间兜底，与页面口径一致）
        #          ② 完成状态为已完工（超时完成/按时完成/已完成）；完工率达标 → 尾料按损耗处理，
        #             在产品清零、成本并入完工（避免"已结束工单在产品永远挂账结转"）
        detail = work_order_detail_map.get(order_no)
        detail_finished = report_time_agg.get(order_no) is not None or (
            detail is not None
            and (
                bool(detail.actual_finish)
                or (detail.completion_status in FINISHED_WORKORDER_STATUSES)
            )
        )
        if (
            wo_obj.production_dept_code in SMALL_ORDER_WORKSHOPS
            and detail_finished
            and plan_qty_val > ZERO
            and finished_qty_val > ZERO
            and wip > ZERO
            and (finished_qty_val / plan_qty_val * Decimal("100"))
            >= MIN_FINISH_RATE_PCT
        ):
            # 用户 2026-09-30 补充：**明显超投料的单不清零**——超出的料是留给在产的，
            # 不是尾料损耗（金额口径判断，机织米/kg 量纲不同也不受影响）。
            need_for_finished = ZERO
            for _line in _bom_lines_by_order.get(order_no, []):
                _price = _need_prices.get(_line.sub_item_code)
                if _price:
                    need_for_finished += (
                        Decimal(_line.qty) * finished_qty_val * Decimal(_price)
                    )
            _opening_row = opening_map.get(order_no)
            on_books = material_by_order.get(order_no, ZERO) + Decimal(
                _opening_row.material if _opening_row else ZERO
            )
            if (
                need_for_finished <= ZERO
                or on_books
                <= need_for_finished * (Decimal("1") + WIP_KEEP_OVERRUN_PCT / Decimal("100"))
            ):
                wip = ZERO
        _wip_cache[order_no] = wip
        return wip

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
        wip_qty = _final_wip_qty(wo)
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

        # 折合产量口径（用户 2026-09-30）：机织（有 BOM 米重）的**在产品以 kg 计**
        #   （＝在产纱线重量），完工也换成 kg（完工米×米重）→ 折合产量＝完工kg＋0.5×在产品kg；
        #   其它车间（以及没有米重的机织单）用各自数量（件/kg/米）。见 equivalent_quantity。
        #   同一个车间内口径一致，工资池按车间折合产量合计分摊。
        workshop_equivalent = ZERO
        equivalent_by_order = {}
        for other in orders:
            if other.production_dept_code != wo.production_dept_code:
                continue
            other_finished = finished_by_order.get(other.order_no, ZERO)
            other_wip = _final_wip_qty(other)
            equivalent = equivalent_quantity(
                other.production_dept_code,
                other.order_no,
                other_finished,
                other_wip,
                jz_kg_per_m,
            )
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
        #   材料、制造费用：一般车间按数量比例；机织（在产品 kg）按**耗料 kg 占比**——
        #     完工耗料kg ＝ 完工米×米重×(1+损耗率)，在产品kg 就是剩下的料 →
        #     相当于"完工只承担报工该用的料、余额留在产品"（用户 2026-09-30 口径）。
        #   工资按折合产量（完工 + 0.5×在产品，机织以 kg 计）占比拆。
        my_finished_eq = equivalent_quantity(
            wo.production_dept_code, wo.order_no, finished_qty, ZERO, jz_kg_per_m
        )
        my_wip_eq = wip_qty
        kg_m_now = jz_kg_per_m.get(wo.order_no, ZERO)
        if wip_qty <= ZERO:
            finished_share = Decimal("1")
        elif finished_qty <= ZERO:
            finished_share = ZERO
        elif wo.order_no in backfilled_orders:
            # 「无投料-有报工」按 BOM 反推补进来的料：全部属于本期完工
            # （这些工单没有投料台账、也没有期初材料，那笔料就是按报工量补的；
            #   若按数量比例拆，"计划−完工"的在产品会吃掉大头 → 与「材料全进完工」口径不符）
            finished_share = Decimal("1")
        elif wo.production_dept_code == "JiZhi" and kg_m_now > ZERO:
            consumed_kg = finished_qty * kg_m_now * (
                Decimal("1") + JIZHI_WIP_LOSS_RATE
            )
            total_kg = consumed_kg + wip_qty
            finished_share = (
                clamp_rate(consumed_kg / total_kg) if total_kg > ZERO else Decimal("1")
            )
        else:
            # 非机织、有在产品的单：**完工材料按「完工数量 × BOM 单耗 × 价」反推**
            # （用户 2026-09-30：上月有在产、本月没投料但有报工的单，完工该承担多少料应按
            #   完工入库量推，而不是把账上材料按数量比例摊一点）。
            # 上限＝材料总额：料不够报工用 → 全部进完工、在产品不承担材料（与情形③/④ 一致）。
            need_finished = _bom_need_for_finished(wo.order_no, finished_qty)
            if need_finished > ZERO and direct_material > ZERO:
                finished_share = clamp_rate(
                    min(need_finished, direct_material) / direct_material
                )
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
        finished_hours = my_finished_eq
        wip_hours = my_wip_eq * WIP_CONVERSION_RATE
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
            processing_fee=allocated["processing_fee"],
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
    txns = period.inventory_transactions.filter(business_type__in=FINISHED_RECEIPT_TYPES)
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
    txns = period.inventory_transactions.filter(business_type__in=FINISHED_RECEIPT_TYPES)
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
            txns, ["unit_price", "amount", "price_source"], batch_size=BULK_UPDATE_BATCH
        )


#: 浆粕车间的固定车间名（与 material_issue_voucher.PULP_WORKSHOP 一致）
PULP_WORKSHOP = "浆粕车间"

#: 展示顺序（用户 2026-09-21）：染色 → 纺纱 → 包覆 → 针织 → 机织 → 浆粕。
#: 若该期间已排过结转顺序（车间结转页可调），以那个顺序为准，其余按本顺序补齐。
WORKSHOP_DISPLAY_ORDER = tuple(SMALL_ORDER_WORKSHOPS) + (PULP_WORKSHOP,)


def build_pulp_cost_rows(period: AccountingPeriod):
    """浆粕车间的「工单成本」行（用户 2026-09-21：把浆粕车间也补进这张表）。

    浆粕没有生产工单/投料/报工，所以一行＝一个**浆粕成品**，数据全部来自上传的浆粕流水
    与浆粕成本归集（collect_pulp_costs）：
      完工数量＝该成品产量（生产入库流水）；材料＝按备注分摊到的原料领料金额；
      制造费用 / 工资＝车间费用按产量比例分摊，并按科目结构拆回 水电/蒸汽/天然气/折旧/机物料/
      染料助剂/加工费；总成本＝材料＋费用＋工资；单位成本＝总成本÷产量；在产品＝0（浆粕全完工）。

    返回可直接喂给模板/导出的行对象（用 SimpleNamespace 模拟工单成本结果的字段）。
    """
    from types import SimpleNamespace

    try:
        from settlement.services.pulp_costing import collect_pulp_costs

        costs = collect_pulp_costs(period)
    except Exception:
        return []
    products = costs.get("products") or {}
    if not products:
        return []
    output_total = Decimal(costs.get("output_total") or 0)
    pool = period.workshop_cost_pools.filter(workshop=PULP_WORKSHOP).first()
    pool_by_field = {
        field: Decimal(getattr(pool, field, 0) or 0) if pool else Decimal("0")
        for field, _label in FEE_FIELDS
    }
    rows = []
    for code, item in products.items():
        qty = Decimal(item.get("output") or 0)
        share = (qty / output_total) if output_total else Decimal("0")
        material = Decimal(item.get("material") or 0)
        wages = Decimal(item.get("wages") or 0)
        total = Decimal(item.get("total") or 0)
        unit = Decimal(item.get("unit") or 0)
        fee_split = {
            field: money(pool_by_field[field] * share) for field, _label in FEE_FIELDS
        }
        row = SimpleNamespace(
            # 工单（浆粕没有工单：给出可展示的占位）
            work_order=SimpleNamespace(
                order_no="浆粕车间（按上传流水）",
                production_dept_code=PULP_WORKSHOP,
                get_status_display=lambda: "浆粕车间",
            ),
            wo_category="浆粕",
            product_code=code,
            product_name=item.get("name") or "",
            wo_planned_finish="",
            wo_actual_finish="",
            wo_completion_status="已完成",
            plan_quantity=None,
            opening_quantity=None,
            opening_amount=None,
            current_issue=None,
            current_unit_price=None,
            current_unit_price_from_bom=False,
            accum_issue=None,
            issue_status="",
            current_finished=qty,
            accum_finished=qty,
            current_finish_rate_pct=Decimal("100") if qty else None,
            finish_rate_pct=Decimal("100") if qty else None,
            uses_roving="—",
            finished_quantity=qty,
            wip_quantity=Decimal("0"),
            equivalent_qty=qty,
            direct_material=money(material),
            wages=money(wages),
            total_cost=money(total),
            finished_cost=money(total),
            finished_unit_cost=unit,
            wip_cost=Decimal("0"),
        )
        for field, _label in FEE_FIELDS:
            setattr(row, field, fee_split[field])
            setattr(row, "rate_%s_pct" % field, None)
        rows.append(row)
    rows.sort(key=lambda r: -(Decimal(r.finished_quantity or 0)))
    return rows


def build_work_order_cost_summary(period: AccountingPeriod, selected_workshop=""):
    """本月生产成本与费用（工单口径；浆粕车间走上传流水口径）。"""
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

    # 车间顺序：染色 → 纺纱 → 包覆 → 针织 → 机织 → 浆粕（用户 2026-09-21）
    # 该期间已排过结转顺序时，以那个顺序为主，不在其内的车间按固定顺序补在后面。
    sequence_order = [normalize_workshop(code) for code in settlements]
    display_index = {code: i for i, code in enumerate(WORKSHOP_DISPLAY_ORDER)}
    workshops = sorted(
        set(by_workshop) | set(pools) | set(settlements),
        key=lambda code: (
            sequence_order.index(code)
            if code in sequence_order
            else len(sequence_order),
            display_index.get(code, len(WORKSHOP_DISPLAY_ORDER)),
            code,
        ),
    )
    if selected_workshop:
        # 按车间筛选时，车间列表也收敛到该车间（否则会出现其它车间的空行）
        workshops = [w for w in workshops if w == selected_workshop]
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
        # 浆粕车间：没有生产工单/投料/报工，材料与完工成本按上传的浆粕流水统计
        # （口径与浆粕车间自己的成本归集一致：材料＝原料领料金额，完工成本＝成品归集总成本）。
        # 用户 2026-09-21：不这样填，浆粕那一行只有车间费用，差额会挂 612,406 当成未分配。
        if workshop == PULP_WORKSHOP:
            pulp_filled = False
            try:
                from settlement.services.pulp_costing import collect_pulp_costs

                pulp = collect_pulp_costs(period)
            except Exception:
                pulp = None
            if pulp and pulp.get("products"):
                current_material = money(
                    Decimal(pulp["material_fanglun"] or 0)
                    + Decimal(pulp["material_jinglun"] or 0)
                )
                finished_cost = money(
                    sum(
                        (Decimal(item["total"] or 0) for item in pulp["products"].values()),
                        ZERO,
                    )
                )
                current_input = money(current_material + current_expenses)
                ending_wip = ZERO
                cost_to_account = money(opening_wip + current_input)
                accounted_output = money(finished_cost + ending_wip)
                # 浆粕的「完工」来自上传流水，与车间结转状态无关 → 直接算已计算
                is_calculated = True
                pulp_filled = True
            if pulp_filled:
                totals_pulp_note = True
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


def annotate_work_order_cost_rows(period, results, openings=None):
    """给工单成本结果行挂上「工单成本明细」页要用的展示字段（用户 2026-09-24 抽出共用）。

    从 `views.work_order_costs` 里整段搬过来的，**口径完全没有改**，所以：
      * 「工单成本」页（work_order_costs）
      * 「生产成本」页的「工单投料报工明细」
      * 导出的「工单成本」工作表
    三处看到的同一张工单的数字完全一致（同一份取数逻辑）。

    挂上的字段包括：期初数量/金额、累计与本期投料、投料单价与状态、本期/累计完工与完工率、
    在产品、折合产量、完成状态、是否用粗纱、类别、计划/实际完工时间、各制造费用分配率、
    标准单位成本与单位成本差异等。
    """
    from decimal import Decimal as _D
    from django.db.models import Max as _Max
    from django.db.models import Sum as _Sum

    from settlement.models import BomLine, IssueProgress  # noqa: F401

    openings = openings or {
        item.order_no: item for item in period.opening_wip.all()
    }
    # BOM 预载（用户 2026-09-27）：没有投料的单要按 BOM 表折算「投料单价」，
    # 一次全量取出按工单号分组，免得每单查一遍库。
    bom_by_order_no = defaultdict(list)
    for _line in period.bom_lines.all():
        bom_by_order_no[_line.work_order_no].append(_line)
    # BOM 100% 全粗纱的工单（「是否用粗纱」列＝是）
    _all_roving_bom_orders = all_roving_bom_orders(period)

    # 车间聚合：材料总额 / 折合数量合计（分母）
    _kg_per_m_map = bom_kg_per_meter(period)

    def _equivalent_qty(item):
        """折合产量（与结算层同一口径，用户 2026-09-30，见 equivalent_quantity）。"""
        return equivalent_quantity(
            item.work_order.production_dept_code,
            item.work_order.order_no,
            item.finished_quantity,
            item.wip_quantity,
            _kg_per_m_map,
        )

    workshop_material = defaultdict(Decimal)
    workshop_equivalent = defaultdict(Decimal)
    for item in results:
        opening = openings.get(item.work_order.order_no)
        current_material = Decimal(item.direct_material or 0) - Decimal(
            opening.material if opening else 0
        )
        workshop_material[item.work_order.production_dept_code] += current_material
        workshop_equivalent[item.work_order.production_dept_code] += _equivalent_qty(item)

    # 车间级制造费用分配率 = 该费用总额 ÷ 车间材料成本总额（每元材料承担的费用）
    from settlement.services.workorder_costing import MANUFACTURING_FIELDS

    workshop_fee_total = defaultdict(lambda: defaultdict(Decimal))
    workshop_material_total = defaultdict(Decimal)
    for item in results:
        ws = item.work_order.production_dept_code
        workshop_material_total[ws] += Decimal(item.direct_material or 0)
        for field in MANUFACTURING_FIELDS:
            workshop_fee_total[ws][field] += Decimal(getattr(item, field) or 0)

    # 每个工单的分配率（制造费用 = 车间级费用/材料；工资 = 折合数量占比）
    for item in results:
        ws = item.work_order.production_dept_code
        for field in MANUFACTURING_FIELDS:
            fee_total = workshop_fee_total[ws][field]
            mat_total = workshop_material_total[ws]
            rate = (
                ((fee_total / mat_total) * Decimal("100"))
                if mat_total
                else Decimal("0")
            )
            setattr(item, f"rate_{field}_pct", rate if rate else None)
        equivalent = _equivalent_qty(item)
        item.equivalent_qty = equivalent
        denom2 = workshop_equivalent[ws]
        wage_rate = (
            (equivalent / denom2 * Decimal("100")) if denom2 else Decimal("0")
        )
        item.rate_wages_pct = wage_rate if wage_rate else None

    # 投料进度（投料进度表）与粗纱标志
    from settlement.models import BomLine, IssueProgress

    progress_rows = list(
        period.issue_progress.values("work_order_no").annotate(
            total=Sum("actual_issue_quantity")
        )
    )
    progress_qty_by_order = {
        row["work_order_no"]: Decimal(row["total"] or 0) for row in progress_rows
    }
    progress_status_by_order = defaultdict(set)
    for status, wo in period.issue_progress.exclude(
        issue_status=""
    ).values_list("issue_status", "work_order_no"):
        progress_status_by_order[wo].add(status)
    roving_codes = set(
        BomLine.objects.filter(period=period, sub_item_code__startswith="BC-")
        .values_list("sub_item_code", flat=True)
    )
    roving_issue_orders = set(
        period.material_issues.filter(material_code__startswith="BC-")
        .exclude(work_order_no="")
        .values_list("work_order_no", flat=True)
    )
    roving_progress_orders = set(
        period.issue_progress.filter(material_code__startswith="BC-")
        .exclude(work_order_no="")
        .values_list("work_order_no", flat=True)
    )
    # 工单明细（用户上传）：类别/计划完工时间/实际完工时间/完成状态 按工单编号关联
    work_order_details = {
        d.order_no: d
        for d in period.work_order_details.all()
    }
    # 各工单实际完工时间兜底：报工记录最大报工时间
    from django.db.models import Max as _MaxDetail

    report_time_agg = {
        row["work_order_no"]: row["max_time"]
        for row in period.work_reports.exclude(work_order_no="")
        .values("work_order_no")
        .annotate(max_time=_MaxDetail("report_time"))
    }
    # 类别解析器（与完工入库结转凭证共用同一套规则，见 build_order_category_resolver）
    resolve_order_category = build_order_category_resolver(period, work_order_details)
    # 产品编码/名称统一解析器（BOM → 标识码提取 → 工单号）
    resolve_product = build_product_resolver(period)
    # 投料物料单价：月末加权价优先（成本口径），无则历史单价兜底
    material_price = {}
    for v in period.inventory_valuations.all():
        material_price.setdefault(v.material_code, Decimal(v.weighted_average_price or 0))
    for hp in period.historical_prices.all():
        if hp.issue_price is not None:
            material_price.setdefault(hp.material_code, Decimal(hp.issue_price))
        elif hp.receipt_price is not None:
            material_price.setdefault(hp.material_code, Decimal(hp.receipt_price))
    # 工单投料数量聚合：(work_order_no, material_code) -> 实际投料数量
    issue_qty_by_pair = defaultdict(Decimal)
    for row in period.material_issues.exclude(work_order_no="").exclude(material_code="").values(
        "work_order_no", "material_code", "actual_issue_quantity"
    ):
        issue_qty_by_pair[(row["work_order_no"], row["material_code"])] += Decimal(
            row["actual_issue_quantity"] or 0
        )
    # 投料台账跨期累计（用户 2026-10-16）：本期及以前各期 actual_issue_quantity 合计
    accum_issue_all = defaultdict(Decimal)
    for _pa in AccountingPeriod.objects.filter(month__lte=period.month):
        for _row in (
            _pa.material_issues.exclude(work_order_no="")
            .values("work_order_no", "actual_issue_quantity")
        ):
            accum_issue_all[_row["work_order_no"]] += Decimal(
                _row["actual_issue_quantity"] or 0
            )

    for item in results:
        order_no = item.work_order.order_no
        ws = item.work_order.production_dept_code
        # 期初在产品：数量/金额（金额=材料+制造费用+工资合计）
        opening = openings.get(order_no)
        item.opening_quantity = Decimal(opening.quantity if opening else Decimal("0"))
        item.opening_amount = Decimal(opening.total if opening else Decimal("0"))
        # 累计投料：**投料台账跨期累计**（本期＋以前各期，用户 2026-10-16）
        #   不再用投料进度表（常没同步），也不再用"仅本期"的合计
        item.accum_issue = accum_issue_all.get(order_no, Decimal(0))
        # 投料状态：进度表该工单各行状态去重
        statuses = progress_status_by_order.get(order_no)
        item.issue_status = "、".join(sorted(statuses)) if statuses else "—"
        # 本期完工 / 累计完工 / 完工率（用户 2026-09-24 定稿口径）：
        #   本期完工＝本月报工数量；累计完工＝工单明细的「入库数量」（ERP 跨月累计，无明细退回本月报工）；
        #   两个完工率＝完工 ÷ 计划数量 × 100。原来写成「本期=累计」「完工÷(完工+在产品)」，
        #   两对列完全重复且完工率其实是成本拆分比例，已按上面口径改掉（见 attach_finish_progress）。
        attach_finish_progress([item], work_order_details)
        wip_qty = Decimal(item.wip_quantity or 0)
        # 本期投料（当月投料；投料无期初结转，本期=累计）
        item.current_issue = Decimal(item.accum_issue or 0)
        # 投料单价（保留 2 位小数，用户 2026-09-27 定；2026-09-30 补第 ③ 条）：
        #   ① 有投料、材料全来自实领 → 实际投料的加权平均单价 Σ(数量×月末加权价) ÷ Σ数量；
        #   ② 没有投料（或投料都没价）→ 按 BOM 表折算 Σ(BOM子项用量×价) ÷ Σ用量；
        #   ③ **材料里含"按 BOM 反推补进来"的部分**（情形③ 差额反推 / 情形④ BOM 反推 / 粗纱链）→
        #      也用 BOM 折算单价：这单的材料是"应有料"而不是"实际投的那几种料"，拿实领的加权价
        #      会和材料成本对不上——实测 TDE025-BF20260720 只投了氨纶（25.66/kg），
        #      而材料 20.80 ÷ 报工 1.6 ＝ 13.00，正是 BOM 折算价 12.9991（含没投的涤纶）。
        _price = issue_unit_price(issue_qty_by_pair, material_price, order_no)
        _bom_price = bom_unit_price(
            period, order_no, material_price,
            bom_lines=bom_by_order_no.get(order_no, []),
        )
        _issue_amount = sum(
            (
                qty * Decimal(material_price.get(code, ZERO) or ZERO)
                for (wo_no, code), qty in issue_qty_by_pair.items()
                if wo_no == order_no
            ),
            ZERO,
        )
        _opening_material = Decimal(opening.material or ZERO) if opening else ZERO
        _calc_backfill = (
            Decimal(item.direct_material or ZERO) - _issue_amount - _opening_material
        )
        item.current_unit_price_from_bom = False
        if _calc_backfill > ZERO and _bom_price is not None:
            _price = _bom_price
            item.current_unit_price_from_bom = True
        elif _price is None:
            _price = _bom_price
            item.current_unit_price_from_bom = _bom_price is not None
        # 两处都取不到价就留 None（页面/导出显示 —，不写成 0 免得像是"零元投料"）
        item.current_unit_price = _price
        # 本期完工（本月报工）已由 attach_finish_progress 一并挂上
        # 完成状态：未开工 / 部分完成 / 已完成（按累计完工 + 在产品）
        if item.accum_finished <= Decimal("0") and wip_qty <= Decimal("0"):
            item.completion_status = "未开工"
        elif wip_qty <= Decimal("0"):
            item.completion_status = "已完成"
        else:
            item.completion_status = "部分完成"
        # 是否用粗纱（用户 2026-09-30 定稿）：**BOM 子项 100% 都是粗纱（BC-）**才显示"是"，
        # 混料单（粗纱＋其它料，如 RAJ031A-HF20260603）显示"否"；所有车间都显示 是/否。
        item.uses_roving = "是" if order_no in _all_roving_bom_orders else "否"
        # 类别/计划完工时间/实际完工时间/完成状态：优先取工单明细（用户上传），无则系统回退
        detail = work_order_details.get(order_no)
        # 产品编码/名称：结算值优先，BOM/标识码兜底（统一 resolver）
        product_code, product_name = resolve_product(
            item.work_order, item.product_code, item.product_name
        )
        item.product_code = product_code
        if product_name:
            item.product_name = product_name
        # 类别：产品物料分类（规则见 build_order_category_resolver，与完工入库凭证共用）
        item.wo_category = resolve_order_category(
            order_no, product_code, ws, item.work_order.order_type
        )
        if detail is not None and detail.planned_finish:
            item.wo_planned_finish = detail.planned_finish
        else:
            item.wo_planned_finish = (
                item.work_order.planned_finish.strftime("%Y-%m-%d %H:%M")
                if item.work_order.planned_finish
                else ""
            )
        if detail is not None and detail.actual_finish:
            item.wo_actual_finish = detail.actual_finish
        else:
            _t = report_time_agg.get(order_no)
            item.wo_actual_finish = _t.strftime("%Y-%m-%d %H:%M") if _t else ""
        if detail is not None and detail.completion_status:
            item.wo_completion_status = detail.completion_status
        else:
            item.wo_completion_status = item.completion_status

    # 标准单位成本（用户 2026-09-22）：放到本页与实际单位成本并排对照
    # 标准单位成本/单位成本差异：实现在 views（历史原因），这里局部导入避免循环依赖
    from settlement.views import attach_standard_unit_costs as _attach_std

    _attach_std(period, results)
    return results


