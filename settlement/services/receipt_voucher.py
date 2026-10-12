# -*- coding: utf-8 -*-
"""完工入库结转凭证生成（用户 2026-09-25；贷方改要素 2026-09-27）。

口径：
- **一张凭证、按车间分块**：
  借方 **库存商品—物料分类**（每个车间的每个物料分类一行）；
  贷方 **生产成本—直接材料 / 直接人工 / 制造费用**（全车间汇总，各一行，含浆粕车间）。
- **摘要按车间**：「{月}月{车间}完工入库{物料分类}」；贷方要素行「{月}月各车间完工入库{要素}」、
  浆粕行「{月}月浆粕车间完工入库」。
- **金额基准＝工单完工成本**（= 报工数量 × 完工单位成本，与「工单成本」页一致）；
  贷方三个要素分别取工单成本的 **完工直接材料 / 完工人工 / 完工制造费用**
  （三者相加＝完工成本，借贷仍然相等）。
  ⚠️ **不用库存「报工」入库流水的金额**：本地 7 月的报工入库流水里有**重复行**
  （例：31.731-20260525 只有 1 条报工 2527.4，流水却有 2 行完全相同；
  JW076-HF20260649 报工 2 条，流水 7 行），流水金额比完工成本多 361.6 万，
  按流水入账会把成本重复计一遍。该差异在 checks 里列示，请到源头核对报工入库流水。
- **浆粕车间**没有工单、不走工单成本：它的成品按库存流水「生产入库」归集
  （本期 3 个成品），科目 库存商品—浆粕；**贷方也按成本要素拆**（用户 2026-09-27）：
  直接材料＝浆粕原料领料、直接人工＝车间工资、制造费用＝水电/蒸汽/折旧/机物料等车间费用，
  三者之和＝「生产入库」金额（取自 `pulp_costing.collect_pulp_costs`，
  本期 3,871,389.85 + 227,516.23 + 384,889.78 = 4,483,795.86）。

生成前检查（blocking 会挡住入账）：
1. 报工入库流水里的**重复行**（同一单据+日期+数量+金额多行）；
2. 报工入库流水金额 与 工单完工成本 的差异（列示差额与涉及工单数）；
3. 有完工成本但**没有完工入库流水**的工单；
4. 产品分类判不出来的工单（会挂到「库存商品—未分类」）；
5. 借贷不平。
"""
from __future__ import annotations

from collections import defaultdict
from decimal import ROUND_HALF_UP, Decimal

from settlement.models import AccountingPeriod
from settlement.services.imports import WORKSHOP_LABELS
from settlement.services.pulp_costing import collect_pulp_costs
from settlement.services.workorder_costing import (
    FINISHED_RECEIPT_TYPES,
    PULP_WORKSHOP,
    SMALL_ORDER_WORKSHOPS,
    build_order_category_resolver,
)

ZERO = Decimal("0")
CENT = Decimal("0.01")

#: 借方科目（库存商品，按物料分类挂）
ACCOUNT_INVENTORY = "库存商品"
#: 贷方科目（生产成本）
ACCOUNT_PRODUCTION = "生产成本"
#: 分类判不出来时的兜底分类
UNCLASSIFIED = "未分类"

#: 贷方二级科目＝成本要素（用户 2026-09-27）：直接材料 / 直接人工 / 制造费用。
#: 金额取工单成本的完工直接材料 / 完工人工 / 完工制造费用（三者相加＝完工成本）。
ELEMENT_MATERIAL = "直接材料"
ELEMENT_WAGES = "直接人工"
ELEMENT_OVERHEAD = "制造费用"
PRODUCTION_ELEMENTS = (ELEMENT_MATERIAL, ELEMENT_WAGES, ELEMENT_OVERHEAD)
#: 工单成本结果里对应要素的字段名
ELEMENT_FIELDS = {
    ELEMENT_MATERIAL: "finished_material",
    ELEMENT_WAGES: "finished_wages",
    ELEMENT_OVERHEAD: "finished_overhead",
}

#: 车间顺序（染色→纺纱→包覆→针织→机织），浆粕车间单独排最后
WORKSHOP_SEQUENCE = tuple(SMALL_ORDER_WORKSHOPS)
UNALLOCATED = "无工单"
UNALLOCATED_LABEL = "无工单入库"

#: 浆粕车间成品识别（浆粕车间没有工单，成品按库存「生产入库」流水归集）
PULP_KEYWORDS = ("浆粕", "微纤")
PULP_CODE_PREFIX = "JP-"

#: 取数来源标签
SOURCE_ORDER = "工单完工成本"
SOURCE_PULP_FLOW = "库存流水·生产入库"
#: 浆粕车间：金额取「生产入库」，拆分取浆粕成本归集
SOURCE_PULP_SPLIT = "浆粕成本归集（材料/工资/车间费用）"


def pulp_element_totals(period):
    """浆粕车间的要素拆分：直接材料 / 直接人工 / 制造费用（用户 2026-09-27）。

    浆粕车间没有工单、不走工单成本，它的完工成本＝库存流水「生产入库」的金额；
    要素拆分取「浆粕车间成本归集」（`pulp_costing.collect_pulp_costs`）：
    - 直接材料＝原料领料（按备注分摊到成品后合计）
    - 直接人工＝车间工资
    - 制造费用＝水电/蒸汽/天然气/折旧/机物料/染料助剂/加工费
    取不到（未上传浆粕原料流水等）时返回 None，由调用方兜底并入直接材料并给出提示。
    """
    try:
        data = collect_pulp_costs(period)
    except Exception:  # noqa: BLE001 —— 归集不了不能让整张凭证挂掉
        return None
    material = money(
        Decimal(data.get("material_fanglun") or ZERO)
        + Decimal(data.get("material_jinglun") or ZERO)
    )
    wages = money(data.get("wages_total"))
    overhead = money(data.get("mfg_total"))
    products = {}
    for code, item in (data.get("products") or {}).items():
        products[code] = {
            ELEMENT_MATERIAL: money(item.get("material")),
            ELEMENT_WAGES: money(item.get("wages")),
            ELEMENT_OVERHEAD: money(item.get("fee")),
            "total": money(item.get("total")),
        }
    return {
        ELEMENT_MATERIAL: material,
        ELEMENT_WAGES: wages,
        ELEMENT_OVERHEAD: overhead,
        "total": money(material + wages + overhead),
        "products": products,
    }

VOUCHER_NOTE = (
    "借：库存商品—物料分类（按车间 × 物料分类逐行）；"
    "贷：生产成本—直接材料 / 直接人工 / 制造费用（全车间汇总，各一行）。"
    "摘要按车间：{月}月{车间}完工入库{物料分类}。"
    "金额基准＝工单完工成本（报工数量 × 完工单位成本，与「工单成本」页一致），"
    "贷方三个要素取工单成本的完工直接材料 / 完工人工 / 完工制造费用（三者相加＝完工成本）；"
    "浆粕车间没有工单，成品按库存流水「生产入库」归集，也拆成这三个要素"
    "（材料＝浆粕原料领料、人工＝车间工资、制造费用＝其余车间费用）。"
)


def money(value) -> Decimal:
    return Decimal(value or ZERO).quantize(CENT, rounding=ROUND_HALF_UP)


def _bucket():
    return {"amount": ZERO, "quantity": ZERO, "rows": 0, "orders": 0}


def _add(bucket, amount, quantity, rows=1, orders=0):
    bucket["amount"] += amount
    bucket["quantity"] += quantity
    bucket["rows"] += rows
    bucket["orders"] += orders


def workshop_label(code: str) -> str:
    if not code:
        return ""
    if code in (PULP_WORKSHOP, UNALLOCATED):
        return UNALLOCATED_LABEL if code == UNALLOCATED else PULP_WORKSHOP
    return WORKSHOP_LABELS.get(code, code)


def debit_summary_text(period, workshop, category="") -> str:
    """借方摘要：{月}月{车间}完工入库{物料分类}。"""
    month = "%d月" % period.month.month
    label = workshop_label(workshop)
    if category:
        return "%s%s完工入库%s" % (month, label, category)
    return "%s%s完工入库" % (month, label)


def credit_summary_text(period, workshop) -> str:
    """贷方摘要（车间行）：{月}月{车间}完工入库。"""
    return "%s%s完工入库" % ("%d月" % period.month.month, workshop_label(workshop))


def element_credit_summary_text(period, element) -> str:
    """贷方摘要（要素行）：{月}月各车间完工入库{直接材料/直接人工/制造费用}。

    含浆粕车间（它没有工单，按库存流水「生产入库」归集后再拆要素）。
    """
    return "%s各车间完工入库%s" % ("%d月" % period.month.month, element)


def inventory_account(category: str) -> str:
    """借方科目：库存商品—物料分类。"""
    if not category or category == UNCLASSIFIED:
        return ACCOUNT_INVENTORY
    return "%s—%s" % (ACCOUNT_INVENTORY, category)


def production_account(workshop: str) -> str:
    """贷方科目（车间行）：生产成本—车间（保留给按车间挂科目时使用）。"""
    return "%s—%s" % (ACCOUNT_PRODUCTION, workshop_label(workshop))


def production_element_account(element: str) -> str:
    """贷方科目（要素行）：生产成本—直接材料 / 直接人工 / 制造费用。"""
    return "%s—%s" % (ACCOUNT_PRODUCTION, element)


def _is_pulp_row(txn) -> bool:
    name = str(txn.material_name or "")
    code = str(txn.material_code or "")
    if (txn.category_name or "").strip() == "浆粕":
        return True
    if any(keyword in name for keyword in PULP_KEYWORDS):
        return True
    return code.upper().startswith(PULP_CODE_PREFIX)


def _workshop_order(order_code):
    if order_code == PULP_WORKSHOP:
        return len(WORKSHOP_SEQUENCE)  # 浆粕排最后
    if order_code == UNALLOCATED:
        return len(WORKSHOP_SEQUENCE) + 1
    try:
        return WORKSHOP_SEQUENCE.index(order_code)
    except ValueError:
        return len(WORKSHOP_SEQUENCE) - 1


def receipt_flow_rows(period):
    """本期完工入库流水（业务类型＝报工 / 生产入库），按单据排序。"""
    rows = [
        txn
        for txn in period.inventory_transactions.all()
        if (txn.business_type or "").strip() in FINISHED_RECEIPT_TYPES
    ]
    rows.sort(key=lambda t: (t.txn_date or "", t.document_no or "", t.id))
    return rows


def find_duplicate_flow_rows(period, rows=None):
    """找出完工入库流水里的重复行（同一单据+日期+数量+金额出现多行）。

    返回 (重复行数, 重复金额)。重复行＝实际多出来的行（同组里除第一行之外的都算）。
    """
    rows = rows if rows is not None else receipt_flow_rows(period)
    groups = defaultdict(list)
    for txn in rows:
        key = (
            (txn.document_no or "").strip(),
            txn.txn_date,
            str(Decimal(txn.quantity or ZERO)),
            str(money(txn.amount)),
        )
        groups[key].append(txn)
    dup_rows = 0
    dup_amount = ZERO
    for key, items in groups.items():
        if len(items) <= 1:
            continue
        dup_rows += len(items) - 1
        dup_amount += money(items[0].amount) * (len(items) - 1)
    return dup_rows, money(dup_amount)


def _element_values(item):
    """一张工单的（直接材料, 直接人工, 制造费用），保证三者相加＝完工成本。

    正常数据里工单成本的 finished_material / finished_wages / finished_overhead
    本来就是从完工成本拆出来的（三者和＝finished_cost）。个别历史/手工数据没拆，
    这里把差额并进直接材料，避免借贷不平。
    """
    cost = money(item.finished_cost)
    material = money(item.finished_material)
    wages = money(item.finished_wages)
    overhead = money(item.finished_overhead)
    residual = money(cost - (material + wages + overhead))
    if residual != ZERO:
        material = money(material + residual)
        return material, wages, overhead, True
    return material, wages, overhead, False


def build_finished_receipt_voucher(period: AccountingPeriod):
    """生成完工入库结转凭证。"""
    resolver = build_order_category_resolver(period)
    results = list(
        period.work_order_cost_results.select_related("work_order").order_by(
            "work_order__production_dept_code", "work_order__order_no"
        )
    )
    flow_rows = receipt_flow_rows(period)

    # ── 借方（库存商品按车间×分类）+ 贷方（生产成本按要素）汇总工单完工成本 ──
    debit = defaultdict(_bucket)
    elements = defaultdict(
        lambda: {
            ELEMENT_MATERIAL: ZERO,
            ELEMENT_WAGES: ZERO,
            ELEMENT_OVERHEAD: ZERO,
            "amount": ZERO,
            "quantity": ZERO,
            "orders": 0,
        }
    )
    detail_rows = []
    zero_cost_orders = 0
    unclassified_orders = 0
    residual_orders = 0
    orders_with_flow = set()
    for txn in flow_rows:
        doc = (txn.document_no or "").strip()
        if doc:
            orders_with_flow.add(doc)

    for item in results:
        order_no = item.work_order.order_no
        workshop = item.work_order.production_dept_code
        cost = money(item.finished_cost)
        if cost == ZERO:
            zero_cost_orders += 1
            continue
        material, wages, overhead, residual = _element_values(item)
        if residual:
            residual_orders += 1
        category = resolver(
            order_no,
            item.product_code,
            workshop,
            item.work_order.order_type,
        ) or UNCLASSIFIED
        if category == UNCLASSIFIED:
            unclassified_orders += 1
        quantity = Decimal(item.finished_quantity or ZERO)
        _add(debit[(workshop, category)], cost, quantity, 1, 1)
        bucket = elements[workshop]
        bucket[ELEMENT_MATERIAL] += material
        bucket[ELEMENT_WAGES] += wages
        bucket[ELEMENT_OVERHEAD] += overhead
        bucket["amount"] += cost
        bucket["quantity"] += quantity
        bucket["orders"] += 1
        detail_rows.append(
            {
                "period": period.month.strftime("%Y-%m"),
                "workshop": workshop,
                "workshop_label": workshop_label(workshop),
                "order_no": order_no,
                "product_code": item.product_code,
                "product_name": item.product_name,
                "category": category,
                "finished_quantity": quantity,
                "finished_unit_cost": Decimal(item.finished_unit_cost or ZERO),
                "material": material,
                "wages": wages,
                "overhead": overhead,
                "amount": cost,
                "source": SOURCE_ORDER,
                "has_flow": order_no in orders_with_flow,
            }
        )

    # ── 浆粕车间：没有工单，按库存流水「生产入库」归集，再按成本归集拆要素 ──
    pulp_split = pulp_element_totals(period) if flow_rows else None
    pulp_split_residual = None
    pulp_products = (pulp_split or {}).get("products") or {}
    pulp_rows = []
    for txn in flow_rows:
        if not _is_pulp_row(txn):
            continue
        category = (txn.category_name or "").strip() or "浆粕"
        amount = money(txn.amount)
        quantity = Decimal(txn.quantity or ZERO)
        # 逐成品的要素拆分（同一成品多条生产入库流水时，金额按流水金额比例摊）
        split = pulp_products.get(txn.material_code)
        prod_total = money(split["total"]) if split else ZERO
        share = (amount / prod_total) if (split and prod_total) else ZERO
        _add(debit[(PULP_WORKSHOP, category)], amount, quantity, 1, 1)
        pulp_rows.append(
            {
                "period": period.month.strftime("%Y-%m"),
                "workshop": PULP_WORKSHOP,
                "workshop_label": PULP_WORKSHOP,
                "order_no": txn.document_no or "",
                "product_code": txn.material_code,
                "product_name": txn.material_name,
                "category": category,
                "finished_quantity": quantity,
                "finished_unit_cost": Decimal(txn.unit_price or ZERO),
                "material": money(split[ELEMENT_MATERIAL] * share) if split else ZERO,
                "wages": money(split[ELEMENT_WAGES] * share) if split else ZERO,
                "overhead": money(split[ELEMENT_OVERHEAD] * share) if split else ZERO,
                "amount": amount,
                "source": SOURCE_PULP_SPLIT if split else SOURCE_PULP_FLOW,
                "has_flow": True,
            }
        )

    # 浆粕的要素合计：把逐成品的拆分加起来（这样附件1的逐行和附件3的浆粕行必然一致）；
    # 与「生产入库」金额对不上时，差额并入直接材料，保证借贷相等。
    pulp_amount = money(sum((row["amount"] for row in pulp_rows), ZERO))
    pulp_quantity = sum((row["finished_quantity"] for row in pulp_rows), ZERO)
    pulp_material = pulp_wages = pulp_overhead = ZERO
    if pulp_rows and pulp_split:
        pulp_material = money(sum((row["material"] for row in pulp_rows), ZERO))
        pulp_wages = money(sum((row["wages"] for row in pulp_rows), ZERO))
        pulp_overhead = money(sum((row["overhead"] for row in pulp_rows), ZERO))
        residual = money(pulp_amount - (pulp_material + pulp_wages + pulp_overhead))
        if residual != ZERO:
            pulp_material = money(pulp_material + residual)
            pulp_split_residual = residual
    elif pulp_rows:
        # 归集不出来（未上传浆粕原料流水等）→ 全额并入直接材料，保证借贷相等
        pulp_material = pulp_amount

    # ── 分录：借方按「车间 × 分类」（摘要写车间+分类），贷方按要素一行（含浆粕） ──
    rows = []
    debit_keys = sorted(
        debit,
        key=lambda key: (_workshop_order(key[0]), str(key[1])),
    )
    for workshop, category in debit_keys:
        item = debit[(workshop, category)]
        if item["amount"] == ZERO and item["quantity"] == ZERO:
            continue
        rows.append(
            {
                "level": "line",
                "summary": debit_summary_text(period, workshop, category),
                "account": inventory_account(category),
                "aux": "%s（%d 张工单）" % (workshop_label(workshop), item["orders"]),
                "debit": money(item["amount"]),
                "credit": ZERO,
                "quantity": item["quantity"],
                "source_rows": item["rows"],
                "side": "debit",
            }
        )

    # 贷方三个要素＝工单部分（各车间）＋ 浆粕车间（按库存流水「生产入库」，再按成本归集拆要素）
    pulp_by_element = {
        ELEMENT_MATERIAL: pulp_material,
        ELEMENT_WAGES: pulp_wages,
        ELEMENT_OVERHEAD: pulp_overhead,
    }
    element_totals = {}
    for element in PRODUCTION_ELEMENTS:
        element_totals[element] = money(
            sum((bucket[element] for bucket in elements.values()), ZERO)
            + pulp_by_element[element]
        )
    element_quantity = (
        sum((bucket["quantity"] for bucket in elements.values()), ZERO) + pulp_quantity
    )
    element_orders = (
        sum((bucket["orders"] for bucket in elements.values()), 0) + len(pulp_rows)
    )
    order_element_totals = {
        element: money(
            sum((bucket[element] for bucket in elements.values()), ZERO)
        )
        for element in PRODUCTION_ELEMENTS
    }
    for element in PRODUCTION_ELEMENTS:
        amount = element_totals[element]
        if amount == ZERO:
            continue
        rows.append(
            {
                "level": "line",
                "summary": element_credit_summary_text(period, element),
                "account": production_element_account(element),
                "aux": "%d 个车间汇总（%d 张工单）%s" % (
                    len(elements) + (1 if pulp_rows else 0),
                    element_orders,
                    "＋浆粕车间" if pulp_rows else "",
                ),
                "debit": ZERO,
                "credit": amount,
                "quantity": element_quantity,
                "source_rows": element_orders,
                "side": "credit",
            }
        )

    debit_total = sum((row["debit"] for row in rows), ZERO)
    credit_total = sum((row["credit"] for row in rows), ZERO)
    quantity_total = sum((row["quantity"] for row in rows if row["debit"]), ZERO)
    rows.append(
        {
            "level": "total",
            "summary": "合计",
            "account": "",
            "aux": "",
            "debit": money(debit_total),
            "credit": money(credit_total),
            "quantity": quantity_total,
            "source_rows": len(detail_rows) + len(pulp_rows),
            "side": "",
        }
    )

    # ── 生成前检查 ────────────────────────────────────────────────────────
    checks = []
    flow_amount = money(sum((money(txn.amount) for txn in flow_rows), ZERO))
    dup_rows, dup_amount = find_duplicate_flow_rows(period, flow_rows)
    if dup_rows:
        checks.append(
            {
                "level": "blocking",
                "text": "完工入库流水里有 %d 行**重复行**（同一单据+日期+数量+金额出现多行，"
                        "重复金额 %s）：本凭证按「工单完工成本」计，没有按这些流水入账，"
                        "否则成本会重复计一遍。请先核对上传的报工入库流水（库存流水）并去重，"
                        "再把库存重算一次。" % (dup_rows, f"{dup_amount:,.2f}"),
            }
        )
    # 口径差异拆两块：① 有完工成本但流水缺失（凭证侧多出）；② 流水金额大于工单完工成本（流水侧多出）
    order_amount = sum((row["amount"] for row in detail_rows), ZERO)
    no_flow = [row for row in detail_rows if not row["has_flow"]]
    no_flow_amount = sum((row["amount"] for row in no_flow), ZERO)
    non_pulp_flow_amount = money(flow_amount - sum(
        (money(txn.amount) for txn in flow_rows if _is_pulp_row(txn)), ZERO
    ))
    flow_for_costed_orders = money(order_amount - no_flow_amount)
    flow_excess = money(non_pulp_flow_amount - flow_for_costed_orders)
    diff = money(order_amount + pulp_amount - flow_amount)
    if diff != ZERO or flow_excess != ZERO or no_flow_amount != ZERO:
        parts = []
        if no_flow_amount:
            parts.append(
                "① 有 %d 张工单有完工成本、但完工入库流水缺失（凭证侧多出 %s）"
                % (len(no_flow), f"{no_flow_amount:,.2f}")
            )
        if flow_excess:
            parts.append(
                "② 部分工单的入库流水金额大于工单完工成本（流水侧多出 %s，其中重复行 %s）"
                % (f"{flow_excess:,.2f}", f"{dup_amount:,.2f}")
            )
        checks.append(
            {
                "level": "warning",
                "text": "完工入库两条口径对不上，本凭证按**工单完工成本**计"
                        "（＝报工数量 × 完工单位成本，与「工单成本」页一致），"
                        "合计 %s；库存「报工/生产入库」流水合计 %s，差 %s。构成：%s。"
                        "两块都要到源头核对（流水漏传 / 流水重复或多传），否则库存商品与生产成本对不上。"
                        % (f"{money(order_amount + pulp_amount):,.2f}",
                           f"{flow_amount:,.2f}", f"{diff:,.2f}",
                           "；".join(parts) if parts else "无"),
            }
        )
    if unclassified_orders:
        checks.append(
            {
                "level": "warning",
                "text": "有 %d 张工单的产品物料分类判不出来，挂在「%s—%s」下，请核对产品编码/物料分类。"
                        % (unclassified_orders, ACCOUNT_INVENTORY, UNCLASSIFIED),
            }
        )
    if zero_cost_orders:
        checks.append(
            {
                "level": "info",
                "text": "另有 %d 张工单本期完工成本为 0（未完工或成本算不出），没有列入本凭证。"
                        % zero_cost_orders,
            }
        )
    if residual_orders:
        checks.append(
            {
                "level": "warning",
                "text": "有 %d 张工单的完工成本没有拆出「直接材料/直接人工/制造费用」"
                        "（工单成本的要素字段为空或加不上），差额已并入**直接材料**，"
                        "保证借贷相等；请核对这几张工单的成本。" % residual_orders,
            }
        )
    if not pulp_rows:
        checks.append(
            {"level": "info", "text": "本期没有浆粕车间的成品入库（库存流水「生产入库」）。"}
        )
    else:
        checks.append(
            {
                "level": "info",
                "text": "浆粕车间完工入库归集 %d 个成品（%s 元）：它没有工单、不走工单成本，"
                        "金额取库存流水「生产入库」，贷方**也按要素拆开**"
                        "（直接材料 %s / 直接人工 %s / 制造费用 %s，取浆粕车间成本归集）。"
                        % (len(pulp_rows), f"{pulp_amount:,.2f}",
                           f"{pulp_material:,.2f}", f"{pulp_wages:,.2f}",
                           f"{pulp_overhead:,.2f}"),
            }
        )
    if pulp_rows and pulp_split is None:
        checks.append(
            {
                "level": "warning",
                "text": "浆粕车间的要素拆分取不到（浆粕成本归集失败：多半是没上传当月"
                        "「浆粕原料流水」，或浆粕车间费用池缺数据），本期把浆粕的 %s 元"
                        "**全额并入直接材料**，保证借贷相等；请补上传后重新结转浆粕车间。"
                        % f"{pulp_amount:,.2f}",
            }
        )
    elif pulp_rows and pulp_split_residual not in (None, ZERO):
        checks.append(
            {
                "level": "warning",
                "text": "浆粕车间的要素拆分与「生产入库」金额对不上，差额 %s 已并入**直接材料**，"
                        "保证借贷相等。请核对浆粕原料流水与浆粕车间费用池："
                        "拆分合计 %s vs 生产入库 %s。" % (
                            f"{pulp_split_residual:,.2f}",
                            f"{money(pulp_material + pulp_wages + pulp_overhead):,.2f}",
                            f"{pulp_amount:,.2f}"),
            }
        )
    if not detail_rows and not pulp_rows:
        checks.append({"level": "warning", "text": "本期没有完工入库，凭证为空。"})
    if money(debit_total - credit_total) != ZERO:
        checks.append(
            {
                "level": "blocking",
                "text": "借贷不平：借方 %s / 贷方 %s，差 %s。"
                        % (f"{debit_total:,.2f}", f"{credit_total:,.2f}",
                           f"{debit_total - credit_total:,.2f}"),
            }
        )

    # ── 附件：按车间（含要素）、按车间×物料分类 ────────────────────────────
    workshop_rows = []
    for workshop in sorted(elements, key=_workshop_order):
        bucket = elements[workshop]
        categories = [
            category for (code, category) in debit_keys if code == workshop
            and (debit[(code, category)]["amount"] or debit[(code, category)]["quantity"])
        ]
        if not categories:
            continue
        workshop_rows.append(
            {
                "workshop": workshop,
                "workshop_label": workshop_label(workshop),
                "order_count": bucket["orders"],
                "category_count": len(categories),
                "categories": "、".join(categories),
                "quantity": bucket["quantity"],
                "material": money(bucket[ELEMENT_MATERIAL]),
                "wages": money(bucket[ELEMENT_WAGES]),
                "overhead": money(bucket[ELEMENT_OVERHEAD]),
                "amount": money(bucket["amount"]),
                "source": SOURCE_ORDER,
                "element_split": True,
            }
        )
    if pulp_rows:
        workshop_rows.append(
            {
                "workshop": PULP_WORKSHOP,
                "workshop_label": PULP_WORKSHOP,
                "order_count": len(pulp_rows),
                "category_count": 1,
                "categories": "、".join(
                    sorted({row["category"] for row in pulp_rows})
                ),
                "quantity": pulp_quantity,
                "material": pulp_material,
                "wages": pulp_wages,
                "overhead": pulp_overhead,
                "amount": pulp_amount,
                "source": SOURCE_PULP_SPLIT if pulp_split else SOURCE_PULP_FLOW,
                "element_split": bool(pulp_split),
            }
        )
    workshop_totals = {
        "order_count": sum((row["order_count"] for row in workshop_rows), 0),
        "quantity": sum((row["quantity"] for row in workshop_rows), ZERO),
        "material": money(sum((row["material"] for row in workshop_rows), ZERO)),
        "wages": money(sum((row["wages"] for row in workshop_rows), ZERO)),
        "overhead": money(sum((row["overhead"] for row in workshop_rows), ZERO)),
        "amount": money(sum((row["amount"] for row in workshop_rows), ZERO)),
    }
    category_rows = []
    for workshop, category in debit_keys:
        item = debit[(workshop, category)]
        if item["amount"] == ZERO and item["quantity"] == ZERO:
            continue
        category_rows.append(
            {
                "workshop_label": workshop_label(workshop),
                "category": category,
                "summary": debit_summary_text(period, workshop, category),
                "order_count": item["orders"],
                "quantity": item["quantity"],
                "amount": money(item["amount"]),
                "account_debit": inventory_account(category),
            }
        )

    return {
        "month": period.month.strftime("%Y-%m"),
        "rows": rows,
        "headers": ["摘要", "科目", "数量", "借方", "贷方"],
        "totals": {
            "debit": money(debit_total),
            "credit": money(credit_total),
            "difference": money(debit_total - credit_total),
            "balanced": money(debit_total - credit_total) == ZERO,
            "quantity": quantity_total,
            "line_count": len([row for row in rows if row["level"] == "line"]),
        },
        "summary": {
            "order_count": len(detail_rows),
            "order_amount": money(order_amount),
            "pulp_count": len(pulp_rows),
            "pulp_amount": pulp_amount,
            "voucher_amount": money(debit_total),
            "flow_amount": flow_amount,
            "duplicate_rows": dup_rows,
            "duplicate_amount": dup_amount,
            "zero_cost_orders": zero_cost_orders,
            "unclassified_orders": unclassified_orders,
            "residual_orders": residual_orders,
            "workshop_count": len([
                code for code in elements if code != UNALLOCATED
            ]),
            "element_totals": element_totals,
            #: 贷方要素里属于工单的部分（不含浆粕），方便页面把「工单 + 浆粕」拆开讲
            "order_element_totals": order_element_totals,
            "pulp_element_totals": dict(pulp_by_element),
            "pulp_split_ok": bool(pulp_split),
            "pulp_quantity": pulp_quantity,
        },
        "checks": checks,
        "detail_rows": detail_rows,
        "pulp_rows": pulp_rows,
        "workshop_rows": workshop_rows,
        "workshop_totals": workshop_totals,
        "category_rows": category_rows,
        "elements": list(PRODUCTION_ELEMENTS),
        "accounts": {
            "inventory": ACCOUNT_INVENTORY,
            "production": ACCOUNT_PRODUCTION,
            "elements": [production_element_account(item) for item in PRODUCTION_ELEMENTS],
        },
        "voucher_note": VOUCHER_NOTE,
        "months": [
            {"value": item.month.strftime("%Y-%m"), "label": str(item)}
            for item in AccountingPeriod.objects.order_by("month")
        ],
    }
