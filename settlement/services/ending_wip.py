# -*- coding: utf-8 -*-
"""期末在产品（工单在制余额）汇总 —— 「成本核算 → 期末在产品」页面与导出共用。

口径（与「总成本表」的「在产品成本」块一致）：
- 数据来源：工单结算结果 WorkOrderCostResult（按生产工单核算）；
- 期末在产品 = 该工单结算结果里的 wip_* 三个科目：
  直接材料 wip_material / 人工费 wip_wages / 制造费用 wip_overhead，合计 wip_cost；
- 只列**还有在产品**的工单（在产品数量或在产品成本任一不为 0）；
- 车间按工单主档的「生产部门编号」归一（与车间结转同口径）；只列有期末在产品的车间；
- 车间顺序沿用车间结转的 sequence（兜底固定顺序）。
"""
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP

from settlement.models import OpeningWip, WorkOrder, WorkOrderCostResult
from settlement.services.imports import normalize_workshop
from settlement.services.total_cost import WORKSHOP_LABELS, _workshop_order

ZERO = Decimal("0.00")
CENT = Decimal("0.01")


#: 期初在产品表的 9 个成本科目（与「总成本表」期初块一致）
OPENING_WIP_COLUMNS = [
    "material", "water_power", "steam", "natural_gas", "dye_auxiliary",
    "consumables", "depreciation", "wages", "processing_fee",
]


def _wip_start_labels(period, order_nos):
    """每张工单「什么时候开始有的在产品」。

    规则（用户 2026-09-21 要求：我要知道这是什么时候开始有的在产品）：
      ① 该工单在本期或以前期间的**期初在产品**里出现过 → 期初就有在产品：
         有「来源期间」用来源期间（YYYY-MM）；没有来源期间（上传的期初台账）→ 显示「期初结转」；
      ② 否则 → 该工单在本期及以前期间**最早**有在产品的那个期间（YYYY-MM）。
    """
    labels = {}
    # ① 期初在产品：优先取带来源期间的
    openings = defaultdict(list)
    for item in OpeningWip.objects.filter(
        period__month__lte=period.month, order_no__in=order_nos
    ).select_related("source_period"):
        openings[item.order_no].append(item)
    for order_no, items in openings.items():
        sources = sorted(
            {o.source_period.month for o in items if o.source_period}
        )
        if sources:
            labels[order_no] = sources[0].strftime("%Y-%m")
        else:
            labels[order_no] = "期初结转"

    # ② 结算结果里最早有在产品的期间
    firsts = {}
    for r in (
        WorkOrderCostResult.objects.filter(
            period__month__lte=period.month, work_order__order_no__in=order_nos
        )
        .select_related("period", "work_order")
        .order_by("period__month")
    ):
        if Decimal(r.wip_cost or ZERO) == ZERO and Decimal(r.wip_quantity or ZERO) == ZERO:
            continue
        order_no = r.work_order.order_no
        if order_no not in firsts:
            firsts[order_no] = r.period.month
    for order_no, month in firsts.items():
        labels.setdefault(order_no, month.strftime("%Y-%m"))
    return labels


def _money(value) -> Decimal:
    return Decimal(value or ZERO).quantize(CENT, rounding=ROUND_HALF_UP)


def _product_lookup(period):
    """工单编号 -> (产品编码, 产品名称)：给结算结果里没带产品的工单兜底。

    优先级（先到先得）：① 工单结算结果（在调用处优先用）；② BOM 的「成品物料编号/名称」；
    ③ 投料记录的「主产出物料编号/名称」；④ 报工记录的「报工物料编号/名称」。
    （用户 2026-09-21：期末在产品明细里 200 行有 172 行产品编码/名称为空）
    """
    lookup = {}

    def _put(order_no, code, name):
        order_no = str(order_no or "").strip()
        code = str(code or "").strip()
        if not order_no or not code or order_no in lookup:
            return
        lookup[order_no] = (code, str(name or "").strip())

    from settlement.models import BomLine, MaterialIssue, WorkReport

    for order_no, code, name in (
        BomLine.objects.filter(period=period)
        .order_by("work_order_no", "id")
        .values_list("work_order_no", "material_code", "material_name")
    ):
        _put(order_no, code, name)
    for order_no, code, name in (
        MaterialIssue.objects.filter(period=period)
        .order_by("work_order_no", "id")
        .values_list("work_order_no", "main_output_code", "main_output_name")
    ):
        _put(order_no, code, name)
    for order_no, code, name in (
        WorkReport.objects.filter(period=period)
        .order_by("work_order_no", "id")
        .values_list("work_order_no", "material_code", "material_name")
    ):
        _put(order_no, code, name)
    return lookup


def build_ending_wip_rows(period, workshop="", material="", status=""):
    """期末在产品明细（一行一个还有在产品的工单）。

    筛选：workshop（车间编号，模糊包含）/ material（产品编码或名称，模糊包含）/
    status（工单状态，精确相等；用户 2026-09-21）。
    """
    rows = []
    products = _product_lookup(period)
    results = (
        WorkOrderCostResult.objects.filter(period=period)
        .select_related("work_order")
    )
    for r in results:
        material_amount = Decimal(r.wip_material or ZERO)
        wages = Decimal(r.wip_wages or ZERO)
        overhead = Decimal(r.wip_overhead or ZERO)
        cost = Decimal(r.wip_cost or ZERO)
        quantity = Decimal(r.wip_quantity or ZERO)
        # 只列还有在产品的工单（数量或金额任一不为 0）
        if cost == ZERO and quantity == ZERO:
            continue
        ws_code = normalize_workshop(r.work_order.production_dept_code or "")
        rows.append(
            {
                "start_period": "",
                "workshop_code": ws_code,
                "workshop": WORKSHOP_LABELS.get(ws_code, ws_code or "—"),
                "order_no": r.work_order.order_no,
                "order_status": r.work_order.status or "",
                "product_code": (
                    (r.product_code or "").strip()
                    or products.get(r.work_order.order_no, ("", ""))[0]
                ),
                "product_name": (
                    (r.product_name or "").strip()
                    or products.get(r.work_order.order_no, ("", ""))[1]
                ),
                "plan_quantity": Decimal(r.plan_quantity or ZERO),
                "finished_quantity": Decimal(r.finished_quantity or ZERO),
                "wip_quantity": quantity,
                "wip_material": _money(material_amount),
                "wip_wages": _money(wages),
                "wip_overhead": _money(overhead),
                "wip_cost": _money(cost),
            }
        )

    start_labels = _wip_start_labels(
        period, {row["order_no"] for row in rows}
    )
    for row in rows:
        row["start_period"] = start_labels.get(row["order_no"], "")

    if workshop:
        needle = workshop.strip()
        rows = [
            row
            for row in rows
            if needle in row["workshop_code"] or needle in row["workshop"]
        ]
    if material:
        needle = material.strip().lower()
        rows = [
            row
            for row in rows
            if needle
            in ("%s %s" % (row["product_code"], row["product_name"])).lower()
        ]
    if status:
        rows = [row for row in rows if (row["order_status"] or "") == status]

    order = _workshop_order(period)
    index = {code: i for i, code in enumerate(order)}
    rows.sort(
        key=lambda row: (
            index.get(row["workshop_code"], 98),
            row["workshop_code"],
            row["order_no"],
        )
    )
    return rows


def summarize_ending_wip(rows, period=None):
    """按车间汇总 + 合计（用于页面汇总表）：只列**有期末在产品**的车间。"""
    buckets = defaultdict(
        lambda: {
            "order_count": 0,
            "wip_quantity": ZERO,
            "wip_material": ZERO,
            "wip_wages": ZERO,
            "wip_overhead": ZERO,
            "wip_cost": ZERO,
        }
    )
    order_seen = []
    for row in rows:
        key = row["workshop_code"]
        item = buckets[key]
        item["workshop"] = row["workshop"]
        item["order_count"] += 1
        item["wip_quantity"] += row["wip_quantity"]
        item["wip_material"] += row["wip_material"]
        item["wip_wages"] += row["wip_wages"]
        item["wip_overhead"] += row["wip_overhead"]
        item["wip_cost"] += row["wip_cost"]
        if key not in order_seen:
            order_seen.append(key)

    workshop_rows = []
    total = {
        "order_count": 0,
        "wip_quantity": ZERO,
        "wip_material": ZERO,
        "wip_wages": ZERO,
        "wip_overhead": ZERO,
        "wip_cost": ZERO,
    }
    for key in order_seen:
        item = buckets[key]
        item["workshop_code"] = key
        item["amount"] = _money(item["wip_cost"])
        workshop_rows.append(item)
        for field in total:
            total[field] += item[field]
    total = {**total, "amount": _money(total["wip_cost"])}
    return workshop_rows, total


def build_ending_wip_data(period, workshop="", material="", status=""):
    """页面/导出共用。

    返回：
      · rows            筛选后的明细（页面表格 / 导出明细）；
      · all_rows        全量明细（用于汇总与合计，不受明细筛选影响）；
      · workshop_rows   各车间汇总（含期初在产品成本；车间＝期初∪期末的并集）；
      · total           汇总合计（全量）；
      · filtered_total  筛选后的合计（明细表金额徽标用）。
    """
    all_rows = build_ending_wip_rows(period)
    rows = build_ending_wip_rows(
        period, workshop=workshop, material=material, status=status
    )
    workshop_rows, total = summarize_ending_wip(all_rows, period)
    filtered_workshop_rows, filtered_total = summarize_ending_wip(rows, period)
    return {
        "rows": rows,
        "all_rows": all_rows,
        "workshop_rows": workshop_rows,
        "total": total,
        "filtered_total": filtered_total,
    }


def ending_wip_status_options(period):
    """工单状态下拉选项：本页明细里出现过的状态（按出现顺序）。"""
    seen = []
    for row in build_ending_wip_rows(period):
        value = row["order_status"] or ""
        if value and value not in seen:
            seen.append(value)
    return seen


def ending_wip_workshop_options(period):
    """车间下拉选项：本页明细里出现过的车间（按车间顺序）。"""
    rows = build_ending_wip_rows(period)
    seen = []
    for row in rows:
        if row["workshop_code"] not in [item["value"] for item in seen]:
            seen.append({"value": row["workshop_code"], "label": row["workshop"]})
    return seen
