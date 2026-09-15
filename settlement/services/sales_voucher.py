# -*- coding: utf-8 -*-
"""销售凭证 + 发出商品及销售成本统计表（成本口径，数据来自销售台账的「结转类型」）。

台账里「结转类型」的取值正好是会计口径的 4 类：
    当月发货当月开票结转成本 / 前期发货本期开票结转成本 /
    当月发货未开票挂发出商品 / 免费样品赠送计入销售费用

- 销售结转成本凭证：统计三类结转（当月发货当月开票结转成本 /
  前期发货本期开票结转成本 / 当月发货未开票挂发出商品），按「结转类型 × 销售大类」汇总，
  列：会计科目、结转类型、销售大类、出库期间、开票状态、开票期间、数量、金额；
  可按 记账月份 / 出库期间 / 开票状态 / 开票期间 / 销售大类 筛选。
- 发出商品及销售成本统计表：项目分类 × 月份（从当前期间起 6 个月）。
"""
from __future__ import annotations

from calendar import monthrange
from collections import defaultdict
from datetime import date
from decimal import ROUND_HALF_UP, Decimal

from settlement.models import SalesLedgerLine, SalesLedgerLine as _Line  # noqa: F401
from settlement.services.profit_statement import category_label

ZERO = Decimal("0")
CENT = Decimal("0.01")

# ── 结转类型（与台账口径一致） ──────────────────────────────────────────
CARRY_CURRENT_INVOICED = "当月发货当月开票结转成本"
CARRY_PRIOR_INVOICED = "前期发货本期开票结转成本"
CARRY_UNINVOICED = "当月发货未开票挂发出商品"
CARRY_SAMPLE = "免费样品赠送计入销售费用"

#: 销售结转成本凭证统计的结转类型（每个类型都会列全 5 个销售大类）
CARRY_COST_TYPES = (
    CARRY_CURRENT_INVOICED,
    CARRY_PRIOR_INVOICED,
    CARRY_UNINVOICED,
)

#: 结转类型 → 会计科目
ACCOUNT_BY_CARRY = {
    CARRY_CURRENT_INVOICED: "主营业务成本",
    CARRY_PRIOR_INVOICED: "主营业务成本",
    CARRY_UNINVOICED: "发出商品",
}

#: 开票状态中文（台账取值 -> 显示）；三个选项固定出现在筛选里
INVOICE_STATUS_LABELS = {
    "invoiced": "已开发票",
    "uninvoiced": "未开发票",
    "sample_free": "样品免费",
}
INVOICE_STATUS_VALUES = ("invoiced", "uninvoiced", "sample_free")

#: 销售大类固定 5 个（各结转类型共用），每类都要出现（没有数据也显示 0）
CATEGORY_ORDER = ("纱线", "面料", "纤维", "其他成品", "原料")
ALL_LABEL = "全部"
BLANK_LABEL = "（空白）"          # 下拉里的空白选项：不限制该字段

#: 汇总表列（已去掉「抬头销售凭证」）
VOUCHER_HEADERS = (
    "会计科目", "结转类型", "销售大类", "出库期间", "开票状态", "开票期间", "数量", "金额",
)

#: 统计表的行分类：一级分类 -> 二级项目
STAT_ROWS = (
    ("一、本月销售成本", (
        "当月发货当月开票结转成本",
        "前期发货本期开票结转成本",
        "本月主营业务成本合计",
    )),
    ("二、发出商品逐月结转", (
        "上月期末发出商品未开票成本",
        "当月发货未开票挂发出商品",
        "前期发货本期开票结转成本",
        "当月期末发出商品未开票成本",
    )),
    ("三、其他", (
        "免费样品赠送计入销售费用",
    )),
)


def money(value) -> Decimal:
    return Decimal(value or ZERO).quantize(CENT, rounding=ROUND_HALF_UP)


def month_label(value) -> str:
    return value.strftime("%Y-%m") if value else ""


def month_first(value) -> date:
    """取该日期所在月的 1 号。"""
    return value.replace(day=1)


def month_add(month: date, count: int) -> date:
    """月份加减（month 须为某月 1 号）。"""
    total = month.year * 12 + (month.month - 1) + count
    return date(total // 12, total % 12 + 1, 1)


def month_axis(period, count: int = 6) -> list:
    """统计表的月份列：从当前期间所在月起连续 count 个月。"""
    start_month = month_first(period.month)
    return [month_add(start_month, index) for index in range(count)]


# ── 台账汇总 ────────────────────────────────────────────────────────────

def _ledger_lines():
    return SalesLedgerLine.objects.all().only(
        "row_no", "out_month", "invoice_month", "invoice_status", "carry_type",
        "sales_category", "order_no", "customer_code", "customer_name",
        "material_code", "material_name", "quantity", "cost_amount",
        "terminal_amount", "batch_id",
    ).order_by("out_month", "row_no")


def _voucher_month(line) -> date | None:
    """凭证的记账月份：前期发货本期开票按「开票月」，其余按「出库月」。"""
    if (line.carry_type or "").strip() == CARRY_PRIOR_INVOICED:
        return month_first(line.invoice_month) if line.invoice_month else None
    return month_first(line.out_month) if line.out_month else None


def _category_sort_key(name):
    if name in CATEGORY_ORDER:
        return (0, CATEGORY_ORDER.index(name), "")
    return (1, len(CATEGORY_ORDER), name)


def _carry_sort_key(carry):
    if carry in CARRY_COST_TYPES:
        return (0, CARRY_COST_TYPES.index(carry), "")
    return (1, len(CARRY_COST_TYPES), carry)


def _options(values, counter, labeler=None):
    """把取值的统计整理成筛选项（按给定顺序）。"""
    return [
        {
            "value": value,
            "label": (labeler or (lambda v: v))(value),
            "row_count": counter[value],
        }
        for value in values
    ]


#: 每个结转类型块的筛选项参数名（带块序号，保证三块互不干扰）
BLOCK_PARAM_TEMPLATES = ("vout{}", "vst{}", "vinv{}")


def block_param_names(index: int) -> dict:
    """第 index 块的三个筛选参数名。"""
    return {
        "out": BLOCK_PARAM_TEMPLATES[0].format(index),
        "status": BLOCK_PARAM_TEMPLATES[1].format(index),
        "invoice": BLOCK_PARAM_TEMPLATES[2].format(index),
    }


def default_block_filters(period) -> dict:
    """三块的默认筛选：前两块按「已开发票 + 当月开票期间」，第三块按「当月出库 + 不限制」。"""
    current = month_label(month_first(period.month))
    return {
        CARRY_CURRENT_INVOICED: {"out": "", "status": "invoiced", "invoice": current},
        CARRY_PRIOR_INVOICED: {"out": "", "status": "invoiced", "invoice": current},
        CARRY_UNINVOICED: {"out": current, "status": "", "invoice": ""},
    }


def build_sales_voucher_summary(period, month=None, block_filters=None) -> dict:
    """销售结转成本凭证：按「结转类型 × 销售大类」汇总，三个结转类型块各自独立筛选。

    - 统计三类结转：当月发货当月开票 / 前期发货本期开票 / 当月发货未开票挂发出商品
      （前两者记主营业务成本，后者记发出商品）；记账月份共用，按各类型的记账规则归属；
    - 每个结转类型都列全 5 个销售大类（纱线/面料/纤维/其他成品/原料），无数据时显示 0；
    - block_filters: {结转类型: {"out": 值, "status": 值, "invoice": 值}}；
      值为 None（未提交）→ 用该块默认值；值为 ""（选了「（空白）」）→ 该字段不限制；
      非法值按「不限制」处理。
    """
    lines = [
        line for line in _ledger_lines()
        if (line.carry_type or "").strip() in CARRY_COST_TYPES
    ]

    # 共用选项：记账月份 + 三个筛选的候选项
    month_counter = defaultdict(int)
    out_counter = defaultdict(int)
    invoice_counter = defaultdict(int)
    status_counter = defaultdict(int)
    for line in lines:
        value = _voucher_month(line)
        if value:
            month_counter[month_first(value)] += 1
        if line.out_month:
            out_counter[month_label(line.out_month)] += 1
        if line.invoice_month:
            invoice_counter[month_label(line.invoice_month)] += 1
        status_counter[line.invoice_status] += 1

    months = [
        {"value": month_label(value), "label": month_label(value),
         "row_count": month_counter[value]}
        for value in sorted(month_counter)
    ]
    out_values = sorted(out_counter)
    invoice_values = sorted(invoice_counter)
    out_options = (
        [{"value": "", "label": BLANK_LABEL}]
        + [{"value": v, "label": v, "row_count": out_counter[v]} for v in out_values]
    )
    invoice_options = (
        [{"value": "", "label": BLANK_LABEL}]
        + [{"value": v, "label": v, "row_count": invoice_counter[v]}
           for v in invoice_values]
    )
    status_options = [{"value": "", "label": BLANK_LABEL}]
    status_options += [
        {
            "value": value,
            "label": INVOICE_STATUS_LABELS[value],
            "row_count": status_counter.get(value, 0),
        }
        for value in INVOICE_STATUS_VALUES
    ]

    # 解析每块的筛选（None = 未提交 → 用默认；非法值 → 不限制）
    selected_month = month or month_first(period.month)
    defaults = default_block_filters(period)
    valid_out = set(out_values) | {""}
    valid_invoice = set(invoice_values) | {""}
    valid_status = set(INVOICE_STATUS_VALUES) | {""}

    blocks = []
    resolved = {}
    for index, carry in enumerate(CARRY_COST_TYPES):
        base = defaults[carry]
        supplied = (block_filters or {}).get(carry) or {}
        values = {}
        for key, valid in (("out", valid_out), ("status", valid_status),
                           ("invoice", valid_invoice)):
            raw_value = supplied.get(key, None)
            if raw_value is None:
                # 未提交 → 用该块默认值（默认值永远有效，即使当前数据里没有该月份）
                values[key] = base[key]
                continue
            value = str(raw_value)
            if value and value not in valid:
                # 用户提交了不存在的取值 → 按「不限制」处理
                value = ""
            values[key] = value
        resolved[carry] = values
        blocks.append(
            {
                "index": index,
                "carry_type": carry,
                "account": ACCOUNT_BY_CARRY.get(carry, "主营业务成本"),
                "params": block_param_names(index),
                "out_value": values["out"],
                "status_value": values["status"],
                "invoice_value": values["invoice"],
            }
        )

    # 按块汇总
    def _blank():
        return {"quantity": ZERO, "amount": ZERO, "rows": 0, "missing": 0}

    rows = []
    grand = _blank()
    for block in blocks:
        carry = block["carry_type"]
        values = resolved[carry]
        buckets = defaultdict(_blank)
        for line in lines:
            if (line.carry_type or "").strip() != carry:
                continue
            if _voucher_month(line) != selected_month:
                continue
            if values["out"] and month_label(line.out_month) != values["out"]:
                continue
            if values["status"] and line.invoice_status != values["status"]:
                continue
            if values["invoice"] and month_label(line.invoice_month) != values["invoice"]:
                continue
            item = buckets[category_label(line.sales_category)]
            item["quantity"] += line.quantity or ZERO
            item["rows"] += 1
            if line.cost_amount is None:
                item["missing"] += 1
            else:
                item["amount"] += Decimal(line.cost_amount)

        # 该块的筛选当前值 + 参数名，挂在每一行上（模板按行渲染下拉并回显选中项）
        block_info = {
            "block_index": block["index"],
            "block_params": block["params"],
            "block_out": values["out"],
            "block_status": values["status"],
            "block_invoice": values["invoice"],
        }
        subtotal = _blank()
        for category in CATEGORY_ORDER:
            item = buckets.get(category) or _blank()
            rows.append(
                {
                    "level": "category",
                    "account": block["account"],
                    "carry_type": carry,
                    "category": category,
                    "quantity": item["quantity"],
                    "amount": money(item["amount"]),
                    "missing_cost_rows": item["missing"],
                    "source_rows": item["rows"],
                    **block_info,
                }
            )
            for key in subtotal:
                subtotal[key] += item[key]
        rows.append(
            {
                "level": "subtotal",
                "account": block["account"],
                "carry_type": carry,
                "category": "小计",
                "quantity": subtotal["quantity"],
                "amount": money(subtotal["amount"]),
                "missing_cost_rows": subtotal["missing"],
                "source_rows": subtotal["rows"],
                **block_info,
            }
        )
        for key in grand:
            grand[key] += subtotal[key]

    rows.append(
        {
            "level": "total",
            "block_index": -1,           # 合计行不属于任何一块（不带表内筛选）
            "block_params": None,
            "block_out": "",
            "block_status": "",
            "block_invoice": "",
            "account": "",
            "carry_type": "合计",
            "category": "",
            "quantity": grand["quantity"],
            "amount": money(grand["amount"]),
            "missing_cost_rows": grand["missing"],
            "source_rows": grand["rows"],
        }
    )

    return {
        "rows": rows,
        "blocks": blocks,
        "months": months,
        "month": month_label(selected_month),
        "headers": VOUCHER_HEADERS,
        "totals": {
            "quantity": money(grand["quantity"]),
            "amount": money(grand["amount"]),
            "row_count": len(rows),
            "source_rows": grand["rows"],
            "missing_cost_rows": grand["missing"],
        },
        "out_options": out_options,
        "invoice_options": invoice_options,
        "status_options": status_options,
        "all_label": ALL_LABEL,
        "blank_label": BLANK_LABEL,
    }


def build_sales_cost_statistics(period, month_count: int = 6) -> dict:
    """发出商品及销售成本统计表：项目分类 × 月份（成本口径）。

    口径：
    - 当月发货当月开票结转成本：出库月 = 该月 且 结转类型 = 当月发货当月开票
    - 前期发货本期开票结转成本：开票月 = 该月 且 结转类型 = 前期发货本期开票
    - 上月期末发出商品未开票成本 = 上月末余额（按「期末 = 上月末 + 当月发货未开票 − 前期发货本期开票」逐月滚动）
    - 当月发货未开票挂发出商品：出库月 = 该月 且 结转类型 = 当月发货未开票
    - 免费样品赠送计入销售费用：出库月 = 该月 且 结转类型 = 免费样品
    """
    axis = month_axis(period, month_count)
    by_out = defaultdict(lambda: ZERO)      # (月份, 结转类型) -> 成本
    by_invoice = defaultdict(lambda: ZERO)  # (月份, 结转类型) -> 成本
    missing_rows = defaultdict(int)         # 月份 -> 无成本行数
    for line in _ledger_lines():
        carry = (line.carry_type or "").strip()
        cost = line.cost_amount
        out_value = month_first(line.out_month) if line.out_month else None
        invoice_value = month_first(line.invoice_month) if line.invoice_month else None
        if out_value:
            if cost is None:
                missing_rows[out_value] += 1
            else:
                by_out[(out_value, carry)] += Decimal(cost)
        if invoice_value and cost is not None:
            by_invoice[(invoice_value, carry)] += Decimal(cost)

    # 期初余额：从最早有数据的月份滚动到统计起始月之前
    all_months = sorted({value for value, _carry in by_out} | {value for value, _carry in by_invoice})
    balance = ZERO
    if all_months:
        cursor = all_months[0]
        while cursor < axis[0]:
            balance += by_out[(cursor, CARRY_UNINVOICED)]
            balance -= by_invoice[(cursor, CARRY_PRIOR_INVOICED)]
            cursor = month_add(cursor, 1)

    months = []
    for value in axis:
        current_invoiced = money(by_out[(value, CARRY_CURRENT_INVOICED)])
        prior_invoiced = money(by_invoice[(value, CARRY_PRIOR_INVOICED)])
        shipped_uninvoiced = money(by_out[(value, CARRY_UNINVOICED)])
        sample = money(by_out[(value, CARRY_SAMPLE)])
        opening = money(balance)
        ending = money(opening + shipped_uninvoiced - prior_invoiced)
        months.append(
            {
                "value": month_label(value),
                "label": "%d年%d月" % (value.year, value.month),
                "cells": {
                    "当月发货当月开票结转成本": current_invoiced,
                    "前期发货本期开票结转成本": prior_invoiced,
                    "本月主营业务成本合计": money(current_invoiced + prior_invoiced),
                    "上月期末发出商品未开票成本": opening,
                    "当月发货未开票挂发出商品": shipped_uninvoiced,
                    "当月期末发出商品未开票成本": ending,
                    "免费样品赠送计入销售费用": sample,
                },
                "missing_cost_rows": missing_rows.get(value, 0),
                "has_data": bool(
                    current_invoiced or prior_invoiced or shipped_uninvoiced or sample
                ),
            }
        )
        balance = ending

    # 展开成「一级分类 + 二级项目」的平铺行，便于模板渲染
    BOLD_ITEMS = {
        "本月主营业务成本合计",
        "当月期末发出商品未开票成本",
    }
    flat = []
    for group, items in STAT_ROWS:
        flat.append({"level": "group", "label": group, "cells": None, "is_total": False})
        for item in items:
            flat.append(
                {
                    "level": "item",
                    "label": item,
                    "cells": [month["cells"][item] for month in months],
                    "is_total": item in BOLD_ITEMS,
                }
            )
    return {
        "months": months,
        "rows": flat,
        "month_headers": [month["label"] for month in months],
        "period": period,
    }
