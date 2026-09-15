# -*- coding: utf-8 -*-
"""销售凭证 + 发出商品及销售成本统计表（成本口径，数据来自销售台账的「结转类型」）。

台账里「结转类型」的取值正好是会计口径的 4 类：
    当月发货当月开票结转成本 / 前期发货本期开票结转成本 /
    当月发货未开票挂发出商品 / 免费样品赠送计入销售费用

- 销售结转成本凭证：只统计两类结转成本（当月发货当月开票结转成本 /
  前期发货本期开票结转成本），按「结转类型 × 销售大类」汇总，
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

#: 销售结转成本凭证只统计这两类结转成本
CARRY_COST_TYPES = (CARRY_CURRENT_INVOICED, CARRY_PRIOR_INVOICED)

#: 结转类型 → 会计科目
ACCOUNT_BY_CARRY = {
    CARRY_CURRENT_INVOICED: "主营业务成本",
    CARRY_PRIOR_INVOICED: "主营业务成本",
}

#: 开票状态中文（台账取值 -> 显示）
INVOICE_STATUS_LABELS = {
    "invoiced": "已开票",
    "uninvoiced": "未开票",
    "sample_free": "样品免费",
}

#: 销售大类显示顺序（用户指定），未列出的排在其后
CATEGORY_ORDER = ("纱线", "面料", "纤维", "其他成品")
ALL_LABEL = "全部"

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


def build_sales_voucher_summary(period, month=None, out_months=None,
                                invoice_months=None, statuses=None,
                                categories=None) -> dict:
    """销售结转成本凭证：按「结转类型 × 销售大类」汇总。

    只统计两类结转成本（当月发货当月开票 / 前期发货本期开票）；
    可按「记账月份 + 出库期间 + 开票状态 + 开票期间 + 销售大类」筛选，
    汇总数字跟随筛选条件重算。
    """
    lines = [
        line for line in _ledger_lines()
        if (line.carry_type or "").strip() in CARRY_COST_TYPES
    ]

    # 筛选选项（按全部数据统计）
    month_counter = defaultdict(int)
    out_counter = defaultdict(int)
    invoice_counter = defaultdict(int)
    status_counter = defaultdict(int)
    category_counter = defaultdict(int)
    for line in lines:
        value = _voucher_month(line)
        if value:
            month_counter[month_first(value)] += 1
        if line.out_month:
            out_counter[month_label(line.out_month)] += 1
        if line.invoice_month:
            invoice_counter[month_label(line.invoice_month)] += 1
        status_counter[line.invoice_status] += 1
        category_counter[category_label(line.sales_category)] += 1

    months = [
        {"value": month_label(value), "label": month_label(value),
         "row_count": month_counter[value]}
        for value in sorted(month_counter)
    ]
    out_options = _options(sorted(out_counter), out_counter)
    invoice_options = _options(sorted(invoice_counter), invoice_counter)
    status_options = _options(
        [value for value in ("invoiced", "uninvoiced", "sample_free")
         if value in status_counter],
        status_counter,
        labeler=lambda value: INVOICE_STATUS_LABELS.get(value, value),
    )
    category_options = _options(
        sorted(category_counter, key=_category_sort_key), category_counter
    )

    # 当前筛选
    selected = month or month_first(period.month)
    out_set = {value for value in (out_months or []) if value}
    invoice_set = {value for value in (invoice_months or []) if value}
    status_set = {value for value in (statuses or []) if value}
    category_set = {value for value in (categories or []) if value}
    has_filter = bool(out_set or invoice_set or status_set or category_set)

    out_label = "、".join(sorted(out_set)) or ALL_LABEL
    invoice_label = "、".join(sorted(invoice_set)) or ALL_LABEL
    status_label = (
        "、".join(INVOICE_STATUS_LABELS.get(v, v) for v in sorted(status_set))
        or ALL_LABEL
    )

    buckets = defaultdict(
        lambda: {"quantity": ZERO, "amount": ZERO, "rows": 0, "missing": 0}
    )
    for line in lines:
        if _voucher_month(line) != selected:
            continue
        if out_set and month_label(line.out_month) not in out_set:
            continue
        if invoice_set and month_label(line.invoice_month) not in invoice_set:
            continue
        if status_set and line.invoice_status not in status_set:
            continue
        category = category_label(line.sales_category)
        if category_set and category not in category_set:
            continue
        item = buckets[((line.carry_type or "").strip(), category)]
        item["quantity"] += line.quantity or ZERO
        item["rows"] += 1
        if line.cost_amount is None:
            item["missing"] += 1
        else:
            item["amount"] += Decimal(line.cost_amount)

    def _row(level, carry_type, category, quantity, amount, missing=0, rows=0):
        return {
            "level": level,
            "account": ACCOUNT_BY_CARRY.get(carry_type, "主营业务成本"),
            "carry_type": carry_type,
            "category": category,
            "out_month": out_label,
            "invoice_status": status_label,
            "invoice_month": invoice_label,
            "quantity": quantity,
            "amount": money(amount),
            "missing_cost_rows": missing,
            "source_rows": rows,
        }

    rows = []
    for carry in sorted({key[0] for key in buckets}, key=_carry_sort_key):
        sub_quantity = ZERO
        sub_amount = ZERO
        sub_rows = sub_missing = 0
        for category in sorted(
            (key[1] for key in buckets if key[0] == carry), key=_category_sort_key
        ):
            item = buckets[(carry, category)]
            rows.append(
                _row("category", carry, category, item["quantity"], item["amount"],
                     item["missing"], item["rows"])
            )
            sub_quantity += item["quantity"]
            sub_amount += item["amount"]
            sub_rows += item["rows"]
            sub_missing += item["missing"]
        rows.append(
            _row("subtotal", carry, "小计", sub_quantity, sub_amount, sub_missing, sub_rows)
        )

    total_quantity = sum((row["quantity"] for row in rows if row["level"] == "category"), ZERO)
    total_amount = sum(
        (row["amount"] for row in rows if row["level"] == "category"), ZERO
    )
    total_rows = sum(row["source_rows"] for row in rows if row["level"] == "category")
    total_missing = sum(
        row["missing_cost_rows"] for row in rows if row["level"] == "category"
    )
    rows.append(
        _row("total", "合计", "", total_quantity, total_amount, total_missing, total_rows)
    )

    return {
        "rows": rows,
        "months": months,
        "month": month_label(selected),
        "headers": VOUCHER_HEADERS,
        "totals": {
            "quantity": money(total_quantity),
            "amount": money(total_amount),
            "row_count": len(rows),
            "source_rows": total_rows,
            "missing_cost_rows": total_missing,
        },
        "out_options": out_options,
        "invoice_options": invoice_options,
        "status_options": status_options,
        "category_options": category_options,
        "selected_out": sorted(out_set),
        "selected_invoice": sorted(invoice_set),
        "selected_status": sorted(status_set),
        "selected_categories": sorted(category_set),
        "all_label": ALL_LABEL,
        "has_filter": has_filter,
        "empty": not any(row["level"] == "category" for row in rows),
    }


# ── 发出商品及销售成本统计表 ────────────────────────────────────────────

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
