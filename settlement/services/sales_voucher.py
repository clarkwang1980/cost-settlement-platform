# -*- coding: utf-8 -*-
"""销售凭证 + 发出商品及销售成本统计表（成本口径，数据来自销售台账的「结转类型」）。

台账里「结转类型」的取值正好是会计口径的 4 类：
    当月发货当月开票结转成本 / 前期发货本期开票结转成本 /
    当月发货未开票挂发出商品 / 免费样品赠送计入销售费用

- 销售凭证：逐行列出凭证抬头、会计科目、结转类型、销售大类、出库期间、开票状态、
  开票期间、数量、金额（金额 = 销售成本，结转成本口径）。
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

#: 结转类型 → (凭证抬头, 会计科目)
VOUCHER_ACCOUNTS = {
    CARRY_CURRENT_INVOICED: ("销售成本结转凭证（当月发货当月开票）", "主营业务成本"),
    CARRY_PRIOR_INVOICED: ("销售成本结转凭证（前期发货本期开票）", "主营业务成本"),
    CARRY_UNINVOICED: ("发出商品结转凭证", "发出商品"),
    CARRY_SAMPLE: ("样品费用凭证", "销售费用"),
}

#: 开票状态中文（台账取值 -> 显示）
INVOICE_STATUS_LABELS = {
    "invoiced": "已开票",
    "uninvoiced": "未开票",
    "sample_free": "样品免费",
}

VOUCHER_HEADERS = (
    "抬头销售凭证", "会计科目", "结转类型", "销售大类", "出库期间",
    "开票状态", "开票期间", "数量", "金额",
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
    start = month_first(period.month)
    return [month_add(start, index) for index in range(count)]


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
    if line.carry_type == CARRY_PRIOR_INVOICED:
        return month_first(line.invoice_month) if line.invoice_month else None
    return month_first(line.out_month) if line.out_month else None


def _voucher_row(line) -> dict:
    carry = (line.carry_type or "").strip()
    title, account = VOUCHER_ACCOUNTS.get(carry, ("销售结转凭证", "—"))
    return {
        "voucher_month": month_label(_voucher_month(line)),
        "title": title,
        "account": account,
        "carry_type": carry or "—",
        "category": category_label(line.sales_category),
        "out_month": month_label(line.out_month),
        "invoice_status": INVOICE_STATUS_LABELS.get(
            line.invoice_status, line.invoice_status or "—"
        ),
        "invoice_month": month_label(line.invoice_month) or "—",
        "quantity": line.quantity or ZERO,
        "amount": line.cost_amount,
        "customer": line.customer_name or line.customer_code or "",
        "material_code": line.material_code or "",
        "material_name": line.material_name or "",
    }


def build_sales_voucher_data(period, month=None):
    """销售凭证：按「记账月份」筛选（默认当前期间月份）。

    返回 {"rows", "months", "month", "totals", "missing_cost_rows", "empty"}。
    """
    lines = list(_ledger_lines())
    month_map = defaultdict(int)
    for line in lines:
        value = _voucher_month(line)
        if value:
            month_map[value] += 1
    months = [
        {"value": month_label(value), "label": month_label(value), "row_count": month_map[value]}
        for value in sorted(month_map)
    ]
    selected = month or month_first(period.month)
    rows = [
        _voucher_row(line)
        for line in lines
        if _voucher_month(line) == selected
    ]
    rows.sort(key=lambda item: (item["carry_type"], item["category"], item["out_month"]))
    totals = {
        "quantity": sum((row["quantity"] for row in rows), ZERO),
        "amount": money(sum((row["amount"] or ZERO for row in rows), ZERO)),
        "row_count": len(rows),
    }
    missing = sum(1 for row in rows if row["amount"] is None)
    return {
        "rows": rows,
        "months": months,
        "month": month_label(selected),
        "headers": VOUCHER_HEADERS,
        "totals": totals,
        "missing_cost_rows": missing,
        "empty": not rows,
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
