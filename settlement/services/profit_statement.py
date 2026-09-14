# -*- coding: utf-8 -*-
"""损益表：销售台账逐行 + 按月录入的当月费用，按「当月终端收入占比」分摊。

口径（2026-09-14 与用户确认）：
- 收入 = 终端销售收入（射阳兰邦销售收入只作列展示，不参与计算）；
- 销售成本 = 平台按月末加权平均价算出的成本（台账行 cost_amount，结算后回填）；
- 销售毛利 = 终端销售收入 − 销售成本；销售毛利率 = 销售毛利 ÷ 终端销售收入；
- 税金及附加 / 管理费用 / 销售费用 / 财务费用 / 营业外收入 / 营业外支出 / 所得税费用
  按【出库月份】分别录入（4 月、5 月…… 各录一套），每行按
  「本行终端销售收入 ÷ 该行所属出库月份的终端销售收入合计」分摊，
  四舍五入到分、尾差计入该月收入最大的一行 ⇒ 每月各项分摊合计 = 该月录入金额；
- 营业利润 = 销售毛利 − 税金及附加 − 管理费用 − 销售费用 − 财务费用；
- 利润总额 = 营业利润 + 营业外收入 − 营业外支出；
- 净利润 = 利润总额 − 所得税费用；净利润率 = 净利润 ÷ 终端销售收入；
- 汇总可按【出库期间】或【开票期间】分组（只改分组口径，每行数字不变），
  分组层级：月份 → 销售渠道（国内销售/国外销售/未分类）→ 销售大类。
"""
from __future__ import annotations

from collections import defaultdict
from decimal import ROUND_HALF_UP, Decimal

from settlement.models import ProfitStatementExpense
from settlement.services.sales_ledger import fill_sales_ledger_costs

ZERO = Decimal("0")
CENT = Decimal("0.01")
PCT_Q = Decimal("0.01")

#: 分摊项：模型字段名 -> 页面/表头名称
ALLOCATION_FIELDS = (
    ("tax_surcharge", "税金及附加"),
    ("admin_expense", "管理费用"),
    ("selling_expense", "销售费用"),
    ("finance_expense", "财务费用"),
    ("non_operating_income", "营业外收入"),
    ("non_operating_expense", "营业外支出"),
    ("income_tax", "所得税费用"),
)

#: 表头（与 ROW_KEYS 一一对应）
HEADERS = (
    "序号", "出库期间", "开票期间", "销售渠道", "销售公司", "订单单号", "客户名称",
    "销售大类", "产品编码", "产品名称", "规格型号", "计量单位",
    "终端销售收入", "终端销售单价", "射阳兰邦销售收入", "射阳兰邦销售单价", "销售数量",
    "销售成本单价", "销售成本", "销售毛利", "销售毛利率(%)",
    "税金及附加", "管理费用", "销售费用", "财务费用", "营业利润",
    "营业外收入", "营业外支出", "利润总额", "所得税费用", "净利润", "净利润率(%)",
)

#: 行/合计字典的键（顺序与 HEADERS 一致）
ROW_KEYS = (
    "seq", "out_month", "invoice_month", "channel", "company", "order_no", "customer",
    "category", "code", "name", "spec", "uom",
    "revenue", "revenue_price", "lb_amount", "lb_price", "quantity",
    "cost_price", "cost", "gross", "gross_margin",
    "tax_surcharge", "admin_expense", "selling_expense", "finance_expense",
    "operating_profit",
    "non_operating_income", "non_operating_expense", "total_profit",
    "income_tax", "net_profit", "net_margin",
)

#: 需要求和的键（合计行 / 汇总行）
TOTAL_KEYS = (
    "revenue", "lb_amount", "quantity", "cost", "gross",
    "tax_surcharge", "admin_expense", "selling_expense", "finance_expense",
    "operating_profit", "non_operating_income", "non_operating_expense",
    "total_profit", "income_tax", "net_profit",
)

EXPENSE_TOTAL_FIELDS = tuple(field for field, _label in ALLOCATION_FIELDS)

#: 渠道归集：台账留空的一律进「未分类」，保证行数不丢
UNCLASSIFIED_CHANNEL = "未分类"
CHANNEL_ORDER = ("国内销售", "国外销售", UNCLASSIFIED_CHANNEL)
UNINVOICED_LABEL = "(未开票)"

#: 汇总口径：按出库期间 / 按开票期间
BASIS_OUT = "out"
BASIS_INVOICE = "invoice"
BASIS_CHOICES = ((BASIS_OUT, "按出库期间"), (BASIS_INVOICE, "按开票期间"))
BASIS_LABELS = dict(BASIS_CHOICES)

#: 汇总表指标（键 -> 表头）
SUMMARY_METRICS = (
    ("quantity", "销售数量"),
    ("revenue", "终端销售收入"),
    ("lb_amount", "射阳兰邦销售收入"),
    ("cost", "销售成本"),
    ("gross", "销售毛利"),
    ("gross_margin", "销售毛利率(%)"),
    ("tax_surcharge", "税金及附加"),
    ("admin_expense", "管理费用"),
    ("selling_expense", "销售费用"),
    ("finance_expense", "财务费用"),
    ("operating_profit", "营业利润"),
    ("non_operating_income", "营业外收入"),
    ("non_operating_expense", "营业外支出"),
    ("total_profit", "利润总额"),
    ("income_tax", "所得税费用"),
    ("net_profit", "净利润"),
    ("net_margin", "净利润率(%)"),
)
SUMMARY_HEADERS = ("月份", "销售渠道", "销售大类") + tuple(
    label for _key, label in SUMMARY_METRICS
)
SUMMARY_KEYS = ("month", "channel", "category") + tuple(
    key for key, _label in SUMMARY_METRICS
)

#: 月度 × 渠道 总览的指标
PIVOT_METRICS = (
    "revenue", "cost", "gross", "gross_margin",
    "operating_profit", "net_profit", "net_margin",
)
PIVOT_HEADERS = ("月份", "销售渠道") + tuple(
    label for key, label in SUMMARY_METRICS if key in PIVOT_METRICS
)

#: 页面顶端「期间 × 销售类别」汇总表的指标（按用户 2026-09-14 指定的 8 列）
CATEGORY_SUMMARY_METRICS = (
    ("revenue", "销售收入"),
    ("cost", "销售成本"),
    ("gross", "销售毛利"),
    ("gross_margin", "销售毛利率(%)"),
    ("operating_profit", "营业利润"),
    ("total_profit", "利润总额"),
    ("net_profit", "净利润"),
    ("net_margin", "净利润率(%)"),
)
CATEGORY_SUMMARY_HEADERS = ("期间", "销售类别") + tuple(
    label for _key, label in CATEGORY_SUMMARY_METRICS
)
CATEGORY_SUMMARY_KEYS = ("month", "category") + tuple(
    key for key, _label in CATEGORY_SUMMARY_METRICS
)


def _dec(value) -> Decimal:
    if value is None or value == "":
        return ZERO
    return Decimal(value)


def _q(value: Decimal) -> Decimal:
    return Decimal(value).quantize(CENT, rounding=ROUND_HALF_UP)


def _pct(numerator: Decimal, denominator: Decimal):
    """百分数（保留2位）；分母为 0/空时返回 None（页面显示 —）。"""
    if denominator is None or Decimal(denominator) == ZERO:
        return None
    return (Decimal(numerator) / Decimal(denominator) * Decimal("100")).quantize(
        PCT_Q, rounding=ROUND_HALF_UP
    )


def month_label(value) -> str:
    return value.strftime("%Y-%m") if value else ""


def row_values(row: dict) -> list:
    """按 HEADERS 顺序取值（用于导出）。"""
    return [row.get(key) for key in ROW_KEYS]


def summary_row_values(row: dict) -> list:
    """按 SUMMARY_HEADERS 顺序取值（用于导出）。"""
    return [row.get(key) for key in SUMMARY_KEYS]


def category_summary_row_values(row: dict) -> list:
    """按 CATEGORY_SUMMARY_HEADERS 顺序取值。"""
    return [row.get(key) for key in CATEGORY_SUMMARY_KEYS]


# ———————————————————————— 费用录入 ————————————————————————


def get_expense(period, month) -> ProfitStatementExpense:
    expense, _created = ProfitStatementExpense.objects.get_or_create(
        period=period, month=month
    )
    return expense


def expense_models(period) -> dict:
    """{月首日期: ProfitStatementExpense}。"""
    return {
        expense.month: expense
        for expense in ProfitStatementExpense.objects.filter(period=period)
    }


def expense_entries(period) -> dict:
    """{月首日期: {字段: Decimal}}（只含已录入的月份，未录入的按 0 处理）。"""
    return {
        month: {
            field: _dec(getattr(expense, field)) for field, _label in ALLOCATION_FIELDS
        }
        for month, expense in expense_models(period).items()
    }


def ledger_months(period) -> list:
    """台账里出现过的【出库月份】（月首日期，升序）——费用按月录入的依据。"""
    months = {
        value.replace(day=1)
        for value in period.sales_ledger_lines.values_list("out_month", flat=True)
        if value
    }
    return sorted(months)


def ensure_ledger_costs(period) -> int:
    """台账行成本还没回填、且本期已有库存计价时，先按月末加权价回填（幂等）。"""
    pending = period.sales_ledger_lines.filter(
        terminal_amount__isnull=False, cost_amount__isnull=True
    ).exists()
    if pending and period.inventory_valuations.exists():
        return fill_sales_ledger_costs(period)
    return 0


# ———————————————————————— 明细 ————————————————————————


def _allocate(month_of_row, revenues, expenses) -> dict:
    """按「本行收入 ÷ 所属出库月收入合计 × 该月录入金额」逐行分摊。

    尾差计入该月收入最大的一行，保证每月各项分摊合计 = 该月录入金额；
    没有录入该月费用时，该月所有行分摊额为 0。
    """
    allocated = {field: [ZERO] * len(revenues) for field in EXPENSE_TOTAL_FIELDS}
    by_month = defaultdict(list)
    for index, month in enumerate(month_of_row):
        by_month[month].append(index)

    for month, indexes in by_month.items():
        entry = expenses.get(month)
        if not entry:
            continue
        month_revenue = sum((revenues[i] for i in indexes), ZERO)
        if month_revenue <= ZERO:
            continue
        biggest = max(indexes, key=lambda i: revenues[i])
        for field in EXPENSE_TOTAL_FIELDS:
            total = _q(_dec(entry.get(field)))
            if total == ZERO:
                continue
            used = ZERO
            for index in indexes:
                revenue = revenues[index]
                if revenue <= ZERO:
                    continue
                amount = _q(revenue * total / month_revenue)
                allocated[field][index] = amount
                used += amount
            residual = total - used
            if residual != ZERO and revenues[biggest] > ZERO:
                allocated[field][biggest] += residual
    return allocated


def build_profit_statement(period, expenses=None) -> dict:
    """组装损益表：返回 {headers, rows, totals, months(按月录入信息), 若干提示标志}。"""
    backfilled = ensure_ledger_costs(period)
    if expenses is None:
        expenses = expense_entries(period)
    models = expense_models(period)
    lines = list(period.sales_ledger_lines.order_by("row_no"))

    revenues = [_dec(line.terminal_amount) for line in lines]
    month_of_row = [
        line.out_month.replace(day=1) if line.out_month else None for line in lines
    ]
    revenue_total = sum(revenues, ZERO)

    allocated = _allocate(month_of_row, revenues, expenses)

    rows = []
    sums = {key: ZERO for key in TOTAL_KEYS}
    missing_cost_rows = 0
    lb_only_rows = 0
    month_stats = defaultdict(lambda: {"row_count": 0, "revenue": ZERO})

    for index, line in enumerate(lines):
        revenue = revenues[index]
        lb_amount = _dec(line.lb_amount)
        cost = _dec(line.cost_amount)
        gross = _q(revenue - cost)

        tax = allocated["tax_surcharge"][index]
        admin = allocated["admin_expense"][index]
        selling = allocated["selling_expense"][index]
        finance = allocated["finance_expense"][index]
        nonop_in = allocated["non_operating_income"][index]
        nonop_out = allocated["non_operating_expense"][index]
        income_tax = allocated["income_tax"][index]

        operating_profit = _q(gross - tax - admin - selling - finance)
        total_profit = _q(operating_profit + nonop_in - nonop_out)
        net_profit = _q(total_profit - income_tax)

        if revenue <= ZERO and lb_amount > ZERO:
            lb_only_rows += 1
        if revenue > ZERO and line.cost_amount is None:
            missing_cost_rows += 1

        out_month = month_of_row[index]
        if out_month:
            month_stats[out_month]["row_count"] += 1
            month_stats[out_month]["revenue"] += revenue

        row = {
            "seq": line.row_no,
            "out_month": month_label(line.out_month),
            "invoice_month": month_label(line.invoice_month),
            "channel": line.channel,
            "company": line.company,
            "order_no": line.order_no,
            "customer": line.customer_name or line.customer_code,
            "category": line.sales_category,
            "code": line.material_code,
            "name": line.material_name,
            "spec": line.specification,
            "uom": line.uom,
            "revenue": revenue,
            "revenue_price": line.terminal_unit_price,
            "lb_amount": lb_amount,
            "lb_price": line.lb_unit_price,
            "quantity": _dec(line.quantity),
            "cost_price": line.cost_price,
            "cost": cost,
            "gross": gross,
            "gross_margin": _pct(gross, revenue),
            "tax_surcharge": tax,
            "admin_expense": admin,
            "selling_expense": selling,
            "finance_expense": finance,
            "operating_profit": operating_profit,
            "non_operating_income": nonop_in,
            "non_operating_expense": nonop_out,
            "total_profit": total_profit,
            "income_tax": income_tax,
            "net_profit": net_profit,
            "net_margin": _pct(net_profit, revenue),
            # 汇总用维度（明细表不显示）
            "month_out_key": month_label(out_month),
            "month_invoice_key": month_label(line.invoice_month) or UNINVOICED_LABEL,
            "channel_key": (line.channel or "").strip() or UNCLASSIFIED_CHANNEL,
            "category_key": (line.sales_category or "").strip() or UNCLASSIFIED_CHANNEL,
        }
        for key in TOTAL_KEYS:
            sums[key] += row[key]
        rows.append(row)

    totals = {key: "" for key in ROW_KEYS}
    for key in TOTAL_KEYS:
        totals[key] = _q(sums[key])
    totals["revenue_price"] = None
    totals["cost_price"] = None
    totals["gross_margin"] = _pct(sums["gross"], sums["revenue"])
    totals["net_margin"] = _pct(sums["net_profit"], sums["revenue"])

    months = []
    for month in sorted(month_stats):
        stats = month_stats[month]
        expense = models.get(month)
        entered = expenses.get(month)
        months.append(
            {
                "month": month,
                "key": month_label(month),
                "row_count": stats["row_count"],
                "revenue": _q(stats["revenue"]),
                "has_revenue": stats["revenue"] > ZERO,
                "entered": entered is not None,
                "has_values": bool(
                    entered
                    and any(_dec(entered.get(field)) != ZERO for field, _label in ALLOCATION_FIELDS)
                ),
                "values": {
                    field: _q(_dec((entered or {}).get(field)))
                    for field, _label in ALLOCATION_FIELDS
                },
                "updated_at": expense.updated_at if expense else None,
                "updated_by": expense.updated_by if expense else None,
            }
        )

    return {
        "period": period,
        "headers": HEADERS,
        "rows": rows,
        "totals": totals,
        "months": months,
        "revenue_total": revenue_total,
        "expense_totals": {
            field: _q(
                sum((_dec(entry.get(field)) for entry in expenses.values()), ZERO)
            )
            for field, _label in ALLOCATION_FIELDS
        },
        "missing_cost_rows": missing_cost_rows,
        "lb_only_rows": lb_only_rows,
        "backfilled": backfilled,
        "has_valuation": period.inventory_valuations.exists(),
        "row_count": len(rows),
        "no_revenue_months": [
            item["key"] for item in months if item["entered"] and not item["has_revenue"]
        ],
        "months_without_expense": [
            item["key"] for item in months if not item["entered"]
        ],
    }


# ———————————————————————— 汇总（月 → 渠道 → 大类）————————————————————————


def _accumulate(rows) -> dict:
    acc = {key: ZERO for key in TOTAL_KEYS}
    for row in rows:
        for key in TOTAL_KEYS:
            acc[key] += row[key]
    metrics = {key: _q(value) for key, value in acc.items()}
    metrics["gross_margin"] = _pct(metrics["gross"], metrics["revenue"])
    metrics["net_margin"] = _pct(metrics["net_profit"], metrics["revenue"])
    return metrics


def _channel_sort_key(channel: str):
    if channel in CHANNEL_ORDER:
        return (0, CHANNEL_ORDER.index(channel), "")
    return (1, len(CHANNEL_ORDER), channel)


def _month_sort_key(label: str):
    # “(未开票)”固定排在最后
    return (label == UNINVOICED_LABEL, label)


def build_summary(rows, basis: str = BASIS_OUT) -> dict:
    """按 月份 → 销售渠道 → 销售大类 汇总；含渠道小计、月度小计与总计。"""
    key_field = "month_invoice_key" if basis == BASIS_INVOICE else "month_out_key"
    tree = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
    for row in rows:
        tree[row[key_field]][row["channel_key"]][row["category_key"]].append(row)

    flat = []
    grand_rows = []
    month_blocks = []

    for month in sorted(tree, key=_month_sort_key):
        channels = tree[month]
        month_rows = []
        channel_blocks = []
        for channel in sorted(channels, key=_channel_sort_key):
            categories = channels[channel]
            channel_items = []
            channel_rows = []

            def _category_sort_key(name, _categories=categories):
                total = sum((item["revenue"] for item in _categories[name]), ZERO)
                return (-total, name)

            for category in sorted(categories, key=_category_sort_key):
                items = categories[category]
                channel_items.extend(items)
                row = {
                    "level": "category",
                    "month": month,
                    "channel": channel,
                    "category": category,
                    **_accumulate(items),
                }
                flat.append(row)
                channel_rows.append(row)

            channel_total = _accumulate(channel_items)
            channel_total_row = {
                "level": "channel",
                "month": month,
                "channel": channel,
                "category": "小计",
                **channel_total,
            }
            flat.append(channel_total_row)
            channel_blocks.append(
                {"channel": channel, "total": channel_total_row, "rows": channel_rows}
            )
            month_rows.extend(channel_items)

        month_total_row = {
            "level": "month",
            "month": month,
            "channel": "小计",
            "category": "",
            **_accumulate(month_rows),
        }
        flat.append(month_total_row)
        month_blocks.append(
            {"month": month, "total": month_total_row, "channels": channel_blocks}
        )
        grand_rows.extend(month_rows)

    grand_total = _accumulate(grand_rows)
    grand_total_row = {
        "level": "grand",
        "month": "合计",
        "channel": "",
        "category": "",
        **grand_total,
    }
    flat.append(grand_total_row)

    return {
        "basis": basis,
        "basis_label": BASIS_LABELS.get(basis, BASIS_LABELS[BASIS_OUT]),
        "months": month_blocks,
        "grand_total": grand_total,
        "grand_total_row": grand_total_row,
        "rows": flat,
        "headers": SUMMARY_HEADERS,
        "keys": SUMMARY_KEYS,
        "pivot_rows": [row for row in flat if row["level"] in ("channel", "month", "grand")],
        "pivot_headers": PIVOT_HEADERS,
        "pivot_keys": ("month", "channel") + PIVOT_METRICS,
    }


def build_category_summary(rows, basis: str = BASIS_OUT) -> dict:
    """页面顶端汇总表：期间 → 销售类别（销售大类），每个期间一个小计，末尾全部合计。

    列固定为用户指定的 8 个指标：销售收入、销售成本、销售毛利、销售毛利率、
    营业利润、利润总额、净利润、净利润率。期间口径随「按出库期间 / 按开票期间」切换。
    """
    key_field = "month_invoice_key" if basis == BASIS_INVOICE else "month_out_key"
    tree = defaultdict(lambda: defaultdict(list))
    for row in rows:
        tree[row[key_field]][row["category_key"]].append(row)

    flat = []
    month_blocks = []
    grand_items = []

    for month in sorted(tree, key=_month_sort_key):
        categories = tree[month]

        def _category_sort_key(name, _categories=categories):
            total = sum((item["revenue"] for item in _categories[name]), ZERO)
            return (-total, name)

        month_items = []
        month_rows = []
        for category in sorted(categories, key=_category_sort_key):
            items = categories[category]
            month_items.extend(items)
            row = {
                "level": "category",
                "month": month,
                "category": category,
                **_accumulate(items),
            }
            flat.append(row)
            month_rows.append(row)

        subtotal = {
            "level": "month",
            "month": month,
            "category": "小计",
            **_accumulate(month_items),
        }
        flat.append(subtotal)
        month_blocks.append({"month": month, "total": subtotal, "rows": month_rows})
        grand_items.extend(month_items)

    grand_total_row = {
        "level": "grand",
        "month": "合计",
        "category": "",
        **_accumulate(grand_items),
    }
    flat.append(grand_total_row)

    return {
        "basis": basis,
        "basis_label": BASIS_LABELS.get(basis, BASIS_LABELS[BASIS_OUT]),
        "rows": flat,
        "headers": CATEGORY_SUMMARY_HEADERS,
        "keys": CATEGORY_SUMMARY_KEYS,
        "months": month_blocks,
        "grand_total_row": grand_total_row,
    }


def build_profit_statement_sheet(period, expenses=None):
    """导出用：(sheet 列表, 数据)。

    sheet = (工作表名, 表头, 数据行)。工作表：损益表明细 + 两张汇总（出库期间/开票期间）。
    """
    data = build_profit_statement(period, expenses=expenses)
    detail_rows = [row_values(row) for row in data["rows"]]
    detail_rows.append(row_values(data["totals"]))
    sheets = [("损益表", list(HEADERS), detail_rows)]
    for basis, title in (
        (BASIS_OUT, "汇总-按出库期间"),
        (BASIS_INVOICE, "汇总-按开票期间"),
    ):
        summary = build_summary(data["rows"], basis)
        sheets.append(
            (
                title,
                list(SUMMARY_HEADERS),
                [summary_row_values(row) for row in summary["rows"]],
            )
        )
    return sheets, data
