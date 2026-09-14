# -*- coding: utf-8 -*-
"""损益表：销售台账逐行 + 当月费用按「终端销售收入占比」分摊。

口径（2026-09-14 与用户确认）：
- 收入 = 终端销售收入（射阳兰邦销售收入只作列展示，不参与计算）；
- 销售成本 = 平台按月末加权平均价算出的成本（台账行 cost_amount，结算后回填）；
- 销售毛利 = 终端销售收入 − 销售成本；销售毛利率 = 销售毛利 ÷ 终端销售收入；
- 用户录入的 7 项【当月发生额】按「本行终端销售收入 ÷ 当月终端销售收入合计」分摊到每一行
  （四舍五入到分，尾差调整到收入最大的一行，保证各列合计 = 录入金额）：
  税金及附加 / 管理费用 / 销售费用 / 财务费用 / 营业外收入 / 营业外支出 / 所得税费用；
- 营业利润 = 销售毛利 − 税金及附加 − 管理费用 − 销售费用 − 财务费用；
- 利润总额 = 营业利润 + 营业外收入 − 营业外支出；
- 净利润 = 利润总额 − 所得税费用；净利润率 = 净利润 ÷ 终端销售收入。
"""
from __future__ import annotations

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

#: 需要求和的键（合计行）
TOTAL_KEYS = (
    "revenue", "lb_amount", "quantity", "cost", "gross",
    "tax_surcharge", "admin_expense", "selling_expense", "finance_expense",
    "operating_profit", "non_operating_income", "non_operating_expense",
    "total_profit", "income_tax", "net_profit",
)

EXPENSE_TOTAL_FIELDS = tuple(field for field, _label in ALLOCATION_FIELDS)


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


def row_values(row: dict) -> list:
    """按 HEADERS 顺序取值（用于导出）。"""
    return [row.get(key) for key in ROW_KEYS]


def get_expense(period) -> ProfitStatementExpense:
    expense, _created = ProfitStatementExpense.objects.get_or_create(period=period)
    return expense


def ensure_ledger_costs(period) -> int:
    """台账行成本还没回填、且本期已有库存计价时，先按月末加权价回填（幂等）。"""
    pending = period.sales_ledger_lines.filter(
        terminal_amount__isnull=False, cost_amount__isnull=True
    ).exists()
    if pending and period.inventory_valuations.exists():
        return fill_sales_ledger_costs(period)
    return 0


def build_profit_statement(period, expense=None) -> dict:
    """组装损益表：返回 {headers, rows, totals, 以及若干提示标志}。"""
    backfilled = ensure_ledger_costs(period)
    expense = expense or get_expense(period)
    lines = list(period.sales_ledger_lines.order_by("row_no"))

    revenues = [_dec(line.terminal_amount) for line in lines]
    revenue_total = sum(revenues, ZERO)

    # —— 当月费用按「行终端收入 ÷ 当月终端收入合计」分摊（尾差给收入最大的行）——
    allocated = {field: [ZERO] * len(lines) for field in EXPENSE_TOTAL_FIELDS}
    if lines and revenue_total > ZERO:
        max_index = max(range(len(lines)), key=lambda i: revenues[i])
        for field in EXPENSE_TOTAL_FIELDS:
            total = _dec(getattr(expense, field))
            if total == ZERO:
                continue
            used = ZERO
            for index, revenue in enumerate(revenues):
                if revenue <= ZERO:
                    continue
                amount = _q(revenue * total / revenue_total)
                allocated[field][index] = amount
                used += amount
            residual = total - used
            if residual != ZERO and revenues[max_index] > ZERO:
                allocated[field][max_index] += residual

    rows = []
    sums = {key: ZERO for key in TOTAL_KEYS}
    missing_cost_rows = 0
    lb_only_rows = 0
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

        row = {
            "seq": line.row_no,
            "out_month": line.out_month.strftime("%Y-%m") if line.out_month else "",
            "invoice_month": (
                line.invoice_month.strftime("%Y-%m") if line.invoice_month else ""
            ),
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

    return {
        "period": period,
        "expense": expense,
        "headers": HEADERS,
        "rows": rows,
        "totals": totals,
        "revenue_total": revenue_total,
        "entered_totals": {
            field: _dec(getattr(expense, field)) for field in EXPENSE_TOTAL_FIELDS
        },
        "missing_cost_rows": missing_cost_rows,
        "lb_only_rows": lb_only_rows,
        "backfilled": backfilled,
        "has_valuation": period.inventory_valuations.exists(),
        "row_count": len(rows),
    }


def build_profit_statement_sheet(period):
    """导出用：(sheet 列表, 数据)。sheet = (工作表名, 表头, 数据行)。"""
    data = build_profit_statement(period)
    rows = [row_values(row) for row in data["rows"]]
    rows.append(row_values(data["totals"]))
    return [("损益表", list(HEADERS), rows)], data
