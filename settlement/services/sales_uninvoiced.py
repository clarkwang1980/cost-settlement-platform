# -*- coding: utf-8 -*-
"""月末已出库未开票 + 发出商品滚动（方案A：全部按「销售台账」口径，2026-09-21 用户确认）。

台账自带「结转类型」，四类正好是会计口径：
  · 当月发货当月开票结转成本   → 本期已开票结转
  · 前期发货本期开票结转成本   → 本期开票结转（冲前期发出商品）
  · 当月发货未开票挂发出商品   → 本期新增发出商品（本页明细就是这些行）
  · 免费样品赠送计入销售费用   → 样品费用（不进发出商品）
不再使用「销售发票核销库存流水」，也不再依赖库存流水的销售出库行。

对外保持原来的函数名（build_sales_uninvoiced_rows / summarize_sales_uninvoiced /
build_goods_issued_rollforward / summarize_current_sales_cost），视图与导出不用改调用点。
"""
from __future__ import annotations

from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal

from settlement.models import (
    AccountingPeriod,
    MatchingMode,
    OpeningGoodsIssued,
    SalesLedgerLine,
)

ZERO = Decimal("0")
CENT = Decimal("0.01")

CARRY_CURRENT_INVOICED = "当月发货当月开票结转成本"
CARRY_PRIOR_INVOICED = "前期发货本期开票结转成本"
CARRY_UNINVOICED = "当月发货未开票挂发出商品"
CARRY_SAMPLE = "免费样品赠送计入销售费用"


def money(value) -> Decimal:
    return Decimal(value or ZERO).quantize(CENT, rounding=ROUND_HALF_UP)


def _carry(line) -> str:
    return (line.carry_type or "").strip()


def _ledger_lines(period):
    return list(period.sales_ledger_lines.select_related("material").all())


def _carry_totals(period):
    """台账按结转类型汇总：{结转类型: {"count": n, "quantity": q, "cost": c}}。"""
    totals = {}
    for line in _ledger_lines(period):
        key = _carry(line) or "(未填结转类型)"
        item = totals.setdefault(key, {"count": 0, "quantity": ZERO, "cost": ZERO})
        item["count"] += 1
        item["quantity"] += Decimal(line.quantity or ZERO)
        item["cost"] += Decimal(line.cost_amount or ZERO)
    return totals


@dataclass
class SalesUninvoicedRow:
    """清单明细行＝台账里「当月发货未开票挂发出商品」的一行（用户自己的台账行）。"""

    line: SalesLedgerLine
    out_month: object
    order_no: str
    customer_name: str
    channel: str
    company: str
    sales_category: str
    material_code: str
    material_name: str
    specification: str
    uom: str
    quantity: Decimal
    unit_cost: Decimal | None
    cost_amount: Decimal
    invoice_status: str
    prior_period: bool          # 出库期间早于本期间（以前月份出库仍未开票）
    outstanding_days: int


def build_sales_uninvoiced_rows(period: AccountingPeriod, include_closed: bool = False):
    """本期「已出库未开票」明细＝台账中结转类型为「当月发货未开票挂发出商品」的行。

    include_closed 保留形参（口径下不再有"已开票关闭"的行）：台账行就是台账行。
    """
    from django.utils import timezone

    today = timezone.localdate()
    rows = []
    for line in _ledger_lines(period):
        if _carry(line) != CARRY_UNINVOICED:
            continue
        out_month = line.out_month
        days = (today - out_month).days if out_month else 0
        rows.append(
            SalesUninvoicedRow(
                line=line,
                out_month=out_month,
                order_no=(line.order_no or "").strip(),
                customer_name=(line.customer_name or line.customer_code or "").strip(),
                channel=(line.channel or "").strip(),
                company=(line.company or "").strip(),
                sales_category=(line.sales_category or "").strip(),
                material_code=line.material_code or "",
                material_name=line.material_name or "",
                specification=line.specification or "",
                uom=line.uom or "",
                quantity=Decimal(line.quantity or ZERO),
                unit_cost=(
                    Decimal(line.cost_price) if line.cost_price is not None else None
                ),
                cost_amount=Decimal(line.cost_amount or ZERO),
                invoice_status=line.get_invoice_status_display(),
                prior_period=bool(out_month and out_month < period.month),
                outstanding_days=max(days, 0),
            )
        )
    rows.sort(key=lambda row: (row.out_month, row.order_no, row.material_code))
    return rows


def summarize_sales_uninvoiced(rows, period: AccountingPeriod):
    """页头六个数：未开票笔数 / 数量 / 期末（本月新增）发出商品 / 以前月份出库 / 未填成本 / 样品。"""
    totals = _carry_totals(period)
    sample = totals.get(CARRY_SAMPLE, {"count": 0, "quantity": ZERO, "cost": ZERO})
    return {
        "issue_count": len(rows),
        "outstanding_quantity": sum((row.quantity for row in rows), ZERO),
        "outstanding_cost": money(sum((row.cost_amount for row in rows), ZERO)),
        "prior_period_count": sum(1 for row in rows if row.prior_period),
        "unpriced_count": sum(1 for row in rows if row.unit_cost is None),
        "sample_count": sample["count"],
        "sample_cost": money(sample["cost"]),
    }


def filter_sales_uninvoiced_rows(rows, issue_month="", customer=""):
    """明细表筛选（页面与「导出Excel」共用同一套，保证两边行数一致）。

    · issue_month：出库期间，YYYY-MM 精确匹配；
    · customer：客户名称，模糊匹配（不分大小写）。
    老链接里的其它参数（销售订单号 / 产品 / 以前月份）不在这里处理，
    由导出分支自己叠加，保证既有链接行为不变。
    """
    month_value = str(issue_month or "").strip()
    if month_value:
        rows = [
            row
            for row in rows
            if row.out_month and row.out_month.strftime("%Y-%m") == month_value
        ]
    needle = str(customer or "").strip().lower()
    if needle:
        rows = [row for row in rows if needle in (row.customer_name or "").lower()]
    return rows


def summarize_uninvoiced_rows(rows):
    """筛选后的明细合计（明细表页脚用）：笔数 / 数量 / 金额 / 以前月份笔数 / 未填成本笔数。"""
    return {
        "issue_count": len(rows),
        "outstanding_quantity": sum((row.quantity for row in rows), ZERO),
        "outstanding_cost": money(sum((row.cost_amount for row in rows), ZERO)),
        "prior_period_count": sum(1 for row in rows if row.prior_period),
        "unpriced_count": sum(1 for row in rows if row.unit_cost is None),
    }


def uninvoiced_period_options(period: AccountingPeriod):
    """明细表「期间」下拉选项：本期明细里实际出现过的出库期间（新→旧）。"""
    seen = []
    for row in build_sales_uninvoiced_rows(period):
        if not row.out_month:
            continue
        value = row.out_month.strftime("%Y-%m")
        if value not in seen:
            seen.append(value)
    return [{"value": value, "label": value} for value in sorted(seen, reverse=True)]


def _opening_stats(period: AccountingPeriod):
    """期初发出商品（笔数, 数量, 金额）。

    优先用本期上传的「期初发出商品明细」；没有上传时金额按上期期末递归推算
    （与 `_opening_amount` 同口径）。此时只有台账里「以前月份出库」的行合计正好等于
    该推算金额时，才给出笔数/数量；否则留空（显示 —），避免出现"笔数数量与金额对不上"。
    """
    opening = list(
        OpeningGoodsIssued.objects.filter(period=period).exclude(
            match_mode=MatchingMode.IGNORED
        )
    )
    if opening:
        return (
            len(opening),
            sum((Decimal(item.opening_quantity or ZERO) for item in opening), ZERO),
            money(
                sum((Decimal(item.opening_amount or ZERO) for item in opening), ZERO)
            ),
        )
    amount = _opening_amount(period)
    prior = [row for row in build_sales_uninvoiced_rows(period) if row.prior_period]
    prior_amount = money(sum((row.cost_amount for row in prior), ZERO))
    if prior and prior_amount == amount:
        return (
            len(prior),
            sum((row.quantity for row in prior), ZERO),
            amount,
        )
    return None, None, amount


def _statistics_ending_amount(period: AccountingPeriod):
    """「发出商品及销售成本统计表」里本期所在月的「当月期末发出商品未开票成本」。

    该表按出库月/开票月逐月滚动，与本页（按台账结转类型）是两套取数；
    拿出来并列是为了让「两张表的期末发出商品为什么不一样」一眼可见。
    """
    from settlement.services.sales_voucher import build_sales_cost_statistics

    label = "%d年%d月" % (period.month.year, period.month.month)
    for item in build_sales_cost_statistics(period).get("months", []):
        if item.get("label") == label:
            return item["cells"].get("当月期末发出商品未开票成本")
    return None


def build_current_uninvoiced_rows(period: AccountingPeriod):
    """**本期新增**的未开票行（出库期间不早于本期）。

    用户 2026-09-27 方案A：首期之前的历史发货由「期初发出商品明细」承担（那是它的唯一来源），
    销售台账里同样带出来的历史未开票行不再计入「本期新增发出商品」——
    否则期初与台账历史行是同一批货，会被算两遍（7月多算 5,651,830.37）。
    明细表仍完整列出月末所有未开票行，历史行只是不重复参与滚动。
    """
    return [row for row in build_sales_uninvoiced_rows(period) if not row.prior_period]


def build_goods_issued_check_rows(period: AccountingPeriod):
    """「未开票笔数 → 免费样品 → 发出商品跨月滚动」勾稽表。

    每行＝{level, label, count, quantity, amount, note}；level 取
    group（分节标题）/ item（明细项）/ total（合计，加粗）。
    """
    totals = _carry_totals(period)

    def _pick(carry_type):
        item = totals.get(carry_type, {"count": 0, "quantity": ZERO, "cost": ZERO})
        return item["count"], item["quantity"], money(item["cost"])

    rows = build_sales_uninvoiced_rows(period)
    prior_rows = [row for row in rows if row.prior_period]
    current_rows = [row for row in rows if not row.prior_period]
    opening_count, opening_quantity, opening_amount = _opening_stats(period)
    uninvoiced_count = len(current_rows)
    uninvoiced_quantity = sum((row.quantity for row in current_rows), ZERO)
    uninvoiced_cost = money(sum((row.cost_amount for row in current_rows), ZERO))
    prior_count, prior_quantity, prior_cost = _pick(CARRY_PRIOR_INVOICED)
    sample_count, sample_quantity, sample_cost = _pick(CARRY_SAMPLE)

    rollforward_ending = money(opening_amount + uninvoiced_cost - prior_cost)
    statistics_ending = _statistics_ending_amount(period)
    gap = (
        money(rollforward_ending - statistics_ending)
        if statistics_ending is not None
        else None
    )

    check_rows = [
        {"level": "group", "label": "一、截至月末仍未开票", "count": None,
         "quantity": None, "amount": None, "note": ""},
        {"level": "item", "label": "未开票笔数（＝下方明细表行数）",
         "count": len(rows),
         "quantity": sum((row.quantity for row in rows), ZERO),
         "amount": money(sum((row.cost_amount for row in rows), ZERO)),
         "note": "台账「当月发货未开票挂发出商品」"},
        {"level": "item", "label": "其中：以前月份出库、本月仍未开票",
         "count": len(prior_rows),
         "quantity": sum((row.quantity for row in prior_rows), ZERO),
         "amount": money(sum((row.cost_amount for row in prior_rows), ZERO)),
         "note": "出库期间早于本期"},
        {"level": "group", "label": "二、免费样品（计入销售费用，不进发出商品）",
         "count": None, "quantity": None, "amount": None, "note": ""},
        {"level": "item", "label": "免费样品赠送计入销售费用",
         "count": sample_count, "quantity": sample_quantity, "amount": sample_cost,
         "note": "台账结转类型＝免费样品赠送计入销售费用"},
        {"level": "group", "label": "三、发出商品跨月滚动勾稽",
         "count": None, "quantity": None, "amount": None, "note": ""},
        {"level": "item", "label": "期初发出商品",
         "count": opening_count, "quantity": opening_quantity, "amount": opening_amount,
         "note": "上传的期初发出商品明细（首个核算月之前的历史发货）"},
        {"level": "item", "label": "加：本期新增发出商品",
         "count": uninvoiced_count, "quantity": uninvoiced_quantity,
         "amount": uninvoiced_cost,
         "note": "台账结转类型＝当月发货未开票挂发出商品，且出库期间在本期及以后"},
        {"level": "item", "label": "减：本期开票结转成本",
         "count": prior_count, "quantity": prior_quantity, "amount": prior_cost,
         "note": "台账结转类型＝前期发货本期开票结转成本"},
        {"level": "total", "label": "＝ 期末发出商品",
         # 笔数/数量不做加减（一张期初行可能对应多张开票行），只给滚出来的金额；
         # 月末仍未开票的笔数与数量看上面第一节
         "count": None, "quantity": None,
         "amount": rollforward_ending,
         "note": "期初＋本期新增－本期开票结转"},
    ]
    if statistics_ending is not None:
        check_rows.extend(
            [
                {"level": "group", "label": "四、与「发出商品及销售成本统计表」对账",
                 "count": None, "quantity": None, "amount": None, "note": ""},
                {"level": "item", "label": "本页期末发出商品（按台账结转类型滚动）",
                 "count": None, "quantity": None, "amount": rollforward_ending,
                 "note": "＝第三节合计"},
                {"level": "item", "label": "统计表期末发出商品（按出库月/开票月滚动）",
                 "count": None, "quantity": None, "amount": statistics_ending,
                 "note": "「发出商品及销售成本统计表」本期列"},
                {"level": "total", "label": "差额（本页 － 统计表）",
                 "count": None, "quantity": None, "amount": gap,
                 "note": "应为 0.00；不为 0 说明台账与统计表的取数范围有出入，需要查台账"},
            ]
        )
    return check_rows


def _opening_amount(period: AccountingPeriod) -> Decimal:
    """期初发出商品：优先用本期上传的「期初发出商品明细」，没有则用上一期间的期末（递归）。"""
    opening = list(
        OpeningGoodsIssued.objects.filter(period=period).exclude(match_mode="ignored")
    )
    if opening:
        return money(sum((Decimal(item.opening_amount or ZERO) for item in opening), ZERO))
    previous = (
        AccountingPeriod.objects.filter(month__lt=period.month)
        .order_by("-month")
        .first()
    )
    if previous is None:
        return ZERO
    previous_rows = build_sales_uninvoiced_rows(previous)
    previous_totals = _carry_totals(previous)
    return money(
        _opening_amount(previous)
        + sum((row.cost_amount for row in previous_rows), ZERO)
        - previous_totals.get(CARRY_PRIOR_INVOICED, {"cost": ZERO})["cost"]
    )


def build_goods_issued_rollforward(period: AccountingPeriod):
    """发出商品跨月滚动（台账口径）：

        期初发出商品 ＋ 本期台账「当月发货未开票挂发出商品」
        － 本期台账「前期发货本期开票结转成本」
        ＝ 期末发出商品

    差额按定义恒为 0（合不上说明台账缺行），页面照旧显示以便自查。
    """
    totals = _carry_totals(period)
    uninvoiced = totals.get(CARRY_UNINVOICED, {"count": 0, "quantity": ZERO, "cost": ZERO})
    prior = totals.get(CARRY_PRIOR_INVOICED, {"count": 0, "quantity": ZERO, "cost": ZERO})
    opening_amount = _opening_amount(period)
    rows = build_sales_uninvoiced_rows(period)
    # 本期新增＝只算本期及以后出库的未开票行（方案A：历史发货由期初承担，见
    # build_current_uninvoiced_rows 的说明）
    current_rows = [row for row in rows if not row.prior_period]
    current_issue_amount = money(
        sum((row.cost_amount for row in current_rows), ZERO)
    )
    invoice_cleared_amount = money(prior["cost"])
    ending_amount = money(opening_amount + current_issue_amount - invoice_cleared_amount)
    # 注意：局部变量不能叫 _opening_amount，否则会遮住上面的同名函数
    opening_count, _open_qty, _open_amt = _opening_stats(period)
    return {
        "opening_amount": opening_amount,
        "current_issue_amount": current_issue_amount,
        "invoice_cleared_amount": invoice_cleared_amount,
        "ending_amount": ending_amount,
        "difference": money(
            opening_amount + current_issue_amount - invoice_cleared_amount - ending_amount
        ),
        "opening_count": opening_count,
        "current_count": len(current_rows),
        "ending_count": len(rows),
        "prior_period_count": sum(1 for row in rows if row.prior_period),
        "unpriced_count": sum(1 for row in rows if row.unit_cost is None),
    }


def summarize_current_sales_cost(period: AccountingPeriod):
    """本期销售成本按台账「结转类型」拆分（勾稽用）：

        本期销售成本合计 = 已开票结转（当月+前期）＋ 未开票发出商品 ＋ 样品费用

    「未开票发出商品」按方案A只取**本期及以后出库**的行（历史发货属期初存量）。
    """
    totals = _carry_totals(period)
    current_invoiced = totals.get(CARRY_CURRENT_INVOICED, {"cost": ZERO})["cost"]
    prior_invoiced = totals.get(CARRY_PRIOR_INVOICED, {"cost": ZERO})["cost"]
    uninvoiced = sum(
        (row.cost_amount for row in build_current_uninvoiced_rows(period)), ZERO
    )
    sample = totals.get(CARRY_SAMPLE, {"cost": ZERO})["cost"]
    return {
        "sales_cost": money(current_invoiced + prior_invoiced + uninvoiced + sample),
        "invoiced_cost": money(current_invoiced + prior_invoiced),
        "goods_issued": money(uninvoiced),
        "sample_cost": money(sample),
        "ignored_cost": money(ZERO),
    }
