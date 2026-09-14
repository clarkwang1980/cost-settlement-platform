# -*- coding: utf-8 -*-
"""损益表页面与导出（独立模块，避免改动 views.py）。

页面结构：
- 按月费用录入：台账里出现过的每个出库月份一块，各录一套当月费用；
- 汇总：按【出库期间】或【开票期间】切换，层级 = 月份 → 销售渠道(国内/国外/未分类) → 销售大类；
- 明细：销售台账逐行；
- 导出 Excel：明细 + 两张汇总。
"""
from decimal import Decimal, InvalidOperation
from io import BytesIO
from urllib.parse import quote

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.http import HttpResponse
from django.shortcuts import get_object_or_404, redirect, render

from settlement.models import AccountingPeriod, ProfitStatementExpense
from settlement.services.exports import build_listing_workbook
from settlement.services.profit_statement import (
    ALLOCATION_FIELDS,
    BASIS_CHOICES,
    BASIS_OUT,
    build_category_summary,
    build_profit_statement,
    build_profit_statement_sheet,
    build_summary,
    ledger_months,
    month_label,
    summary_period_options,
)

PAGE_SIZE = 200
BASIS_VALUES = tuple(value for value, _label in BASIS_CHOICES)


def _parse_amount(raw, label):
    """解析页面填写的金额：允许千分位/空格、允许为空（按 0 处理）。"""
    text = str(raw or "").replace(",", "").replace("，", "").strip()
    if not text:
        return Decimal("0")
    try:
        return Decimal(text).quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError):
        raise ValueError(f"「{label}」不是有效金额：{raw}")


def _field_name(month, field):
    """表单字段名：exp_2026-07_admin_expense。"""
    return f"exp_{month_label(month)}_{field}"


def _month_forms(months):
    """把服务返回的按月信息整理成模板直接可渲染的字段列表。"""
    forms = []
    for item in months:
        forms.append(
            {
                **item,
                "fields": [
                    {
                        "name": _field_name(item["month"], field),
                        "label": label,
                        "value": item["values"].get(field, Decimal("0")),
                    }
                    for field, label in ALLOCATION_FIELDS
                ],
            }
        )
    return forms


def _normalise_range(period_options, period_from, period_to):
    """校验/纠正期间区间：非法值忽略；顺序反了自动对调；返回 (from, to)。"""
    values = [item["value"] for item in period_options]
    if period_from not in values:
        period_from = ""
    if period_to not in values:
        period_to = ""
    if period_from and period_to and values.index(period_from) > values.index(period_to):
        period_from, period_to = period_to, period_from
    return period_from, period_to


def _range_query(request):
    """「从/到」两个参数拼成跳转用的查询串（保留口径与区间）。"""
    pairs = []
    for key in ("basis", "pfrom", "pto"):
        value = request.GET.get(key)
        if value:
            pairs.append((key, value))
    return "&".join("%s=%s" % pair for pair in pairs)


@login_required
def profit_statement_view(request, period_id):
    period = get_object_or_404(AccountingPeriod, pk=period_id)
    months = ledger_months(period)

    basis = request.POST.get("basis") or request.GET.get("basis") or BASIS_OUT
    if basis not in BASIS_VALUES:
        basis = BASIS_OUT
    querystring = _range_query(request)
    back_url = f"{request.path}?{querystring}" if querystring else request.path

    if request.method == "POST":
        if period.status == AccountingPeriod.Status.LOCKED:
            messages.error(
                request,
                "该期间已锁账，不能修改损益表费用；如需调整请让管理员重开期间。",
            )
            return redirect(back_url)
        try:
            values_by_month = {}
            for month in months:
                values_by_month[month] = {
                    field: _parse_amount(request.POST.get(_field_name(month, field)), label)
                    for field, label in ALLOCATION_FIELDS
                }
        except ValueError as exc:
            messages.error(request, str(exc))
            return redirect(back_url)

        for month, values in values_by_month.items():
            ProfitStatementExpense.objects.update_or_create(
                period=period,
                month=month,
                defaults={**values, "updated_by": request.user},
            )
        summary = "；".join(
            "%s %s" % (
                month_label(month),
                "、".join(
                    f"{label} {values[field]:,.2f}" for field, label in ALLOCATION_FIELDS
                ),
            )
            for month, values in values_by_month.items()
        )
        messages.success(
            request,
            f"已保存 {len(values_by_month)} 个月的费用：{summary}。"
            "每个月按该月终端销售收入占比分摊到该月各台账行。",
        )
        return redirect(back_url)

    data = build_profit_statement(period)
    summary = build_summary(data["rows"], basis)

    # 顶端汇总表：只显示「从…到…」选定的期间
    period_options = summary_period_options(data["rows"], basis)
    period_from, period_to = _normalise_range(
        period_options, request.GET.get("pfrom"), request.GET.get("pto")
    )
    category_summary = build_category_summary(
        data["rows"], basis, period_from or None, period_to or None
    )

    page = Paginator(data["rows"], PAGE_SIZE).get_page(request.GET.get("page"))
    query = request.GET.copy()
    query.pop("page", None)
    page.querystring = query.urlencode()

    return render(
        request,
        "settlement/profit_statement.html",
        {
            "period": period,
            "data": data,
            "summary": summary,
            "category_summary": category_summary,
            "period_options": period_options,
            "period_from": period_from,
            "period_to": period_to,
            "months": _month_forms(data["months"]),
            "basis": basis,
            "basis_choices": BASIS_CHOICES,
            "basis_values": BASIS_VALUES,
            "page": page,
        },
    )


@login_required
def profit_statement_export(request, period_id):
    period = get_object_or_404(AccountingPeriod, pk=period_id)
    sheets, _data = build_profit_statement_sheet(period)
    stream = build_listing_workbook(sheets)
    filename = f"损益表-{period.month:%Y%m}.xlsx"
    response = HttpResponse(
        stream.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = f"attachment; filename*=UTF-8''{quote(filename)}"
    return response
