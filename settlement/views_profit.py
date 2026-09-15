# -*- coding: utf-8 -*-
"""损益表页面与导出（独立模块，避免改动 views.py）。

页面结构：
- 按月费用录入：台账里出现过的每个出库月份一块，各录一套当月费用；
- 损益汇总表（一张表）：期间 × 销售渠道 × 销售大类，可按口径（出库/开票）、期间、销售渠道、销售大类筛选；
  打开页面默认只看当月（当前期间所属月份），可点「显示全部」或另勾期间；
- 明细：销售台账逐行；
- 导出 Excel：明细 + 两张汇总。
"""
from decimal import Decimal, InvalidOperation
from io import BytesIO
from urllib.parse import quote, urlencode

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
    build_profit_statement,
    build_profit_statement_sheet,
    build_summary,
    filter_summary_rows,
    ledger_months,
    month_label,
    UNINVOICED_LABEL,
    summary_category_options,
    summary_channel_options,
    summary_period_options,
)

PAGE_SIZE = 200
BASIS_VALUES = tuple(value for value, _label in BASIS_CHOICES)

#: 筛选表单 / 「显示全部」提交时带的标记：把「刚打开页面」和「主动筛选（哪怕一个都没勾）」区分开
FILTER_MARKER = "f"


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


def _normalise_selection(options, selected):
    """按选项顺序保留确实存在的筛选项（去重）；非法值忽略。"""
    wanted = {value for value in (selected or []) if value}
    return [item["value"] for item in options if item["value"] in wanted]


def _selection_query(basis, periods, channels, categories):
    """按「口径 + 期间 + 销售渠道 + 销售大类」拼查询串（口径切换链接 / 跳转回跳用）。"""
    pairs = [("basis", basis)]
    pairs += [("m", value) for value in periods]
    pairs += [("ch", value) for value in channels]
    pairs += [("cat", value) for value in categories]
    return urlencode([(key, value) for key, value in pairs if value])


def _filter_query(request):
    """当前筛选（口径 + 期间 + 销售渠道 + 销售大类）拼成跳转用的查询串。"""
    return _selection_query(
        request.GET.get("basis") or "",
        request.GET.getlist("m"),
        request.GET.getlist("ch"),
        request.GET.getlist("cat"),
    )


def _has_explicit_filter(request) -> bool:
    """本次请求是不是「用户主动动过的筛选」。

    带提交标记（筛选表单 / 显示全部）或显式给了期间 = 主动筛选，按 URL 原样显示；
    否则算「刚打开页面」，套用默认当月。
    """
    return FILTER_MARKER in request.GET or "m" in request.GET


def _default_period_selection(period_options, period) -> list:
    """打开页面时的默认期间：优先「当前期间所属月份」，台账里没有该月份时退回最近有数据的期间。

    只影响「期间」这一项；销售渠道 / 销售大类仍保持「全部」。
    """
    if not period_options:
        return []
    current = month_label(period.month)
    values = [item["value"] for item in period_options]
    if current in values:
        return [current]
    months = [value for value in values if value != UNINVOICED_LABEL]
    return [months[-1]] if months else [values[-1]]


@login_required
def profit_statement_view(request, period_id):
    period = get_object_or_404(AccountingPeriod, pk=period_id)
    months = ledger_months(period)

    basis = request.POST.get("basis") or request.GET.get("basis") or BASIS_OUT
    if basis not in BASIS_VALUES:
        basis = BASIS_OUT
    querystring = _filter_query(request)
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
        return redirect(back_url)

    data = build_profit_statement(period)

    # 汇总（一张表）：期间 × 销售渠道 × 销售大类；口径 / 期间 / 销售渠道 / 销售大类 四个筛选
    period_options = summary_period_options(data["rows"], basis)
    channel_options = summary_channel_options(data["rows"])
    category_options = summary_category_options(data["rows"])
    selected_periods = _normalise_selection(period_options, request.GET.getlist("m"))
    selected_channels = _normalise_selection(channel_options, request.GET.getlist("ch"))
    selected_categories = _normalise_selection(
        category_options, request.GET.getlist("cat")
    )
    # 刚打开页面（没带筛选标记、也没指定期间）默认只看当月；
    # 「期间」= 当前期间所属月份，台账里没有该月份时退回最近有数据的期间。
    default_period_applied = False
    if not _has_explicit_filter(request):
        selected_periods = _default_period_selection(period_options, period)
        default_period_applied = bool(selected_periods)
    has_filter = bool(selected_periods or selected_channels or selected_categories)
    summary = build_summary(
        filter_summary_rows(
            data["rows"], basis, selected_periods, selected_channels, selected_categories
        ),
        basis,
        has_filter=has_filter,
    )
    basis_links = [
        {
            "value": value,
            "label": label,
            "query": _selection_query(
                value, selected_periods, selected_channels, selected_categories
            ),
        }
        for value, label in BASIS_CHOICES
    ]

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
            "period_options": period_options,
            "channel_options": channel_options,
            "category_options": category_options,
            "selected_periods": selected_periods,
            "default_period_applied": default_period_applied,
            "selected_channels": selected_channels,
            "selected_categories": selected_categories,
            "basis_links": basis_links,
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
