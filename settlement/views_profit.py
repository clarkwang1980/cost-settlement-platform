# -*- coding: utf-8 -*-
"""损益表页面与导出（独立模块，避免改动 views.py）。"""
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
    build_profit_statement,
    build_profit_statement_sheet,
    get_expense,
)

PAGE_SIZE = 200


def _parse_amount(raw, label):
    """解析页面填写的金额：允许千分位/空格、允许为空（按 0 处理）。"""
    text = str(raw or "").replace(",", "").replace("，", "").strip()
    if not text:
        return Decimal("0")
    try:
        return Decimal(text).quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError):
        raise ValueError(f"「{label}」不是有效金额：{raw}")


@login_required
def profit_statement_view(request, period_id):
    period = get_object_or_404(AccountingPeriod, pk=period_id)

    if request.method == "POST":
        if period.status == AccountingPeriod.Status.LOCKED:
            messages.error(
                request,
                "该期间已锁账，不能修改损益表费用；如需调整请让管理员重开期间。",
            )
            return redirect("profit_statement", period_id=period.id)
        try:
            values = {
                field: _parse_amount(request.POST.get(field), label)
                for field, label in ALLOCATION_FIELDS
            }
        except ValueError as exc:
            messages.error(request, str(exc))
            return redirect("profit_statement", period_id=period.id)
        ProfitStatementExpense.objects.update_or_create(
            period=period,
            defaults={**values, "updated_by": request.user},
        )
        summary = "、".join(
            f"{label} {values[field]:,.2f}" for field, label in ALLOCATION_FIELDS
        )
        messages.success(
            request, f"本月费用已保存：{summary}。已按各行终端销售收入占比分摊。",
        )
        return redirect("profit_statement", period_id=period.id)

    expense = get_expense(period)
    data = build_profit_statement(period, expense=expense)

    page = Paginator(data["rows"], PAGE_SIZE).get_page(request.GET.get("page"))
    query = request.GET.copy()
    query.pop("page", None)
    page.querystring = query.urlencode()

    fields = [
        {"name": field, "label": label, "value": getattr(expense, field)}
        for field, label in ALLOCATION_FIELDS
    ]
    return render(
        request,
        "settlement/profit_statement.html",
        {
            "period": period,
            "data": data,
            "fields": fields,
            "expense": expense,
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
