# -*- coding: utf-8 -*-
"""车间/部门编码与中文名显示过滤器。"""
from decimal import Decimal, InvalidOperation

from django import template

from settlement.services.imports import WORKSHOP_LABELS

register = template.Library()


@register.filter
def workshop_label(value):
    """车间编码 -> 中文名（如 RanZheng -> 染色车间）。"""
    return WORKSHOP_LABELS.get(value or "", value or "")


@register.filter
def thousands(value):
    """千分位（分节号）格式化：1234567.89 -> 1,234,567.89；保留原小数位。"""
    if value is None or value == "":
        return ""
    try:
        d = Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError):
        return value
    text = format(d, "f")
    if "." in text:
        int_part, dec_part = text.split(".")
    else:
        int_part, dec_part = text, ""
    neg = int_part.startswith("-")
    if neg:
        int_part = int_part[1:]
    grouped = []
    while int_part:
        grouped.insert(0, int_part[-3:])
        int_part = int_part[:-3]
    result = ",".join(grouped) if grouped else "0"
    if neg:
        result = "-" + result
    return f"{result}.{dec_part}" if dec_part else result
