# -*- coding: utf-8 -*-
"""车间/部门编码与中文名显示过滤器。"""
from django import template

from settlement.services.imports import WORKSHOP_LABELS

register = template.Library()


@register.filter
def workshop_label(value):
    """车间编码 -> 中文名（如 RanZhengXianWei -> 染色纤维车间）。"""
    return WORKSHOP_LABELS.get(value or "", value or "")
