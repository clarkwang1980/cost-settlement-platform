# -*- coding: utf-8 -*-
"""车间/部门编码与中文名显示过滤器。"""
from django.utils import timezone
from decimal import Decimal, InvalidOperation

from django import template
from django.template.defaultfilters import floatformat

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

@register.filter
def dash_num(value, digits=2):
    """报表金额/数量单元格：空或 0 → —；其余保留 digits 位小数并加千分位。

    用在「工单成本明细」：口径上"没有金额/没有数量"和"0.00"是一回事，
    整屏 0.00 会把真正的数埋掉。四舍五入沿用 Django floatformat（和站内其它表一致），
    所以 0.004 在 2 位小数下也是 0.00，同样显示 —。
    """
    if value is None or value == "":
        return "—"
    try:
        number = Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError):
        return value
    if number == 0:
        return "—"
    try:
        digits = int(digits)
    except (TypeError, ValueError):
        digits = 2
    text = floatformat(value, digits)
    if not text:
        return "—"
    try:
        if Decimal(str(text).replace(",", "")) == 0:
            return "—"
    except (InvalidOperation, ValueError):
        return text
    return thousands(text)


@register.filter
def percent(value, digits=2):
    """比率 -> 百分数（0.1234 -> 12.34）；空值原样返回。"""
    if value is None or value == "":
        return ""
    try:
        number = Decimal(str(value)) * 100
    except (InvalidOperation, ValueError, TypeError):
        return value
    try:
        digits = int(digits)
    except (TypeError, ValueError):
        digits = 2
    return format(number.quantize(Decimal("1." + "0" * digits)), "f")

@register.filter
def trim_num(value):
    """数字去掉多余的小数尾巴：11.000000 → 11、0.500000 → 0.5（与上传文件显示一致）。

    用于「生产工单」页：上传文件里计划数量写 11，不要显示成 11.000。
    """
    if value is None or value == "":
        return ""
    try:
        number = Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError):
        return value
    if number == number.to_integral_value():
        return str(number.quantize(Decimal("1")))
    trimmed = number.normalize()
    text = format(trimmed, "f")
    return text


@register.filter
def dash_blank(value):
    """录入框回填用：None / 空 / 0 → ""，其余原样（不带千分位）。

    用户 2026-09-23：构成表里「材料成本占比」是可选填的，没填过就留空，
    不要回填一个 0 把「按金额自动算」的含义弄丢。
    """
    if value is None or value == "":
        return ""
    try:
        number = Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError):
        return value
    if number == 0:
        return ""
    text = format(number.normalize(), "f")
    return text


@register.filter
def dash_pct(value, digits=2):
    """占比单元格：空或 0 → —；其余带上百分号（76.07 → 76.07%）。"""
    text = dash_num(value, digits)
    if text == "—":
        return text
    return "%s%%" % text


@register.filter
def group_num(value, digits=2):
    """录入框回填：带千分位（1,234.56）；空或 0 → ""（不要塞一个 0 进去）。"""
    if value is None or value == "":
        return ""
    try:
        number = Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError):
        return value
    if number == 0:
        return ""
    try:
        digits = int(digits)
    except (TypeError, ValueError):
        digits = 2
    return thousands(floatformat(number, digits))


@register.filter
def rate_factor(material_pct):
    """标准成本加成系数 ＝ 100 ÷ 材料成本占比（＝ 1 ＋ 人工加成比 ＋ 制造费用加成比）。"""
    try:
        material = Decimal(str(material_pct))
    except (InvalidOperation, ValueError, TypeError):
        return None
    if not material:
        return None
    return Decimal("100") / material


@register.filter
def yes_no(value):
    """布尔 → 是 / 否（与上传文件里的写法一致）。"""
    if value is None or value == "":
        return ""
    return "是" if value else "否"

@register.filter
def cell_text(value):
    """报表单元格的通用显示：空 → —；时间 → 本地时间；数字去掉多余小数尾巴；其余原样。

    用在「工单报工」明细表：那一版是按上传的「报工记录」原样出列，
    单元格类型不确定（有的是时间、有的是数字、有的是文本），不能逐列写死格式。
    """
    if value is None or value == "":
        return "—"
    if hasattr(value, "strftime"):
        return time_text(value)
    if isinstance(value, bool):
        return yes_no(value)
    if isinstance(value, (int, float, Decimal)):
        return trim_num(value)
    text_value = str(value).strip()
    return text_value or "—"


@register.filter
def num_class(value):
    """数字单元格右对齐（与全站 .num 的写法一致）；非数字返回空。"""
    if isinstance(value, bool):
        return ""
    if isinstance(value, (int, float, Decimal)):
        return "num"
    return ""


@register.filter
def time_text(value):
    """时间显示成上传文件里的写法：datetime → 2026-07-31 15:45:42；date → 2026-07-31。

    为什么不用 |date:"Y-m-d H:i:s"：浆粕车间的投料时间来自库存流水（DateField），
    Django 的 date 过滤器遇到纯日期 + 带时间的格式串会直接报错。
    """
    if value is None or value == "":
        return ""
    if hasattr(value, "strftime"):
        if hasattr(value, "hour"):
            # 库里存的是 UTC（USE_TZ=True）：先转成本地时间再格式化，
            # 否则会比上传文件里的时间早 8 小时
            if timezone.is_aware(value):
                value = timezone.localtime(value)
            return value.strftime("%Y-%m-%d %H:%M:%S")
        return value.strftime("%Y-%m-%d")
    return str(value)
