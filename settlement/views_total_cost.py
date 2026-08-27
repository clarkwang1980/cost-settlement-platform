# -*- coding: utf-8 -*-
"""总成本表页面与导出视图（独立模块，避免改动 views.py）。"""
from io import BytesIO
from urllib.parse import quote

from django.http import HttpResponse
from django.shortcuts import get_object_or_404, render
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from settlement.models import AccountingPeriod
from settlement.services.total_cost import (
    BLOCKS,
    SUBJECTS,
    build_total_cost_data,
)


def total_cost_view(request, period_id):
    period = get_object_or_404(AccountingPeriod, pk=period_id)
    data = build_total_cost_data(period)
    return render(request, "settlement/total_cost.html", {
        "data": data,
        "period": period,
    })


# ---------------- 导出 Excel ----------------

def _build_total_cost_workbook(data):
    """生成横向宽表：行=科目，列=四大块（各车间+合计）。"""
    period = data["period"]
    workshops = data["workshops"]
    n_ws = len(workshops)
    block_w = n_ws + 1
    ncols = 1 + len(BLOCKS) * block_w
    last_col = get_column_letter(ncols)

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    total_fill = PatternFill("solid", fgColor="FCE4D6")
    title_font = Font(name="微软雅黑", size=14, bold=True)
    normal_font = Font(name="微软雅黑", size=10)
    bold_font = Font(name="微软雅黑", size=10, bold=True)
    center = Alignment(horizontal="center", vertical="center")
    right = Alignment(horizontal="right", vertical="center")
    numfmt = "#,##0.00"

    wb = Workbook()
    ws = wb.active
    ws.title = "总成本表"

    # 标题
    ws.merge_cells(f"A1:{last_col}1")
    ws["A1"] = "总 成 本 表"
    ws["A1"].font = title_font
    ws["A1"].alignment = center
    ws.row_dimensions[1].height = 30
    ws.merge_cells(f"A2:{last_col}2")
    ws["A2"] = f"期间：{period.month:%Y年%m月}"
    ws["A2"].font = normal_font
    ws["A2"].alignment = center

    # 表头两行
    ws.merge_cells("A3:A4")
    c = ws["A3"]
    c.value = "成本项目"
    c.font = bold_font
    c.fill = header_fill
    c.border = border
    c.alignment = center

    col = 2
    for block in BLOCKS:
        l1 = get_column_letter(col)
        l2 = get_column_letter(col + block_w - 1)
        ws.merge_cells(f"{l1}3:{l2}3")
        c = ws[f"{l1}3"]
        c.value = block
        c.font = bold_font
        c.fill = header_fill
        c.border = border
        c.alignment = center
        for wi, ws_info in enumerate(workshops):
            cell = ws.cell(row=4, column=col + wi, value=ws_info["label"])
            cell.font = bold_font
            cell.fill = header_fill
            cell.border = border
            cell.alignment = center
        cell = ws.cell(row=4, column=col + n_ws, value="合计")
        cell.font = bold_font
        cell.fill = total_fill
        cell.border = border
        cell.alignment = center
        col += block_w
    ws.row_dimensions[3].height = 22
    ws.row_dimensions[4].height = 22

    def _write_row(r, values, fill=None, font=None):
        for i, v in enumerate(values):
            c2 = ws.cell(row=r, column=2 + i,
                         value="--" if v is None else float(v))
            c2.font = font or normal_font
            c2.border = border
            c2.alignment = right
            if v is not None:
                c2.number_format = numfmt
            if fill:
                c2.fill = fill

    # 数据行
    r = 5
    for row in data["rows"]:
        c = ws.cell(row=r, column=1, value=row["subject"])
        c.font = normal_font
        c.border = border
        c.alignment = Alignment(horizontal="left", vertical="center")
        _write_row(r, [cell["value"] for cell in row["cells"]])
        r += 1

    # 合计行
    c = ws.cell(row=r, column=1, value="合计")
    c.font = bold_font
    c.fill = total_fill
    c.border = border
    c.alignment = Alignment(horizontal="left", vertical="center")
    _write_row(r, [cell["value"] for cell in data["total_cells"]],
               fill=total_fill, font=bold_font)
    r += 2

    # 说明
    notes = [
        "口径说明：",
        "① 期初 = 期初在产品表按车间汇总；本期投入 = 工单结算结果 − 期初；产成品成本/在产品 = 结算结果完工/在产品拆分。",
        "② 产成品/在产品中水电/蒸汽/燃气/染料助剂/机物料/折旧按该车间制造费用结构比例拆分。",
        "③ 未结算车间本期投入/产成品/在产品无数据（--），仅有期初在产品，待结算后结转补全。",
        "④ 勾稽：期初 + 本期投入 = 产成品成本 + 在产品（对已结算车间逐科目校验，页面同步显示）。",
    ]
    for note in notes:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
        ws.cell(row=r, column=1, value=note).font = Font(
            name="微软雅黑", size=9, color="808080")
        ws.cell(row=r, column=1).alignment = Alignment(
            horizontal="left", vertical="center")
        r += 1

    # 列宽 / 冻结
    ws.column_dimensions["A"].width = 16
    for c2 in range(2, ncols + 1):
        ws.column_dimensions[get_column_letter(c2)].width = 10.5
    ws.freeze_panes = "B5"
    return wb


def total_cost_export(request, period_id):
    period = get_object_or_404(AccountingPeriod, pk=period_id)
    data = build_total_cost_data(period)
    wb = _build_total_cost_workbook(data)
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    filename = f"总成本表-{period.month:%Y-%m}.xlsx"
    response = HttpResponse(
        stream.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = f"attachment; filename*=UTF-8''{quote(filename)}"
    return response
