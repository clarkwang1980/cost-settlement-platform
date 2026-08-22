from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime, time
from decimal import Decimal
from io import BytesIO

from django.db.models import Sum
from django.http import HttpResponse
from django.utils import timezone
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from settlement.models import MatchAllocation
from settlement.services.matching import is_purchase_receipt, is_sales_issue
from settlement.services.purchase_accruals import (
    build_purchase_accrual_rows,
    build_purchase_difference_rows,
)
from settlement.services.sales_uninvoiced import (
    build_goods_issued_rollforward,
    build_sales_uninvoiced_rows,
)


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
SUBTOTAL_FILL = PatternFill("solid", fgColor="D9EAF7")
# 数字列统一显示：千分位分隔 + 两位小数（底层仍是数值，可继续计算）
NUMBER_FORMAT = "#,##0.00"


def _is_number(value) -> bool:
    return isinstance(value, (int, float, Decimal)) and not isinstance(
        value, bool
    )


def _clean_cell_value(value):
    """Excel 不支持带时区的日期时间：统一转成本地时间的 naive datetime。

    Django 开启 USE_TZ 后 DateTimeField 返回的 datetime 带 tzinfo，
    直接写入 openpyxl 会抛
    "Excel does not support timezones in datetimes"。
    这里在写入前做全局兜底，覆盖所有报表。
    """
    if isinstance(value, datetime):
        if value.tzinfo is not None:
            return timezone.localtime(value).replace(tzinfo=None)
        return value
    return value


def _clean_row(row):
    return [_clean_cell_value(value) for value in row]


def _append_table(workbook, title, headers, rows, highlight_rows=frozenset()):
    worksheet = workbook.create_sheet(title=title)
    worksheet.append(headers)
    for cell in worksheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row_index, row in enumerate(rows, start=2):
        worksheet.append(_clean_row(row))
        if row_index in highlight_rows:
            for cell in worksheet[row_index]:
                cell.fill = SUBTOTAL_FILL
                cell.font = Font(bold=True)
    # 数据区数字单元格设置千分位+两位小数格式
    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            if _is_number(cell.value):
                cell.number_format = NUMBER_FORMAT
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for index, column in enumerate(worksheet.columns, 1):
        width = max(
            len(str(cell.value or "")) for cell in list(column)[:200]
        )
        worksheet.column_dimensions[get_column_letter(index)].width = min(
            max(width + 2, 10), 42
        )
    return worksheet


def read_uploaded_rows(path):
    """读取用户上传的原始表格文件，返回 (表头, 数据行列表)。保留原始行数与列。"""
    suffix = path.suffix.lower()
    rows = []
    header = []
    if suffix in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook

        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        it = ws.iter_rows(values_only=True)
        try:
            header = list(next(it))
        except StopIteration:
            wb.close()
            return [], []
        for r in it:
            if any(v not in (None, "") for v in r):
                rows.append(tuple(r))
        wb.close()
    elif suffix == ".csv":
        import csv
        import io

        raw = path.read_bytes()
        decoded = None
        for enc in ("utf-8-sig", "gb18030", "utf-8"):
            try:
                decoded = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if decoded is None:
            return [], []
        csv_rows = list(csv.reader(io.StringIO(decoded)))
        if not csv_rows:
            return [], []
        header = csv_rows[0]
        rows = [
            tuple(r)
            for r in csv_rows[1:]
            if any(str(v).strip() for v in r)
        ]
    else:
        return [], []
    # 去掉尾部空列
    while header and header[-1] in (None, ""):
        header.pop()
        rows = [r[: len(header)] for r in rows]
    return header, rows


def build_settlement_workbook(period):
    workbook = Workbook()
    workbook.remove(workbook.active)

    _append_table(
        workbook,
        "库存计价",
        [
            "物料编码",
            "物料名称",
            "期初数量",
            "期初金额",
            "本期入库数量",
            "本期入库金额",
            "月末加权平均价",
            "本期出库数量",
            "本期出库金额",
            "期末数量",
            "期末金额",
            "是否收敛",
        ],
        (
            (
                item.material_code,
                item.material_name,
                item.opening_quantity,
                item.opening_amount,
                item.receipt_quantity,
                item.receipt_amount,
                item.weighted_average_price,
                item.issue_quantity,
                item.issue_amount,
                item.ending_quantity,
                item.ending_amount,
                "是" if item.converged else "否",
            )
            for item in period.inventory_valuations.all()
        ),
    )

    carry_rows = []
    for item in period.opening_inventory_balances.all():
        carry_rows.append(
            (
                "期初库存",
                item.source_period.month if item.source_period else None,
                item.material_code,
                item.material_name,
                item.quantity,
                item.amount,
                item.unit_price,
                (
                    "系统自动承接"
                    if item.is_system_generated
                    else "人工上传/覆盖"
                ),
            )
        )
    for item in period.opening_wip.all():
        carry_rows.append(
            (
                "期初在产品",
                item.source_period.month if item.source_period else None,
                item.order_no,
                "",
                "",
                item.total,
                "",
                "系统自动承接" if item.is_system_generated else "人工上传/覆盖",
            )
        )
    for item in period.purchase_contracts.filter(is_system_generated=True):
        carry_rows.append(
            (
                "合同余额",
                item.carried_from.period.month if item.carried_from else None,
                item.contract_no,
                item.material_name,
                item.quantity,
                "",
                item.unit_price_excl_tax,
                item.supplier_name,
            )
        )
    _append_table(
        workbook,
        "跨期结转",
        [
            "结转类型",
            "来源期间",
            "编码/单号",
            "名称",
            "数量",
            "金额",
            "单价",
            "说明",
        ],
        carry_rows,
    )

    _append_table(
        workbook,
        "期初采购暂估",
        [
            "原入库期间",
            "原入库日期",
            "原入库单号",
            "行号",
            "供应商",
            "合同号",
            "物料编码",
            "物料名称",
            "原入库数量",
            "以前到票数量",
            "期初未到票数量",
            "暂估单价",
            "期初暂估金额",
            "剩余未到票数量",
            "最近到票期间",
            "匹配状态",
        ],
        (
            (
                item.original_period,
                item.document_date,
                item.document_no,
                item.line_no,
                item.supplier_name or item.supplier_code,
                item.contract_no,
                item.material_code,
                item.material_name,
                item.original_quantity,
                item.prior_invoiced_quantity,
                item.opening_quantity,
                item.provisional_unit_price,
                item.opening_amount,
                item.remaining_quantity,
                item.invoice_period,
                item.get_match_status_display(),
            )
            for item in period.opening_purchase_accruals.all()
        ),
    )

    _append_table(
        workbook,
        "期初发出商品",
        [
            "原出库期间",
            "原出库日期",
            "原出库单号",
            "行号",
            "销售订单号",
            "客户",
            "产品编码",
            "产品名称",
            "原出库数量",
            "以前开票数量",
            "期初未开票数量",
            "原出库单位成本",
            "期初发出商品成本",
            "剩余未开票数量",
            "最近开票期间",
            "匹配状态",
        ],
        (
            (
                item.original_period,
                item.document_date,
                item.document_no,
                item.line_no,
                item.sales_order_no,
                item.customer_name or item.customer_code,
                item.material_code,
                item.material_name,
                item.original_quantity,
                item.prior_invoiced_quantity,
                item.opening_quantity,
                item.unit_cost,
                item.opening_amount,
                item.remaining_quantity,
                item.invoice_period,
                item.get_match_status_display(),
            )
            for item in period.opening_goods_issued.all()
        ),
    )

    _append_table(
        workbook,
        "期末存货",
        [
            "期末日期",
            "物料号",
            "物料名称",
            "分类名称",
            "仓库名称",
            "批次号",
            "结存数量",
            "单位名称",
            "单价",
            "金额",
        ],
        (
            (
                item.ending_date,
                item.material_code,
                item.material_name,
                item.category_name,
                item.warehouse,
                item.batch_no,
                item.quantity,
                item.uom,
                item.unit_price,
                item.amount,
            )
            for item in period.ending_inventory_balances.all()
        ),
    )

    _append_table(
        workbook,
        "车间结转",
        [
            "结转顺序",
            "车间/成本中心",
            "状态",
            "工单成本",
            "完工入库成本",
            "循环次数",
            "是否收敛",
            "结转时间",
            "结转人",
        ],
        (
            (
                item.sequence,
                item.workshop,
                item.get_status_display(),
                item.order_cost,
                item.production_receipt_amount,
                item.iterations,
                "是" if item.converged else "否",
                (
                    timezone.localtime(item.settled_at).replace(tzinfo=None)
                    if item.settled_at
                    else None
                ),
                str(item.settled_by or ""),
            )
            for item in period.workshop_settlements.all()
        ),
    )

    _append_table(
        workbook,
        "库存流水",
        [
            "创建时间",
            "物料名称",
            "物料编号",
            "分类",
            "物料规格",
            "仓库名称",
            "批号",
            "操作数量",
            "主单位",
            "辅助数量",
            "辅助单位",
            "备注",
            "业务单据编号",
            "业务类型",
            "供应商名称",
            "单价",
            "金额",
        ],
        (
            (
                item.txn_date,
                item.material_name,
                item.material_code,
                item.category_name,
                item.specification,
                item.warehouse,
                item.batch_no,
                item.quantity,
                item.uom,
                item.auxiliary_quantity,
                item.auxiliary_uom,
                item.remark,
                item.document_no,
                item.business_type,
                item.supplier_name or item.customer_name
                or item.supplier_code or item.customer_code or "",
                item.unit_price,
                item.amount,
            )
            for item in period.inventory_transactions.all()
        ),
    )

    purchase_rows = []
    for txn in period.inventory_transactions.all():
        if not is_purchase_receipt(txn):
            continue
        contract_allocs = txn.allocations.filter(
            kind=MatchAllocation.Kind.PURCHASE_CONTRACT
        ).select_related("contract_line")
        invoice_allocs = txn.allocations.filter(
            kind=MatchAllocation.Kind.PURCHASE_INVOICE,
            period__month__lte=period.month,
        ).select_related("purchase_invoice_line")
        purchase_rows.append(
            (
                txn.id,
                txn.txn_date,
                txn.document_no,
                txn.supplier_name,
                txn.material_code,
                txn.material_name,
                txn.quantity,
                "；".join(
                    f"{item.contract_line.contract_no}:{item.matched_quantity}"
                    for item in contract_allocs
                ),
                "；".join(
                    f"{item.purchase_invoice_line.invoice_no}:{item.matched_quantity}"
                    for item in invoice_allocs
                ),
                txn.unit_price,
                txn.amount,
                txn.price_source,
                txn.invoice_period,
            )
        )
    _append_table(
        workbook,
        "采购匹配",
        [
            "库存流水ID",
            "入库日期",
            "入库单号",
            "供应商",
            "物料编码",
            "物料名称",
            "入库数量",
            "匹配合同及数量",
            "匹配发票及数量",
            "暂估入库单价",
            "固定暂估入库金额",
            "价格来源",
            "最新发票期间",
        ],
        purchase_rows,
    )

    accrual_rows = build_purchase_accrual_rows(period)
    _append_table(
        workbook,
        "月末采购暂估",
        [
            "入库期间",
            "入库日期",
            "入库单号",
            "合同号",
            "供应商",
            "物料编码",
            "物料名称",
            "入库数量",
            "累计到票数量",
            "尚未到票数量",
            "暂估单价",
            "月末暂估金额",
            "最近到票期间",
            "暂估依据",
            "状态",
        ],
        (
            (
                row.receipt_month,
                row.document_date,
                row.document_no,
                row.contract_nos,
                row.supplier_name or row.supplier_code,
                row.material_code,
                row.material_name,
                row.receipt_quantity,
                row.invoiced_quantity,
                row.outstanding_quantity,
                row.provisional_unit_price,
                row.provisional_amount,
                row.latest_invoice_period,
                row.price_source,
                row.status,
            )
            for row in accrual_rows
        ),
    )

    difference_rows = build_purchase_difference_rows(period)
    _append_table(
        workbook,
        "采购暂估差异",
        [
            "发票期间",
            "发票号",
            "供应商",
            "物料编码",
            "物料名称",
            "入库期间",
            "入库日期",
            "入库单号",
            "核销数量",
            "暂估单价",
            "暂估金额",
            "发票单价",
            "发票金额",
            "暂估价差",
            "是否跨期",
            "状态",
        ],
        (
            (
                row.invoice.invoice_period,
                row.invoice.invoice_no,
                row.invoice.supplier_name or row.invoice.supplier_code,
                row.material_code,
                row.material_name,
                row.receipt_month,
                row.document_date,
                row.document_no,
                row.matched_quantity,
                row.provisional_unit_price,
                row.provisional_amount,
                row.invoice_unit_price,
                row.invoice_amount,
                row.difference_amount,
                "是" if row.cross_period else "否",
                row.status,
            )
            for row in difference_rows
        ),
    )

    sales_rows = []
    for txn in period.inventory_transactions.all():
        if not is_sales_issue(txn):
            continue
        allocations = txn.allocations.filter(
            kind=MatchAllocation.Kind.SALES_INVOICE
        ).select_related("sales_invoice_line")
        sales_rows.append(
            (
                txn.id,
                txn.txn_date,
                txn.document_no,
                txn.customer_name,
                txn.material_code,
                txn.material_name,
                txn.quantity,
                txn.amount,
                "；".join(
                    f"{item.sales_invoice_line.invoice_no}:{item.matched_quantity}"
                    for item in allocations
                ),
                txn.invoice_period,
                txn.get_match_status_display(),
            )
        )
    _append_table(
        workbook,
        "销售匹配",
        [
            "库存流水ID",
            "出库日期",
            "出库单号",
            "客户",
            "产品编码",
            "产品名称",
            "出库数量",
            "出库成本",
            "匹配发票及数量",
            "开票期间",
            "匹配状态",
        ],
        sales_rows,
    )

    uninvoiced_rows = build_sales_uninvoiced_rows(period)
    _append_table(
        workbook,
        "月末已出库未开票",
        [
            "出库期间",
            "出库日期",
            "出库单号",
            "销售订单号",
            "客户",
            "产品编码",
            "产品名称",
            "出库数量",
            "累计开票数量",
            "未开票数量",
            "出库成本",
            "已开票成本",
            "未开票成本",
            "匹配发票",
            "最近开票期间",
            "未开票天数",
            "状态",
        ],
        (
            (
                row.issue_month,
                row.document_date,
                row.document_no,
                row.sales_order_no,
                row.customer_name or row.customer_code,
                row.material_code,
                row.material_name,
                row.issue_quantity,
                row.invoiced_quantity,
                row.outstanding_quantity,
                row.issue_amount,
                row.invoiced_cost,
                row.outstanding_cost,
                row.invoice_nos,
                row.latest_invoice_period,
                row.outstanding_days,
                row.status,
            )
            for row in uninvoiced_rows
        ),
    )

    goods_issued = build_goods_issued_rollforward(period)
    _append_table(
        workbook,
        "发出商品滚动",
        [
            "核算期间",
            "期初发出商品",
            "本月销售出库",
            "本月开票结转成本",
            "期末发出商品",
            "差额",
            "期初未开票笔数",
            "期末未开票笔数",
            "其中以前月份出库",
        ],
        [
            (
                period.month,
                goods_issued["opening_amount"],
                goods_issued["current_issue_amount"],
                goods_issued["invoice_cleared_amount"],
                goods_issued["ending_amount"],
                goods_issued["difference"],
                goods_issued["opening_count"],
                goods_issued["ending_count"],
                goods_issued["prior_period_count"],
            )
        ],
    )

    # 工单成本（与页面一致：每个制造费用列后加分配率列，每车间末尾加合计行）
    WO_HEADERS = [
        "核算期间",
        "工单编号",
        "车间",
        "产品编码",
        "产品名称",
        "计划数量",
        "累计投料",
        "投料状态",
        "完成状态",
        "累计完工",
        "累计完工率",
        "是否用粗纱",
        "报工数量",
        "在产品数量",
        "材料成本",
        "水电",
        "水电分配率%",
        "蒸汽",
        "蒸汽分配率%",
        "天然气",
        "天然气分配率%",
        "折旧",
        "折旧分配率%",
        "机物料",
        "机物料分配率%",
        "染料助剂",
        "染料助剂分配率%",
        "工资",
        "工资分配率%",
        "工单总成本",
        "完工成本",
        "完工单位成本",
        "在产品成本",
    ]
    from settlement.services.workorder_costing import MANUFACTURING_FIELDS
    from settlement.models import BomLine

    wo_items = list(
        period.work_order_cost_results.select_related("work_order").order_by(
            "work_order__production_dept_code", "work_order__order_no"
        )
    )
    wo_rows = []
    wo_highlight = set()
    if wo_items:
        # 投料进度与粗纱标志
        progress_rows = list(
            period.issue_progress.values("work_order_no").annotate(
                total=Sum("actual_issue_quantity")
            )
        )
        progress_qty_by_order = {
            row["work_order_no"]: Decimal(row["total"] or 0) for row in progress_rows
        }
        progress_status_by_order = defaultdict(set)
        for status, wo in period.issue_progress.exclude(
            issue_status=""
        ).values_list("issue_status", "work_order_no"):
            progress_status_by_order[wo].add(status)
        roving_codes = set(
            BomLine.objects.filter(period=period, sub_item_code__startswith="BC-")
            .values_list("sub_item_code", flat=True)
        )
        roving_issue_orders = set(
            period.material_issues.filter(material_code__startswith="BC-")
            .exclude(work_order_no="")
            .values_list("work_order_no", flat=True)
        )
        roving_progress_orders = set(
            period.issue_progress.filter(material_code__startswith="BC-")
            .exclude(work_order_no="")
            .values_list("work_order_no", flat=True)
        )
        for item in wo_items:
            order_no = item.work_order.order_no
            ws = item.work_order.production_dept_code
            if order_no in progress_qty_by_order:
                item.accum_issue = progress_qty_by_order[order_no]
            else:
                item.accum_issue = Decimal(
                    period.material_issues.filter(work_order_no=order_no)
                    .aggregate(total=Sum("actual_issue_quantity"))["total"]
                    or 0
                )
            statuses = progress_status_by_order.get(order_no)
            item.issue_status = "、".join(sorted(statuses)) if statuses else "—"
            item.accum_finished = Decimal(item.finished_quantity or 0)
            wip_qty = Decimal(item.wip_quantity or 0)
            item.finish_rate_pct = (
                (item.accum_finished / (item.accum_finished + wip_qty) * Decimal("100"))
                if (item.accum_finished + wip_qty)
                else None
            )
            if item.accum_finished <= Decimal("0") and wip_qty <= Decimal("0"):
                item.completion_status = "未开工"
            elif wip_qty <= Decimal("0"):
                item.completion_status = "已完成"
            else:
                item.completion_status = "部分完成"
            if ws == "tefang":
                uses_roving = (
                    (order_no in roving_issue_orders)
                    or (order_no in roving_progress_orders)
                    or BomLine.objects.filter(
                        period=period,
                        work_order_no=order_no,
                        sub_item_code__in=roving_codes,
                    ).exists()
                )
                item.uses_roving = "是" if uses_roving else "否"
            else:
                item.uses_roving = "—"

        mat_total = {}
        fee_total = {field: {} for field in MANUFACTURING_FIELDS}
        equiv_total = {}
        for item in wo_items:
            ws = item.work_order.production_dept_code
            mat_total[ws] = mat_total.get(ws, Decimal("0")) + Decimal(
                item.direct_material or 0
            )
            for field in MANUFACTURING_FIELDS:
                fee_total[field][ws] = fee_total[field].get(ws, Decimal("0")) + Decimal(
                    getattr(item, field) or 0
                )
            equiv_total[ws] = equiv_total.get(ws, Decimal("0")) + Decimal(
                item.finished_quantity or 0
            ) + Decimal(item.wip_quantity or 0) * Decimal("0.5")

        def _rate(fee_field, ws):
            denominator = mat_total.get(ws, Decimal("0"))
            if not denominator:
                return Decimal("0")
            return (fee_total[fee_field].get(ws, Decimal("0")) / denominator) * Decimal("100")

        def _wage_rate(item, ws):
            denominator = equiv_total.get(ws, Decimal("0"))
            if not denominator:
                return Decimal("0")
            equivalent = Decimal(item.finished_quantity or 0) + Decimal(
                item.wip_quantity or 0
            ) * Decimal("0.5")
            return (equivalent / denominator) * Decimal("100")

        SUM_FIELDS = (
            "direct_material", "water_power", "steam", "natural_gas", "depreciation",
            "consumables", "dye_auxiliary", "wages", "total_cost",
            "finished_cost", "wip_cost",
        )
        subtotals = {}
        for item in wo_items:
            ws = item.work_order.production_dept_code
            st = subtotals.setdefault(
                ws,
                {"count": 0, **{f: Decimal("0") for f in SUM_FIELDS}},
            )
            st["count"] += 1
            for f in SUM_FIELDS:
                st[f] += Decimal(getattr(item, f) or 0)

        def _detail_row(item):
            ws = item.work_order.production_dept_code
            rate_text = (
                f"{item.finish_rate_pct:.2f}%" if item.finish_rate_pct is not None else "—"
            )
            return (
                f"{period.month:%Y-%m}",
                item.work_order.order_no,
                item.work_order.production_dept_code,
                item.product_code,
                item.product_name,
                item.plan_quantity,
                item.accum_issue,
                item.issue_status,
                item.completion_status,
                item.accum_finished,
                rate_text,
                item.uses_roving,
                item.finished_quantity,
                item.wip_quantity,
                item.direct_material,
                item.water_power,
                _rate("water_power", ws),
                item.steam,
                _rate("steam", ws),
                item.natural_gas,
                _rate("natural_gas", ws),
                item.depreciation,
                _rate("depreciation", ws),
                item.consumables,
                _rate("consumables", ws),
                item.dye_auxiliary,
                _rate("dye_auxiliary", ws),
                item.wages,
                _wage_rate(item, ws),
                item.total_cost,
                item.finished_cost,
                item.finished_unit_cost,
                item.wip_cost,
            )

        def _subtotal_row(ws, st):
            return (
                f"{ws} 合计（{st['count']}张工单）",
                "", "", "", "", "", "", "", "", "", "", "", "", "",
                st["direct_material"],
                st["water_power"],
                "",
                st["steam"],
                "",
                st["natural_gas"],
                "",
                st["depreciation"],
                "",
                st["consumables"],
                "",
                st["dye_auxiliary"],
                "",
                st["wages"],
                "",
                st["total_cost"],
                st["finished_cost"],
                "",
                st["wip_cost"],
            )

        current_ws = None
        for item in wo_items:
            ws = item.work_order.production_dept_code
            if current_ws is not None and ws != current_ws:
                wo_rows.append(_subtotal_row(current_ws, subtotals[current_ws]))
                wo_highlight.add(len(wo_rows))
            current_ws = ws
            wo_rows.append(_detail_row(item))
        if current_ws is not None:
            wo_rows.append(_subtotal_row(current_ws, subtotals[current_ws]))
            wo_highlight.add(len(wo_rows))

    _append_table(
        workbook,
        "工单成本",
        WO_HEADERS,
        wo_rows,
        highlight_rows=wo_highlight,
    )

    _append_table(
        workbook,
        "勾稽结果",
        [
            "检查编码",
            "检查项目",
            "左侧金额",
            "右侧金额",
            "差额",
            "状态",
            "分项明细",
        ],
        (
            (
                item.check_code,
                item.check_name,
                item.left_amount,
                item.right_amount,
                item.difference,
                item.get_status_display(),
                "；".join(
                    f"{key}={value}"
                    for key, value in item.details.items()
                ),
            )
            for item in period.reconciliations.all()
        ),
    )

    _append_table(
        workbook,
        "异常清单",
        ["严重程度", "类别", "引用", "异常内容", "处理状态"],
        (
            (
                item.get_severity_display(),
                item.category,
                item.reference,
                item.message,
                item.get_status_display(),
            )
            for item in period.exceptions.all()
        ),
    )

    stream = BytesIO()
    workbook.save(stream)
    stream.seek(0)
    return stream


def settlement_excel_response(period):
    stream = build_settlement_workbook(period)
    filename = f"成本结算结果-{period.month:%Y%m}.xlsx"
    response = HttpResponse(
        stream.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = (
        f"attachment; filename*=UTF-8''{__import__('urllib.parse').parse.quote(filename)}"
    )
    return response


def build_listing_workbook(sheets):
    """生成单文件多工作表 Excel。

    sheets: [(工作表名, 表头列表, 数据行(可迭代的元组/列表), 可选highlight_rows集合), ...]
    """
    workbook = Workbook()
    workbook.remove(workbook.active)
    for sheet in sheets:
        if len(sheet) >= 4:
            title, headers, rows, highlight_rows = sheet
        else:
            title, headers, rows = sheet
            highlight_rows = frozenset()
        _append_table(
            workbook, title, headers, rows, highlight_rows=highlight_rows
        )
    stream = BytesIO()
    workbook.save(stream)
    stream.seek(0)
    return stream

