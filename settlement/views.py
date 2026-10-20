import io
from pathlib import Path
from calendar import monthrange
from collections import defaultdict
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

from django.conf import settings
from urllib.parse import urlencode

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.db import transaction
from django.db.models import Count, Max, Q, Sum
from django.http import (
    FileResponse,
    Http404,
    HttpResponse,
    HttpResponseBadRequest,
    JsonResponse,
)
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
import logging

logger = logging.getLogger(__name__)

from django.utils import timezone
from django.views.decorators.cache import never_cache
from django.views.decorators.http import require_POST

from settlement.forms import (
    InventoryDocumentMatchForm,
    ManualMatchForm,
    MatchModeForm,
    OpeningDocumentMatchForm,
    PeriodForm,
    SourceDocumentMatchForm,
    UploadForm,
)
from settlement.models import (
    AccountingPeriod,
    contract_line_amount,
    contract_line_received_amount,
    ExceptionItem,
    HistoricalPrice,
    ImportBatch,
    InventoryTransaction,
    InventoryValuation,
    MatchAllocation,
    MatchingMode,
    OpeningGoodsIssued,
    OpeningPurchaseAccrual,
    PurchaseContractLine,
    PurchaseInvoiceLine,
    ReconciliationResult,
    SalesInvoiceLine,
    SalesLedgerLine,
    SettlementRun,
    StandardCostRate,
    WorkshopCostPool,
    CustomerOrder,
    CustomerReceipt,
    WorkOrder,)
from settlement.services.costing import (
    BOM_QTY_NORMALIZE_OVER,
    SETTLE_STALE_SECONDS,
    SettlePreconditionError,
    can_lock,
    ensure_workshop_settlements,
    reset_workshop_settlements,
    run_settlement,
    settle_workshop as settle_one_workshop,
    update_workshop_sequences,

    bom_material_prices,)
from settlement.services.exports import (
    build_listing_workbook,
    build_settlement_workbook,
    read_uploaded_rows,
    settlement_excel_response,
)
from settlement.services.imports import (
    ALIASES,
    WORKSHOP_LABELS,
    create_import_batch,
    normalize_header,
    process_batch,
)
from settlement.services.profit_statement import (
    build_profit_statement_sheet,
)
from settlement.services.total_cost import DEFAULT_WORKSHOPS
from settlement.services.ending_wip import (
    build_ending_wip_data,
    ending_wip_workshop_options,
    ending_wip_status_options,
)
from settlement.services.matching import (
    auto_match_all,
    is_purchase_receipt,
    is_sales_issue,
    manual_match,
    manual_match_opening,
    remove_manual_allocation,
    set_source_match_mode,
    set_transaction_match_mode,
)
from settlement.services.purchase_accruals import (
    build_purchase_accrual_rows,
    build_purchase_difference_rows,
    build_purchase_reversal_rows,
    contract_amount_by_month,
    invoice_attribution_month,
    provisional_basis,
    purchase_inbound_rows_of_period,
    purchase_receipts_of_period,
    receipt_attribution_month,
    summarize_purchase_differences,
    summarize_purchase_reversals,
)
from settlement.services.workorder_costing import (
    CASE_CHOICES,
    CASE_NORMAL,
    CASE_NO_PRICE,
    CASE_OTHER,
    CASE_SUGGESTIONS,
    CASE_COST_BASIS,
    FINISHED_RECEIPT_TYPES,
    all_roving_bom_orders,
    annotate_work_order_cost_rows,
    attach_finish_progress,
    unit_cost_diff_reason,
    build_issue_report_rows,
    build_product_resolver,
    build_work_order_cost_summary,
    finish_rate_text,
    summarize_issue_rows,
)
from settlement.services.rollforward import (
    carry_forward_context,
    invalidate_downstream_periods,
)
from settlement.services.purchase_voucher import (
    build_cost_difference_rows,
    build_purchase_voucher,
)
from settlement.services.receipt_voucher import build_finished_receipt_voucher
from settlement.services.material_issue_voucher import (
    PULP_WORKSHOP,
    WORKSHOP_ORDER_INDEX,
    build_material_issue_voucher,
    pulp_issue_rows,
)
from settlement.services.pulp_costing import _is_pulp_product, collect_pulp_costs
from settlement.services.voucher_center import build_voucher_center
from settlement.services.sales_voucher import (
    CARRY_COST_TYPES,
    block_param_names,
    build_sales_cost_statistics,
    build_sales_voucher_entries,
    build_sales_voucher_summary,
)
from settlement.services.sales_uninvoiced import (
    build_goods_issued_check_rows,
    build_goods_issued_rollforward,
    build_sales_uninvoiced_rows,
    filter_sales_uninvoiced_rows,
    summarize_uninvoiced_rows,
    uninvoiced_period_options,
)
from urllib.parse import quote
from io import BytesIO
import re
import zipfile


TEMPLATE_FILES = {
    ImportBatch.Kind.INVENTORY: "库存出入库流水.csv",
    ImportBatch.Kind.PURCHASE_CONTRACT: "采购合同.csv",
    ImportBatch.Kind.PURCHASE_INVOICE: "采购发票.csv",
    ImportBatch.Kind.SALES_INVOICE: "销售发票.csv",
    ImportBatch.Kind.SALES_LEDGER: "销售台账（出库+开票）.xlsx",
    ImportBatch.Kind.WORKSHOP_COST: "车间费用.csv",
    ImportBatch.Kind.OPENING_WIP: "期初在产品.xlsx",
    ImportBatch.Kind.OPENING_PURCHASE_ACCRUAL: "期初采购暂估明细.csv",
    ImportBatch.Kind.OPENING_GOODS_ISSUED: "期初发出商品明细（同销售台账）.xlsx",
    ImportBatch.Kind.OPENING_INVENTORY: "期初存货.csv",
    ImportBatch.Kind.ENDING_INVENTORY: "期末存货.csv",
    ImportBatch.Kind.WORK_ORDER: "生产工单.csv",
    ImportBatch.Kind.MATERIAL_ISSUE: "投料记录.csv",
    ImportBatch.Kind.WORK_REPORT: "报工记录.csv",
        "pulp_opening": "浆粕期初成品.xlsx",
    "pulp_flow": "浆粕成品流水.xlsx",
    "pulp_ending": "浆粕期末成品.xlsx",
    "pulp_material": "浆粕原料流水.xlsx",
    "pulp_ledger": "浆粕流水.xlsx",
    ImportBatch.Kind.BOM: "BOM表.xlsx",
    ImportBatch.Kind.ISSUE_PROGRESS: "投料进度表.xlsx",
    ImportBatch.Kind.HISTORICAL_PRICE: "历史单价（入库+出库）.xlsx",
    ImportBatch.Kind.WORK_ORDER_DETAIL: "工单明细.xlsx",
}


def _page(request, queryset, size=100):
    page = Paginator(queryset, size).get_page(request.GET.get("page"))
    query = request.GET.copy()
    query.pop("page", None)
    page.querystring = query.urlencode()
    return page


def _text_filters(request, queryset, fields):
    for parameter, field in fields.items():
        value = str(request.GET.get(parameter, "")).strip()
        if value:
            queryset = queryset.filter(**{f"{field}__icontains": value})
    return queryset


def _choice_filter(request, queryset, parameter, field):
    value = str(request.GET.get(parameter, "")).strip()
    if value:
        queryset = queryset.filter(**{field: value})
    return queryset


def _party_filter(request, queryset, parameter, code_field, name_field):
    value = str(request.GET.get(parameter, "")).strip()
    if value:
        queryset = queryset.filter(
            Q(**{f"{code_field}__icontains": value})
            | Q(**{f"{name_field}__icontains": value})
        )
    return queryset


def _date_filter(request, queryset, parameter, field):
    value = str(request.GET.get(parameter, "")).strip()
    if value:
        queryset = queryset.filter(**{field: value})
    return queryset


def _inventory_match_category(txn):
    if is_purchase_receipt(txn):
        return "采购入库"
    # 用户 2026-09-21：销售出库不再参与匹配——销售的开票状态/成本以上传的「销售台账」为准
    # （台账自带 开票状态/开票期间/销售成本/结转类型），平台不再拿销售发票去核销库存流水。
    business_type = (txn.business_type or "").strip()
    if business_type in FINISHED_RECEIPT_TYPES:
        return "完工入库"
    if business_type in {"投料", "投料回撤", "投料撤回"}:
        return "生产投料"
    return "无需匹配"


@login_required
def dashboard(request):
    if request.method == "POST":
        form = PeriodForm(request.POST)
        if form.is_valid():
            period = form.save(commit=False)
            period.month = period.month.replace(day=1)
            period.save()
            messages.success(request, f"{period}已建立。")
            return redirect("period_detail", period_id=period.id)
    else:
        form = PeriodForm()
    periods = AccountingPeriod.objects.all()
    return render(
        request,
        "settlement/dashboard.html",
        {
            "periods": periods,
            "form": form,
            "total_periods": periods.count(),
            "locked_periods": periods.filter(
                status=AccountingPeriod.Status.LOCKED
            ).count(),
        },
    )


@login_required
def period_detail(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    settlements = list(ensure_workshop_settlements(period))
    latest_run = period.runs.first()
    blocking = period.exceptions.filter(
        severity=ExceptionItem.Severity.BLOCKING,
        status=ExceptionItem.Status.OPEN,
    ).count()
    failed_checks = period.reconciliations.filter(
        status=ReconciliationResult.Status.FAIL
    ).count()
    counts = {
        "imports": period.import_batches.filter(
            status=ImportBatch.Status.SUCCESS
        ).count(),
        "inventory": period.inventory_transactions.count(),
        "contracts": period.purchase_contracts.count(),
        "purchase_invoices": period.purchase_invoices.count(),
        # 「张」口径（与采购发票页「上传发票（张 / 行）」一致）：不同发票号个数，
        # 多行发票只算一张；行数看 purchase_invoices。
        "purchase_invoice_count": period.purchase_invoices.aggregate(
            n=Count("invoice_no", distinct=True)
        )["n"],
        "sales_invoices": period.sales_invoices.count(),
        # 销售侧不再核销（用户 2026-09-21）：第 03 步的「待处理」改成台账里未开票的行数
        "sales_uninvoiced": period.sales_ledger_lines.filter(
            carry_type="当月发货未开票挂发出商品"
        ).count(),
        "orders": period.work_orders.count(),
        "valuations": period.inventory_valuations.count(),
        "ending_inventory": period.ending_inventory_balances.count(),
        "blocking": blocking,
        "failed_checks": failed_checks,
        "workshops": len(settlements),
        "settled_workshops": sum(
            1 for item in settlements if item.status == "settled"
        ),
        # 数据自检（用户 2026-09-30）：标准成本/损益表算不出来的常见原因都在这里
        "bom_lines": period.bom_lines.count(),
        "bom_priced": period.bom_lines.exclude(standard_price=None).count(),
        "historical_prices": period.historical_prices.count()
        + HistoricalPrice.objects.filter(
            period__month__lt=period.month
        ).exclude(
            material_code__in=period.historical_prices.values_list(
                "material_code", flat=True
            )
        ).values("material_code").distinct().count(),
        "standard_rates": period.standard_cost_rates.count(),
        "workshop_cost_pools": period.workshop_cost_pools.count(),
        "sales_ledger": period.sales_ledger_lines.count(),
        "material_issues": period.material_issues.count(),
        "work_reports": period.work_reports.count(),
        "work_order_costs": period.work_order_cost_results.count(),
        "batch_kinds": period.import_batches.values("kind").distinct().count(),
    }
    lockable, lock_reason = can_lock(period)
    return render(
        request,
        "settlement/period_detail.html",
        {
            "period": period,
            "latest_run": latest_run,
            "counts": counts,
            "lockable": lockable,
            "lock_reason": lock_reason,
            "batches": period.import_batches.all()[:20],
            "carry": carry_forward_context(period),
        },
    )


@login_required
@login_required
@require_POST
def delete_import_batch(request, period_id, batch_id):
    """删除一个上传批次及其导入的明细数据，并重置本月结算输出（需重新结转）。
    仅限未锁账期间操作。"""
    period = get_object_or_404(AccountingPeriod, id=period_id)
    if period.status == AccountingPeriod.Status.LOCKED:
        messages.error(request, "该期间已锁账，不能删除上传批次。")
        return redirect("import_batches", period_id=period.id)
    batch = get_object_or_404(ImportBatch, id=batch_id, period=period)
    label = f"{batch.get_kind_display()}（{batch.original_name}）"
    from settlement.services.imports import clear_calculation_outputs
    clear_calculation_outputs(period)
    batch.delete()  # CASCADE 删除该批次导入的全部明细
    messages.success(
        request,
        f"已删除上传批次：{label}。相关结算结果已重置，请按顺序重新结转。",
    )
    return redirect("import_batches", period_id=period.id)


def import_batches(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    batches = period.import_batches.select_related("uploaded_by")
    batches = _choice_filter(request, batches, "kind", "kind")
    batches = _choice_filter(request, batches, "status", "status")
    batches = batches.order_by("-created_at", "-id")
    return render(
        request,
        "settlement/import_batches.html",
        {
            "period": period,
            "page": _page(request, batches),
            "batch_kinds": ImportBatch.Kind.choices,
            "batch_statuses": ImportBatch.Status.choices,
        },
    )


@login_required
def download_import_file(request, period_id, batch_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    batch = get_object_or_404(ImportBatch, id=batch_id, period=period)
    disk = ""
    if batch.file:
        try:
            disk = batch.file.path
        except Exception:  # noqa: BLE001
            disk = ""
    if disk and Path(disk).is_file():
        return FileResponse(
            batch.file.open("rb"),
            as_attachment=True,
            filename=batch.original_name,
        )
    # 磁盘上没有（容器重建被清空）→ 回退到数据库里的副本
    if batch.file_content:
        response = HttpResponse(
        bytes(batch.file_content), content_type="application/octet-stream"
        )
        response["Content-Disposition"] = (
            "attachment; filename*=UTF-8''" + quote(batch.original_name or "原始文件")
        )
        return response
    raise Http404("原始文件已丢失（数据库里也没有副本）。")


#: 手工补建的工单号（每月都必须在当期工单主档里，否则车间结转会失败）
MANUAL_WORK_ORDER_CODES = ("31.295-20260320", "PULP202607", "R1")

#: 每月必传资料 + 作用说明（用户 2026-10-14：上传页要有"本月资料清单"）
MONTHLY_UPLOAD_CHECKLIST = [
    ("inventory", "必传", "本期全部出入库明细。只要数量，金额按月末加权平均价自动计算；整个结算的基础。"),
    ("sales_ledger", "必传", "本期销售出库与开票。销售成本结转、损益表、发出商品都靠它。"),
    ("work_order", "必传", "本期生产工单（含上月未完工）。车间结转与工单成本的依据。"),
    ("material_issue", "必传", "本期车间投料/领料。材料成本的依据。"),
    ("work_report", "必传", "本期完工报工。完工入库结转的数量依据。"),
    ("issue_progress", "必传", "投料进度表。投料与工单的对应关系。"),
    ("work_order_detail", "必传", "工单用料明细。BOM 反推投料与标准成本的依据。"),
    ("bom", "必传", "BOM 表。单位标准成本、BOM 标准价回填的依据。"),
    ("pulp_ledger", "必传", "浆粕车间流水（原料+成品通用）。不传，浆粕车间无法结转。"),
    ("purchase_invoice", "必传", "本期采购发票。核销采购入库与暂估，决定入库单价。"),
    ("purchase_contract", "必传", "本期采购合同。采购入库取价依据。"),
    ("pulp_ending", "可选", "浆粕期末成品。不传按平台计算值，会与上传值做核对。"),
    ("ending_inventory", "可选", "期末存货快照。只作核对参考（平台自己算期末）。"),
    ("historical_price", "新增才传", "历史单价。上传一次后各期自动沿用，只有新增物料才需补。"),
    ("sales_invoice", "不用传", "销售发票。开票信息已含在销售台账里。"),
]


def _monthly_checklist(period):
    """本月资料清单：必传项 + 是否已上传 + 导入结果。"""
    latest = {}
    for batch in period.import_batches.order_by("-id"):
        latest.setdefault(batch.kind, batch)
    kinds = dict(ImportBatch.Kind.choices)
    rows = []
    for kind, level, note in MONTHLY_UPLOAD_CHECKLIST:
        batch = latest.get(kind)
        if batch is None:
            if level == "必传":
                state, tone = "未上传", "bad"
            elif level == "可选":
                state, tone = "未传（可不传）", "muted"
            else:
                state, tone = "—", "muted"
            row_count = error_count = None
        elif batch.status == ImportBatch.Status.FAILED:
            state, tone = "导入失败", "bad"
            row_count, error_count = batch.row_count, batch.error_count
        elif (batch.error_count or 0) > 0:
            state = "部分成功（%d 行未导入）" % batch.error_count
            tone, row_count, error_count = "warn", batch.row_count, batch.error_count
        else:
            state, tone, row_count, error_count = "导入成功", "ok", batch.row_count, 0
        rows.append({
            "name": kinds.get(kind, kind), "level": level, "note": note,
            "state": state, "tone": tone, "row_count": row_count,
            "error_count": error_count, "uploaded_at": getattr(batch, "created_at", None),
        })
    # 固定项：3 个手工补建工单必须出现在当期工单主档（用户 2026-10-15）
    #   平台不会自动把它们带到下月，缺了车间结转会报"工单主档中不存在"。
    codes = list(MANUAL_WORK_ORDER_CODES)
    if codes:
        have = set(
            WorkOrder.objects.filter(period=period, order_no__in=codes)
            .values_list("order_no", flat=True)
        )
        missing = [c for c in codes if c not in have]
        rows.append({
            "name": "补充工单（手工补建）",
            "level": "必传",
            "note": "这 3 个工单是手工补建的（%s）。平台不会自动带到下个月，"
                    "缺了「车间结转」会报『工单主档中不存在』。用「补充工单.xlsx」重新上传即可。"
                    % "、".join(codes),
            "state": ("已上传 %d/%d" % (len(codes) - len(missing), len(codes)))
                     if not missing else ("缺失 %d 个：%s" % (len(missing), "、".join(missing))),
            "tone": "ok" if not missing else "bad",
            "row_count": len(codes) - len(missing),
            "error_count": len(missing),
            "uploaded_at": None,
        })
    return rows


def _render_upload_page(request, period, form):
    """渲染资料导入页（上传后也回到这里，不跳页）。"""
    checklist = _monthly_checklist(period)
    return render(
        request,
        "settlement/upload.html",
        {
            "period": period,
            "form": form,
            "checklist": checklist,
            "check_must": sum(1 for c in checklist if c["level"] == "必传"),
            "check_missing": sum(
                1 for c in checklist
                if c["level"] == "必传" and c["state"] == "未上传"
            ),
            "check_ok": sum(1 for c in checklist if c["tone"] == "ok"),
            "template_kinds": [
                k for k in ImportBatch.Kind.choices
                if not k[0].startswith("pulp_") and k[0] != "workshop_cost"
            ],
            "carry": carry_forward_context(period),
        },
    )


@login_required
def upload_file(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    if request.method == "POST":
        form = UploadForm(request.POST, request.FILES)
        if form.is_valid():
            try:
                batch = create_import_batch(
                    period=period,
                    kind=form.cleaned_data["kind"],
                    uploaded_file=form.cleaned_data["file"],
                    user=request.user,
                    replace_existing=form.cleaned_data["replace_existing"],
                )
                process_batch(batch)
                messages.success(
                    request,
                    f"{batch.get_kind_display()}导入成功：{batch.row_count}行，"
                    f"{batch.error_count}行未导入。",
                )
                return _render_upload_page(request, period, form)
            except Exception as exc:
                # 处理中途异常（如解析/整批校验失败）：以数据库真实状态为准，清掉回滚后残留的
                # PENDING 批次，避免占住文件哈希导致无法重传（process_batch 的保存已随事务回滚）
                pending = locals().get("batch")
                if pending is not None and pending.pk:
                    stale = ImportBatch.objects.filter(
                        pk=pending.pk, status=ImportBatch.Status.PENDING
                    ).first()
                    if stale is not None:
                        stale.delete()
                messages.error(request, str(exc))
    else:
        form = UploadForm()
    return _render_upload_page(request, period, form)


@login_required
def download_template(request, kind):
    # 历史单价：模板结构随年度变化（列=当年1~12月），按最新核算期间动态生成
    if kind == ImportBatch.Kind.HISTORICAL_PRICE:
        from settlement.services.price_table import build_price_template_workbook

        stream = build_price_template_workbook()
        filename = "历史单价（入库+出库）.xlsx"
        response = HttpResponse(
            stream.getvalue(),
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response["Content-Disposition"] = (
            f"attachment; filename*=UTF-8''{quote(filename)}"
        )
        return response
    filename = TEMPLATE_FILES.get(kind)
    if not filename:
        raise Http404("模板不存在")
    path = Path(settings.BASE_DIR) / "sample_templates" / filename
    if not path.is_file():
        raise Http404("模板文件不存在")
    return FileResponse(path.open("rb"), as_attachment=True, filename=filename)


@login_required
@require_POST
def run_period(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    try:
        run = run_settlement(period, request.user)
        messages.success(
            request,
            f"结算完成，共循环{run.iterations}次，"
            f"{'已经收敛' if run.converged else '尚未收敛'}。",
        )
    except Exception as exc:
        logger.exception("完整结算失败 period=%s", period.id)
        messages.error(
            request,
            "结算失败：" + str(exc) + "。如持续出现，请把本提示内容发给管理员。",
        )
    return redirect("period_detail", period_id=period.id)


@login_required
@require_POST
def rematch_period(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    try:
        reset_workshop_settlements(period)
        auto_match_all(period)
        messages.success(request, "自动匹配已重新执行。")
    except Exception as exc:
        messages.error(request, str(exc))
    return redirect("period_detail", period_id=period.id)


def _running_settle_context(period):
    """页面用：当前是否有结转在跑（含已运行秒数、车间名、是否已超时）。"""
    run = (
        SettlementRun.objects.filter(
            period=period,
            status=SettlementRun.Status.RUNNING,
        )
        .order_by("-started_at", "-id")
        .first()
    )
    if run is None or run.started_at is None:
        return None
    elapsed = int((timezone.now() - run.started_at).total_seconds())
    code = (run.summary or {}).get("workshop") or ""
    return {
        "elapsed": elapsed,
        "workshop": WORKSHOP_LABELS.get(code, code),
        "is_full": (run.summary or {}).get("mode") == "full",
        "stale": elapsed >= SETTLE_STALE_SECONDS,
    }


@login_required
def workshop_settlements_view(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    settlements = list(ensure_workshop_settlements(period))
    first_unsettled = next(
        (item for item in settlements if item.status != "settled"),
        None,
    )
    for item in settlements:
        item.can_settle_now = bool(
            period.status != AccountingPeriod.Status.LOCKED
            and first_unsettled
            and item.id == first_unsettled.id
        )
    return render(
        request,
        "settlement/workshop_settlements.html",
        {
            "period": period,
            "settlements": settlements,
            "has_settled": any(item.status == "settled" for item in settlements),
            "all_settled": bool(settlements)
            and all(item.status == "settled" for item in settlements),
            "carry": carry_forward_context(period),
            "running_settle": _running_settle_context(period),
        },
    )


@login_required
@require_POST
def update_workshop_order(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    sequence_by_id = {
        item.id: request.POST.get(f"sequence_{item.id}")
        for item in ensure_workshop_settlements(period)
    }
    try:
        update_workshop_sequences(period, sequence_by_id)
        messages.success(request, "车间结转顺序已更新。")
    except Exception as exc:
        messages.error(request, str(exc))
    return redirect("workshop_settlements", period_id=period.id)


@login_required
@require_POST
def settle_workshop_view(request, period_id, settlement_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    try:
        run = settle_one_workshop(
            period,
            settlement_id,
            request.user,
        )
        workshop_code = run.summary.get("workshop", "")
        workshop = WORKSHOP_LABELS.get(workshop_code, workshop_code or "该车间")
        messages.success(
            request,
            f"{workshop}结转完成，共循环{run.iterations}次。"
            + (
                "全部车间已结转，可以进行勾稽和锁账。"
                if run.summary.get("all_settled")
                else "现在可以结转下一车间。"
            ),
        )
    except SettlePreconditionError as exc:
        # 前置条件不满足（已结转 / 前序未完成）：给友好提示，不显示“结转失败”
        messages.info(request, str(exc))
    except Exception as exc:
        logger.exception("车间结转失败 period=%s settlement=%s", period.id, settlement_id)
        messages.error(
            request,
            "车间结转失败：" + str(exc) + "。如持续出现，请把本提示内容发给管理员。",
        )
    return redirect("workshop_settlements", period_id=period.id)


@login_required
@require_POST
def reset_workshops_view(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    try:
        reset_workshop_settlements(period)
        messages.warning(
            request,
            "车间结转结果已重置，可以重新排序并依次结转。",
        )
    except Exception as exc:
        messages.error(request, str(exc))
    return redirect("workshop_settlements", period_id=period.id)


def purchase_invoice_summary(queryset):
    """采购发票合计：上传张数/行数、数量、不含税金额，以及核销的已匹配/未匹配情况。

    queryset 传入的是已经套用页面筛选条件的发票明细行，所以合计随筛选条件变化。
    注意：必须用**不带** `matched_total` 注解的 queryset 调用，否则 Count/Sum 会被核销行放大。
    """
    totals = queryset.aggregate(
        invoice_count=Count("invoice_no", distinct=True),
        row_count=Count("id"),
        quantity=Sum("quantity"),
        amount=Sum("amount_excl_tax"),
        remaining_quantity=Sum("remaining_quantity"),
    )
    matched = MatchAllocation.objects.filter(
        kind=MatchAllocation.Kind.PURCHASE_INVOICE,
        purchase_invoice_line__in=queryset.values("pk"),
    ).aggregate(
        quantity=Sum("matched_quantity"),
        amount=Sum("amount"),
    )
    unmatched = queryset.filter(matchallocation__isnull=True).aggregate(
        count=Count("invoice_no", distinct=True),
        amount=Sum("amount_excl_tax"),
    )
    # 上传文件里「不含税金额」与「不含税单价 × 数量」不一致的行（常见于单价被四舍五入，
    # 或金额取自票面总额）：已核销金额按单价 × 核销数量计，所以两列会对不上，需要提示。
    gap_count = 0
    gap_amount = Decimal("0")
    for quantity, unit_price, amount in queryset.values_list(
        "quantity", "unit_price_excl_tax", "amount_excl_tax"
    ):
        diff = (amount or Decimal("0")) - (quantity * (unit_price or Decimal("0")))
        if abs(diff) > Decimal("0.01"):
            gap_count += 1
            gap_amount += diff
    totals["amount_gap_count"] = gap_count
    totals["amount_gap_amount"] = gap_amount.quantize(Decimal("0.01"))
    for key in ("quantity", "amount", "remaining_quantity"):
        totals[key] = totals.get(key) or Decimal("0")
    totals["invoice_count"] = totals["invoice_count"] or 0
    totals["row_count"] = totals["row_count"] or 0
    totals["matched_quantity"] = matched["quantity"] or Decimal("0")
    totals["matched_amount"] = matched["amount"] or Decimal("0")
    totals["unmatched_invoice_count"] = unmatched["count"] or 0
    totals["unmatched_amount"] = unmatched["amount"] or Decimal("0")
    # 金额统一到分（SQLite 下 Sum 会带出多余尾数）
    for key in ("amount", "matched_amount", "unmatched_amount"):
        totals[key] = totals[key].quantize(Decimal("0.01"))
    return totals


def _purchase_invoice_query_filters(request, queryset):
    """采购发票匹配页的筛选：供应商（编码或名称）、合同号。

    这两个条件同时作用于「各期间上传发票合计」表和下方明细表，保证两张表口径一致。
    期间不在这里处理——它是页面级别的「切换到哪个月」，见 purchase_matches。
    """
    queryset = _party_filter(
        request, queryset, "supplier", "supplier_code", "supplier_name"
    )
    return _text_filters(request, queryset, {"contract_no": "contract_no"})


def purchase_invoice_period_summary(request, period):
    """各期间上传发票合计：一行一个核算期间（按月份升序），另返回一行合计。

    每个期间都套用同一套筛选条件（供应商 / 合同号），所以各行的数字与进该期间页面看到的一致。
    """
    rows = []
    for item in AccountingPeriod.objects.order_by("month"):
        totals = purchase_invoice_summary(
            _purchase_invoice_query_filters(
                request, PurchaseInvoiceLine.objects.filter(period=item)
            )
        )
        totals["period"] = item
        totals["current"] = item.id == period.id
        rows.append(totals)
    total = {
        "invoice_count": sum(row["invoice_count"] for row in rows),
        "row_count": sum(row["row_count"] for row in rows),
        "quantity": sum((row["quantity"] for row in rows), Decimal("0")),
        "amount": sum((row["amount"] for row in rows), Decimal("0")),
        "matched_quantity": sum(
            (row["matched_quantity"] for row in rows), Decimal("0")
        ),
        "matched_amount": sum((row["matched_amount"] for row in rows), Decimal("0")),
        "remaining_quantity": sum(
            (row["remaining_quantity"] for row in rows), Decimal("0")
        ),
    }
    for key in ("amount", "matched_amount"):
        total[key] = total[key].quantize(Decimal("0.01"))
    return rows, total


@login_required
def purchase_matches(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    # 「期间」是页面级切换：选了别的期间就跳到那个期间的页面，并带上已填的供应商/合同号
    target_period_id = str(request.GET.get("period", "")).strip()
    if target_period_id and target_period_id != str(period.id):
        target = AccountingPeriod.objects.filter(id=target_period_id).first()
        if target is not None:
            kept = {
                key: str(request.GET.get(key, "")).strip()
                for key in ("supplier", "contract_no")
            }
            kept = {k: v for k, v in kept.items() if v}
            url = reverse("purchase_matches", args=[target.id])
            if kept:
                url = f"{url}?{urlencode(kept)}"
            return redirect(url)
    summary_rows, summary_total = purchase_invoice_period_summary(request, period)
    invoices = _purchase_invoice_query_filters(
        request, period.purchase_invoices.all()
    ).order_by("invoice_date", "invoice_no", "id")
    invoices = invoices.annotate(
        matched_total=Sum("matchallocation__matched_quantity")
    ).prefetch_related(
        "matchallocation_set__inventory_transaction__period",
        "matchallocation_set__opening_purchase_accrual",
    )
    return render(
        request,
        "settlement/purchase_matches.html",
        {
            "period": period,
            "page": _page(request, invoices),
            "summary_rows": summary_rows,
            "summary_total": summary_total,
            # 期间下拉用：全部核算期间
            "periods": [row["period"] for row in summary_rows],
        },
    )


def _contract_summary_rows(contracts):
    """采购合同汇总（按「收货期间」分组）→ (明细行列表, 合计行)。

    字段：收货期间 / 合同数量 / 签收数量 / 合同金额 / 签收金额。
      · 合同金额 = Σ 合同行金额（优先文件的「原币未税总价」，老数据退回 数量×单价）
      · 签收金额 = Σ（已签收数量 × 不含税单价）
    只取需要的列（values_list）在 Python 里聚合：行数不多，但避免把整表对象读进内存。
    """
    buckets = {}
    for receipt_period, qty, received_qty, unit_price, amount in contracts.order_by().values_list(
        "receipt_period",
        "quantity",
        "received_quantity",
        "unit_price_excl_tax",
        "amount_excl_tax",
    ):
        bucket = buckets.get(receipt_period)
        if bucket is None:
            bucket = {
                "receipt_period": receipt_period,
                "contract_quantity": Decimal("0"),
                "received_quantity": Decimal("0"),
                "contract_amount": Decimal("0"),
                "received_amount": Decimal("0"),
            }
            buckets[receipt_period] = bucket
        bucket["contract_quantity"] += Decimal(qty or 0)
        bucket["received_quantity"] += Decimal(received_qty or 0)
        bucket["contract_amount"] += contract_line_amount(qty, unit_price, amount)
        bucket["received_amount"] += contract_line_received_amount(received_qty, unit_price)
    rows = sorted(
        buckets.values(),
        key=lambda row: (
            row["receipt_period"] is None,               # 没填「收货期间」的排最后
            row["receipt_period"] or date.max,
        ),
    )
    total = {
        "receipt_period": None,
        "contract_quantity": sum((r["contract_quantity"] for r in rows), Decimal("0")),
        "received_quantity": sum((r["received_quantity"] for r in rows), Decimal("0")),
        "contract_amount": sum((r["contract_amount"] for r in rows), Decimal("0")),
        "received_amount": sum((r["received_amount"] for r in rows), Decimal("0")),
    }
    return rows, total


@login_required
def purchase_contracts(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    # 取数只做必要的事（用户 2026-09-19 要求「快一点」）：
    #   · 不再 prefetch 批次（模板没用 batch，白跑一条 SQL）；
    #   · 不再传 match_modes（筛选模块已从页面删除）；
    #   · 刻意**不用 .only()** —— 漏列字段会变成逐行延迟加载（实测 SQL 从 7 条涨到 109 条）。
    contracts = period.purchase_contracts.all()
    contracts = _text_filters(
        request,
        contracts,
        {
            "contract_no": "contract_no",
            "material": "material_code",
        },
    )
    contracts = _party_filter(
        request,
        contracts,
        "supplier",
        "supplier_code",
        "supplier_name",
    )
    contracts = _choice_filter(
        request, contracts, "match_mode", "match_mode"
    )
    contracts = contracts.order_by("contract_date", "contract_no", "id")
    summary_rows, summary_total = _contract_summary_rows(contracts)
    return render(
        request,
        "settlement/purchase_contracts.html",
        {
            "period": period,
            "page": _page(request, contracts),
            "summary_rows": summary_rows,
            "summary_total": summary_total,
        },
    )


def _is_virtual_document_no(value):
    """入库单号是否为「虚拟」编号（不该展示）。

    判定：空值，或纯数字的 ERP 系统内部编号（本机数据是 16 位，如 1785465500777398）。
    真实的业务单号一般含字母/短横线（如 GR-001、RK20260701）。
    """
    text = str(value or "").strip()
    if not text:
        return True
    return text.isdigit()


ACCRUAL_PAGE_FILTERS = ("receipt_month", "supplier", "material")


def period_month_end(month):
    """核算期间的月末日期（＝暂估截止日）：2026-07-01 -> 2026-07-31。"""
    return date(month.year, month.month, monthrange(month.year, month.month)[1])


def _accrual_page_filters(request, rows):
    """月末采购暂估**页面**的筛选：归属月份 / 供应商 / 物料编码。

    只认页面上看得见的这三个参数（ACCRUAL_PAGE_FILTERS）；入库单号、合同号、
    暂估状态等一律忽略，避免旧链接/旧书签带上参数后出现「看不见的筛选」。
    """
    receipt_month = str(request.GET.get("receipt_month", "")).strip()
    supplier = str(request.GET.get("supplier", "")).strip().lower()
    material = str(request.GET.get("material", "")).strip().lower()
    if receipt_month:
        rows = [
            r for r in rows if r.receipt_month.strftime("%Y-%m") == receipt_month
        ]
    if supplier:
        rows = [
            r for r in rows
            if supplier in (f"{r.supplier_code} {r.supplier_name}").lower()
        ]
    if material:
        rows = [
            r for r in rows
            if material in (f"{r.material_code} {r.material_name}").lower()
        ]
    return rows


def _accrual_filter_labels(request):
    """当前筛选条件的显示文字（明细表下方的提示行）。"""
    labels = []
    receipt_month = str(request.GET.get("receipt_month", "")).strip()
    if receipt_month:
        labels.append(f"归属月份 {receipt_month}")
    supplier = str(request.GET.get("supplier", "")).strip()
    if supplier:
        labels.append(f"供应商含「{supplier}」")
    material = str(request.GET.get("material", "")).strip()
    if material:
        labels.append(f"物料编码/名称含「{material}」")
    return labels


def _accrual_receipt_month_options(request, rows):
    """归属月份下拉：全部 + 明细里出现过的月份（新→旧）。

    当前参数值即使本月没有未到票明细，也保留为选项，避免出现「看不见的筛选」。
    """
    values = sorted(
        {r.receipt_month.strftime("%Y-%m") for r in rows if r.receipt_month},
        reverse=True,
    )
    current = str(request.GET.get("receipt_month", "")).strip()
    if current and current not in values:
        values.insert(0, current)
    return [{"value": value, "label": value} for value in values]


def _accrual_export_query(request):
    """「导出Excel」链接的查询串：只带页面上看得见的三个筛选条件，保证“筛选后导出＝所见”。"""
    kept = {
        key: str(request.GET.get(key, "")).strip() for key in ACCRUAL_PAGE_FILTERS
    }
    return urlencode({key: value for key, value in kept.items() if value})


def _filter_accrual_rows(request, rows):
    """月末采购暂估的筛选（归属月份/入库单号/供应商/物料/合同号/暂估状态）。

    供导出 Excel 使用：导出链接只会带上页面上看得见的三个条件（见 _accrual_export_query），
    这里多支持的单号/合同号/状态是给手动拼链接或旧链接用的，页面上不体现。
    """
    receipt_month = str(request.GET.get("receipt_month", "")).strip()
    document_no = str(request.GET.get("document_no", "")).strip().lower()
    supplier = str(request.GET.get("supplier", "")).strip().lower()
    material = str(request.GET.get("material", "")).strip().lower()
    contract_no = str(request.GET.get("contract_no", "")).strip().lower()
    status = str(request.GET.get("status", "")).strip()
    if receipt_month:
        rows = [r for r in rows if r.receipt_month.strftime("%Y-%m") == receipt_month]
    if document_no:
        rows = [r for r in rows if document_no in (r.document_no or "").lower()]
    if supplier:
        rows = [
            r for r in rows
            if supplier in (f"{r.supplier_code} {r.supplier_name}").lower()
        ]
    if material:
        rows = [
            r for r in rows
            if material in (f"{r.material_code} {r.material_name}").lower()
        ]
    if contract_no:
        rows = [r for r in rows if contract_no in r.contract_nos.lower()]
    if status:
        rows = [r for r in rows if r.status == status]
    return rows


def _purchase_flow_stats(period, accrual_rows=None):
    """当月发生额四件套：当月到货到票 / 本月暂估 / 本月到票 / 本月回冲金额。

    **严格按本期取数**（与「回冲汇总」页原有的"截至本月末累计"口径不同）：
    · 归属月份（用户 2026-09-19）：发票填了「收货期间」→ 以发票为准；否则按该入库的暂估归属
      （＝合同有到货签收时按合同「收货期间」，否则入库流水月份）；核销期初台账的沿用台账原月份。
    · 本月到票 ＝ `MatchAllocation.period＝本期` 的核销金额合计；
      其中**归属本月**的入库 → 当月到货到票（直接入账，不进暂估）；
      其余（归属以前月份 / 期初暂估台账）→ 本月回冲金额（冲回以前暂估）。
    · 本月暂估 ＝ **归属本月**的入库、截至本期末仍未到票的暂估金额（各期不重复）。
    · 恒等式：本月到票 ＝ 当月到货到票 ＋ 本月回冲金额。

    accrual_rows 传入已算好的 `build_purchase_accrual_rows(period)` 可省一次重算。
    """
    same_month = Decimal("0")
    cross = Decimal("0")
    same_month_count = 0
    month_alloc_count = 0
    allocations = MatchAllocation.objects.filter(
        kind=MatchAllocation.Kind.PURCHASE_INVOICE,
        period=period,
    ).select_related("inventory_transaction__period", "purchase_invoice_line")
    for allocation in allocations:
        month_alloc_count += 1
        amount = (Decimal(allocation.amount or 0)).quantize(Decimal("0.01"))
        txn = allocation.inventory_transaction
        # 归属月份（用户 2026-09-19）：发票「收货期间」优先 → 其次该入库的暂估归属
        # （＝合同收货期间优先，否则流水月份）→ 期初台账沿用原规则。
        attribution = invoice_attribution_month(allocation.purchase_invoice_line, txn)
        if attribution is not None:
            same = attribution == period.month
        else:
            same = txn is not None and txn.period.month == period.month
        if same:
            same_month += amount
            same_month_count += 1
        else:
            cross += amount
    rows = (
        accrual_rows if accrual_rows is not None
        else build_purchase_accrual_rows(period)
    )
    month_rows = [row for row in rows if row.receipt_month == period.month]
    month_accrual = sum(
        (row.provisional_amount or Decimal("0") for row in month_rows), Decimal("0")
    ).quantize(Decimal("0.01"))
    # 期末暂估：截至本期末仍未到票的余额（＝「各期间月末采购暂估」里本期那一行的金额）
    ending_accrual = sum(
        (row.provisional_amount or Decimal("0") for row in rows), Decimal("0")
    ).quantize(Decimal("0.01"))
    return {
        "same_month": same_month.quantize(Decimal("0.01")),
        "same_month_count": same_month_count,
        "month_accrual": month_accrual,
        "month_accrual_count": len(month_rows),
        "invoice_matched": (same_month + cross).quantize(Decimal("0.01")),
        "invoice_matched_count": month_alloc_count,
        "cross": cross.quantize(Decimal("0.01")),
        "cross_count": month_alloc_count - same_month_count,
        "ending_accrual": ending_accrual,
    }


def _accrual_balance_as_of(period):
    """截至该期末仍未到票的暂估余额（＝build_purchase_accrual_rows 各行的暂估金额之和）。"""
    return sum(
        (row.provisional_amount or Decimal("0")
         for row in build_purchase_accrual_rows(period)),
        Decimal("0"),
    ).quantize(Decimal("0.01"))


def _accrual_opening_balance(period):
    """**首个核算期间**的期初暂估：期初采购暂估台账里「本期之前尚未到票」的部分。

    非首个期间不用这个函数 —— 那些期间的期初＝上一期间的期末（在表里逐行传递）。
    """
    total = Decimal("0")
    openings = (
        OpeningPurchaseAccrual.objects.filter(period__month__lte=period.month)
        .exclude(match_mode=MatchingMode.IGNORED)
    )
    for item in openings:
        quantity = Decimal(item.opening_quantity or 0)
        if quantity <= 0:
            continue
        matched_before = item.allocations.filter(
            kind=MatchAllocation.Kind.PURCHASE_INVOICE,
            period__month__lt=period.month,
        ).exclude(
            purchase_invoice_line__match_mode=MatchingMode.IGNORED
        ).aggregate(quantity=Sum("matched_quantity"))["quantity"] or Decimal("0")
        outstanding = max(quantity - Decimal(matched_before), Decimal("0"))
        total += Decimal(item.opening_amount or 0) * outstanding / quantity
    return total.quantize(Decimal("0.01"))


def _accrual_opening_detail_rows(period):
    """首个期间的期初暂估明细：期初采购暂估台账里「本期之前尚未到票」的行。

    （非首个期间不用这里 —— 那些期间的期初＝上一期间的期末，直接拿上一期的未到票明细。）
    字段与 build_purchase_accrual_rows 对齐，模板可共用一张表。
    """
    rows = []
    openings = (
        OpeningPurchaseAccrual.objects.filter(period__month__lte=period.month)
        .exclude(match_mode=MatchingMode.IGNORED)
        .order_by("original_period", "document_date", "document_no", "line_no", "id")
    )
    for item in openings:
        quantity = Decimal(item.opening_quantity or 0)
        if quantity <= 0:
            continue
        matched = item.allocations.filter(
            kind=MatchAllocation.Kind.PURCHASE_INVOICE,
            period__month__lt=period.month,
        ).exclude(
            purchase_invoice_line__match_mode=MatchingMode.IGNORED
        ).aggregate(
            quantity=Sum("matched_quantity")
        )["quantity"] or Decimal("0")
        matched = min(Decimal(matched), quantity)
        outstanding = quantity - matched
        if outstanding <= 0:
            continue
        if item.provisional_unit_price is None:
            status = "缺少暂估价格"
        elif matched > 0:
            status = "部分到票"
        else:
            status = "全部暂估"
        rows.append(
            {
                "receipt_month": item.original_period,
                "document_date": item.document_date,
                "document_no": item.document_no,
                "transaction": None,
                "supplier": item.supplier_name or item.supplier_code or "",
                "material_code": item.material_code,
                "material_name": item.material_name,
                "receipt_quantity": quantity,
                "invoiced_quantity": matched,
                "outstanding_quantity": outstanding,
                "unit_price": item.provisional_unit_price,
                "amount": (
                    Decimal(item.opening_amount or 0) * outstanding / quantity
                ).quantize(Decimal("0.01")),
                "status": status,
            }
        )
    return rows


def _accrual_rows_as_dicts(rows):
    """本月暂估明细：把暂估行整理成模板好用的字典。"""
    return [
        {
            "receipt_month": row.receipt_month,
            "document_date": row.document_date,
            "document_no": row.document_no,
            "transaction": row.transaction,
            "supplier": row.supplier_name or row.supplier_code or "",
            "material_code": row.material_code,
            "material_name": row.material_name,
            "receipt_quantity": row.receipt_quantity,
            "invoiced_quantity": row.invoiced_quantity,
            "outstanding_quantity": row.outstanding_quantity,
            "unit_price": row.provisional_unit_price,
            "amount": row.provisional_amount,
            "status": row.status,
        }
        for row in rows
    ]


def purchase_accrual_detail_tables(period, all_rows):
    """「到票与暂估」上表五个数字的明细表（金额与上表逐列对得上）。

    · 期初暂估明细 ＝ 上一期间期末的未到票明细（首个期间取期初采购暂估台账的余额明细）
    · 本月到票明细 ＝ 本期发票核销逐笔（＝ 当月到货到票 ＋ 本月回冲）
    · 当月到货到票明细 ＝ 本月到票里核销「归属本月」入库的行
    · 本月回冲明细 ＝ 本月到票里核销「归属以前月份入库／期初台账」的行
    · 本月暂估明细 ＝ 归属本月、截至本月末仍未到票的行（＝ 期末暂估里的本月归属部分）
    """
    periods = list(AccountingPeriod.objects.order_by("month"))
    index = [item.id for item in periods].index(period.id)
    previous = periods[index - 1] if index > 0 else None
    opening_rows = (
        _accrual_rows_as_dicts(build_purchase_accrual_rows(previous))
        if previous is not None
        else _accrual_opening_detail_rows(period)
    )
    invoices = build_purchase_difference_rows(period)
    month_accrual_rows = [row for row in all_rows if row.receipt_month == period.month]
    return {
        "opening": opening_rows,
        "opening_from_previous": previous,
        "invoice": invoices,
        "same_month": [row for row in invoices if not row.cross_period],
        "cross": [row for row in invoices if row.cross_period],
        "month_accrual": _accrual_rows_as_dicts(month_accrual_rows),
    }


#: 「到票与暂估」4 张明细小表各自的筛选参数（用户 2026-09-19）。
#: 每张表一套独立前缀：改一张表的筛选不影响其它表，也不与页面主筛选
#: （receipt_month / supplier / material）同名。
ACCRUAL_DETAIL_FILTER_KEYS = {
    "invoice": {"period": "inv_period", "supplier": "inv_supplier", "material": "inv_material"},
    "same_month": {"period": "sm_period", "supplier": "sm_supplier", "material": "sm_material"},
    "month_accrual": {"month": "ma_month", "supplier": "ma_supplier"},
    "cross": {"period": "cr_period", "supplier": "cr_supplier"},
}

#: 每张明细表「期间」列的名称（下拉框标题）与「全部」选项文字
ACCRUAL_DETAIL_PERIOD_LABELS = {
    "invoice": ("开票期间", "全部期间"),
    "same_month": ("开票期间", "全部期间"),
    "month_accrual": ("归属月份", "全部月份"),
    "cross": ("开票期间", "全部期间"),
}

#: 每张明细表筛选栏右侧显示的金额徽标名称（与上表列名一致，用户 2026-09-19）
ACCRUAL_DETAIL_AMOUNT_LABELS = {
    "invoice": "本月到票",
    "same_month": "当月到货到票",
    "month_accrual": "本月暂估",
    "cross": "本月回冲",
}

#: 每张明细表对应的导出种类（「导出Excel」链接用）
ACCRUAL_DETAIL_EXPORT_KIND_NAMES = {
    "invoice": "purchase_accrual_invoice",
    "same_month": "purchase_accrual_same_month",
    "month_accrual": "purchase_accrual_month",
    "cross": "purchase_accrual_cross",
}

#: 供应商/物料编码建议值（datalist）最多给多少个，避免明细行多时把页面撑大
ACCRUAL_DETAIL_SUGGESTION_LIMIT = 300


def _detail_period_value(key, row):
    """这一行属于哪个期间：本月暂估看「归属月份」，其余三张表看发票的「开票期间」。"""
    if key == "month_accrual":
        return row.get("receipt_month")
    return row.invoice.invoice_period


def _detail_supplier_text(key, row):
    """这一行的供应商显示文字（本月暂估的行是字典，其余三张表挂在发票上）。"""
    if key == "month_accrual":
        return (row.get("supplier") or "").strip()
    invoice = row.invoice
    return (invoice.supplier_name or invoice.supplier_code or "").strip()


def _detail_material_text(key, row):
    if key == "month_accrual":
        code, name = row.get("material_code"), row.get("material_name")
    else:
        code, name = row.material_code, row.material_name
    return f"{code or ''} {name or ''}".strip()


def _detail_suggestion_values(values):
    """筛选框的建议值：去重、排序、限量（给 datalist 用，仍可自由输入）。"""
    return sorted({value for value in values if value})[:ACCRUAL_DETAIL_SUGGESTION_LIMIT]


def _detail_filter_amount(key, row):
    """每行金额：本月暂估是暂估金额，其余三张表是发票金额（与标题徽标同口径）。"""
    if key == "month_accrual":
        return row.get("amount") or Decimal("0")
    return row.invoice_amount or Decimal("0")


def _accrual_detail_filters(request, details):
    """4 张明细小表各自的筛选：过滤后的行 + 回显值 + 下拉选项 + 命中行数/金额合计。

    · 只作用于这 4 张明细小表；上表「到票与暂估」的数字和页面主筛选都不受影响。
    · 「导出Excel」会带上同一套参数，保证「筛选后导出＝所见」（见 _listing_rows）。
    """
    result = {}
    for key, params in ACCRUAL_DETAIL_FILTER_KEYS.items():
        # 统一成 period / supplier / material 三个字段名，4 张表就能共用同一段模板
        normalized = {
            "period": params.get("period") or params.get("month"),
            "supplier": params.get("supplier"),
            "material": params.get("material"),
        }
        normalized = {field: name for field, name in normalized.items() if name}
        rows = details[key]
        period_label, period_all_label = ACCRUAL_DETAIL_PERIOD_LABELS[key]
        values = {
            field: str(request.GET.get(name, "")).strip()
            for field, name in normalized.items()
        }
        want_period = values.get("period", "")
        want_supplier = values.get("supplier", "")
        want_material = values.get("material", "")

        shown = rows
        if want_period:
            shown = [
                row for row in shown
                if _detail_period_value(key, row)
                and _detail_period_value(key, row).strftime("%Y-%m") == want_period
            ]
        if want_supplier:
            needle = want_supplier.lower()
            shown = [row for row in shown if needle in _detail_supplier_text(key, row).lower()]
        if want_material:
            needle = want_material.lower()
            shown = [row for row in shown if needle in _detail_material_text(key, row).lower()]

        period_values = sorted(
            {_detail_period_value(key, row).strftime("%Y-%m")
             for row in rows if _detail_period_value(key, row)},
            reverse=True,
        )
        # 当前筛选值即使本表里没有，也留成选项，避免出现「看不见的筛选」
        if want_period and want_period not in period_values:
            period_values.insert(0, want_period)

        labels = []
        if want_period:
            labels.append(f"{period_label} {want_period}")
        if want_supplier:
            labels.append(f"供应商含「{want_supplier}」")
        if want_material:
            labels.append(f"物料编码/名称含「{want_material}」")

        # 提交本表筛选时，把其它表/页面主筛选的条件原样带上（题面隐藏域）
        own_names = set(normalized.values())
        hidden = [
            (name, value) for name, value in request.GET.items()
            if name not in own_names and str(value).strip()
        ]
        amount = sum((_detail_filter_amount(key, row) for row in shown), Decimal("0"))
        amount_all = sum((_detail_filter_amount(key, row) for row in rows), Decimal("0"))
        result[key] = {
            "key": key,
            "rows": shown,
            "shown": len(shown),
            "total": len(rows),
            "labels": labels,
            "filtered": bool(labels),
            "period_label": period_label,
            "period_all_label": period_all_label,
            "period_options": period_values,
            # 当前条件：模板回显用（下拉选中项、输入框里的文字）
            "period": want_period,
            "supplier": want_supplier,
            "material": want_material,
            "period_param": normalized.get("period", ""),
            "supplier_param": normalized.get("supplier", ""),
            "material_param": normalized.get("material", ""),
            "has_material": bool(normalized.get("material")),
            # 金额徽标：筛选后就显示「筛选后」合计（与表格里看到的行一致），未筛选＝上表本期数
            "amount": amount.quantize(Decimal("0.01")),
            "amount_all": amount_all.quantize(Decimal("0.01")),
            "amount_label": ACCRUAL_DETAIL_AMOUNT_LABELS[key],
            "export_kind": ACCRUAL_DETAIL_EXPORT_KIND_NAMES[key],
            "supplier_options": _detail_suggestion_values(
                [_detail_supplier_text(key, row) for row in rows]
            ),
            "material_options": _detail_suggestion_values(
                [_detail_material_text(key, row) for row in rows]
            ),
            "hidden": hidden,
            "export_query": urlencode({
                name: values[field] for field, name in normalized.items() if values.get(field)
            }),
            "clear_query": urlencode(dict(hidden)),
        }
    return result


def purchase_accrual_flow_tables(period, all_rows):
    """「到票与暂估」表：一行一个核算期间 + 一行合计。

    ⚠️ 关键：每一行都按**该期间当时的口径**独立计算 ——
    核销只取该期间的发票核销记录、余额只取截至该期末的未到票余额，
    所以以后月份新增数据**不会改动以前月份那一行**，历史月份的数据会被保留。
    期初暂估＝上一期间的期末暂估（首个期间取期初采购暂估台账余额）。
    「合同金额」＝归属到该期间的采购合同不含税金额（见 contract_amount_by_month）。
    返回 (flow_rows, flow_total)。
    """
    periods = list(AccountingPeriod.objects.order_by("month"))
    rows_by_period = {
        item.id: (all_rows if item.id == period.id else build_purchase_accrual_rows(item))
        for item in periods
    }
    contract_amounts = contract_amount_by_month()
    flow_rows = []
    previous_ending = None
    for index, item in enumerate(periods):
        stats = _purchase_flow_stats(item, rows_by_period[item.id])
        opening = previous_ending if index else _accrual_opening_balance(item)
        stats["opening_accrual"] = (opening or Decimal("0")).quantize(Decimal("0.01"))
        stats["contract_amount"] = contract_amounts.get(item.month, Decimal("0.00"))
        previous_ending = stats["ending_accrual"]
        stats["period"] = item
        stats["current"] = item.id == period.id
        flow_rows.append(stats)
    flow_total = {
        key: sum((row[key] for row in flow_rows), Decimal("0"))
        for key in ("same_month", "month_accrual", "invoice_matched", "cross",
                    "contract_amount")
    }
    for key in ("same_month_count", "month_accrual_count",
                "invoice_matched_count", "cross_count"):
        flow_total[key] = sum(row[key] for row in flow_rows)
    # 期初/期末暂估是余额、不可加：合计行分别取第一期期初、最后一期期末
    flow_total["opening_accrual"] = (
        flow_rows[0]["opening_accrual"] if flow_rows else Decimal("0")
    )
    flow_total["ending_accrual"] = (
        flow_rows[-1]["ending_accrual"] if flow_rows else Decimal("0")
    )
    return flow_rows, flow_total


def _purchase_receipt_row(period, txn):
    """「采购入库」页的一行：一条采购入库流水。

    period ＝ 参考期间（「累计到票 / 尚未到票」都截至该期末）；
    其余列＝数量、暂估单价/金额（来源见 provisional_basis）、最近到票期间、状态。
    """
    basis = provisional_basis(txn)
    quantity = abs(Decimal(txn.quantity or Decimal("0")))
    allocated = (
        txn.allocations.filter(
            kind=MatchAllocation.Kind.PURCHASE_INVOICE,
            period__month__lte=period.month,
        )
        .exclude(purchase_invoice_line__match_mode=MatchingMode.IGNORED)
        .aggregate(
            quantity=Sum("matched_quantity"),
            latest_period=Max("purchase_invoice_line__invoice_period"),
        )
    )
    invoiced = min(Decimal(allocated["quantity"] or Decimal("0")), quantity)
    outstanding = max(quantity - invoiced, Decimal("0"))
    return {
        "business_type": (txn.business_type or "").strip(),
        "receipt_period": txn.period,
        # 「入库期间」（用户 2026-09-21）：取**入库日期的年月**（如 2026-07），
        # 不用流水所属期间 —— 历史单据补录到本期时期间仍按单据日期算，也不会留空。
        "receipt_month": txn.txn_date.replace(day=1),
        "receipt_month_label": txn.txn_date.strftime("%Y-%m"),
        "txn_date": txn.txn_date,
        "document_no": txn.document_no,
        "supplier_name": txn.supplier_name or txn.supplier_code,
        "contract_no": txn.contract_no,
        "material_code": txn.material_code,
        "material_name": txn.material_name,
        "specification": txn.specification,
        "uom": txn.uom,
        "quantity": quantity,
        "provisional_unit_price": basis["unit_price"],
        "provisional_amount": basis["amount"],
        "price_source": basis["source"],
        "invoiced_quantity": invoiced,
        "outstanding_quantity": outstanding,
        "latest_invoice_period": allocated["latest_period"],
        "attribution_month": receipt_attribution_month(txn),
        "status": (
            "全部到票"
            if outstanding == Decimal("0")
            else "部分到票"
            if invoiced > Decimal("0")
            else "全部暂估"
        ),
    }


def _purchase_receipt_periods_up_to(period):
    """截至本期的所有核算期间（升序）。"""
    return list(
        AccountingPeriod.objects.filter(month__lte=period.month).order_by("month")
    )


def _purchase_receipt_rows(period, periods=None, filters=None):
    """「采购入库」明细行。

    · periods=None → 只取本期（默认）；
    · 传入期间列表 → 取这些期间录入的入库流水（用于「全部期间」）；
    · 到票进度一律**截至 period（页面当前期间）期末**，方便同一页横向比较；
    · 每行带「入库期间」（取入库日期的年月）。
    """
    if periods is None:
        periods = [period]
    include_returns = bool((filters or {}).get("include_returns"))
    return [
        _purchase_receipt_row(period, txn)
        for item in periods
        for txn in purchase_inbound_rows_of_period(
            item, include_returns=include_returns
        )
    ]


def _purchase_receipt_all_rows(period, filters=None):
    """本期及以前所有期间录入的采购侧入库流水（到票进度截至 period）。"""
    return _purchase_receipt_rows(
        period, _purchase_receipt_periods_up_to(period), filters=filters
    )


def _purchase_receipt_filters(request, period):
    """采购入库明细的筛选（用户 2026-09-21）：入库期间 / 供应商 / 物料编码。

    只影响**明细表与导出**；上面的「各期间采购入库合计」始终是全量。
    入库期间：current（默认，本期）/ all（全部期间）/ YYYY-MM（按入库日期的年月）。
    """
    return {
        # 入库期间：默认「全部期间」（用户 2026-09-21 去掉了「本期」选项）
        "period": str(request.GET.get("receipt_period", "")).strip() or "all",
        "supplier": str(request.GET.get("supplier", "")).strip(),
        "material": str(request.GET.get("material", "")).strip(),
        # 「入库类型」：页面上的筛选项已按用户要求删除；
        # 这里保留 URL 参数（?type=with_return），需要与「库存结转」对账时仍可用。
        "include_returns": str(request.GET.get("type", "")).strip() == "with_return",
    }


def _filtered_purchase_receipt_rows(request, period):
    """按页面筛选条件（入库期间 / 供应商 / 物料编码）取明细行。"""
    filters = _purchase_receipt_filters(request, period)
    rows = _purchase_receipt_all_rows(period, filters=filters)
    if filters["period"] == "current":
        current_label = period.month.strftime("%Y-%m")
        rows = [row for row in rows if row["receipt_month_label"] == current_label]
    elif filters["period"] != "all":
        rows = [row for row in rows if row["receipt_month_label"] == filters["period"]]
    if filters["supplier"]:
        needle = filters["supplier"].lower()
        rows = [
            row for row in rows if needle in str(row["supplier_name"] or "").lower()
        ]
    if filters["material"]:
        needle = filters["material"].lower()
        rows = [
            row
            for row in rows
            if needle
            in ("%s %s" % (row["material_code"], row["material_name"])).lower()
        ]
    return rows, filters


def _purchase_receipt_month_options(period):
    """「入库期间」下拉可选项：有入库流水的年月（升序），值＝YYYY-MM。"""
    labels = sorted(
        {row["receipt_month_label"] for row in _purchase_receipt_all_rows(period)}
    )
    return [{"value": label, "label": label} for label in labels]


def _purchase_receipt_category_map(period):
    """物料编码 → 物料分类，**跨期间**查找（截至本期）。

    为什么不能只看本期：采购汇总是「截至本期的全部到货」，历史月份的物料
    在本期可能一笔流水都没有；只按本期取分类，历史行会全变成「未分类」。
    取法：从**最近的期间往回**逐期合并（近期的分类优先），分类来源见 inventory_category_map。
    """
    mapping = {}
    for item in reversed(_purchase_receipt_periods_up_to(period)):
        for code, category in inventory_category_map(item).items():
            mapping.setdefault(code, category)
    return mapping


def _purchase_receipt_period_summary(period, filters=None):
    """「采购汇总」：一行一个 **入库期间（入库日期的年月）× 物料分类** ＋ 入库数量、暂估金额。

    用户 2026-09-21：本表只留 入库期间 / 物料分类 / 入库数量 / 暂估金额 四列
    （原来的 入库单数、累计到票数量、尚未到票数量 已删除；「物料分类」为同日新增）。
    物料分类取法见 inventory_category_map（期初/期末存货、库存流水、投料记录，取不到记「未分类」）。
    """
    category_by_code = _purchase_receipt_category_map(period)
    buckets = {}
    for row in _purchase_receipt_all_rows(period, filters=filters):
        category = category_by_code.get(row["material_code"]) or "未分类"
        item = buckets.setdefault(
            (row["receipt_month_label"], category),
            {
                "month_label": row["receipt_month_label"],
                "category": category,
                "quantity": Decimal("0"),
                "amount": Decimal("0"),
                "uncleared_price": 0,
            },
        )
        item["quantity"] += row["quantity"]
        if row["provisional_amount"] is None:
            item["uncleared_price"] += 1
        else:
            item["amount"] += Decimal(row["provisional_amount"])

    summary = []
    # 同一个入库期间里，按分类名排序（「未分类」排最后）
    for label, category in sorted(
        buckets, key=lambda key: (key[0], key[1] == "未分类", key[1])
    ):
        item = buckets[(label, category)]
        item["period"] = next(
            (
                p
                for p in _purchase_receipt_periods_up_to(period)
                if p.month.strftime("%Y-%m") == label
            ),
            None,
        )
        item["is_current"] = label == period.month.strftime("%Y-%m")
        summary.append(item)
    return summary


@login_required
def ending_wip(request, period_id):
    """期末在产品：按生产工单列本期期末的在制余额（材料 / 人工 / 制造费用）。

    口径与「总成本表」的「在产品成本」块一致（都取工单结算结果 wip_*）；
    本页到工单级，并给出车间汇总。只列**还有在产品**的工单。
    """
    period = get_object_or_404(AccountingPeriod, id=period_id)
    workshop = str(request.GET.get("workshop", "")).strip()
    material = str(request.GET.get("material", "")).strip()
    status = str(request.GET.get("status", "")).strip()
    data = build_ending_wip_data(
        period, workshop=workshop, material=material, status=status
    )
    return render(
        request,
        "settlement/ending_wip.html",
        {
            "period": period,
            "page": _page(request, data["rows"]),
            "workshop_rows": data["workshop_rows"],
            # 汇总/合计＝全量口径（不被明细筛选带走）
            "total": data["total"],
            "filtered_total": data["filtered_total"],
            "row_count": len(data["rows"]),
            "all_row_count": len(data["all_rows"]),
            "wip_filter": {
                "workshop": workshop,
                "material": material,
                "status": status,
            },
            "wip_filter_active": bool(workshop or material or status),
            "workshop_options": ending_wip_workshop_options(period),
            "status_options": ending_wip_status_options(period),
        },
    )


@login_required
def purchase_receipts(request, period_id):
    """采购入库：本期间录入的采购入库流水（到货台账）。

    与「月末采购暂估」的区别：那里按暂估**归属月份**取「仍未到票」的余额，
    这里按流水**上传期间**列出本期到的货，含每单的暂估单价/金额与到票进度。
    """
    period = get_object_or_404(AccountingPeriod, id=period_id)
    # 明细：按 入库期间 / 供应商 / 物料编码 筛选（2026-09-21 用户要求）；
    # 上面的「各期间采购入库合计」不受筛选影响，始终是全量。
    rows, filters = _filtered_purchase_receipt_rows(request, period)
    return render(
        request,
        "settlement/purchase_receipts.html",
        {
            "period": period,
            "page": _page(request, rows),
            # 「采购汇总」是全量口径（一行一个入库年月）：只跟「入库类型」走，
            # 不随 入库期间 / 供应商 / 物料编码 这三个明细筛选变化（页面已注明）。
            "summary_rows": _purchase_receipt_period_summary(
                period, {"include_returns": filters["include_returns"]}
            ),
            "receipt_count": len(rows),
            "receipt_filter": filters,
            "receipt_filter_active": any(
                filters[key] and filters[key] != "all"
                for key in ("period", "supplier", "material")
            ) or filters["include_returns"],
            "receipt_period_options": _purchase_receipt_month_options(period),
        },
    )


@login_required
def purchase_accruals(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    all_rows = build_purchase_accrual_rows(period)
    # 页面筛选：归属月份 / 供应商 / 物料编码（三个可见字段，见 _accrual_page_filters）
    rows = _accrual_page_filters(request, all_rows)
    flow_rows, flow_total = purchase_accrual_flow_tables(period, all_rows)
    # 「暂估价差」列（用户 2026-09-21）：该期间发票核销产生的价差合计 ＝ 发票金额 − 暂估金额。
    # 每条核销只归属一个期间，所以各期可相加。
    diff_by_month = {
        item.month: summarize_purchase_differences(build_purchase_difference_rows(item))[0][
            "difference_amount"
        ]
        for item in AccountingPeriod.objects.order_by("month")
    }
    for row in flow_rows:
        row["difference_amount"] = diff_by_month.get(
            row["period"].month, Decimal("0.00")
        )
    flow_total["difference_amount"] = sum(
        (row["difference_amount"] for row in flow_rows), Decimal("0.00")
    )
    # 明细块都是「本期」口径，徽标必须用本期那一行；
    # flow_total 是跨期间合计（期初/期末取首末期），只适合上表的合计行。
    flow_current = next((row for row in flow_rows if row["current"]), None) or {}
    details = purchase_accrual_detail_tables(period, all_rows)
    # 4 张明细小表各自的筛选（用户 2026-09-19）：只影响这几张表；上表数字仍是本期全口径，
    # 页面主筛选（归属月份/供应商/物料编码）也照旧只管下面那张主明细表。
    detail_filters = _accrual_detail_filters(request, details)
    details = {**details, **{key: item["rows"] for key, item in detail_filters.items()}}
    # 页面主筛选栏提交时，把这 4 张明细表已生效的条件也带上（否则一筛主表就把明细条件冲掉）
    detail_filter_hidden = [
        (name, value)
        for item in detail_filters.values()
        for name, value in (
            (item["period_param"], item["period"]),
            (item["supplier_param"], item["supplier"]),
            (item["material_param"], item["material"]),
        )
        if name and value
    ]
    # 「入库单号」若全是虚拟编号（空 / 纯数字系统编号）就整列不显示（按全部明细判断列是否保留）
    show_doc_no = any(not _is_virtual_document_no(row.document_no) for row in all_rows)
    # 方案A 的保底提示（用户 2026-09-19 选A：合同表管到货/暂估，发票表管到票/实际价）：
    # 本期没上传采购发票时，「本月到票 / 当月到货到票 / 本月回冲」会全是 0，
    # 容易被误读成"本期没有到票"（其实是没数据，未到票的采购入库会一直挂着暂估）。
    no_invoice_data = period.purchase_invoices.count() == 0
    return render(
        request,
        "settlement/purchase_accruals.html",
        {
            "period": period,
            "cutoff_date": period_month_end(period.month),
            "page": _page(request, rows),
            "flow_rows": flow_rows,
            "flow_total": flow_total,
            "flow_current": flow_current,
            "details": details,
            "detail_filters": detail_filters,
            "detail_filter_hidden": detail_filter_hidden,
            "show_doc_no": show_doc_no,
            "no_invoice_data": no_invoice_data,
            "receipt_months": _accrual_receipt_month_options(request, all_rows),
            "filter_labels": _accrual_filter_labels(request),
            "filtered_count": len(rows),
            "total_count": len(all_rows),
            # 「月末采购暂估」筛选行的金额徽标（用户 2026-09-21）：筛选后＝表格里看到的行合计
            "filtered_amount": sum(
                (r.provisional_amount or Decimal("0") for r in rows), Decimal("0")
            ),
            "export_query": _accrual_export_query(request),
            # 「本月采购暂估差异」三块（2026-09-21 并入本页，只列有价差的行）
            **_purchase_difference_blocks(request, period),
            # 「回冲明细（逐笔）」板块（2026-09-21 由「暂估回冲」页并入本页）
            **_purchase_reversal_blocks(request, period),
        },
    )


def _difference_page_filters(request, rows):
    """本月采购暂估差异的筛选，与页面筛选栏一一对应。

    页面与「各期间暂估差异」汇总表共用，保证两处口径一致（筛选后汇总也跟着变）。
    """
    invoice_no = str(request.GET.get("invoice_no", "")).strip().lower()
    receipt_month = str(request.GET.get("receipt_month", "")).strip()
    document_no = str(request.GET.get("document_no", "")).strip().lower()
    supplier = str(request.GET.get("supplier", "")).strip().lower()
    material = str(request.GET.get("material", "")).strip().lower()
    status = str(request.GET.get("status", "")).strip()
    cross_period = str(request.GET.get("cross_period", "")).strip()
    if invoice_no:
        rows = [
            row
            for row in rows
            if invoice_no in (row.invoice.invoice_no or "").lower()
        ]
    if receipt_month:
        rows = [
            row
            for row in rows
            if row.receipt_month.strftime("%Y-%m") == receipt_month
        ]
    if document_no:
        rows = [
            row
            for row in rows
            if document_no in (row.document_no or "").lower()
        ]
    if supplier:
        rows = [
            row
            for row in rows
            if supplier
            in (
                f"{row.invoice.supplier_code} "
                f"{row.invoice.supplier_name}"
            ).lower()
        ]
    if material:
        rows = [
            row
            for row in rows
            if material
            in (
                f"{row.material_code} "
                f"{row.material_name}"
            ).lower()
        ]
    if status:
        rows = [row for row in rows if row.status == status]
    if cross_period in {"yes", "no"}:
        rows = [
            row
            for row in rows
            if row.cross_period == (cross_period == "yes")
        ]
    return rows


def purchase_difference_period_summary(request, period):
    """各期间暂估差异：一行一个核算期间 + 一行合计。

    与明细套用同一套筛选。这里各期间**可以相加**：每条发票核销只归属一个核算期间
    （`build_purchase_difference_rows` 按 MatchAllocation.period 取数），
    所以合计＝各期间之和，正好是各月交总账的价差合计。
    """
    rows = []
    for item in AccountingPeriod.objects.order_by("month"):
        period_rows = _difference_page_filters(
            request, build_purchase_difference_rows(item)
        )
        totals, _ = summarize_purchase_differences(period_rows)
        totals["period"] = item
        totals["current"] = item.id == period.id
        rows.append(totals)

    def _sum(key):
        return sum((row[key] for row in rows), Decimal("0"))

    grand_total = {
        "matched_quantity": _sum("matched_quantity"),
        "provisional_amount": _sum("provisional_amount"),
        "invoice_amount": _sum("invoice_amount"),
        "difference_amount": _sum("difference_amount"),
        "cross_period_count": _sum("cross_period_count"),
        "pending_count": _sum("pending_count"),
    }
    return rows, grand_total


def _purchase_difference_blocks(request, period):
    """「本月采购暂估差异」并入「月末采购暂估」页后用的三块数据（用户 2026-09-21）。

    · 各期间汇总 / 按供应商物料汇总 / 价差明细 都**只保留真有暂估价差的核销行**
      （暂估金额 ≠ 发票金额）；无价差的行不显示；
    · 价差明细单独分页（参数 diff_page），避免与页面主表的 page 冲突。
    """
    rows = _difference_page_filters(request, build_purchase_difference_rows(period))

    def has_difference(row):
        return (
            row.provisional_amount is not None
            and (row.difference_amount or Decimal("0")) != 0
        )

    diff_rows = [row for row in rows if has_difference(row)]
    pending_count = len([row for row in rows if row.provisional_amount is None])
    totals, grouped_rows = summarize_purchase_differences(diff_rows)
    grouped_rows = [
        item
        for item in grouped_rows
        if (item["difference_amount"] or Decimal("0")) != 0
    ]

    summary_rows, _summary_total = purchase_difference_period_summary(request, period)
    summary_rows = [
        row
        for row in summary_rows
        if (row["difference_amount"] or Decimal("0")) != 0
    ]
    keys = (
        "matched_quantity",
        "provisional_amount",
        "invoice_amount",
        "difference_amount",
        "cross_period_count",
        "pending_count",
    )
    summary_total = {
        key: sum((row[key] for row in summary_rows), Decimal("0")) for key in keys
    }

    page = Paginator(diff_rows, 100).get_page(request.GET.get("diff_page"))
    query = request.GET.copy()
    query.pop("diff_page", None)
    page.querystring = query.urlencode()

    return {
        "diff_totals": totals,
        "diff_grouped_rows": grouped_rows,
        "diff_summary_rows": summary_rows,
        "diff_summary_total": summary_total,
        "diff_page": page,
        "diff_pending_count": pending_count,
        "diff_row_count": len(diff_rows),
    }


@login_required
def purchase_differences(request, period_id):
    """「本月采购暂估差异」已并入「月末采购暂估」页（用户 2026-09-21）。

    这里保留 URL 只做跳转，老链接 / 书签照样能用（查询参数原样带过去）。
    """
    url = reverse("purchase_accruals", args=[period_id])
    query = request.GET.urlencode()
    return redirect("%s?%s" % (url, query) if query else url)


def _sales_voucher_param_names():
    """凭证自己的筛选参数名（「清空筛选」时只清这些，保留页面上其它参数）。"""
    names = {"voucher_month", "vpage"}
    for index in range(len(CARRY_COST_TYPES)):
        names |= set(block_param_names(index).values())
    return names


@login_required


def _sales_voucher_filters(request):
    """销售结转成本凭证的筛选解析：记账月份 + 三个结转类型块各自的出库期间/开票状态/开票期间。

    参数名带块序号（vout0/vst0/vinv0、vout1/...），保证三块互不干扰；
    没提交过的字段传 None → 用该块默认值；提交了空串 = 选了「（空白）」= 该字段不限制。
    """
    voucher_month_raw = str(request.GET.get("voucher_month", "")).strip()
    voucher_month = None
    if voucher_month_raw:
        try:
            year, month_no = voucher_month_raw.split("-")
            voucher_month = date(int(year), int(month_no), 1)
        except (TypeError, ValueError):
            voucher_month = None
    block_filters = {}
    for index, carry in enumerate(CARRY_COST_TYPES):
        names = block_param_names(index)
        block_filters[carry] = {
            "out": request.GET.get(names["out"]),
            "status": request.GET.get(names["status"]),
            "invoice": request.GET.get(names["invoice"]),
        }
    return voucher_month, block_filters


@login_required
def sales_cost_voucher(request, period_id):
    """销售结转成本凭证：凭证分录口径（与采购暂估凭证同格式）。

    - 数据来源：销售台账（结转类型由台账带来；金额由平台结算后按销售数量 × 月末加权平均价回填）；
    - 凭证期间：voucher_month（记账月份）；平台把「前期发货本期开票」记在开票月，其余记在出库月；
    - 三个结转类型块各自独立筛选（出库期间 / 开票状态 / 开票期间）；
    - 借方：主营业务成本（同月开票、前期发货本期开票）/ 发出商品（未开票）；
      贷方：库存商品（同月开票、未开票）/ 发出商品（前期发货本期开票）。
    """
    period = get_object_or_404(AccountingPeriod, id=period_id)
    voucher_month, block_filters = _sales_voucher_filters(request)
    group = str(request.GET.get("group", "")).strip() or "category"
    voucher = build_sales_voucher_entries(
        period, month=voucher_month, block_filters=block_filters, group=group
    )
    voucher_clear_query = urlencode(
        [
            (key, value)
            for key, value in request.GET.items()
            if key not in _sales_voucher_param_names()
        ]
    )
    return render(
        request,
        "settlement/sales_cost_voucher.html",
        {
            "period": period,
            "voucher": voucher,
            "voucher_clear_query": voucher_clear_query,
        },
    )


@login_required
def sales_uninvoiced(request, period_id):
    """月末已出库未开票（台账口径，用户 2026-09-21 方案A）。

    明细＝你上传的「销售台账」里 结转类型=「当月发货未开票挂发出商品」的那些行；
    发出商品滚动＝期初发出商品 ＋ 本期台账未开票 － 本期台账「前期发货本期开票结转成本」。
    不再用「销售发票核销库存流水」（那条线只匹配上 12% 的数量，会把发出商品算虚增 524 万）。

    「销售结转成本凭证」已拆成独立页面（sales_cost_voucher），本页不再展示。
    """
    period = get_object_or_404(AccountingPeriod, id=period_id)
    rows = build_sales_uninvoiced_rows(period)
    # 明细表筛选（用户 2026-09-27）：期间 + 客户；与「导出Excel」共用同一套筛选函数，
    # 保证页面看到的行数与导出的行数完全一致。
    issue_filter = str(request.GET.get("issue_month", "")).strip()
    customer_filter = str(request.GET.get("customer", "")).strip()
    filtered_rows = filter_sales_uninvoiced_rows(
        rows, issue_month=issue_filter, customer=customer_filter
    )
    detail_filter = {
        "issue_month": issue_filter,
        "customer": customer_filter,
        "active": bool(issue_filter or customer_filter),
    }
    return render(
        request,
        "settlement/sales_uninvoiced.html",
        {
            "period": period,
            "page": _page(request, filtered_rows),
            # 明细表页脚：跟着筛选走；顶部六个数是期末口径，不受筛选影响
            "detail_total": summarize_uninvoiced_rows(filtered_rows),
            "detail_filter": detail_filter,
            "period_options": uninvoiced_period_options(period),
            "rollforward": build_goods_issued_rollforward(period),
            "check_rows": build_goods_issued_check_rows(period),
            "statistics": build_sales_cost_statistics(period),
        },
    )



def _ledger_effective_invoice_month(request, period):
    """期间筛选的实际取值：未传参数 -> 当期核算月份；传空 -> 全部。"""
    if "invoice_month" not in request.GET:
        return period.month.strftime("%Y-%m")
    return str(request.GET.get("invoice_month") or "").strip()


def _sales_ledger_month(value):
    """把 YYYY-MM 解析成月首日期；解析不出来就返回 None（等于不筛选）。"""
    raw = str(value or "").strip()
    if not raw:
        return None
    try:
        year, month_no = raw.split("-")[:2]
        return date(int(year), int(month_no), 1)
    except (TypeError, ValueError):
        return None


def sales_ledger_queryset(request, period):
    """销售台账的筛选口径：页面与「导出Excel」共用，保证“筛选后导出 = 所见”。

    可选筛选（不填即全部，全部为下拉选择）：
    invoice_month 开票期间 YYYY-MM；out_month 出库期间 YYYY-MM；
    invoice_status 开票状态；category 销售大类；channel 销售渠道。
    """
    queryset = period.sales_ledger_lines.select_related("material")

    out_month = _sales_ledger_month(request.GET.get("out_month"))
    if out_month:
        queryset = queryset.filter(out_month=out_month)
    # 期间默认当月：不带 invoice_month 参数时取当期核算月份；
    # 显式传空（?invoice_month=）才表示「全部」。
    invoice_month = _sales_ledger_month(_ledger_effective_invoice_month(request, period))
    if invoice_month:
        queryset = queryset.filter(invoice_month=invoice_month)

    status = str(request.GET.get("invoice_status", "")).strip()
    if status:
        queryset = queryset.filter(invoice_status=status)

    # 销售渠道 / 销售大类：下拉选值（包含匹配，兼容手写参数）
    for parameter, field in (("channel", "channel"), ("category", "sales_category")):
        value = str(request.GET.get(parameter, "")).strip()
        if value:
            queryset = queryset.filter(**{f"{field}__icontains": value})

    # 与「报表导出 → 销售台账（成本·毛利）」同一次序；同号时按主键，结果稳定
    return queryset.order_by("out_month", "row_no", "id")


SALES_CATEGORY_ORDER = ("纱线", "面料", "纤维", "其他成品")


def _ledger_category_label(value):
    """大类显示名：去掉括号里的补充说明（如「其他成品（手套、服装、床垫等）」->「其他成品」）。"""
    text = (value or "").strip()
    if not text:
        return "（未填大类）"
    for sep in ("（", "("):
        index = text.find(sep)
        if index > 0:
            text = text[:index].strip()
    return text or value


def _ledger_category_sort_key(label):
    """大类排序：纱线 → 面料 → 纤维 → 其他成品 → 其余按名称。"""
    if label in SALES_CATEGORY_ORDER:
        return (0, SALES_CATEGORY_ORDER.index(label), "")
    return (1, len(SALES_CATEGORY_ORDER), label)


def _ledger_category_options(ledger):
    """销售大类下拉选项：去重 + 去括号显示名 + 固定顺序。

    注意：必须先 .order_by() 清掉 Model.Meta.ordering，否则 distinct()
    会把 batch_id/row_no 也带进 SELECT，导致筛选下拉出现大量重复选项。
    """
    seen = {}
    for value in (
        ledger.order_by().values_list("sales_category", flat=True).distinct()
    ):
        if not value:
            continue
        seen.setdefault(value, _ledger_category_label(value))
    return [
        {"value": value, "label": label}
        for value, label in sorted(
            seen.items(), key=lambda pair: _ledger_category_sort_key(pair[1])
        )
    ]


def _ledger_amount(value):
    return (value or Decimal("0")).quantize(Decimal("0.01"))


def _ledger_margin(revenue, profit):
    """毛利率 = 毛利 ÷ 销售收入（收入为 0 时不计算）。"""
    base = Decimal(revenue or 0)
    if not base:
        return None
    return Decimal(profit or 0) / base


def summarize_sales_ledger(queryset):
    """台账明细的合计（当前筛选）。

    口径与「损益表」一致：**毛利 = 终端销售收入 − 销售成本**，
    成本缺失（该产品本月没有库存计价）的行按成本 0 计，毛利会偏高，页面会提示；
    毛利率 = 毛利 ÷ 终端销售收入。
    """
    rows = list(
        queryset.order_by().values_list(
            "quantity", "terminal_amount", "lb_amount", "cost_amount"
        )
    )
    zero = Decimal("0")
    quantity = sum((Decimal(r[0] or 0) for r in rows), zero)
    revenue = sum((Decimal(r[1] or 0) for r in rows), zero)
    lb_amount = sum((Decimal(r[2] or 0) for r in rows), zero)
    cost = sum((Decimal(r[3] or 0) for r in rows), zero)
    # 毛利逐行取整后再合计，与「损益表」的 _q(收入-成本) 完全一致
    gross = sum(
        ((Decimal(r[1] or 0) - Decimal(r[3] or 0)).quantize(Decimal("0.01")) for r in rows),
        zero,
    )
    totals = {
        "rows": len(rows),
        "quantity": quantity,
        "terminal_amount": revenue,
        "lb_amount": lb_amount,
        "cost_amount": cost,
        "gross_profit": gross.quantize(Decimal("0.01")),
    }
    totals["margin"] = _ledger_margin(revenue, totals["gross_profit"])
    # 有收入却没成本单价的行：毛利会偏高，页面要提示（与损益表提示口径一致）
    totals["missing_cost_rows"] = queryset.filter(
        Q(cost_price__isnull=True) | Q(cost_price=zero)
    ).count()
    return totals


def summarize_sales_ledger_by_group(queryset, basis="invoice_month"):
    """按「期间 × 销售大类」汇总（页面上的销售毛利汇总表）。

    期间口径默认 invoice_month（开票期间），与损益表的「按开票期间」一致；
    毛利 = 终端销售收入 − 销售成本（成本缺失按 0），毛利率 = 毛利 ÷ 收入。
    """
    field = "invoice_month" if basis == "invoice_month" else "out_month"
    zero = Decimal("0")
    # 必须先清掉默认排序：否则 order_by 的字段会被带进 GROUP BY，分组会变成逐行
    buckets = {}
    for month, category, revenue, cost in queryset.order_by().values_list(
        field, "sales_category", "terminal_amount", "cost_amount"
    ):
        key = (month, category or "")
        bucket = buckets.setdefault(
            key, {"revenue": zero, "cost": zero, "gross": zero}
        )
        row_revenue = Decimal(revenue or 0)
        row_cost = Decimal(cost or 0)
        bucket["revenue"] += row_revenue
        bucket["cost"] += row_cost
        # 逐行取整后再合计，与「损益表」一致
        bucket["gross"] += (row_revenue - row_cost).quantize(Decimal("0.01"))

    rows = []
    for (month, category), bucket in buckets.items():
        revenue = _ledger_amount(bucket["revenue"])
        cost = _ledger_amount(bucket["cost"])
        profit = bucket["gross"].quantize(Decimal("0.01"))
        if not revenue and not cost:
            continue  # 收入、成本都为 0 的组不占位置
        rows.append(
            {
                "month": month,
                "month_label": month.strftime("%Y-%m") if month else "未开票",
                "category_label": _ledger_category_label(category),
                "terminal_amount": revenue,
                "cost_amount": cost,
                "gross_profit": profit,
                "margin": _ledger_margin(revenue, profit),
            }
        )
    # 排序：无期间的行放最后；其余按期间，期间内按固定大类顺序
    rows.sort(key=lambda row: (
        row["month"] is None,
        row["month"] or date.min,
        _ledger_category_sort_key(row["category_label"]),
    ))

    groups = []
    total = {"terminal_amount": zero, "cost_amount": zero, "gross_profit": zero}
    for row in rows:
        if not groups or groups[-1]["month"] != row["month"]:
            groups.append(
                {
                    "month": row["month"],
                    "month_label": row["month_label"],
                    "rows": [],
                    "subtotal": {"terminal_amount": zero, "cost_amount": zero,
                                 "gross_profit": zero},
                }
            )
        group = groups[-1]
        group["rows"].append(row)
        for key in ("terminal_amount", "cost_amount", "gross_profit"):
            group["subtotal"][key] += row[key]
            total[key] += row[key]

    for group in groups:
        group["subtotal"]["margin"] = _ledger_margin(
            group["subtotal"]["terminal_amount"], group["subtotal"]["gross_profit"]
        )
    total["margin"] = _ledger_margin(total["terminal_amount"], total["gross_profit"])
    return {"basis": field, "groups": groups, "total": total, "row_count": len(rows)}


@login_required
def sales_ledger(request, period_id):
    """销售台账：逐行查看出库/开票、数量、收入、成本单价、销售成本、毛利与毛利率。"""
    period = get_object_or_404(AccountingPeriod, id=period_id)
    queryset = sales_ledger_queryset(request, period)
    ledger = period.sales_ledger_lines
    # 汇总表固定按「开票期间 × 销售大类」汇总：
    # 页面上的「期间」筛选就是开票期间（与台账明细的「开票期间」同一个条件）。
    month_values = (
        ledger.values_list("out_month", flat=True).distinct().order_by("out_month")
    )
    invoice_month_values = (
        ledger.exclude(invoice_month__isnull=True)
        .values_list("invoice_month", flat=True)
        .distinct()
        .order_by("invoice_month")
    )
    page = _page(request, queryset, size=200)
    # 明细行也按「收入 - 成本」口径给出毛利与毛利率（缺成本按 0），与损益表、汇总表一致
    for item in page.object_list:
        revenue = Decimal(item.terminal_amount or 0)
        cost = Decimal(item.cost_amount or 0)
        gross = (revenue - cost).quantize(Decimal("0.01"))
        item.ledger_revenue = revenue.quantize(Decimal("0.01"))
        item.ledger_cost = cost.quantize(Decimal("0.01"))
        item.ledger_gross = gross
        item.ledger_margin = (gross / revenue) if revenue else None
        item.ledger_category = _ledger_category_label(item.sales_category)
    return render(
        request,
        "settlement/sales_ledger.html",
        {
            "period": period,
            "page": page,
            "summary": summarize_sales_ledger(queryset),
            "grouped": summarize_sales_ledger_by_group(queryset, "invoice_month"),
            "selected_invoice_month": _ledger_effective_invoice_month(request, period),
            "invoice_statuses": SalesLedgerLine.InvoiceStatus.choices,
            "months": [
                {"value": item.strftime("%Y-%m"), "label": item.strftime("%Y-%m")}
                for item in month_values
                if item
            ],
            "invoice_months": [
                {"value": item.strftime("%Y-%m"), "label": item.strftime("%Y-%m")}
                for item in invoice_month_values
                if item
            ],
            # 下拉显示名去掉括号补充；value 仍是原值，筛选口径不变
            "categories": _ledger_category_options(ledger),
            "channels": sorted(
                {
                    value
                    for value in ledger.values_list("channel", flat=True).distinct()
                    if value
                }
            ),
        },
    )


@login_required
def voucher_center(request, period_id):
    """凭证中心：本期间全部成本凭证一览（已实现的可打开明细/导出，规划中的列取数口径）。"""
    period = get_object_or_404(AccountingPeriod, id=period_id)
    from settlement.services.backfill_material_voucher import (
        backfill_material_voucher_totals,
    )

    _backfill = backfill_material_voucher_totals(period)
    return render(
        request,
        "settlement/voucher_center.html",
        {
            "period": period,
            "backfill_voucher": _backfill,
            "voucher_center": build_voucher_center(
                period,
                sales_month=request.GET.get("sales_month"),
                purchase_month=request.GET.get("purchase_month"),
                material_issue_month=request.GET.get("material_issue_month"),
                receipt_month=request.GET.get("receipt_month"),
            ),
        },
    )


@login_required
def purchase_accrual_voucher(request, period_id):
    """采购暂估 / 暂估回冲凭证（独立页面）。

    业务原理：本月到货本月到票 → 直接入账；本月到货月末未到票 → 暂估；
    前期暂估本月到票 → 冲回原暂估 + 按票入账（价差进材料成本差异）。
    凭证金额不随明细筛选变化（借贷必须相等）；筛选只过滤下面三张附件。
    """
    period = get_object_or_404(AccountingPeriod, id=period_id)
    group = str(request.GET.get("group", "")).strip() or "category"
    # 凭证期间筛选：voucher_month=YYYY-MM，默认当前核算期间
    raw_month = str(request.GET.get("voucher_month", "")).strip()
    target_period = period
    if raw_month:
        try:
            year, month_no = raw_month.split("-")
            target_period = (
                AccountingPeriod.objects.filter(
                    month=date(int(year), int(month_no), 1)
                ).first()
                or period
            )
        except (TypeError, ValueError):
            target_period = period
    voucher = build_purchase_voucher(
        target_period,
        group=group,
        detail_supplier=request.GET.get("supplier", ""),
        detail_material=request.GET.get("material", ""),
    )
    return render(
        request,
        "settlement/purchase_accrual_voucher.html",
        {
            "period": period,
            "target_period": target_period,
            "voucher": voucher,
            "period_options": [
                {"id": item.id, "value": item.month.strftime("%Y-%m"),
                 "label": str(item)}
                for item in AccountingPeriod.objects.order_by("month")
            ],
        },
    )


@login_required
def material_issue_voucher(request, period_id):
    """生产领料（投料）凭证（独立页面），格式与「采购暂估 / 回冲凭证」一致。

    摘要按车间：{月}月{车间}领用{物料大类}。
    借方 生产成本—直接材料—XX车间（一行一个车间 × 物料大类）；
    贷方 原材料／库存商品（按科目合并，不分块）。
    取库存出库流水口径：投料出库 ＋ 无工单领料 − 投料回撤。
    凭证金额不随附件筛选变化（借贷必须相等）；筛选只过滤下方附件。
    """
    period = get_object_or_404(AccountingPeriod, id=period_id)
    # 凭证期间筛选：voucher_month=YYYY-MM，默认当前核算期间
    raw_month = str(request.GET.get("voucher_month", "")).strip()
    target_period = period
    if raw_month:
        try:
            year, month_no = raw_month.split("-")
            target_period = (
                AccountingPeriod.objects.filter(
                    month=date(int(year), int(month_no), 1)
                ).first()
                or period
            )
        except (TypeError, ValueError):
            target_period = period
    voucher = build_material_issue_voucher(
        target_period,
        detail_workshop=request.GET.get("workshop", ""),
        detail_material=request.GET.get("material", ""),
    )
    from settlement.services.backfill_material_voucher import backfill_material_voucher_totals
    _bf = backfill_material_voucher_totals(period)
    return render(
        request,
        "settlement/material_issue_voucher.html",
        {
            "backfill_voucher": _bf,
            "period": period,
            "target_period": target_period,
            "voucher": voucher,
            "period_options": [
                {"id": item.id, "value": item.month.strftime("%Y-%m"),
                 "label": str(item)}
                for item in AccountingPeriod.objects.order_by("month")
            ],
        },
    )


@login_required
def finished_receipt_voucher(request, period_id):
    """完工入库结转凭证（独立明细页，用户 2026-09-27）。

    摘要按车间、科目按物料分类；贷方改成成本要素：生产成本—直接材料 / 直接人工 / 制造费用
    （全车间汇总，各一行；浆粕车间也并进这三个要素）。
    金额基准＝工单完工成本（＝报工数量 × 完工单位成本，与「工单成本」页一致，不用库存流水）；
    浆粕车间没有工单，按库存流水「生产入库」归集，再按浆粕车间成本归集拆成这三个要素。
    凭证金额不随附件筛选变化（借贷必须相等）；筛选只过滤下方附件。
    """
    period = get_object_or_404(AccountingPeriod, id=period_id)
    # 凭证期间筛选：voucher_month=YYYY-MM，默认当前核算期间
    raw_month = str(request.GET.get("voucher_month", "")).strip()
    target_period = period
    if raw_month:
        try:
            year, month_no = raw_month.split("-")
            target_period = (
                AccountingPeriod.objects.filter(
                    month=date(int(year), int(month_no), 1)
                ).first()
                or period
            )
        except (TypeError, ValueError):
            target_period = period
    voucher = build_finished_receipt_voucher(target_period)
    return render(
        request,
        "settlement/finished_receipt_voucher.html",
        {
            "period": period,
            "target_period": target_period,
            "voucher": voucher,
            "period_options": [
                {"id": item.id, "value": item.month.strftime("%Y-%m"),
                 "label": str(item)}
                for item in AccountingPeriod.objects.order_by("month")
            ],
        },
    )


def _work_order_product(period, order_no):
    report = (
        period.work_reports.filter(work_order_no=order_no)
        .exclude(material_code="")
        .first()
    )
    if report:
        return report.material_code, report.material_name
    return "", ""


@login_required
def inventory_ledger(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    txns = period.inventory_transactions.select_related("batch", "material")
    txns = _date_filter(request, txns, "txn_date", "txn_date")
    txns = _text_filters(
        request,
        txns,
        {
            "document_no": "document_no",
            "business_type": "business_type",
            "material": "material_code",
            "warehouse": "warehouse",
            "order_no": "order_no",
        },
    )
    txns = _choice_filter(request, txns, "direction", "direction")
    party = str(request.GET.get("party", "")).strip()
    if party:
        txns = txns.filter(
            Q(supplier_code__icontains=party)
            | Q(supplier_name__icontains=party)
            | Q(customer_code__icontains=party)
            | Q(customer_name__icontains=party)
        )
    txns = txns.order_by("txn_date", "document_no", "id")
    return render(
        request,
        "settlement/inventory_ledger.html",
        {
            "period": period,
            "page": _page(request, txns),
            "directions": InventoryTransaction.Direction.choices,
        },
    )


def _inventory_match_queryset(request, period):
    """库存流水匹配页的取数（页面与「导出Excel」共用同一份）。

    页面上的筛选栏已删除（用户 2026-09-21），但老链接/书签里带的参数照旧生效：
    业务日期 txn_date / 库存单据号 document_no / 业务类型 business_type / 方向 direction /
    仓库 warehouse / 物料编码 material / 供应商·客户 party / 订单号 order_no /
    处理方式 match_mode / 匹配状态 match_status。
    """
    txns = period.inventory_transactions.prefetch_related(
        "allocations__contract_line",
        "allocations__purchase_invoice_line",
        "allocations__sales_invoice_line",
    )
    txns = _text_filters(
        request,
        txns,
        {
            "document_no": "document_no",
            "business_type": "business_type",
            "material": "material_code",
            "order_no": "order_no",
            "warehouse": "warehouse",
        },
    )
    txns = _date_filter(request, txns, "txn_date", "txn_date")
    txns = _choice_filter(request, txns, "direction", "direction")
    txns = _choice_filter(request, txns, "match_mode", "match_mode")
    txns = _choice_filter(request, txns, "match_status", "match_status")
    party = str(request.GET.get("party", "")).strip()
    if party:
        txns = txns.filter(
            Q(supplier_code__icontains=party)
            | Q(supplier_name__icontains=party)
            | Q(customer_code__icontains=party)
            | Q(customer_name__icontains=party)
        )
    return txns


def _inventory_match_month(value):
    """业务日期的年月，显示成 2026-7（用户 2026-09-21：格式就是业务日期的年月）。"""
    if value is None:
        return ""
    return "%d-%d" % (value.year, value.month)


def _inventory_match_allocation_text(txn):
    """「匹配到的单据」列的文字：合同 C2606…：1,234.000；采购发票 B…：12.000。"""
    parts = []
    for allocation in txn.allocations.all():
        if allocation.contract_line:
            label = "合同 %s" % (allocation.contract_line.contract_no or "")
        elif allocation.purchase_invoice_line:
            label = "采购发票 %s" % (allocation.purchase_invoice_line.invoice_no or "")
        elif allocation.sales_invoice_line:
            label = "销售发票 %s" % (allocation.sales_invoice_line.invoice_no or "")
        else:
            label = "—"
        parts.append("%s：%s" % (label, allocation.matched_quantity))
    return "；".join(parts)


@login_required
def inventory_matches(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    page = _page(request, _inventory_match_queryset(request, period))
    for txn in page.object_list:
        txn.match_category = _inventory_match_category(txn)
        txn.period_label = _inventory_match_month(txn.txn_date)
    return render(
        request,
        "settlement/inventory_matches.html",
        {
            "period": period,
            "page": page,
            "match_modes": MatchingMode.choices,
            "match_statuses": InventoryTransaction.MatchStatus.choices,
            "directions": InventoryTransaction.Direction.choices,
        },
    )


SOURCE_DOCUMENTS = {
    "purchase_invoice": {
        "model": PurchaseInvoiceLine,
        "kind": MatchAllocation.Kind.PURCHASE_INVOICE,
        "title": "采购发票",
        "list_view": "purchase_matches",
    },
    "purchase_contract": {
        # 用户 2026-09-19：合同也能「打开修改」处理方式（与采购发票页一致）
        "model": PurchaseContractLine,
        "kind": MatchAllocation.Kind.PURCHASE_CONTRACT,
        "title": "采购合同",
        "list_view": "purchase_contracts",
    },
    "sales_invoice": {
        "model": SalesInvoiceLine,
        "kind": MatchAllocation.Kind.SALES_INVOICE,
        "title": "销售发票",
        "list_view": "sales_matches",
    },
}


def _source_candidates(period, source_type, source):
    if source_type in {"purchase_invoice", "sales_invoice"}:
        txns = (
            InventoryTransaction.objects.filter(
                period__month__lte=period.month,
                material_code=source.material_code,
            ).exclude(match_mode=MatchingMode.IGNORED)
            .select_related("period")
            .prefetch_related("allocations")
        )
        candidates = []
        for txn in txns:
            expected_kind = (
                is_purchase_receipt(txn)
                if source_type == "purchase_invoice"
                else is_sales_issue(txn)
            )
            if not expected_kind:
                continue
            matched = sum(
                (
                    allocation.matched_quantity
                    for allocation in txn.allocations.all()
                    if allocation.kind
                    == (
                        MatchAllocation.Kind.PURCHASE_INVOICE
                        if source_type == "purchase_invoice"
                        else MatchAllocation.Kind.SALES_INVOICE
                    )
                ),
                0,
            )
            txn.available_match_quantity = max(
                txn.quantity - matched,
                0,
            )
            if txn.available_match_quantity > 0:
                candidates.append(txn)
        return candidates
    if source_type == "purchase_contract":
        # 合同的手工匹配对象：**本期**仍可被合同占用的采购入库流水（自动匹配也是按本期分配的）
        txns = (
            InventoryTransaction.objects.filter(
                period=period,
                material_code=source.material_code,
            )
            .exclude(match_mode=MatchingMode.IGNORED)
            .select_related("period")
            .prefetch_related("allocations")
        )
        candidates = []
        for txn in txns:
            if not is_purchase_receipt(txn):
                continue
            matched = sum(
                (
                    allocation.matched_quantity
                    for allocation in txn.allocations.all()
                    if allocation.kind == MatchAllocation.Kind.PURCHASE_CONTRACT
                ),
                0,
            )
            txn.available_match_quantity = max(txn.quantity - matched, 0)
            if txn.available_match_quantity > 0:
                candidates.append(txn)
        return candidates
    return []


def _source_opening_candidates(period, source_type, source):
    if source_type == "purchase_invoice":
        rows = OpeningPurchaseAccrual.objects.filter(
            period__month__lte=period.month,
            material_code=source.material_code,
        ).exclude(match_mode=MatchingMode.IGNORED)
        kind = MatchAllocation.Kind.PURCHASE_INVOICE
    elif source_type == "sales_invoice":
        rows = OpeningGoodsIssued.objects.filter(
            period__month__lte=period.month,
            material_code=source.material_code,
        ).exclude(match_mode=MatchingMode.IGNORED)
        kind = MatchAllocation.Kind.SALES_INVOICE
    else:
        return []
    rows = rows.prefetch_related("allocations")
    candidates = []
    for row in rows:
        matched = sum(
            (
                allocation.matched_quantity
                for allocation in row.allocations.all()
                if allocation.kind == kind
                and allocation.period.month <= period.month
            ),
            0,
        )
        row.available_match_quantity = max(
            row.opening_quantity - matched,
            0,
        )
        if row.available_match_quantity > 0:
            candidates.append(row)
    return candidates


@login_required
def source_document_detail(request, period_id, source_type, source_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    config = SOURCE_DOCUMENTS.get(source_type)
    if not config:
        return HttpResponseBadRequest("不支持的单据类型。")
    source = get_object_or_404(
        config["model"], id=source_id, period=period
    )
    if request.method == "POST":
        action = request.POST.get("action")
        try:
            if action == "set_mode":
                form = MatchModeForm(request.POST)
                if not form.is_valid():
                    raise ValueError("处理方式填写不完整。")
                set_source_match_mode(
                    period,
                    source,
                    form.cleaned_data["match_mode"],
                    form.cleaned_data["ignore_reason"],
                )
                messages.success(request, "单据处理方式已更新。")
            elif action == "add_match":
                form = SourceDocumentMatchForm(request.POST)
                if not form.is_valid():
                    raise ValueError("请选择库存流水并填写匹配数量。")
                txn_filters = {
                    "id": form.cleaned_data["transaction_id"],
                }
                if source_type in {"purchase_invoice", "sales_invoice"}:
                    txn_filters["period__month__lte"] = period.month
                else:
                    txn_filters["period"] = period
                txn = get_object_or_404(
                    InventoryTransaction,
                    **txn_filters,
                )
                kind = config["kind"]
                manual_match(
                    period,
                    kind,
                    txn.id,
                    source.id,
                    form.cleaned_data["matched_quantity"],
                )
                messages.success(request, "手工匹配已保存。")
            elif action == "add_opening_match":
                if source_type not in {
                    "purchase_invoice",
                    "sales_invoice",
                }:
                    raise ValueError("该单据不能核销期初台账。")
                form = OpeningDocumentMatchForm(request.POST)
                if not form.is_valid():
                    raise ValueError("请选择期初明细并填写匹配数量。")
                manual_match_opening(
                    period,
                    config["kind"],
                    form.cleaned_data["opening_id"],
                    source.id,
                    form.cleaned_data["matched_quantity"],
                )
                messages.success(request, "期初台账核销关系已保存。")
            elif action == "remove_match":
                remove_manual_allocation(
                    period, int(request.POST.get("allocation_id", 0))
                )
                messages.success(request, "匹配关系已删除。")
            else:
                raise ValueError("不支持的操作。")
            reset_workshop_settlements(period)
        except Exception as exc:
            messages.error(request, str(exc))
        return redirect(
            "source_document_detail",
            period_id=period.id,
            source_type=source_type,
            source_id=source.id,
        )

    allocations = MatchAllocation.objects.filter(period=period)
    if source_type == "purchase_invoice":
        allocations = allocations.filter(purchase_invoice_line=source)
    elif source_type == "purchase_contract":
        allocations = allocations.filter(contract_line=source)
    elif source_type == "sales_invoice":
        allocations = allocations.filter(sales_invoice_line=source)
    allocations = allocations.select_related(
        "inventory_transaction__period",
        "opening_purchase_accrual",
        "opening_goods_issued",
    )
    candidates = _source_candidates(period, source_type, source)
    opening_candidates = _source_opening_candidates(
        period, source_type, source
    )
    for txn in candidates:
        txn.match_category = _inventory_match_category(txn)
    return render(
        request,
        "settlement/source_document_detail.html",
        {
            "period": period,
            "source": source,
            "source_type": source_type,
            "source_title": config["title"],
            "list_view": config["list_view"],
            "allocations": allocations,
            "candidates": candidates[:300],
            "opening_candidates": opening_candidates[:300],
            "match_modes": MatchingMode.choices,
        },
    )


def _transaction_source_groups(period, txn):
    groups = []
    if is_purchase_receipt(txn):
        groups.append(
            {
                "kind": MatchAllocation.Kind.PURCHASE_CONTRACT,
                "label": "采购合同",
                "items": period.purchase_contracts.filter(
                    material_code=txn.material_code
                ).exclude(match_mode=MatchingMode.IGNORED)[:300],
            }
        )
        groups.append(
            {
                "kind": MatchAllocation.Kind.PURCHASE_INVOICE,
                "label": "采购发票",
                "items": period.purchase_invoices.filter(
                    material_code=txn.material_code
                ).exclude(match_mode=MatchingMode.IGNORED)[:300],
            }
        )
    # 销售出库不再参与匹配（用户 2026-09-21）：不给「销售发票」可匹配组，页面上也就没有手工匹配入口
    return groups


@login_required
def inventory_match_detail(request, period_id, transaction_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    txn = get_object_or_404(
        InventoryTransaction, id=transaction_id, period=period
    )
    if request.method == "POST":
        action = request.POST.get("action")
        try:
            if action == "set_mode":
                form = MatchModeForm(request.POST)
                if not form.is_valid():
                    raise ValueError("处理方式填写不完整。")
                set_transaction_match_mode(
                    period,
                    txn,
                    form.cleaned_data["match_mode"],
                    form.cleaned_data["ignore_reason"],
                )
                messages.success(request, "库存流水处理方式已更新。")
            elif action == "add_match":
                form = InventoryDocumentMatchForm(request.POST)
                if not form.is_valid():
                    raise ValueError("目标单据和匹配数量填写不完整。")
                allowed_kinds = {
                    group["kind"]
                    for group in _transaction_source_groups(period, txn)
                }
                kind = form.cleaned_data["kind"]
                if kind not in allowed_kinds:
                    raise ValueError("匹配类型与该库存流水不一致。")
                if txn.match_mode != MatchingMode.MANUAL:
                    set_transaction_match_mode(
                        period, txn, MatchingMode.MANUAL
                    )
                    txn.refresh_from_db()
                manual_match(
                    period,
                    kind,
                    txn.id,
                    form.cleaned_data["source_id"],
                    form.cleaned_data["matched_quantity"],
                    set_source_manual=False,
                )
                messages.success(request, "手工匹配已保存。")
            elif action == "remove_match":
                remove_manual_allocation(
                    period, int(request.POST.get("allocation_id", 0))
                )
                messages.success(request, "匹配关系已删除。")
            else:
                raise ValueError("不支持的操作。")
            reset_workshop_settlements(period)
        except Exception as exc:
            messages.error(request, str(exc))
        return redirect(
            "inventory_match_detail",
            period_id=period.id,
            transaction_id=txn.id,
        )
    txn.match_category = _inventory_match_category(txn)
    return render(
        request,
        "settlement/inventory_match_detail.html",
        {
            "period": period,
            "txn": txn,
            "allocations": txn.allocations.select_related(
                "contract_line",
                "purchase_invoice_line__period",
                "sales_invoice_line__period",
            ),
            "source_groups": _transaction_source_groups(period, txn),
            "match_modes": MatchingMode.choices,
        },
    )


@login_required
@require_POST
def manual_match_view(request, period_id, kind):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    form = ManualMatchForm(request.POST)
    allowed = {value for value, _ in MatchAllocation.Kind.choices}
    if kind not in allowed:
        return HttpResponseBadRequest("不支持的匹配类型。")
    if form.is_valid():
        try:
            manual_match(
                period,
                kind,
                form.cleaned_data["transaction_id"],
                form.cleaned_data["source_id"],
                form.cleaned_data["matched_quantity"],
            )
            reset_workshop_settlements(period)
            messages.success(request, "人工匹配已保存，并自动补齐剩余匹配。")
        except Exception as exc:
            messages.error(request, str(exc))
    else:
        messages.error(request, "人工匹配信息不完整。")
    target = (
        "sales_matches"
        if kind == MatchAllocation.Kind.SALES_INVOICE
        else "purchase_matches"
    )
    return redirect(target, period_id=period.id)


#: 物料分类展示顺序（与单价表一致；其余分类按名称排在后面）
INVENTORY_CATEGORY_ORDER = (
    "原料", "染色纤维", "纤维", "熟条", "浆粞",
    "纱线", "面料", "坯布", "产成品",
)


def _inventory_category_sort_key(label):
    if label in INVENTORY_CATEGORY_ORDER:
        return (0, INVENTORY_CATEGORY_ORDER.index(label), "")
    if label == "未分类":
        return (2, 0, "")
    return (1, 0, label or "")


def inventory_category_map(period):
    """物料编码 -> 物料分类。

    库存结转表本身不存分类，按优先级从下面几处取（先到先得）：
    ① 期初存货明细的分类名称（覆盖最全）；
    ② 期末存货明细；③ 本期库存流水的分类名称；④ 本期投料记录的物料分类。
    取不到的显示为「未分类」。
    """
    mapping = {}
    layers = (
        period.opening_inventory_balances.exclude(category_name="")
        .order_by("material_code", "id").values_list("material_code", "category_name"),
        period.ending_inventory_balances.exclude(category_name="")
        .order_by("material_code", "id").values_list("material_code", "category_name"),
        period.inventory_transactions.exclude(category_name="")
        .order_by("material_code", "id").values_list("material_code", "category_name"),
        period.material_issues.exclude(material_category="")
        .order_by("material_code", "id").values_list("material_code", "material_category"),
    )
    for pairs in layers:
        for code, category in pairs:
            if code and category:
                mapping.setdefault(code, category.strip())
    return mapping


def build_inventory_category_rows(period, category_by_code):
    """按「期间 × 物料分类」汇总各项的数量与金额（期初 / 本期入库 / 本期出库 / 期末）。"""
    zero = Decimal("0")
    buckets = {}
    period_label = period.month.strftime("%Y-%m")
    for item in period.inventory_valuations.all():
        category = category_by_code.get(item.material_code) or "未分类"
        bucket = buckets.setdefault(
            category,
            {
                "opening_quantity": zero, "opening_amount": zero,
                "receipt_quantity": zero, "receipt_amount": zero,
                "issue_quantity": zero, "issue_amount": zero,
                "ending_quantity": zero, "ending_amount": zero,
            },
        )
        for key in ("opening", "receipt", "issue", "ending"):
            bucket[f"{key}_quantity"] += Decimal(getattr(item, f"{key}_quantity") or 0)
            bucket[f"{key}_amount"] += Decimal(getattr(item, f"{key}_amount") or 0)

    rows = []
    totals = {
        "opening_quantity": zero, "opening_amount": zero,
        "receipt_quantity": zero, "receipt_amount": zero,
        "issue_quantity": zero, "issue_amount": zero,
        "ending_quantity": zero, "ending_amount": zero,
    }
    for category, bucket in buckets.items():
        row = {"period_label": period_label, "category": category}
        for key in ("opening", "receipt", "issue", "ending"):
            row[f"{key}_quantity"] = bucket[f"{key}_quantity"]
            row[f"{key}_amount"] = bucket[f"{key}_amount"]
            totals[f"{key}_quantity"] += bucket[f"{key}_quantity"]
            totals[f"{key}_amount"] += bucket[f"{key}_amount"]
        rows.append(row)
    rows.sort(key=lambda row: _inventory_category_sort_key(row["category"]))
    return rows, totals


def _inventory_valuation_periods(current, selected):
    """库存金额结转页要取哪些期间（用户 2026-09-21）：

    默认＝当前核算期间（只显示当月）；selected == "all" → 有数据的全部期间；
    selected 是某个月（2026-7 / 2026-07）→ 那个月。这样既能默认看当月，也能查历史。
    """
    if selected == "all":
        ids = (
            InventoryValuation.objects.values_list("period_id", flat=True).distinct()
        )
        return list(AccountingPeriod.objects.filter(id__in=ids).order_by("month"))
    month = _parse_month(selected)
    if month:
        matched = [
            item
            for item in AccountingPeriod.objects.order_by("month")
            if (item.month.year, item.month.month) == month
        ]
        if matched:
            return matched
    return [current]


def _inventory_period_options(current):
    """下拉里的期间候选：有库存金额结转数据的期间（本期已由「本期（YYYY-MM）」覆盖，不再重复列）。"""
    ids = set(
        InventoryValuation.objects.values_list("period_id", flat=True).distinct()
    )
    ids.add(current.id)
    current_label = current.month.strftime("%Y-%m")
    return [
        {"value": item.month.strftime("%Y-%m"), "label": item.month.strftime("%Y-%m")}
        for item in AccountingPeriod.objects.filter(id__in=ids).order_by("month")
        if item.month.strftime("%Y-%m") != current_label
    ]


def build_inventory_results_data(request, period):
    """库存金额结转页的取数（页面与两个导出共用同一份）。

    汇总（按物料分类）：筛选 s_period（期间，默认当月）/ s_category（物料分类）
    明细（逐物料）    ：筛选 d_period（期间，默认当月）/ d_material（物料编码，包含即可）
    """
    zero = Decimal("0")
    category_cache = {}

    def categories(target):
        if target.id not in category_cache:
            category_cache[target.id] = inventory_category_map(target)
        return category_cache[target.id]

    summary_filters = {
        "period": str(request.GET.get("s_period", "")).strip(),
        "category": str(request.GET.get("s_category", "")).strip(),
    }
    detail_filters = {
        "period": str(request.GET.get("d_period", "")).strip(),
        "material": str(request.GET.get("d_material", "")).strip(),
    }

    # —— 汇总：所选期间的「期间 × 物料分类」行 ——
    summary_periods = _inventory_valuation_periods(
        period, summary_filters["period"]
    )
    all_summary_rows = []
    for target in summary_periods:
        rows, _totals = build_inventory_category_rows(target, categories(target))
        all_summary_rows.extend(rows)
    category_options = sorted(
        {row["category"] for row in all_summary_rows}, key=_inventory_category_sort_key
    )
    summary_rows = [
        row
        for row in all_summary_rows
        if not summary_filters["category"]
        or row["category"] == summary_filters["category"]
    ]
    summary_rows.sort(
        key=lambda row: (row["period_label"], _inventory_category_sort_key(row["category"]))
    )
    summary_totals = {
        key: zero
        for key in (
            "opening_quantity", "opening_amount",
            "receipt_quantity", "receipt_amount",
            "issue_quantity", "issue_amount",
            "ending_quantity", "ending_amount",
        )
    }
    for row in summary_rows:
        for key in summary_totals:
            summary_totals[key] += row[key]

    # —— 明细：所选期间 ×（物料编码包含） ——
    detail_periods = _inventory_valuation_periods(period, detail_filters["period"])
    valuations = (
        InventoryValuation.objects.filter(period__in=detail_periods)
        .select_related("period")
        .order_by("period__month", "material_code")
    )
    if detail_filters["material"]:
        valuations = valuations.filter(
            material_code__icontains=detail_filters["material"]
        )
    for item in valuations:
        item.material_category = (
            categories(item.period).get(item.material_code) or "未分类"
        )
        item.period_label = item.period.month.strftime("%Y-%m")

    unclassified_count = sum(
        1 for item in valuations if item.material_category == "未分类"
    )
    return {
        "summary_rows": summary_rows,
        "summary_totals": summary_totals,
        "category_options": category_options,
        "period_options": _inventory_period_options(period),
        "detail_rows": valuations,
        "summary_filters": summary_filters,
        "detail_filters": detail_filters,
        "unclassified_count": unclassified_count,
        "current_month": period.month.strftime("%Y-%m"),
    }


@login_required
def inventory_results(request, period_id):
    """库存金额结转：按物料分类的汇总 + 逐物料明细（都能按期间查历史，默认当月）。

    用户 2026-09-21：两个表各自带「导出Excel」（放在表标题旁）；汇总加「期间/物料分类」筛选、
    明细加「期间/物料编码」筛选并补一列「期间」。
    """
    period = get_object_or_404(AccountingPeriod, id=period_id)
    data = build_inventory_results_data(request, period)
    page = _page(request, data["detail_rows"])
    return render(
        request,
        "settlement/inventory_results.html",
        {
            "period": period,
            "page": page,
            "category_rows": data["summary_rows"],
            "category_totals": data["summary_totals"],
            "category_options": data["category_options"],
            "period_options": data["period_options"],
            "summary_filters": data["summary_filters"],
            "detail_filters": data["detail_filters"],
            "unclassified_count": data["unclassified_count"],
            "current_month": data["current_month"],
        },
    )


@login_required
def work_order_report(request, period_id):
    """生产工单报表：工单主档台账（生产工单口径）。"""
    period = get_object_or_404(AccountingPeriod, id=period_id)
    orders = period.work_orders.all()
    # 页面上的筛选栏已按用户 2026-09-21 要求删除（工单编号 / 车间·部门 / 状态 + 筛选 / 清空）；
    # 老链接 / 书签里带的这几个参数仍然生效，导出也会把它们带过去。
    orders = _text_filters(
        request,
        orders,
        {
            "order_no": "order_no",
            "workshop": "production_dept_code",
            "product": "identifier",
        },
    )
    status = str(request.GET.get("status", "")).strip()
    if status:
        orders = orders.filter(status__icontains=status)
    return render(
        request,
        "settlement/work_order_report.html",
        {
            "period": period,
            # 列＝上传文件的 25 列（列名/顺序一致），这里只用于页头提示行数
            "columns": WORK_ORDER_COLUMNS,
            "page": _page(
                request,
                orders.order_by("order_no"),
            ),
        },
    )


def build_material_issue_rows(period, order_no="", workshop="", material=""):
    """投料记录明细行（与导出模板字段一致）：每条投料记录一行，含全部模板字段 + 单价 + 金额。
    单价 = 月末加权价（inventory_valuations）→ 历史单价（receipt_price）兜底；
    无单价时 unit_price/amount 为 None（页面/导出提示"无单价"）。"""
    prices = {
        item.material_code: Decimal(item.weighted_average_price or 0)
        for item in period.inventory_valuations.all()
    }
    for hp in period.historical_prices.all():
        if hp.issue_price is not None:
            prices.setdefault(hp.material_code, Decimal(hp.issue_price))
        elif hp.receipt_price is not None:
            prices.setdefault(hp.material_code, Decimal(hp.receipt_price))
    wo_workshop = {
        wo.order_no: wo.production_dept_code for wo in period.work_orders.all()
    }
    issues = period.material_issues.all()
    if order_no:
        issues = issues.filter(work_order_no__icontains=order_no)
    if material:
        issues = issues.filter(material_code__icontains=material)
    rows = []
    for issue in issues:
        if not issue.work_order_no or not issue.material_code:
            continue
        ws = wo_workshop.get(issue.work_order_no, "") or ""
        if workshop and workshop not in ws:
            continue
        price = prices.get(issue.material_code) or None
        qty = Decimal(issue.actual_issue_quantity or 0)
        unit_price = (
            price.quantize(Decimal("0.0001"))
            if price is not None else None
        )
        amount = (
            (price * qty).quantize(Decimal("0.01"))
            if price is not None and qty else None
        )
        rows.append(
            {
                "task_no": issue.task_no,
                "operation_no": issue.operation_no,
                "operation_name": issue.operation_name,
                "material_code": issue.material_code,
                "material_name": issue.material_name,
                "material_category": issue.material_category,
                "material_attribute": issue.material_attribute,
                "material_spec": issue.material_spec,
                "issue_quantity": issue.issue_quantity,
                "returned_quantity": issue.returned_quantity,
                "actual_issue_quantity": qty,
                "issue_uom": issue.issue_uom,
                "auxiliary_quantity": issue.auxiliary_quantity,
                "auxiliary_uom": issue.auxiliary_uom,
                "identifier": issue.identifier,
                "batch_no": issue.batch_no,
                "quality_status": issue.quality_status,
                "issue_location": issue.issue_location,
                "replaced_material_code": issue.replaced_material_code,
                "replaced_material_name": issue.replaced_material_name,
                "issue_time": issue.issue_time,
                "issuer": issue.issuer,
                "executor": issue.executor,
                "work_order_no": issue.work_order_no,
                "main_output_code": issue.main_output_code,
                "main_output_name": issue.main_output_name,
                "sales_order_no": issue.sales_order_no,
                "record_no": issue.record_no,
                "workshop": ws,
                "unit_price": unit_price,
                "amount": amount,
            }
        )
    rows.sort(
        key=lambda r: (r["workshop"], r["work_order_no"], r["material_code"])
    )
    no_price = [r for r in rows if r["unit_price"] is None]
    no_price_materials = {}
    for r in no_price:
        no_price_materials.setdefault(r["material_code"], r["material_name"] or "")
    totals = {
        "issue_quantity": sum((r["actual_issue_quantity"] for r in rows), Decimal("0")),
        "amount": sum((r["amount"] for r in rows if r["amount"] is not None), Decimal("0")),
        "no_price_count": len(no_price),
        "no_price_materials": no_price_materials,
    }
    return rows, totals


def build_material_issue_summary(rows, period, extra_rows=()):
    """工单投料汇总：按「投料期间 × 投料车间 × 物料分类」汇总数量、金额与无单价行数。

    rows：工单投料台账明细行（build_material_issue_rows 的返回值）。
    extra_rows：额外并入的行（目前是浆粕车间的领料出库流水，按备注认领）、
                每条需要 month / workshop / category / quantity / amount / no_price 字段。
    金额＝数量 × 单价（月末加权价，历史单价兜底），无单价行为 0 并单独计数。
    """
    from settlement.services.imports import WORKSHOP_ORDER as _WS_ORDER

    groups = {}

    def _add(month, workshop, category, quantity, amount, no_price):
        key = (month, workshop, category)
        item = groups.setdefault(
            key,
            {
                "month": month,
                "workshop": workshop,
                "workshop_label": WORKSHOP_LABELS.get(workshop, workshop or "未归属"),
                "category": category,
                "quantity": Decimal("0"),
                "amount": Decimal("0"),
                "no_price": 0,
                "rows": 0,
            },
        )
        item["quantity"] += quantity
        item["amount"] += amount
        item["rows"] += 1
        if no_price:
            item["no_price"] += 1

    for row in rows:
        when = row.get("issue_time")
        month = when.strftime("%Y-%m") if when else period.month.strftime("%Y-%m")
        _add(
            month,
            row.get("workshop") or "",
            (row.get("material_category") or "").strip() or "未分类",
            row.get("actual_issue_quantity") or Decimal("0"),
            row.get("amount") or Decimal("0"),
            row.get("unit_price") is None,
        )

    extra_count = 0
    extra_no_price = 0
    for row in extra_rows:
        extra_count += 1
        if row.get("no_price"):
            extra_no_price += 1
        _add(
            row.get("month") or period.month.strftime("%Y-%m"),
            row.get("workshop") or "",
            (row.get("category") or "").strip() or "未分类",
            row.get("quantity") or Decimal("0"),
            row.get("amount") or Decimal("0"),
            row.get("no_price"),
        )

    order_index = {ws: index for index, ws in enumerate(_WS_ORDER)}
    summary = sorted(
        groups.values(),
        key=lambda item: (
            item["month"],
            order_index.get(item["workshop"], 99),
            item["category"],
        ),
    )
    totals = {
        "quantity": sum((item["quantity"] for item in summary), Decimal("0")),
        "amount": sum((item["amount"] for item in summary), Decimal("0")),
        "no_price": sum(item["no_price"] for item in summary),
        "rows": sum(item["rows"] for item in summary),
        "ledger_rows": len(rows),
        "extra_rows": extra_count,
        "extra_no_price": extra_no_price,
    }
    return summary, totals


#: 「车间费用」页面录入列（顺序＝用户模板列顺序：人工费、折旧费、水电费、蒸汽费、燃气费、染料助剂费、机物料费、加工费）
WORKSHOP_COST_COLUMNS = (
    ("wages", "人工费"),
    ("depreciation", "折旧费"),
    ("water_power", "水电费"),
    ("steam", "蒸汽费"),
    ("natural_gas", "燃气费"),
    ("dye_auxiliary", "染料助剂费"),
    ("consumables", "机物料费"),
    ("processing_fee", "加工费"),
)

#: 页面录入生成的批次（每个期间一条，便于和上传批次区分；sha256 用固定哨兵值占位）
MANUAL_WORKSHOP_COST_NAME = "（车间费用·页面录入）"
MANUAL_WORKSHOP_COST_SHA = "manual-workshop-cost-entry"
ZERO_CNY = Decimal("0")


def _cost_money(value):
    return (Decimal(value or ZERO_CNY)).quantize(Decimal("0.01"))


def _manual_workshop_cost_batch(period, user=None):
    """页面手工录入用的批次（每个期间一条）。"""
    batch = ImportBatch.objects.filter(
        period=period,
        kind=ImportBatch.Kind.WORKSHOP_COST,
        sha256=MANUAL_WORKSHOP_COST_SHA,
    ).first()
    if batch is None:
        batch = ImportBatch.objects.create(
            period=period,
            kind=ImportBatch.Kind.WORKSHOP_COST,
            file="",
            original_name=MANUAL_WORKSHOP_COST_NAME,
            sha256=MANUAL_WORKSHOP_COST_SHA,
            status=ImportBatch.Status.SUCCESS,
            uploaded_by=user,
        )
    return batch


def _workshop_cost_periods():
    """所有核算期间（升序）：车间费用表按期间排在一起，7月/8月/9月…都能填。"""
    return list(AccountingPeriod.objects.order_by("month"))


def _workshop_cost_rows(period):
    """某期间的行：车间顺序＝车间结转顺序（费用按这个顺序参与分配），再补上费用池里出现的车间。

    输入框名带上期间 id（ws_{期间id}_{车间}_{字段}），便于在同一张表里录入多个月份。
    """
    pools = {pool.workshop: pool for pool in period.workshop_cost_pools.all()}
    order = list(
        period.workshop_settlements.order_by("sequence", "workshop").values_list(
            "workshop", flat=True
        )
    )
    for code in DEFAULT_WORKSHOPS:
        if code not in order:
            order.append(code)
    for code in pools:
        if code not in order:
            order.append(code)
    rows = []
    for index, code in enumerate(order, start=1):
        pool = pools.get(code)
        cells = []
        total = ZERO_CNY
        for field, label in WORKSHOP_COST_COLUMNS:
            value = _cost_money(getattr(pool, field, ZERO_CNY) if pool else ZERO_CNY)
            total += value
            cells.append(
                {
                    "field": field,
                    "label": label,
                    "name": "ws_%s_%s_%s" % (period.id, code, field),
                    "value": value,
                    # 输入框里也带千分位（用户 2026-09-21）；保存时前后端都会去掉逗号
                    "input_value": "{:,.2f}".format(value),
                }
            )
        rows.append(
            {
                "period_id": period.id,
                "sequence": index,
                "workshop": code,
                "label": WORKSHOP_LABELS.get(code, code),
                "cells": cells,
                "remark_name": "ws_%s_%s_remark" % (period.id, code),
                "remark_value": (pool.remark or "") if pool else "",
                "has_pool": pool is not None,
                "from_upload": bool(
                    pool is not None
                    and pool.batch_id
                    and pool.batch.original_name != MANUAL_WORKSHOP_COST_NAME
                ),
                "total": _cost_money(total),
                "row_total_key": "%s_%s" % (period.id, code),
            }
        )
    return rows


def _workshop_cost_totals(rows):
    """列合计与总计（页面首屏就显示；前端输入时实时重算）。"""
    column_totals = []
    for field, label in WORKSHOP_COST_COLUMNS:
        amount = ZERO_CNY
        for row in rows:
            for cell in row["cells"]:
                if cell["field"] == field:
                    amount += Decimal(cell["value"] or ZERO_CNY)
        amount = _cost_money(amount)
        column_totals.append(
            {
                "field": field,
                "label": label,
                "amount": amount,
                "text": "{:,.2f}".format(amount),
            }
        )
    grand = sum((Decimal(row["total"] or ZERO_CNY) for row in rows), ZERO_CNY)
    return column_totals, _cost_money(grand)


def _summary_amount(value):
    """汇总用：把单元格里的值安全转成金额（输入出错回填时可能是文本）。

    ⚠ 页面上的金额文本带千分位（"100,723.08"），必须先去掉逗号再解析——
    否则 Decimal(...) 直接抛 InvalidOperation，总合计会变成 0（2026-09-21 修）。
    """
    text = str(value or "").replace(",", "").replace("，", "").strip()
    if not text:
        return ZERO_CNY
    try:
        return _cost_money(Decimal(text))
    except (InvalidOperation, ValueError):
        return ZERO_CNY


def _workshop_cost_summary(groups, workshop=""):
    """累计汇总（按车间）：列与「各车间费用」录入表完全一致，但数据是所有期间累计。

    每个车间一行：人工费/折旧费/…/加工费 都是各期间相加（8月就是7月+8月，以后逐月累加），
    末列「合计」＝该车间各科目累计之和；期间列显示累计区间（如 2026-07~2026-08）；
    最后一行给出各科目的总合计。

    workshop 非空时只保留该车间（用户 2026-09-27：汇总表自己的「车间」筛选），
    总合计行也跟着只算筛出来的行。
    """
    if not groups:
        return [], None
    labels = [group["period_label"] for group in groups]
    period_range = labels[0] if len(labels) == 1 else "%s~%s" % (labels[0], labels[-1])

    order = []
    seen = set()
    for group in groups:
        for row in group["rows"]:
            code = row["workshop"]
            if code in seen:
                continue
            seen.add(code)
            order.append((row["sequence"], code, row["label"]))
    order.sort()
    if workshop:
        order = [item for item in order if item[1] == workshop]

    summary_rows = []
    for sequence, code, label in order:
        cells = []
        total = ZERO_CNY
        for field, column_label in WORKSHOP_COST_COLUMNS:
            amount = ZERO_CNY
            for group in groups:
                for row in group["rows"]:
                    if row["workshop"] != code:
                        continue
                    for cell in row["cells"]:
                        if cell["field"] == field:
                            amount += _summary_amount(cell["value"])
            amount = _cost_money(amount)
            total += amount
            cells.append(
                {"field": field, "label": column_label,
                 "amount": amount, "text": "{:,.2f}".format(amount)}
            )
        summary_rows.append(
            {
                "sequence": sequence,
                "workshop": code,
                "label": label,
                "period_range": period_range,
                "cells": cells,
                "total_text": "{:,.2f}".format(_cost_money(total)),
            }
        )

    footer_cells = []
    grand_total = ZERO_CNY
    for field, column_label in WORKSHOP_COST_COLUMNS:
        amount = ZERO_CNY
        for row in summary_rows:
            for cell in row["cells"]:
                if cell["field"] == field:
                    # 优先用行里存的数值（文本带千分位，直接解析会算成 0）
                    amount += _summary_amount(cell.get("amount", cell["text"]))
        amount = _cost_money(amount)
        grand_total += amount
        footer_cells.append(
            {"field": field, "label": column_label, "text": "{:,.2f}".format(amount)}
        )
    footer = {
        "period_range": period_range,
        "cells": footer_cells,
        "total_text": "{:,.2f}".format(_cost_money(grand_total)),
    }
    return summary_rows, footer


def _workshop_cost_groups(current_period):
    """页面数据：按期间分组（每期间＝该期间车间行 + 该期间列合计），并给出跨期间汇总。"""
    periods = _workshop_cost_periods()
    if all(item.id != current_period.id for item in periods):
        periods.append(current_period)
        periods.sort(key=lambda item: item.month)
    groups = []
    for item in periods:
        rows = _workshop_cost_rows(item)
        column_totals, subtotal = _workshop_cost_totals(rows)
        groups.append(
            {
                "period": item,
                "period_id": item.id,
                "period_label": item.month.strftime("%Y-%m"),
                "locked": item.status == AccountingPeriod.Status.LOCKED,
                "rows": rows,
                "column_totals": column_totals,
                "subtotal": subtotal,
            }
        )
    return groups


def _workshop_cost_write(period, groups_rows, posted_period, user):
    """把某个期间 POST 上来的值写进该期间的费用池，返回是否有变化。

    只有**真的改了数字或备注**才写库、才把批次标成「页面录入」；
    完全没改的行原样不动（这样只打开页面点保存不会重置结算、也不会丢掉「原上传」标记）。
    """
    changed = False
    batch = _manual_workshop_cost_batch(period, user)
    with transaction.atomic():
        for row in groups_rows:
            entry = posted_period.get(row["workshop"])
            if not entry:
                continue
            data, remark = entry
            pool = WorkshopCostPool.objects.filter(
                period=period, workshop=row["workshop"]
            ).first()
            if pool is None:
                # 该车间本来没有费用记录：全 0（且备注为空）就不建记录，避免"打开页面点保存"也产生改动
                if not remark and not any(
                    data[field] for field, _label in WORKSHOP_COST_COLUMNS
                ):
                    continue
                pool = WorkshopCostPool(
                    period=period, batch=batch, workshop=row["workshop"]
                )
                for field, _label in WORKSHOP_COST_COLUMNS:
                    setattr(pool, field, data[field])
                pool.remark = remark
                pool.save()
                changed = True
                continue
            values_changed = False
            for field, _label in WORKSHOP_COST_COLUMNS:
                if _cost_money(getattr(pool, field, ZERO_CNY)) != data[field]:
                    values_changed = True
                    setattr(pool, field, data[field])
            remark_changed = (pool.remark or "") != remark
            if not (values_changed or remark_changed):
                continue
            if remark_changed:
                pool.remark = remark
            pool.batch = batch        # 改过的行记为「页面录入」
            pool.save()
            changed = True
    return changed


@login_required
@never_cache
def workshop_costs(request, period_id):
    """车间费用：直接在页面表格里填各车间费用（不再需要上传文件），按期间分块列出所有月份。

    列＝期间｜车间｜人工费｜折旧费｜水电费｜蒸汽费｜燃气费｜染料助剂费｜机物料费｜加工费｜合计｜备注
    （与用户模板顺序一致）。行顺序＝期间 × 车间结转顺序，费用就按这个顺序参与结转/分配。
    保存后相关期间的结算结果重置，需按顺序重新结转。已锁账的期间只读。
    """
    period = get_object_or_404(AccountingPeriod, id=period_id)
    groups = _workshop_cost_groups(period)
    # 录入表只显示「当前这个期间」要填的数据（7 月填 7 月、8 月填 8 月）；
    # 以前月份的数据在下面的「车间费用汇总表（按车间）」里累计体现。
    entry_group = next(
        (group for group in groups if group["period_id"] == period.id), None
    )
    if entry_group is None:
        rows = _workshop_cost_rows(period)
        column_totals, subtotal = _workshop_cost_totals(rows)
        entry_group = {
            "period": period,
            "period_id": period.id,
            "period_label": period.month.strftime("%Y-%m"),
            "locked": period.status == AccountingPeriod.Status.LOCKED,
            "rows": rows,
            "column_totals": column_totals,
            "subtotal": subtotal,
        }
    current_locked = period.status == AccountingPeriod.Status.LOCKED
    can_save = not current_locked

    if request.method == "POST" and request.POST.get("action") == "copy_previous":
        # 用户 2026-10-14：车间费用每月都要重录，加「照抄上月」降低重复劳动。
        if current_locked:
            messages.error(request, "该期间已锁账，不能修改车间费用。")
        else:
            previous = (
                AccountingPeriod.objects.filter(month__lt=period.month)
                .order_by("-month")
                .first()
            )
            src = (
                list(WorkshopCostPool.objects.filter(period=previous))
                if previous else []
            )
            if not src:
                messages.warning(
                    request,
                    f"上一个期间（{previous.month:%Y-%m}）没有车间费用可照抄。"
                    if previous else "没有上一个期间。",
                )
            else:
                fields = ["water_power", "steam", "natural_gas", "depreciation",
                          "consumables", "dye_auxiliary", "wages", "processing_fee"]
                copied = 0
                batch = _manual_workshop_cost_batch(period, request.user)
                for row in src:
                    target, _created = WorkshopCostPool.objects.get_or_create(
                        period=period, workshop=row.workshop,
                        defaults={"batch": batch},
                    )
                    if target.batch_id is None:
                        target.batch = batch
                    for name in fields:
                        setattr(target, name, getattr(row, name))
                    target.remark = row.remark
                    target.save()
                    copied += 1
                messages.success(
                    request,
                    f"已照抄 {previous.month:%Y-%m} 的车间费用（{copied} 个车间），"
                    "请按本月实际修改后点「保存车间费用」。",
                )
        return redirect("workshop_costs", period_id=period.id)

    if request.method == "POST":
        if current_locked:
            messages.error(request, "该期间已锁账，不能修改车间费用。")
            return redirect("workshop_costs", period_id=period.id)
        errors = []
        parsed = {}
        for row in entry_group["rows"]:
            data = {}
            posted_row = {}
            for cell in row["cells"]:
                raw = str(request.POST.get(cell["name"], "") or "").strip()
                raw = raw.replace(",", "").replace("，", "")
                posted_row[cell["field"]] = raw
                if raw == "":
                    data[cell["field"]] = ZERO_CNY
                    continue
                try:
                    value = Decimal(raw)
                except (InvalidOperation, ValueError):
                    errors.append(
                        "%s %s 的「%s」不是数字：%s"
                        % (entry_group["period_label"], row["label"], cell["label"], raw)
                    )
                    continue
                if value < ZERO_CNY:
                    errors.append(
                        "%s %s 的「%s」不能为负数：%s"
                        % (entry_group["period_label"], row["label"], cell["label"], raw)
                    )
                    continue
                data[cell["field"]] = _cost_money(value)
            remark = str(request.POST.get(row["remark_name"], "") or "").strip()[:500]
            posted_row["remark"] = remark
            parsed[row["workshop"]] = (data, remark)
            entry_group.setdefault("posted", {})[row["workshop"]] = posted_row

        if errors:
            for item in errors:
                messages.error(request, item)
            # 回填用户刚填的内容，避免重打
            posted_group = entry_group.get("posted", {})
            for row in entry_group["rows"]:
                posted_row = posted_group.get(row["workshop"])
                if not posted_row:
                    continue
                total = ZERO_CNY
                for cell in row["cells"]:
                    raw_value = posted_row.get(cell["field"], "")
                    try:
                        cell["value"] = _cost_money(
                            Decimal(raw_value) if raw_value else ZERO_CNY
                        )
                        # 能解析成数字就带千分位显示
                        cell["input_value"] = "{:,.2f}".format(cell["value"])
                    except (InvalidOperation, ValueError):
                        # 填的不是数字：值按 0 算，但把用户填的内容原样留着（别吃掉）
                        cell["value"] = ZERO_CNY
                        cell["input_value"] = raw_value
                    total += cell["value"]
                row["remark_value"] = posted_row.get("remark", "")
                row["total"] = _cost_money(total)
            entry_group["column_totals"], entry_group["subtotal"] = (
                _workshop_cost_totals(entry_group["rows"])
            )
        else:
            changed = _workshop_cost_write(
                period, entry_group["rows"], parsed, request.user
            )
            if changed:
                from settlement.services.imports import clear_calculation_outputs

                clear_calculation_outputs(period)
                messages.success(
                    request,
                    "%s 的车间费用已保存，费用按车间顺序参与分配。"
                    "该期间结算结果已重置，请到「车间结转」按顺序重新结转。"
                    % entry_group["period_label"],
                )
            else:
                messages.success(request, "车间费用没有变化，未做改动。")
            return redirect("workshop_costs", period_id=period.id)

    # 兜底（用户 2026-10-14）：各期间的金额**直接从费用池取**再回填一遍，
    #   避免任何间接取数（结转记录等）导致汇总表显示 0。
    for group in groups:
        pools = {
            pool.workshop: pool
            for pool in group["period"].workshop_cost_pools.all()
        }
        group["pool_count"] = len(pools)
        for row in group["rows"]:
            pool = pools.get(row["workshop"])
            subtotal = ZERO_CNY
            for cell in row["cells"]:
                value = _cost_money(
                    getattr(pool, cell["field"], ZERO_CNY) if pool else ZERO_CNY
                )
                cell["value"] = value
                cell["input_value"] = "{:,.2f}".format(value)
                subtotal += value
            row["total"] = _cost_money(subtotal)

    # 「车间费用汇总表（按车间）」自己的筛选（用户 2026-09-27）：期间 + 车间。
    # 只影响这张汇总表与它的「导出Excel」，不影响上面的录入表（录入表固定=当前期间）。
    summary_filter = {
        "period": str(request.GET.get("sum_period", "")).strip(),
        "workshop": str(request.GET.get("sum_workshop", "")).strip(),
    }
    summary_filter["active"] = bool(summary_filter["period"] or summary_filter["workshop"])
    summary_groups = groups
    if summary_filter["period"]:
        summary_groups = [
            group for group in groups
            if group["period_label"] == summary_filter["period"]
        ]
    if summary_filter["period"]:
        summary_rows, summary_footer = _workshop_cost_summary(
            summary_groups, workshop=summary_filter["workshop"]
        )
    else:
        # 用户 2026-10-14：汇总表默认**按月分行**显示历史（7 月一块、8 月一块…），
        #   末尾再给一行"各期间合计"，既能看到历史、也能看累计。
        summary_rows = []
        for group in summary_groups:
            part, _foot = _workshop_cost_summary(
                [group], workshop=summary_filter["workshop"]
            )
            summary_rows.extend(part)
        _all_rows, summary_footer = _workshop_cost_summary(
            summary_groups, workshop=summary_filter["workshop"]
        )
    # 诊断（用户 2026-10-14）：明确告诉页面"这段期间在数据库里到底有没有费用记录"
    _diag = []
    for group in summary_groups:
        n = group.get("pool_count", 0)
        if n:
            _diag.append(f"{group['period_label']}：库里有 {n} 条费用记录")
        else:
            _diag.append(f"{group['period_label']}：库里没有费用记录")
    summary_note = "；".join(_diag)

    summary_period_options = [
        {"value": group["period_label"], "label": group["period_label"]}
        for group in sorted(groups, key=lambda g: g["period_label"], reverse=True)
    ]
    summary_workshop_options = []
    _seen_ws = set()
    for group in groups:
        for row in group["rows"]:
            if row["workshop"] in _seen_ws:
                continue
            _seen_ws.add(row["workshop"])
            summary_workshop_options.append(
                {"value": row["workshop"], "label": row["label"],
                 "sequence": row["sequence"]}
            )
    summary_workshop_options.sort(key=lambda item: item["sequence"])
    return render(
        request,
        "settlement/workshop_costs.html",
        {
        "summary_note": summary_note,
        "groups": groups,
            "period": period,
            "entry_group": entry_group,
            "columns": WORKSHOP_COST_COLUMNS,
            "summary_rows": summary_rows,
            "summary_footer": summary_footer,
            "summary_filter": summary_filter,
            "summary_period_options": summary_period_options,
            "summary_workshop_options": summary_workshop_options,
            "current_locked": current_locked,
            "can_save": can_save,
            "manual_batch_name": MANUAL_WORKSHOP_COST_NAME,
        },
    )


@login_required
def order_material_report(request, period_id):
    """工单投料：除浆粕车间外用「工单投料台账」、浆粕车间按库存流水备注；按车间分页签展示。

    两个模块各自带筛选（互不影响）：明细模块 d_* 前缀，汇总模块 s_* 前缀。
    投料期间 month / 投料车间 workshop / 物料分类 category / 工单号·单据号 order_no / 物料 material。
    """
    period = get_object_or_404(AccountingPeriod, id=period_id)
    rows, totals = build_material_issue_rows(period)

    # 台账行补齐「投料期间」和明细口径的统一字段（筛选与汇总都用这些字段）
    for _r in rows:
        _when = _r.get("issue_time")
        _r["month"] = (
            _when.strftime("%Y-%m") if _when else period.month.strftime("%Y-%m")
        )
        _r["source"] = "工单投料台账"
        _r["reference"] = _r.get("work_order_no") or ""
        _r["no_price"] = _r.get("unit_price") is None

    # 浆粕车间没有工单：按库存流水「领料出库」的备注认领，既并入汇总表，也作为明细行进车间页签
    pulp_detail_rows = []
    for item in pulp_issue_rows(period):
        when = item.get("date")
        pulp_detail_rows.append(
            {
                "month": when.strftime("%Y-%m") if when else period.month.strftime("%Y-%m"),
                "work_order_no": "",
                "document_no": item["document_no"],
                "reference": item["document_no"] or "（领料出库）",
                "source": "库存领料出库",
                "material_category": item["category"],
                "material_code": item["material_code"],
                "material_name": item["material_name"],
                "material_spec": "",
                "issue_uom": item.get("uom", ""),
                # 用户 2026-09-21 要求明细补这 6 列；浆粕没有工单台账，
                # 只能填流水里有的（仓位、时间），其余留空（页面上显示 —）
                "batch_no": item.get("batch_no", ""),
                "issue_location": item.get("warehouse", ""),
                "issue_time": item.get("date"),
                "issuer": item.get("issuer", ""),
                "main_output_code": item.get("main_output_code", ""),
                "main_output_name": item.get("main_output_name", ""),
                "actual_issue_quantity": item["quantity"],
                "unit_price": item["unit_price"],
                "amount": item["amount"],
                "basis": item["basis"],
                "workshop": item["workshop"],
                "no_price": item["unit_price"] is None
                or (item["amount"] == 0 and item["quantity"] != 0),
            }
        )

    all_rows = rows + pulp_detail_rows

    # ── 筛选：两个模块各自独立 ──────────────────────────────────────────────
    # 「投料汇总」模块：s_month / s_workshop / s_category
    # 「投料明细表」模块：d_month / d_workshop / d_category / d_order_no / d_material
    def _pick(prefix):
        return {
            "month": str(request.GET.get(prefix + "month", "")).strip(),
            "workshop": str(request.GET.get(prefix + "workshop", "")).strip(),
            "category": str(request.GET.get(prefix + "category", "")).strip(),
            "order_no": str(request.GET.get(prefix + "order_no", "")).strip(),
            "material": str(request.GET.get(prefix + "material", "")).strip(),
        }

    summary_filters = _pick("s_")
    detail_filters = _pick("d_")

    def _keep(row, filters):
        if filters["month"] and row["month"] != filters["month"]:
            return False
        if filters["workshop"] and (row["workshop"] or "") != filters["workshop"]:
            return False
        if filters["category"] and (row.get("material_category") or "") != filters["category"]:
            return False
        if filters["order_no"] and filters["order_no"].lower() not in (
                row.get("reference") or "").lower():
            return False
        if filters["material"]:
            haystack = "%s %s" % (
                row.get("material_code") or "", row.get("material_name") or "")
            if filters["material"].lower() not in haystack.lower():
                return False
        return True

    # 汇总模块只看期间/车间/分类；明细模块还看工单号与物料
    summary_filter_fields = ("month", "workshop", "category")
    detail_filter_fields = ("month", "workshop", "category", "order_no", "material")
    summary_active = any(summary_filters[k] for k in summary_filter_fields)
    detail_active = any(detail_filters[k] for k in detail_filter_fields)
    summary_source_rows = [row for row in all_rows if _keep(row, summary_filters)]
    filtered_rows = [row for row in all_rows if _keep(row, detail_filters)]

    from collections import defaultdict as _dd_om
    by_ws = _dd_om(list)
    for _r in filtered_rows:
        by_ws[_r["workshop"] or ""].append(_r)

    ordered = list(
        period.workshop_settlements.order_by("sequence").values_list(
            "workshop", flat=True
        )
    )
    if not ordered:
        ordered = list(WORKSHOP_LABELS.keys())
    present = [ws for ws in ordered if ws in by_ws] + [
        ws for ws in by_ws if ws not in ordered
    ]
    sections = []
    for ws in present:
        ws_rows = by_ws[ws]
        qty = sum((r["actual_issue_quantity"] for r in ws_rows), Decimal("0"))
        amt = sum((r["amount"] for r in ws_rows if r["amount"] is not None), Decimal("0"))
        no_price_n = sum(1 for r in ws_rows if r["unit_price"] is None)
        no_price_mats = {}
        for r in ws_rows:
            if r["unit_price"] is None:
                no_price_mats.setdefault(r["material_code"], r["material_name"] or "")
        sections.append(
            {
                "workshop": ws,
                "label": WORKSHOP_LABELS.get(ws, ws or "未归属"),
                "rows": ws_rows,
                "no_price_materials": no_price_mats,
                "subtotal": {"issue_quantity": qty, "amount": amt, "no_price_count": no_price_n},
            }
        )
    active_ws = detail_filters["workshop"]
    if active_ws not in present:
        active_ws = present[0] if present else ""
    show_all = request.GET.get("full") == "1"
    # 汇总表用「投料汇总」模块自己的筛选结果
    summary_rows, summary_totals = build_material_issue_summary(
        [],
        period,
        extra_rows=[
            {
                "month": row["month"],
                "workshop": row["workshop"] or "",
                "category": (row.get("material_category") or "").strip() or "未分类",
                "quantity": row["actual_issue_quantity"],
                "amount": row["amount"],
                "no_price": row["no_price"],
            }
            for row in summary_source_rows
        ],
    )
    # 行数按来源拆开（台账 / 浆粕流水），供页面显示
    summary_totals["ledger_rows"] = sum(
        1 for row in summary_source_rows if row.get("source") == "工单投料台账"
    )
    summary_totals["extra_rows"] = sum(
        1 for row in summary_source_rows if row.get("source") == "库存领料出库"
    )
    # 筛选项候选值（用全量数据，便于来回切换）
    month_options = sorted({row["month"] for row in all_rows})
    ws_codes = {row["workshop"] or "" for row in all_rows}
    workshop_options = [
        {"value": code, "label": WORKSHOP_LABELS.get(code, code or "未归属")}
        for code in sorted(ws_codes, key=lambda c: WORKSHOP_ORDER_INDEX.get(c, 99))
    ]
    category_options = sorted({
        (row.get("material_category") or "").strip()
        for row in all_rows
        if (row.get("material_category") or "").strip()
    })
    response = render(
        request,
        "settlement/order_material_report.html",
        {
            "period": period,
            "rows": rows,
            "totals": totals,
            "sections": sections,
            "active_ws": active_ws,
            "show_all": show_all,
            "summary_rows": summary_rows,
            "summary_totals": summary_totals,
            "pulp_rows": len(pulp_detail_rows),
            "summary_filters": summary_filters,
            "detail_filters": detail_filters,
            "summary_active": summary_active,
            "detail_active": detail_active,
            "month_options": month_options,
            "workshop_options": workshop_options,
            "category_options": category_options,
        },
    )
    # 禁止浏览器缓存，避免改版后用户看到旧页面
    response["Cache-Control"] = "no-store"
    return response


def _parse_month(value):
    """「2026-07」→ (2026, 7)；不是这个形状就返回 None。"""
    text = str(value or "").strip()
    if not re.match(r"^\d{4}-\d{1,2}$", text):
        return None
    year, month = text.split("-")
    month = int(month)
    if not 1 <= month <= 12:
        return None
    return int(year), month


def _month_label(value):
    """时间 → 「2026-7」（用户 2026-09-21：期间就显示 2026-7 这种，不补零）；空 → 空串。

    库里存的是 UTC（USE_TZ=True）：先转本地时间再取年月，才不会比上传文件里的时间
    早 8 小时而跨月（与 time_text 过滤器、页面显示的报工时间一致）。
    """
    if value is None:
        return ""
    # 只对「带时间的」做时区换算：period.month 是纯日期，没有 utcoffset
    if hasattr(value, "hour") and timezone.is_aware(value):
        value = timezone.localtime(value)
    return "%d-%d" % (value.year, value.month)


def _order_completion_report_rows(request, period):
    """工单报工明细表的取数（页面与「导出Excel」共用同一份）。

    筛选参数（两个模块各带前缀，互不影响）：
      期间 d_period（报工时间的年月，YYYY-MM；空＝全部）
      工单部门 d_department（报工记录的「工单部门」，包含即可）
      工单编号 d_order_no（包含即可）
    老链接/书签里不带前缀的 order_no / workshop / product（＝报工物料编号，兼容导出
    原来的 material）/ quality_status 照旧生效。
    """
    reports = period.work_reports.all()

    def _param(name):
        return str(
            request.GET.get("d_%s" % name, "") or request.GET.get(name, "")
        ).strip()

    order_no = _param("order_no")
    workshop = str(request.GET.get("workshop", "")).strip()
    department = _param("department")
    product = _param("product") or str(request.GET.get("material", "")).strip()
    quality_status = _param("quality_status")
    month = _parse_month(request.GET.get("d_period", ""))
    if order_no:
        reports = reports.filter(work_order_no__icontains=order_no)
    if department:
        reports = reports.filter(work_order_department__icontains=department)
    if product:
        reports = reports.filter(material_code__icontains=product)
    if quality_status:
        reports = reports.filter(quality_status__icontains=quality_status)
    if month:
        # USE_TZ=True 时 Django 会按当前时区换算年月，与页面显示的本地时间一致
        reports = reports.filter(report_time__year=month[0], report_time__month=month[1])
    if workshop:
        reports = reports.filter(
            work_order_no__in=period.work_orders.filter(
                production_dept_code__icontains=workshop
            ).values("order_no")
        )
    return reports


# 报工记录原样的列（上传文件缺失时的兜底：列名与字段顺序同「导出Excel」）
WORK_REPORT_FALLBACK_COLUMNS = (
    ("报工时间", "report_time"),
    ("生产任务编号", "task_no"),
    ("工序编号", "operation_no"),
    ("工序名称", "operation_name"),
    ("报工物料编码", "material_code"),
    ("报工物料名称", "material_name"),
    ("生产人员", "operator"),
    ("质量状态", "quality_status"),
    ("不良原因", "defect_reason"),
    ("报工数量", "quantity"),
    ("报工单位", "uom"),
    ("辅助数量", "auxiliary_quantity"),
    ("辅助单位", "auxiliary_uom"),
    ("批次号", "batch_no"),
    ("标识码", "identifier"),
    ("流转卡编号", "route_card_no"),
    ("工时", "work_hours"),
    ("工时单位", "work_hours_uom"),
    ("备注", "remark"),
    ("部门", "department"),
    ("设备", "equipment"),
    ("工单编号", "work_order_no"),
    ("工单物料编码", "work_order_material_code"),
    ("工单物料名称", "work_order_material_name"),
    ("订单号", "order_no"),
    ("纸管颜色", "paper_tube_color"),
    ("纺纱班别", "spinning_shift"),
    ("含水率", "moisture_content"),
    ("批号流水", "dye_batch_flow"),
    ("织布序号", "weaving_sequence"),
    ("织布机台", "weaving_machine"),
    ("产出仓位", "output_location"),
    ("工单部门", "work_order_department"),
)

# 明细表在报工记录之后追加的 6 个计算列（与「导出Excel」逐列一致）
WORK_REPORT_EXTRA_COLUMNS = (
    "完工成本",
    "单位完工成本",
    "在产品成本",
    "总成本",
    "报工入库数量",
    "报工入库金额",
)

#: 明细表最前面加的一列：期间＝报工时间的年月（用户 2026-09-21）
WORK_REPORT_PERIOD_COLUMN = "期间"


def _order_completion_report_sheets(period, reports):
    """明细表＝期间 + 上传的报工记录原样（列名/顺序按上传文件）+ 6 个计算列。

    返回 (列名, 数据行, 工单编号所在列下标)。页面与「导出Excel」用的是同一份，
    所以页面上看到的就是导出的那张表。
    """
    header, rows, _ = _uploaded_ledger_rows(
        period, ImportBatch.Kind.WORK_REPORT, reports
    )
    results = {
        item.work_order.order_no: item
        for item in period.work_order_cost_results.select_related("work_order").all()
    }
    receipt = {}
    for row in (
        period.inventory_transactions.filter(business_type__in=FINISHED_RECEIPT_TYPES)
        .values("document_no")
        .annotate(txn_qty=Sum("quantity"), txn_amount=Sum("amount"))
    ):
        receipt[row["document_no"]] = row

    def extra(order_no):
        item = results.get(order_no)
        record = receipt.get(order_no)
        return (
            item.finished_cost if item else None,
            item.finished_unit_cost if item else None,
            item.wip_cost if item else None,
            item.total_cost if item else None,
            record["txn_qty"] if record else None,
            record["txn_amount"]
            if record and record["txn_amount"] is not None
            else None,
        )

    alias = {}
    for field, names in ALIASES[ImportBatch.Kind.WORK_REPORT].items():
        for name in names:
            alias[normalize_header(name)] = field

    if header:
        # 有上传文件：列名/列顺序完全照上传文件，逐列按别名映射到字段
        column_names = [str(name) for name in header]
        fields = [alias.get(normalize_header(name)) for name in header]
        raw_rows = [tuple(row) for row in rows]
    else:
        # 干净环境/测试：列名与字段顺序与导出兜底一致
        column_names = [name for name, _ in WORK_REPORT_FALLBACK_COLUMNS]
        fields = [field for _, field in WORK_REPORT_FALLBACK_COLUMNS]
        raw_rows = [
            tuple(getattr(obj, field, None) for field in fields) for obj in reports
        ]

    time_idx = fields.index("report_time") if "report_time" in fields else None
    order_idx = fields.index("work_order_no") if "work_order_no" in fields else None

    columns = [WORK_REPORT_PERIOD_COLUMN] + column_names + list(
        WORK_REPORT_EXTRA_COLUMNS
    )
    data = []
    for row in raw_rows:
        order_no = (
            str(row[order_idx]) if order_idx is not None and row[order_idx] else ""
        )
        month = _month_label(row[time_idx]) if time_idx is not None else ""
        data.append((month,) + row + extra(order_no))
    return columns, data, (order_idx + 1 if order_idx is not None else None)


def _pulp_workshop_totals(period):
    """浆粕车间的六个数字（用户 2026-09-21：按浆粕车间自己的流水来）。

    浆粕没有工单、没有投料、没有报工记录，所以照它自己的口径补一行：
      报工数量 / 报工入库数量 = 浆粕成品「生产入库」流水数量合计（＝产量）
      报工入库金额 = 这些流水的金额合计（按浆粕完工成本回填过的金额）
      完工成本 = 浆粕成本归集（原料领料 + 车间费用，按产量分到成品）；
                 原料流水缺失/归类不了时退回用流水的金额合计，绝不让汇总页挂掉
      在产品成本 = 0（浆粕成本全部分摊到成品，不留在产品）
      总成本 = 完工成本
    该期间没有任何浆粕成品流水时返回 None（不出这一行）。
    """
    zero = Decimal("0")
    quantity = zero
    amount = zero
    for txn in period.inventory_transactions.filter(business_type="生产入库"):
        if not _is_pulp_product(txn.material_code, txn.material_name):
            continue
        quantity += Decimal(txn.quantity or zero)
        amount += Decimal(txn.amount or zero)

    finished = None
    try:
        costs = collect_pulp_costs(period)
        finished = sum(
            (Decimal(item["total"] or zero) for item in costs["products"].values()),
            zero,
        )
    except Exception:
        # 未上传浆粕原料流水、或原料无法归类 → 退回用流水上的金额
        finished = None
    if finished is None:
        finished = amount
    if not (quantity or amount or finished):
        return None
    return {
        "report_quantity": quantity,
        "finished_cost": finished,
        "wip_cost": zero,
        "total_cost": finished,
        "receipt_quantity": quantity,
        "receipt_amount": amount,
    }


#: 汇总表的六个数字（＝改版前页头的六个数字，只是按期间 × 车间拆开）
COMPLETION_SUMMARY_KEYS = (
    "report_quantity",
    "finished_cost",
    "wip_cost",
    "total_cost",
    "receipt_quantity",
    "receipt_amount",
)


def _completion_period_buckets(period):
    """一个结算期间：{车间编码: 六个数字 + order_count}（含浆粕车间）。

    车间取工单主档的「生产部门编号」（与页面显示一致）；浆粕车间没有工单，
    按浆粕自己的流水补上（见 _pulp_workshop_totals）。
    """
    zero = Decimal("0")
    masters = {wo.order_no: wo for wo in period.work_orders.all()}
    report_agg = {
        row["work_order_no"]: row
        for row in period.work_reports.exclude(work_order_no="")
        .values("work_order_no")
        .annotate(report_qty=Sum("quantity"))
    }
    results = {
        item.work_order.order_no: item
        for item in period.work_order_cost_results.select_related("work_order").all()
    }
    receipt_agg = {}
    for row in (
        period.inventory_transactions.filter(business_type__in=FINISHED_RECEIPT_TYPES)
        .values("document_no")
        .annotate(txn_qty=Sum("quantity"), txn_amount=Sum("amount"))
    ):
        receipt_agg[row["document_no"]] = row

    buckets = {}

    def bucket(code):
        return buckets.setdefault(
            code,
            dict(
                {key: zero for key in COMPLETION_SUMMARY_KEYS},
                order_count=0,
            ),
        )

    for number in sorted(set(masters) | set(report_agg)):
        wo = masters.get(number)
        row_bucket = bucket((wo.production_dept_code if wo else "") or "")
        row_bucket["order_count"] += 1
        agg = report_agg.get(number)
        if agg and agg["report_qty"] is not None:
            row_bucket["report_quantity"] += Decimal(agg["report_qty"])
        item = results.get(number)
        if item:
            row_bucket["finished_cost"] += item.finished_cost
            row_bucket["wip_cost"] += item.wip_cost
            row_bucket["total_cost"] += item.total_cost
        record = receipt_agg.get(number)
        if record:
            if record["txn_qty"] is not None:
                row_bucket["receipt_quantity"] += Decimal(record["txn_qty"])
            if record["txn_amount"] is not None:
                row_bucket["receipt_amount"] += Decimal(record["txn_amount"])

    pulp = _pulp_workshop_totals(period)
    if pulp:
        pulp_bucket = bucket(PULP_WORKSHOP)
        for key in COMPLETION_SUMMARY_KEYS:
            pulp_bucket[key] += pulp[key]
    return buckets


def _completion_summary_periods(period, selected):
    """汇总要统计哪些期间。

    用户 2026-09-21：这张表是累计的（以后的月份也要加进去），但**默认只显示当月**；
    选「全部期间」时统计库里已有的全部期间。
    """
    if selected == "all":
        return list(AccountingPeriod.objects.order_by("month"))
    month = _parse_month(selected)
    if month:
        matched = [
            item
            for item in AccountingPeriod.objects.order_by("month")
            if (item.month.year, item.month.month) == month
        ]
        if matched:
            return matched
    return [period]


def _order_completion_report_summary(period, summary_filters):
    """报工汇总（按期间 × 车间，累计口径）：(数据行, 合计, 合计文案, 车间选项)。

    六个数字的口径与改版前页头完全一致，只是拆成「期间 × 车间」并加上浆粕车间；
    合计＝当前筛选下的加总（默认只算本期，所以就是原来页头那六个数字 + 浆粕车间）。
    """
    zero = Decimal("0")
    periods = _completion_summary_periods(period, summary_filters["period"])
    rows = []
    for item in periods:
        month_label = _month_label(item.month)
        for code, values in _completion_period_buckets(item).items():
            row = dict(values)
            row["month_label"] = month_label
            row["workshop_code"] = code
            row["workshop_label"] = WORKSHOP_LABELS.get(code, code or "未归属")
            rows.append(row)

    workshop_codes = sorted(
        {row["workshop_code"] for row in rows},
        key=lambda code: (WORKSHOP_ORDER_INDEX.get(code, 99), code),
    )
    workshop_options = [
        {"value": code, "label": WORKSHOP_LABELS.get(code, code or "未归属")}
        for code in workshop_codes
        if code
    ]
    if summary_filters["workshop"]:
        rows = [
            row
            for row in rows
            if row["workshop_code"] == summary_filters["workshop"]
        ]
    # 六个数字全是 0 的车间行不显示（本期没有报工/成本/入库的车间），避免空行噪音
    rows = [
        row for row in rows if any(row[key] for key in COMPLETION_SUMMARY_KEYS)
    ]
    rows.sort(
        key=lambda row: (
            row["month_label"],
            WORKSHOP_ORDER_INDEX.get(row["workshop_code"], 99),
            row["workshop_code"],
        )
    )

    totals = {
        key: sum((row[key] for row in rows), zero) for key in COMPLETION_SUMMARY_KEYS
    }
    order_count = sum(row["order_count"] for row in rows)
    months = sorted({row["month_label"] for row in rows})
    if len(months) == 1:
        total_label = "合计（%s · %d 张工单）" % (months[0], order_count)
    else:
        total_label = "合计（%d 个期间 · %d 张工单）" % (len(months), order_count)
    return rows, totals, total_label, workshop_options


def _order_completion_detail_options(period):
    """明细表的筛选候选项：期间＝报工时间的年月；工单部门＝报工记录里的取值。"""
    months = set()
    departments = set()
    for report_time, department in period.work_reports.values_list(
        "report_time", "work_order_department"
    ):
        label = _month_label(report_time)
        if label:
            months.add(label)
        department = str(department or "").strip()
        if department:
            departments.add(department)
    return sorted(months), sorted(departments)




@login_required
def order_completion_report(request, period_id):
    """工单报工：明细表＝期间 + 上传的报工记录原样 + 6 个计算列（与「导出Excel」一致）；
    汇总＝按期间 × 车间的六个数字（累计，默认本月，含浆粕车间）。

    用户 2026-09-21：①汇总表把浆粕车间也放进去（按浆粕自己的流水），累计、默认只显示当月，
    加「期间」「车间」筛选；②明细表删掉标题里的括注，加一列「期间」（＝报工时间的年月），
    加「期间 / 工单部门 / 工单编号」筛选。
    """
    period = get_object_or_404(AccountingPeriod, id=period_id)
    reports = _order_completion_report_rows(request, period)
    columns, data, order_idx = _order_completion_report_sheets(period, reports)

    detail_filters = {
        "period": str(request.GET.get("d_period", "")).strip(),
        "department": str(request.GET.get("d_department", "")).strip(),
        "order_no": str(request.GET.get("d_order_no", "")).strip(),
    }
    summary_filters = {
        "period": str(request.GET.get("s_period", "")).strip(),
        "workshop": str(request.GET.get("s_workshop", "")).strip(),
    }
    summary_rows, totals, total_label, workshop_options = (
        _order_completion_report_summary(period, summary_filters)
    )
    detail_months, detail_departments = _order_completion_detail_options(period)
    current_month = _month_label(period.month)
    period_options = [
        _month_label(item.month)
        for item in AccountingPeriod.objects.order_by("month")
        if _month_label(item.month) != current_month
    ]

    order_total = len(
        {row[order_idx] for row in data if order_idx is not None and row[order_idx]}
    )
    # 页面上把「当前生效的筛选条件」写出来（用户 2026-09-21：选了没变化 → 至少能看出筛了什么）
    summary_active_labels = []
    if summary_filters["period"]:
        summary_active_labels.append(
            "期间＝全部期间（累计）"
            if summary_filters["period"] == "all"
            else "期间＝%s" % summary_filters["period"]
        )
    if summary_filters["workshop"]:
        summary_active_labels.append(
            "车间＝%s"
            % WORKSHOP_LABELS.get(
                summary_filters["workshop"], summary_filters["workshop"]
            )
        )
    detail_active_labels = []
    if detail_filters["period"]:
        detail_active_labels.append("期间＝%s" % detail_filters["period"])
    if detail_filters["department"]:
        detail_active_labels.append("工单部门＝%s" % detail_filters["department"])
    if detail_filters["order_no"]:
        detail_active_labels.append("工单编号含「%s」" % detail_filters["order_no"])
    response = render(
        request,
        "settlement/order_completion_report.html",
        {
            "period": period,
            "page": _page(request, data),
            "columns": columns,
            "row_total": len(data),
            "order_total": order_total,
            "summary_rows": summary_rows,
            "totals": totals,
            "total_label": total_label,
            "current_month": current_month,
            "period_options": period_options,
            "workshop_options": workshop_options,
            "detail_months": detail_months,
            "detail_departments": detail_departments,
            "summary_filters": summary_filters,
            "detail_filters": detail_filters,
            # 两个模块各自的筛选状态（页面上给个提示，看得见「筛了什么」）
            "summary_active": bool(summary_active_labels),
            "summary_active_labels": summary_active_labels,
            "detail_active": bool(detail_active_labels),
            "detail_active_labels": detail_active_labels,
            # 两个模块的筛选互不影响：各自的「清空」只清自己，另一个模块的条件原样带过去
            "summary_hidden_qs": urlencode(
                {"s_%s" % key: value for key, value in summary_filters.items() if value}
            ),
            "detail_hidden_qs": urlencode(
                {"d_%s" % key: value for key, value in detail_filters.items() if value}
            ),
        },
    )
    # 禁止浏览器缓存，避免改版后用户看到旧页面
    response["Cache-Control"] = "no-store"
    return response


@login_required
def purchase_reversal_report(request, period_id):
    """收到发票暂估回冲明细 —— 2026-09-19 用户要求：已并入「暂估回冲」页。

    这里保留 URL 只做跳转，保证旧链接/书签/凭证页按钮照旧可用（查询参数原样带过去）。
    """
    url = reverse("purchase_accruals", args=[period_id])
    query = request.GET.urlencode()
    return redirect("%s?%s" % (url, query) if query else url)


def _purchase_summary_stats(period):
    """暂估回冲页的统计口径：返回与 purchase_summary 模板一致的 context 数据。

    供单期间汇总页（purchase_summary）与各期间汇总页（purchase_accrual_overview）复用。
    """
    from settlement.services.purchase_accruals import provisional_basis

    ZERO = Decimal("0")

    def money(v):
        return (Decimal(v or ZERO)).quantize(Decimal("0.01"))

    # —— 1. 当月入库（本期采购入库流水）——
    receipts = [
        t for t in period.inventory_transactions.select_related("period").all()
        if is_purchase_receipt(t)
    ]
    receipt_qty = sum((abs(Decimal(t.quantity or ZERO)) for t in receipts), ZERO)
    receipt_amount = ZERO
    receipt_unpriced = 0
    for t in receipts:
        basis = provisional_basis(t)
        if basis["unit_price"] is not None:
            receipt_amount += Decimal(t.quantity or ZERO) * basis["unit_price"]
        else:
            receipt_unpriced += 1

    # —— 2. 发票 ——
    invoices = list(period.purchase_invoices.all())
    total_invoice = sum((money(i.amount_excl_tax) for i in invoices), ZERO)
    matched_ids = set(
        MatchAllocation.objects.filter(
            kind=MatchAllocation.Kind.PURCHASE_INVOICE,
            period__month__lte=period.month,
        ).values_list("purchase_invoice_line_id", flat=True)
    )
    matched = [i for i in invoices if i.id in matched_ids]
    matched_amount = sum((money(i.amount_excl_tax) for i in matched), ZERO)
    unmatched_amount = total_invoice - matched_amount

    # —— 3. 回冲（发票核销分配）——
    allocations = MatchAllocation.objects.filter(
        kind=MatchAllocation.Kind.PURCHASE_INVOICE,
        period__month__lte=period.month,
    ).select_related(
        "inventory_transaction__period",
        "opening_purchase_accrual",
        "purchase_invoice_line",
    )
    reversal_qty = sum(
        (Decimal(a.matched_quantity or ZERO) for a in allocations), ZERO
    )
    reversal_amount = sum((money(a.amount) for a in allocations), ZERO)
    # 累计口径的跨期笔数（既有「其中跨期核销」沿用）
    cross_count = 0
    for a in allocations:
        # 归属月份优先（用户 2026-09-19）：算出来早于本期 → 跨期核销（冲以前暂估）
        attribution = invoice_attribution_month(
            a.purchase_invoice_line, a.inventory_transaction
        )
        if attribution is not None:
            if attribution < period.month:
                cross_count += 1
        elif (
            a.inventory_transaction is not None
            and a.inventory_transaction.period.month < period.month
        ):
            cross_count += 1
        elif (
            a.opening_purchase_accrual is not None
            and a.opening_purchase_accrual.original_period < period.month
        ):
            cross_count += 1

    # —— 3b. 本月发生额（严格当月）见 _purchase_flow_stats()，供「月末采购暂估」页的四件套表用 ——

    # —— 4. 月末采购暂估 ——
    rows = build_purchase_accrual_rows(period)
    month_rows = [r for r in rows if r.receipt_month == period.month]
    prior_rows = [r for r in rows if r.receipt_month < period.month]
    total_accrual = sum((r.provisional_amount or ZERO for r in rows), ZERO)
    month_accrual = sum((r.provisional_amount or ZERO for r in month_rows), ZERO)
    prior_accrual = sum((r.provisional_amount or ZERO for r in prior_rows), ZERO)

    # —— 5. 期初采购暂估台账 ——
    opening = OpeningPurchaseAccrual.objects.filter(period__month__lte=period.month)
    opening_amount = sum((money(o.opening_amount) for o in opening), ZERO)

    # —— 6. 可视化数据（万元 + 占比）——
    def wan(v):
        return (Decimal(v or ZERO) / Decimal("10000")).quantize(Decimal("0.01"))

    receipt_amt = money(receipt_amount)
    reversal_amt = money(reversal_amount)
    month_accrual_amt = money(month_accrual)
    total_accrual_amt = money(total_accrual)
    inv_total = money(total_invoice)
    inv_matched = money(matched_amount)
    inv_unmatched = money(unmatched_amount)

    def pct(part, whole):
        if not whole:
            return 0
        return int((Decimal(part) / Decimal(whole) * 100).quantize(Decimal("0")))

    return {
        "period": period,
        "receipts": {
            "count": len(receipts),
            "quantity": receipt_qty,
            "amount": receipt_amt,
            "unpriced": receipt_unpriced,
        },
        "invoice": {
            "count": len(invoices),
            "total": inv_total,
            "matched_count": len(matched),
            "matched": inv_matched,
            "unmatched_count": len(invoices) - len(matched),
            "unmatched": inv_unmatched,
            "matched_pct": pct(inv_matched, inv_total),
            "unmatched_pct": pct(inv_unmatched, inv_total),
        },
        "reversal": {
            "count": allocations.count(),
            "quantity": reversal_qty,
            "amount": reversal_amt,
            "cross_count": cross_count,
        },
        "accrual": {
            "count": len(rows),
            "total": total_accrual_amt,
            "month_count": len(month_rows),
            "month": month_accrual_amt,
            "prior_count": len(prior_rows),
            "prior": money(prior_accrual),
        },
        "opening": {"count": opening.count(), "amount": money(opening_amount)},
        "flow": {
            "reversal_pct": pct(reversal_amt, receipt_amt),
            "accrual_pct": pct(month_accrual_amt, receipt_amt),
        },
        "kpi": {
            "receipts": wan(receipt_amt),
            "invoice": wan(inv_total),
            "invoice_matched": wan(inv_matched),
            "invoice_unmatched": wan(inv_unmatched),
            "reversal": wan(reversal_amt),
            "month": wan(month_accrual_amt),
            "total": wan(total_accrual_amt),
            "opening": wan(opening_amount),
        },
    }


def _purchase_reversal_blocks(request, period):
    """「回冲明细（逐笔）」板块的数据（2026-09-21 由「暂估回冲」页并入月末暂估）。

    返回：
      · detail_rows     本期发票核销逐笔（老链接里带的 invoice_no/supplier/material/cross_period
                        参数照旧生效，只是页面上已无筛选栏）；
      · detail_totals   本期逐笔合计（笔数/核销数量/暂估金额/发票金额/价差/跨期笔数）；
      · reversal_period_rows / reversal_period_total  逐期间汇总表（一个期间一行，可相加）。
    """
    periods = list(AccountingPeriod.objects.order_by("month"))
    detail_rows = build_purchase_reversal_rows(
        period,
        invoice_no=request.GET.get("invoice_no", "").strip(),
        supplier=request.GET.get("supplier", "").strip(),
        material=request.GET.get("material", "").strip(),
        cross_period=request.GET.get("cross_period", "").strip(),
    )
    filter_active = any(
        request.GET.get(key, "").strip()
        for key in ("invoice_no", "supplier", "material", "cross_period")
    )
    period_total = {
        "count": 0,
        "matched_quantity": Decimal("0"),
        "provisional_amount": Decimal("0"),
        "invoice_amount": Decimal("0"),
        "difference_amount": Decimal("0"),
        "cross_period_count": 0,
    }
    reversal_period_rows = []
    for item in periods:
        if item.id == period.id:
            period_rows = (
                detail_rows if not filter_active else build_purchase_reversal_rows(item)
            )
        else:
            period_rows = build_purchase_reversal_rows(item)
        totals = summarize_purchase_reversals(period_rows)
        reversal_period_rows.append(
            {
                "period": item,
                "current": item.id == period.id,
                **{key: totals[key] for key in period_total},
            }
        )
        for key in period_total:
            period_total[key] += totals[key]
    return {
        "detail_rows": detail_rows,
        "detail_totals": summarize_purchase_reversals(detail_rows),
        "detail_filter_active": filter_active,
        "reversal_period_rows": reversal_period_rows,
        "reversal_period_total": period_total,
    }


@login_required
def purchase_summary(request, period_id):
    """「暂估回冲」已并入「月末采购暂估」页（用户 2026-09-21）。

    这里保留 URL 只做跳转，老链接 / 书签照样能用（查询参数原样带过去）。
    """
    url = reverse("purchase_accruals", args=[period_id])
    query = request.GET.urlencode()
    return redirect("%s?%s" % (url, query) if query else url)


@login_required
def purchase_accrual_overview(request):
    """各期间暂估已并入「月末采购暂估」页：直接跳转到最近期间。"""
    latest = AccountingPeriod.objects.order_by("month").last()
    if latest is None:
        return redirect("dashboard")
    return redirect("purchase_accruals", period_id=latest.id)
@login_required
def work_order_imports(request, period_id):
    """生产工单导入台账（ERP 原始资料，不参与成本计算）。"""
    period = get_object_or_404(AccountingPeriod, id=period_id)
    orders = period.work_orders.all()
    orders = _text_filters(
        request,
        orders,
        {
            "order_no": "order_no",
            "line_code": "line_code",
            "status": "status",
        },
    )
    return render(
        request,
        "settlement/work_order_imports.html",
        {
            "period": period,
            "page": _page(request, orders),
        },
    )


@login_required
def material_issue_imports(request, period_id):
    """投料记录导入台账（ERP 原始资料，不参与成本计算）。"""
    period = get_object_or_404(AccountingPeriod, id=period_id)
    issues = period.material_issues.all()
    issues = _text_filters(
        request,
        issues,
        {
            "work_order_no": "work_order_no",
            "material_code": "material_code",
            "task_no": "task_no",
        },
    )
    return render(
        request,
        "settlement/material_issue_imports.html",
        {
            "period": period,
            "page": _page(request, issues),
        },
    )


@login_required
def work_report_imports(request, period_id):
    """报工记录导入台账（ERP 原始资料，不参与成本计算）。"""
    period = get_object_or_404(AccountingPeriod, id=period_id)
    reports = period.work_reports.all()
    reports = _text_filters(
        request,
        reports,
        {
            "work_order_no": "work_order_no",
            "material_code": "material_code",
            "quality_status": "quality_status",
        },
    )
    return render(
        request,
        "settlement/work_report_imports.html",
        {
            "period": period,
            "page": _page(request, reports),
        },
    )


def bom_qty_total_by_order(period):
    """工单编号 → 该工单 BOM 子项用量合计（Σ 分子÷分母）。

    口径与结算层 services/costing.py::_fill_bom_prices 一致：**整张工单的全部 BOM 行**，
    不受页面筛选影响（否则筛选后 Σ 变小，「折算到 100%」就失真了）。
    用于「用量(转100%)」与「转换后子项用量:分子」两列的折算。
    """
    totals = defaultdict(Decimal)
    for order_no, numerator, denominator in period.bom_lines.values_list(
        "work_order_no", "qty_numerator", "qty_denominator"
    ):
        if not order_no:
            continue
        denominator = denominator or Decimal("1")
        if not denominator:
            denominator = Decimal("1")
        totals[order_no] += (numerator or Decimal("0")) / denominator
    return totals


def bom_qty_is_scaled(qty_total):
    """该工单的 Σ用量是否**超过 100%**（结算层「标准成品单价」的折算阈值，同 costing 常量）。

    注意：这是**结算口径**的阈值（Σ用量 > 100% 才把配方折算到 100% 再算标准成品单价）；
    页面上的「用量(转100%)／转换后子项用量:分子」两列**一律按 Σ用量 折算**（用户 2026-09-30），
    见 bom_normalized_qty。
    """
    if not BOM_QTY_NORMALIZE_OVER:
        return False
    return (qty_total or Decimal("0")) > Decimal(BOM_QTY_NORMALIZE_OVER)


def bom_normalized_qty(qty, qty_total):
    """折算后用量（＝该子项占整张工单的**配比**）：**用量 ÷ 该工单 Σ用量**。

    用户 2026-09-30 更正：**一律用该工单标准用量之和折算**（每个物料 ÷ 总用量），
    各子项合计＝1（＝100%）；不再是「Σ用量 > 100% 才折算」——否则 Σ用量 < 100% 的工单
    （如机织 39.034H1-HF20250734-3，Σ用量 ＝ 0.38）会原样显示成 0.3537/0.0063/0.02，
    而正确配比是 0.930789473684211 / 0.0165789473684211 / 0.0526315789473684（合计 100%）。

    ⚠ 结算层的「标准成品单价」没有跟着改：那里仍是 Σ用量 > 100% 才折算
    （机织等的 Σ用量 就是实际单耗/米重，拿占比去乘单价会把单价放大 1/Σ 倍）。
    """
    total = Decimal(qty_total or 0)
    if total <= 0:
        return qty
    return qty / total


def bom_converted_numerator(qty, qty_total):
    """「转换后子项用量:分子」（**用量口径**）：折算后用量 × 100。

    用户 2026-09-23：源表「子项用量:分子 / 子项用量:分母」的分母写法很杂（1 / 10 / 100 / 110 / 1000）。
    用户 2026-09-30 更正：折算一律做（每个物料 ÷ 该工单 Σ用量）。

    ⚠ 该工单**有标准用量**时，「转换后子项用量:分子」列走的是 `bom_share_resolver`
    （＝标准用量 ÷ Σ标准用量，用户 2026-09-30）；本函数只剩「没有标准用量」时的兜底口径。
    """
    return bom_normalized_qty(qty, qty_total) * Decimal("100")


@login_required
def bom_imports(request, period_id):
    """BOM 表导入台账：子项单价与标准成品单价由系统按月末加权价回填。

    用户 2026-09-23：子项用量合计 **超过 100%** 的工单，标准成品单价按「用量折算到 100%」
    后再算（＝Σ(用量×单价) ÷ Σ用量）。页面上给每行加上「用量(转100%)」列便于核对。
    用户 2026-09-23 追加：再加一列「转换后子项用量:分子」——按标准用量把源表各种分母的
    分子统一换算成「分母＝100」的写法（口径见 bom_converted_numerator）。
    用户 2026-09-30 更正：这两列**一律**按该工单口径折算（不再只在 Σ用量 > 100% 时才折算）：
    「用量(转100%)」＝用量 ÷ 该工单 Σ用量；「转换后子项用量:分子」＝**标准用量 ÷ 该工单 Σ标准用量**
    （用户 2026-09-30：分子用标准用量；该工单有行缺标准用量时退回用量口径）——见 bom_share_resolver。
    """
    period = get_object_or_404(AccountingPeriod, id=period_id)
    lines = period.bom_lines.all()
    lines = _text_filters(
        request,
        lines,
        {
            "work_order_no": "work_order_no",
            "material_code": "material_code",
            "sub_item_code": "sub_item_code",
        },
    )
    lines = list(lines)
    # 按工单汇总：Σ用量（用量(转100%)）＋ Σ标准用量（转换后子项用量:分子）
    qty_total = bom_qty_total_by_order(period)
    std_total, std_missing = bom_standard_qty_basis(period)
    share = bom_share_resolver(period)
    for line in lines:
        total = qty_total.get(line.work_order_no) or Decimal("0")
        qty_ratio, converted_ratio = share(line)
        line.qty_total = total
        line.qty_is_scaled = bom_qty_is_scaled(total)
        line.qty_normalized = qty_ratio
        line.qty_converted_numerator = converted_ratio
        line.qty_share_standard = (
            line.standard_qty is not None
            and (std_total.get(line.work_order_no) or Decimal("0")) > 0
            and line.work_order_no not in std_missing
        )
        line.qty_share_total = (
            std_total.get(line.work_order_no) if line.qty_share_standard else total
        )
        # 取价口径：月末加权价 → 历史单价（出库价优先）；两处都没有 = 没价（页面显示「—」而不是 0.0000）
        line.unit_price_blank = not (line.unit_price or Decimal("0")) > Decimal("0")
        line.standard_price_blank = not (line.standard_price or Decimal("0")) > Decimal("0")
    return render(
        request,
        "settlement/bom_imports.html",
        {
            "period": period,
            "page": _page(request, lines),
        },
    )


def _production_cost_target_period(current, selected):
    """期间筛选：默认本期；选到某个月（2026-7 / 2026-07）就取那个月。"""
    month = _parse_month(selected)
    if month:
        for item in AccountingPeriod.objects.order_by("month"):
            if (item.month.year, item.month.month) == month:
                return item
    return current


def _production_cost_period_options(current):
    """下拉候选：有工单结算结果的期间（本期自身不重复列）。"""
    from settlement.models import WorkOrderCostResult

    ids = set(
        WorkOrderCostResult.objects.values_list("period_id", flat=True).distinct()
    )
    ids.add(current.id)
    current_label = current.month.strftime("%Y-%m")
    return [
        {"value": item.month.strftime("%Y-%m"), "label": item.month.strftime("%Y-%m")}
        for item in AccountingPeriod.objects.filter(id__in=ids).order_by("month")
        if item.month.strftime("%Y-%m") != current_label
    ]


def _workshop_issue_average_price(period, order_no):
    """该工单所在车间的「投料均价」＝车间投料金额 ÷ 车间投料数量（取不到返回 None）。

    用途：人工只填了投料数量、而这张工单原来根本没有投料记录时，用它折算金额。
    """
    workshop = (
        period.work_orders.filter(order_no=order_no)
        .values_list("production_dept_code", flat=True)
        .first()
    )
    if not workshop:
        return None
    prices = {
        item.material_code: Decimal(item.weighted_average_price or 0)
        for item in period.inventory_valuations.all()
    }
    for hp in period.historical_prices.all():
        if hp.issue_price is not None and not prices.get(hp.material_code):
            prices[hp.material_code] = Decimal(hp.issue_price)
        elif hp.receipt_price is not None and not prices.get(hp.material_code):
            prices[hp.material_code] = Decimal(hp.receipt_price)
    order_nos = set(
        period.work_orders.filter(production_dept_code=workshop)
        .values_list("order_no", flat=True)
    )
    qty_total = Decimal("0")
    amount_total = Decimal("0")
    for row in (
        period.material_issues.exclude(work_order_no="")
        .exclude(material_code="")
        .values("work_order_no", "material_code", "actual_issue_quantity")
    ):
        if row["work_order_no"] not in order_nos:
            continue
        qty = Decimal(row["actual_issue_quantity"] or 0)
        qty_total += qty
        amount_total += qty * Decimal(prices.get(row["material_code"], 0) or 0)
    if qty_total <= 0:
        return None
    return amount_total / qty_total


def _parse_adjustment_amount(raw, label):
    """人工调整录入：空 → None（表示"用系统值"）；其余转成数字（允许 0）。"""
    text = str(raw or "").strip().replace(",", "").replace("，", "")
    if not text:
        return None
    try:
        value = Decimal(text)
    except (InvalidOperation, ValueError):
        raise ValueError("%s「%s」不是数字。" % (label, raw))
    if value < 0:
        raise ValueError("%s不能为负数。" % label)
    return value


def _save_issue_adjustments(request, period):
    """保存「投料报工情形速查」里的人工调整，并重算涉及的车间。

    口径（用户 2026-09-24，方案 A「覆盖」）：填了就以人工值为准；清空＝恢复系统值。
    只处理本次真正提交的行（按工单号定位），保存后立刻重算这些工单所在车间的工单成本，
    这样「工单成本」「生产成本」页面的数字会跟着变；车间结转状态不动（页面会提示需重新结转）。
    """
    from settlement.models import WorkOrderManualAdjustment
    from settlement.services.costing import recalculate_workshop

    order_nos = [
        str(value).strip()
        for value in request.POST.getlist("order_no")
        if str(value).strip()
    ]
    if not order_nos:
        return "", "info", []
    workshops = set()
    saved = cleared = 0
    errors = []
    with transaction.atomic():
        for order_no in order_nos:
            fields = {
                "issue_quantity": _parse_adjustment_amount(
                    request.POST.get("adj_issue_qty_%s" % order_no), "投料数量"
                ),
                "issue_amount": _parse_adjustment_amount(
                    request.POST.get("adj_issue_amount_%s" % order_no), "投料金额"
                ),
                "report_quantity": _parse_adjustment_amount(
                    request.POST.get("adj_report_qty_%s" % order_no), "报工数量"
                ),
            }
            remark = str(request.POST.get("adj_remark_%s" % order_no, "")).strip()
            empty = all(value is None for value in fields.values())
            existing = WorkOrderManualAdjustment.objects.filter(
                period=period, order_no=order_no
            )
            workshop = (
                period.work_orders.filter(order_no=order_no)
                .values_list("production_dept_code", flat=True)
                .first()
            )
            if empty:
                if existing.exists():
                    existing.delete()
                    cleared += 1
                    # 清空也要重算，否则工单成本还停在人工值上
                    if workshop:
                        workshops.add(workshop)
                continue
            if fields["issue_quantity"] is not None and fields["issue_amount"] is None:
                # 只填数量：先按该单原投料单价折算；该单原来没有投料记录时，
                # 退回「车间投料均价」（车间投料金额÷车间投料数量），取不到才报错。
                has_issue = period.material_issues.filter(
                    work_order_no=order_no
                ).exists()
                if not has_issue:
                    average_price = _workshop_issue_average_price(period, order_no)
                    if average_price is None:
                        errors.append(
                            "%s 只填了投料数量，但这张工单和它所在车间本月都没有投料记录，"
                            "取不到投料单价，无法折算金额（请先补投料台账，或改用投料金额）。"
                            % order_no
                        )
                        continue
                    fields["issue_amount"] = (
                        Decimal(fields["issue_quantity"]) * average_price
                    )
            WorkOrderManualAdjustment.objects.update_or_create(
                period=period,
                order_no=order_no,
                defaults={
                    **fields,
                    "remark": remark,
                    "updated_by": request.user,
                },
            )
            saved += 1
            if workshop:
                workshops.add(workshop)
    if errors:
        return "；".join(errors), "error", []

    recalculated = []
    for workshop in sorted(workshops):
        try:
            recalculate_workshop(period, workshop)
            recalculated.append(WORKSHOP_LABELS.get(workshop, workshop))
        except Exception as exc:  # noqa: BLE001
            return (
                "已保存 %d 笔调整，但重算「%s」失败：%s"
                % (saved, WORKSHOP_LABELS.get(workshop, workshop), exc),
                "error",
                recalculated,
            )
    parts = []
    if saved:
        parts.append("已保存 %d 笔人工调整" % saved)
    if cleared:
        parts.append("已清除 %d 笔调整（恢复系统值）" % cleared)
    if recalculated:
        parts.append("并重算了 %s" % "、".join(recalculated))
    if parts:
        parts.append("如该车间已结转，请在「车间结转」重新结转一次（下游车间同理），账才闭环。")
    return "，".join(parts), "success" if not cleared else "success", recalculated


#: 工单明细里「创建时间」（ERP 建单时间）的月份提取（用户 2026-09-27）
_ORDER_PERIOD_RE = re.compile(r"(20\d{2})\D?(0?[1-9]|1[0-2])")


def _order_period_label(created_at) -> str:
    """把工单明细的「创建时间」解析成 YYYY-MM（工单期间＝下单月份）。

    上传格式可能是 2026-08-26 / 2026/8/26 / 20260826，统一取年月；
    解析不出来返回空串（页面上显示 —）。
    """
    match = _ORDER_PERIOD_RE.search(str(created_at or ""))
    if not match:
        return ""
    return "%s-%02d" % (match.group(1), int(match.group(2)))


def order_period_map(period) -> dict:
    """{工单号: 工单期间 YYYY-MM}＝工单明细上传的「创建时间」（ERP 建单时间）所属月份。"""
    mapping = {}
    for detail in period.work_order_details.only("order_no", "created_at"):
        label = _order_period_label(detail.created_at)
        if label and detail.order_no not in mapping:
            mapping[detail.order_no] = label
    return mapping


def attach_order_period_labels(period, rows):
    """给工单成本行挂上「工单期间」（下单月份），返回查不到创建时间的行数。

    用户 2026-09-27：核算期间＝这张单的成本算在哪个月；工单期间＝这张单的下单月份。
    """
    mapping = order_period_map(period)
    missing = 0
    for row in rows:
        work_order = getattr(row, "work_order", None)
        order_no = (
            getattr(work_order, "order_no", None)
            or getattr(row, "order_no", "")
        )
        label = mapping.get(order_no, "")
        row.order_period_label = label
        if not label:
            missing += 1
    return missing


def build_issue_detail_table(issue_period, issue_filters, issue_rows_all):
    """「二、工单投料报工明细」的取数（页面与导出共用同一套，用户 2026-09-27）。

    取数＝「工单成本」页的同一套（`annotate_work_order_cost_rows`），
    另外挂上：核算期间（成本算在哪个月）、工单期间（下单月份）、情形、人工调整标记。
    筛选：工单号 / 情形 / 工单期间（用户 2026-09-27 加了「工单期间」这一栏筛选）。
    返回 (detail_rows, detail_totals, order_period_missing)。
    """
    detail_qs = issue_period.work_order_cost_results.select_related("work_order")
    if issue_filters["workshop"] and issue_filters["workshop"] != PULP_WORKSHOP:
        detail_qs = detail_qs.filter(
            work_order__production_dept_code=issue_filters["workshop"]
        )
    elif issue_filters["workshop"] == PULP_WORKSHOP:
        detail_qs = detail_qs.none()   # 浆粕车间没有工单
    if issue_filters["order_no"]:
        detail_qs = detail_qs.filter(
            work_order__order_no__icontains=issue_filters["order_no"]
        )
    detail_rows = annotate_work_order_cost_rows(
        issue_period,
        list(
            detail_qs.order_by(
                "work_order__production_dept_code", "work_order__order_no"
            )
        ),
    )
    issue_by_order = {row["order_no"]: row for row in issue_rows_all}
    # 工单期间（用户 2026-09-27）：＝工单明细上传的「创建时间」（ERP 建单时间）所属月份，
    # 与「核算期间」（这张单的成本算在哪个月）区分开。先挂上再筛，否则筛不了。
    for row in detail_rows:
        row.detail_period_label = issue_period.month.strftime("%Y-%m")
    order_period_missing = attach_order_period_labels(issue_period, detail_rows)
    if issue_filters.get("case"):
        # 情形筛选也作用到明细表（只留该情形下的工单）
        # ⑤「正常投料生产工单」在小结里是 CASE_OTHER，明细行上是 CASE_NORMAL（"— 正常"）→ 两个都算命中。
        wanted = {issue_filters["case"]}
        if issue_filters["case"] == CASE_OTHER:
            wanted.add(CASE_NORMAL)
        keep = {
            row["order_no"] for row in issue_rows_all
            if row["case"] in wanted
        }
        detail_rows = [row for row in detail_rows if row.work_order.order_no in keep]
    if issue_filters.get("order_period"):
        # 「工单期间」筛选（用户 2026-09-27）：只看某个下单月份的单
        detail_rows = [
            row for row in detail_rows
            if row.order_period_label == issue_filters["order_period"]
        ]
    detail_totals = {"count": len(detail_rows)}
    for field in (
        "direct_material", "water_power", "steam", "natural_gas", "depreciation",
        "consumables", "dye_auxiliary", "processing_fee", "wages", "total_cost",
        "finished_cost", "wip_cost", "plan_quantity", "finished_quantity",
        "wip_quantity", "current_issue", "opening_quantity", "opening_amount",
    ):
        detail_totals[field] = sum(
            (Decimal(getattr(row, field) or 0) for row in detail_rows), Decimal("0")
        )
    for row in detail_rows:
        info = issue_by_order.get(row.work_order.order_no) or {}
        row.case_label = info.get("case", CASE_NORMAL)
        row.case_basis = CASE_COST_BASIS.get(row.case_label, "")  # 情形列悬停：成本怎么算
        row.case_has_manual = info.get("has_manual", False)
        row.case_manual_issue_qty = info.get("manual_issue_qty")
        row.case_raw_issue_qty = info.get("raw_issue_qty")
    return detail_rows, detail_totals, order_period_missing


#: 「二、工单投料报工明细」导出 Excel 的列（顺序与页面表格一致，用户 2026-09-27）
ISSUE_DETAIL_EXPORT_COLUMNS = (
    ("核算期间", "detail_period_label", "text"),
    ("工单期间", "order_period_label", "text"),
    ("车间", "_workshop_label", "text"),
    ("工单", "_order_no", "text"),
    ("类别", "wo_category", "text"),
    ("产品编码", "product_code", "text"),
    ("产品名称", "product_name", "text"),
    ("计划完工时间", "wo_planned_finish", "text"),
    ("实际完工时间", "wo_actual_finish", "text"),
    ("完成状态", "wo_completion_status", "text"),
    ("计划数量", "plan_quantity", "num"),
    ("期初数量", "opening_quantity", "num"),
    ("期初金额", "opening_amount", "num"),
    ("本期投料", "_issue_qty_effective", "num"),
    ("投料单价", "current_unit_price", "num2"),
    ("累计投料", "accum_issue", "num"),
    ("投料状态", "issue_status", "text"),
    ("本期完工", "current_finished", "num"),
    ("累计完工", "accum_finished", "num"),
    ("本期完工率%", "current_finish_rate_pct", "num"),
    ("累计完工率%", "finish_rate_pct", "num"),
    ("是否用粗纱", "uses_roving", "text"),
    ("报工数量", "finished_quantity", "num"),
    ("在产品", "wip_quantity", "num"),
    ("折合产量", "equivalent_qty", "num"),
    ("实际材料", "direct_material", "num"),
    ("反推材料", "backfill_material", "num"),
    ("水电", "water_power", "num"),
    ("水电分配率%", "rate_water_power_pct", "num"),
    ("蒸汽", "steam", "num"),
    ("蒸汽分配率%", "rate_steam_pct", "num"),
    ("天然气", "natural_gas", "num"),
    ("天然气分配率%", "rate_natural_gas_pct", "num"),
    ("折旧", "depreciation", "num"),
    ("折旧分配率%", "rate_depreciation_pct", "num"),
    ("机物料", "consumables", "num"),
    ("机物料分配率%", "rate_consumables_pct", "num"),
    ("染料助剂", "dye_auxiliary", "num"),
    ("染料助剂分配率%", "rate_dye_auxiliary_pct", "num"),
    ("加工费", "processing_fee", "num"),
    ("加工费分配率%", "rate_processing_fee_pct", "num"),
    ("工资", "wages", "num"),
    ("工资分配率%", "rate_wages_pct", "num"),
    ("总成本", "total_cost", "num"),
    ("完工成本", "finished_cost", "num"),
    ("单位成本", "finished_unit_cost", "num"),
    ("标准单位成本", "standard_unit_cost", "num"),
    ("单位成本差异率%", "unit_cost_diff_pct_text", "text"),
    ("投料差异", "issue_diff_text", "text"),
    ("单位成本差异原因", "unit_cost_diff_reason", "text"),
    ("期末在产品数量", "wip_quantity", "num"),
    ("在产品成本", "wip_cost", "num"),
    ("情形", "case_label", "text"),
)


def issue_detail_export_row(item):
    """按 ISSUE_DETAIL_EXPORT_COLUMNS 把一行明细转成导出用的元组。"""
    values = []
    for _header, attr, kind in ISSUE_DETAIL_EXPORT_COLUMNS:
        if attr == "_workshop_label":
            value = WORKSHOP_LABELS.get(
                item.work_order.production_dept_code,
                item.work_order.production_dept_code,
            )
        elif attr == "_order_no":
            value = item.work_order.order_no
        elif attr == "_issue_qty_effective":
            # 与页面上那个可手改的「本期投料」一致：填了人工值就以人工值为准
            manual = getattr(item, "case_manual_issue_qty", None)
            if manual in (None, ""):
                manual = getattr(item, "case_raw_issue_qty", None)
            value = Decimal(manual) if manual not in (None, "") else None
        else:
            value = getattr(item, attr, None)
        if kind == "num2":
            # 保留 2 位小数（用户 2026-09-27）：取不到价的显示 —，不写成 0
            value = (
                Decimal(value).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                if value not in (None, "")
                else "—"
            )
        elif kind == "num4":
            value = Decimal(value or 0) if value not in (None, "") else None
        elif kind == "num":
            value = Decimal(value) if value not in (None, "") else None
        else:
            value = value if value not in (None, "") else "—"
        values.append(value)
    return tuple(values)


def _export_query(**params):
    """把非空筛选参数拼成 "?a=1&b=2"（拼不出就返回空串）。"""
    clean = {key: value for key, value in params.items() if value not in (None, "")}
    return "?%s" % urlencode(clean) if clean else ""


def build_case_workshop_matrix(rows, workshop_options):
    """「情形 × 车间」工单数矩阵（用户 2026-09-27）。

    行＝各情形（含「正常投料生产工单」与合计），列＝车间（按车间顺序，只列本期出现的）
    ＋合计；单元格＝工单数。跟着「期间 / 车间」筛选走（情形筛选不影响它，否则看不到分布）。
    """
    present = [
        {"workshop": item["value"], "workshop_label": item["label"]}
        for item in workshop_options
        if any(row["workshop"] == item["value"] for row in rows)
    ]
    # 车间顺序里没有的（兜底）：按出现顺序补在后面，避免漏列
    present_codes = [item["workshop"] for item in present]
    for row in rows:
        if row["workshop"] not in present_codes:
            present_codes.append(row["workshop"])
            present.append({
                "workshop": row["workshop"],
                "workshop_label": row.get("workshop_label") or row["workshop"],
            })
    columns = [
        (item["workshop"], item["workshop_label"]) for item in present
    ]

    def _cells(case_rows):
        per = defaultdict(int)
        for row in case_rows:
            per[row["workshop"]] += 1
        return per

    matrix_rows = []
    for case in list(CASE_CHOICES) + [CASE_OTHER, CASE_NO_PRICE]:
        # 「正常投料生产工单」那一行＝情形为「— 正常」的单（与小结同一套口径）
        source = CASE_NORMAL if case == CASE_OTHER else case
        per = _cells([row for row in rows if row["case"] == source])
        cells = [
            {
                "workshop": code,
                "workshop_label": label,
                "count": per.get(code, 0),
            }
            for code, label in columns
        ]
        matrix_rows.append(
            {"case": case, "cells": cells, "count": sum(c["count"] for c in cells)}
        )
    totals = [
        {
            "workshop": code,
            "workshop_label": label,
            "count": sum(1 for row in rows if row["workshop"] == code),
        }
        for code, label in columns
    ]
    return {
        "workshops": present,
        "rows": matrix_rows,
        "totals": totals,
        "total": sum(item["count"] for item in totals),
    }


def _issue_order_period_options(issue_period, issue_rows_all):
    """「工单期间」下拉的候选：**本期这张表里真的出现过**的下单月份，新的在前。

    不直接用 order_period_map(period)（它覆盖该期间全部工单明细，会列出与本期这张表
    无关的老月份，选了却是 0 行）。
    """
    mapping = order_period_map(issue_period)
    labels = {
        mapping[row["order_no"]]
        for row in issue_rows_all
        if mapping.get(row["order_no"])
    }
    return sorted(labels, reverse=True)


def _issue_filters_from_request(request):
    """「二、工单投料报工明细」自己的筛选（页面与导出共用）。"""
    return {
        "period": str(request.GET.get("i_period", "")).strip(),
        "workshop": str(request.GET.get("i_workshop", "")).strip(),
        "case": str(request.GET.get("i_case", "")).strip(),
        "order_period": str(request.GET.get("i_order_period", "")).strip(),
        "order_no": str(request.GET.get("i_order_no", "")).strip(),
        "roving": str(request.GET.get("i_roving", "")).strip(),
    }


@login_required
def production_cost_summary(request, period_id):
    """生产成本（2026-09-21 与「总成本表」合并为一个页面）。

    页面 = 车间生产成本汇总（含合计行）→ 成本构成表（原总成本表主体；原名「成本项目×四大块」）
    → 选中车间的工单明细（下钻）。两张表各自带「期间 + 车间」筛选，互不影响；
    老链接里带的 `workshop` 参数仍然作用于两张表。
    """
    period = get_object_or_404(AccountingPeriod, id=period_id)
    from settlement.services.total_cost import build_total_cost_data

    # 保存人工调整（投料/报工）→ 立即重算相关车间的工单成本（用户 2026-09-24）
    if request.method == "POST":
        message, level, recalculated = _save_issue_adjustments(request, period)
        if message:
            {"success": messages.success, "error": messages.error}.get(
                level, messages.info
            )(request, message)
        # 重算后回到同一筛选（PRG 模式，避免刷新重复提交）
        target = request.get_full_path()
        return redirect(target)

    legacy_workshop = str(request.GET.get("workshop", "")).strip()
    summary_filters = {
        "period": str(request.GET.get("p_period", "")).strip(),
        "workshop": str(request.GET.get("p_workshop", "")).strip() or legacy_workshop,
    }
    matrix_filters = {
        "period": str(request.GET.get("m_period", "")).strip(),
        "workshop": str(request.GET.get("m_workshop", "")).strip() or legacy_workshop,
    }
    # 「投料报工情形速查」自己的筛选（用户 2026-09-24）：
    # 核算期间 / 车间 / 情形 / 工单号（＋2026-09-27 加的「工单期间」）
    issue_filters = {
        "period": str(request.GET.get("i_period", "")).strip(),
        "workshop": str(request.GET.get("i_workshop", "")).strip(),
        "case": str(request.GET.get("i_case", "")).strip(),
        "order_period": str(request.GET.get("i_order_period", "")).strip(),
        "order_no": str(request.GET.get("i_order_no", "")).strip(),
        "roving": str(request.GET.get("i_roving", "")).strip(),
    }
    summary_period = _production_cost_target_period(
        period, summary_filters["period"]
    )
    matrix_period = _production_cost_target_period(period, matrix_filters["period"])
    issue_period = _production_cost_target_period(period, issue_filters["period"])

    summary = build_work_order_cost_summary(
        summary_period, summary_filters["workshop"]
    )
    total_cost = build_total_cost_data(
        matrix_period, only_workshop=matrix_filters["workshop"]
    )
    # 一次算全量（不带"情形"筛选），再按情形筛出要显示的行：
    # ① 情形下拉/小结里的条数要显示"全部情形"下的条数，不能被当前的情形筛选影响；
    # ② 少跑一遍重活。
    include_roving_only = issue_filters["roving"] == "1"
    issue_rows_all, issue_totals_display, issue_by_case = build_issue_report_rows(
        issue_period,
        workshop_filter=issue_filters["workshop"],
        order_filter=issue_filters["order_no"],
        include_roving_only=include_roving_only,
    )
    if issue_filters["case"]:
        wanted = {issue_filters["case"]}
        if issue_filters["case"] == CASE_OTHER:
            wanted.add(CASE_NORMAL)     # ⑤ 在小结里是 CASE_OTHER、明细行上是"— 正常"
        issue_rows = [row for row in issue_rows_all if row["case"] in wanted]
    else:
        issue_rows = issue_rows_all
    issue_totals = summarize_issue_rows(issue_rows)
    # 明细表（用户 2026-09-24）：**直接复用「工单成本明细」页的同一套取数**
    # （annotate_work_order_cost_rows），所以同一张工单在这里和「工单成本」页的数字完全一致；
    # 只在最右边多一列「情形」，并把「本期投料」做成可人工修改。
    # 取数抽到 build_issue_detail_table（页面与「导出Excel」共用同一套，用户 2026-09-27）。
    detail_rows, detail_totals, order_period_missing = build_issue_detail_table(
        issue_period, issue_filters, issue_rows_all
    )

    # 情形小结（用户 2026-09-24：每个情形都要有序号、按序号排列）
    # 本期没出现的情形也列出来（工单数 0），序号才连续、一眼看出哪类没有；
    # 最后补一行「正常投料生产工单」＝"正常"的行（用户 2026-09-27 改名），
    # 这样**合计才等于明细表的行数**。
    issue_case_summary = [
        {
            "case": case,
            "count": issue_by_case[case]["count"] if issue_by_case.get(case) else 0,
            "material": issue_by_case[case]["direct_material"] if issue_by_case.get(case) else None,
            "advice": CASE_SUGGESTIONS.get(case, ""),
            "basis": CASE_COST_BASIS.get(case, ""),
            "case_link": True,
        }
        for case in CASE_CHOICES
    ]
    _normal = issue_by_case.get(CASE_NORMAL) or {"count": 0, "direct_material": Decimal("0")}
    issue_case_summary.append(
        {
            "case": CASE_OTHER,
            "count": _normal["count"],
            "material": _normal["direct_material"] if _normal["count"] else None,
            "advice": CASE_SUGGESTIONS.get(CASE_OTHER, ""),
            "basis": CASE_COST_BASIS.get(CASE_OTHER, ""),
            "case_link": True,
        }
    )
    # 最后再补一行「BOM 缺单价的工单」（用户 2026-09-30：整理出来好补单价）。
    # 注意：这一行是**标注类**，与 ①~⑤ 有重叠，不计入合计、也不进"需要人工介入"的计数。
    # ⑥ BOM 子项缺单价的工单（用户 2026-09-30）：作为**第 6 种情形**常驻一行
    #（和其它情形一样是蓝色链接、可点击筛选；已经是独立情形，不再与 ①~⑤ 重叠 → 计入合计）。
    # 用户 2026-10-15：两处修正
    #  ① 只统计**当月有动**的工单（本期投料/报工/期初在产品台账反查）
    #  ② 缺价用**实时取价表**判断 —— 以前读 BOM 行里已回填的 unit_price，
    #     而它只有在"完整结算"之后才回填 → 新月份没结算时会把**全部 BOM 行**都算成缺价。
    active_orders = set()
    try:
        active_orders |= {
            str(x).strip()
            for x in MaterialIssue.objects.filter(period=period)
            .values_list("work_order_no", flat=True) if x
        }
        active_orders |= {
            str(x).strip()
            for x in WorkReport.objects.filter(period=period)
            .values_list("work_order_no", flat=True) if x
        }
        active_orders |= {
            str(x).strip()
            for x in OpeningWip.objects.filter(period=period)
            .values_list("order_no", flat=True) if x
        }
    except Exception:  # noqa: BLE001
        active_orders = set()

    # 本月确有投料 / 报工 的工单（更贴近"当月核算"）
    current_orders = set()
    try:
        current_orders |= {
            str(x).strip()
            for x in MaterialIssue.objects.filter(period=period)
            .values_list("work_order_no", flat=True) if x
        }
        current_orders |= {
            str(x).strip()
            for x in WorkReport.objects.filter(period=period)
            .values_list("work_order_no", flat=True) if x
        }
    except Exception:  # noqa: BLE001
        current_orders = set()

    def _order_no_of(row):
        for key in ("work_order_no", "order_no", "work_order", "mo_no"):
            value = row.get(key)
            if value:
                return str(value).strip()
        return ""

    # 实时取价表：月末加权价 → 历史单价（含跨期沿用）
    try:
        live_prices = bom_material_prices(period)
    except Exception:  # noqa: BLE001
        live_prices = None

    # 工单 → 真正取不到价的子项物料
    bom_unpriced = {}
    if live_prices is not None:
        for ln in period.bom_lines.all().only("work_order_no", "sub_item_code"):
            code = (ln.sub_item_code or "").strip()
            if not code:
                continue
            try:
                priced = Decimal(live_prices.get(code) or 0) > 0
            except Exception:  # noqa: BLE001
                priced = False
            if not priced:
                bom_unpriced.setdefault((ln.work_order_no or "").strip(), set()).add(code)

    # 用户 2026-10-15：口径＝**本月确有投料或报工**的工单（上一版把"只有期初在产品"的
    #   历史单也算进来了，数字偏大）。期初在产品只是结转过来的余额，不算"当月核算"。
    if live_prices is not None:
        _unpriced_rows = [
            r for r in issue_rows
            if _order_no_of(r) in bom_unpriced
            and (not current_orders or _order_no_of(r) in current_orders)
        ]
    else:
        _unpriced_rows = [
            r for r in issue_rows
            if r.get("unpriced_codes")
            and (not active_orders or _order_no_of(r) in active_orders)
        ]

    _unpriced_codes = []
    if live_prices is not None:
        for _r in _unpriced_rows:
            for _code in sorted(bom_unpriced.get(_order_no_of(_r), ())):
                if _code not in _unpriced_codes:
                    _unpriced_codes.append(_code)
    else:
        for _r in _unpriced_rows:
            for _code in _r.get("unpriced_codes") or ():
                if _code not in _unpriced_codes:
                    _unpriced_codes.append(_code)
    _unpriced_material = sum(
        (Decimal(r.get("direct_material") or 0) for r in _unpriced_rows), Decimal("0")
    )
    issue_case_summary.append(
        {
            "case": CASE_NO_PRICE,
            "count": len(_unpriced_rows),
            "material": _unpriced_material if _unpriced_rows else None,
            "advice": CASE_SUGGESTIONS.get(CASE_NO_PRICE, ""),
            "basis": (
                # 用户 2026-10-15：这里不再罗列物料编码（太占地方），只给数量；
                #   缺的是哪些物料看下方的 BOM 明细表即可。
                "本月有投料/报工的工单里，缺单价的物料共 %d 个（明细见下方 BOM 明细表）"
                % len(_unpriced_codes)
                if _unpriced_codes
                else CASE_COST_BASIS.get(CASE_NO_PRICE, "")
            ),
            "case_link": True,
        }
    )
    issue_attention_count = sum(
        (item["count"] for item in issue_case_summary if item["case"] != CASE_OTHER), 0
    )
    # 情形下拉：6 种情况都列出来（①②③④ 的情形 + ⑤ 正常 + ⑥ 缺单价）
    # ⑤ 的计数在 issue_by_case 里记在「— 正常」上（明细行的取值），所以下面映射一下。
    def _case_count(case):
        source = CASE_NORMAL if case == CASE_OTHER else case
        return (issue_by_case.get(source) or {}).get("count", 0)

    issue_case_choices = [
        {
            "value": case,
            "label": "%s（%d 张）" % (case, _case_count(case)) if _case_count(case) else case,
        }
        for case in list(CASE_CHOICES) + [CASE_OTHER, CASE_NO_PRICE]
    ]
    # 车间下拉：完整车间清单（不受「矩阵/车间汇总」当前筛选影响，否则选了车间就换不了别的车间）
    workshop_options = _workshop_filter_options(period)
    return render(
        request,
        "settlement/production_cost_summary.html",
        {
            "period": period,
            **summary,
            "total_cost": total_cost,
            "summary_filters": summary_filters,
            "matrix_filters": matrix_filters,
            "issue_rows": issue_rows,
            "issue_totals": issue_totals,
            "detail_rows": detail_rows,
            "detail_totals": detail_totals,
            "issue_totals_display": issue_totals_display,
            "issue_by_case": issue_by_case,
            "issue_case_summary": issue_case_summary,
            "issue_attention_count": issue_attention_count,
            "issue_case_choices": issue_case_choices,
            "issue_filters": issue_filters,
            "issue_period_label": issue_period.month.strftime("%Y-%m"),
            "order_period_missing": order_period_missing,
            # 「工单期间」下拉：只列**本期这张表里真的出现过**的下单月份（新的在前），
            # 否则会列出与本期无关的老月份，选了却是 0 行（用户 2026-09-27）。
            "order_period_options": _issue_order_period_options(
                issue_period, issue_rows_all
            ),
            # 情形 × 车间（工单数）：一眼看出每类情形分别在哪些车间、各多少张
            "case_workshop": build_case_workshop_matrix(issue_rows_all, workshop_options),
            # 两张表各自的「导出Excel」（带各自筛选条件，用户 2026-09-27）
            "matrix_export_qs": _export_query(
                m_period=matrix_filters["period"], m_workshop=matrix_filters["workshop"]
            ),
            "issue_export_qs": _export_query(
                i_period=issue_filters["period"],
                i_workshop=issue_filters["workshop"],
                i_case=issue_filters["case"],
                i_order_period=issue_filters["order_period"],
                i_order_no=issue_filters["order_no"],
            ),
            "summary_period_label": summary_period.month.strftime("%Y-%m"),
            "matrix_period_label": matrix_period.month.strftime("%Y-%m"),
            "current_month": period.month.strftime("%Y-%m"),
            "period_options": _production_cost_period_options(period),
            "workshop_options": workshop_options,
            # 两张表各自「清空」时要保留对方的条件
            "summary_hidden_qs": urlencode(
                {
                    "m_%s" % key: value
                    for key, value in matrix_filters.items()
                    if value
                }
            ),
            "matrix_hidden_qs": urlencode(
                {
                    "p_%s" % key: value
                    for key, value in summary_filters.items()
                    if value
                }
            ),
        },
    )


@login_required
def issue_report_export(request, period_id):
    """「二、工单投料报工明细」单张工作表导出（用户 2026-09-27）。

    与页面**同一套取数**（build_issue_detail_table），跟着该表自己的筛选
    （i_period / i_workshop / i_case / i_order_no）走；列顺序与页面表格一致，
    含「核算期间」「工单期间」「情形」。
    """
    period = get_object_or_404(AccountingPeriod, id=period_id)
    issue_filters = _issue_filters_from_request(request)
    issue_period = _production_cost_target_period(period, issue_filters["period"])
    issue_rows_all, _totals, _by_case = build_issue_report_rows(
        issue_period,
        workshop_filter=issue_filters["workshop"],
        order_filter=issue_filters["order_no"],
        include_roving_only=issue_filters["roving"] == "1",
    )
    detail_rows, detail_totals, _missing = build_issue_detail_table(
        issue_period, issue_filters, issue_rows_all
    )
    headers = [header for header, _attr, _kind in ISSUE_DETAIL_EXPORT_COLUMNS]
    rows = [issue_detail_export_row(item) for item in detail_rows]
    # 合计行：只合计「数量/金额」列，用页面同一套 detail_totals（单价、分配率% 这类
    # 跨行相加没有意义，留空）。
    total_row = ["合计（%d 张工单）" % detail_totals["count"]] + [""] * (len(headers) - 1)
    for index, (_header, attr, _kind) in enumerate(ISSUE_DETAIL_EXPORT_COLUMNS):
        if attr in detail_totals:
            total_row[index] = detail_totals[attr]
    rows.append(tuple(total_row))
    stream = build_listing_workbook([
        ("工单投料报工明细", headers, rows),
    ])
    suffix = ""
    if issue_filters["workshop"]:
        suffix = "-%s" % WORKSHOP_LABELS.get(
            issue_filters["workshop"], issue_filters["workshop"]
        )
    filename = "工单投料报工明细-%s%s.xlsx" % (
        issue_period.month.strftime("%Y-%m"), suffix)
    response = HttpResponse(
        stream.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = "attachment; filename*=UTF-8''%s" % quote(filename)
    return response


def _workshop_filter_options(period):
    """车间下拉的候选项：**完整**车间清单（与当前筛选无关）。

    顺序＝该期间「车间结转顺序」优先、其余按固定顺序（染色→纺纱→包覆→针织→机织→浆粕）补齐，
    与「生产成本」页矩阵的车间列同源。用户 2026-09-21：以前下拉只列筛选后的车间，
    选了纺纱之后就看不到别的车间了，必须先切回「全部车间」。
    """
    from settlement.services.total_cost import (
        WORKSHOP_LABELS as _TC_LABELS,
        _workshop_order,
    )

    return [
        {"value": code, "label": _TC_LABELS.get(code, code)}
        for code in _workshop_order(period)
    ]


def _work_order_cost_period(current, selected):
    """工单成本页的「期间」筛选：默认本期；选到某个月就取那个月的结算结果。"""
    month = _parse_month(selected)
    if month:
        for item in AccountingPeriod.objects.order_by("month"):
            if (item.month.year, item.month.month) == month:
                return item
    return current


def _work_order_cost_period_options(current):
    """期间下拉：有工单结算结果的期间（本期自身不重复列）。"""
    from settlement.models import WorkOrderCostResult

    ids = set(
        WorkOrderCostResult.objects.values_list("period_id", flat=True).distinct()
    )
    ids.add(current.id)
    current_label = current.month.strftime("%Y-%m")
    return [
        {"value": item.month.strftime("%Y-%m"), "label": item.month.strftime("%Y-%m")}
        for item in AccountingPeriod.objects.filter(id__in=ids).order_by("month")
        if item.month.strftime("%Y-%m") != current_label
    ]


@login_required
def work_order_costs(request, period_id):
    """工单成本明细（生产工单+投料+报工核算结果；每个车间独立一张表，页签切换）。

    分配率口径（与分摊逻辑一致）：
    - 制造费用（水电/蒸汽/天然气/折旧/机物料/染料助剂/加工费）分配率
      = 车间该费用总额 ÷ 车间材料成本总额（每元材料承担的费用）
    - 工资分配率 = 工单折合数量（完工+0.5×在产品）÷ 车间折合数量合计
    """
    period = get_object_or_404(AccountingPeriod, id=period_id)
    # 筛选（用户 2026-09-21）：期间（默认本期，可选有结算结果的历史期间）+ 车间
    target_period = _work_order_cost_period(
        period, str(request.GET.get("w_period", "")).strip()
    )
    period = target_period
    workshop_filter = str(request.GET.get("w_workshop", "")).strip()
    results = period.work_order_cost_results.select_related("work_order")
    results = _text_filters(
        request,
        results,
        {
            "order_no": "work_order__order_no",
            "product": "product_code",
        },
    )
    if workshop_filter and workshop_filter != "浆粕车间":
        results = results.filter(
            work_order__production_dept_code=workshop_filter
        )
    elif workshop_filter == "浆粕车间":
        # 只看浆粕：没有工单结果，全部靠下面的浆粕流水行
        results = results.none()
    results = list(
        results.order_by("work_order__production_dept_code", "work_order__order_no")
    )
    openings = {
        item.order_no: item for item in period.opening_wip.all()
    }

    results = annotate_work_order_cost_rows(period, results, openings)
    # 工单期间（用户 2026-09-27）：工单成本明细也加这一列（＝下单月份），
    # 原来的「期间」列改名「核算期间」（＝这张单的成本算在哪个月）。
    attach_order_period_labels(period, results)

    # 车间合计（全量，金额列汇总）
    SUM_FIELDS = (
        "direct_material", "water_power", "steam", "natural_gas", "depreciation",
        "consumables", "dye_auxiliary", "wages", "total_cost",
        "finished_cost", "wip_cost",
    )
    subtotals = {}
    for item in results:
        ws = item.work_order.production_dept_code
        st = subtotals.setdefault(
            ws,
            {"workshop": ws, "count": 0, **{f: Decimal("0") for f in SUM_FIELDS}},
        )
        st["count"] += 1
        for f in SUM_FIELDS:
            st[f] += Decimal(getattr(item, f) or 0)

    # 每个车间独立一张表：车间顺序优先取「车间结算」里用户自定义 sequence，
    # 兜底按 WORKSHOP_LABELS 顺序；每车间含本车间行、合计行与本车间汇总条。
    from settlement.services.imports import WORKSHOP_LABELS

    present = []
    for item in results:
        ws = item.work_order.production_dept_code
        if ws not in present:
            present.append(ws)
    ordered = list(
        period.workshop_settlements.order_by("sequence").values_list(
            "workshop", flat=True
        )
    )
    if not ordered:
        ordered = list(WORKSHOP_LABELS.keys())
    present_sorted = [ws for ws in ordered if ws in present] + [
        ws for ws in present if ws not in ordered
    ]

    # 只保留一张表（用户 2026-09-21）：把所选期间/车间 的结果按车间顺序拉平
    rows = []
    for ws in present_sorted:
        rows.extend(
            item for item in results
            if item.work_order.production_dept_code == ws
        )

    # 浆粕车间（没有工单）：按上传的浆粕流水出「一行一个成品」，排在最后。
    # ⚠ 2026-09-30：浆粕行没有工单编号、也不走工单查询，
    #   所以按「工单编号」筛选时必须排除（否则筛了单号它还跟着显示）；
    #   按「产品编码」筛选时对它的产品编码/名称也做同样的包含匹配。
    order_no_filter = str(request.GET.get("order_no", "")).strip()
    product_filter = str(request.GET.get("product", "")).strip()

    if not workshop_filter or workshop_filter == "浆粕车间":
        from settlement.services.workorder_costing import build_pulp_cost_rows

        pulp_rows = build_pulp_cost_rows(period)
        if order_no_filter:
            pulp_rows = []
        elif product_filter:
            keyword = product_filter.lower()
            pulp_rows = [
                r
                for r in pulp_rows
                if keyword in str(getattr(r, "product_code", "") or "").lower()
                or keyword in str(getattr(r, "product_name", "") or "").lower()
            ]
        rows.extend(pulp_rows)

    # 合计行：当前筛选下的合计
    sum_fields = (
        "direct_material", "water_power", "steam", "natural_gas", "depreciation",
        "consumables", "dye_auxiliary", "processing_fee", "wages", "total_cost",
        "finished_cost", "wip_cost",
    )
    subtotal = {"count": len(rows)}
    for field in sum_fields:
        subtotal[field] = _cost_money(
            sum((Decimal(getattr(r, field) or 0) for r in rows), Decimal("0"))
        )

    # 合计行的单位成本对照（用户 2026-09-22）：标准/实际都按「算得出标准成本」的工单的
    # 完工数量加权，免得算不出标准成本的工单（标准为 —）把加权拉低。
    _std_qty = Decimal("0")
    _std_amount = Decimal("0")
    _act_amount = Decimal("0")
    for item in rows:
        _unit = getattr(item, "standard_unit_cost", None)
        _qty = Decimal(getattr(item, "finished_quantity", 0) or 0)
        if _unit and _qty > 0:
            _std_qty += _qty
            _std_amount += Decimal(_unit) * _qty
            _act_amount += Decimal(getattr(item, "finished_cost", 0) or 0)
    subtotal["standard_unit_cost"] = (_std_amount / _std_qty) if _std_qty else None
    subtotal["unit_cost_diff"] = (
        ((_act_amount - _std_amount) / _std_qty) if _std_qty else None
    )
    # 单位成本差异率（%，用户 2026-09-24）：差异 ÷ 标准单位成本 × 100
    subtotal["unit_cost_diff_pct"] = (
        subtotal["unit_cost_diff"] / subtotal["standard_unit_cost"] * Decimal("100")
        if subtotal["standard_unit_cost"]
        else None
    )
    subtotal["unit_cost_diff_class"] = (
        "woc-diff-bad"
        if subtotal["unit_cost_diff"] and subtotal["unit_cost_diff"] > 0
        else "woc-diff-good"
        if subtotal["unit_cost_diff"] and subtotal["unit_cost_diff"] < 0
        else ""
    )

    # 车间下拉：用**完整**车间清单（不受当前筛选影响，否则选了车间就换不了别的车间）
    workshop_options = _workshop_filter_options(period)

    return render(
        request,
        "settlement/work_order_costs.html",
        {
            "period": period,
            "rows": rows,
            "subtotal": subtotal,
            "row_count": len(rows),
            "period_label": period.month.strftime("%Y-%m"),
            "current_month": target_period.month.strftime("%Y-%m"),
            "period_options": _work_order_cost_period_options(
                get_object_or_404(AccountingPeriod, id=period_id)
            ),
            "workshop_options": workshop_options,
            "w_filter": {
                "workshop": workshop_filter,
                "period": str(request.GET.get("w_period", "")).strip(),
            },
            # 导出按「所选期间」导出：链接用该期间 id（导出分支按 URL 的期间取数）
            "export_period_id": period.id,
        },
    )


# ---------------------------------------------------------------------------
# 标准成本（用户 2026-09-22）：按车间录入「人工费用占比 / 制造费用占比」，
# 用 BOM 口径算出标准成本，与本页实际成本并排对比。
# ---------------------------------------------------------------------------


def standard_cost_workshops(period):
    """标准成本页的车间清单：与「工单成本」页下拉同源（结转顺序优先，其余按固定顺序补齐）。"""
    return [item["value"] for item in _workshop_filter_options(period)]


def effective_standard_cost_rates(period):
    """本期标准成本占比；本期没录过就沿用上一个录过的期间（省得每月重录）。

    返回 {workshop: StandardCostRate}（沿用来的对象也会带上，模板里据此标注「沿用 X 月」）。
    """
    current = {item.workshop: item for item in period.standard_cost_rates.all()}
    if current:
        return current, None
    carried_source = (
        AccountingPeriod.objects.filter(
            month__lt=period.month, standard_cost_rates__isnull=False
        )
        .order_by("-month")
        .first()
    )
    if not carried_source:
        return {}, None
    carried = {
        item.workshop: item for item in carried_source.standard_cost_rates.all()
    }
    return carried, carried_source


def standard_unit_cost_by_order(period):
    """工单编号 → 该工单的 BOM「标准成品单价」＝Σ(子项用量 × 子项单价)。

    结算后由系统回填（services/costing.py::_fill_bom_prices），同一工单的多行写的是
    同一个汇总值。标准价为空或算成 0 的工单不会出现在结果里，调用方据此显示「—」。
    """
    values = {}
    for line in period.bom_lines.exclude(standard_price=None):
        if line.work_order_no:
            values[line.work_order_no] = Decimal(line.standard_price or 0)
    return values


def standard_unit_cost_factor(period, workshop, rates=None):
    """该车间的标准成本加成系数 ＝ 1 ＋ 人工占比 ＋ 制造费用占比。

    占比来自「半年滚动」：上一个完整半年**各月占比的平均数**（用户 2026-09-23，见
    rolling_standard_cost_rates）；基期半年一条数据都没有时，退回本页手工录入的参数。
    """
    if rates is None:
        rates, _meta = rolling_standard_cost_rates(period)
    return standard_cost_factor(rates, workshop)


# ---------------------------------------------------------------------------
# 半年滚动的标准占比（用户 2026-09-23）
#   口径：某期间的标准占比 ＝ **上一个完整半年**的实际占比平均值
#         （7-12 月的工单 → 用当年 1-6 月的实际平均；次年 1-6 月 → 用上年 7-12 月的实际平均），
#         即每半年滚动一次。
#   数据来源：该半年内每个月的实际发生数（工单成本结果：直接材料/工资/制造费用）；
#         系统里没有的月份（2026 年 1-6 月）用「标准成本构成表」手工录入的数。
# ---------------------------------------------------------------------------


def month_key(value):
    """把日期/期间统一成「该月 1 日」的 date。"""
    if hasattr(value, "day"):
        return date(value.year, value.month, 1)
    return value


def half_year_before(month):
    """某个期间所属的「上一个完整半年」＝半年滚动的基期。

    7-12 月 → 当年 1-6 月；1-6 月 → 上一年 7-12 月。
    """
    month = month_key(month)
    if month.month <= 6:
        return (
            date(month.year - 1, 7, 1),
            date(month.year - 1, 12, 1),
            "%d年下半年" % (month.year - 1),
        )
    return date(month.year, 1, 1), date(month.year, 6, 1), "%d年上半年" % month.year


def months_between(start, end):
    """[start, end] 之间的月份列表（每月 1 日）。"""
    months = []
    year, month = start.year, start.month
    while (year, month) <= (end.year, end.month):
        months.append(date(year, month, 1))
        month += 1
        if month > 12:
            year, month = year + 1, 1
    return months


def actual_composition_by_month(start, end):
    """{月份: {车间: {产量/直接材料/人工/制造费用}}} —— 来自工单成本结果的实际发生数。"""
    from settlement.models import WorkOrderCostResult

    data = defaultdict(lambda: defaultdict(lambda: {
        "output_qty": Decimal("0"),
        "direct_material": Decimal("0"),
        "wages": Decimal("0"),
        "manufacturing_overhead": Decimal("0"),
    }))
    results = (
        WorkOrderCostResult.objects.filter(
            period__month__gte=start, period__month__lte=end
        )
        .select_related("work_order")
    )
    for item in results:
        workshop = item.work_order.production_dept_code
        cell = data[month_key(item.period.month)][workshop]
        cell["output_qty"] += Decimal(item.finished_quantity or 0)
        cell["direct_material"] += Decimal(item.direct_material or 0)
        cell["wages"] += Decimal(item.wages or 0)
        cell["manufacturing_overhead"] += _manaoverhead_of(item)
    return data


def manual_composition_by_month(start, end):
    """{月份: {车间: StandardCostCompositionEntry}} —— 构成表手工录入的数。"""
    from settlement.models import StandardCostCompositionEntry

    data = defaultdict(dict)
    entries = StandardCostCompositionEntry.objects.filter(
        period_month__gte=start, period_month__lte=end
    )
    for entry in entries:
        data[month_key(entry.period_month)][entry.workshop] = entry
    return data


def resolve_composition_pcts(cell):
    """某月某车间的三个占比（%）：一律按金额算（各自 ÷ 三者合计）。

    用户 2026-09-23：占比不再手工填（上半年录金额、占比自动算），所以这里只按金额算；
    金额都为空时返回 (None, None, None)。
    """
    if cell is None:
        return None, None, None
    base = (
        Decimal(cell.get("direct_material") or 0)
        + Decimal(cell.get("wages") or 0)
        + Decimal(cell.get("manufacturing_overhead") or 0)
    )
    if not base:
        return None, None, None
    hundred = Decimal("100")
    return (
        Decimal(cell.get("direct_material") or 0) / base * hundred,
        Decimal(cell.get("wages") or 0) / base * hundred,
        Decimal(cell.get("manufacturing_overhead") or 0) / base * hundred,
    )


def composition_cells_by_month(start, end):
    """{月份: {车间: cell}}：实际数优先，实际没有该车间数据的才用手工录入的数。

    cell 里带 `source`（actual / manual）和 `entry`（手工录入对象，可能为空）。
    """
    actuals = actual_composition_by_month(start, end)
    manual = manual_composition_by_month(start, end)
    cells = {}
    for month in months_between(start, end):
        per_workshop = {}
        for workshop, cell in (actuals.get(month) or {}).items():
            per_workshop[workshop] = dict(cell, source="actual", entry=None)
        for workshop, entry in (manual.get(month) or {}).items():
            if workshop in per_workshop:
                continue  # 系统有实际数据 → 实际优先
            per_workshop[workshop] = {
                "output_qty": entry.output_qty,
                "direct_material": entry.direct_material,
                "wages": entry.wages,
                "manufacturing_overhead": entry.manufacturing_overhead,
                "material_pct": entry.material_pct,
                "labor_pct": entry.labor_pct,
                "overhead_pct": entry.overhead_pct,
                "source": "manual",
                "entry": entry,
            }
        cells[month] = per_workshop
    return cells


def parse_month_param(raw):
    """把 "YYYY-MM" 形式的筛选值解析成该月 1 日；不合法返回 None。"""
    raw = str(raw or "").strip()
    if not raw:
        return None
    try:
        parts = raw.split("-")
        return date(int(parts[0]), int(parts[1]), 1)
    except (ValueError, IndexError):
        return None


def rolling_standard_cost_rates(period, month=None):
    """本期适用的标准占比（半年滚动）→ ({车间: 占比信息}, 基期信息)。

    month 给定时按那个月算基期；不给就用页面期间的月份。

    **口径（用户 2026-09-23 定的）**：
      * 取基期半年里**每个月**的「材料/人工/制造费用」占比（各自 ÷ 三者合计），
        再求**算术平均**——也就是「1-6 月各月占比的平均数」。
      * 标准成本公式：**标准成本 ＝ BOM 标准材料成本 ×（1 ＋ 人工占比 ＋ 制造费用占比）**。
        页面「适用的标准占比」表直接给出「人工占比和费用占比的和」＝ 人工占比 ＋ 制造费用占比。

    占比信息含：材料/人工/制造费用占比（%，各月平均）、人工占比＋制造费用占比之和、
    来源（rolling＝半年实际；manual＝基期没数据、退回本页手工参数）、用了几个月的数。
    """
    target = month_key(month) if month else month_key(period.month)
    start, end, label = half_year_before(target)
    cells = composition_cells_by_month(start, end)

    hundred = Decimal("100")
    # 车间 -> 各国每月的（材料%, 人工%, 制造费用%）
    monthly = defaultdict(list)
    for month in months_between(start, end):
        for workshop, cell in (cells.get(month) or {}).items():
            values = (
                Decimal(cell.get("direct_material") or 0),
                Decimal(cell.get("wages") or 0),
                Decimal(cell.get("manufacturing_overhead") or 0),
            )
            base = values[0] + values[1] + values[2]
            if not base:
                continue
            monthly[workshop].append(tuple(v / base * hundred for v in values))

    rates = {}
    for workshop, rows in monthly.items():
        count = Decimal(len(rows))
        material_pct = sum((r[0] for r in rows), Decimal("0")) / count
        labor_pct = sum((r[1] for r in rows), Decimal("0")) / count
        overhead_pct = sum((r[2] for r in rows), Decimal("0")) / count
        rates[workshop] = {
            "material_pct": material_pct,
            "labor_pct": labor_pct,
            "overhead_pct": overhead_pct,
            # 公式直接用：加成比 ＝ 各占比（占成本合计）；系数 ＝ 1 ＋ 人工 ＋ 费用
            "labor_ratio": labor_pct / hundred,
            "overhead_ratio": overhead_pct / hundred,
            "labor_plus_overhead_pct": labor_pct + overhead_pct,
            "months": len(rows),
            "source": "rolling",
        }

    # 兜底：基期半年里该车间一点数据都没有 → 沿用本页手工录入的占比参数
    fallback, carried_from = effective_standard_cost_rates(period)
    for workshop in standard_cost_workshops(period):
        if workshop in rates:
            continue
        rate = fallback.get(workshop)
        if not rate:
            continue
        labor = Decimal(rate.labor_rate_pct or 0)
        overhead = Decimal(rate.overhead_rate_pct or 0)
        if labor + overhead >= hundred:
            continue
        rates[workshop] = {
            "material_pct": hundred - labor - overhead,
            "labor_pct": labor,
            "overhead_pct": overhead,
            "labor_ratio": labor / hundred,
            "overhead_ratio": overhead / hundred,
            "labor_plus_overhead_pct": labor + overhead,
            "months": 0,
            "source": "manual",
        }

    return rates, {
        "basis_start": start,
        "basis_end": end,
        "basis_label": label,
        "carried_from": carried_from,
    }


def standard_cost_factor(rates, workshop):
    """加成系数 ＝ 1 ＋ 人工占比 ＋ 制造费用占比（占成本合计）。

    用户 2026-09-23：标准成本 ＝ BOM 标准材料成本 × 该系数。
    """
    rate = rates.get(workshop)
    if not rate:
        return Decimal("1")
    return (
        Decimal("1") + Decimal(rate["labor_ratio"]) + Decimal(rate["overhead_ratio"])
    )



def attach_standard_unit_costs(period, items):
    """给工单成本结果行挂上「标准单位成本」「单位成本差异」（页面/导出共用，保证一致）。

    用户 2026-09-22：标准侧只需要**单位标准成本**，实际成本以「工单成本」结果为准，
    所以这里只给每张工单算一个单位标准成本，拿来与实际单位成本对照：
      标准单位成本 ＝ BOM 标准成品单价 ×（1 ＋ 车间人工占比 ＋ 车间制造费用占比）
      单位成本差异 ＝ 实际单位成本（完工成本÷完工数量）− 标准单位成本（正数＝超支）
    用户 2026-09-24：单位成本差异改成**差异率百分比**（÷ 标准单位成本 × 100），
    所以同时给出 unit_cost_diff（金额）与 unit_cost_diff_pct（百分数）。
    工单没有可用的标准成品单价（BOM 子项原料当月没有月末加权价）时三者皆为 None。
    """
    standard_unit_by_order = standard_unit_cost_by_order(period)
    # 单位成本差异原因（用户 2026-09-30）：要看出"为什么差"，例如 BOM 两种料只投了一种。
    #   只取"一点都没投"的 BOM 子项（粗纱 BC- 除外：它走粗纱链，不算漏投）。
    bom_codes_by_order = defaultdict(list)
    for line in period.bom_lines.all():
        if line.work_order_no and not str(line.sub_item_code).upper().startswith("BC-"):
            bom_codes_by_order[line.work_order_no].append(
                (line.sub_item_code, line.sub_item_name)
            )
    issued_by_pair = defaultdict(Decimal)
    for row in period.material_issues.exclude(work_order_no="").exclude(material_code="").values(
        "work_order_no", "material_code", "actual_issue_quantity"
    ):
        issued_by_pair[(row["work_order_no"], row["material_code"])] += Decimal(
            row["actual_issue_quantity"] or 0
        )
    factors = {}
    for item in items:
        workshop = item.work_order.production_dept_code
        unit_material = standard_unit_by_order.get(item.work_order.order_no)
        item.standard_unit_cost = None
        item.unit_cost_diff = None
        item.unit_cost_diff_pct = None
        item.unit_cost_diff_class = ""
        item.unit_cost_diff_reason = ""
        if not unit_material:
            continue
        if workshop not in factors:
            factors[workshop] = standard_unit_cost_factor(period, workshop)
        unit_cost = Decimal(unit_material) * factors[workshop]
        item.standard_unit_cost = unit_cost
        if Decimal(item.finished_quantity or 0) > 0:
            diff = Decimal(item.finished_unit_cost or 0) - unit_cost
            item.unit_cost_diff = diff
            # 差异率（%）：单位成本差异 ÷ 标准单位成本 × 100；标准为 0 时留空
            if unit_cost:
                item.unit_cost_diff_pct = (
                    diff / unit_cost * Decimal("100")
                )
                item.unit_cost_diff_pct_text = (
                    f"{item.unit_cost_diff_pct:.2f}%"
                    if item.unit_cost_diff_pct is not None else "—"
                )
            item.unit_cost_diff_class = (
                "woc-diff-bad" if diff > 0 else "woc-diff-good" if diff < 0 else ""
            )
        order_no = item.work_order.order_no
        missing_codes = [
            (code, name)
            for code, name in bom_codes_by_order.get(order_no, [])
            if issued_by_pair.get((order_no, code), Decimal("0")) <= 0
        ]
        # 用户 2026-09-30：把"少投/超投的是哪种原料、差多少数量"也算出来（简写进原因）
        finish_qty = Decimal(item.finished_quantity or 0)
        short_items = []
        over_items = []
        if finish_qty > 0:
            for line in period.bom_lines.filter(work_order_no=order_no):
                code = line.sub_item_code
                if str(code).upper().startswith("BC-"):
                    continue  # 粗纱走粗纱链，不算漏投
                need_qty = Decimal(line.qty or 0) * finish_qty
                if need_qty <= 0:
                    continue
                issued_qty = issued_by_pair.get((order_no, code), Decimal("0"))
                diff_qty = need_qty - issued_qty
                uom = line.uom or ""
                if diff_qty > Decimal("0.01"):
                    short_items.append((code, line.sub_item_name, diff_qty, uom))
                elif diff_qty < Decimal("-0.01"):
                    over_items.append((code, line.sub_item_name, -diff_qty, uom))
            short_items.sort(key=lambda x: -x[2])
            over_items.sort(key=lambda x: -x[2])
        # 用户 2026-10-16：**反推材料成本** = 材料成本 − 实际投料金额（>0 才有反推）
        _actual_issue_amt = Decimal(0)
        try:
            _pmap2 = bom_material_prices(period)
        except Exception:  # noqa: BLE001
            _pmap2 = {}
        for (_o2, _c2), _q2 in issued_by_pair.items():
            if _o2 != order_no:
                continue
            _pr2 = _pmap2.get(_c2) or Decimal(0)
            _actual_issue_amt += Decimal(_q2 or 0) * Decimal(_pr2)
        _mat2 = Decimal(getattr(item, "direct_material", 0) or 0)
        item.actual_issue_material = _actual_issue_amt
        item.backfill_material = (_mat2 - _actual_issue_amt) if _mat2 > _actual_issue_amt else Decimal(0)
        # 用户 2026-10-16：**投料差异**单独一列、永远显示（不依赖差异率是否 ≥30%），
        #   否则"报工超投料→按 BOM 反推"会把少投的料补平，页面上看不出来。
        # 用户 2026-10-16：**互斥**——一点没投的只写「漏投」，投了不够的只写「少投」，
        #   同一个物料不再同时出现在两类里；数字统一用（）包起来。
        _miss_codes = {c for c, _n in missing_codes}
        _parts = []
        if short_items:
            _miss_txt = []
            _short_txt = []
            for c, _n, q, u in short_items[:4]:
                if c in _miss_codes:
                    _miss_txt.append("漏投 %s（%s%s）" % (c, ("%g" % float(q)), u or ""))
                else:
                    _short_txt.append("少投 %s（%s%s）" % (c, ("%g" % float(q)), u or ""))
            if _miss_txt:
                _parts.append("、".join(_miss_txt))
            if _short_txt:
                _parts.append("、".join(_short_txt))
        if over_items:
            _parts.append("多投 %s" % "、".join(
                "%s（%s%s）" % (c, ("%g" % float(q)), u or "") for c, _n, q, u in over_items[:3]
            ))
        if _parts:
            item.issue_diff_text = "；".join(_parts)
        else:
            # 没有任何可比对的差异时，别一律写"一致"（会把"没投料"说成一致 ✗）
            _has_issue = any(
                issued_by_pair.get((order_no, _c), Decimal("0")) > 0
                for _c in {ln.sub_item_code for ln in period.bom_lines.filter(
                    work_order_no=order_no)}
            )
            _cmp = 0
            for _ln in period.bom_lines.filter(work_order_no=order_no):
                if not str(_ln.sub_item_code or "").upper().startswith("BC-"):
                    _cmp += 1
            if not _has_issue and finish_qty > 0:
                item.issue_diff_text = "未投料（有报工，材料按 BOM 反推）"
            elif _cmp == 0:
                item.issue_diff_text = "不适用（BOM 只有粗纱/无可比子项）"
            elif finish_qty <= 0:
                item.issue_diff_text = "无完工（未报工，无法比对）"
            else:
                item.issue_diff_text = "一致"
        # 用户 2026-10-16：情形③（有投料-报工超投料量）也会**按报工量 × BOM 反推补足**材料，
        #   在"投料差异"里说明，避免"实际投料不够却看不到"。
        try:
            _case = getattr(item, "case", "") or ""
            if ("报工超投料量" in _case) and getattr(item, "issue_diff_text", ""):
                item.issue_diff_text = "%s（材料按报工量×BOM 反推补足）" % item.issue_diff_text
        except Exception:  # noqa: BLE001
            pass
        item.unit_cost_diff_reason = unit_cost_diff_reason(
            item, unit_material, missing_codes,
            short_items=short_items, over_items=over_items,
            has_issue_ledger=any(
                key[0] == order_no for key in issued_by_pair
            ),
        )
    return items


def _manaoverhead_of(item):
    """工单实际制造费用 = 水电+蒸汽+天然气+折旧+机物料+染料助剂+加工费。"""
    from settlement.services.workorder_costing import MANUFACTURING_FIELDS

    return sum(
        (Decimal(getattr(item, field) or 0) for field in MANUFACTURING_FIELDS),
        Decimal("0"),
    )


def build_standard_cost_rows(
    period, workshop_filter="", order_filter="", product_filter=""
):
    """标准成本行（页面与导出共用，保证两处完全一致）。

    用户 2026-09-22：本函数只算**标准侧**（实际成本在「工单成本」结果里已有，不在这里
    重复算、也不在这里比较）：标准材料单位成本＝BOM「标准成品单价」＝Σ(子项用量×子项单价)；
    标准材料成本＝标准材料单位成本×完工数量；标准人工、标准制造费用＝标准材料成本×各自占比；
    标准单位成本＝标准材料单位成本×(1＋人工占比＋制造费用占比)。

    用户 2026-09-23：占比改为**半年滚动**——取上一个完整半年**各月占比的平均数**
    （见 rolling_standard_cost_rates），不再按月手工录。

    用户 2026-09-23（产品编码/名称）：结算结果没带产品的工单（本期间无报工，
    多为完工=0 的单）原来编码/名称空着，现在用与「期末在产品」页同一套解析器
    （build_product_resolver：报工 → BOM 成品物料 → 工单标识码 → 工单号）兜底显示；
    产品筛选也按兜底后的编码/名称匹配。
    """
    rates, rates_meta = rolling_standard_cost_rates(period)
    carried_from = rates_meta["carried_from"]
    results = period.work_order_cost_results.select_related("work_order")
    if workshop_filter and workshop_filter != PULP_WORKSHOP:
        results = results.filter(work_order__production_dept_code=workshop_filter)
    elif workshop_filter == PULP_WORKSHOP:
        # 浆粕车间没有工单、也没有 BOM，标准成本侧为空（页面上单独说明）
        results = results.none()
    if order_filter:
        results = results.filter(work_order__order_no__icontains=order_filter)
    results = list(results)

    # 产品编码/名称兜底（与「期末在产品」页同一套解析器）：报工 → BOM 成品物料 → 标识码 → 工单号
    resolve_product = build_product_resolver(period)

    # 该工单的 BOM 标准成品单价（单位成品的标准材料成本）
    standard_unit_by_order = standard_unit_cost_by_order(period)

    rows = []
    no_standard_price = 0
    for item in results:
        order_no = item.work_order.order_no
        workshop = item.work_order.production_dept_code
        product_code, product_name = resolve_product(
            item.work_order, item.product_code, item.product_name
        )
        # 产品筛选：按兜底后的编码/名称匹配（原来只查结算结果的编码，空编码的单永远筛不到）
        if product_filter:
            haystack = ("%s %s" % (product_code, product_name)).lower()
            if product_filter.lower() not in haystack:
                continue
        rate = rates.get(workshop) or {}
        labor_ratio = Decimal(rate.get("labor_ratio") or 0)
        overhead_ratio = Decimal(rate.get("overhead_ratio") or 0)
        labor_pct = Decimal(rate.get("labor_pct") or 0)
        overhead_pct = Decimal(rate.get("overhead_pct") or 0)

        finished_qty = Decimal(item.finished_quantity or 0)
        unit_material = standard_unit_by_order.get(order_no)
        has_standard_price = unit_material is not None and unit_material != 0
        if has_standard_price:
            standard_material = Decimal(unit_material) * finished_qty
            standard_labor = standard_material * labor_ratio
            standard_overhead = standard_material * overhead_ratio
            standard_total = standard_material + standard_labor + standard_overhead
            # 单位标准成本＝标准材料单位成本×(1+两加成比)：与完工数量无关（线性缩放）
            standard_unit_cost = (
                Decimal(unit_material)
                * (Decimal("1") + labor_ratio + overhead_ratio)
            )
        else:
            # 注意：不是「没有 BOM」。绝大多数是 BOM 有、但子项原料当月没有月末加权价
            # → Σ(用量×单价) 算出来是 0（用户 2026-09-22 核实：90 张全部是这种）。
            no_standard_price += 1
            standard_material = standard_labor = standard_overhead = None
            standard_total = standard_unit_cost = None

        rows.append(
            {
                "workshop": workshop,
                "workshop_label": WORKSHOP_LABELS.get(workshop, workshop or "—"),
                "order_no": order_no,
                "product_code": product_code,
                "product_name": product_name,
                "finished_quantity": finished_qty,
                "has_bom": has_standard_price,
                "has_standard_price": has_standard_price,
                "labor_rate_pct": labor_pct,
                "overhead_rate_pct": overhead_pct,
                "standard_unit_material": Decimal(unit_material) if has_standard_price else None,
                "standard_material": standard_material,
                "standard_labor": standard_labor,
                "standard_overhead": standard_overhead,
                "standard_total": standard_total,
                "standard_unit_cost": standard_unit_cost,
            }
        )

    rows.sort(key=lambda r: (r["workshop_label"], r["order_no"]))
    return rows, {
        "no_standard_price": no_standard_price,
        "carried_from": carried_from,
        "rates": rates,
        "basis_label": rates_meta["basis_label"],
        "basis_start": rates_meta["basis_start"],
        "basis_end": rates_meta["basis_end"],
    }


def bom_detail_head_map(period):
    """工单编号 → 「BOM 明细表」的行头信息（车间 / 产品编码 / 产品名称）。

    产品编码/名称用与「工单标准成本计算表」同一套兜底解析（build_product_resolver），
    两表显示一致；「BOM 明细表」自己的期间/车间/工单号筛选就按这张表过滤。

    用户 2026-09-30：本表要覆盖**本期全部 BOM 行**（原来只列有结算结果的工单，9944 行只剩 992 行），
    所以「本期有 BOM 行、但没有结算结果」的工单也要建上行头：
      · 车间：工单主档的生产部门（没有主档就留空）；
      · 产品编码/名称：同一套兜底（BOM 成品物料 → 工单标识码前缀 → 工单号）。
    """
    resolve_product = build_product_resolver(period)
    heads = {}
    for item in period.work_order_cost_results.select_related("work_order"):
        order_no = item.work_order.order_no
        if not order_no:
            continue
        workshop = item.work_order.production_dept_code
        product_code, product_name = resolve_product(
            item.work_order, item.product_code, item.product_name
        )
        heads[order_no] = {
            "order_no": order_no,
            "workshop": workshop,
            "workshop_label": WORKSHOP_LABELS.get(workshop, workshop or "—"),
            "product_code": product_code,
            "product_name": product_name,
        }

    class _OrderStub:
        """没有工单主档的 BOM 工单：resolver 只用到 order_no / identifier。"""

        __slots__ = ("order_no", "identifier")

        def __init__(self, order_no, identifier=""):
            self.order_no = order_no
            self.identifier = identifier

    masters = {wo.order_no: wo for wo in period.work_orders.all() if wo.order_no}
    for order_no in (
        period.bom_lines.values_list("work_order_no", flat=True).distinct()
    ):
        if not order_no or order_no in heads:
            continue
        master = masters.get(order_no)
        workshop = master.production_dept_code if master is not None else ""
        product_code, product_name = resolve_product(
            master if master is not None else _OrderStub(order_no), "", ""
        )
        heads[order_no] = {
            "order_no": order_no,
            "workshop": workshop,
            "workshop_label": WORKSHOP_LABELS.get(workshop, workshop or "—"),
            "product_code": product_code,
            "product_name": product_name,
        }
    return heads


#: 「BOM 明细」源列取值的字段（＝ BomLine 上与上传文件一一对应的字段）
BOM_SOURCE_FIELDS = (
    "work_order_no",
    "material_code",
    "material_name",
    "line_no",
    "sub_item_code",
    "sub_item_name",
    "qty_numerator",
    "qty_denominator",
    "uom",
    "unit_price",
    "standard_price",
    "loss_rate",
    "standard_qty",
    "required_qty",
)

#: 右对齐（数值）的源字段
_BOM_NUMERIC_FIELDS = frozenset(
    {
        "qty_numerator",
        "qty_denominator",
        "unit_price",
        "standard_price",
        "loss_rate",
        "standard_qty",
        "required_qty",
    }
)

#: 系统列（追加在上传原样列之后）：(列名, 行字典键, 页面小数位)
#: 「BOM表台账/导出」不带「金额」；「标准成本页 BOM 明细表」带（该表口径＝金额＝用量×单价）
BOM_LEDGER_SYSTEM_COLUMNS = (
    ("用量", "qty", 2),
    ("用量(转100%)", "qty_normalized", 2),
    ("转换后子项用量:分子", "qty_converted_numerator", 2),
    ("单价", "unit_price", 4),
    ("标准成品单价", "standard_price", 2),
)
BOM_DETAIL_SYSTEM_COLUMNS = (
    ("用量", "qty", 2),
    ("用量(转100%)", "qty_normalized", 2),
    ("转换后子项用量:分子", "qty_converted_numerator", 2),
    ("单价", "unit_price", 4),
    ("金额", "amount", 2),
    ("标准成品单价", "standard_price", 2),
)

#: 标准成本页「BOM 明细表」每页行数（用户 2026-09-30：本表列本期全部 BOM 行，
#: 近万行、40 列一起渲染页面会到十几 MB，所以页面分页；导出 Excel 仍是全部行）
BOM_DETAIL_PAGE_SIZE = 200


def bom_source_header(period):
    """用户上传的 BOM 文件表头（列名/列序原样）；该期间没上传过 BOM 文件时返回 None。

    选文件的口径与「导出中心 → BOM表」一致（行数最多的那个批次，行数取导入时记的 row_count），
    但**只读表头所在的前几行**——近万行的源文件不必整表读一遍（页面每次渲染都要用列定义）。
    """
    import csv as _csv
    import io as _io

    batch = (
        period.import_batches.filter(kind=ImportBatch.Kind.BOM)
        .exclude(file="")
        .order_by("-row_count", "-created_at", "-id")
        .first()
    )
    if batch is None:
        return None
    path = Path(batch.file.path)
    if not path.is_file():
        return None
    suffix = path.suffix.lower()
    head_rows = []
    try:
        if suffix in (".xlsx", ".xlsm"):
            from openpyxl import load_workbook

            workbook = load_workbook(path, read_only=True, data_only=True)
            try:
                for index, row in enumerate(workbook.worksheets[0].iter_rows(values_only=True)):
                    head_rows.append(list(row))
                    if index >= 14:
                        break
            finally:
                workbook.close()
        elif suffix == ".csv":
            raw = path.read_bytes()
            for encoding in ("utf-8-sig", "gb18030", "utf-8"):
                try:
                    text = raw.decode(encoding)
                    break
                except UnicodeDecodeError:
                    continue
            else:
                return None
            for index, row in enumerate(_csv.reader(_io.StringIO(text))):
                head_rows.append(list(row))
                if index >= 14:
                    break
    except Exception:
        return None
    if not head_rows:
        return None
    alias = bom_alias_field_map()
    located = None
    for index, row in enumerate(head_rows):
        hits = sum(1 for value in row if normalize_header(value) in alias)
        if located is None or hits > located[0]:
            located = (hits, index)
    if located is None or located[0] < 2:
        return None
    return [str(name) for name in head_rows[located[1]]]


def bom_alias_field_map():
    """归一化列名 → BomLine 字段名（与导入用的同一套 ALIASES）。"""
    mapping = {}
    for field, names in ALIASES["bom"].items():
        for name in names:
            mapping[normalize_header(name)] = field
    return mapping


def bom_standard_qty_basis(period):
    """工单编号 → Σ标准用量，以及「标准用量有缺失」的工单集合。

    用户 2026-09-30：「转换后子项用量:分子」的**分子用标准用量**（＝标准用量 ÷ 该工单 Σ标准用量）；
    老数据（没导入标准用量）的工单退回用 用量 ÷ Σ用量。
    """
    totals = defaultdict(Decimal)
    missing = set()
    for order_no, standard_qty in period.bom_lines.values_list(
        "work_order_no", "standard_qty"
    ):
        if not order_no:
            continue
        if standard_qty is None:
            missing.add(order_no)
        else:
            totals[order_no] += standard_qty
    return totals, missing


def bom_share_resolver(period):
    """返回 share(line) → (用量占比, 转换后占比)，页面/导出共用同一套口径。

    * 用量占比（「用量(转100%)」列）＝ 用量（分子÷分母）÷ 该工单 Σ用量；
    * 转换后占比（「转换后子项用量:分子」列，用户 2026-09-30 更正）
      ＝ **标准用量 ÷ 该工单 Σ标准用量**（每个物料 ÷ 该工单标准用量之和，合计＝1＝100%）；
      该工单有行缺标准用量时，退回 用量 ÷ Σ用量。
    """
    qty_totals = bom_qty_total_by_order(period)
    std_totals, std_missing = bom_standard_qty_basis(period)

    def share(line):
        order_no = line.work_order_no
        qty_ratio = bom_normalized_qty(line.qty, qty_totals.get(order_no))
        std_total = std_totals.get(order_no) or Decimal("0")
        if (
            line.standard_qty is not None
            and std_total > 0
            and order_no not in std_missing
        ):
            converted = line.standard_qty / std_total
        else:
            converted = qty_ratio
        return qty_ratio, converted

    return share


def bom_row_from_line(line, qty_total, share=None):
    """BomLine → 「BOM 明细」行字典（页面/导出共用，折算口径只在这里算一次）。

    share 由 bom_share_resolver(period) 给出；不给时「转换后」退回用量口径（老调用方/测试）。
    """
    qty = line.qty
    total = qty_total.get(line.work_order_no) or Decimal("0")
    if share is not None:
        qty_ratio, converted_ratio = share(line)
    else:
        qty_ratio = bom_normalized_qty(qty, total)
        converted_ratio = qty_ratio
    return {
        "id": line.id,
        "order_no": line.work_order_no,
        "workshop_label": "",
        "product_code": line.material_code,
        "product_name": line.material_name,
        "line_no": line.line_no,
        "sub_item_code": line.sub_item_code,
        "sub_item_name": line.sub_item_name,
        "qty_numerator": line.qty_numerator,
        "qty_denominator": line.qty_denominator,
        "qty": qty,
        "qty_total": total,
        "qty_is_scaled": bom_qty_is_scaled(total),
        "qty_normalized": qty_ratio,
        "qty_converted_numerator": converted_ratio,
        "uom": line.uom,
        "unit_price": line.unit_price,
        "amount": qty * Decimal(line.unit_price or 0),
        "standard_price": line.standard_price,
        "loss_rate": line.loss_rate,
        "standard_qty": line.standard_qty,
        "required_qty": line.required_qty,
        "src": {field: getattr(line, field, "") for field in BOM_SOURCE_FIELDS},
    }


def bom_cell_value(row, col, index=0, group_no=0):
    """取一格的**原始值**（页面显示与导出共用同一套取值）。"""
    special = col.get("special")
    if special == "serial":
        return index
    if special == "group":
        return group_no
    if "const" in col:
        return col["const"]
    field = col.get("field")
    if field:
        return (row.get("src") or {}).get(field, "")
    key = col.get("key")
    if key:
        return row.get(key)
    return ""


def _bom_cell_for_excel(value):
    """写进 Excel 的单元格值：Decimal 超过 6 位小数就截到 6 位（去掉无意义长尾）。"""
    if isinstance(value, Decimal):
        try:
            if -value.as_tuple().exponent > 6:
                return value.quantize(Decimal("0.000001"))
        except (ArithmeticError, ValueError):
            return value
    return value


def _bom_group_thousands(text):
    """加千分位（分节号）。"""
    if "." in text:
        int_part, dec_part = text.split(".")
    else:
        int_part, dec_part = text, ""
    negative = int_part.startswith("-")
    if negative:
        int_part = int_part[1:]
    grouped = []
    while int_part:
        grouped.insert(0, int_part[-3:])
        int_part = int_part[:-3]
    result = ",".join(grouped) if grouped else "0"
    if negative:
        result = "-" + result
    return f"{result}.{dec_part}" if dec_part else result


def bom_cell_text(value, decimals=None, blank_zero=False):
    """一格在页面上的显示文字：空值 → 「—」；数量列按指定小数位（带千分位）。

    blank_zero=True 的列（单价 / 标准成品单价）把 0 也显示成「—」——
    0 表示「取不到价」（月末加权价与历史单价都没有），不是真的 0 元。
    """
    if value is None or value == "":
        return "—"
    if blank_zero:
        try:
            if Decimal(str(value)) == 0:
                return "—"
        except (InvalidOperation, ValueError, TypeError):
            pass
    if decimals is not None:
        try:
            fixed = Decimal(str(value)).quantize(
                Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP
            )
        except (InvalidOperation, ValueError, ArithmeticError):
            return str(value)
        return _bom_group_thousands(f"{fixed:f}")
    try:
        number = Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError):
        return str(value)
    text = format(number.normalize(), "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return _bom_group_thousands(text or "0")


def bom_columns(period, ledger=False):
    """「BOM 明细」的列定义（页面与导出共用一套）。

    有上传文件时：**上传文件的列原样同序**（列名照抄），系统列追加在最右边；
    上传列里已经有的字段（12 列老格式自带「单价（子项物料的单价）」「标准成品单价」）不再重复追加。
    没上传过 BOM 文件时回退系统固定列（与改版前一致）。
    """
    period_label = period.month.strftime("%Y-%m")
    source = bom_source_header(period)
    if not source:
        if ledger:
            return [
                {"name": "工单编号", "key": "order_no"},
                {"name": "物料编号", "field": "material_code"},
                {"name": "物料名称", "field": "material_name"},
                {"name": "项次", "field": "line_no"},
                {"name": "子项物料编号", "field": "sub_item_code"},
                {"name": "子项物料名称", "field": "sub_item_name"},
                {"name": "子项用量:分子", "field": "qty_numerator", "num": True},
                {"name": "子项用量:分母", "field": "qty_denominator", "num": True},
                {"name": "用量", "key": "qty", "num": True},
                {"name": "用量(转100%)", "key": "qty_normalized", "num": True},
                {
                    "name": "转换后子项用量:分子",
                    "key": "qty_converted_numerator",
                    "num": True,
                },
                {"name": "单位名称", "field": "uom"},
                {"name": "损耗率(%)", "field": "loss_rate", "num": True},
                {"name": "标准用量", "field": "standard_qty", "num": True},
                {"name": "需求数", "field": "required_qty", "num": True},
                {"name": "单价", "key": "unit_price", "num": True, "blank_zero": True},
                {
                    "name": "标准成品单价",
                    "key": "standard_price",
                    "num": True,
                    "blank_zero": True,
                },
            ]
        return [
            {"name": "期间", "const": period_label},
            {"name": "车间", "key": "workshop_label"},
            {"name": "工单编号", "key": "order_no"},
            {"name": "产品编码", "key": "product_code"},
            {"name": "产品名称", "key": "product_name"},
            {"name": "项次", "field": "line_no"},
            {"name": "子项物料编号", "field": "sub_item_code"},
            {"name": "子项物料名称", "field": "sub_item_name"},
            {"name": "子项用量:分子", "field": "qty_numerator", "decimals": 2, "num": True},
            {"name": "子项用量:分母", "field": "qty_denominator", "decimals": 2, "num": True},
            {"name": "用量", "key": "qty", "decimals": 2, "num": True},
            {"name": "用量(转100%)", "key": "qty_normalized", "decimals": 2, "num": True},
            {
                "name": "转换后子项用量:分子",
                "key": "qty_converted_numerator",
                "decimals": 2,
                "num": True,
            },
            {"name": "单位名称", "field": "uom"},
            {
                "name": "单价",
                "key": "unit_price",
                "decimals": 4,
                "num": True,
                "blank_zero": True,
            },
            {"name": "金额", "key": "amount", "decimals": 2, "num": True},
            {
                "name": "标准成品单价",
                "key": "standard_price",
                "decimals": 2,
                "num": True,
                "blank_zero": True,
            },
        ]

    alias = bom_alias_field_map()
    columns = []
    for name in source:
        normalized = normalize_header(name)
        col = {"name": name}
        if normalized == "序列":
            col["special"] = "serial"
            col["num"] = True
        elif normalized == "序号":
            col["special"] = "group"
            col["num"] = True
        else:
            field = alias.get(normalized)
            if field:
                col["field"] = field
                if field in ("qty_numerator", "qty_denominator"):
                    # 用户 2026-09-24：本表这两列按两位小数显示
                    col["decimals"] = 2
                if field in ("unit_price", "standard_price"):
                    # 取不到价（0）显示「—」，不是真的 0 元
                    col["blank_zero"] = True
                col["num"] = field in _BOM_NUMERIC_FIELDS
        columns.append(col)
    mapped = {col["field"] for col in columns if col.get("field")}
    for name, key, decimals in (
        BOM_LEDGER_SYSTEM_COLUMNS if ledger else BOM_DETAIL_SYSTEM_COLUMNS
    ):
        field = alias.get(normalize_header(name))
        if field and field in mapped:
            continue  # 上传列里已有这一列（老格式）
        col = {"name": name, "key": key, "decimals": decimals, "num": True}
        if key in ("unit_price", "standard_price"):
            col["blank_zero"] = True
        columns.append(col)
    return columns


def build_bom_detail_table(
    period, rows=None, workshop_filter="", order_filter="", ledger=False,
    missing_only=False,

):
    """「BOM 明细」的列 + 行（取数、列定义、序列/序号、折算一次算好，页面与导出共用）。

    返回 (columns, data_rows)：columns＝列定义；data_rows＝行字典（带 _index/_group，
    即「序列」「序号」的取值）。取值/显示文字分别由 attach_bom_cells /
    attach_bom_cell_texts 按需挂上——页面只给当前页做，导出才给全部行做（近万行不做无用功）。
    """
    detail_rows = build_standard_cost_bom_rows(
        period, rows=rows, workshop_filter=workshop_filter, order_filter=order_filter,
            missing_only=missing_only,
        )
    columns = bom_columns(period, ledger=ledger)
    previous = None
    group_no = 0
    for index, row in enumerate(detail_rows, start=1):
        if row["order_no"] != previous:
            group_no += 1
            previous = row["order_no"]
        row["_index"] = index
        row["_group"] = group_no
    return columns, detail_rows


def attach_bom_cells(rows, columns):
    """按列取原始值挂到行上（cells）：页面只给当前页，导出给全部行。"""
    for row in rows:
        row["cells"] = [
            bom_cell_value(
                row, col, index=row.get("_index", 0), group_no=row.get("_group", 0)
            )
            for col in columns
        ]
    return rows


def attach_bom_cell_texts(rows, columns):
    """给（当前页的）行挂上显示文字 cell_texts＝[(文字, 是否数值右对齐), …]。"""
    for row in rows:
        values = row.get("cells")
        if values is None:
            values = [
                bom_cell_value(
                    row, col, index=row.get("_index", 0), group_no=row.get("_group", 0)
                )
                for col in columns
            ]
            row["cells"] = values
        row["cell_texts"] = [
            (
                bom_cell_text(
                    value, col.get("decimals"), col.get("blank_zero", False)
                ),
                bool(col.get("num")),
            )
            for value, col in zip(values, columns)
        ]
    return rows


def build_standard_cost_bom_rows(
    period, rows=None, workshop_filter="", order_filter="", head=None,
    missing_only=False
):
    """标准成本「BOM 明细」行（页面与导出共用，用户 2026-09-23）。

    把「工单标准成本计算表」里那些工单的 BOM 子项逐行展开，用来解释该工单的
    「标准材料单位成本」（＝BOM 标准成品单价）是怎么来的：
    **用户 2026-09-30：范围改成「本期全部 BOM 行」**（＝上传的 BOM 文件全部行，
    不再只限有结算结果的工单；没有结算结果的工单其单价/标准成品单价为空）。

    - rows 给定时只列这些工单（导出「标准成本」工作簿里与计算表筛选一致）；
      rows 为空时按 workshop_filter / order_filter 自己取工单范围
      （用户 2026-09-24：页面「BOM 明细表」有自己的一套期间/车间/工单号筛选）；
    - 金额 ＝ 子项用量 × 子项单价；同一工单各行的金额合计＝该工单的标准成品单价
      （Σ用量 > 100% 的工单，页面的标准成品单价按 Σ(用量×单价)÷Σ用量 折算，
      见「BOM表台账」页说明）；
    - 转换后子项用量:分子 ＝ 折算后用量 × 100（与 BOM 表台账页同一列口径一致）。
    """
    if head is None:
        if rows is not None:
            head = {row["order_no"]: row for row in rows if row.get("order_no")}
        else:
            head = bom_detail_head_map(period)
    order_set = {order_no for order_no in head if order_no}
    if workshop_filter:
        if workshop_filter == PULP_WORKSHOP:
            order_set = set()  # 浆粕车间没有工单、也没有 BOM
        else:
            order_set = {
                order_no
                for order_no in order_set
                if (head.get(order_no) or {}).get("workshop") == workshop_filter
            }
    if order_filter:
        needle = order_filter.lower()
        order_set = {order_no for order_no in order_set if needle in order_no.lower()}
    if not order_set:
        return []
    qty_total = bom_qty_total_by_order(period)
    share = bom_share_resolver(period)
    # 用户 2026-10-15：勾选"只看缺单价（当月核算工单）"时——
    #   只保留 ①子项取不到价 且 ②该工单本月确有投料/报工 的 BOM 行，方便一次导出补价。
    missing_prices = bom_material_prices(period) if missing_only else {}
    current_orders = set()
    if missing_only:
        try:
            current_orders |= {
                str(x).strip()
                for x in MaterialIssue.objects.filter(period=period)
                .values_list("work_order_no", flat=True) if x
            }
            current_orders |= {
                str(x).strip()
                for x in WorkReport.objects.filter(period=period)
                .values_list("work_order_no", flat=True) if x
            }
        except Exception:  # noqa: BLE001
            current_orders = set()
    result = []
    for line in period.bom_lines.all():
        if line.work_order_no not in order_set:
            continue
        if missing_only:
            code = (line.sub_item_code or "").strip()
            try:
                priced = Decimal(missing_prices.get(code) or 0) > 0
            except Exception:  # noqa: BLE001
                priced = False
            if priced:
                continue
            if current_orders and (line.work_order_no or "").strip() not in current_orders:
                continue
        row = bom_row_from_line(line, qty_total, share)
        head_row = head.get(row["order_no"]) or {}
        row["workshop_label"] = head_row.get("workshop_label", "")
        row["product_code"] = head_row.get("product_code") or line.material_code
        row["product_name"] = head_row.get("product_name") or line.material_name
        result.append(row)
    # 行序＝上传顺序（导入顺序），与「导出中心 → BOM表」一致，便于逐行跟上传文件核对
    result.sort(key=lambda row: row["id"])
    return result


def actual_workshop_cost_rates(period):
    """本月各车间的**实际**占比（参考值）：工资÷材料、制造费用÷材料。

    只用于页面上给用户一个对照，好判断录入的标准占比是否离谱；不参与任何计算。
    """
    labor = defaultdict(Decimal)
    overhead = defaultdict(Decimal)
    material = defaultdict(Decimal)
    for item in period.work_order_cost_results.all():
        workshop = item.work_order.production_dept_code
        material[workshop] += Decimal(item.direct_material or 0)
        labor[workshop] += Decimal(item.wages or 0)
        overhead[workshop] += _manaoverhead_of(item)
    reference = {}
    for workshop, base in material.items():
        if not base:
            continue
        reference[workshop] = {
            "material": base,
            "labor_pct": labor[workshop] / base * Decimal("100"),
            "overhead_pct": overhead[workshop] / base * Decimal("100"),
        }
    return reference


def build_standard_cost_composition_rows(period, month_filter="", workshop_filter=""):
    """标准成本构成表（用户 2026-09-23）：参考年 1-12 月 × 车间。

    列＝期间｜车间｜产量｜直接材料｜人工｜制造费用｜材料成本占比%｜人工成本占比%｜制造费用占比%。
    取数（用户 2026-09-23 定）：
      * 下半年（7-12 月）＝系统自动：取该月实际发生数（工单成本结果的完工数量、
        直接材料、工资、制造费用合计），没有数据显示 —；
      * 上半年（1-6 月）＝人工录入：系统里没有这些期间的数据，由用户在表内填写
        （产量/直接材料/人工/制造费用），三个占比由金额自动算出（各自 ÷ 三者合计）。
    """
    year = period.month.year
    start, end = date(year, 1, 1), date(year, 12, 1)
    cells = composition_cells_by_month(start, end)
    workshops = standard_cost_workshops(period)

    # 期间筛选：支持 "YYYY-MM" 字符串
    if month_filter:
        month_filter = parse_month_param(month_filter)

    rows = []
    for month in months_between(start, end):
        editable = month.month <= 6          # 上半年人工录；下半年系统算
        if month_filter and month_key(month) != month_key(month_filter):
            continue
        for workshop in workshops:
            if workshop_filter and workshop != workshop_filter:
                continue
            cell = (cells.get(month) or {}).get(workshop)
            material, labor, overhead = resolve_composition_pcts(cell)
            rows.append(
                {
                    "month": month,
                    "period_key": month.strftime("%Y-%m"),
                    "period_label": month.strftime("%Y-%m"),
                    "workshop": workshop,
                    "label": WORKSHOP_LABELS.get(workshop, workshop),
                    "output_qty": (cell or {}).get("output_qty"),
                    "direct_material": (cell or {}).get("direct_material"),
                    "wages": (cell or {}).get("wages"),
                    "manufacturing_overhead": (
                        (cell or {}).get("manufacturing_overhead")
                    ),
                    "material_pct": material,
                    "labor_pct": labor,
                    "overhead_pct": overhead,
                    # 人工录入过的占比（回填输入框用）
                    "has_data": cell is not None,
                    "source": (cell or {}).get("source", ""),
                    "editable": editable,
                }
            )
    return rows

def month_filter_options(period):
    """构成表「期间」下拉：参考年 1-12 月。"""
    year = period.month.year
    return [
        {"value": month.strftime("%Y-%m"), "label": month.strftime("%Y-%m")}
        for month in months_between(date(year, 1, 1), date(year, 12, 1))
    ]



def _standard_cost_totals(rows):
    """合计行：金额求和；标准价算不出（显示 —）的行不计入标准合计。"""
    fields = (
        "standard_material", "standard_labor", "standard_overhead", "standard_total",
    )
    totals = {
        "count": len(rows),
        "no_standard_price": sum(1 for r in rows if not r["has_standard_price"]),
    }
    for field in fields:
        totals[field] = sum(
            (r[field] for r in rows if r[field] is not None), Decimal("0")
        )
    totals["finished_quantity"] = sum(
        (r["finished_quantity"] for r in rows), Decimal("0")
    )
    std_qty = sum(
        (r["finished_quantity"] for r in rows if r["has_standard_price"]), Decimal("0")
    )
    totals["standard_unit_cost"] = (
        totals["standard_total"] / std_qty if std_qty else None
    )
    return totals


@login_required
def standard_costs(request, period_id):
    """标准成本：构成表（期间×车间）+ 各工单的标准成本（只出标准侧）。

    口径（用户 2026-09-23）：标准占比＝**上一个完整半年各月占比的平均数**（每半年滚动一次）；
    标准成本 ＝ 标准材料成本 ×（1 ＋ 人工占比 ＋ 制造费用占比）；标准材料成本来自 BOM
    「标准成品单价」（Σ(子项用量×子项单价)，结算后自动回填）。实际成本不在这里重复算：
    单位标准成本与实际单位成本的对照在「工单成本」页。
    """
    period = get_object_or_404(AccountingPeriod, id=period_id)
    url_period = period
    # 「工单标准成本计算表」的期间筛选（用户 2026-09-23）：与「工单成本」页同款——
    # 选到哪个月就看那个月的工单标准成本（页面上其它块也跟着走那一期）。
    period = _work_order_cost_period(
        period, str(request.GET.get("w_period", "")).strip()
    )
    workshop_filter = str(request.GET.get("w_workshop", "")).strip()
    order_filter = str(request.GET.get("order_no", "")).strip()
    product_filter = str(request.GET.get("product", "")).strip()
    # 构成表「期间」筛选（用户 2026-09-24）：默认显示当月份的数据；
    # 想一次看全年就把下拉选成「全部」（参数在、但为空 → 不按月过滤）。
    c_month_raw = request.GET.get("c_month")
    c_month_filter = (
        period.month.strftime("%Y-%m")
        if c_month_raw is None
        else str(c_month_raw).strip()
    )
    c_workshop_filter = str(request.GET.get("c_workshop", "")).strip()
    # BOM 明细表筛选（用户 2026-09-24）：本表有自己的一套期间/车间/工单号字段。
    bom_period = _work_order_cost_period(
        period, str(request.GET.get("b_period", "")).strip()
    )
    b_workshop_filter = str(request.GET.get("b_workshop", "")).strip()
    b_order_filter = str(request.GET.get("b_order_no", "")).strip()

    rows, meta = build_standard_cost_rows(
        period,
        workshop_filter=workshop_filter,
        order_filter=order_filter,
        product_filter=product_filter,
    )

    # BOM 明细（用户 2026-09-23）：把对应工单的 BOM 子项逐行列出，
    # 说明「标准材料单位成本」怎么来的；2026-09-24 起按本表自己的筛选字段取数；
    # 用户 2026-09-30：改成**本期全部 BOM 行**（＝上传文件的全部行），列＝上传文件原样 +
    # 系统列在最右边；行数上千，页面按 BOM_DETAIL_PAGE_SIZE 分页（导出不受分页影响）。
    bom_columns, bom_table_rows = build_bom_detail_table(
        bom_period,
        workshop_filter=b_workshop_filter,
        order_filter=b_order_filter,
        missing_only=str(request.GET.get("b_missing", "")).strip() in ("1", "on", "true"),
    )
    bom_page = Paginator(bom_table_rows, BOM_DETAIL_PAGE_SIZE).get_page(
        request.GET.get("b_page")
    )
    _bom_query = request.GET.copy()
    _bom_query.pop("b_page", None)
    bom_page.querystring = _bom_query.urlencode()
    attach_bom_cell_texts(bom_page.object_list, bom_columns)

    # 构成表：参考年 1-12 月 × 车间（上半年可录入，下半年系统自动）
    composition_rows = build_standard_cost_composition_rows(
        period,
        month_filter=c_month_filter,
        workshop_filter=c_workshop_filter,
    )

    # 标准占比表：用本页期间的基期（半年滚动），不再单独筛选
    view_rates, view_meta = rolling_standard_cost_rates(period)
    standard_rate_rows = []
    for code in standard_cost_workshops(period):
        rate = view_rates.get(code) or {}
        standard_rate_rows.append(
            {
                "workshop": code,
                "label": WORKSHOP_LABELS.get(code, code),
                "material_pct": rate.get("material_pct"),
                "labor_pct": rate.get("labor_pct"),
                "overhead_pct": rate.get("overhead_pct"),
                # 用户 2026-09-23：表上只留「人工占比和费用占比的和」一列
                "labor_plus_overhead_pct": rate.get("labor_plus_overhead_pct"),
                "months": rate.get("months", 0),
                "source": rate.get("source", ""),
            }
        )

    return render(
        request,
        "settlement/standard_costs.html",
        {
            "period": period,
            "rows": rows,
            "row_count": len(rows),
            "bom_columns": bom_columns,
            "bom_rows": bom_page.object_list,
            "bom_page": bom_page,
            "bom_row_count": len(bom_table_rows),
            "bom_order_count": len({r["order_no"] for r in bom_table_rows}),
            "bom_page_size": BOM_DETAIL_PAGE_SIZE,
            "bom_period_label": bom_period.month.strftime("%Y-%m"),
            "bom_period_options": _work_order_cost_period_options(url_period),
            "bom_export_period_id": bom_period.id,
            "b_filter": {
                "missing": str(request.GET.get("b_missing", "")).strip() in ("1", "on", "true"),
                "period": str(request.GET.get("b_period", "")).strip(),
                "workshop": b_workshop_filter,
                "order_no": b_order_filter,
            },
            "totals": _standard_cost_totals(rows),
            "composition_rows": composition_rows,
            "composition_count": len(composition_rows),
            "standard_rate_rows": standard_rate_rows,
            "rate_month_label": period.month.strftime("%Y-%m"),
            "view_basis_label": view_meta["basis_label"],
            "basis_label": meta["basis_label"],
            "basis_start": meta["basis_start"],
            "basis_end": meta["basis_end"],
            "carried_from": meta["carried_from"],
            "month_options": month_filter_options(period),
            "c_filter": {
                "month": c_month_filter,
                "workshop": c_workshop_filter,
            },
            "period_label": period.month.strftime("%Y-%m"),
            "url_period_id": url_period.id,
            "url_period_label": url_period.month.strftime("%Y-%m"),
            "calc_period_options": _work_order_cost_period_options(url_period),
            "workshop_options": _workshop_filter_options(period),
            "w_filter": {
                "period": str(request.GET.get("w_period", "")).strip(),
                "workshop": workshop_filter,
                "order_no": order_filter,
                "product": product_filter,
            },
            "export_period_id": period.id,
        },
    )


def _parse_pct(raw, label):
    """占比按百分数录入，如 12.5 表示 12.5%。"""
    if not raw:
        return None
    try:
        value = Decimal(str(raw).replace(",", "").replace("%", "").strip())
    except (InvalidOperation, ValueError):
        raise ValueError("%s「%s」不是数字，请填如 12.5 表示 12.5%%。" % (label, raw))
    if value < 0:
        raise ValueError("%s不能为负数。" % label)
    if value > 1000:
        raise ValueError(
            "%s「%s」过大（填的是百分数吗？如 12.5 表示 12.5%%）。" % (label, raw)
        )
    return value.quantize(Decimal("0.0001"))


def _parse_amount(raw, label):
    """金额/数量录入。"""
    if not raw:
        return Decimal("0")
    try:
        value = Decimal(str(raw).replace(",", "").replace("%", "").strip())
    except (InvalidOperation, ValueError):
        raise ValueError("%s「%s」不是数字。" % (label, raw))
    if value < 0:
        raise ValueError("%s不能为负数。" % label)
    return value.quantize(Decimal("0.000001"))


@login_required
@require_POST
def save_standard_cost_rates(request, period_id):
    """保存标准成本构成表（页面用 AJAX 提交，返回 JSON，不跳页）。

    用户 2026-09-23：点「保存构成表」不要跳页——所以这里只回 JSON，
    由页面弹出「已保存」提示；需要跳转的调用方（比如表单直提）请用下面的
    save_standard_cost_composition_redirect。

    两件事一起存：
      ① 上半年（1-6 月）各车间的人工录入（产量/直接材料/人工/制造费用）；
      ② 兜底的手工占比参数（基期半年一点数据都没有时沿用，仍是按百分数录入）。
    """
    try:
        _save_composition_entries(request, period_id)
        _save_fallback_rates(request, period_id)
    except Exception as exc:
        return JsonResponse({"ok": False, "message": "保存失败：" + str(exc)})
    period = get_object_or_404(AccountingPeriod, id=period_id)
    rates, rate_meta = rolling_standard_cost_rates(period)
    return JsonResponse(
        {
            "ok": True,
            "message": "标准成本构成表已保存。",
            "basis_label": rate_meta["basis_label"],
            "rates": [
                {
                    "workshop": code,
                    "material_pct": (
                        str(item["material_pct"]) if item.get("material_pct") is not None
                        else None
                    ),
                    "labor_pct": (
                        str(item["labor_pct"]) if item.get("labor_pct") is not None
                        else None
                    ),
                    "overhead_pct": (
                        str(item["overhead_pct"]) if item.get("overhead_pct") is not None
                        else None
                    ),
                    "labor_plus_overhead_pct": (
                        str(item["labor_plus_overhead_pct"])
                        if item.get("labor_plus_overhead_pct") is not None else None
                    ),
                    "months": item.get("months", 0),
                    "source": item.get("source", ""),
                }
                for code, item in rates.items()
            ],
        }
    )


def _save_composition_entries(request, period_id):
    """把构成表上半月（1-6 月）的录入写库；返回写入/删除的记录数。

    只处理页面里真正提交了的格子（按 c_month / c_workshop 筛选后可能只提交一部分），
    四个金额都为空的格子＝清空（删掉已有记录），避免旧数据留着骗人。
    """
    from settlement.models import StandardCostCompositionEntry

    period = get_object_or_404(AccountingPeriod, id=period_id)
    saved = 0
    with transaction.atomic():
        year = period.month.year
        for month in months_between(date(year, 1, 1), date(year, 6, 1)):
            key = month.strftime("%Y-%m")
            for code in standard_cost_workshops(period):
                qty_field = "qty_%s_%s" % (key, code)
                if qty_field not in request.POST:
                    continue  # 这一格这次没提交（被筛选掉了）
                amounts = {
                    "qty": request.POST.get(qty_field, ""),
                    "material": request.POST.get("material_%s_%s" % (key, code), ""),
                    "wages": request.POST.get("wages_%s_%s" % (key, code), ""),
                    "overhead": request.POST.get("overhead_%s_%s" % (key, code), ""),
                }
                label = "%s %s" % (key, WORKSHOP_LABELS.get(code, code))
                if not any(str(value).strip() for value in amounts.values()):
                    StandardCostCompositionEntry.objects.filter(
                        period_month=month, workshop=code
                    ).delete()
                    continue
                StandardCostCompositionEntry.objects.update_or_create(
                    period_month=month,
                    workshop=code,
                    defaults={
                        "output_qty": _parse_amount(amounts["qty"], label + " 产量"),
                        "direct_material": _parse_amount(
                            amounts["material"], label + " 直接材料"
                        ),
                        "wages": _parse_amount(amounts["wages"], label + " 人工"),
                        "manufacturing_overhead": _parse_amount(
                            amounts["overhead"], label + " 制造费用"
                        ),
                        "updated_by": request.user,
                    },
                )
                saved += 1
    return saved


def _save_fallback_rates(request, period_id):
    """兜底的手工占比参数（人工费用占比 / 制造费用占比，基数＝材料成本）。"""
    period = get_object_or_404(AccountingPeriod, id=period_id)
    with transaction.atomic():
        for code in standard_cost_workshops(period):
            raw_labor = str(request.POST.get("labor_%s" % code, "")).strip()
            raw_overhead = str(request.POST.get("overhead_%s" % code, "")).strip()
            if not raw_labor and not raw_overhead:
                continue
            StandardCostRate.objects.update_or_create(
                period=period,
                workshop=code,
                defaults={
                    "labor_rate_pct": (
                        _parse_pct(raw_labor, "人工费用占比") or Decimal("0")
                    ),
                    "overhead_rate_pct": (
                        _parse_pct(raw_overhead, "制造费用占比") or Decimal("0")
                    ),
                    "updated_by": request.user,
                },
            )


@login_required
@require_POST
def save_standard_cost_rates_redirect(request, period_id):
    """老入口（表单直提）：保存后跳回页面并给一条 messages 提示。"""
    period = get_object_or_404(AccountingPeriod, id=period_id)
    try:
        _save_composition_entries(request, period_id)
        _save_fallback_rates(request, period_id)
        messages.success(request, "标准成本构成表已保存。")
    except Exception as exc:
        messages.error(request, "保存标准成本构成表失败：" + str(exc))
    return redirect("standard_costs", period_id=period.id)


def prepare_price_table_data(period, n=6):
    """准备单价表（跨月对比）数据：期间列表、物料行、类别筛选。

    物料列表 = 各期间库存计价 ∪ 上传历史单价（未结算期间的上传价也显示）。
    单价优先上传历史单价；未上传用系统计算价。含历史单价自动补全。
    """
    from settlement.models import AccountingPeriod, InventoryTransaction
    from settlement.services.historical_prices import autofill_historical_prices

    periods = list(
        AccountingPeriod.objects.filter(month__lte=period.month)
        .order_by("-month")[:n]
    )
    periods = list(reversed(periods))

    # 历史单价自动补全：用户只需上传一次基准，期间内新增物料自动按系统计算单价补入
    autofill_historical_prices(periods)

    per_period = {
        p.id: {v.material_code: v for v in p.inventory_valuations.all()}
        for p in periods
    }
    # 用户上传的历史单价（优先展示；未上传的物料仍用系统计算单价）
    historical = {
        p.id: {hp.material_code: hp for hp in p.historical_prices.all()}
        for p in periods
    }
    # 物料列表 = 库存计价 ∪ 历史单价：上传过单价的物料即使该期间未结算也显示
    codes = set()
    for d in per_period.values():
        codes.update(d.keys())
    for d in historical.values():
        codes.update(d.keys())

    # 用户 2026-09-30：类别名称要用**上传的原始分类**（期初/期末存货 → 库存流水 → 投料记录），
    #   取不到才记「未分类」；不再只压成 原料/成品/其它 三个大类（那样看不出来是什么类别）。
    # 内部流转（调拨/库存调整…）的数量：算单价时要扣掉（它们没有成本金额）
    from settlement.services.price_table import internal_transfer_qty as _internal_qty

    internal_qty = _internal_qty(periods)

    cat_by_code = {}
    uom_by_code = {}
    for code, cat, uom in (
        InventoryTransaction.objects.filter(period__in=periods)
        .exclude(category_name="")
        .order_by("material_code")
        .values_list("material_code", "category_name", "uom")
    ):
        cat_by_code.setdefault(code, cat)
        uom_by_code.setdefault(code, uom)
    # 与「库存结果」页同一取法（期初存货 → 期末存货 → 库存流水 → 投料记录），逐期间合并
    for _p in periods:
        for _code, _cat in (inventory_category_map(_p) or {}).items():
            if _cat:
                cat_by_code.setdefault(_code, _cat)

    name_by_code = {}
    for p in periods:
        for v in p.inventory_valuations.all():
            name_by_code.setdefault(v.material_code, v.material_name)
    for p in periods:
        for hp in p.historical_prices.all():
            name_by_code.setdefault(hp.material_code, hp.material_name)

    RAW = {"原料", "染色纤维", "纤维", "熟条", "浆粕"}
    FIN = {"纱线", "面料", "坯布", "布", "成品"}
    # 名称关键字推断（用户 2026-09-30：有的看名字就知道是纱线还是面料，BC- 开头就是粗纱）
    NAME_RULES = (
        ("面料", ("面料", "坯布", "布")),
        ("纱线", ("纱线", "纱", "线")),
        ("长丝", ("长丝",)),
        ("纤维", ("纤维", "短纤", "短纤维")),
        ("浆粕", ("浆粕",)),
    )

    def infer_cat(code, name=""):
        if str(code).upper().startswith("BC-"):
            return "粗纱"
        text = (name or "") + str(code)
        for label, keywords in NAME_RULES:
            if any(k in text for k in keywords):
                return label
        return ""

    def big_cat(code):
        cat = cat_by_code.get(code, "")
        if cat in RAW:
            return "原料"
        if cat in FIN:
            return "成品"
        if infer_cat(code, name_by_code.get(code, "")) in ("纱线", "面料"):
            return "成品"
        if infer_cat(code, name_by_code.get(code, "")):
            return "原料"
        return "其它"

    rows = []
    for code in sorted(codes):
        cat = big_cat(code)
        cells = []
        for p in periods:
            v = per_period[p.id].get(code)
            hp = historical.get(p.id, {}).get(code)
            if v is None:
                if hp is not None:
                    rp = Decimal(hp.receipt_price) if hp.receipt_price is not None else Decimal("0")
                    ip = Decimal(hp.issue_price) if hp.issue_price is not None else Decimal("0")
                    cells.append((Decimal("0"), Decimal("0"), rp, Decimal("0"), Decimal("0"), ip))
                else:
                    cells.append((Decimal("0"), Decimal("0"), Decimal("0"), Decimal("0"), Decimal("0"), Decimal("0")))
                continue
            rq = Decimal(v.receipt_quantity or 0)
            ra = Decimal(v.receipt_amount or 0)
            iq = Decimal(v.issue_quantity or 0)
            ia = Decimal(v.issue_amount or 0)
            # 单价的分母要扣掉内部流转数量（调拨/库存调整等没有成本金额）
            in_internal, out_internal = internal_qty.get((p.month.month, code), [Decimal("0"), Decimal("0")])
            rq_price = rq - Decimal(in_internal)
            iq_price = iq - Decimal(out_internal)
            rp = (ra / rq_price).quantize(Decimal("0.0001")) if rq_price > 0 else Decimal("0")
            ip = (ia / iq_price).quantize(Decimal("0.0001")) if iq_price > 0 else Decimal("0")
            if hp is not None:
                if hp.receipt_price is not None:
                    rp = Decimal(hp.receipt_price)
                if hp.issue_price is not None:
                    ip = Decimal(hp.issue_price)
            cells.append((rq, ra, rp, iq, ia, ip))
        # 用户 2026-10-16：单价表也**沿用上月**——当月没有价（空或 0）时，取最近一个有值的月份
        _last_in = None
        _last_out = None
        _fixed = []
        for _rq, _ra, _rp, _iq, _ia, _ip in cells:
            if (_rp is None or Decimal(_rp) == Decimal("0")) and _last_in is not None:
                _rp = _last_in
            if (_ip is None or Decimal(_ip) == Decimal("0")) and _last_out is not None:
                _ip = _last_out
            if _rp is not None and Decimal(_rp) != Decimal("0"):
                _last_in = _rp
            if _ip is not None and Decimal(_ip) != Decimal("0"):
                _last_out = _ip
            _fixed.append((_rq, _ra, _rp, _iq, _ia, _ip))
        cells = _fixed
        rows.append({
            "category": cat,
            "category_name": (
                cat_by_code.get(code)
                or infer_cat(code, name_by_code.get(code, ""))
                or "未分类"
            ),
            "code": code,
            "name": name_by_code.get(code, ""),
            "uom": uom_by_code.get(code, ""),
            "cells": cells,
        })

    rows.sort(key=lambda r: ({"原料": 0, "成品": 1, "其它": 2}.get(r["category"], 9), r["code"]))
    return periods, rows


def build_price_table_workbook(period, n=None, months=None, code_filter=""):
    """单价表：两个工作表（入库单价 / 出库单价），列为 序号/品类/品名编码/品名名称 + 当年 1~12 月。

    口径见 settlement.services.price_table（n 参数保留仅为兼容旧调用）。
    """
    from settlement.services.price_table import (
        build_price_table_workbook as _build,
    )

    return _build(period, n, months=months, code_filter=code_filter)


@login_required
def price_table_view(request, period_id):
    """单价表（跨月对比）：每个物料一行，每月一组「入库单价、出库单价」。

    物料列表 = 各期间库存计价 ∪ 上传的历史单价（上传过单价的物料即使该期间未结算也显示）；
    单价优先用上传的历史单价，未上传的用系统计算价（入库金额÷数量、出库金额÷数量）。
    默认取最近 6 期（含当前）。页面已取消「期间数 / 类别」筛选控件，
    仅保留 ?periods=N（1~12）与 ?category=all|raw|fin|other 两个 URL 参数的兼容支持；
    ?export=1 导出 Excel。
    """
    from io import BytesIO

    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    from settlement.models import AccountingPeriod, InventoryTransaction

    period = get_object_or_404(AccountingPeriod, id=period_id)
    try:
        n = min(max(int(request.GET.get("periods", "6")), 1), 12)
    except (TypeError, ValueError):
        n = 6
    periods = list(
        AccountingPeriod.objects.filter(month__lte=period.month)
        .order_by("-month")[:n]
    )
    periods = list(reversed(periods))

    # 历史单价自动补全：用户只需上传一次基准，期间内新增物料自动按系统计算单价补入
    from settlement.services.historical_prices import autofill_historical_prices

    autofill_historical_prices(periods)

    per_period = {
        p.id: {v.material_code: v for v in p.inventory_valuations.all()}
        for p in periods
    }
    # 用户上传的历史单价（优先展示；未上传的物料仍用系统计算单价）
    historical = {
        p.id: {hp.material_code: hp for hp in p.historical_prices.all()}
        for p in periods
    }
    # 物料列表 = 库存计价 ∪ 历史单价：上传过单价的物料即使该期间未结算也显示
    codes = set()
    for d in per_period.values():
        codes.update(d.keys())
    for d in historical.values():
        codes.update(d.keys())

    # 用户 2026-09-30：类别名称要用**上传的原始分类**（期初/期末存货 → 库存流水 → 投料记录），
    #   取不到才记「未分类」；不再只压成 原料/成品/其它 三个大类（那样看不出来是什么类别）。
    cat_by_code = {}
    uom_by_code = {}
    for code, cat, uom in (
        InventoryTransaction.objects.filter(period__in=periods)
        .exclude(category_name="")
        .order_by("material_code")
        .values_list("material_code", "category_name", "uom")
    ):
        cat_by_code.setdefault(code, cat)
        uom_by_code.setdefault(code, uom)
    # 与「库存结果」页同一取法（期初存货 → 期末存货 → 库存流水 → 投料记录），逐期间合并
    for _p in periods:
        for _code, _cat in (inventory_category_map(_p) or {}).items():
            if _cat:
                cat_by_code.setdefault(_code, _cat)

    name_by_code = {}
    for p in periods:
        for v in p.inventory_valuations.all():
            name_by_code.setdefault(v.material_code, v.material_name)
    for p in periods:
        for hp in p.historical_prices.all():
            name_by_code.setdefault(hp.material_code, hp.material_name)

    RAW = {"原料", "染色纤维", "纤维", "熟条", "浆粕"}
    FIN = {"纱线", "面料", "坯布", "布", "成品"}
    # 名称关键字推断（用户 2026-09-30：有的看名字就知道是纱线还是面料，BC- 开头就是粗纱）
    NAME_RULES = (
        ("面料", ("面料", "坯布", "布")),
        ("纱线", ("纱线", "纱", "线")),
        ("长丝", ("长丝",)),
        ("纤维", ("纤维", "短纤", "短纤维")),
        ("浆粕", ("浆粕",)),
    )

    def infer_cat(code, name=""):
        if str(code).upper().startswith("BC-"):
            return "粗纱"
        text = (name or "") + str(code)
        for label, keywords in NAME_RULES:
            if any(k in text for k in keywords):
                return label
        return ""

    def big_cat(code):
        cat = cat_by_code.get(code, "")
        if cat in RAW:
            return "原料"
        if cat in FIN:
            return "成品"
        if infer_cat(code, name_by_code.get(code, "")) in ("纱线", "面料"):
            return "成品"
        if infer_cat(code, name_by_code.get(code, "")):
            return "原料"
        return "其它"

    category_filter = (request.GET.get("category") or "all").strip()
    want_cat = {"raw": "原料", "fin": "成品", "other": "其它"}.get(category_filter)
    # 用户 2026-09-30：页面加字段筛选 —— 物料编码（包含）+ 期间范围（从/到，YYYY-MM）
    code_filter = (request.GET.get("code") or "").strip()
    month_from = (request.GET.get("p_from") or "").strip()
    month_to = (request.GET.get("p_to") or "").strip()
    if month_from or month_to:
        def _in_range(p):
            m = p.month.strftime("%Y-%m")
            if month_from and m < month_from:
                return False
            if month_to and m > month_to:
                return False
            return True

        periods = [p for p in periods if _in_range(p)] or periods

    rows = []
    for code in sorted(codes):
        if code_filter and code_filter.lower() not in code.lower():
            continue
        cat = big_cat(code)
        if want_cat and cat != want_cat:
            continue
        cells = []
        for p in periods:
            v = per_period[p.id].get(code)
            hp = historical.get(p.id, {}).get(code)
            if v is None:
                # 该期无结算计价：有上传历史单价则显示上传价，否则留空
                if hp is not None:
                    rp = Decimal(hp.receipt_price) if hp.receipt_price is not None else Decimal("0")
                    ip = Decimal(hp.issue_price) if hp.issue_price is not None else Decimal("0")
                    cells.append((Decimal("0"), Decimal("0"), rp, Decimal("0"), Decimal("0"), ip))
                else:
                    cells.append((Decimal("0"), Decimal("0"), Decimal("0"), Decimal("0"), Decimal("0"), Decimal("0")))
                continue
            rq = Decimal(v.receipt_quantity or 0)
            ra = Decimal(v.receipt_amount or 0)
            iq = Decimal(v.issue_quantity or 0)
            ia = Decimal(v.issue_amount or 0)
            rp = (ra / rq).quantize(Decimal("0.0001")) if rq else Decimal("0")
            ip = (ia / iq).quantize(Decimal("0.0001")) if iq else Decimal("0")
            if hp is not None:
                if hp.receipt_price is not None:
                    rp = Decimal(hp.receipt_price)
                if hp.issue_price is not None:
                    ip = Decimal(hp.issue_price)
            cells.append((rq, ra, rp, iq, ia, ip))
        # 用户 2026-10-16（口径说明里早就写了）：当月没有价（空或 0）→ **沿用上月（及更早）**单价
        _last_in = None
        _last_out = None
        _fixed = []
        for _rq, _ra, _rp, _iq, _ia, _ip in cells:
            if _rp is not None and Decimal(_rp) != Decimal("0"):
                _last_in = _rp
            elif _last_in is not None:
                _rp = _last_in
            if _ip is not None and Decimal(_ip) != Decimal("0"):
                _last_out = _ip
            elif _last_out is not None:
                _ip = _last_out
            _fixed.append((_rq, _ra, _rp, _iq, _ia, _ip))
        cells = _fixed
        rows.append({
            "category": cat,
            "category_name": (
                cat_by_code.get(code)
                or infer_cat(code, name_by_code.get(code, ""))
                or "未分类"
            ),
            "code": code,
            "name": name_by_code.get(code, ""),
            "uom": uom_by_code.get(code, ""),
            "cells": cells,
        })

    rows.sort(key=lambda r: ({"原料": 0, "成品": 1, "其它": 2}.get(r["category"], 9), r["code"]))

    def _price(a, q):
        q = Decimal(q or 0)
        a = Decimal(a or 0)
        return (a / q).quantize(Decimal("0.0001")) if q else Decimal("0")

    if request.GET.get("export"):
        stream = build_price_table_workbook(
            period,
            n,
            months=[p.month.month for p in periods],
            code_filter=code_filter,
        )
        response = HttpResponse(
            stream.getvalue(),
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response["Content-Disposition"] = (
            f"attachment; filename*=UTF-8''{quote('单价表-跨月对比-' + period.month.strftime('%Y%m') + '.xlsx')}"
        )
        return response

    return render(
        request,
        "settlement/price_table.html",
        {
            "period": period,
            "periods": periods,
            "rows": rows,
            "price": _price,
            "filters": {
                "code": code_filter,
                "p_from": month_from,
                "p_to": month_to,
                "category": category_filter,
            },
        },
    )


@login_required
def reconciliation(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    return render(
        request,
        "settlement/reconciliation.html",
        {
            "period": period,
            "checks": period.reconciliations.all(),
            "exceptions": _page(request, period.exceptions.all(), 80),
        },
    )


@login_required
@require_POST

@login_required
@require_POST
def confirm_reconciliation(request, period_id, check_id):
    """确认/取消确认一条勾稽差异：确认后该条 fail 不再阻断锁账。"""
    period = get_object_or_404(AccountingPeriod, id=period_id)
    if period.status == AccountingPeriod.Status.LOCKED:
        messages.error(request, "该期间已锁账，不能再调整勾稽确认。")
        return redirect("reconciliation", period_id=period.id)
    check = get_object_or_404(ReconciliationResult, id=check_id, period=period)
    action = request.POST.get("action", "")
    if action == "confirm":
        check.confirmed = True
        check.confirmed_at = timezone.now()
        check.confirmed_by = request.user
        check.confirmed_note = (request.POST.get("note", "") or "")[:255]
        check.save(update_fields=["confirmed", "confirmed_at", "confirmed_by", "confirmed_note"])
        messages.success(request, f"已确认差异「{check.check_name}」，不再阻断锁账。")
    elif action == "unconfirm":
        check.confirmed = False
        check.confirmed_at = None
        check.confirmed_by = None
        check.confirmed_note = ""
        check.save(update_fields=["confirmed", "confirmed_at", "confirmed_by", "confirmed_note"])
        messages.success(request, f"已取消确认「{check.check_name}」。")
    return redirect("reconciliation", period_id=period.id)


def resolve_exception(request, period_id, exception_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    item = get_object_or_404(ExceptionItem, id=exception_id, period=period)
    if period.status == AccountingPeriod.Status.LOCKED:
        messages.error(request, "期间已锁账，不能修改异常状态。")
    else:
        item.status = (
            ExceptionItem.Status.IGNORED
            if request.POST.get("action") == "ignore"
            else ExceptionItem.Status.RESOLVED
        )
        item.resolved_at = timezone.now()
        item.resolved_by = request.user
        item.save(update_fields=["status", "resolved_at", "resolved_by"])
        messages.success(request, "异常状态已更新。")
    return redirect("reconciliation", period_id=period.id)


@login_required
@require_POST
def lock_period(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    lockable, reason = can_lock(period)
    if not lockable:
        messages.error(request, reason)
    else:
        period.status = AccountingPeriod.Status.LOCKED
        period.locked_at = timezone.now()
        period.locked_by = request.user
        period.save(update_fields=["status", "locked_at", "locked_by"])
        messages.success(request, f"{period}已锁账。")
    return redirect("period_detail", period_id=period.id)


@login_required
@require_POST
def reopen_period(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    if not request.user.is_superuser:
        return HttpResponseBadRequest("只有系统管理员可以重开期间。")
    # 用户 2026-10-13：重开期间是不可逆操作（会重置车间结转、删除结算记录），
    #   必须输入**登录密码**确认，密码不对就不重开。
    password = (request.POST.get("password") or "").strip()
    if not password:
        messages.error(request, "重开期间需要输入登录密码确认，本次未重开。")
        return redirect("period_detail", period_id=period.id)
    if not request.user.check_password(password):
        messages.error(request, "密码不正确，本次未重开。")
        return redirect("period_detail", period_id=period.id)
    try:
        invalidated = invalidate_downstream_periods(period)
        period.status = AccountingPeriod.Status.OPEN
        period.locked_at = None
        period.locked_by = None
        period.save(update_fields=["status", "locked_at", "locked_by"])
        reset_workshop_settlements(period)
        period.runs.all().delete()
        messages.warning(
            request,
            f"{period}已重新打开，重新结算后才能再次锁账。"
            + (
                f"同时已使{invalidated}个后续未锁账期间的自动承接和结算结果失效。"
                if invalidated
                else ""
            ),
        )
    except Exception as exc:
        messages.error(request, str(exc))
    return redirect("period_detail", period_id=period.id)


#: 「生产工单」页 / 导出的列：与上传的生产工单文件完全一致（列名、顺序）
WORK_ORDER_COLUMNS = [
    "工单编号", "工单类型", "标识码", "计划数量", "计划开始时间", "计划完工时间",
    "优先级", "产线设备", "班次", "生产部门编号", "生产主管编号", "计划部门编号",
    "计划员编号", "备注", "指定用料", "外部单据编号", "状态",
    "是否使用物料清单", "是否使用工艺路线", "工序计划类型", "启用SOP",
    "生产模式", "产线编号", "订单号", "物料颜色",
]


LISTING_EXPORTS = {
    "inventory_ledger": {"title": "库存流水"},
    "inventory_matches": {"title": "库存流水匹配"},
    "purchase_contract": {"title": "采购合同"},
    "purchase_receipt": {"title": "采购入库"},
    "ending_wip": {"title": "期末在产品"},
    "standard_cost": {"title": "标准成本"},
    "standard_cost_bom": {"title": "标准成本BOM明细"},
    "purchase_invoice": {"title": "采购发票匹配"},
    "purchase_accrual": {"title": "月末采购暂估"},
    # 月末暂估页 5 张明细小表各自的导出（用户 2026-09-19：按钮放在各表「打印本表」旁边）
    "purchase_accrual_opening": {"title": "期初暂估明细"},
    "purchase_accrual_invoice": {"title": "本月到票明细"},
    "purchase_accrual_same_month": {"title": "当月到货到票明细"},
    "purchase_accrual_month": {"title": "本月暂估明细"},
    "purchase_accrual_cross": {"title": "本月回冲明细"},
    "purchase_difference": {"title": "本月采购暂估差异"},
    "material_cost_difference": {"title": "材料成本差异明细"},
    "purchase_reversal": {"title": "发票暂估回冲明细"},
    "sales_invoice": {"title": "销售发票匹配"},
    "sales_uninvoiced": {"title": "月末已出库未开票"},
    "sales_voucher": {"title": "销售结转成本凭证"},
    "purchase_voucher": {"title": "采购暂估回冲凭证"},
    "material_issue_voucher": {"title": "生产领料（投料）凭证"},
    "finished_receipt_voucher": {"title": "完工入库结转凭证"},
    "sales_ledger": {"title": "销售台账（成本·毛利）"},
    "profit_statement": {"title": "损益表"},
    "inventory_valuation": {"title": "期末库存（平台计价）"},
    "inventory_valuation_summary": {"title": "期末库存汇总（按物料分类，平台计价）"},
    "work_order_cost": {"title": "工单成本"},
    "work_order": {"title": "生产工单"},
    "order_material": {"title": "工单投料"},
    "order_completion": {"title": "工单报工"},
    "work_order_ledger": {"title": "工单台账"},
    "material_issue_ledger": {"title": "投料台账"},
    "work_report_ledger": {"title": "报工台账"},
    "work_order_detail": {"title": "工单明细"},
    "material_check": {"title": "投料核对"},
    "deemed_finished": {"title": "视同完工清单"},
    "reconciliation": {"title": "勾稽结果"},
    "workshop_settlement": {"title": "车间结转"},
    "workshop_cost": {"title": "车间费用"},
    "workshop_cost_summary": {"title": "车间费用汇总（按车间）"},
    "opening_inventory": {"title": "期初存货"},
    "ending_inventory": {"title": "期末存货"},
    "bom": {"title": "BOM表"},
    "pulp_opening": {"title": "浆粕期初成品"},
    "pulp_flow": {"title": "浆粕成品流水"},
    "pulp_ending": {"title": "浆粕期末成品"},
    "pulp_material": {"title": "浆粕原料流水"},
    "pulp_ledger": {"title": "浆粕流水"},
}



def _uploaded_ledger_rows(period, kind, queryset):
    """按用户上传文件的格式导出台账：
    - 格式（表头列名/顺序）取该类型行数最多的上传文件（即主文件）；
    - 表头通过 ALIASES 别名匹配自动定位（兼容模板含标题行/说明行的情况）；
    - 数据取数据库该类型全部记录，按上传表头逐列映射字段输出。
    无上传文件或读取失败时返回 (None, None, None)。
    """
    import csv as _csv
    import io as _io
    from pathlib import Path

    def _raw_rows(path):
        suffix = path.suffix.lower()
        if suffix in (".xlsx", ".xlsm"):
            from openpyxl import load_workbook

            wb = load_workbook(path, read_only=True, data_only=True)
            try:
                return [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
            finally:
                wb.close()
        if suffix == ".csv":
            raw = path.read_bytes()
            decoded = None
            for enc in ("utf-8-sig", "gb18030", "utf-8"):
                try:
                    decoded = raw.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            if decoded is None:
                return []
            return [list(r) for r in _csv.reader(_io.StringIO(decoded))]
        return []

    # 找到行数最多的上传批次作为格式基准
    best_raw = None
    best_count = -1
    for batch in (
        period.import_batches.filter(kind=kind)
        .exclude(file="")
        .order_by("-created_at")
    ):
        try:
            if not Path(batch.file.path).is_file():
                continue
            raw = _raw_rows(Path(batch.file.path))
        except Exception:
            continue
        if raw and len(raw) > best_count:
            best_raw = raw
            best_count = len(raw)
    if not best_raw:
        return None, None, None

    # 用 ALIASES 匹配定位真正的表头行（跳过标题行/说明行）
    alias_map = {}
    for field, names in ALIASES[kind].items():
        for name in names:
            alias_map[normalize_header(name)] = field
    best = None
    for idx, row in enumerate(best_raw[:15]):
        hits = sum(1 for v in row if normalize_header(v) in alias_map)
        if best is None or hits > best[0]:
            best = (hits, idx)
    if best is None or best[0] < 2:
        return None, None, None
    header = best_raw[best[1]]
    fields = [alias_map.get(normalize_header(h)) for h in header]

    rows_out = []
    for obj in queryset:
        row = []
        for f in fields:
            if f:
                v = getattr(obj, f, None)
                row.append(v)
            else:
                row.append("")
        rows_out.append(tuple(row))
    return header, rows_out, best_count




def _uploaded_inventory_raw(period):
    """读最近一次成功上传的库存流水原始文件，按上传原样返回 (sheet名, 表头, 数据行)。

    以最近一次上传的批次为准（不拼接多个批次，避免表头列顺序差异导致导出与上传不一致），
    表头通过 ALIASES 匹配定位；数据取该文件全部数据行，
    单价/金额保持用户上传的原值，不受结算回填影响；兼容 xlsx 与 csv。
    无上传文件时返回 None。"""
    import csv as _csv
    import io as _io
    from pathlib import Path

    def _read(path):
        suffix = path.suffix.lower()
        if suffix in (".xlsx", ".xlsm"):
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True, data_only=True)
            try:
                return [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
            finally:
                wb.close()
        if suffix == ".csv":
            raw = path.read_bytes()
            for enc in ("utf-8-sig", "gb18030", "utf-8"):
                try:
                    return [list(r) for r in _csv.reader(_io.StringIO(raw.decode(enc)))]
                except UnicodeDecodeError:
                    continue
        return []

    alias_map = {}
    for _f, _names in ALIASES["inventory"].items():
        for _n in _names:
            alias_map[normalize_header(_n)] = _f

    batches = list(
        period.import_batches.filter(kind="inventory")
        .exclude(file="")
        .order_by("-created_at", "-id")
    )
    for b in batches:
        fp = Path(b.file.path)
        if not fp.is_file():
            continue
        rows = _read(fp)
        if not rows:
            continue
        # 定位真正表头行（跳过标题/说明行）
        best = None
        for idx, row in enumerate(rows[:15]):
            hits = sum(1 for v in row if normalize_header(v) in alias_map)
            if best is None or hits > best[0]:
                best = (hits, idx)
        if best is None or best[0] < 2:
            continue
        header = rows[best[1]]
        data = rows[best[1] + 1:]
        # 单价/金额列回填：取该批次库存流水（按导入顺序），用结算后的单价/金额覆盖上传原值；
        # 尚无单价的行（如调拨/库存调整等内部流转、数量为0）按该物料月末加权价兜底回填，确保全表单价/金额有值
        txns = list(InventoryTransaction.objects.filter(batch=b).order_by("id"))
        if data and txns and len(txns) == len(data):
            price_idx = None
            amount_idx = None
            for _i, _h in enumerate(header):
                _f = alias_map.get(normalize_header(_h))
                if _f == "unit_price":
                    price_idx = _i
                elif _f == "amount":
                    amount_idx = _i
            if price_idx is not None or amount_idx is not None:
                val_prices = {
                    v.material_code: v.weighted_average_price
                    for v in period.inventory_valuations.all()
                }
                filled = []
                for _i, _row in enumerate(data):
                    _r = list(_row)
                    _t = txns[_i]
                    _up = _t.unit_price
                    _am = _t.amount
                    if _up is None:
                        _vp = val_prices.get(_t.material_code)
                        if _vp is not None:
                            _up = _vp
                            _am = (Decimal(_t.quantity or 0) * Decimal(_vp)).quantize(Decimal("0.01"))
                    if price_idx is not None:
                        _r[price_idx] = _up
                    if amount_idx is not None:
                        _r[amount_idx] = _am
                    filled.append(_r)
                data = filled
        return ("库存流水", header, data)
    return None


def _accrual_detail_opening_rows(details):
    """期初暂估明细（与页面同一口径：上一期期末仍未到票，首个期间取期初台账）。"""
    return [
        (
            r.get("receipt_month"),
            r.get("document_date"),
            r.get("document_no"),
            r.get("supplier"),
            r.get("material_code"),
            r.get("material_name"),
            r.get("receipt_quantity"),
            r.get("invoiced_quantity"),
            r.get("outstanding_quantity"),
            r.get("unit_price"),
            r.get("amount"),
            r.get("status"),
        )
        for r in details["opening"]
    ]


def _accrual_detail_month_rows(details):
    """本月暂估明细（归属本月、截至月末未到票）。"""
    return [
        (
            r.get("receipt_month"),
            r.get("document_date"),
            r.get("document_no"),
            r.get("supplier"),
            r.get("material_code"),
            r.get("material_name"),
            r.get("receipt_quantity"),
            r.get("invoiced_quantity"),
            r.get("outstanding_quantity"),
            r.get("unit_price"),
            r.get("amount"),
            r.get("status"),
        )
        for r in details["month_accrual"]
    ]


def _accrual_detail_invoice_rows(details):
    """本月到票明细（逐笔核销＝当月到货到票＋本月回冲）。"""
    return [
        (
            r.invoice.invoice_period,
            r.invoice.supplier_name or r.invoice.supplier_code,
            r.material_code,
            r.material_name,
            "%s · %s%s"
            % (
                r.receipt_month.strftime("%Y-%m") if r.receipt_month else "",
                r.document_no or "—",
                "" if r.transaction else "（期初台账）",
            ),
            r.matched_quantity,
            r.provisional_unit_price,
            r.provisional_amount,
            r.invoice_unit_price,
            r.invoice_amount,
            r.difference_amount,
            "本月回冲" if r.cross_period else "当月到货到票",
        )
        for r in details["invoice"]
    ]


def _accrual_detail_same_month_rows(details):
    """当月到货到票明细。"""
    return [
        (
            r.invoice.invoice_period,
            r.invoice.supplier_name or r.invoice.supplier_code,
            r.material_code,
            r.material_name,
            "%s · %s"
            % (
                r.receipt_month.strftime("%Y-%m") if r.receipt_month else "",
                r.document_no or "—",
            ),
            r.matched_quantity,
            r.provisional_unit_price,
            r.provisional_amount,
            r.invoice_unit_price,
            r.invoice_amount,
            r.difference_amount,
        )
        for r in details["same_month"]
    ]


def _accrual_detail_cross_rows(details):
    """本月回冲明细（冲回的以前暂估）。"""
    return [
        (
            r.invoice.invoice_period,
            r.invoice.supplier_name or r.invoice.supplier_code,
            r.material_code,
            r.material_name,
            "%s · %s%s"
            % (
                r.receipt_month.strftime("%Y-%m") if r.receipt_month else "",
                r.document_no or "—",
                "" if r.transaction else "（期初台账）",
            ),
            r.matched_quantity,
            r.provisional_unit_price,
            r.provisional_amount,
            r.invoice_unit_price,
            r.invoice_amount,
            r.difference_amount,
        )
        for r in details["cross"]
    ]


#: 月末暂估页 5 张明细小表的导出定义（用户 2026-09-19）
#: ⚠️ 用户 2026-09-19：本月到票 / 当月到货到票 / 本月回冲 三张表不再出「发票号」字段
#: （页面表格和这里的导出表头都去掉；列索引随之下移）。
ACCRUAL_DETAIL_EXPORT_KINDS = {
    "purchase_accrual_opening": {
        "sheet": "期初暂估明细",
        "headers": ["归属月份", "入库日期", "入库单号", "供应商", "物料编码", "物料名称",
                    "入库数量", "累计到票", "尚未到票", "暂估单价", "暂估金额", "状态"],
        "amount_index": 10,
        "total_note": "合计 {n} 行",
        "rows": _accrual_detail_opening_rows,
    },
    "purchase_accrual_invoice": {
        "sheet": "本月到票明细",
        "headers": ["开票期间", "供应商", "物料编码", "物料名称", "核销对象",
                    "核销数量", "暂估单价", "暂估金额", "发票单价", "发票金额", "暂估价差", "期间性质"],
        "amount_index": 9,
        "total_note": "合计 {n} 行（发票金额＝当月到货到票 ＋ 本月回冲）",
        "rows": _accrual_detail_invoice_rows,
    },
    "purchase_accrual_same_month": {
        "sheet": "当月到货到票明细",
        "headers": ["开票期间", "供应商", "物料编码", "物料名称", "入库单",
                    "核销数量", "暂估单价", "暂估金额", "发票单价", "发票金额", "暂估价差"],
        "amount_index": 9,
        "total_note": "合计 {n} 行（发票金额）",
        "rows": _accrual_detail_same_month_rows,
    },
    "purchase_accrual_month": {
        "sheet": "本月暂估明细",
        "headers": ["归属月份", "入库日期", "入库单号", "供应商", "物料编码", "物料名称",
                    "入库数量", "累计到票", "尚未到票", "暂估单价", "暂估金额", "状态"],
        "amount_index": 10,
        "total_note": "合计 {n} 行",
        "rows": _accrual_detail_month_rows,
    },
    "purchase_accrual_cross": {
        "sheet": "本月回冲明细",
        "headers": ["开票期间", "供应商", "物料编码", "物料名称", "冲回的入库/期初台账",
                    "核销数量", "暂估单价", "冲回暂估金额", "发票单价", "发票金额", "暂估价差"],
        "amount_index": 7,
        "total_note": "合计 {n} 行；发票金额合计 {invoice_total}",
        "rows": _accrual_detail_cross_rows,
    },
}


def _listing_rows(request, period, kind):
    """按 kind 返回 (工作表名, 表头, 数据行) 列表，筛选逻辑与列表页一致。"""
    def _pulp_mat_price(code):
        """浆粕原料出库单价：历史单价出库价优先、入库价兜底（与浆粕成本计算口径一致）。"""
        hp = period.historical_prices.filter(material_code=code).order_by("-id").first()
        if not hp:
            return None
        if hp.issue_price:
            return Decimal(hp.issue_price)
        if hp.receipt_price:
            return Decimal(hp.receipt_price)
        return None

    def _pulp_unit_price(m):
        if m.unit_price is not None:
            return m.unit_price
        return _pulp_mat_price(m.material_code)

    def _pulp_amount(m, price):
        if m.amount:
            return m.amount
        if price is not None:
            return (Decimal(m.quantity or 0) * Decimal(price)).quantize(Decimal("0.01"))
        return m.amount

    def _pulp_product_unit(code):
        """浆粕成品完工单位成本（用于回填成品流水单价）。"""
        try:
            from settlement.services.pulp_costing import collect_pulp_costs
            costs = collect_pulp_costs(period)
        except Exception:
            return None
        p = costs.get("products", {}).get(code or "")
        return p.get("unit") if p else None

    sheets = []

    # ===== 台账公共生成器：投料台账（上传原样+单价）/ 报工台账（上传原样）=====
    def _mi_ledger_sheet(include_amount=False, workshop=""):
        """投料台账：上传原样列 + 单价列（月末加权价→历史单价兜底）；无上传回退数据库明细+单价。

        workshop 非空时只导出该车间（工单编号 → 工单主档的生产部门编号归一后的车间）的行，
        供「工单投料（按投料车间）」每个车间页签自己的「导出Excel」使用。
        """
        from settlement.services.imports import normalize_workshop

        _ws_orders = None
        if workshop:
            _ws_orders = {
                _no for _no, _dept in period.work_orders.values_list(
                    "order_no", "production_dept_code"
                )
                if normalize_workshop(_dept or "") == workshop
            }
        header, raw_rows, _ = _uploaded_ledger_rows(
            period, ImportBatch.Kind.MATERIAL_ISSUE, period.material_issues.all()
        )
        _mi_prices = {}
        for _v in period.inventory_valuations.all():
            _mi_prices.setdefault(_v.material_code, Decimal(_v.weighted_average_price or 0))
        for _hp in period.historical_prices.all():
            if _hp.issue_price is not None:
                _mi_prices.setdefault(_hp.material_code, Decimal(_hp.issue_price))
            elif _hp.receipt_price is not None:
                _mi_prices.setdefault(_hp.material_code, Decimal(_hp.receipt_price))

        def _mi_unit_price(_code):
            _p = _mi_prices.get(_code or "")
            return Decimal(_p).quantize(Decimal("0.0001")) if _p else None

        if header:
            _alias = {}
            for _f, _names in ALIASES["material_issue"].items():
                for _n in _names:
                    _alias[normalize_header(_n)] = _f
            _mat_idx = None
            for _i, _h in enumerate(header):
                if _alias.get(normalize_header(_h)) == "material_code":
                    _mat_idx = _i
                    break
            _qty_idx = None
            for _i, _h in enumerate(header):
                if _alias.get(normalize_header(_h)) == "actual_issue_quantity":
                    _qty_idx = _i
                    break
            _extra = ["单价"] + (["金额"] if include_amount else [])
            _header2 = list(header) + _extra
            _rows2 = []
            # 按车间过滤：先找到「工单编号」列（按别名匹配表头）
            _wo_idx = None
            if _ws_orders is not None:
                for _i, _h in enumerate(header):
                    if _alias.get(normalize_header(_h)) == "work_order_no":
                        _wo_idx = _i
                        break
            for _row in raw_rows:
                if _ws_orders is not None:
                    if _wo_idx is None:
                        _rows2 = []
                        break
                    if str(_row[_wo_idx] or "") not in _ws_orders:
                        continue
                _price = _mi_unit_price(
                    str(_row[_mat_idx]) if _mat_idx is not None and _row[_mat_idx] else ""
                )
                _qty = (
                    Decimal(str(_row[_qty_idx]).replace(",", ""))
                    if _qty_idx is not None and _row[_qty_idx] not in (None, "")
                    else Decimal("0")
                )
                _amount = (
                    (_price * _qty).quantize(Decimal("0.01"))
                    if _price and _qty else None
                )
                _row2 = tuple(_row) + (_price,)
                if include_amount:
                    _row2 = _row2 + (_amount,)
                _rows2.append(_row2)
            return _header2, _rows2
        issues = period.material_issues.all()
        issues = _text_filters(
            request,
            issues,
            {"order_no": "work_order_no", "material": "material_code"},
        )
        if _ws_orders is not None:
            issues = issues.filter(work_order_no__in=_ws_orders)
        return (
            ["生产任务编号", "工序编号", "工序名称", "物料编码", "物料名称", "物料分类",
             "物料属性", "物料规格", "投料数量", "已回撤数量", "实际投料数量", "投料单位",
             "辅助数量", "辅助单位", "标识码", "批次号", "质量状态", "投料仓位",
             "被替代物料编号", "被替代物料名称", "投料时间", "投料人员", "执行人员",
             "工单编号", "主产出物料编号", "主产出物料名称", "销售订单编号", "投料记录编号",
             "单价"],
            [
                (
                    item.task_no, item.operation_no, item.operation_name, item.material_code,
                    item.material_name, item.material_category, item.material_attribute,
                    item.material_spec, item.issue_quantity, item.returned_quantity,
                    item.actual_issue_quantity, item.issue_uom, item.auxiliary_quantity,
                    item.auxiliary_uom, item.identifier, item.batch_no, item.quality_status,
                    item.issue_location, item.replaced_material_code, item.replaced_material_name,
                    item.issue_time, item.issuer, item.executor, item.work_order_no,
                    item.main_output_code, item.main_output_name, item.sales_order_no,
                    item.record_no, _mi_unit_price(item.material_code),
                )
                for item in issues.order_by("id")
            ],
        )

    def _wr_ledger_sheet():
        """报工台账：上传原样列；无上传回退数据库明细。"""
        header, raw_rows, _ = _uploaded_ledger_rows(
            period, ImportBatch.Kind.WORK_REPORT, period.work_reports.all()
        )
        if header:
            return header, raw_rows
        reports = period.work_reports.all()
        reports = _text_filters(
            request,
            reports,
            {"order_no": "work_order_no", "material": "material_code",
             "quality_status": "quality_status"},
        )
        return (
            ["报工时间", "生产任务编号", "工序编号", "工序名称", "报工物料编码", "报工物料名称",
             "生产人员", "质量状态", "不良原因", "报工数量", "报工单位", "辅助数量", "辅助单位",
             "批次号", "标识码", "流转卡编号", "工时", "工时单位", "备注", "部门", "设备",
             "工单编号", "工单物料编码", "工单物料名称", "订单号", "纸管颜色", "纺纱班别",
             "含水率", "批号流水", "织布序号", "织布机台", "产出仓位", "工单部门"],
            [
                (
                    item.report_time, item.task_no, item.operation_no, item.operation_name,
                    item.material_code, item.material_name, item.operator, item.quality_status,
                    item.defect_reason, item.quantity, item.uom, item.auxiliary_quantity,
                    item.auxiliary_uom, item.batch_no, item.identifier, item.route_card_no,
                    item.work_hours, item.work_hours_uom, item.remark, item.department,
                    item.equipment, item.work_order_no, item.work_order_material_code,
                    item.work_order_material_name, item.order_no, item.paper_tube_color,
                    item.spinning_shift, item.moisture_content, item.dye_batch_flow,
                    item.weaving_sequence, item.weaving_machine, item.output_location,
                    item.work_order_department,
                )
                for item in reports.order_by("id")
            ],
        )

    if kind == "inventory_matches":
        # 库存流水匹配页的导出：列与页面一致（含期间 / 匹配到的单据 / 处理方式 / 匹配状态），
        # 筛选逻辑与页面共用 _inventory_match_queryset（老链接带的参数照旧生效）
        txns = _inventory_match_queryset(request, period).order_by(
            "txn_date", "document_no", "id"
        )
        sheets.append(
            (
                "库存流水匹配",
                [
                    "期间", "库存单据", "业务日期", "匹配类别", "业务类型", "方向",
                    "仓库", "物料编码", "物料名称", "数量", "供应商/客户", "订单号",
                    "匹配到的单据", "处理方式", "匹配状态",
                ],
                (
                    (
                        _inventory_match_month(item.txn_date),
                        item.document_no,
                        item.txn_date,
                        _inventory_match_category(item),
                        item.business_type,
                        item.get_direction_display(),
                        item.warehouse,
                        item.material_code,
                        item.material_name,
                        item.quantity,
                        item.supplier_name
                        or item.customer_name
                        or item.supplier_code
                        or item.customer_code,
                        item.order_no,
                        _inventory_match_allocation_text(item),
                        item.get_match_mode_display(),
                        item.get_match_status_display(),
                    )
                    for item in txns
                ),
            )
        )
        return sheets

    if kind == "inventory_ledger":
        raw_sheet = _uploaded_inventory_raw(period)
        if raw_sheet is not None:
            # 按用户上传原样导出（数据取原始文件，不经数据库回填）
            sheets.append(raw_sheet)
            return sheets

        txns = period.inventory_transactions.select_related("batch", "material")
        txns = _date_filter(request, txns, "txn_date", "txn_date")
        txns = _text_filters(
            request,
            txns,
            {
                "document_no": "document_no",
                "business_type": "business_type",
                "material": "material_code",
                "warehouse": "warehouse",
                "order_no": "order_no",
            },
        )
        txns = _choice_filter(request, txns, "direction", "direction")
        party = str(request.GET.get("party", "")).strip()
        if party:
            txns = txns.filter(
                Q(supplier_code__icontains=party)
                | Q(supplier_name__icontains=party)
                | Q(customer_code__icontains=party)
                | Q(customer_name__icontains=party)
            )
        txns = txns.order_by("txn_date", "document_no", "id")
        sheets.append(
            (
                "库存流水",
                [
                    "创建时间", "物料名称", "物料编号", "分类", "物料规格", "仓库名称",
                    "批号", "操作数量", "主单位", "辅助数量", "辅助单位", "备注",
                    "业务单据编号", "业务类型", "供应商名称", "单价", "金额",
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
                    for item in txns
                ),
            )
        )

    elif kind == "purchase_contract":
        contracts = period.purchase_contracts
        contracts = _text_filters(
            request,
            contracts,
            {"contract_no": "contract_no", "material": "material_code"},
        )
        contracts = _party_filter(
            request, contracts, "supplier", "supplier_code", "supplier_name"
        )
        contracts = _choice_filter(request, contracts, "match_mode", "match_mode")
        contracts = contracts.order_by("contract_date", "contract_no", "id")
        sheets.append(
            (
                "采购合同",
                ["收货期间", "已签收数量", "合同号", "合同日期", "供应商", "品名编码", "品名", "合同数量",
                 "剩余数量", "不含税单价", "金额", "来源", "处理方式"],
                (
                    (
                        item.receipt_period,
                        item.received_quantity,
                        item.contract_no,
                        item.contract_date,
                        item.supplier_name or item.supplier_code,
                        item.material_code,
                        item.material_name,
                        item.quantity,
                        item.remaining_quantity,
                        item.unit_price_excl_tax,
                        item.contract_amount,
                        "自动承接" if item.is_system_generated else "导入",
                        item.get_match_mode_display(),
                    )
                    for item in contracts
                ),
            )
        )

    elif kind == "purchase_invoice":
        # 与页面同一口径：供应商 / 合同号筛选生效（页面导出链接会把这两个参数带过来）
        invoices = _purchase_invoice_query_filters(
            request, period.purchase_invoices.all()
        ).order_by("invoice_date", "invoice_no", "id")
        invoice_totals = purchase_invoice_summary(invoices)
        invoices = invoices.annotate(
            matched_total=Sum("matchallocation__matched_quantity")
        ).prefetch_related(
            "matchallocation_set__inventory_transaction__period",
            "matchallocation_set__opening_purchase_accrual",
        )

        def _invoice_matches(invoice):
            parts = []
            for a in invoice.matchallocation_set.all():
                if a.inventory_transaction:
                    parts.append(
                        f"{a.inventory_transaction.period.month:%Y-%m}·"
                        f"{a.inventory_transaction.document_no or '流水'}:{a.matched_quantity}"
                    )
                else:
                    parts.append(
                        f"{a.opening_purchase_accrual.original_period:%Y-%m}·"
                        f"{a.opening_purchase_accrual.document_no}(期初暂估):{a.matched_quantity}"
                    )
            return "；".join(parts)

        sheets.append(
            (
                "采购发票匹配",
                ["收货期间", "发票号", "收票日期", "发票期间", "合同号", "供应商", "物料编码",
                 "物料名称", "单位", "签收数量", "已匹配数量", "未匹配数量", "不含税金额",
                 "匹配的入库/期初暂估", "处理方式"],
                [
                    (
                        item.receipt_period,
                        item.invoice_no,
                        item.invoice_date,
                        item.invoice_period,
                        item.contract_no,
                        item.supplier_name or item.supplier_code,
                        item.material_code,
                        item.material_name,
                        item.uom,
                        item.quantity,
                        item.matched_total or 0,
                        item.remaining_quantity,
                        item.amount_excl_tax,
                        _invoice_matches(item),
                        item.get_match_mode_display(),
                    )
                    for item in invoices
                ]
                + [
                    (
                        "合计",
                        "",
                        "",
                        "",
                        "",
                        f"{invoice_totals['invoice_count']} 张发票",
                        "",
                        "",
                        "",
                        invoice_totals["quantity"],
                        invoice_totals["matched_quantity"],
                        invoice_totals["remaining_quantity"],
                        invoice_totals["amount"],
                        f"已核销金额 {invoice_totals['matched_amount']}；"
                        f"完全未核销 {invoice_totals['unmatched_invoice_count']} 张"
                        f"（{invoice_totals['unmatched_amount']}）",
                        "",
                    )
                ],
            )
        )

    elif kind == "purchase_accrual":
        rows = _filter_accrual_rows(
            request, build_purchase_accrual_rows(period)
        )
        sheets.append(
            (
                "月末采购暂估",
                ["归属期间", "流水期间", "入库日期", "入库单号", "供应商", "物料编码", "物料名称",
                 "入库数量", "已到票数量", "未到票数量", "暂估单价", "未到票金额",
                 "状态", "匹配合同"],
                (
                    (
                        r.receipt_month,
                        getattr(r, "flow_month", None),
                        r.document_date,
                        r.document_no,
                        r.supplier_name or r.supplier_code,
                        r.material_code,
                        r.material_name,
                        r.receipt_quantity,
                        r.invoiced_quantity,
                        r.outstanding_quantity,
                        r.provisional_unit_price,
                        r.provisional_amount,
                        r.status,
                        r.contract_nos,
                    )
                    for r in rows
                ),
            )
        )

    elif kind in ACCRUAL_DETAIL_EXPORT_KINDS:
        # 「到票与暂估」下面 5 张明细小表：各自单独导出（用户 2026-09-19）
        all_rows = build_purchase_accrual_rows(period)
        details = purchase_accrual_detail_tables(period, all_rows)
        # 每张表自己的筛选条件也套到导出上（用户 2026-09-19）：筛选后导出＝所见
        detail_filters = _accrual_detail_filters(request, details)
        detail_key = {
            "purchase_accrual_invoice": "invoice",
            "purchase_accrual_same_month": "same_month",
            "purchase_accrual_month": "month_accrual",
            "purchase_accrual_cross": "cross",
        }.get(kind)
        if detail_key:
            details = {**details, detail_key: detail_filters[detail_key]["rows"]}
        sheet_rows = ACCRUAL_DETAIL_EXPORT_KINDS[kind]["rows"](details)

        def _money_of(item):
            if isinstance(item, dict):
                return item.get("amount") or Decimal("0")
            return item.provisional_amount or Decimal("0")

        source = {
            "purchase_accrual_opening": details["opening"],
            "purchase_accrual_invoice": details["invoice"],
            "purchase_accrual_same_month": details["same_month"],
            "purchase_accrual_month": details["month_accrual"],
            "purchase_accrual_cross": details["cross"],
        }[kind]
        if kind == "purchase_accrual_invoice":
            # 本月到票明细的金额列是「发票金额」
            total_amount = sum((r.invoice_amount or Decimal("0") for r in source), Decimal("0"))
        elif kind == "purchase_accrual_cross":
            # 回冲明细的金额列是「冲回暂估金额」，另外发票金额合计放备注
            total_amount = sum((r.provisional_amount or Decimal("0") for r in source), Decimal("0"))
        elif kind == "purchase_accrual_same_month":
            total_amount = sum((r.invoice_amount or Decimal("0") for r in source), Decimal("0"))
        else:
            total_amount = sum((_money_of(r) for r in source), Decimal("0"))
        headers = ACCRUAL_DETAIL_EXPORT_KINDS[kind]["headers"]
        rows = list(sheet_rows)
        if rows:
            # 合计行：首列写「合计」、金额列放合计、末列写口径备注
            total_row = [""] * len(headers)
            total_row[0] = "合计"
            total_row[ACCRUAL_DETAIL_EXPORT_KINDS[kind]["amount_index"]] = (
                Decimal(total_amount).quantize(Decimal("0.01"))
            )
            total_row[-1] = ACCRUAL_DETAIL_EXPORT_KINDS[kind]["total_note"].format(
                n=len(source),
                invoice_total=sum(
                    (r.invoice_amount or Decimal("0") for r in source), Decimal("0")
                ).quantize(Decimal("0.01"))
                if kind == "purchase_accrual_cross"
                else "",
            )
            rows.append(tuple(total_row))
        sheets.append(
            (ACCRUAL_DETAIL_EXPORT_KINDS[kind]["sheet"], headers, rows)
        )

    elif kind == "purchase_receipt":
        # 「采购入库」页导出：与页面同口径（含 入库期间 列，并套用页面筛选条件）
        sheet_rows, _receipt_filters = _filtered_purchase_receipt_rows(request, period)
        sheets.append(
            (
                "采购入库",
                ["入库期间", "入库日期", "入库单号", "业务类型", "供应商", "合同号",
                 "物料编码", "物料名称",
                 "规格", "单位", "数量", "暂估单价", "暂估金额", "暂估依据",
                 "累计到票数量", "尚未到票数量", "最近到票期间", "状态"],
                (
                    (
                        r["receipt_month_label"],
                        r["txn_date"],
                        r["document_no"],
                        r["business_type"],
                        r["supplier_name"],
                        r["contract_no"],
                        r["material_code"],
                        r["material_name"],
                        r["specification"],
                        r["uom"],
                        r["quantity"],
                        r["provisional_unit_price"],
                        r["provisional_amount"],
                        r["price_source"],
                        r["invoiced_quantity"],
                        r["outstanding_quantity"],
                        r["latest_invoice_period"],
                        r["status"],
                    )
                    for r in sheet_rows
                ),
            )
        )

    elif kind == "purchase_difference":
        rows = build_purchase_difference_rows(period)
        invoice_no = str(request.GET.get("invoice_no", "")).strip().lower()
        receipt_month = str(request.GET.get("receipt_month", "")).strip()
        document_no = str(request.GET.get("document_no", "")).strip().lower()
        supplier = str(request.GET.get("supplier", "")).strip().lower()
        material = str(request.GET.get("material", "")).strip().lower()
        status = str(request.GET.get("status", "")).strip()
        cross_period = str(request.GET.get("cross_period", "")).strip()
        if invoice_no:
            rows = [r for r in rows if invoice_no in (r.invoice.invoice_no or "").lower()]
        if receipt_month:
            rows = [r for r in rows if r.receipt_month.strftime("%Y-%m") == receipt_month]
        if document_no:
            rows = [r for r in rows if document_no in (r.document_no or "").lower()]
        if supplier:
            rows = [r for r in rows if supplier in (f"{r.invoice.supplier_code} {r.invoice.supplier_name}").lower()]
        if material:
            rows = [r for r in rows if material in (f"{r.material_code} {r.material_name}").lower()]
        if status:
            rows = [r for r in rows if r.status == status]
        if cross_period in {"yes", "no"}:
            rows = [r for r in rows if r.cross_period == (cross_period == "yes")]
        sheets.append(
            (
                "采购暂估差异",
                ["发票号", "发票日期", "供应商", "物料编码", "物料名称", "入库期间", "归属期间",
                 "入库单号", "核销数量", "暂估单价", "发票单价", "暂估金额", "发票金额",
                 "价差", "是否跨期", "状态"],
                (
                    (
                        r.invoice.invoice_no,
                        r.invoice.invoice_date,
                        r.invoice.supplier_name or r.invoice.supplier_code,
                        r.material_code,
                        r.material_name,
                        r.receipt_month,
                        r.attribution_month or r.receipt_month,
                        r.document_no,
                        r.matched_quantity,
                        r.provisional_unit_price,
                        r.invoice_unit_price,
                        r.provisional_amount,
                        r.invoice_amount,
                        r.difference_amount,
                        "是" if r.cross_period else "否",
                        r.status,
                    )
                    for r in rows
                ),
            )
        )

    elif kind == "purchase_reversal":
        rows = build_purchase_reversal_rows(
            period,
            invoice_no=request.GET.get("invoice_no", "").strip(),
            supplier=request.GET.get("supplier", "").strip(),
            material=request.GET.get("material", "").strip(),
            cross_period=request.GET.get("cross_period", "").strip(),
        )
        sheets.append(
            (
                "发票暂估回冲明细",
                ["发票号", "发票日期", "发票期间", "核销对象类型", "原入库单号",
                 "原入库月份", "归属期间", "原入库日期", "物料编码", "物料名称", "供应商",
                 "核销数量", "暂估单价", "暂估金额", "发票单价", "发票金额",
                 "价差", "该入库单入库数量", "该入库单累计已核销数量",
                 "核销后剩余未到票", "是否跨期", "状态"],
                (
                    (
                        r["invoice_no"],
                        r["invoice_date"],
                        r["invoice_period"],
                        r["target_type"],
                        r["document_no"],
                        r["receipt_month"],
                        r.get("attribution_month") or r["receipt_month"],
                        r["document_date"],
                        r["material_code"],
                        r["material_name"],
                        r["supplier_name"],
                        r["matched_quantity"],
                        r["provisional_unit_price"],
                        r["provisional_amount"],
                        r["invoice_unit_price"],
                        r["invoice_amount"],
                        r["difference_amount"],
                        r["receipt_quantity"],
                        r["cumulative_quantity"],
                        r["outstanding_quantity"],
                        "是" if r["cross_period"] else "否",
                        r["status"],
                    )
                    for r in rows
                ),
            )
        )

    elif kind == "material_cost_difference":
        # 进「材料成本差异」科目的价差明细：口径与「采购暂估 / 回冲凭证」完全一致，
        # 因此本表合计 = 凭证里「材料成本差异」的金额，可直接对账。
        diff_rows, diff_totals = build_cost_difference_rows(period)
        sheets.append(
            (
                "材料成本差异明细",
                ["归属", "发票号", "发票日期", "发票期间", "供应商", "物料大类",
                 "物料编码", "物料名称", "入库单号", "到货月份", "核销数量",
                 "暂估单价", "暂估金额", "发票单价", "发票金额", "价差",
                 "是否跨期", "暂估来源"],
                [
                    (
                        r["group"],
                        r["invoice_no"],
                        r["invoice_date"],
                        r["invoice_period"],
                        r["supplier"],
                        r["category"],
                        r["material_code"],
                        r["material_name"],
                        r["document_no"],
                        r["receipt_month"],
                        r["matched_quantity"],
                        r["provisional_unit_price"],
                        r["provisional_amount"],
                        r["invoice_unit_price"],
                        r["invoice_amount"],
                        r["difference_amount"],
                        "是" if r["cross_period"] else "否",
                        r["price_source"],
                    )
                    for r in diff_rows
                ]
                + [
                    (
                        "合计", "", "", "", "", "", "", "", "", "",
                        diff_totals["matched_quantity"],
                        "", diff_totals["provisional_amount"],
                        "", diff_totals["invoice_amount"],
                        diff_totals["difference_amount"],
                        "", "",
                    )
                ],
            )
        )

        def _agg(key_fn):
            agg = {}
            for r in diff_rows:
                key = key_fn(r)
                entry = agg.setdefault(key, {"rows": 0, "prov": Decimal("0"),
                                             "inv": Decimal("0"), "diff": Decimal("0")})
                entry["rows"] += 1
                entry["prov"] += r["provisional_amount"]
                entry["inv"] += r["invoice_amount"]
                entry["diff"] += r["difference_amount"]
            return agg

        by_supplier = _agg(lambda r: r["supplier"])
        sheets.append(
            (
                "按供应商汇总",
                ["供应商", "行数", "暂估金额", "发票金额", "价差"],
                [
                    (k, v["rows"], v["prov"], v["inv"], v["diff"])
                    for k, v in sorted(by_supplier.items(),
                                       key=lambda kv: -abs(kv[1]["diff"]))
                ]
                + [
                    ("合计", diff_totals["count"], diff_totals["provisional_amount"],
                     diff_totals["invoice_amount"], diff_totals["difference_amount"])
                ],
            )
        )

        by_material = _agg(lambda r: (r["material_code"], r["material_name"]))
        sheets.append(
            (
                "按物料汇总",
                ["物料编码", "物料名称", "行数", "暂估金额", "发票金额", "价差"],
                [
                    (k[0], k[1], v["rows"], v["prov"], v["inv"], v["diff"])
                    for k, v in sorted(by_material.items(),
                                       key=lambda kv: -abs(kv[1]["diff"]))
                ]
                + [
                    ("合计", "", diff_totals["count"], diff_totals["provisional_amount"],
                     diff_totals["invoice_amount"], diff_totals["difference_amount"])
                ],
            )
        )

    elif kind == "sales_uninvoiced":
        # 台账口径（用户 2026-09-21 方案A）：明细＝销售台账「当月发货未开票挂发出商品」的行；
        # 老链接里带的筛选参数照旧生效（出库期间 / 销售订单号 / 客户 / 产品 / 以前月份）；
        # 其中「期间 / 客户」走与页面同一个筛选函数（用户 2026-09-27），保证页面＝导出。
        rows = filter_sales_uninvoiced_rows(
            build_sales_uninvoiced_rows(period),
            issue_month=str(request.GET.get("issue_month", "")).strip(),
            customer=str(request.GET.get("customer", "")).strip(),
        )
        sales_order_no = str(request.GET.get("sales_order_no", "")).strip().lower()
        material = str(request.GET.get("material", "")).strip().lower()
        cross_period = str(request.GET.get("cross_period", "")).strip()
        if sales_order_no:
            rows = [r for r in rows if sales_order_no in (r.order_no or "").lower()]
        if material:
            rows = [
                r for r in rows
                if material in f"{r.material_code} {r.material_name}".lower()
            ]
        if cross_period in {"yes", "no"}:
            rows = [r for r in rows if r.prior_period == (cross_period == "yes")]
        sheets.append(
            (
                "月末未开票",
                ["出库期间", "销售订单号", "客户", "销售渠道", "销售公司", "销售大类",
                 "产品编码", "产品名称", "规格型号", "单位", "数量", "成本单价",
                 "销售成本", "未开票天数", "是否以前月份出库"],
                (
                    (
                        r.out_month,
                        r.order_no,
                        r.customer_name,
                        r.channel,
                        r.company,
                        r.sales_category,
                        r.material_code,
                        r.material_name,
                        r.specification,
                        r.uom,
                        r.quantity,
                        r.unit_cost,
                        r.cost_amount,
                        r.outstanding_days,
                        "是" if r.prior_period else "否",
                    )
                    for r in rows
                ),
            )
        )
        # 导出＝整页内容（用户 2026-09-27）：明细之外再带上「勾稽表」与「统计表」，
        # 这样导出的文件与页面看到的三块表一致。
        check_rows = build_goods_issued_check_rows(period)
        sheets.append(
            (
                "发出商品滚动勾稽",
                ["项目", "笔数", "数量", "金额", "说明"],
                tuple(
                    (
                        item["label"],
                        item["count"],
                        item["quantity"],
                        item["amount"],
                        item["note"],
                    )
                    for item in check_rows
                ),
            )
        )
        statistics = build_sales_cost_statistics(period)
        sheets.append(
            (
                "发出商品及销售成本统计表",
                ["项目分类"] + list(statistics["month_headers"]),
                tuple(
                    (row["label"],) + tuple(row["cells"] or [])
                    for row in statistics["rows"]
                ),
            )
        )

    elif kind == "purchase_voucher":
        # 采购暂估 / 暂估回冲凭证：凭证分录 + 三张附件（与页面口径一致）
        _pv = build_purchase_voucher(
            period,
            group=request.GET.get("group", "") or "category",
            detail_supplier=request.GET.get("supplier", ""),
            detail_material=request.GET.get("material", ""),
        )
        _rows = []
        for _block in _pv["blocks"]:
            for _row in _block["rows"]:
                _rows.append(
                    (
                        _pv["month"],
                        _row["summary"],
                        _row["account"],
                        _row["quantity"],
                        _row["debit"],
                        _row["credit"],
                    )
                )
        _rows.append(
            (
                _pv["month"], "合计", "", _pv["totals"]["quantity"],
                _pv["totals"]["debit"], _pv["totals"]["credit"],
            )
        )
        sheets.append(
            (
                "凭证分录",
                ["凭证期间", "摘要", "科目", "数量", "借方", "贷方"],
                _rows,
            )
        )
        sheets.append(
            (
                "附件1-本月到货明细",
                ["入库单号", "入库日期", "到货月份", "供应商", "物料编码", "物料名称",
                 "物料大类", "入库数量", "入库金额", "暂估单价", "单价来源",
                 "本次到票数量", "已到票金额", "未到票数量", "暂估金额", "状态"],
                [
                    (
                        r["document_no"], r["document_date"],
                        r["receipt_month"].strftime("%Y-%m") if r["receipt_month"] else "",
                        r["supplier"], r["material_code"], r["material_name"],
                        r["category"], r["quantity"], r["amount"],
                        r["provisional_unit_price"], r["price_source"],
                        r["matched_quantity"], r["matched_amount"],
                        r["unmatched_quantity"], r["unmatched_amount"], r["status"],
                    )
                    for r in _pv["receipt_rows"]
                ],
            )
        )
        sheets.append(
            (
                "附件2-暂估回冲明细",
                ["发票号", "开票日期", "开票期间", "供应商", "核销对象", "入库单号",
                 "到货月份", "物料编码", "物料名称", "核销数量", "暂估单价", "暂估金额",
                 "发票单价", "发票金额", "价差"],
                [
                    (
                        r["invoice_no"], r["invoice_date"],
                        r["invoice_period"].strftime("%Y-%m") if r["invoice_period"] else "",
                        r["supplier"], r["target_type"], r["document_no"],
                        r["receipt_month"].strftime("%Y-%m") if r["receipt_month"] else "",
                        r["material_code"], r["material_name"], r["matched_quantity"],
                        r["provisional_unit_price"], r["provisional_amount"],
                        r["invoice_unit_price"], r["invoice_amount"],
                        r["difference_amount"],
                    )
                    for r in _pv["reversal_rows"]
                ],
            )
        )
        sheets.append(
            (
                "附件3-应付暂估余额",
                ["供应商", "未到票数量", "暂估余额", "行数"],
                [
                    (r["supplier"], r["quantity"], r["amount"], r["rows"])
                    for r in _pv["balance_rows"]
                ],
            )
        )
        sheets.append(
            (
                "附件4-暂估账龄",
                ["到货月份", "暂估金额", "行数", "占比%"],
                [
                    (r["month"], r["amount"], r["rows"], round(r["ratio"], 2))
                    for r in _pv["aging_rows"]
                ],
            )
        )

    elif kind == "finished_receipt_voucher":
        # 完工入库结转凭证（用户 2026-09-25）：摘要按车间、科目按物料分类。
        # 与凭证中心同一套函数，保证页面/导出一致。
        _fr_period = period
        _fr_month = str(
            request.GET.get("receipt_month")
            or request.GET.get("voucher_month")
            or ""
        ).strip()
        if _fr_month:
            try:
                _y, _m = _fr_month.split("-")
                _fr_period = (
                    AccountingPeriod.objects.filter(
                        month=date(int(_y), int(_m), 1)
                    ).first()
                    or period
                )
            except (TypeError, ValueError):
                _fr_period = period
        _fr = build_finished_receipt_voucher(_fr_period)
        sheets.append(
            (
                "凭证分录",
                ["凭证期间", "摘要", "科目", "辅助说明", "数量", "借方", "贷方"],
                [
                    (
                        _fr["month"], row["summary"], row["account"], row["aux"],
                        row["quantity"], row["debit"], row["credit"],
                    )
                    for row in _fr["rows"]
                ],
            )
        )
        sheets.append(
            (
                "附件1-逐单完工入库",
                ["期间", "车间", "工单编号", "产品编码", "产品名称", "物料分类",
                 "完工数量", "完工单位成本", "直接材料", "直接人工", "制造费用",
                 "金额", "取数来源", "有入库流水"],
                [
                    (
                        row["period"], row["workshop_label"], row["order_no"],
                        row["product_code"], row["product_name"], row["category"],
                        row["finished_quantity"], row["finished_unit_cost"],
                        row["material"], row["wages"], row["overhead"],
                        row["amount"], row["source"],
                        "是" if row["has_flow"] else "否（流水缺失）",
                    )
                    for row in _fr["detail_rows"] + _fr["pulp_rows"]
                ],
            )
        )
        sheets.append(
            (
                "附件2-车间×物料分类",
                ["车间", "物料分类", "摘要", "工单数", "数量", "金额", "借方科目"],
                [
                    (
                        row["workshop_label"], row["category"], row["summary"],
                        row["order_count"], row["quantity"], row["amount"],
                        row["account_debit"],
                    )
                    for row in _fr["category_rows"]
                ],
            )
        )
        sheets.append(
            (
                "附件3-车间×成本要素",
                ["车间", "工单数", "物料分类数", "物料分类", "完工数量",
                 "直接材料", "直接人工", "制造费用", "金额", "取数来源"],
                [
                    (
                        row["workshop_label"], row["order_count"],
                        row["category_count"], row["categories"], row["quantity"],
                        row["material"] if row["element_split"] else "—",
                        row["wages"] if row["element_split"] else "—",
                        row["overhead"] if row["element_split"] else "—",
                        row["amount"], row["source"],
                    )
                    for row in _fr["workshop_rows"]
                ],
            )
        )

    elif kind == "material_issue_voucher":
        # 生产领料（投料）凭证：凭证分录 + 领料明细 + 按车间×物料大类汇总（与页面口径一致）
        _mi_period = period
        _mi_month = str(request.GET.get("voucher_month", "")).strip()
        if _mi_month:
            try:
                _y, _m = _mi_month.split("-")
                _mi_period = (
                    AccountingPeriod.objects.filter(
                        month=date(int(_y), int(_m), 1)
                    ).first()
                    or period
                )
            except (TypeError, ValueError):
                _mi_period = period
        _mi = build_material_issue_voucher(
            _mi_period,
            detail_workshop=request.GET.get("workshop", ""),
            detail_material=request.GET.get("material", ""),
        )
        _mi_rows = [
            (
                _mi["month"], _row["summary"], _row["account"],
                _row["quantity"], _row["debit"], _row["credit"],
            )
            for _row in _mi["rows"]
        ]
        sheets.append(
            ("凭证分录", ["凭证期间", "摘要", "科目", "数量", "借方", "贷方"], _mi_rows)
        )
        sheets.append(
            (
                "附件1-领料明细",
                ["日期", "来源", "工单号/单据号", "车间", "物料大类", "物料编码",
                 "物料名称", "仓库/库位", "数量", "单价", "金额", "取数依据"],
                [
                    (
                        r["date"], r["source"],
                        r["work_order_no"] or r["document_no"], r["workshop_label"],
                        r["category"], r["material_code"], r["material_name"],
                        r["warehouse"], r["quantity"], r["unit_price"], r["amount"],
                        r["basis"],
                    )
                    for r in _mi["detail_rows"]
                ],
            )
        )
        sheets.append(
            (
                "附件2-车间×物料大类",
                ["车间", "物料大类", "摘要", "数量", "金额", "借方科目", "贷方科目", "行数"],
                [
                    (
                        r["workshop_label"], r["category"], r["summary"], r["quantity"],
                        r["amount"], r["account_debit"], r["account_credit"], r["rows"],
                    )
                    for r in _mi["category_rows"]
                ],
            )
        )
        sheets.append(
            (
                "附件3-贷方科目汇总",
                ["科目", "物料大类", "摘要", "数量", "金额", "流水行数"],
                [
                    (r["account"], r["category"], r["summary"], r["quantity"],
                     r["amount"], r["rows"])
                    for r in _mi["credit_rows"]
                ],
            )
        )

    elif kind == "sales_voucher":
        # 销售结转成本凭证：与页面口径一致（凭证分录 + 结转类型×销售大类附件）
        _vm, _vf = _sales_voucher_filters(request)
        _voucher = build_sales_voucher_entries(
            period,
            month=_vm,
            block_filters=_vf,
            group=request.GET.get("group", "") or "category",
        )
        _rows = []
        for _block in _voucher["blocks"]:
            for _row in _block["entries"]:
                _rows.append(
                    (
                        _voucher["month"], _row["summary"], _row["account"],
                        _row["quantity"], _row["debit"], _row["credit"],
                    )
                )
        _rows.append(
            (
                _voucher["month"], "合计", "",
                _voucher["totals"]["quantity"], _voucher["totals"]["debit"],
                _voucher["totals"]["credit"],
            )
        )
        _status_labels = {
            item["value"]: item["label"] for item in _voucher["status_options"]
        }
        sheets.append(
            (
                "凭证分录",
                ["凭证期间", "摘要", "科目", "数量", "借方", "贷方"],
                _rows,
            )
        )
        sheets.append(
            (
                "附件-结转类型与销售大类",
                ["凭证期间", "会计科目", "贷方科目", "结转类型", "销售大类",
                 "出库期间", "开票状态", "开票期间", "数量", "金额"],
                (
                    (
                        _voucher["month"] if row["block_index"] >= 0 else "",
                        row["account"],
                        row.get("credit_account", ""),
                        row["carry_type"],
                        row["category"],
                        row["block_out"] if row["block_index"] >= 0 else "",
                        (
                            _status_labels.get(
                                row["block_status"], row["block_status"] or ""
                            )
                            if row["block_index"] >= 0
                            else ""
                        ),
                        row["block_invoice"] if row["block_index"] >= 0 else "",
                        row["quantity"],
                        row["amount"],
                    )
                    for row in _voucher["summary_rows"]
                ),
            )
        )

    elif kind == "sales_ledger":
        # 与「销售台账」页共用同一套筛选：页面上筛选后导出的就是当前看到的行。
        items = sales_ledger_queryset(request, period)
        sheets.append(
            (
                "销售台账（成本·毛利）",
                ["序号", "出库期间", "开票状态", "开票期间", "销售渠道", "销售公司",
                 "订单单号", "客户名称", "销售大类", "产品编码", "产品名称", "规格型号",
                 "计量单位", "终端销售收入", "终端销售单价", "射阳兰邦销售收入",
                 "射阳兰邦销售单价", "销售数量", "销售成本单价", "销售成本",
                 "销售毛利", "销售毛利率", "结转类型"],
                (
                    (
                        item.row_no,
                        item.out_month.strftime("%Y-%m") if item.out_month else "",
                        item.get_invoice_status_display(),
                        item.invoice_month.strftime("%Y-%m") if item.invoice_month else "",
                        item.channel,
                        item.company,
                        item.order_no,
                        item.customer_name or item.customer_code,
                        item.sales_category,
                        item.material_code,
                        item.material_name,
                        item.specification,
                        item.uom,
                        item.terminal_amount,
                        item.terminal_unit_price,
                        item.lb_amount,
                        item.lb_unit_price,
                        item.quantity,
                        item.cost_price,
                        item.cost_amount,
                        # 与页面一致：毛利 = 终端销售收入 - 销售成本（缺成本按 0）
                        (
                            Decimal(item.terminal_amount or 0)
                            - Decimal(item.cost_amount or 0)
                        ).quantize(Decimal("0.01")),
                        (
                            (Decimal(item.terminal_amount or 0)
                             - Decimal(item.cost_amount or 0))
                            / Decimal(item.terminal_amount)
                            * 100
                        ).quantize(Decimal("0.01"))
                        if item.terminal_amount
                        else None,
                        item.carry_type,
                    )
                    for item in items
                ),
            )
        )

    elif kind == "profit_statement":
        sheets.extend(build_profit_statement_sheet(period)[0])

    elif kind == "work_order_detail":
        # 工单明细：优先输出用户上传的工单明细（格式=上传文件原样，数据=上传的 WorkOrderDetail 记录）；
        # 无上传时回退系统计算导出
        header, raw_rows, _ = _uploaded_ledger_rows(
            period, ImportBatch.Kind.WORK_ORDER_DETAIL, period.work_order_details.all()
        )
        if header:
            sheets.append(("工单明细", header, raw_rows))
            return sheets
        # 工单明细：创建时间/车间/工单号/计划完工/实际完工/完成状态/计划数量/投料数量/入库数量
        from django.db.models import Max as _Max

        # 各工单投料数量（实际投料量合计）
        issue_agg = {
            row["work_order_no"]: row["total"]
            for row in period.material_issues.exclude(work_order_no="")
            .values("work_order_no")
            .annotate(total=Sum("actual_issue_quantity"))
        }
        # 各工单报工入库数量（"报工"入库流水合计，按单据号=工单号关联）
        receipt_agg = {
            row["document_no"]: row["total"]
            for row in period.inventory_transactions.filter(business_type__in=FINISHED_RECEIPT_TYPES)
            .exclude(document_no="")
            .values("document_no")
            .annotate(total=Sum("quantity"))
        }
        # 各工单实际完工时间（报工记录最大报工时间）
        report_time_agg = {
            row["work_order_no"]: row["max_time"]
            for row in period.work_reports.exclude(work_order_no="")
            .values("work_order_no")
            .annotate(max_time=_Max("report_time"))
        }
        # 完成状态（与工单成本页一致：累计完工/在产品）
        results = {
            item.work_order.order_no: item
            for item in period.work_order_cost_results.select_related("work_order").all()
        }

        def _completion(wo, finished_qty, wip_qty):
            if finished_qty <= 0 and wip_qty <= 0:
                return "未开工"
            if wip_qty <= 0:
                return "已完成"
            return "部分完成"

        rows = []
        for wo in period.work_orders.all().order_by("production_dept_code", "order_no"):
            finished_qty = Decimal(0)
            wip_qty = Decimal(0)
            if wo.order_no in results:
                finished_qty = Decimal(results[wo.order_no].finished_quantity or 0)
                wip_qty = Decimal(results[wo.order_no].wip_quantity or 0)
            else:
                finished_qty = Decimal(receipt_agg.get(wo.order_no, 0) or 0)
            rows.append((
                wo.created_at,
                WORKSHOP_LABELS.get(wo.production_dept_code, wo.production_dept_code),
                wo.order_no,
                wo.planned_finish,
                report_time_agg.get(wo.order_no),
                _completion(wo, finished_qty, wip_qty),
                Decimal(wo.plan_quantity or 0),
                Decimal(issue_agg.get(wo.order_no, 0) or 0),
                Decimal(receipt_agg.get(wo.order_no, 0) or 0),
            ))
        sheets.append(
            (
                "工单明细",
                ["创建时间", "生产车间", "工单编号", "计划完工时间", "实际完工时间",
                 "完成状态", "计划数量", "投料数量", "入库数量"],
                rows,
            )
        )

    elif kind == "work_order_ledger":
        # 工单台账：格式/字段按用户上传文件原样，数据=本期全部工单；无上传时回退数据库明细
        header, raw_rows, _ = _uploaded_ledger_rows(
            period, ImportBatch.Kind.WORK_ORDER, period.work_orders.all()
        )
        if header:
            sheets.append(("工单台账", header, raw_rows))
            return sheets
        orders = period.work_orders.all()
        orders = _text_filters(
            request,
            orders,
            {
                "order_no": "order_no",
                "workshop": "production_dept_code",
                "status": "status",
            },
        )
        sheets.append(
            (
                "工单台账",
                ["工单编号", "工单类型", "标识码", "计划数量", "计划开始时间", "计划完工时间",
                 "优先级", "产线设备", "班次", "生产部门", "生产主管", "计划部门", "计划员",
                 "备注", "指定用料", "外部单据编号", "状态", "使用物料清单", "使用工艺路线",
                 "工序计划类型", "启用SOP", "生产模式", "产线编号", "订单号", "物料颜色"],
                (
                    (
                        item.order_no,
                        item.order_type,
                        item.identifier,
                        item.plan_quantity,
                        item.planned_start,
                        item.planned_finish,
                        item.priority,
                        item.line_equipment,
                        item.shift,
                        WORKSHOP_LABELS.get(item.production_dept_code, item.production_dept_code),
                        item.production_supervisor_code,
                        item.planning_dept_code,
                        item.planner_code,
                        item.remark,
                        item.specified_material,
                        item.external_document_no,
                        item.status,
                        "是" if item.use_bom else "否",
                        "是" if item.use_routing else "否",
                        item.process_plan_type,
                        "是" if item.enable_sop else "否",
                        item.production_mode,
                        item.line_code,
                        item.source_order_no,
                        item.material_color,
                    )
                    for item in orders.order_by("production_dept_code", "order_no")
                ),
            )
        )

    elif kind == "material_issue_ledger":
        _h, _r = _mi_ledger_sheet()
        sheets.append(("投料台账", _h, _r))
    elif kind == "work_report_ledger":
        _h, _r = _wr_ledger_sheet()
        sheets.append(("报工台账", _h, _r))
    elif kind == "work_order":
        # 「生产工单」页导出：与页面、与上传文件一致（25 列，列名/顺序同上传文件）
        orders = period.work_orders.all()
        orders = _text_filters(
            request,
            orders,
            {
                "order_no": "order_no",
                "workshop": "production_dept_code",
                "product": "identifier",
            },
        )
        status = str(request.GET.get("status", "")).strip()
        if status:
            orders = orders.filter(status__icontains=status)

        def _local_naive(value):
            """时间导出成「本地时间、不带时区」（Excel 不接受带时区的时间）。"""
            if value is None:
                return None
            if timezone.is_aware(value):
                value = timezone.localtime(value)
            return value.replace(tzinfo=None)

        def _order_cells(item):
            return (
                item.order_no,
                item.order_type,
                item.identifier,
                item.plan_quantity,
                _local_naive(item.planned_start),
                _local_naive(item.planned_finish),
                item.priority,
                item.line_equipment,
                item.shift,
                item.production_dept_code,
                item.production_supervisor_code,
                item.planning_dept_code,
                item.planner_code,
                item.remark,
                item.specified_material,
                item.external_document_no,
                item.status,
                "是" if item.use_bom else "否",
                "是" if item.use_routing else "否",
                item.process_plan_type,
                "是" if item.enable_sop else "否",
                item.production_mode,
                item.line_code,
                item.source_order_no,
                item.material_color,
            )

        sheets.append(
            (
                "生产工单",
                list(WORK_ORDER_COLUMNS),
                tuple(_order_cells(item) for item in orders.order_by("order_no")),
            )
        )

    elif kind == "work_order_ledger":
        # 工单台账：格式/字段按用户上传文件原样，数据=本期全部工单；无上传时回退数据库明细
        header, raw_rows, _ = _uploaded_ledger_rows(
            period, ImportBatch.Kind.WORK_ORDER, period.work_orders.all()
        )
        if header:
            sheets.append(("工单台账", header, raw_rows))
            return sheets
        orders = period.work_orders.all()
        orders = _text_filters(
            request,
            orders,
            {
                "order_no": "order_no",
                "workshop": "production_dept_code",
                "status": "status",
            },
        )
        sheets.append(
            (
                "工单台账",
                ["工单编号", "工单类型", "标识码", "计划数量", "计划开始时间", "计划完工时间",
                 "优先级", "产线设备", "班次", "生产部门", "生产主管", "计划部门", "计划员",
                 "备注", "指定用料", "外部单据编号", "状态", "使用物料清单", "使用工艺路线",
                 "工序计划类型", "启用SOP", "生产模式", "产线编号", "订单号", "物料颜色"],
                (
                    (
                        item.order_no,
                        item.order_type,
                        item.identifier,
                        item.plan_quantity,
                        item.planned_start,
                        item.planned_finish,
                        item.priority,
                        item.line_equipment,
                        item.shift,
                        WORKSHOP_LABELS.get(item.production_dept_code, item.production_dept_code),
                        item.production_supervisor_code,
                        item.planning_dept_code,
                        item.planner_code,
                        item.remark,
                        item.specified_material,
                        item.external_document_no,
                        item.status,
                        "是" if item.use_bom else "否",
                        "是" if item.use_routing else "否",
                        item.process_plan_type,
                        "是" if item.enable_sop else "否",
                        item.production_mode,
                        item.line_code,
                        item.source_order_no,
                        item.material_color,
                    )
                    for item in orders.order_by("production_dept_code", "order_no")
                ),
            )
        )

    elif kind == "material_issue_ledger":
        _h, _r = _mi_ledger_sheet()
        sheets.append(("投料台账", _h, _r))
    elif kind == "work_report_ledger":
        _h, _r = _wr_ledger_sheet()
        sheets.append(("报工台账", _h, _r))
    elif kind == "work_order":
        orders = period.work_orders.all()
        orders = _text_filters(
            request,
            orders,
            {
                "order_no": "order_no",
                "workshop": "production_dept_code",
                "product": "identifier",
            },
        )
        status = str(request.GET.get("status", "")).strip()
        if status:
            orders = orders.filter(status__icontains=status)
        sheets.append(
            (
                "生产工单",
                ["工单编号", "工单类型", "车间/部门", "状态", "计划数量",
                 "计划开始时间", "计划完工时间", "产线/设备", "订单号", "备注"],
                (
                    (
                        item.order_no,
                        item.order_type,
                        WORKSHOP_LABELS.get(
                            item.production_dept_code, item.production_dept_code
                        ),
                        item.status,
                        item.plan_quantity,
                        item.planned_start,
                        item.planned_finish,
                        item.line_equipment,
                        item.source_order_no,
                        item.remark,
                    )
                    for item in orders.order_by(
                        "production_dept_code", "order_no"
                    )
                ),
            )
        )

    elif kind == "order_material":
        # 工单投料：字段与上传模板一致（原样）+ 单价列 + 金额列（金额=实际投料数量×单价）
        # 页面每个车间页签的「导出Excel」会带 ?workshop=<车间>，这里按车间过滤
        _h, _r = _mi_ledger_sheet(
            include_amount=True,
            workshop=str(request.GET.get("workshop", "")).strip(),
        )
        sheets.append(("工单投料", _h, _r))
    elif kind == "order_completion":
        # 工单报工：与页面明细表同一份数据（期间 + 上传的报工记录原样 + 6 个计算列）
        _reports = _order_completion_report_rows(request, period)
        _cols, _rows, _ = _order_completion_report_sheets(period, _reports)
        sheets.append(("工单报工", _cols, _rows))
        return sheets
    elif kind == "inventory_valuation":
        # 期末库存（逐物料）：与页面同一套筛选（期间 / 物料编码），并补「期间」列
        data = build_inventory_results_data(request, period)
        sheets.append(
            (
                "期末库存（平台计价）",
                ["期间", "物料编码", "物料名称", "物料分类", "期初数量", "期初金额",
                 "入库数量", "入库金额", "出库数量", "出库金额", "月末加权平均价",
                 "期末数量", "期末金额", "收敛"],
                (
                    (
                        item.period_label,
                        item.material_code,
                        item.material_name,
                        item.material_category,
                        item.opening_quantity,
                        item.opening_amount,
                        item.receipt_quantity,
                        item.receipt_amount,
                        item.issue_quantity,
                        item.issue_amount,
                        item.weighted_average_price,
                        item.ending_quantity,
                        item.ending_amount,
                        "是" if item.converged else "否",
                    )
                    for item in data["detail_rows"]
                ),
            )
        )

    elif kind == "inventory_valuation_summary":
        # 库存金额结转汇总（按物料分类）：与页面同一套筛选（期间 / 物料分类）+ 合计行
        data = build_inventory_results_data(request, period)
        totals = data["summary_totals"]
        sheets.append(
            (
                "期末库存汇总（按物料分类）",
                ["期间", "物料分类", "期初数量", "期初金额", "入库数量", "入库金额",
                 "出库数量", "出库金额", "期末数量", "期末金额"],
                [
                    (
                        row["period_label"], row["category"],
                        row["opening_quantity"], row["opening_amount"],
                        row["receipt_quantity"], row["receipt_amount"],
                        row["issue_quantity"], row["issue_amount"],
                        row["ending_quantity"], row["ending_amount"],
                    )
                    for row in data["summary_rows"]
                ]
                + [
                    (
                        "合计", "",
                        totals["opening_quantity"], totals["opening_amount"],
                        totals["receipt_quantity"], totals["receipt_amount"],
                        totals["issue_quantity"], totals["issue_amount"],
                        totals["ending_quantity"], totals["ending_amount"],
                    )
                ],
            )
        )

    elif kind == "ending_wip":
        # 「期末在产品」页导出：与页面同一套筛选（车间 / 产品编码或名称）
        data = build_ending_wip_data(
            period,
            workshop=str(request.GET.get("workshop", "")).strip(),
            material=str(request.GET.get("material", "")).strip(),
            status=str(request.GET.get("status", "")).strip(),
        )
        rows = data["rows"]
        workshop_rows = data["workshop_rows"]
        total = data["total"]
        sheets.append(
            (
                "期末在产品",
                ["车间", "起始期间", "工单编号", "工单状态", "产品编码", "产品名称",
                 "计划数量", "完工数量", "在产品数量", "直接材料", "人工费",
                 "制造费用"],
                tuple(
                    (
                        r["workshop"], r["start_period"], r["order_no"], r["order_status"],
                        r["product_code"], r["product_name"], r["plan_quantity"],
                        r["finished_quantity"], r["wip_quantity"], r["wip_material"],
                        r["wip_wages"], r["wip_overhead"],
                    )
                    for r in rows
                )
                + (
                    (
                        "合计", "", "", "", "", "", "", "", total["wip_quantity"],
                        total["wip_material"], total["wip_wages"],
                        total["wip_overhead"],
                    ),
                ),
            )
        )
        sheets.append(
            (
                "按车间汇总",
                ["车间", "在产品工单数", "在产品数量", "直接材料", "人工费",
                 "制造费用", "在产品成本"],
                tuple(
                    (
                        item["workshop"], item["order_count"], item["wip_quantity"],
                        item["wip_material"], item["wip_wages"],
                        item["wip_overhead"], item["amount"],
                    )
                    for item in workshop_rows
                )
                + (
                    (
                        "合计", total["order_count"], total["wip_quantity"],
                        total["wip_material"], total["wip_wages"],
                        total["wip_overhead"], total["amount"],
                    ),
                ),
            )
        )

    elif kind == "work_order_cost":
        results = period.work_order_cost_results.select_related("work_order")
        results = _text_filters(
            request,
            results,
            {
                "order_no": "work_order__order_no",
                "workshop": "work_order__production_dept_code",
                "product": "product_code",
            },
        )
        results = list(results)
        # 按结转顺序排序：车间按「车间结算」sequence（与页面一致），车间内按工单号
        ws_seq = {
            s.workshop: s.sequence
            for s in period.workshop_settlements.all()
        }
        results.sort(
            key=lambda item: (
                ws_seq.get(item.work_order.production_dept_code, 999),
                item.work_order.order_no,
            )
        )
        from settlement.services.workorder_costing import MANUFACTURING_FIELDS
        from settlement.models import BomLine, IssueProgress
        opening_map = {o.order_no: o for o in period.opening_wip.all()}

        # 投料进度与粗纱标志（与页面一致）
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
        # 投料台账跨期累计（用户 2026-10-16）：本期及以前各期的投料数量合计
        accum_issue_by_order = defaultdict(Decimal)
        for _p in AccountingPeriod.objects.filter(month__lte=period.month):
            for _row in (
                _p.material_issues.exclude(work_order_no="")
                .values("work_order_no", "actual_issue_quantity")
            ):
                accum_issue_by_order[_row["work_order_no"]] += Decimal(
                    _row["actual_issue_quantity"] or 0
                )

        # 工单明细（用户上传）：类别/计划完工/实际完工/完成状态
        work_order_details = {
            d.order_no: d
            for d in period.work_order_details.all()
        }
        from django.db.models import Max as _MaxDetail2

        report_time_agg = {
            row["work_order_no"]: row["max_time"]
            for row in period.work_reports.exclude(work_order_no="")
            .values("work_order_no")
            .annotate(max_time=_MaxDetail2("report_time"))
        }
        # 类别兜底：产品物料分类
        category_by_code = {}
        for code, cat in (
            period.inventory_transactions.exclude(category_name="")
            .order_by("material_code")
            .values_list("material_code", "category_name")
        ):
            category_by_code.setdefault(code, cat)
        bom_product = {}
        for line in period.bom_lines.exclude(material_code="").order_by("id"):
            bom_product.setdefault(
                line.work_order_no, (line.material_code, line.material_name)
            )
        resolve_product = build_product_resolver(period)
        # 投料物料单价（月末加权价优先，历史单价兜底）+ 工单投料数量聚合
        material_price = {}
        for v in period.inventory_valuations.all():
            material_price.setdefault(v.material_code, Decimal(v.weighted_average_price or 0))
        for hp in period.historical_prices.all():
            if hp.issue_price is not None:
                material_price.setdefault(hp.material_code, Decimal(hp.issue_price))
            elif hp.receipt_price is not None:
                material_price.setdefault(hp.material_code, Decimal(hp.receipt_price))
        issued_by_pair = defaultdict(Decimal)
        for row in period.material_issues.exclude(work_order_no="").exclude(material_code="").values(
            "work_order_no", "material_code", "actual_issue_quantity"
        ):
            issued_by_pair[(row["work_order_no"], row["material_code"])] += Decimal(
                row["actual_issue_quantity"] or 0
            )
        _CATEGORY_KEYWORDS2 = (
            ("面料", ("面料", "坯布", "织布", "摇粒绒", "螺纹", "绒布", "布", "绒")),
            ("纱线", ("纱线", "纱")),
            ("染色纤维", (
                "纤维", "长丝", "丝", "腈纶", "粘胶", "涤纶", "莱赛尔",
                "芳纶", "氨纶", "短纤", "棉纱",
            )),
            ("原料", ("原料", "棉", "麻")),
        )

        def _cat_from_name2(name):
            if not name:
                return ""
            for cat, kws in _CATEGORY_KEYWORDS2:
                if any(k in name for k in kws):
                    return cat
            return ""

        # BOM 100% 全粗纱的工单（「是否用粗纱」＝是）；循环外算一次，避免每单查库
        _all_roving_orders = all_roving_bom_orders(period)
        for item in results:
            order_no = item.work_order.order_no
            ws = item.work_order.production_dept_code
            # 期初在产品：数量/金额（金额=材料+制造费用+工资合计）
            opening = opening_map.get(order_no)
            item.opening_quantity = Decimal(opening.quantity if opening else Decimal("0"))
            item.opening_amount = Decimal(opening.total if opening else Decimal("0"))
            # 用户 2026-10-16：累计投料＝**投料台账跨期累计**（本期＋以前各期），
            #   不再用投料进度表（它常没同步，会出现"累计 < 本期"的矛盾）。
            item.accum_issue = accum_issue_by_order.get(order_no, Decimal(0))
            statuses = progress_status_by_order.get(order_no)
            item.issue_status = "、".join(sorted(statuses)) if statuses else "—"
            # 本期完工 / 累计完工 / 完工率（用户 2026-09-24 定稿口径，与页面同一套函数）
            attach_finish_progress([item], work_order_details)
            wip_qty = Decimal(item.wip_quantity or 0)
            # 本期投料/完工/完工率（与页面一致）
            item.current_issue = Decimal(item.accum_issue or 0)
            # 投料单价（保留 2 位小数，用户 2026-09-27）：
            #   ① 有投料 → 按实际投料的加权平均单价；② 没有投料 → 按 BOM 表折算。
            # 与「生产成本 / 工单成本」页共用同一套函数，保证两处一致。
            from settlement.services.workorder_costing import (
                bom_unit_price as _bom_unit_price,
                issue_unit_price as _issue_unit_price,
            )

            _price = _issue_unit_price(issued_by_pair, material_price, order_no)
            item.current_unit_price_from_bom = False
            if _price is None:
                _price = _bom_unit_price(
                    period, order_no, material_price,
                    bom_lines=list(period.bom_lines.filter(work_order_no=order_no)),
                )
                item.current_unit_price_from_bom = _price is not None
            # 两处都取不到价就留 None（页面/导出显示 —）
            item.current_unit_price = _price
            # 本期完工已由 attach_finish_progress 一并挂上
            if item.accum_finished <= Decimal("0") and wip_qty <= Decimal("0"):
                item.completion_status = "未开工"
            elif wip_qty <= Decimal("0"):
                item.completion_status = "已完成"
            else:
                item.completion_status = "部分完成"
            # 是否用粗纱（用户 2026-09-30 定稿）：BOM 100% 全粗纱才"是"，混料单"否"；所有车间都显示
            item.uses_roving = "是" if order_no in _all_roving_orders else "否"
            # 类别/计划完工/实际完工/完成状态（优先工单明细）
            detail = work_order_details.get(order_no)
            # 产品编码/名称：结算值优先，BOM/标识码兜底（统一 resolver）
            product_code, product_name = resolve_product(
                item.work_order, item.product_code, item.product_name
            )
            item.product_code = product_code
            if product_name:
                item.product_name = product_name
            # 类别：流水匹配优先 → BOM主产出码/物料名推断 → 明细类别 → 工单类型
            _bom = bom_product.get(order_no)
            item.wo_category = (
                category_by_code.get(product_code)
                or category_by_code.get(item.product_code or "")
                or (category_by_code.get(_bom[0]) if _bom and _bom[0] else "")
                or _cat_from_name2(_bom[1] if _bom else "")
                or (detail.category if detail is not None else "")
                or {
                    "RanZheng": "染色纤维",
                    "tefang": "纱线",
                    "BaoFu": "纱线",
                    "ZhenZhi": "面料",
                    "JiZhi": "面料",
                    "浆粕车间": "浆粕",
                }.get(ws, "")
                or item.work_order.order_type
                or ""
            )
            if detail is not None and detail.planned_finish:
                item.wo_planned_finish = detail.planned_finish
            else:
                item.wo_planned_finish = (
                    item.work_order.planned_finish.strftime("%Y-%m-%d %H:%M")
                    if item.work_order.planned_finish
                    else ""
                )
            if detail is not None and detail.actual_finish:
                item.wo_actual_finish = detail.actual_finish
            else:
                _t = report_time_agg.get(order_no)
                item.wo_actual_finish = _t.strftime("%Y-%m-%d %H:%M") if _t else ""
            if detail is not None and detail.completion_status:
                item.wo_completion_status = detail.completion_status
            else:
                item.wo_completion_status = item.completion_status

        # 标准单位成本与单位成本差异（与页面同一套函数，保证两处一致）
        attach_standard_unit_costs(period, results)

        # 与页面/结算工作簿一致：每个制造费用列后加分配率列，每车间末尾加合计行
        # 纺纱车间无天然气费用：单独导出纺纱时去掉天然气两列（与页面一致）
        workshop_param = (
            request.GET.get("workshop") or request.GET.get("w_workshop") or ""
        ).strip()
        hide_gas = workshop_param == "tefang"
        # 浆粕车间（用户 2026-09-21）：没有工单，按上传的浆粕流水补成品行（与页面一致）
        if not workshop_param or workshop_param == "浆粕车间":
            from settlement.services.workorder_costing import build_pulp_cost_rows

            pulp_rows = build_pulp_cost_rows(period)
            if workshop_param == "浆粕车间":
                results = pulp_rows
            else:
                results = list(results) + pulp_rows
        headers = [
            "核算期间", "工单编号", "车间", "类别", "产品编码", "产品名称",
            "计划完工时间", "实际完工时间", "完成状态", "计划数量",
            "期初数量", "期初金额",
            "本期投料", "投料单价", "累计投料", "投料状态",
            "本期完工", "累计完工", "本期完工率", "累计完工率", "是否用粗纱", "报工数量",
            "在产品数量", "折合产量", "材料成本", "水电", "水电分配率%", "蒸汽", "蒸汽分配率%",
            "天然气", "天然气分配率%", "折旧", "折旧分配率%", "机物料", "机物料分配率%",
            "染料助剂", "染料助剂分配率%", "加工费", "加工费分配率%", "工资", "工资分配率%", "工单总成本",
            "完工成本", "完工单位成本", "标准单位成本", "单位成本差异",
            "期末在产品数量", "在产品成本",
        ]
        if hide_gas:
            headers = [h for h in headers if h not in ("天然气", "天然气分配率%")]
        rows = []
        highlight = set()
        if results:
            mat_total = {}
            fee_total = {f: {} for f in MANUFACTURING_FIELDS}
            equiv_total = {}
            # 折合产量与结算层同一口径（用户 2026-09-30：机织在产品以 kg 计、完工换成 kg）
            from settlement.services.workorder_costing import (
                bom_kg_per_meter,
                equivalent_quantity,
            )

            _kg_per_m_map = bom_kg_per_meter(period)
            for item in results:
                ws = item.work_order.production_dept_code
                mat_total[ws] = mat_total.get(ws, Decimal("0")) + Decimal(
                    item.direct_material or 0
                )
                for f in MANUFACTURING_FIELDS:
                    fee_total[f][ws] = fee_total[f].get(ws, Decimal("0")) + Decimal(
                        getattr(item, f) or 0
                    )
                equiv_total[ws] = equiv_total.get(ws, Decimal("0")) + equivalent_quantity(
                    ws,
                    item.work_order.order_no,
                    item.finished_quantity,
                    item.wip_quantity,
                    _kg_per_m_map,
                )

            def _rate(fee_field, ws):
                denom = mat_total.get(ws, Decimal("0"))
                return (
                    (fee_total[fee_field].get(ws, Decimal("0")) / denom) * Decimal("100")
                    if denom
                    else Decimal("0")
                )

            def _rate_text(fee_field, ws):
                v = _rate(fee_field, ws)
                return "—" if not v else f"{v:.2f}%"

            def _wage_rate(item, ws):
                denom = equiv_total.get(ws, Decimal("0"))
                if not denom:
                    return Decimal("0")
                equivalent = equivalent_quantity(
                    ws,
                    item.work_order.order_no,
                    item.finished_quantity,
                    item.wip_quantity,
                    _kg_per_m_map,
                )
                return (equivalent / denom) * Decimal("100")

            def _wage_rate_text(item, ws):
                v = _wage_rate(item, ws)
                return "—" if not v else f"{v:.2f}%"

            def _unit_cost_text(value):
                """完工单位成本：保留两位小数；0（含四舍五入后为 0）显示 —。

                与页面 |dash_num:2 同一口径（ROUND_HALF_UP，和 Django floatformat 一致）。
                非零值仍然写数值（Excel 里可继续算，自带 #,##0.00 显示格式），
                只有 0/空才写 "—"。
                """
                if value is None or value == "":
                    return "—"
                try:
                    number = Decimal(str(value))
                except (InvalidOperation, ValueError, TypeError):
                    return value
                if number == 0:
                    return "—"
                rounded = number.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                return "—" if rounded == 0 else rounded

            def _pct_text(value, digits=2):
                """差异率（%，用户 2026-09-24）：保留两位小数；0/空显示 —。"""
                if value is None or value == "":
                    return "—"
                try:
                    number = Decimal(str(value))
                except (InvalidOperation, ValueError, TypeError):
                    return value
                quantum = Decimal("1." + "0" * int(digits)) if digits else Decimal("1")
                rounded = number.quantize(quantum, rounding=ROUND_HALF_UP)
                return "—" if rounded == 0 else "%s%%" % rounded

            SUM_FIELDS = (
                "direct_material", "water_power", "steam", "natural_gas",
                "depreciation", "consumables", "dye_auxiliary", "processing_fee",
                "wages", "total_cost", "finished_cost", "wip_cost",
            )
            subtotals = {}
            for item in results:
                ws = item.work_order.production_dept_code
                st = subtotals.setdefault(
                    ws, {"count": 0, **{f: Decimal("0") for f in SUM_FIELDS}}
                )
                st["count"] += 1
                for f in SUM_FIELDS:
                    st[f] += Decimal(getattr(item, f) or 0)

            def _detail_row(item):
                ws = item.work_order.production_dept_code
                if ws == "浆粕车间":
                    # 浆粕行没有工单字段：按成品行填（与页面同一套口径）
                    return [
                        f"{period.month:%Y-%m}",
                        item.work_order.order_no,
                        WORKSHOP_LABELS.get(ws, ws),
                        item.wo_category,
                        item.product_code,
                        item.product_name,
                        item.wo_planned_finish,
                        item.wo_actual_finish,
                        item.wo_completion_status,
                        item.plan_quantity,
                        item.opening_quantity,
                        item.opening_amount,
                        item.current_issue,
                        item.current_unit_price,
                        item.accum_issue,
                        item.issue_status,
                        item.current_finished,
                        item.accum_finished,
                        "100.00%" if item.current_finished else "—",
                        "100.00%" if item.accum_finished else "—",
                        item.uses_roving,
                        item.finished_quantity,
                        item.wip_quantity,
                        item.equivalent_qty,
                        item.direct_material,
                        item.water_power, "—",
                        item.steam, "—",
                    ] + ([] if hide_gas else [item.natural_gas, "—"]) + [
                        item.depreciation, "—",
                        item.consumables, "—",
                        item.dye_auxiliary, "—",
                        item.processing_fee, "—",
                        item.wages, "—",
                        item.total_cost,
                        item.finished_cost,
                        _unit_cost_text(item.finished_unit_cost),
                        "—",  # 标准单位成本（浆粕无工单/BOM）
                        "—",  # 单位成本差异
                        item.wip_quantity,
                        item.wip_cost,
                    ]
                rate_text = (
                    f"{item.finish_rate_pct:.2f}%" if item.finish_rate_pct else "—"
                )
                # 本期完工率／累计完工率：写成百分比文本（与页面显示完全一致：两位小数、0 → —）
                current_rate_text = finish_rate_text(item.current_finish_rate_pct)
                rate_text = finish_rate_text(item.finish_rate_pct)
                row = [
                    f"{period.month:%Y-%m}",
                    item.work_order.order_no,
                    WORKSHOP_LABELS.get(ws, ws),
                    item.wo_category,
                    item.product_code,
                    item.product_name,
                    item.wo_planned_finish,
                    item.wo_actual_finish,
                    item.wo_completion_status,
                    item.plan_quantity,
                    item.opening_quantity,
                    item.opening_amount,
                    item.current_issue,
                    item.current_unit_price,
                    item.accum_issue,
                    item.issue_status,
                    item.current_finished,
                    item.accum_finished,
                    current_rate_text,
                    rate_text,
                    item.uses_roving,
                    item.finished_quantity,
                    item.wip_quantity,
                    Decimal(item.finished_quantity or 0) + Decimal(item.wip_quantity or 0) * Decimal("0.5"),
                    item.direct_material,
                    item.water_power,
                    _rate_text("water_power", ws),
                    item.steam,
                    _rate_text("steam", ws),
                ]
                if not hide_gas:
                    row += [item.natural_gas, _rate_text("natural_gas", ws)]
                row += [
                    item.depreciation,
                    _rate_text("depreciation", ws),
                    item.consumables,
                    _rate_text("consumables", ws),
                    item.dye_auxiliary,
                    _rate_text("dye_auxiliary", ws),
                    item.processing_fee,
                    _rate_text("processing_fee", ws),
                    item.wages,
                    _wage_rate_text(item, ws),
                    item.total_cost,
                    item.finished_cost,
                    _unit_cost_text(item.finished_unit_cost),
                    _unit_cost_text(item.standard_unit_cost),
                    _pct_text(getattr(item, "unit_cost_diff_pct", None)),
                    item.wip_quantity,
                    item.wip_cost,
                ]
                return tuple(row)

            def _subtotal_row(ws, st):
                row = [f"{WORKSHOP_LABELS.get(ws, ws)}合计（{st['count']}张工单）"]
                # 列2-24：工单编号~折合产量（含期初数量/期初金额/报工数量/在产品数量/折合产量）
                row += [""] * 23
                row += [
                    st["direct_material"],
                    st["water_power"], "",
                    st["steam"], "",
                ]
                if not hide_gas:
                    row += [st["natural_gas"], ""]
                row += [
                    st["depreciation"], "",
                    st["consumables"], "",
                    st["dye_auxiliary"], "",
                    st["processing_fee"], "",
                    st["wages"], "",
                    st["total_cost"],
                    st["finished_cost"],
                    "",  # 完工单位成本
                    "",  # 标准单位成本
                    "",  # 单位成本差异
                    "",  # 期末在产品数量
                    st["wip_cost"],
                ]
                return tuple(row)

            current_ws = None
            for item in results:
                ws = item.work_order.production_dept_code
                if current_ws is not None and ws != current_ws:
                    rows.append(_subtotal_row(current_ws, subtotals[current_ws]))
                    highlight.add(len(rows))
                current_ws = ws
                rows.append(_detail_row(item))
            if current_ws is not None:
                rows.append(_subtotal_row(current_ws, subtotals[current_ws]))
                highlight.add(len(rows))

        sheets.append(( "工单成本", headers, rows, highlight ))


    elif kind == "standard_cost":
        # 标准成本（只出标准侧；与「标准成本」页面同一套函数，保证两处完全一致）
        _sc_workshop = (
            request.GET.get("workshop") or request.GET.get("w_workshop") or ""
        ).strip()
        _sc_rows, _sc_meta = build_standard_cost_rows(
            period,
            workshop_filter=_sc_workshop,
            order_filter=(request.GET.get("order_no") or "").strip(),
            product_filter=(request.GET.get("product") or "").strip(),
        )
        headers = [
            "期间", "车间", "工单编号", "产品编码", "产品名称",
            "标准材料成本", "标准人工成本", "标准制造费用", "标准单位成本",
        ]
        rows = []
        for row in _sc_rows:
            rows.append([
                f"{period.month:%Y-%m}",
                row["workshop_label"],
                row["order_no"],
                row["product_code"],
                row["product_name"],
                row["standard_material"],
                row["standard_labor"],
                row["standard_overhead"],
                row["standard_unit_cost"],
            ])
        _sc_totals = _standard_cost_totals(_sc_rows)
        # 列顺序对齐表头：期间｜车间｜工单编号｜产品编码｜产品名称｜
        #                  标准材料成本｜标准人工成本｜标准制造费用｜标准单位成本
        _sc_total_row = ["合计（%d 张工单，其中 %d 张标准成本为 —）"
                         % (_sc_totals["count"], _sc_totals["no_standard_price"])]
        _sc_total_row += [""] * 4                      # 车间 / 工单编号 / 产品编码 / 产品名称
        _sc_total_row += [
            _sc_totals["standard_material"], _sc_totals["standard_labor"],
            _sc_totals["standard_overhead"],
        ]
        _sc_total_row += [_sc_totals["standard_unit_cost"]]
        rows.append(_sc_total_row)
        _sc_highlight = {len(rows) - 1}
        sheets.append(("工单标准成本计算表", headers, rows, _sc_highlight))

        # 第二张：标准成本构成表（期间×车间；与页面「标准成本构成表」同一套函数，保证两处一致）
        # 页面上的「期间 / 车间」筛选（c_month / c_workshop）在这里同样生效
        _sc_comp = build_standard_cost_composition_rows(
            period,
            month_filter=(request.GET.get("c_month") or "").strip(),
            workshop_filter=(request.GET.get("c_workshop") or "").strip(),
        )
        comp_rows = []
        for row in _sc_comp:
            comp_rows.append([
                row["period_label"], row["label"], row["output_qty"],
                row["direct_material"], row["wages"],
                row["manufacturing_overhead"],
                row["material_pct"], row["labor_pct"], row["overhead_pct"],
            ])
        # 末两段：本期适用的标准占比（半年滚动平均）＋ 来源说明
        comp_rows.append([])
        comp_rows.append([
            "本期适用的标准占比（半年滚动）", _sc_meta["basis_label"],
            "取上一个完整半年各月占比的平均数", "", "", "", "", "", "",
        ])
        for code in standard_cost_workshops(period):
            _rate = _sc_meta["rates"].get(code) or {}
            if not _rate:
                continue
            comp_rows.append([
                "标准占比", WORKSHOP_LABELS.get(code, code), "", "", "", "",
                _rate.get("material_pct"), _rate.get("labor_pct"),
                _rate.get("overhead_pct"),
            ])
        sheets.append((
            "标准成本构成表",
            ["期间", "车间", "产量", "直接材料", "人工", "制造费用",
             "材料成本占比%", "人工成本占比%", "制造费用占比%"],
            comp_rows,
        ))

        # 第三张：BOM 明细（用户 2026-09-23）：把 BOM 子项逐行列出，
        # 说明「标准材料单位成本」怎么来的（与页面「BOM 明细表」同一套函数、同一口径）；
        # 用户 2026-09-30：范围＝本期全部 BOM 行（9944），列＝上传文件原样 + 系统列在右。
        _bom_columns, _bom_rows = build_bom_detail_table(period)
        attach_bom_cells(_bom_rows, _bom_columns)
        sheets.append((
            "BOM 明细",
            [col["name"] for col in _bom_columns],
            [
                [_bom_cell_for_excel(value) for value in row["cells"]]
                for row in _bom_rows
            ],
        ))

    elif kind == "standard_cost_bom":
        # BOM 明细表单独导出（用户 2026-09-24）：与「标准成本」页「BOM 明细表」同一套函数、
        # 同一口径；页面上的期间/车间/工单号筛选（b_period / b_workshop / b_order_no）同样生效。
        # 用户 2026-09-30：范围＝本期全部 BOM 行；列＝上传文件原样 + 系统列在右；导出不受分页影响。
        _bom_columns, _bom_rows = build_bom_detail_table(
            period,
            workshop_filter=(request.GET.get("b_workshop") or "").strip(),
            order_filter=(request.GET.get("b_order_no") or "").strip(),
            missing_only=str(request.GET.get("b_missing", "")).strip() in ("1", "on", "true"),
        )
        attach_bom_cells(_bom_rows, _bom_columns)
        sheets.append((
            "BOM 明细",
            [col["name"] for col in _bom_columns],
            [
                [_bom_cell_for_excel(value) for value in row["cells"]]
                for row in _bom_rows
            ],
        ))



    elif kind == "material_check":
        # ===== 投料核对报表：真实缺料工单（有报工、无投料、无期初、非粗纱）+ 供料匹配（单表）=====
        # 机织(JiZhi)、浆粕车间暂不纳入
        from collections import defaultdict as _dd_mc
        from settlement.models import OpeningWip as _OpeningWipMC

        _WS_SCOPE = ("RanZheng", "tefang", "BaoFu", "ZhenZhi", "JiZhi")
        _reported_mc = set(
            period.work_reports.exclude(work_order_no="")
            .values_list("work_order_no", flat=True)
        )
        _issued_mc = set(
            period.material_issues.exclude(work_order_no="")
            .values_list("work_order_no", flat=True)
        )
        _ws_of_mc = {
            o.order_no: o.production_dept_code
            for o in period.work_orders.all()
        }
        _roving_mc = set(
            period.bom_lines.filter(sub_item_code__startswith="BC-")
            .values_list("work_order_no", flat=True)
        )
        _roving_mc |= set(
            period.material_issues.filter(material_code__startswith="BC-")
            .exclude(work_order_no="")
            .values_list("work_order_no", flat=True)
        )
        _roving_mc |= set(
            period.issue_progress.filter(material_code__startswith="BC-")
            .exclude(work_order_no="")
            .values_list("work_order_no", flat=True)
        )

        def _fin_qty_mc(ono):
            return sum(
                (Decimal(x.quantity or 0) for x in period.work_reports.filter(work_order_no=ono)),
                Decimal("0"),
            )

        def _bom_consume_mc(ono, mat):
            return sum(
                (
                    Decimal(l.qty or 0) * _fin_qty_mc(ono)
                    for l in period.bom_lines.filter(work_order_no=ono, sub_item_code=mat)
                ),
                Decimal("0"),
            )

        # 真实缺料工单：有报工、无投料、无期初在产品、非粗纱、车间范围内
        _targets = []
        for _o in sorted(_reported_mc - _issued_mc):
            _ws = _ws_of_mc.get(_o)
            if _ws not in _WS_SCOPE:
                continue
            if _o in _roving_mc:
                continue
            _op = _OpeningWipMC.objects.filter(period=period, order_no=_o).first()
            if _op and Decimal(_op.material or 0) > 0:
                continue  # 有期初在产品 = 账面不缺料，不统计
            _r = period.work_order_cost_results.select_related("work_order").filter(
                work_order__order_no=_o
            ).first()
            _wb = sorted(
                {
                    str(x.batch_no or "")
                    for x in period.work_reports.filter(work_order_no=_o)
                    if x.batch_no
                }
            )
            _targets.append(
                {
                    "order_no": _o,
                    "ws": _ws,
                    "fin": _fin_qty_mc(_o),
                    "product": (_r.product_name if _r else ""),
                    "report_batch": "、".join(_wb),
                    "bom": list(period.bom_lines.filter(work_order_no=_o)),
                }
            )

        # 投料汇总（物料,工单）-> 数量 / 批次
        _issue_by_mc = _dd_mc(Decimal)
        _issue_batch_mc = _dd_mc(list)
        for _mi in period.material_issues.exclude(work_order_no="").exclude(material_code=""):
            _k = (_mi.material_code, _mi.work_order_no)
            _issue_by_mc[_k] += Decimal(_mi.actual_issue_quantity or 0)
            _issue_batch_mc[_k].append(str(_mi.batch_no or ""))

        # 匹配并生成单表行（每行 = 缺料工单 × 原料需求）
        _rows = []
        for _t in _targets:
            for _line in _t["bom"]:
                _mat = _line.sub_item_code
                _need = Decimal(_line.qty or 0) * _t["fin"]
                if _need <= 0:
                    continue
                _cands = []
                for (_mm, _so), _qty in _issue_by_mc.items():
                    if _mm != _mat or _so == _t["order_no"]:
                        continue
                    if _so in _roving_mc and _ws_of_mc.get(_so) == "tefang":
                        continue
                    _rem = _qty - _bom_consume_mc(_so, _mat)
                    if _rem > Decimal("0.0001"):
                        _cands.append((_so, _rem, _issue_batch_mc[(_mm, _so)]))
                _pref = _t["order_no"].split("-")[0]
                if _cands:
                    # 依据：同产品前缀 > 仅同原料；剩余多者优先
                    _cands.sort(
                        key=lambda c: (
                            0 if c[0].split("-")[0] == _pref else 1,
                            -float(c[1]),
                        )
                    )
                    _so, _rem, _batches = _cands[0]
                    _transfer = min(_need, _rem)
                    _same_prod = _so.split("-")[0] == _pref
                    _basis = "同产品系列" if _same_prod else "仅同原料"
                    _advice = "建议采用" if _same_prod else "需人工确认"
                    _so_batch = "、".join(sorted({b for b in _batches if b}))
                else:
                    _so, _transfer, _rem = "", Decimal("0"), Decimal("0")
                    _basis, _advice, _so_batch = "无供料", "BOM反推补投", ""
                _rows.append(
                    (
                        f"{period.month:%Y-%m}",
                        WORKSHOP_LABELS.get(_t["ws"], _t["ws"]),
                        _t["order_no"],
                        _t["product"],
                        _t["fin"],
                        _mat,
                        _line.sub_item_name,
                        _need,
                        _t["report_batch"],
                        _so or "(无供料)",
                        _so_batch,
                        _transfer,
                        _rem,
                        _basis,
                        _advice,
                    )
                )

        # ===== 按车间分组整理（含车间小计，Sheet1/Sheet2 共用车间顺序）=====
        from collections import OrderedDict as _od_ws
        _ws_label2code = {v: k for k, v in WORKSHOP_LABELS.items()}
        _ws_codes = list(
            period.workshop_settlements.order_by("sequence").values_list(
                "workshop", flat=True
            )
        )
        if not _ws_codes:
            _ws_codes = list(WORKSHOP_LABELS.keys())

        def _ws_key(_code):
            return _ws_codes.index(_code) if _code in _ws_codes else 999

        # Sheet1 投料核对：按车间分组 + 车间小计
        _by_ws1 = _od_ws()
        for _row in _rows:
            _code = _ws_label2code.get(_row[1], _row[1])
            _by_ws1.setdefault(_code, []).append(_row)
        _rows_sorted = []
        for _code in sorted(_by_ws1, key=_ws_key):
            _ws_rows1 = _by_ws1[_code]
            _rows_sorted.extend(_ws_rows1)
            _rows_sorted.append(
                (
                    "",
                    WORKSHOP_LABELS.get(_code, _code) + " 小计",
                    f"{len(_ws_rows1)} 条",
                    "",
                    sum((Decimal(r[4] or 0) for r in _ws_rows1), Decimal("0")),
                    "",
                    "",
                    sum((Decimal(r[7] or 0) for r in _ws_rows1), Decimal("0")),
                    "",
                    "",
                    "",
                    sum((Decimal(r[11] or 0) for r in _ws_rows1), Decimal("0")),
                    "",
                    "",
                    "",
                )
            )
        _rows = _rows_sorted

        sheets.append(
            (
                "投料核对",
                [
                    "期间", "车间", "缺料工单", "产品名称", "完工数量",
                    "原料编码", "原料名称", "需求数量", "缺料工单报工批次",
                    "供料工单", "供料工单投料批次",
                    "转出数量", "供料方扣减后剩余", "匹配依据", "建议",
                ],
                _rows,
            )
        )

        # ===== Sheet2：BOM反推补投对比（与核算层 _backfill_material_by_bom 同口径）=====
        # 有报工+无投料+无期初+非粗纱链 的染色/纺纱/包覆/针织工单，
        # 核算材料成本 = Σ(BOM单耗×报工数量×月末加权价，历史单价兜底)，
        # 即"总成本报表本期材料 − 投料记录金额"的差异来源。
        from settlement.models import (
            InventoryValuation as _InvValMC,
            HistoricalPrice as _HistPriceMC,
        )
        _mc_prices = {
            _v.material_code: Decimal(_v.weighted_average_price or 0)
            for _v in _InvValMC.objects.filter(period=period)
        }
        for _hp in _HistPriceMC.objects.filter(period=period):
            if _hp.issue_price is not None:
                _mc_prices.setdefault(_hp.material_code, Decimal(_hp.issue_price))
            elif _hp.receipt_price is not None:
                _mc_prices.setdefault(_hp.material_code, Decimal(_hp.receipt_price))

        _backfill_rows = []
        for _t in _targets:
            _fin = _t["fin"]
            if _fin <= 0:
                continue
            _mat_cost = Decimal("0")
            _no_price_parts = []
            for _line in _t["bom"]:
                _sub = _line.sub_item_code
                _sub_qty = Decimal(_line.qty or 0) * _fin
                if _sub_qty <= 0:
                    continue
                _pr = _mc_prices.get(_sub)
                if not _pr:
                    _no_price_parts.append(_sub)
                    continue
                _mat_cost += _sub_qty * _pr
            _mat_cost = _mat_cost.quantize(Decimal("0.01"))
            _backfill_rows.append(
                (
                    f"{period.month:%Y-%m}",
                    WORKSHOP_LABELS.get(_t["ws"], _t["ws"]),
                    _t["order_no"],
                    _t["product"],
                    _fin,
                    _mat_cost,
                    Decimal("0.00"),
                    _mat_cost,
                    "、".join(_no_price_parts) or "—",
                )
            )
        if _backfill_rows:
            _backfill_rows.append(
                (
                    "",
                    "合计",
                    f"{len(_backfill_rows)} 张工单",
                    "",
                    sum((r[4] for r in _backfill_rows), Decimal("0")),
                    sum((r[5] for r in _backfill_rows), Decimal("0")),
                    sum((r[6] for r in _backfill_rows), Decimal("0")),
                    sum((r[7] for r in _backfill_rows), Decimal("0")),
                    "",
                )
            )
        # Sheet2 BOM反推补投对比：按车间分组 + 车间小计（保留总合计行）
        _by_ws2 = _od_ws()
        for _row in _backfill_rows[:-1]:
            _code = _ws_label2code.get(_row[1], _row[1])
            _by_ws2.setdefault(_code, []).append(_row)
        _bf_rows_sorted = []
        for _code in sorted(_by_ws2, key=_ws_key):
            _ws_rows2 = _by_ws2[_code]
            _bf_rows_sorted.extend(_ws_rows2)
            _bf_rows_sorted.append(
                (
                    "",
                    WORKSHOP_LABELS.get(_code, _code) + " 小计",
                    f"{len(_ws_rows2)} 张工单",
                    "",
                    sum((Decimal(r[4] or 0) for r in _ws_rows2), Decimal("0")),
                    sum((Decimal(r[5] or 0) for r in _ws_rows2), Decimal("0")),
                    sum((Decimal(r[6] or 0) for r in _ws_rows2), Decimal("0")),
                    sum((Decimal(r[7] or 0) for r in _ws_rows2), Decimal("0")),
                    "",
                )
            )
        _bf_rows_sorted.append(_backfill_rows[-1])  # 总合计
        _backfill_rows = _bf_rows_sorted

        sheets.append(
            (
                "BOM反推补投对比",
                [
                    "期间", "车间", "工单号", "产品名称", "报工数量",
                    "BOM反推材料成本(核算)", "投料记录金额", "差异(反推-投料)", "无单价子项",
                ],
                _backfill_rows,
            )
        )

    elif kind == "deemed_finished":
        # ===== 视同完工清单：工单明细确认已完工 + 完工率达标 → 在产品清零，成本并入完工 =====
        # 与核算逻辑一致：浆粕不适用（机织已纳入视同完工）；判定 = ①明细实际完工时间非空或完成状态为已完工 ②完工率≥70%
        from settlement.services.workorder_costing import (
            MIN_FINISH_RATE_PCT,
            FINISHED_WORKORDER_STATUSES,
            SMALL_ORDER_WORKSHOPS,
        )
        from collections import defaultdict as _dd_df

        _issue_qty_df = _dd_df(Decimal)
        for _mi in period.material_issues.exclude(work_order_no=""):
            _issue_qty_df[_mi.work_order_no] += Decimal(_mi.actual_issue_quantity or 0)
        _roving_df = set(
            period.bom_lines.filter(sub_item_code__startswith="BC-")
            .values_list("work_order_no", flat=True)
        )
        _roving_df |= set(
            period.material_issues.filter(material_code__startswith="BC-")
            .exclude(work_order_no="")
            .values_list("work_order_no", flat=True)
        )
        _detail_df = {d.order_no: d for d in period.work_order_details.all()}
        # 报工记录最大报工时间兜底（与核算 v1.5.25 视同完工口径一致）
        from django.db.models import Max as _MaxDF
        _report_time_df = {
            row["work_order_no"]: row["max_time"]
            for row in period.work_reports.exclude(work_order_no="")
            .values("work_order_no")
            .annotate(max_time=_MaxDF("report_time"))
        }

        _df_rows = []
        for _r in period.work_order_cost_results.select_related("work_order").filter(
            work_order__production_dept_code__in=SMALL_ORDER_WORKSHOPS
        ).order_by("work_order__production_dept_code", "work_order__order_no"):
            _ws = _r.work_order.production_dept_code
            _plan = Decimal(_r.plan_quantity or 0)
            _fin = Decimal(_r.finished_quantity or 0)
            if _plan <= 0 or _fin <= 0:
                continue
            _ono = _r.work_order.order_no
            # ① 工单明细确认已完工（实际完工时间 或 完成状态为已完工）
            _d = _detail_df.get(_ono)
            _detail_finished = (
                _report_time_df.get(_ono) is not None
                or (
                    _d is not None
                    and (
                        bool(_d.actual_finish)
                        or (_d.completion_status in FINISHED_WORKORDER_STATUSES)
                    )
                )
            )
            if not _detail_finished:
                continue
            # 原在产品数量（处理前口径，与核算 _wip_qty 一致）
            if _ws == "tefang" and not (_ono.upper().startswith("BC-") or _ono in _roving_df):
                _wip_before = max(_issue_qty_df.get(_ono, Decimal("0")) - _fin, Decimal("0"))
            else:
                _wip_before = max(_plan - _fin, Decimal("0"))
            if _wip_before <= 0:
                continue
            # ② 完工率达标
            if (_fin / _plan * Decimal("100")) < MIN_FINISH_RATE_PCT:
                continue
            _df_rows.append(
                (
                    f"{period.month:%Y-%m}",
                    WORKSHOP_LABELS.get(_ws, _ws),
                    _ono,
                    _r.product_name,
                    _plan,
                    _fin,
                    (_fin / _plan * Decimal("100")).quantize(Decimal("0.1")),
                    _wip_before,
                    Decimal(_r.wip_quantity or 0),
                    (
                        (_d.actual_finish if _d and _d.actual_finish else "")
                        or (
                            _report_time_df[_ono].strftime("%Y-%m-%d %H:%M")
                            if _report_time_df.get(_ono) else ""
                        )
                        or (_d.completion_status if _d and _d.completion_status else "")
                    ),
                    _r.direct_material,
                    _r.finished_cost,
                    _r.finished_unit_cost,
                    "工单明细确认已完工，尾料按损耗并入完工成本，在产品清零不再结转",
                )
            )
        sheets.append(
            (
                "视同完工清单",
                [
                    "期间", "车间", "工单号", "产品名称",
                    "计划数量", "完工数量", "完工率%",
                    "原在产品数量", "处理后在产品数量",
                    "实际完工时间/状态（明细→报工时间兜底）",
                    "直接材料", "完工成本", "完工单位成本", "处理说明",
                ],
                _df_rows,
            )
        )

    elif kind == "reconciliation":
        sheets.append(
            (
                "勾稽结果",
                ["检查编码", "检查项目", "左侧金额", "右侧金额", "差额", "状态", "分项明细"],
                (
                    (
                        item.check_code,
                        item.check_name,
                        item.left_amount,
                        item.right_amount,
                        item.difference,
                        item.get_status_display(),
                        "；".join(f"{k}={v}" for k, v in item.details.items()),
                    )
                    for item in period.reconciliations.all()
                ),
            )
        )
        sheets.append(
            (
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
        )

    elif kind == "workshop_cost_summary":
        # 「车间费用汇总表（按车间）」的导出：与页面一致（累计口径）
        # 页面上的「期间 / 车间」筛选同样生效（用户 2026-09-27）
        _groups = _workshop_cost_groups(period)
        _sum_period = str(request.GET.get("sum_period", "")).strip()
        _sum_workshop = str(request.GET.get("sum_workshop", "")).strip()
        if _sum_period:
            _groups = [
                group for group in _groups
                if group["period_label"] == _sum_period
            ]
        _rows, _footer = _workshop_cost_summary(_groups, workshop=_sum_workshop)
        sheets.append(
            (
                "车间费用汇总",
                ["顺序", "期间", "车间"]
                + [label for _field, label in WORKSHOP_COST_COLUMNS]
                + ["合计"],
                [
                    (row["sequence"], row["period_range"], row["label"])
                    + tuple(
                        Decimal(cell["text"].replace(",", "") or 0)
                        for cell in row["cells"]
                    )
                    + (Decimal(row["total_text"].replace(",", "") or 0),)
                    for row in _rows
                ]
                + (
                    [
                        ("", _footer["period_range"], "总合计")
                        + tuple(
                            Decimal(cell["text"].replace(",", "") or 0)
                            for cell in _footer["cells"]
                        )
                        + (Decimal(_footer["total_text"].replace(",", "") or 0),)
                    ]
                    if _footer
                    else []
                ),
            )
        )
        return sheets

    elif kind == "workshop_cost":
        # 车间费用：格式/字段按用户上传文件原样，数据=本期全部车间费用；无上传时回退数据库明细
        # 补充（2026-08-27 用户要求）：车间名转中文、合计列填每行合计、末尾追加合计行
        header, raw_rows, _ = _uploaded_ledger_rows(
            period, ImportBatch.Kind.WORKSHOP_COST, period.workshop_cost_pools.all()
        )
        if header:
            rows = [list(r) for r in raw_rows]
            ws_idx = next(
                (i for i, h in enumerate(header) if normalize_header(h) == "车间"),
                None,
            )
            total_idx = next(
                (i for i, h in enumerate(header) if normalize_header(h) == "合计"),
                None,
            )
            skip_idx = {ws_idx, total_idx}
            num_idx = [
                i for i, h in enumerate(header)
                if i not in skip_idx
                and normalize_header(h) not in ("备注", "说明", "序号")
            ]
            if ws_idx is not None:
                for r in rows:
                    if r[ws_idx] not in (None, ""):
                        r[ws_idx] = WORKSHOP_LABELS.get(str(r[ws_idx]), r[ws_idx])
            if total_idx is not None:
                for r in rows:
                    nums = []
                    for i in num_idx:
                        try:
                            nums.append(Decimal(r[i] or 0))
                        except Exception:
                            pass
                    if nums:
                        r[total_idx] = sum(nums)
            if total_idx is not None:
                total_row = [""] * len(header)
                total_row[ws_idx if ws_idx is not None else 0] = "合计"
                for i in num_idx:
                    vals = []
                    for r in rows:
                        try:
                            vals.append(Decimal(r[i] or 0))
                        except Exception:
                            pass
                    if vals:
                        total_row[i] = sum(vals)
                vals = [
                    Decimal(r[total_idx] or 0)
                    for r in rows if r[total_idx] not in (None, "")
                ]
                if vals:
                    total_row[total_idx] = sum(vals)
                rows.append(tuple(total_row))
            sheets.append(("车间费用", header, rows))
            return sheets
        pools = list(period.workshop_cost_pools.all().order_by("workshop"))
        _rows = [
            (
                period.month.strftime("%Y-%m"),
                WORKSHOP_LABELS.get(p.workshop, p.workshop),
                p.wages,
                p.depreciation,
                p.water_power,
                p.steam,
                p.natural_gas,
                p.dye_auxiliary,
                p.consumables,
                p.processing_fee,
                (p.wages or 0) + (p.depreciation or 0) + (p.water_power or 0)
                + (p.steam or 0) + (p.natural_gas or 0) + (p.dye_auxiliary or 0)
                + (p.consumables or 0) + (p.processing_fee or 0),
                p.remark or "",
            )
            for p in pools
        ]
        _rows.append(
            ("合计",) + tuple(sum(r[i] for r in _rows) for i in range(2, 11)) + ("",)
        )
        sheets.append(
            (
                "车间费用",
                ["期间", "车间", "人工费", "折旧费", "水电费", "蒸汽费", "燃气费", "染料助剂费",
                 "机物料费", "加工费", "合计", "备注"],
                _rows,
            )
        )

    elif kind == "workshop_settlement":
        sheets.append(
            (
                "车间结转",
                ["结转顺序", "车间/成本中心", "状态", "工单成本", "完工入库成本",
                 "循环次数", "是否收敛", "结转时间", "结转人"],
                (
                    (
                        item.sequence,
                        WORKSHOP_LABELS.get(item.workshop, item.workshop),
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
        )

    elif kind == "opening_inventory":
        # 优先原样输出用户上传的期初台账（保留仓库/批次明细，行数与上传一致）
        batch = (
            period.import_batches.filter(
                kind=ImportBatch.Kind.OPENING_INVENTORY
            )
            .exclude(file="")
            .order_by("-created_at")
            .first()
        )
        header = []
        rows = []
        if batch and Path(batch.file.path).is_file():
            header, rows = read_uploaded_rows(Path(batch.file.path))
        if not header:
            # 无上传文件时回退到数据库
            db_items = period.opening_inventory_balances.all()
            header = ["期初日期", "物料号", "物料名称", "分类名称", "仓库名称",
                      "批次号", "结存数量", "单位名称", "单价", "金额", "来源"]
            rows = [
                (
                    item.opening_date,
                    item.material_code,
                    item.material_name,
                    item.category_name,
                    item.warehouse,
                    item.batch_no,
                    item.quantity,
                    item.uom,
                    item.unit_price,
                    item.amount,
                    "系统自动承接" if item.is_system_generated else "人工上传/覆盖",
                )
                for item in db_items
            ]
        sheets.append(("期初存货", header, rows))

    elif kind == "ending_inventory":
        rows = period.ending_inventory_balances.all()
        sheets.append(
            (
                "期末存货",
                ["期末日期", "物料号", "物料名称", "分类名称", "仓库名称", "批次号",
                 "结存数量", "单位名称", "单价", "金额", "来源"],
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
                        "系统自动计算" if item.is_system_generated else "人工上传/覆盖",
                    )
                    for item in rows
                ),
            )
        )

    elif kind == "bom":
        # 用户 2026-09-30：导出**与上传的 BOM 文件列一致**（列名、列序逐列照上传文件），
        # 系统算出来的列（用量 / 用量(转100%) / 转换后子项用量:分子 / 单价 / 标准成品单价）
        # 追加在最右边；行序＝上传顺序（同一工单的行挨在一起）。
        # 没上传过 BOM 文件的期间（数据是别的途径进来的）回退系统口径（同一套列定义，见 bom_columns）。
        # 用户 2026-09-23：子项用量合计超过 100% 的工单，标准成品单价按用量折算到 100% 后再算；
        # 导出里带一列「用量(转100%)」便于核对。
        bom_lines = list(period.bom_lines.all().order_by("id"))
        _qty_total = bom_qty_total_by_order(period)
        _share = bom_share_resolver(period)
        _columns = bom_columns(period, ledger=True)
        _prev_order_no = None
        _group_no = 0
        _rows = []
        for _i, _line in enumerate(bom_lines, start=1):
            if _line.work_order_no != _prev_order_no:
                _group_no += 1
                _prev_order_no = _line.work_order_no
            _row = bom_row_from_line(_line, _qty_total, _share)
            _rows.append(
                [
                    _bom_cell_for_excel(
                        bom_cell_value(_row, _col, index=_i, group_no=_group_no)
                    )
                    for _col in _columns
                ]
            )
        sheets.append(("BOM表", [col["name"] for col in _columns], _rows))

    elif kind == "pulp_opening":
        def _is_pulp(code, name):
            return ("浆粕" in str(name or "")) or ("微纤" in str(name or "")) or str(code or "").upper().startswith("JP-")
        pulp_items = [it for it in period.opening_inventory_balances.all()
                      if _is_pulp(it.material_code, it.material_name)]
        sheets.append(("浆粕期初成品",
            ["期初日期", "物料号", "物料名称", "分类名称", "仓库名称", "批次号",
             "结存数量", "单位名称", "单价", "金额", "来源"],
            ((it.opening_date, it.material_code, it.material_name, it.category_name,
              it.warehouse, it.batch_no, it.quantity, it.uom, it.unit_price, it.amount,
              "系统自动承接" if it.is_system_generated else "人工上传/覆盖")
             for it in pulp_items)))

    elif kind == "pulp_flow":
        def _is_pulp2(code, name):
            return ("浆粕" in str(name or "")) or ("微纤" in str(name or "")) or str(code or "").upper().startswith("JP-")
        flow_txns = [t for t in period.inventory_transactions.all()
                     if (t.business_type or "").strip() in ("生产入库", "销售出库")
                     and _is_pulp2(t.material_code, t.material_name)]
        def _fill_txn(t):
            _tu = t.unit_price
            _ta = t.amount
            if _tu is None:
                _cu = _pulp_product_unit(t.material_code)
                if _cu is not None:
                    _tu = _cu
                    _ta = (Decimal(t.quantity or 0) * Decimal(_cu)).quantize(Decimal("0.01"))
            return (t.txn_date, t.material_code, t.material_name, t.category_name, t.specification,
                    t.warehouse, t.batch_no, t.quantity, t.uom, _tu, _ta,
                    t.business_type, t.workshop or "", t.direction, t.remark or "")
        sheets.append(("浆粕成品流水",
            ["创建时间", "物料编号", "物料名称", "分类名称", "物料规格",
             "仓库名称", "批次号", "操作数量", "单位名称", "单价", "金额",
             "业务类型", "车间", "方向", "备注"],
            (_fill_txn(t) for t in flow_txns)))

    elif kind == "pulp_ending":
        def _is_pulp3(code, name):
            return ("浆粕" in str(name or "")) or ("微纤" in str(name or "")) or str(code or "").upper().startswith("JP-")
        pulp_end = [it for it in period.ending_inventory_balances.all()
                    if _is_pulp3(it.material_code, it.material_name)]
        sheets.append(("浆粕期末成品",
            ["期末日期", "物料号", "物料名称", "分类名称", "仓库名称", "批次号",
             "结存数量", "单位名称", "单价", "金额", "来源"],
            ((it.ending_date, it.material_code, it.material_name, it.category_name,
              it.warehouse, it.batch_no, it.quantity, it.uom, it.unit_price, it.amount,
              "系统自动计算" if it.is_system_generated else "人工上传/覆盖")
             for it in pulp_end)))

    elif kind == "pulp_material":
        mats = period.pulp_materials.all().order_by("id")
        sheets.append(("浆粕原料流水",
            ["日期", "单据号", "物料编码", "物料名称", "数量", "单价", "金额", "备注"],
            ((m.txn_date, m.document_no, m.material_code, m.material_name,
              m.quantity, _pulp_unit_price(m), _pulp_amount(m, _pulp_unit_price(m)), m.remark)
             for m in mats)))

    elif kind == "pulp_ledger":
        # 浆粕流水（原料+成品合并导出）：领料出库行（PulpMaterial）+ 生产入库/销售出库行（InventoryTransaction）
        rows = []
        for m in period.pulp_materials.all().order_by("id"):
            _pu = _pulp_unit_price(m)
            rows.append((m.txn_date, m.document_no, m.material_code, m.material_name,
                         m.category_name, m.specification, m.warehouse, m.batch_no,
                         m.quantity, m.uom, _pu, _pulp_amount(m, _pu),
                         m.business_type, m.workshop or "", m.remark or ""))
        for t in period.inventory_transactions.filter(
            business_type__in=["生产入库", "销售出库"]
        ).order_by("id"):
            if not (("浆粕" in str(t.material_name or "")) or ("微纤" in str(t.material_name or "")) or str(t.material_code or "").upper().startswith("JP-")):
                continue
            _tu = t.unit_price
            _ta = t.amount
            if _tu is None:
                _cu = _pulp_product_unit(t.material_code)
                if _cu is not None:
                    _tu = _cu
                    _ta = (Decimal(t.quantity or 0) * Decimal(_cu)).quantize(Decimal("0.01"))
            rows.append((t.txn_date, t.document_no, t.material_code, t.material_name,
                         t.category_name, t.specification, t.warehouse, t.batch_no,
                         t.quantity, t.uom, _tu, _ta,
                         t.business_type, t.workshop or "", t.remark or ""))
        sheets.append(("浆粕流水",
            ["创建时间", "单据号", "物料编号", "物料名称", "分类名称", "物料规格", "仓库名称", "批次号",
             "操作数量", "单位名称", "单价", "金额", "业务类型", "车间", "备注"],
            rows))
    return sheets


@login_required
def export_listing(request, period_id, kind):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    if kind not in LISTING_EXPORTS:
        raise Http404("不支持的导出类型。")
    sheets = _listing_rows(request, period, kind)
    stream = build_listing_workbook(sheets)
    title = LISTING_EXPORTS[kind]["title"]
    filename = f"{title}-{period.month:%Y%m}.xlsx"
    response = HttpResponse(
        stream.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = (
        f"attachment; filename*=UTF-8''{quote(filename)}"
    )
    return response


@login_required
def export_center(request, period_id):
    """报表导出中心：勾选需要导出的报表，打包成 ZIP 下载。"""
    period = get_object_or_404(AccountingPeriod, id=period_id)

    REPORT_GROUPS = [
        {
            "group": "原材料与采购",
            "items": [
                {"kind": "opening_inventory", "name": "期初存货（上传原样）",
                 "desc": "期初库存余额台账，按上传文件原样导出（保留仓库/批次明细，行数与上传一致）"},
                {"kind": "ending_inventory", "name": "期末存货",
                 "desc": "期末库存余额台账（物料、结存数量、单价、金额，系统自动计算）"},
                {"kind": "inventory_ledger", "name": "库存流水",
                 "desc": "本月期间全部库存出入库流水，按你上传的原始文件原样导出（单价/金额保持上传原值，不含系统回填）"},
                {"kind": "inventory_matches", "name": "库存流水匹配",
                 "desc": "库存流水的匹配情况：期间、匹配类别、匹配到的合同/发票及数量、处理方式、匹配状态"},
                {"kind": "inventory_valuation", "name": "期末库存（平台计价）",
                 "desc": "各物料期初/入库/出库/期末数量金额及月末加权平均价（含期间列）"},
                {"kind": "inventory_valuation_summary", "name": "期末库存汇总（按物料分类，平台计价）",
                 "desc": "按「期间 × 物料分类」汇总期初/入库/出库/期末的数量与金额，末行合计"},
                {"kind": "price_table", "name": "单价表（入库/出库·跨月对比）",
                 "desc": "两个工作表：入库单价、出库单价；列=序号/品类/品名编码/品名名称 + 当年1~12月，"
                         "按月填该月新入库/新出库的单价（含采购入库、完工核算入库、其他入库；上传的历史单价优先）"},
                {"kind": "purchase_contract", "name": "采购合同",
                 "desc": "采购合同台账（合同号、供应商、品名、数量、不含税单价、剩余数量）"},
                {"kind": "purchase_receipt", "name": "采购入库",
                 "desc": "本期间采购入库流水（到货台账）：数量、暂估单价/金额、累计到票与尚未到票"},
                {"kind": "purchase_invoice", "name": "采购发票匹配",
                 "desc": "采购发票及核销的入库/期初暂估明细"},
                {"kind": "purchase_accrual", "name": "月末采购暂估",
                 "desc": "截止月末仍未到票的采购入库及暂估金额"},
                {"kind": "purchase_difference", "name": "本月采购暂估差异",
                 "desc": "本月发票核销产生的暂估价差（交总账处理）"},
                {"kind": "material_cost_difference", "name": "材料成本差异明细",
                 "desc": "进「材料成本差异」科目的价差逐笔明细（含按供应商、按物料汇总三张表），"
                         "口径与「采购暂估 / 回冲凭证」一致，合计可直接与凭证对账"},
                {"kind": "purchase_reversal", "name": "发票暂估回冲明细",
                 "desc": "每张采购发票核销了哪笔入库/期初暂估：暂估金额、发票金额、价差、累计核销与剩余未到票"},
                {"kind": "pulp_opening", "name": "浆粕期初成品",
                 "desc": "浆粕车间成品期初结存（浆粕/微纤等成品）"},
                {"kind": "pulp_ledger", "name": "浆粕流水（原料+成品）",
                 "desc": "浆粕车间流水：领料出库行=原料（材料成本来源），生产入库/销售出库行=成品"}
                ,
                {"kind": "pulp_ending", "name": "浆粕期末成品",
                 "desc": "浆粕车间成品期末结存"},
            ],
        },
        {
            "group": "生产与成本",
            "items": [
                {"kind": "work_order", "name": "生产工单",
                 "desc": "生产工单主档台账（数量、状态、计划时间）"},
                {"kind": "work_order_ledger", "name": "工单台账（原样明细）",
                 "desc": "ERP 工单主档原样台账：工单号、类型、车间、计划数量、状态、计划时间、订单号、物料颜色等全字段"},
                {"kind": "order_material", "name": "工单投料",
                 "desc": "工单投料台账原样字段 + 单价 + 金额（除浆粕车间外按工单投料台账；「工单投料」页面的导出）"},
                {"kind": "material_issue_ledger", "name": "投料台账（原样明细）",
                 "desc": "ERP 投料记录原样台账：物料、投料/实际投料数量、批次、仓位、投料时间、人员、工单号等全字段"},
                {"kind": "order_completion", "name": "工单报工",
                 "desc": "工单完工汇报：报工数量、完工成本、报工入库数量与金额"},
                {"kind": "work_report_ledger", "name": "报工台账（原样明细）",
                 "desc": "ERP 报工记录原样台账：报工数量、质量状态、工时、部门、设备、工单号、颜色/批次等全字段"},
                {"kind": "work_order_detail", "name": "工单明细",
                 "desc": "各工单：创建时间、车间、计划/实际完工时间、完成状态、计划/投料/入库数量"},
                {"kind": "material_check", "name": "投料核对（异常排查）",
                 "desc": "有报工无投料的工单、被认领的供料工单、缺料←供料匹配关系；仅染色/纺纱/包覆/针织，用于与生产核对剩余料去向"},
                {"kind": "work_order_cost", "name": "工单成本",
                 "desc": "各生产工单按投料/报工核算的完工与在制成本明细，"
                         "含标准单位成本与单位成本差异（实际−标准）"},
                {"kind": "standard_cost", "name": "标准成本",
                 "desc": "三张表：工单标准成本计算表（各工单 BOM 标准成品单价 × 完工数量，"
                         "含标准材料成本/标准人工成本/标准制造费用/标准单位成本）、"
                         "标准成本构成表（按车间的产量、直接材料、人工、制造费用与材料/人工/制造费用占比）、"
                         "BOM 明细（**本期全部 BOM 行**，列＝你上传的 BOM 文件原样 + "
                         "系统列用量、用量(转100%)、转换后子项用量:分子、单价、金额、标准成品单价）；"
                         "实际成本与差异见「工单成本」导出"},
                {"kind": "standard_cost_bom", "name": "标准成本BOM明细",
                 "desc": "「标准成本」页「BOM 明细表」单独导出：**本期全部 BOM 行**"
                         "（＝你上传的 BOM 文件全部行），列＝上传文件原样同序，"
                         "系统列（用量、用量(转100%)、转换后子项用量:分子、单价、金额、标准成品单价）"
                         "追加在最右边；页面分页但导出是全部行"},
                {"kind": "ending_wip", "name": "期末在产品",
                 "desc": "各生产工单本期期末的在制余额：在产品数量、直接材料、人工费、"
                         "制造费用与在产品成本合计（只列还有在产品的工单）"},
                {"kind": "workshop_cost", "name": "车间费用",
                 "desc": "各车间本期费用（人工费/折旧费/水电费/蒸汽费/燃气费/染料助剂费/机物料费/加工费 + 合计 + 备注），"
                         "按「车间费用」页面录入的列顺序导出"},
                {"kind": "workshop_cost_summary", "name": "车间费用汇总（按车间）",
                 "desc": "各车间费用的累计汇总（期间×车间逐月累加）：8 个费用科目 + 合计，末行总合计"},
                {"kind": "workshop_settlement", "name": "车间结转",
                 "desc": "车间结转顺序、状态与成本结果"},
                {"kind": "deemed_finished", "name": "视同完工清单（小工单）",
                 "desc": "工单明细确认已完工且完工率≥70% 视同完工：在产品清零、尾料按损耗并入完工成本，不再结转"},
                {"kind": "material_issue_voucher", "name": "生产领料（投料）凭证",
                 "desc": "摘要按车间（{月}月{车间}领用{物料大类}）的凭证分录：借 生产成本—直接材料—XX车间／"
                         "贷 原材料、库存商品；含领料明细与按车间×物料大类汇总。"
                         "口径＝库存出库流水（投料出库 ＋ 无工单领料 − 投料回撤）"},
                {"kind": "finished_receipt_voucher", "name": "完工入库结转凭证",
                 "desc": "摘要按车间、科目按物料分类：借方 库存商品—物料分类；"
                         "贷方 生产成本—直接材料/直接人工/制造费用（全车间汇总，含浆粕车间）；"
                         "金额基准＝工单完工成本（报工数量 × 完工单位成本）；"
                         "含凭证分录、逐单完工入库明细、按车间×物料分类、按车间×成本要素四张表"},
                {"kind": "bom", "name": "BOM表",
                 "desc": "与你上传的 BOM 文件同列同序（序列/序号/工单编号/物料编号/…，行序＝上传顺序），"
                         "系统列（用量、用量(转100%)、转换后子项用量:分子、单价、标准成品单价）"
                         "追加在最右边；该期间没上传过 BOM 文件时按系统口径导出"},
            ],
        },
        {
            "group": "销售与开票",
            "items": [
                                {"kind": "sales_uninvoiced", "name": "月末已出库未开票",
                 "desc": "截止月末仍未开票的销售出库及发出商品成本"},
                {"kind": "purchase_voucher", "name": "采购暂估回冲凭证",
                 "desc": "本月到货到票直接入账 / 到货未到票暂估 / 前期暂估本月到票回冲，"
                         "含凭证分录、本月到货明细、回冲明细、应付暂估余额表与账龄表"},
                {"kind": "sales_voucher", "name": "销售结转成本凭证",
                 "desc": "按「结转类型 × 销售大类」汇总的结转成本凭证（会计科目、数量、金额），"
                         "含各结转类型小计与合计；会计条件与凭证页面当前的筛选一致"},
                {"kind": "sales_ledger", "name": "销售台账（成本·毛利）",
                 "desc": "你上传的销售台账 + 平台算出的成本单价（产品月末加权平均价）、销售成本、销售毛利与毛利率"},
                {"kind": "profit_statement", "name": "损益表",
                 "desc": "销售台账逐行 + 当月税金及附加/三费/营业外收支/所得税按终端销售收入占比分摊，含销售毛利、营业利润、利润总额、净利润与净利润率"},
            ],
        },
        {
            "group": "勾稽与综合",
            "items": [
                {"kind": "reconciliation", "name": "勾稽结果与异常清单",
                 "desc": "全部勾稽项目（含异常清单），两个工作表"},
                {"kind": "settlement_workbook", "name": "结算工作簿（完整）",
                 "desc": "含 16 个工作表的完整结算底稿：库存计价、跨期结转、月末暂估、勾稽等"},
                {"kind": "uploaded_files", "name": "上传的原始文件",
                 "desc": "本期间全部上传的原始文件（CSV/XLSX）"},
            ],
        },
    ]
    if request.method == "POST":
        selected = request.POST.getlist("reports")
        if not selected:
            messages.error(request, "请至少勾选一张报表。")
            return redirect("export_center", period_id=period.id)

        stream = BytesIO()
        with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as zf:
            for kind in selected:
                if kind in LISTING_EXPORTS:
                    sheets = _listing_rows(request, period, kind)
                    wb_stream = build_listing_workbook(sheets)
                    filename = f"{LISTING_EXPORTS[kind]['title']}-{period.month:%Y%m}.xlsx"
                    zf.writestr(filename, wb_stream.getvalue())
                elif kind == "price_table":
                    wb_stream = build_price_table_workbook(period)
                    filename = f"单价表-{period.month:%Y%m}.xlsx"
                    zf.writestr(filename, wb_stream.getvalue())
                elif kind == "settlement_workbook":
                    wb_stream = build_settlement_workbook(period)
                    filename = f"成本结算结果-{period.month:%Y%m}.xlsx"
                    zf.writestr(filename, wb_stream.getvalue())
                elif kind == "uploaded_files":
                    for batch in period.import_batches.filter(
                        file__gt=""
                    ).order_by("id"):
                        try:
                            if not Path(batch.file.path).is_file():
                                continue
                            data = batch.file.read()
                            safe_name = (
                                f"上传原始文件-{batch.get_kind_display()}-"
                                f"批次{batch.id}-{batch.original_name}"
                            )
                            zf.writestr(safe_name, data)
                        except OSError:
                            continue

        filename = f"报表导出-{period.month:%Y%m}.zip"
        response = HttpResponse(
            stream.getvalue(),
            content_type="application/zip",
        )
        response["Content-Disposition"] = (
            f"attachment; filename*=UTF-8''{quote(filename)}"
        )
        return response

    return render(
        request,
        "settlement/export_center.html",
        {
            "period": period,
            "grouped_reports": REPORT_GROUPS,
        },
    )


@login_required
def export_period(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    if not period.inventory_valuations.exists():
        messages.error(request, "请先执行结算，再导出结果。")
        return redirect("period_detail", period_id=period.id)
    return settlement_excel_response(period)

@login_required
def download_all_import_files(request, period_id):
    """把该期间上传过的原始文件打包成一个 ZIP（按资料类型分目录，附「上传文件清单.csv」）。

    用户 2026-09-30：报表导出页要能直接拿到"我上传的那些文件"，
    原来只能到「资料导入 → 全部批次」一行行下载。
    """
    import csv as _csv
    import zipfile
    from io import BytesIO, StringIO

    period = get_object_or_404(AccountingPeriod, id=period_id)
    batches = list(period.import_batches.order_by("kind", "id"))

    buffer = BytesIO()
    manifest = StringIO()
    writer = _csv.writer(manifest)
    writer.writerow(["资料类型", "原始文件名", "导入行数", "未导入行数",
                     "上传时间", "包内路径", "备注"])
    written = set()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for batch in batches:
            kind_label = batch.get_kind_display()
            inner = ""
            note = ""
            src = ""
            if batch.file:
                try:
                    src = batch.file.path
                except Exception:  # noqa: BLE001
                    src = ""
            if src and Path(src).is_file():
                pass
            elif batch.file_content:
                # 磁盘被清空 → 用数据库副本写进包里
                name = Path(batch.original_name or f"batch-{batch.id}").name
                inner = f"{kind_label}/{name}"
                index = 2
                while inner in written:
                    inner = f"{kind_label}/{Path(name).stem}_{index}{Path(name).suffix}"
                    index += 1
                written.add(inner)
                zf.writestr(inner, bytes(batch.file_content))
                writer.writerow([kind_label, batch.original_name, batch.row_count,
                                 batch.error_count,
                                 batch.created_at.strftime("%Y-%m-%d %H:%M"),
                                 inner, "来自数据库副本"])
                continue
            if src and Path(src).is_file():
                name = Path(src).name
                inner = f"{kind_label}/{name}"
                index = 2
                while inner in written:
                    inner = f"{kind_label}/{Path(name).stem}_{index}{Path(name).suffix}"
                    index += 1
                written.add(inner)
                zf.write(src, inner)
            elif batch.file_content:
                note = "来自数据库副本"
            elif batch.file:
                note = "原始文件已丢失（数据库也没有副本）"
            else:
                note = "页面录入，无原始文件"
            writer.writerow([kind_label, batch.original_name, batch.row_count,
                             batch.error_count, batch.created_at.strftime("%Y-%m-%d %H:%M"),
                             inner, note])
        zf.writestr("上传文件清单.csv", manifest.getvalue().encode("utf-8-sig"))

    response = HttpResponse(buffer.getvalue(), content_type="application/zip")
    response["Content-Disposition"] = (
        "attachment; filename*=UTF-8''"
        + quote(f"上传原始文件-{period.month:%Y%m}.zip")
    )
    return response

# ---------------------------------------------------------------- 客户订单执行
# 用户 2026-10-14：做一个"中石化仪征化纤订单执行情况"页面（汇总给领导看、明细给业务查）。
#   数据来源：客户订单(CustomerOrder，可导入/录入) + 销售台账(发货) + 销售发票(开票) + 客户收款。
DEFAULT_ORDER_CUSTOMER = "中国石化仪征化纤"
BUILD_STAMP = "2026-10-14 17:40"  # 部署标记（用于确认线上版本）
ZERO = Decimal("0")


def _order_execution_data(request, period):
    """组装「客户订单执行」页面的全部数据。"""
    from datetime import date, timedelta

    customer = (request.GET.get("customer") or DEFAULT_ORDER_CUSTOMER).strip()
    as_of_raw = (request.GET.get("as_of") or "").strip()
    if as_of_raw:
        try:
            y, m, d = [int(x) for x in as_of_raw.split("-")]
            as_of = date(y, m, d)
        except Exception:
            as_of = date(period.month.year, period.month.month, 1)
            as_of = (as_of.replace(day=1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)
    else:
        nxt = (period.month.replace(day=28) + timedelta(days=4)).replace(day=1)
        as_of = nxt - timedelta(days=1)

    cents = lambda v: (Decimal(v or 0) / Decimal("10000")).quantize(Decimal("0.01"))

    orders = list(CustomerOrder.objects.filter(customer_name__icontains=customer)
                  .order_by("order_no"))
    shipments = list(SalesLedgerLine.objects.filter(
        customer_name__icontains=customer, out_month__lte=as_of))
    invoices = list(SalesInvoiceLine.objects.filter(
        customer_name__icontains=customer, invoice_date__lte=as_of))
    receipts = list(CustomerReceipt.objects.filter(
        customer_name__icontains=customer, receipt_date__lte=as_of))

    order_qty = sum((o.quantity or ZERO) for o in orders)
    order_amt = sum((o.amount or ZERO) for o in orders)
    ship_qty = sum((s.quantity or ZERO) for s in shipments)
    ship_amt = sum((s.terminal_amount or ZERO) for s in shipments)
    lb_amt = sum((s.lb_amount or ZERO) for s in shipments)
    invoiced_amt = sum((i.amount_excl_tax or ZERO) for i in invoices)
    # 开票兜底：销售发票里没有该客户时，用台账里"已开票"行的终端金额
    ledger_invoiced = sum(
        (s.terminal_amount or ZERO) for s in shipments
        if getattr(s, "invoice_status", "") == "invoiced"
    )
    if ledger_invoiced > invoiced_amt:
        invoiced_amt = ledger_invoiced
    receipt_amt = sum((r.amount or ZERO) for r in receipts)

    def cat_of(text):
        text = text or ""
        for key, label in (("阻燃", "阻燃"), ("全涤", "全涤"), ("芳纶", "阻燃"),
                           ("纱线", "纱线"), ("面料", "面料")):
            if key in text:
                return label
        return "其他"

    # 按产品大类
    by_cat = {}
    for o in orders:
        key = (o.product_category or cat_of(o.product_name)).strip() or "其他"
        d = by_cat.setdefault(key, {"order_qty": ZERO, "ship_qty": ZERO})
        d["order_qty"] += o.quantity or ZERO
    for s in shipments:
        key = cat_of(s.material_name) if not s.order_no else None
        if key is None:
            key = next((c for c in by_cat if c in ((s.material_name or "") + (s.sales_category or ""))), None)
        key = key or cat_of(s.material_name)
        d = by_cat.setdefault(key, {"order_qty": ZERO, "ship_qty": ZERO})
        d["ship_qty"] += s.quantity or ZERO
    cat_rows = []
    for key, d in sorted(by_cat.items(), key=lambda kv: -kv[1]["order_qty"]):
        oq, sq = d["order_qty"], d["ship_qty"]
        if oq <= 0:
            # 只显示"有订单"的产品大类；没有订单的发货统一放在下面的提示行里
            continue
        cat_rows.append({
            "category": key, "order_qty": oq, "ship_qty": sq,
            "left_qty": max(ZERO, oq - sq),
            "rate": (sq / oq * 100) if oq else ZERO,
        })

    # 按订单
    ship_by_order, inv_by_order, rec_by_order = {}, {}, {}
    for s in shipments:
        k = (s.order_no or "").strip()
        d = ship_by_order.setdefault(k, [ZERO, ZERO])
        d[0] += s.quantity or ZERO
        d[1] += s.terminal_amount or ZERO
    for i in invoices:
        k = (i.sales_order_no or "").strip()
        inv_by_order[k] = inv_by_order.get(k, ZERO) + (i.amount_excl_tax or ZERO)
    for r in receipts:
        k = (r.order_no or "").strip()
        rec_by_order[k] = rec_by_order.get(k, ZERO) + (r.amount or ZERO)

    order_rows = []
    for o in orders:
        k = (o.order_no or "").strip()
        sq, sa = ship_by_order.get(k, [ZERO, ZERO])
        inv = inv_by_order.get(k, ZERO)
        rec = rec_by_order.get(k, ZERO)
        left = max(ZERO, (o.quantity or ZERO) - sq)
        rate = (sq / o.quantity * 100) if o.quantity else ZERO
        if left <= 0:
            status, tone = "已发完", "ok"
        elif o.due_date and as_of > o.due_date:
            status, tone = "超期未发完", "bad"
        elif rate < 40:
            status, tone = "偏慢", "warn"
        else:
            status, tone = "执行中", "ok"
        if left <= 0 and (o.amount or ZERO) > rec:
            status, tone = "已发完·待收款", "warn"
        order_rows.append({
            "order": o, "ship_qty": sq, "left_qty": left, "rate": rate,
            "invoiced": inv, "received": rec,
            "unreceived": (o.amount or ZERO) - rec, "status": status, "tone": tone,
        })

    no_order_ship = ship_by_order.get("", [ZERO, ZERO])
    summary = {
        "order_qty": order_qty, "ship_qty": ship_qty,
        "left_qty": max(ZERO, order_qty - ship_qty),
        "rate": (ship_qty / order_qty * 100) if order_qty else ZERO,
        "order_amt": order_amt, "ship_amt": ship_amt, "lb_amt": lb_amt,
        "invoiced_amt": invoiced_amt, "receipt_amt": receipt_amt,
        "unreceived": order_amt - receipt_amt,
        "uninvoiced": ship_amt - invoiced_amt,
        "invoice_rate": (invoiced_amt / ship_amt * 100) if ship_amt else ZERO,
        "receipt_rate": (receipt_amt / order_amt * 100) if order_amt else ZERO,
        "no_order_qty": no_order_ship[0], "no_order_amt": no_order_ship[1],
        "wan": lambda v: (Decimal(v or 0) / Decimal("10000")).quantize(Decimal("0.01")),
    }

    customers = {}
    for name in SalesLedgerLine.objects.values_list("customer_name", flat=True):
        if name:
            customers[name] = customers.get(name, 0) + 1
    for name in CustomerOrder.objects.values_list("customer_name", flat=True):
        if name:
            customers[name] = customers.get(name, 0) + 1
    top_customers = [n for n, _c in sorted(customers.items(), key=lambda kv: -kv[1])[:60]]

    return {
        "period": period, "customer": customer, "as_of": as_of,
        "summary": summary, "cat_rows": cat_rows, "order_rows": order_rows,
        "shipments": shipments[:2000], "invoices": invoices[:2000], "receipts": receipts[:2000],
        "top_customers": top_customers,
        "has_orders": bool(orders),
    }


@login_required
def order_execution(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    return render(request, "settlement/order_execution.html", _order_execution_data(request, period))


@login_required
def order_execution_export(request, period_id):
    """导出：汇总 / 订单明细 / 发货明细 / 开票收款 四个 sheet。"""
    from urllib.parse import quote

    from openpyxl import Workbook
    from openpyxl.styles import Font

    period = get_object_or_404(AccountingPeriod, id=period_id)
    ctx = _order_execution_data(request, period)
    s, wan = ctx["summary"], ctx["summary"]["wan"]
    wb = Workbook()

    ws = wb.active
    ws.title = "汇总"
    bold = Font(bold=True)
    rows = [
        [f"{period.month:%Y年%m月} {ctx['customer']} 订单执行情况汇总"],
        [f"数据截至 {ctx['as_of']:%Y-%m-%d}；数量单位：万米，金额单位：万元"],
        [],
        ["一、订单与发货进度", "数量（万米）", "备注"],
        ["订单总量", wan(s["order_qty"])],
        ["已发货", wan(s["ship_qty"])],
        ["其中：用兰邦库存发货", wan(s["lb_amt"])],
        ["待发货", wan(s["left_qty"])],
        ["执行率", f"{s['rate']:.2f}%"],
        [],
        ["按产品大类", "订单量", "已发货", "待发货", "执行率"],
    ]
    for c in ctx["cat_rows"]:
        rows.append([c["category"], wan(c["order_qty"]), wan(c["ship_qty"]),
                     wan(c["left_qty"]), f"{c['rate']:.2f}%"])
    rows += [
        [],
        ["二、订单回款情况", "金额（万元）", "备注"],
        ["订单应收款", wan(s["order_amt"])],
        ["已发货金额", wan(s["ship_amt"])],
        ["实际开票金额", wan(s["invoiced_amt"])],
        ["已收款金额", wan(s["receipt_amt"])],
        ["未收款", wan(s["unreceived"])],
        ["已发货未开票", wan(s["uninvoiced"])],
        ["开票率", f"{s['invoice_rate']:.2f}%"],
        ["收款率", f"{s['receipt_rate']:.2f}%"],
    ]
    for r in rows:
        ws.append(r)
    for cell in ws["A"]:
        if cell.value and str(cell.value).startswith(("一、", "二、", "按产品")):
            cell.font = bold
    ws["A1"].font = Font(bold=True, size=13)

    ws2 = wb.create_sheet("订单明细")
    ws2.append(["订单号", "合同号", "签订年", "品名", "规格", "产品大类", "订单数量",
                "已发货", "待发货", "执行率", "订单金额", "已开票", "已收款", "未收款", "状态"])
    for r in ctx["order_rows"]:
        o = r["order"]
        ws2.append([o.order_no, o.contract_no, o.sign_year, o.product_name, o.specification,
                    o.product_category, o.quantity, r["ship_qty"], r["left_qty"],
                    f"{r['rate']:.2f}%", o.amount, r["invoiced"], r["received"],
                    r["unreceived"], r["status"]])

    ws3 = wb.create_sheet("发货明细")
    ws3.append(["订单号", "出库月", "客户", "品名", "规格", "数量", "单位",
                "终端金额", "兰邦金额", "开票状态"])
    for x in ctx["shipments"]:
        ws3.append([x.order_no, x.out_month, x.customer_name, x.material_name, x.specification,
                    x.quantity, x.uom, x.terminal_amount, x.lb_amount,
                    x.get_invoice_status_display()])

    ws4 = wb.create_sheet("开票收款")
    ws4.append(["类型", "单号", "日期", "客户", "订单号", "金额"])
    for i in ctx["invoices"]:
        ws4.append(["发票", i.invoice_no, i.invoice_date, i.customer_name,
                    i.sales_order_no, i.amount_excl_tax])
    for r in ctx["receipts"]:
        ws4.append(["收款", "", r.receipt_date, r.customer_name, r.order_no, r.amount])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    name = f"{period.month:%Y%m}_{ctx['customer']}_订单执行.xlsx"
    resp = HttpResponse(
        buf.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    resp["Content-Disposition"] = "attachment; filename*=UTF-8''" + quote(name)
    return resp


@login_required
def fix_opening_carry(request, period_id):
    """运维用：清空该期间全部期初，让库存结转按规则重建一份（不重复）。"""
    from django.http import JsonResponse

    from settlement.models import AccountingPeriod, OpeningInventoryBalance

    period = get_object_or_404(AccountingPeriod, id=period_id)
    deleted, _d = OpeningInventoryBalance.objects.filter(period=period).delete()
    return JsonResponse({"period": str(period.month)[:7], "deleted": deleted})



@login_required
def unify_material_code(request, period_id):
    """运维用：把不规范编码（如 14.35）统一改为规范编码（14.350）。"""
    from django.http import JsonResponse

    from settlement.models import (
        AccountingPeriod, InventoryTransaction, InventoryValuation,
        OpeningInventoryBalance,
    )

    period = get_object_or_404(AccountingPeriod, id=period_id)
    raw_pairs = (request.GET.get("pairs") or "").strip()
    result = {}
    for item in [x for x in raw_pairs.split(",") if ":" in x]:
        bad, good = item.split(":", 1)
        n1 = InventoryTransaction.objects.filter(period=period, material_code=bad).update(material_code=good)
        n2 = InventoryValuation.objects.filter(period=period, material_code=bad).update(material_code=good)
        n3 = OpeningInventoryBalance.objects.filter(period=period, material_code=bad).update(material_code=good)
        result[f"{bad}->{good}"] = [n1, n2, n3]
    return JsonResponse({"period": str(period.month)[:7], "result": result})
