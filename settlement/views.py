from pathlib import Path
from collections import defaultdict
from decimal import Decimal

from django.conf import settings
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.db.models import Count, Q, Sum
from django.http import FileResponse, Http404, HttpResponse, HttpResponseBadRequest
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
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
    ExceptionItem,
    HistoricalPrice,
    ImportBatch,
    InventoryTransaction,
    MatchAllocation,
    MatchingMode,
    OpeningGoodsIssued,
    OpeningPurchaseAccrual,
    PurchaseContractLine,
    PurchaseInvoiceLine,
    ReconciliationResult,
    SalesInvoiceLine,
)
from settlement.services.costing import (
    can_lock,
    ensure_workshop_settlements,
    reset_workshop_settlements,
    run_settlement,
    settle_workshop as settle_one_workshop,
    update_workshop_sequences,
)
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
    summarize_purchase_accruals,
    summarize_purchase_differences,
    summarize_purchase_reversals,
)
from settlement.services.workorder_costing import (
    build_work_order_cost_summary,
)
from settlement.services.rollforward import (
    carry_forward_context,
    invalidate_downstream_periods,
)
from settlement.services.sales_uninvoiced import (
    build_goods_issued_rollforward,
    build_sales_uninvoiced_rows,
    summarize_sales_uninvoiced,
)
from urllib.parse import quote
from io import BytesIO
import zipfile


TEMPLATE_FILES = {
    ImportBatch.Kind.INVENTORY: "库存出入库流水.csv",
    ImportBatch.Kind.PURCHASE_CONTRACT: "采购合同.csv",
    ImportBatch.Kind.PURCHASE_INVOICE: "采购发票.csv",
    ImportBatch.Kind.SALES_INVOICE: "销售发票.csv",
    ImportBatch.Kind.WORKSHOP_COST: "车间费用.csv",
    ImportBatch.Kind.OPENING_WIP: "期初在产品.xlsx",
    ImportBatch.Kind.OPENING_PURCHASE_ACCRUAL: "期初采购暂估明细.csv",
    ImportBatch.Kind.OPENING_GOODS_ISSUED: "期初发出商品明细.csv",
    ImportBatch.Kind.OPENING_INVENTORY: "期初存货.csv",
    ImportBatch.Kind.ENDING_INVENTORY: "期末存货.csv",
    ImportBatch.Kind.WORK_ORDER: "生产工单.csv",
    ImportBatch.Kind.MATERIAL_ISSUE: "投料记录.csv",
    ImportBatch.Kind.WORK_REPORT: "报工记录.csv",
    ImportBatch.Kind.PULP_RECEIPT: "浆粕成品入库.xlsx",
    ImportBatch.Kind.BOM: "BOM表.xlsx",
    ImportBatch.Kind.ISSUE_PROGRESS: "投料进度表.xlsx",
    ImportBatch.Kind.HISTORICAL_PRICE: "历史单价.csv",
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
    if is_sales_issue(txn):
        return "销售出库"
    business_type = (txn.business_type or "").strip()
    if business_type == "报工":
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
        "sales_invoices": period.sales_invoices.count(),
        "orders": period.work_orders.count(),
        "valuations": period.inventory_valuations.count(),
        "ending_inventory": period.ending_inventory_balances.count(),
        "blocking": blocking,
        "failed_checks": failed_checks,
        "workshops": len(settlements),
        "settled_workshops": sum(
            1 for item in settlements if item.status == "settled"
        ),
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
    if not batch.file:
        raise Http404("该批次没有原始文件。")
    if not Path(batch.file.path).is_file():
        raise Http404("原始文件已丢失。")
    return FileResponse(
        batch.file.open("rb"),
        as_attachment=True,
        filename=batch.original_name,
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
                return redirect("period_detail", period_id=period.id)
            except Exception as exc:
                messages.error(request, str(exc))
    else:
        form = UploadForm()
    return render(
        request,
        "settlement/upload.html",
        {
            "period": period,
            "form": form,
            "template_kinds": ImportBatch.Kind.choices,
            "carry": carry_forward_context(period),
        },
    )


@login_required
def download_template(request, kind):
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
        messages.error(request, f"结算失败：{exc}")
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
        workshop = run.summary.get("workshop", "该车间")
        messages.success(
            request,
            f"{workshop}结转完成，共循环{run.iterations}次。"
            + (
                "全部车间已结转，可以进行勾稽和锁账。"
                if run.summary.get("all_settled")
                else "现在可以结转下一车间。"
            ),
        )
    except Exception as exc:
        messages.error(request, f"车间结转失败：{exc}")
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


@login_required
def purchase_matches(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    invoices = period.purchase_invoices.annotate(
        matched_total=Sum("matchallocation__matched_quantity")
    ).prefetch_related(
        "matchallocation_set__inventory_transaction__period",
        "matchallocation_set__opening_purchase_accrual",
    )
    invoices = _text_filters(
        request,
        invoices,
        {
            "invoice_no": "invoice_no",
            "contract_no": "contract_no",
            "material": "material_code",
        },
    )
    invoices = _party_filter(
        request,
        invoices,
        "supplier",
        "supplier_code",
        "supplier_name",
    )
    invoices = _date_filter(
        request, invoices, "invoice_date", "invoice_date"
    )
    invoices = _choice_filter(
        request, invoices, "match_mode", "match_mode"
    ).order_by("invoice_date", "invoice_no", "id")
    return render(
        request,
        "settlement/purchase_matches.html",
        {
            "period": period,
            "page": _page(request, invoices),
            "match_modes": MatchingMode.choices,
        },
    )


@login_required
def purchase_contracts(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    contracts = period.purchase_contracts.prefetch_related("batch")
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
    return render(
        request,
        "settlement/purchase_contracts.html",
        {
            "period": period,
            "page": _page(request, contracts),
            "match_modes": MatchingMode.choices,
        },
    )


@login_required
def purchase_accruals(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    rows = build_purchase_accrual_rows(period)
    receipt_month = str(request.GET.get("receipt_month", "")).strip()
    document_no = str(request.GET.get("document_no", "")).strip().lower()
    supplier = str(request.GET.get("supplier", "")).strip().lower()
    material = str(request.GET.get("material", "")).strip().lower()
    contract_no = str(request.GET.get("contract_no", "")).strip().lower()
    status = str(request.GET.get("status", "")).strip()
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
                f"{row.supplier_code} "
                f"{row.supplier_name}"
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
    if contract_no:
        rows = [
            row for row in rows if contract_no in row.contract_nos.lower()
        ]
    if status:
        rows = [row for row in rows if row.status == status]
    return render(
        request,
        "settlement/purchase_accruals.html",
        {
            "period": period,
            "page": _page(request, rows),
            "summary": summarize_purchase_accruals(rows, period),
            "statuses": ["全部暂估", "部分到票", "缺少暂估价格"],
        },
    )


@login_required
def purchase_differences(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    rows = build_purchase_difference_rows(period)
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
    totals, grouped_rows = summarize_purchase_differences(rows)
    return render(
        request,
        "settlement/purchase_differences.html",
        {
            "period": period,
            "page": _page(request, rows),
            "totals": totals,
            "grouped_rows": grouped_rows,
            "statuses": ["有差异", "无差异", "待确认暂估价格"],
        },
    )


@login_required
def sales_matches(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    invoices = period.sales_invoices.annotate(
        matched_total=Sum("matchallocation__matched_quantity")
    ).prefetch_related(
        "matchallocation_set__inventory_transaction__period",
        "matchallocation_set__opening_goods_issued",
    )
    invoices = _text_filters(
        request,
        invoices,
        {
            "invoice_no": "invoice_no",
            "sales_order_no": "sales_order_no",
            "material": "material_code",
        },
    )
    invoices = _party_filter(
        request,
        invoices,
        "customer",
        "customer_code",
        "customer_name",
    )
    invoices = _date_filter(
        request, invoices, "invoice_date", "invoice_date"
    )
    invoices = _choice_filter(
        request, invoices, "match_mode", "match_mode"
    ).order_by("invoice_date", "invoice_no", "id")
    return render(
        request,
        "settlement/sales_matches.html",
        {
            "period": period,
            "page": _page(request, invoices),
            "match_modes": MatchingMode.choices,
        },
    )


@login_required
def sales_uninvoiced(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    rows = build_sales_uninvoiced_rows(period)
    issue_month = str(request.GET.get("issue_month", "")).strip()
    document_no = str(request.GET.get("document_no", "")).strip().lower()
    sales_order_no = str(
        request.GET.get("sales_order_no", "")
    ).strip().lower()
    customer = str(request.GET.get("customer", "")).strip().lower()
    material = str(request.GET.get("material", "")).strip().lower()
    status = str(request.GET.get("status", "")).strip()
    cross_period = str(request.GET.get("cross_period", "")).strip()
    if issue_month:
        rows = [
            row
            for row in rows
            if row.issue_month.strftime("%Y-%m") == issue_month
        ]
    if document_no:
        rows = [
            row
            for row in rows
            if document_no in (row.document_no or "").lower()
        ]
    if sales_order_no:
        rows = [
            row
            for row in rows
            if sales_order_no in (row.sales_order_no or "").lower()
        ]
    if customer:
        rows = [
            row
            for row in rows
            if customer
            in (
                f"{row.customer_code} "
                f"{row.customer_name}"
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
    return render(
        request,
        "settlement/sales_uninvoiced.html",
        {
            "period": period,
            "page": _page(request, rows),
            "summary": summarize_sales_uninvoiced(rows, period),
            "rollforward": build_goods_issued_rollforward(period),
            "statuses": ["全部未开票", "部分开票", "缺少出库成本"],
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
def work_order_matches(request, period_id):
    """工单匹配：以生产工单为主，查看投料/报工台账归集与报工入库关联情况。

    生产工单口径下不再做"生产工单"的流水匹配；
    投料记录按工单归集直接材料，报工入库按单据号=工单编号关联计价。
    """
    period = get_object_or_404(AccountingPeriod, id=period_id)
    masters = {wo.order_no: wo for wo in period.work_orders.all()}
    issue_nos = set(
        period.material_issues.exclude(work_order_no="").values_list(
            "work_order_no", flat=True
        )
    )
    report_nos = set(
        period.work_reports.exclude(work_order_no="").values_list(
            "work_order_no", flat=True
        )
    )
    all_nos = sorted(set(masters) | issue_nos | report_nos)

    issue_agg = {
        row["work_order_no"]: row
        for row in period.material_issues.exclude(work_order_no="")
        .values("work_order_no")
        .annotate(
            issue_count=Count("id"),
            issue_qty=Sum("actual_issue_quantity"),
        )
    }
    report_agg = {
        row["work_order_no"]: row
        for row in period.work_reports.exclude(work_order_no="")
        .values("work_order_no")
        .annotate(
            report_count=Count("id"),
            report_qty=Sum("quantity"),
        )
    }
    receipt_agg = {}
    for row in (
        period.inventory_transactions.filter(business_type="报工")
        .values("document_no")
        .annotate(
            txn_count=Count("id"),
            txn_qty=Sum("quantity"),
            priced_count=Count("id", filter=Q(amount__isnull=False)),
        )
    ):
        receipt_agg[row["document_no"]] = row

    enriched = []
    for order_no in all_nos:
        wo = masters.get(order_no)
        issue = issue_agg.get(order_no)
        report = report_agg.get(order_no)
        receipt = receipt_agg.get(order_no)
        product_code, product_name = _work_order_product(period, order_no)
        report_qty = Decimal(report["report_qty"]) if report else Decimal("0")
        receipt_qty = Decimal(receipt["txn_qty"]) if receipt else Decimal("0")
        enriched.append(
            {
                "order_no": order_no,
                "workshop": wo.production_dept_code if wo else "",
                "status": wo.status if wo else "",
                "plan_quantity": Decimal(wo.plan_quantity) if wo else Decimal("0"),
                "product_code": product_code,
                "product_name": product_name,
                "issue_count": issue["issue_count"] if issue else 0,
                "issue_qty": Decimal(issue["issue_qty"]) if issue else Decimal("0"),
                "report_count": report["report_count"] if report else 0,
                "report_qty": report_qty,
                "receipt_count": receipt["txn_count"] if receipt else 0,
                "receipt_qty": receipt_qty,
                "receipt_priced": receipt["priced_count"] if receipt else 0,
                "qty_mismatch": bool(
                    report and receipt
                    and abs(report_qty - receipt_qty) > Decimal("0.001")
                ),
            }
        )

    order_no = str(request.GET.get("order_no", "")).strip()
    workshop = str(request.GET.get("workshop", "")).strip()
    product = str(request.GET.get("product", "")).strip()
    if order_no:
        enriched = [row for row in enriched if order_no in row["order_no"]]
    if workshop:
        enriched = [
            row for row in enriched if workshop in (row["workshop"] or "")
        ]
    if product:
        enriched = [
            row for row in enriched if product in (row["product_code"] or "")
        ]

    page = Paginator(enriched, 100).get_page(request.GET.get("page"))
    query = request.GET.copy()
    query.pop("page", None)
    page.querystring = query.urlencode()
    return render(
        request,
        "settlement/work_order_matches.html",
        {
            "period": period,
            "page": page,
        },
    )


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


@login_required
def inventory_matches(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
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
    page = _page(request, txns)
    for txn in page.object_list:
        txn.match_category = _inventory_match_category(txn)
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
                )[:300],
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
    elif is_sales_issue(txn):
        groups.append(
            {
                "kind": MatchAllocation.Kind.SALES_INVOICE,
                "label": "销售发票",
                "items": period.sales_invoices.filter(
                    material_code=txn.material_code
                ).exclude(match_mode=MatchingMode.IGNORED)[:300],
            }
        )
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


@login_required
def inventory_results(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    totals = period.inventory_valuations.aggregate(
        opening=Sum("opening_amount"),
        receipts=Sum("receipt_amount"),
        issues=Sum("issue_amount"),
        ending=Sum("ending_amount"),
    )
    return render(
        request,
        "settlement/inventory_results.html",
        {
            "period": period,
            "page": _page(request, period.inventory_valuations.all()),
            "totals": totals,
        },
    )


@login_required
def work_order_report(request, period_id):
    """生产工单报表：工单主档台账（生产工单口径）。"""
    period = get_object_or_404(AccountingPeriod, id=period_id)
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
    return render(
        request,
        "settlement/work_order_report.html",
        {
            "period": period,
            "page": _page(
                request,
                orders.order_by("production_dept_code", "order_no"),
            ),
        },
    )


def build_material_issue_rows(period, order_no="", workshop="", material=""):
    """按工单 × 物料汇总投料记录台账（实际投料数量、金额按月末加权价）。"""
    prices = {
        item.material_code: Decimal(item.weighted_average_price or 0)
        for item in period.inventory_valuations.all()
    }
    wo_workshop = {
        wo.order_no: wo.production_dept_code for wo in period.work_orders.all()
    }
    issues = period.material_issues.all()
    if order_no:
        issues = issues.filter(work_order_no__icontains=order_no)
    if material:
        issues = issues.filter(material_code__icontains=material)
    buckets = {}
    for issue in issues:
        if not issue.work_order_no or not issue.material_code:
            continue
        if workshop and workshop not in (
            wo_workshop.get(issue.work_order_no, "") or ""
        ):
            continue
        key = (issue.work_order_no, issue.material_code)
        bucket = buckets.setdefault(
            key,
            {
                "order_no": issue.work_order_no,
                "workshop": wo_workshop.get(issue.work_order_no, ""),
                "material_code": issue.material_code,
                "material_name": issue.material_name,
                "specification": issue.material_spec,
                "uom": issue.issue_uom,
                "issue_quantity": Decimal("0"),
                "amount": Decimal("0"),
            },
        )
        qty = Decimal(issue.actual_issue_quantity or 0)
        bucket["issue_quantity"] += qty
        bucket["amount"] += qty * prices.get(issue.material_code, Decimal("0"))
    rows = []
    for bucket in buckets.values():
        bucket["amount"] = bucket["amount"].quantize(Decimal("0.01"))
        rows.append(bucket)
    rows.sort(
        key=lambda r: (r["workshop"], r["order_no"], r["material_code"])
    )
    totals = {
        "issue_quantity": sum((r["issue_quantity"] for r in rows), Decimal("0")),
        "amount": sum((r["amount"] for r in rows), Decimal("0")),
    }
    return rows, totals


@login_required
def order_material_report(request, period_id):
    """工单投料报表：按工单 × 物料汇总投料记录台账（实际投料数量、金额）。"""
    period = get_object_or_404(AccountingPeriod, id=period_id)
    rows, totals = build_material_issue_rows(
        period,
        order_no=request.GET.get("order_no", "").strip(),
        workshop=request.GET.get("workshop", "").strip(),
        material=request.GET.get("material", "").strip(),
    )
    return render(
        request,
        "settlement/order_material_report.html",
        {
            "period": period,
            "rows": rows,
            "totals": totals,
        },
    )


@login_required
def order_completion_report(request, period_id):
    """工单报工报表：报工数量、完工成本、报工入库。"""
    period = get_object_or_404(AccountingPeriod, id=period_id)
    masters = {wo.order_no: wo for wo in period.work_orders.all()}
    report_nos = set(
        period.work_reports.exclude(work_order_no="").values_list(
            "work_order_no", flat=True
        )
    )
    all_nos = sorted(set(masters) | report_nos)

    report_agg = {
        row["work_order_no"]: row
        for row in period.work_reports.exclude(work_order_no="")
        .values("work_order_no")
        .annotate(
            report_count=Count("id"),
            report_qty=Sum("quantity"),
        )
    }
    receipt_agg = {}
    for row in (
        period.inventory_transactions.filter(business_type="报工")
        .values("document_no")
        .annotate(txn_qty=Sum("quantity"), txn_amount=Sum("amount"))
    ):
        receipt_agg[row["document_no"]] = row
    results = {
        item.work_order.order_no: item
        for item in period.work_order_cost_results.select_related(
            "work_order"
        ).all()
    }

    enriched = []
    for order_no in all_nos:
        wo = masters.get(order_no)
        result = results.get(order_no)
        report = report_agg.get(order_no)
        receipt = receipt_agg.get(order_no)
        product_code, product_name = _work_order_product(period, order_no)
        enriched.append(
            {
                "order_no": order_no,
                "workshop": wo.production_dept_code if wo else "",
                "status": wo.status if wo else "",
                "product_code": product_code,
                "product_name": product_name,
                "plan_quantity": Decimal(wo.plan_quantity) if wo else Decimal("0"),
                "report_count": report["report_count"] if report else 0,
                "report_qty": Decimal(report["report_qty"]) if report else Decimal("0"),
                "result": result,
                "receipt_qty": (
                    Decimal(receipt["txn_qty"]) if receipt else Decimal("0")
                ),
                "receipt_amount": (
                    Decimal(receipt["txn_amount"])
                    if receipt and receipt["txn_amount"] is not None
                    else Decimal("0")
                ),
            }
        )

    order_no = str(request.GET.get("order_no", "")).strip()
    workshop = str(request.GET.get("workshop", "")).strip()
    product = str(request.GET.get("product", "")).strip()
    if order_no:
        enriched = [row for row in enriched if order_no in row["order_no"]]
    if workshop:
        enriched = [
            row for row in enriched if workshop in (row["workshop"] or "")
        ]
    if product:
        enriched = [
            row for row in enriched if product in (row["product_code"] or "")
        ]

    totals = {
        "report_quantity": sum(
            (item["report_qty"] for item in enriched), Decimal("0")
        ),
        "finished_cost": sum(
            (
                item["result"].finished_cost
                for item in enriched
                if item["result"]
            ),
            Decimal("0"),
        ),
        "wip_cost": sum(
            (item["result"].wip_cost for item in enriched if item["result"]),
            Decimal("0"),
        ),
        "total_cost": sum(
            (item["result"].total_cost for item in enriched if item["result"]),
            Decimal("0"),
        ),
        "receipt_quantity": sum(
            (item["receipt_qty"] for item in enriched), Decimal("0")
        ),
        "receipt_amount": sum(
            (item["receipt_amount"] for item in enriched), Decimal("0")
        ),
    }
    return render(
        request,
        "settlement/order_completion_report.html",
        {
            "period": period,
            "page": _page(request, enriched),
            "totals": totals,
        },
    )


@login_required
def purchase_reversal_report(request, period_id):
    """发票暂估回冲明细：每张采购发票核销了哪笔入库/期初暂估。"""
    period = get_object_or_404(AccountingPeriod, id=period_id)
    rows = build_purchase_reversal_rows(
        period,
        invoice_no=request.GET.get("invoice_no", "").strip(),
        supplier=request.GET.get("supplier", "").strip(),
        material=request.GET.get("material", "").strip(),
        cross_period=request.GET.get("cross_period", "").strip(),
    )
    totals = summarize_purchase_reversals(rows)
    return render(
        request,
        "settlement/purchase_reversal_report.html",
        {
            "period": period,
            "rows": rows,
            "totals": totals,
        },
    )


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


@login_required
def bom_imports(request, period_id):
    """BOM 表导入台账：子项单价与标准成品单价由系统按月末加权价回填。"""
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
    return render(
        request,
        "settlement/bom_imports.html",
        {
            "period": period,
            "page": _page(request, lines),
        },
    )


@login_required
def production_cost_summary(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    selected_workshop = request.GET.get("workshop", "").strip()
    summary = build_work_order_cost_summary(period, selected_workshop)
    return render(
        request,
        "settlement/production_cost_summary.html",
        {
            "period": period,
            **summary,
        },
    )


@login_required
def work_order_costs(request, period_id):
    """工单成本明细（生产工单+投料+报工核算结果；每个车间独立一张表，页签切换）。

    分配率口径（与分摊逻辑一致）：
    - 制造费用（水电/蒸汽/天然气/折旧/机物料/染料助剂）分配率
      = 车间该费用总额 ÷ 车间材料成本总额（每元材料承担的费用）
    - 工资分配率 = 工单折合数量（完工+0.5×在产品）÷ 车间折合数量合计
    """
    period = get_object_or_404(AccountingPeriod, id=period_id)
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
    results = list(
        results.order_by("work_order__production_dept_code", "work_order__order_no")
    )
    openings = {
        item.order_no: item for item in period.opening_wip.all()
    }

    # 车间聚合：材料总额 / 折合数量合计（分母）
    workshop_material = defaultdict(Decimal)
    workshop_equivalent = defaultdict(Decimal)
    for item in results:
        opening = openings.get(item.work_order.order_no)
        current_material = Decimal(item.direct_material or 0) - Decimal(
            opening.material if opening else 0
        )
        workshop_material[item.work_order.production_dept_code] += current_material
        equivalent = Decimal(item.finished_quantity or 0) + Decimal(
            item.wip_quantity or 0
        ) * Decimal("0.5")
        workshop_equivalent[item.work_order.production_dept_code] += equivalent

    # 车间级制造费用分配率 = 该费用总额 ÷ 车间材料成本总额（每元材料承担的费用）
    from settlement.services.workorder_costing import MANUFACTURING_FIELDS

    workshop_fee_total = defaultdict(lambda: defaultdict(Decimal))
    workshop_material_total = defaultdict(Decimal)
    for item in results:
        ws = item.work_order.production_dept_code
        workshop_material_total[ws] += Decimal(item.direct_material or 0)
        for field in MANUFACTURING_FIELDS:
            workshop_fee_total[ws][field] += Decimal(getattr(item, field) or 0)

    # 每个工单的分配率（制造费用 = 车间级费用/材料；工资 = 折合数量占比）
    for item in results:
        ws = item.work_order.production_dept_code
        for field in MANUFACTURING_FIELDS:
            fee_total = workshop_fee_total[ws][field]
            mat_total = workshop_material_total[ws]
            setattr(
                item,
                f"rate_{field}_pct",
                ((fee_total / mat_total) * Decimal("100"))
                if mat_total
                else Decimal("0"),
            )
        equivalent = Decimal(item.finished_quantity or 0) + Decimal(
            item.wip_quantity or 0
        ) * Decimal("0.5")
        denom2 = workshop_equivalent[ws]
        item.rate_wages_pct = (
            (equivalent / denom2 * Decimal("100")) if denom2 else Decimal("0")
        )

    # 投料进度（投料进度表）与粗纱标志
    from settlement.models import BomLine, IssueProgress

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

    for item in results:
        order_no = item.work_order.order_no
        ws = item.work_order.production_dept_code
        # 累计投料：投料进度表优先，无则用实际投料记录合计
        if order_no in progress_qty_by_order:
            item.accum_issue = progress_qty_by_order[order_no]
        else:
            item.accum_issue = Decimal(
                period.material_issues.filter(work_order_no=order_no)
                .aggregate(total=Sum("actual_issue_quantity"))["total"]
                or 0
            )
        # 投料状态：进度表该工单各行状态去重
        statuses = progress_status_by_order.get(order_no)
        item.issue_status = "、".join(sorted(statuses)) if statuses else "—"
        # 累计完工 / 完工率（完工÷(完工+在产品)，与核算口径一致）
        item.accum_finished = Decimal(item.finished_quantity or 0)
        wip_qty = Decimal(item.wip_quantity or 0)
        item.finish_rate_pct = (
            (item.accum_finished / (item.accum_finished + wip_qty) * Decimal("100"))
            if (item.accum_finished + wip_qty)
            else None
        )
        # 完成状态：未开工 / 部分完成 / 已完成
        if item.accum_finished <= Decimal("0") and wip_qty <= Decimal("0"):
            item.completion_status = "未开工"
        elif wip_qty <= Decimal("0"):
            item.completion_status = "已完成"
        else:
            item.completion_status = "部分完成"
        # 是否用粗纱：BOM子项/实际投料/投料进度 含 BC- 粗纱
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

    sections = []
    for ws in present_sorted:
        ws_rows = [
            item for item in results if item.work_order.production_dept_code == ws
        ]
        sections.append(
            {
                "workshop": ws,
                "label": WORKSHOP_LABELS.get(ws, ws),
                "rows": ws_rows,
                "subtotal": subtotals.get(ws),
                "totals": {
                    "material": sum(
                        (Decimal(r.direct_material or 0) for r in ws_rows),
                        Decimal("0"),
                    ),
                    "finished": sum(
                        (Decimal(r.finished_cost or 0) for r in ws_rows),
                        Decimal("0"),
                    ),
                    "wip": sum(
                        (Decimal(r.wip_cost or 0) for r in ws_rows),
                        Decimal("0"),
                    ),
                    "total": sum(
                        (Decimal(r.total_cost or 0) for r in ws_rows),
                        Decimal("0"),
                    ),
                },
            }
        )

    active_ws = (request.GET.get("workshop") or "").strip()
    if active_ws not in present_sorted:
        active_ws = present_sorted[0] if present_sorted else ""

    return render(
        request,
        "settlement/work_order_costs.html",
        {
            "period": period,
            "sections": sections,
            "active_ws": active_ws,
        },
    )


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

    name_by_code = {}
    for p in periods:
        for v in p.inventory_valuations.all():
            name_by_code.setdefault(v.material_code, v.material_name)
    for p in periods:
        for hp in p.historical_prices.all():
            name_by_code.setdefault(hp.material_code, hp.material_name)

    RAW = {"原料", "染色纤维", "纤维", "熟条", "浆粕"}
    FIN = {"纱线", "面料", "坯布", "布", "成品"}

    def big_cat(code):
        cat = cat_by_code.get(code, "")
        if cat in RAW:
            return "原料"
        if cat in FIN:
            return "成品"
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
            rp = (ra / rq).quantize(Decimal("0.0001")) if rq else Decimal("0")
            ip = (ia / iq).quantize(Decimal("0.0001")) if iq else Decimal("0")
            if hp is not None:
                if hp.receipt_price is not None:
                    rp = Decimal(hp.receipt_price)
                if hp.issue_price is not None:
                    ip = Decimal(hp.issue_price)
            cells.append((rq, ra, rp, iq, ia, ip))
        rows.append({
            "category": cat,
            "code": code,
            "name": name_by_code.get(code, ""),
            "uom": uom_by_code.get(code, ""),
            "cells": cells,
        })

    rows.sort(key=lambda r: ({"原料": 0, "成品": 1, "其它": 2}.get(r["category"], 9), r["code"]))
    return periods, rows


def build_price_table_workbook(period, n=6):
    """生成单价表（跨月对比）Excel 工作簿，返回 BytesIO。"""
    from io import BytesIO

    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    periods, rows = prepare_price_table_data(period, n)

    wb = Workbook()
    ws = wb.active
    ws.title = "单价表-跨月对比"
    ws.append([f"物料单价表（跨月对比）"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4 + 2 * len(periods))
    ws.cell(1, 1).font = Font(bold=True, size=14)
    headers = ["类别", "物料编码", "物料名称", "单位"]
    for p in periods:
        headers += [
            f"{p.month:%Y-%m} 入库单价", f"{p.month:%Y-%m} 出库单价",
        ]
    ws.append(headers)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill = PatternFill("solid", fgColor="1F4E79")
    for c in range(1, len(headers) + 1):
        cell = ws.cell(2, c)
        cell.fill = fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    raw_fill = PatternFill("solid", fgColor="FCE4D6")
    fin_fill = PatternFill("solid", fgColor="E2EFDA")
    for row in rows:
        line = [row["category"], row["code"], row["name"], row["uom"]]
        for rq, ra, rp, iq, ia, ip in row["cells"]:
            line += [rp, ip]
        ws.append(line)
        rf = raw_fill if row["category"] == "原料" else (fin_fill if row["category"] == "成品" else None)
        for c in range(1, len(headers) + 1):
            cell = ws.cell(ws.max_row, c)
            cell.border = border
            if c >= 5:
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = "#,##0.0000"
            if rf:
                cell.fill = rf
    for i, h in enumerate(headers, 1):
        ws.column_dimensions[chr(64 + i) if i <= 26 else "A"].width = 16
    ws.freeze_panes = "E3"
    stream = BytesIO()
    wb.save(stream)
    return stream


@login_required
def price_table_view(request, period_id):
    """单价表（跨月对比）：每个物料一行，每月一组「入库单价、出库单价」。

    物料列表 = 各期间库存计价 ∪ 上传的历史单价（上传过单价的物料即使该期间未结算也显示）；
    单价优先用上传的历史单价，未上传的用系统计算价（入库金额÷数量、出库金额÷数量）。
    默认取最近 6 期（含当前），支持 ?periods=N（1~12）与 ?category=all|raw|fin|other 筛选；
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

    name_by_code = {}
    for p in periods:
        for v in p.inventory_valuations.all():
            name_by_code.setdefault(v.material_code, v.material_name)
    for p in periods:
        for hp in p.historical_prices.all():
            name_by_code.setdefault(hp.material_code, hp.material_name)

    RAW = {"原料", "染色纤维", "纤维", "熟条", "浆粕"}
    FIN = {"纱线", "面料", "坯布", "布", "成品"}

    def big_cat(code):
        cat = cat_by_code.get(code, "")
        if cat in RAW:
            return "原料"
        if cat in FIN:
            return "成品"
        return "其它"

    category_filter = (request.GET.get("category") or "all").strip()
    want_cat = {"raw": "原料", "fin": "成品", "other": "其它"}.get(category_filter)

    rows = []
    for code in sorted(codes):
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
        rows.append({
            "category": cat,
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
        stream = build_price_table_workbook(period, n)
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
            "category_filter": category_filter or "all",
            "n_periods": len(periods),
            "periods_options": list(range(1, 13)),
            "price": _price,
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


LISTING_EXPORTS = {
    "inventory_ledger": {"title": "库存流水"},
    "purchase_contract": {"title": "采购合同"},
    "purchase_invoice": {"title": "采购发票匹配"},
    "purchase_accrual": {"title": "月末采购暂估"},
    "purchase_difference": {"title": "本月采购暂估差异"},
    "purchase_reversal": {"title": "发票暂估回冲明细"},
    "sales_invoice": {"title": "销售发票匹配"},
    "sales_uninvoiced": {"title": "月末已出库未开票"},
    "work_order_match": {"title": "工单匹配"},
    "inventory_valuation": {"title": "库存金额结转"},
    "work_order_cost": {"title": "工单成本"},
    "work_order": {"title": "生产工单"},
    "order_material": {"title": "工单投料"},
    "order_completion": {"title": "工单报工"},
    "work_order_ledger": {"title": "工单台账"},
    "material_issue_ledger": {"title": "投料台账"},
    "work_report_ledger": {"title": "报工台账"},
    "reconciliation": {"title": "勾稽结果"},
    "workshop_settlement": {"title": "车间结转"},
    "opening_inventory": {"title": "期初存货"},
    "ending_inventory": {"title": "期末存货"},
}



def _uploaded_ledger_rows(period, kind, queryset):
    """按用户上传文件的格式导出台账：
    - 格式（表头列名/顺序）取该类型行数最多的上传文件（即主文件）；
    - 数据取数据库该类型全部记录，按上传表头逐列映射字段输出。
    无上传文件或读取失败时返回 (None, None, None)。
    """
    from pathlib import Path

    from settlement.services.exports import read_uploaded_rows

    # 找到行数最多的上传批次作为格式基准
    best_header = None
    best_count = -1
    for batch in (
        period.import_batches.filter(kind=kind)
        .exclude(file="")
        .order_by("-created_at")
    ):
        try:
            if not Path(batch.file.path).is_file():
                continue
            header, rows = read_uploaded_rows(Path(batch.file.path))
        except Exception:
            continue
        if header and len(rows) > best_count:
            best_header = header
            best_count = len(rows)
    if not best_header:
        return None, None, None

    # 上传表头 -> 模型字段（用 ALIASES 反查，兼容列名差异）
    alias_map = {}
    for field, names in ALIASES[kind].items():
        for name in names:
            alias_map[normalize_header(name)] = field
    fields = [alias_map.get(normalize_header(h)) for h in best_header]

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
    return best_header, rows_out, best_count


def _listing_rows(request, period, kind):
    """按 kind 返回 (工作表名, 表头, 数据行) 列表，筛选逻辑与列表页一致。"""
    sheets = []

    if kind == "inventory_ledger":
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
                ["合同号", "合同日期", "供应商", "品名编码", "品名", "合同数量",
                 "剩余数量", "不含税单价", "来源"],
                (
                    (
                        item.contract_no,
                        item.contract_date,
                        item.supplier_name or item.supplier_code,
                        item.material_code,
                        item.material_name,
                        item.quantity,
                        item.remaining_quantity,
                        item.unit_price_excl_tax,
                        "自动承接" if item.is_system_generated else "导入",
                    )
                    for item in contracts
                ),
            )
        )

    elif kind == "purchase_invoice":
        invoices = period.purchase_invoices.annotate(
            matched_total=Sum("matchallocation__matched_quantity")
        ).prefetch_related(
            "matchallocation_set__inventory_transaction__period",
            "matchallocation_set__opening_purchase_accrual",
        )
        invoices = _text_filters(
            request,
            invoices,
            {
                "invoice_no": "invoice_no",
                "contract_no": "contract_no",
                "material": "material_code",
            },
        )
        invoices = _party_filter(
            request, invoices, "supplier", "supplier_code", "supplier_name"
        )
        invoices = _date_filter(request, invoices, "invoice_date", "invoice_date")
        invoices = _choice_filter(request, invoices, "match_mode", "match_mode")
        invoices = invoices.order_by("invoice_date", "invoice_no", "id")

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
                ["发票号", "收票日期", "发票期间", "合同号", "供应商", "物料编码",
                 "物料名称", "单位", "签收数量", "已匹配数量", "未匹配数量", "不含税金额",
                 "匹配的入库/期初暂估", "处理方式"],
                (
                    (
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
                ),
            )
        )

    elif kind == "purchase_accrual":
        rows = build_purchase_accrual_rows(period)
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
            rows = [r for r in rows if supplier in (f"{r.supplier_code} {r.supplier_name}").lower()]
        if material:
            rows = [r for r in rows if material in (f"{r.material_code} {r.material_name}").lower()]
        if contract_no:
            rows = [r for r in rows if contract_no in r.contract_nos.lower()]
        if status:
            rows = [r for r in rows if r.status == status]
        sheets.append(
            (
                "月末采购暂估",
                ["入库期间", "入库日期", "入库单号", "供应商", "物料编码", "物料名称",
                 "入库数量", "已到票数量", "未到票数量", "暂估单价", "未到票金额",
                 "状态", "匹配合同"],
                (
                    (
                        r.receipt_month,
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
                ["发票号", "发票日期", "供应商", "物料编码", "物料名称", "入库期间",
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
                 "原入库月份", "原入库日期", "物料编码", "物料名称", "供应商",
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

    elif kind == "sales_invoice":
        invoices = period.sales_invoices.annotate(
            matched_total=Sum("matchallocation__matched_quantity")
        ).prefetch_related(
            "matchallocation_set__inventory_transaction__period",
            "matchallocation_set__opening_goods_issued",
        )
        invoices = _text_filters(
            request,
            invoices,
            {
                "invoice_no": "invoice_no",
                "sales_order_no": "sales_order_no",
                "material": "material_code",
            },
        )
        invoices = _party_filter(
            request, invoices, "customer", "customer_code", "customer_name"
        )
        invoices = _date_filter(request, invoices, "invoice_date", "invoice_date")
        invoices = _choice_filter(request, invoices, "match_mode", "match_mode")
        invoices = invoices.order_by("invoice_date", "invoice_no", "id")

        def _sales_matches(invoice):
            parts = []
            for a in invoice.matchallocation_set.all():
                if a.inventory_transaction:
                    parts.append(
                        f"{a.inventory_transaction.period.month:%Y-%m}·"
                        f"{a.inventory_transaction.document_no or '流水'}:{a.matched_quantity}"
                    )
                else:
                    parts.append(
                        f"{a.opening_goods_issued.original_period:%Y-%m}·"
                        f"{a.opening_goods_issued.document_no}(期初发出):{a.matched_quantity}"
                    )
            return "；".join(parts)

        sheets.append(
            (
                "销售发票匹配",
                ["发票号", "发票日期", "发票期间", "销售订单号", "客户", "产品编码",
                 "产品名称", "开票数量", "已匹配数量", "未匹配数量", "不含税金额",
                 "匹配的出库/期初发出", "处理方式"],
                (
                    (
                        item.invoice_no,
                        item.invoice_date,
                        item.invoice_period,
                        item.sales_order_no,
                        item.customer_name or item.customer_code,
                        item.material_code,
                        item.material_name,
                        item.quantity,
                        item.matched_total or 0,
                        item.remaining_quantity,
                        item.amount_excl_tax,
                        _sales_matches(item),
                        item.get_match_mode_display(),
                    )
                    for item in invoices
                ),
            )
        )

    elif kind == "sales_uninvoiced":
        rows = build_sales_uninvoiced_rows(period)
        issue_month = str(request.GET.get("issue_month", "")).strip()
        document_no = str(request.GET.get("document_no", "")).strip().lower()
        sales_order_no = str(request.GET.get("sales_order_no", "")).strip().lower()
        customer = str(request.GET.get("customer", "")).strip().lower()
        material = str(request.GET.get("material", "")).strip().lower()
        status = str(request.GET.get("status", "")).strip()
        cross_period = str(request.GET.get("cross_period", "")).strip()
        if issue_month:
            rows = [r for r in rows if r.issue_month.strftime("%Y-%m") == issue_month]
        if document_no:
            rows = [r for r in rows if document_no in (r.document_no or "").lower()]
        if sales_order_no:
            rows = [r for r in rows if sales_order_no in (r.sales_order_no or "").lower()]
        if customer:
            rows = [r for r in rows if customer in (f"{r.customer_code} {r.customer_name}").lower()]
        if material:
            rows = [r for r in rows if material in (f"{r.material_code} {r.material_name}").lower()]
        if status:
            rows = [r for r in rows if r.status == status]
        if cross_period in {"yes", "no"}:
            rows = [r for r in rows if r.cross_period == (cross_period == "yes")]
        sheets.append(
            (
                "月末未开票",
                ["出库期间", "出库日期", "出库单号", "销售订单号", "客户", "产品编码",
                 "产品名称", "出库数量", "累计开票数量", "未开票数量", "出库成本",
                 "已开票成本", "未开票成本", "匹配发票", "最近开票期间", "未开票天数", "状态"],
                (
                    (
                        r.issue_month,
                        r.document_date,
                        r.document_no,
                        r.sales_order_no,
                        r.customer_name or r.customer_code,
                        r.material_code,
                        r.material_name,
                        r.issue_quantity,
                        r.invoiced_quantity,
                        r.outstanding_quantity,
                        r.issue_amount,
                        r.invoiced_cost,
                        r.outstanding_cost,
                        r.invoice_nos,
                        r.latest_invoice_period,
                        r.outstanding_days,
                        r.status,
                    )
                    for r in rows
                ),
            )
        )

    elif kind == "work_order_match":
        # 与"工单匹配"页一致：工单 + 投料/报工台账 + 报工入库汇总
        masters = {wo.order_no: wo for wo in period.work_orders.all()}
        issue_nos = set(
            period.material_issues.exclude(work_order_no="").values_list(
                "work_order_no", flat=True
            )
        )
        report_nos = set(
            period.work_reports.exclude(work_order_no="").values_list(
                "work_order_no", flat=True
            )
        )
        all_nos = sorted(set(masters) | issue_nos | report_nos)
        issue_agg = {
            row["work_order_no"]: row
            for row in period.material_issues.exclude(work_order_no="")
            .values("work_order_no")
            .annotate(issue_qty=Sum("actual_issue_quantity"))
        }
        report_agg = {
            row["work_order_no"]: row
            for row in period.work_reports.exclude(work_order_no="")
            .values("work_order_no")
            .annotate(report_qty=Sum("quantity"))
        }
        receipt_agg = {}
        for row in (
            period.inventory_transactions.filter(business_type="报工")
            .values("document_no")
            .annotate(txn_qty=Sum("quantity"))
        ):
            receipt_agg[row["document_no"]] = row
        match_rows = []
        for order_no in all_nos:
            wo = masters.get(order_no)
            product_code, _name = _work_order_product(period, order_no)
            issue = issue_agg.get(order_no)
            report = report_agg.get(order_no)
            receipt = receipt_agg.get(order_no)
            match_rows.append(
                (
                    order_no,
                    WORKSHOP_LABELS.get(
                        wo.production_dept_code if wo else "",
                        wo.production_dept_code if wo else "",
                    ),
                    product_code,
                    Decimal(wo.plan_quantity) if wo else Decimal("0"),
                    issue["issue_qty"] if issue else Decimal("0"),
                    report["report_qty"] if report else Decimal("0"),
                    receipt["txn_qty"] if receipt else Decimal("0"),
                )
            )
        sheets.append(
            (
                "工单匹配",
                ["工单编号", "车间/部门", "产品编码", "计划数量",
                 "投料数量", "报工数量", "报工入库数量"],
                match_rows,
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
        # 投料台账：格式/字段按用户上传文件原样，数据=本期全部投料记录；无上传时回退数据库明细
        header, raw_rows, _ = _uploaded_ledger_rows(
            period, ImportBatch.Kind.MATERIAL_ISSUE, period.material_issues.all()
        )
        if header:
            sheets.append(("投料台账", header, raw_rows))
            return sheets
        issues = period.material_issues.all()
        issues = _text_filters(
            request,
            issues,
            {
                "order_no": "work_order_no",
                "material": "material_code",
            },
        )
        sheets.append(
            (
                "投料台账",
                ["生产任务编号", "工序编号", "工序名称", "物料编码", "物料名称", "物料分类",
                 "物料属性", "物料规格", "投料数量", "已撤回数量", "实际投料数量", "投料单位",
                 "辅助数量", "辅助单位", "标识码", "批次号", "质量状态", "投料仓位",
                 "被替代物料编码", "被替代物料名称", "投料时间", "投料人员", "执行人员",
                 "工单编号", "主产出物料编码", "主产出物料名称", "销售订单编号", "投料记录编号"],
                (
                    (
                        item.task_no,
                        item.operation_no,
                        item.operation_name,
                        item.material_code,
                        item.material_name,
                        item.material_category,
                        item.material_attribute,
                        item.material_spec,
                        item.issue_quantity,
                        item.returned_quantity,
                        item.actual_issue_quantity,
                        item.issue_uom,
                        item.auxiliary_quantity,
                        item.auxiliary_uom,
                        item.identifier,
                        item.batch_no,
                        item.quality_status,
                        item.issue_location,
                        item.replaced_material_code,
                        item.replaced_material_name,
                        item.issue_time,
                        item.issuer,
                        item.executor,
                        item.work_order_no,
                        item.main_output_code,
                        item.main_output_name,
                        item.sales_order_no,
                        item.record_no,
                    )
                    for item in issues.order_by("id")
                ),
            )
        )

    elif kind == "work_report_ledger":
        # 报工台账：格式/字段按用户上传文件原样，数据=本期全部报工记录；无上传时回退数据库明细
        header, raw_rows, _ = _uploaded_ledger_rows(
            period, ImportBatch.Kind.WORK_REPORT, period.work_reports.all()
        )
        if header:
            sheets.append(("报工台账", header, raw_rows))
            return sheets
        reports = period.work_reports.all()
        reports = _text_filters(
            request,
            reports,
            {
                "order_no": "work_order_no",
                "material": "material_code",
                "quality_status": "quality_status",
            },
        )
        sheets.append(
            (
                "报工台账",
                ["报工时间", "生产任务编号", "工序编号", "工序名称", "报工物料编码", "报工物料名称",
                 "生产人员", "质量状态", "不良原因", "报工数量", "报工单位", "辅助数量", "辅助单位",
                 "批次号", "标识码", "流转卡编号", "工时", "工时单位", "备注", "部门", "设备",
                 "工单编号", "工单物料编码", "工单物料名称", "订单号", "纸管颜色", "纺纱班别",
                 "含水率", "批号流水", "织布序号", "织布机台", "产出仓位", "工单部门"],
                (
                    (
                        item.report_time,
                        item.task_no,
                        item.operation_no,
                        item.operation_name,
                        item.material_code,
                        item.material_name,
                        item.operator,
                        item.quality_status,
                        item.defect_reason,
                        item.quantity,
                        item.uom,
                        item.auxiliary_quantity,
                        item.auxiliary_uom,
                        item.batch_no,
                        item.identifier,
                        item.route_card_no,
                        item.work_hours,
                        item.work_hours_uom,
                        item.remark,
                        item.department,
                        item.equipment,
                        item.work_order_no,
                        item.work_order_material_code,
                        item.work_order_material_name,
                        item.order_no,
                        item.paper_tube_color,
                        item.spinning_shift,
                        item.moisture_content,
                        item.dye_batch_flow,
                        item.weaving_sequence,
                        item.weaving_machine,
                        item.output_location,
                        item.work_order_department,
                    )
                    for item in reports.order_by("id")
                ),
            )
        )

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
        rows, _totals = build_material_issue_rows(
            period,
            order_no=request.GET.get("order_no", "").strip(),
            workshop=request.GET.get("workshop", "").strip(),
            material=request.GET.get("material", "").strip(),
        )
        sheets.append(
            (
                "工单投料",
                ["工单编号", "车间", "物料编码", "物料名称", "规格", "单位",
                 "实际投料数量", "投料金额"],
                (
                    (
                        r["order_no"],
                        WORKSHOP_LABELS.get(r["workshop"], r["workshop"]),
                        r["material_code"],
                        r["material_name"],
                        r["specification"],
                        r["uom"],
                        r["issue_quantity"],
                        r["amount"],
                    )
                    for r in rows
                ),
            )
        )

    elif kind == "order_completion":
        # 工单报工：报工数量、完工成本、报工入库
        masters = {wo.order_no: wo for wo in period.work_orders.all()}
        report_nos = set(
            period.work_reports.exclude(work_order_no="").values_list(
                "work_order_no", flat=True
            )
        )
        all_nos = sorted(set(masters) | report_nos)
        report_agg = {
            row["work_order_no"]: row
            for row in period.work_reports.exclude(work_order_no="")
            .values("work_order_no")
            .annotate(report_qty=Sum("quantity"))
        }
        receipt_agg = {}
        for row in (
            period.inventory_transactions.filter(business_type="报工")
            .values("document_no")
            .annotate(txn_qty=Sum("quantity"), txn_amount=Sum("amount"))
        ):
            receipt_agg[row["document_no"]] = row
        results = {
            item.work_order.order_no: item
            for item in period.work_order_cost_results.select_related(
                "work_order"
            ).all()
        }
        completion_rows = []
        for order_no in all_nos:
            wo = masters.get(order_no)
            result = results.get(order_no)
            report = report_agg.get(order_no)
            receipt = receipt_agg.get(order_no)
            product_code, product_name = _work_order_product(period, order_no)
            completion_rows.append(
                (
                    order_no,
                    WORKSHOP_LABELS.get(
                        wo.production_dept_code if wo else "",
                        wo.production_dept_code if wo else "",
                    ),
                    product_code,
                    product_name,
                    Decimal(wo.plan_quantity) if wo else Decimal("0"),
                    Decimal(report["report_qty"]) if report else Decimal("0"),
                    result.finished_cost if result else Decimal("0"),
                    result.finished_unit_cost if result else Decimal("0"),
                    result.wip_cost if result else Decimal("0"),
                    result.total_cost if result else Decimal("0"),
                    Decimal(receipt["txn_qty"]) if receipt else Decimal("0"),
                    Decimal(receipt["txn_amount"])
                    if receipt and receipt["txn_amount"] is not None
                    else Decimal("0"),
                    wo.status if wo else "",
                )
            )
        sheets.append(
            (
                "工单报工",
                ["工单编号", "车间", "产品编码", "产品名称", "计划数量", "报工数量",
                 "完工成本", "单位完工成本", "在产品成本", "总成本",
                 "报工入库数量", "报工入库金额", "状态"],
                completion_rows,
            )
        )

    elif kind == "inventory_valuation":
        sheets.append(
            (
                "库存金额结转",
                ["物料编码", "物料名称", "期初数量", "期初金额", "入库数量", "入库金额",
                 "出库数量", "出库金额", "月末加权平均价", "期末数量", "期末金额", "收敛"],
                (
                    (
                        item.material_code,
                        item.material_name,
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
                    for item in period.inventory_valuations.all()
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
        results = list(
            results.order_by(
                "work_order__production_dept_code", "work_order__order_no"
            )
        )
        from settlement.services.workorder_costing import MANUFACTURING_FIELDS
        from settlement.models import BomLine, IssueProgress

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
        for item in results:
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
            # 完工率 = 完工÷(完工+在产品)：纺纱不用粗纱工单的在产品=投料-完工，即 完工÷投料
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

        # 与页面/结算工作簿一致：每个制造费用列后加分配率列，每车间末尾加合计行
        # 纺纱车间无天然气费用：单独导出纺纱时去掉天然气两列（与页面一致）
        workshop_param = (request.GET.get("workshop") or "").strip()
        hide_gas = workshop_param == "tefang"
        headers = [
            "核算期间", "工单编号", "车间", "产品编码", "产品名称", "计划数量", "累计投料", "投料状态", "完成状态",
            "累计完工", "累计完工率", "是否用粗纱", "报工数量",
            "在产品数量", "材料成本", "水电", "水电分配率%", "蒸汽", "蒸汽分配率%",
            "天然气", "天然气分配率%", "折旧", "折旧分配率%", "机物料", "机物料分配率%",
            "染料助剂", "染料助剂分配率%", "工资", "工资分配率%", "工单总成本",
            "完工成本", "完工单位成本", "在产品成本",
        ]
        if hide_gas:
            headers = [h for h in headers if h not in ("天然气", "天然气分配率%")]
        rows = []
        highlight = set()
        if results:
            mat_total = {}
            fee_total = {f: {} for f in MANUFACTURING_FIELDS}
            equiv_total = {}
            for item in results:
                ws = item.work_order.production_dept_code
                mat_total[ws] = mat_total.get(ws, Decimal("0")) + Decimal(
                    item.direct_material or 0
                )
                for f in MANUFACTURING_FIELDS:
                    fee_total[f][ws] = fee_total[f].get(ws, Decimal("0")) + Decimal(
                        getattr(item, f) or 0
                    )
                equiv_total[ws] = equiv_total.get(ws, Decimal("0")) + Decimal(
                    item.finished_quantity or 0
                ) + Decimal(item.wip_quantity or 0) * Decimal("0.5")

            def _rate(fee_field, ws):
                denom = mat_total.get(ws, Decimal("0"))
                return (
                    (fee_total[fee_field].get(ws, Decimal("0")) / denom) * Decimal("100")
                    if denom
                    else Decimal("0")
                )

            def _wage_rate(item, ws):
                denom = equiv_total.get(ws, Decimal("0"))
                if not denom:
                    return Decimal("0")
                equivalent = Decimal(item.finished_quantity or 0) + Decimal(
                    item.wip_quantity or 0
                ) * Decimal("0.5")
                return (equivalent / denom) * Decimal("100")

            SUM_FIELDS = (
                "direct_material", "water_power", "steam", "natural_gas",
                "depreciation", "consumables", "dye_auxiliary", "wages",
                "total_cost", "finished_cost", "wip_cost",
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
                rate_text = (
                    f"{item.finish_rate_pct:.2f}%" if item.finish_rate_pct is not None else "—"
                )
                row = [
                    f"{period.month:%Y-%m}",
                    item.work_order.order_no,
                    WORKSHOP_LABELS.get(ws, ws),
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
                ]
                if not hide_gas:
                    row += [item.natural_gas, _rate("natural_gas", ws)]
                row += [
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
                ]
                return tuple(row)

            def _subtotal_row(ws, st):
                row = [f"{WORKSHOP_LABELS.get(ws, ws)}合计（{st['count']}张工单）"]
                row += [""] * 13
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
                    st["wages"], "",
                    st["total_cost"],
                    st["finished_cost"], "",
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

    REPORT_CATALOG = [
        {"kind": "opening_inventory", "name": "期初存货（上传原样）",
         "desc": "期初库存余额台账，按上传文件原样导出（保留仓库/批次明细，行数与上传一致）"},
        {"kind": "ending_inventory", "name": "期末存货",
         "desc": "期末库存余额台账（物料、结存数量、单价、金额，系统自动计算）"},
        {"kind": "inventory_ledger", "name": "库存流水",
         "desc": "本期间全部库存出入库流水（单据、物料、数量、单价、金额）"},
        {"kind": "purchase_contract", "name": "采购合同",
         "desc": "采购合同台账（合同号、供应商、品名、数量、不含税单价、剩余数量）"},
        {"kind": "purchase_invoice", "name": "采购发票匹配",
         "desc": "采购发票及核销的入库/期初暂估明细"},
        {"kind": "purchase_accrual", "name": "月末采购暂估",
         "desc": "截至月末仍未到票的采购入库及暂估金额"},
        {"kind": "purchase_difference", "name": "本月采购暂估差异",
         "desc": "本月发票核销产生的暂估价差（交总账处理）"},
        {"kind": "purchase_reversal", "name": "发票暂估回冲明细",
         "desc": "每张采购发票核销了哪笔入库/期初暂估：暂估金额、发票金额、价差、累计核销与剩余未到票"},
        {"kind": "sales_invoice", "name": "销售发票匹配",
         "desc": "销售发票及核销的出库/期初发出商品明细"},
        {"kind": "sales_uninvoiced", "name": "月末已出库未开票",
         "desc": "截至月末仍未开票的销售出库及发出商品成本"},
        {"kind": "inventory_valuation", "name": "库存金额结转",
         "desc": "各物料期初/入库/出库/期末数量金额及月末加权平均价"},
        {"kind": "work_order_cost", "name": "工单成本",
         "desc": "各生产工单按投料/报工核算的完工与在制成本明细"},
        {"kind": "work_order", "name": "生产工单",
         "desc": "生产工单主档台账（数量、状态、计划时间）"},
        {"kind": "work_order_ledger", "name": "工单台账（原样明细）",
         "desc": "ERP 工单主档原样台账：工单号、类型、车间、计划数量、状态、计划时间、订单号、物料颜色等全字段"},
        {"kind": "material_issue_ledger", "name": "投料台账（原样明细）",
         "desc": "ERP 投料记录原样台账：物料、投料/实际投料数量、批次、仓位、投料时间、人员、工单号等全字段"},
        {"kind": "work_report_ledger", "name": "报工台账（原样明细）",
         "desc": "ERP 报工记录原样台账：报工数量、质量状态、工时、部门、设备、工单号、颜色/批次等全字段"},
        {"kind": "order_material", "name": "工单投料",
         "desc": "按工单 × 物料汇总投料记录台账数量与金额"},
        {"kind": "order_completion", "name": "工单报工",
         "desc": "工单完工汇报：报工数量、完工成本、报工入库数量与金额"},
        {"kind": "reconciliation", "name": "勾稽结果与异常清单",
         "desc": "全部勾稽项目（含异常清单），两个工作表"},
        {"kind": "workshop_settlement", "name": "车间结转",
         "desc": "车间结转顺序、状态与成本结果"},
    ]
    SPECIAL_REPORTS = [
        {"key": "price_table", "name": "单价表（历史单价·跨月对比）",
         "desc": "各物料最近多期的入库/出库单价（上传的历史单价优先，未上传用系统计算价）；含新增物料自动补全"},
        {"key": "settlement_workbook", "name": "结算工作簿（完整）",
         "desc": "含 16 个工作表的完整结算底稿：库存计价、跨期结转、月末暂估、勾稽等"},
        {"key": "uploaded_files", "name": "上传的原始文件",
         "desc": "本期间全部上传的原始文件（CSV/XLSX）"},
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
            "reports": REPORT_CATALOG,
            "special_reports": SPECIAL_REPORTS,
        },
    )


@login_required
def export_period(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    if not period.inventory_valuations.exists():
        messages.error(request, "请先执行结算，再导出结果。")
        return redirect("period_detail", period_id=period.id)
    return settlement_excel_response(period)
