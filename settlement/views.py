from pathlib import Path

from django.conf import settings
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.db.models import Q, Sum
from django.http import FileResponse, Http404, HttpResponseBadRequest
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
    ImportBatch,
    InventoryTransaction,
    MatchAllocation,
    MatchingMode,
    OpeningGoodsIssued,
    OpeningPurchaseAccrual,
    ProductionOrder,
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
from settlement.services.exports import settlement_excel_response
from settlement.services.imports import (
    create_import_batch,
    process_batch,
)
from settlement.services.matching import (
    auto_match_all,
    is_order_material,
    is_production_receipt,
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
    summarize_purchase_accruals,
    summarize_purchase_differences,
)
from settlement.services.production_costs import build_production_cost_summary
from settlement.services.rollforward import (
    carry_forward_context,
    invalidate_downstream_periods,
)
from settlement.services.sales_uninvoiced import (
    build_goods_issued_rollforward,
    build_sales_uninvoiced_rows,
    summarize_sales_uninvoiced,
)


TEMPLATE_FILES = {
    ImportBatch.Kind.INVENTORY: "库存出入库流水.csv",
    ImportBatch.Kind.PURCHASE_CONTRACT: "采购合同.csv",
    ImportBatch.Kind.PURCHASE_INVOICE: "采购发票.csv",
    ImportBatch.Kind.SALES_INVOICE: "销售发票.csv",
    ImportBatch.Kind.PRODUCTION_ORDER: "生产订单.csv",
    ImportBatch.Kind.WORKSHOP_COST: "车间费用.csv",
    ImportBatch.Kind.OPENING_WIP: "期初在产品.csv",
    ImportBatch.Kind.OPENING_PURCHASE_ACCRUAL: "期初采购暂估明细.csv",
    ImportBatch.Kind.OPENING_GOODS_ISSUED: "期初发出商品明细.csv",
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
    if is_production_receipt(txn):
        return "生产入库"
    if is_order_material(txn):
        return "生产领退料"
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
        "orders": period.production_orders.count(),
        "valuations": period.inventory_valuations.count(),
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


@login_required
def order_matches(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    orders = period.production_orders.prefetch_related(
        "matchallocation_set__inventory_transaction"
    )
    orders = _text_filters(
        request,
        orders,
        {
            "order_no": "order_no",
            "workshop": "workshop",
            "product": "product_code",
        },
    )
    orders = _choice_filter(request, orders, "status", "status")
    orders = _choice_filter(
        request, orders, "match_mode", "match_mode"
    )
    return render(
        request,
        "settlement/order_matches.html",
        {
            "period": period,
            "page": _page(request, orders),
            "match_modes": MatchingMode.choices,
            "order_statuses": ProductionOrder.Status.choices,
        },
    )


@login_required
def inventory_matches(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    txns = period.inventory_transactions.prefetch_related(
        "allocations__contract_line",
        "allocations__purchase_invoice_line",
        "allocations__sales_invoice_line",
        "allocations__production_order",
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
    "production_order": {
        "model": ProductionOrder,
        "kind": None,
        "title": "生产订单",
        "list_view": "order_matches",
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
    txns = period.inventory_transactions.prefetch_related("allocations")
    return [
        txn
        for txn in txns
        if is_order_material(txn) or is_production_receipt(txn)
    ]


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
                if source_type == "production_order":
                    if is_production_receipt(txn):
                        kind = MatchAllocation.Kind.PRODUCTION_RECEIPT
                    elif is_order_material(txn):
                        kind = MatchAllocation.Kind.ORDER_MATERIAL
                    else:
                        raise ValueError("该库存流水不是生产领退料或生产入库。")
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
    else:
        allocations = allocations.filter(production_order=source)
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
    elif is_production_receipt(txn):
        groups.append(
            {
                "kind": MatchAllocation.Kind.PRODUCTION_RECEIPT,
                "label": "生产订单",
                "items": period.production_orders.exclude(
                    match_mode=MatchingMode.IGNORED
                )[:300],
            }
        )
    elif is_order_material(txn):
        groups.append(
            {
                "kind": MatchAllocation.Kind.ORDER_MATERIAL,
                "label": "生产订单",
                "items": period.production_orders.exclude(
                    match_mode=MatchingMode.IGNORED
                )[:300],
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
                "production_order__period",
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
        else "order_matches"
        if kind
        in {
            MatchAllocation.Kind.ORDER_MATERIAL,
            MatchAllocation.Kind.PRODUCTION_RECEIPT,
        }
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
def order_costs(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    totals = period.order_cost_results.aggregate(
        total=Sum("total_cost"),
        finished=Sum("finished_cost"),
        wip=Sum("wip_cost"),
    )
    return render(
        request,
        "settlement/order_costs.html",
        {
            "period": period,
            "page": _page(
                request, period.order_cost_results.select_related("order").all()
            ),
            "totals": totals,
        },
    )


@login_required
def production_cost_summary(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    selected_workshop = request.GET.get("workshop", "").strip()
    summary = build_production_cost_summary(period, selected_workshop)
    return render(
        request,
        "settlement/production_cost_summary.html",
        {
            "period": period,
            **summary,
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


@login_required
def export_period(request, period_id):
    period = get_object_or_404(AccountingPeriod, id=period_id)
    if not period.inventory_valuations.exists():
        messages.error(request, "请先执行结算，再导出结果。")
        return redirect("period_detail", period_id=period.id)
    return settlement_excel_response(period)
