from __future__ import annotations

from calendar import monthrange
from dataclasses import dataclass
from datetime import date, timedelta
from decimal import Decimal, ROUND_HALF_UP

from django.db.models import Max, Sum

from settlement.models import (
    InventoryTransaction,
    MatchAllocation,
    MatchingMode,
    OpeningGoodsIssued,
)
from settlement.services.matching import is_sales_issue


ZERO = Decimal("0")
CENT = Decimal("0.01")


def money(value) -> Decimal:
    return Decimal(value or ZERO).quantize(CENT, rounding=ROUND_HALF_UP)


def _month_end(month: date) -> date:
    return month.replace(day=monthrange(month.year, month.month)[1])


def _previous_month(month: date) -> date:
    return (month.replace(day=1) - timedelta(days=1)).replace(day=1)


def _is_sample_issue(txn: InventoryTransaction) -> bool:
    return "样品" in str(txn.business_type or "")


@dataclass
class SalesUninvoicedRow:
    transaction: InventoryTransaction | None
    opening_item: OpeningGoodsIssued | None
    issue_period: object | None
    issue_month: object
    document_date: object
    document_no: str
    sales_order_no: str
    customer_code: str
    customer_name: str
    material_code: str
    material_name: str
    issue_quantity: Decimal
    invoiced_quantity: Decimal
    outstanding_quantity: Decimal
    issue_amount: Decimal | None
    invoiced_cost: Decimal | None
    outstanding_cost: Decimal | None
    invoice_nos: str
    latest_invoice_period: object | None
    outstanding_days: int
    cross_period: bool
    status: str


def _sales_issues_as_of(month: date):
    queryset = (
        InventoryTransaction.objects.filter(period__month__lte=month)
        .exclude(match_mode=MatchingMode.IGNORED)
        .select_related("period")
        .prefetch_related("allocations__sales_invoice_line")
        .order_by("period__month", "txn_date", "document_no", "id")
    )
    return [
        txn
        for txn in queryset
        if is_sales_issue(txn) and not _is_sample_issue(txn)
    ]


def _build_sales_rows_as_of(month: date, include_closed=False):
    rows = []
    for txn in _sales_issues_as_of(month):
        allocations = txn.allocations.filter(
            kind=MatchAllocation.Kind.SALES_INVOICE,
            period__month__lte=month,
            sales_invoice_line__invoice_period__lte=month,
        ).exclude(
            sales_invoice_line__match_mode=MatchingMode.IGNORED
        )
        totals = allocations.aggregate(
            quantity=Sum("matched_quantity"),
            latest_period=Max("sales_invoice_line__invoice_period"),
        )
        issue_quantity = abs(Decimal(txn.quantity or ZERO))
        invoiced_quantity = min(
            Decimal(totals["quantity"] or ZERO),
            issue_quantity,
        )
        outstanding_quantity = max(
            issue_quantity - invoiced_quantity,
            ZERO,
        )
        if outstanding_quantity == ZERO and not include_closed:
            continue

        issue_amount = (
            money(abs(Decimal(txn.amount)))
            if txn.amount is not None
            else None
        )
        if issue_amount is None or issue_quantity == ZERO:
            invoiced_cost = None
            outstanding_cost = None
        else:
            outstanding_cost = money(
                issue_amount * outstanding_quantity / issue_quantity
            )
            invoiced_cost = money(issue_amount - outstanding_cost)

        invoice_nos = "；".join(
            allocations.order_by(
                "sales_invoice_line__invoice_period",
                "sales_invoice_line__invoice_no",
                "id",
            )
            .values_list("sales_invoice_line__invoice_no", flat=True)
            .distinct()
        )
        if issue_amount is None:
            status = "缺少出库成本"
        elif outstanding_quantity == ZERO:
            status = "全部开票"
        elif invoiced_quantity > ZERO:
            status = "部分开票"
        else:
            status = "全部未开票"
        rows.append(
            SalesUninvoicedRow(
                transaction=txn,
                opening_item=None,
                issue_period=txn.period,
                issue_month=txn.period.month,
                document_date=txn.txn_date,
                document_no=txn.document_no,
                sales_order_no=txn.order_no,
                customer_code=txn.customer_code,
                customer_name=txn.customer_name,
                material_code=txn.material_code,
                material_name=txn.material_name,
                issue_quantity=issue_quantity,
                invoiced_quantity=invoiced_quantity,
                outstanding_quantity=outstanding_quantity,
                issue_amount=issue_amount,
                invoiced_cost=invoiced_cost,
                outstanding_cost=outstanding_cost,
                invoice_nos=invoice_nos,
                latest_invoice_period=totals["latest_period"],
                outstanding_days=max((_month_end(month) - txn.txn_date).days, 0),
                cross_period=txn.period.month < month,
                status=status,
            )
        )
    opening_rows = (
        OpeningGoodsIssued.objects.filter(period__month__lte=month)
        .exclude(match_mode=MatchingMode.IGNORED)
        .prefetch_related("allocations__sales_invoice_line")
        .order_by(
            "original_period",
            "document_date",
            "document_no",
            "line_no",
            "id",
        )
    )
    for opening in opening_rows:
        allocations = opening.allocations.filter(
            kind=MatchAllocation.Kind.SALES_INVOICE,
            period__month__lte=month,
            sales_invoice_line__invoice_period__lte=month,
        ).exclude(
            sales_invoice_line__match_mode=MatchingMode.IGNORED
        )
        totals = allocations.aggregate(
            quantity=Sum("matched_quantity"),
            latest_period=Max("sales_invoice_line__invoice_period"),
        )
        issue_quantity = Decimal(opening.opening_quantity or ZERO)
        invoiced_quantity = min(
            Decimal(totals["quantity"] or ZERO),
            issue_quantity,
        )
        outstanding_quantity = max(
            issue_quantity - invoiced_quantity,
            ZERO,
        )
        if outstanding_quantity == ZERO and not include_closed:
            continue
        issue_amount = money(opening.opening_amount)
        outstanding_cost = money(
            issue_amount * outstanding_quantity / issue_quantity
        )
        invoiced_cost = money(issue_amount - outstanding_cost)
        invoice_nos = "；".join(
            allocations.order_by(
                "sales_invoice_line__invoice_period",
                "sales_invoice_line__invoice_no",
                "id",
            )
            .values_list("sales_invoice_line__invoice_no", flat=True)
            .distinct()
        )
        status = (
            "全部开票"
            if outstanding_quantity == ZERO
            else "部分开票"
            if invoiced_quantity > ZERO
            else "全部未开票"
        )
        rows.append(
            SalesUninvoicedRow(
                transaction=None,
                opening_item=opening,
                issue_period=None,
                issue_month=opening.original_period,
                document_date=opening.document_date,
                document_no=opening.document_no,
                sales_order_no=opening.sales_order_no,
                customer_code=opening.customer_code,
                customer_name=opening.customer_name,
                material_code=opening.material_code,
                material_name=opening.material_name,
                issue_quantity=issue_quantity,
                invoiced_quantity=invoiced_quantity,
                outstanding_quantity=outstanding_quantity,
                issue_amount=issue_amount,
                invoiced_cost=invoiced_cost,
                outstanding_cost=outstanding_cost,
                invoice_nos=invoice_nos,
                latest_invoice_period=totals["latest_period"],
                outstanding_days=max(
                    (_month_end(month) - opening.document_date).days,
                    0,
                ),
                cross_period=opening.original_period < month,
                status=status,
            )
        )
    rows.sort(
        key=lambda row: (
            row.issue_month,
            row.document_date,
            row.document_no,
        )
    )
    return rows


def build_sales_uninvoiced_rows(period, include_closed=False):
    return _build_sales_rows_as_of(period.month, include_closed=include_closed)


def summarize_sales_uninvoiced(rows, period):
    return {
        "issue_count": len(rows),
        "outstanding_quantity": sum(
            (row.outstanding_quantity for row in rows),
            ZERO,
        ),
        "outstanding_cost": sum(
            (
                row.outstanding_cost
                for row in rows
                if row.outstanding_cost is not None
            ),
            ZERO,
        ),
        "partial_count": sum(1 for row in rows if row.invoiced_quantity > ZERO),
        "prior_period_count": sum(
            1 for row in rows if row.issue_month < period.month
        ),
        "unpriced_count": sum(
            1 for row in rows if row.outstanding_cost is None
        ),
    }


def build_goods_issued_rollforward(period):
    opening_rows = _build_sales_rows_as_of(
        _previous_month(period.month),
        include_closed=False,
    )
    current_rows = _build_sales_rows_as_of(
        period.month,
        include_closed=True,
    )
    opening_amount = money(
        sum(
            (
                row.outstanding_cost
                for row in opening_rows
                if row.outstanding_cost is not None
            ),
            ZERO,
        )
        + sum(
            (
                item.opening_amount
                for item in OpeningGoodsIssued.objects.filter(period=period)
                .exclude(match_mode=MatchingMode.IGNORED)
            ),
            ZERO,
        )
    )
    current_issue_amount = money(
        sum(
            (
                row.issue_amount
                for row in current_rows
                if row.transaction is not None
                and row.issue_month == period.month
                and row.issue_amount is not None
            ),
            ZERO,
        )
    )
    ending_amount = money(
        sum(
            (
                row.outstanding_cost
                for row in current_rows
                if row.outstanding_cost is not None
            ),
            ZERO,
        )
    )
    invoice_cleared_amount = money(
        opening_amount + current_issue_amount - ending_amount
    )
    difference = money(
        opening_amount
        + current_issue_amount
        - invoice_cleared_amount
        - ending_amount
    )
    return {
        "opening_amount": opening_amount,
        "current_issue_amount": current_issue_amount,
        "invoice_cleared_amount": invoice_cleared_amount,
        "ending_amount": ending_amount,
        "difference": difference,
        "opening_count": len(opening_rows),
        "ending_count": sum(
            1 for row in current_rows if row.outstanding_quantity > ZERO
        ),
        "prior_period_count": sum(
            1
            for row in current_rows
            if row.outstanding_quantity > ZERO
            and row.issue_month < period.month
        ),
        "unpriced_count": sum(
            1
            for row in current_rows
            if row.outstanding_quantity > ZERO
            and row.outstanding_cost is None
        ),
    }


def summarize_current_sales_cost(period):
    rows_by_transaction = {
        row.transaction.id: row
        for row in _build_sales_rows_as_of(
            period.month,
            include_closed=True,
        )
        if row.transaction is not None
        and row.issue_period.id == period.id
    }
    totals = {
        "sales_cost": ZERO,
        "invoiced_cost": ZERO,
        "goods_issued": ZERO,
        "sample_cost": ZERO,
        "ignored_cost": ZERO,
    }
    for txn in period.inventory_transactions.all():
        if not is_sales_issue(txn):
            continue
        amount = money(abs(Decimal(txn.amount or ZERO)))
        totals["sales_cost"] += amount
        if _is_sample_issue(txn):
            totals["sample_cost"] += amount
        elif txn.match_mode == MatchingMode.IGNORED:
            totals["ignored_cost"] += amount
        else:
            row = rows_by_transaction.get(txn.id)
            if row is None:
                totals["goods_issued"] += amount
                continue
            totals["invoiced_cost"] += money(row.invoiced_cost)
            totals["goods_issued"] += money(row.outstanding_cost)
    return {key: money(value) for key, value in totals.items()}
