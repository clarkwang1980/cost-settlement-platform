from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP

from django.db.models import Max, Q, Sum

from settlement.models import (
    InventoryTransaction,
    MatchAllocation,
    MatchingMode,
    OpeningPurchaseAccrual,
)
from settlement.services.matching import is_purchase_receipt


ZERO = Decimal("0")
CENT = Decimal("0.01")
PRICE_PRECISION = Decimal("0.00000001")


def money(value) -> Decimal:
    return Decimal(value or ZERO).quantize(CENT, rounding=ROUND_HALF_UP)


def price(value) -> Decimal:
    return Decimal(value or ZERO).quantize(
        PRICE_PRECISION, rounding=ROUND_HALF_UP
    )


def provisional_basis(txn: InventoryTransaction):
    contract_allocations = txn.allocations.filter(
        kind=MatchAllocation.Kind.PURCHASE_CONTRACT
    )
    totals = contract_allocations.aggregate(
        quantity=Sum("matched_quantity"),
        amount=Sum("amount"),
    )
    contract_quantity = Decimal(totals["quantity"] or ZERO)
    contract_amount = Decimal(totals["amount"] or ZERO)
    if contract_quantity > ZERO:
        unit_price = price(contract_amount / contract_quantity)
        return {
            "unit_price": unit_price,
            "amount": money(Decimal(txn.quantity) * unit_price),
            "source": "合同暂估",
            "contract_nos": "；".join(
                contract_allocations.order_by("id")
                .values_list("contract_line__contract_no", flat=True)
                .distinct()
            ),
        }
    if (
        txn.unit_price is not None
        and txn.price_source in {"原始金额", "手工暂估"}
    ):
        unit_price = price(txn.unit_price)
        return {
            "unit_price": unit_price,
            "amount": money(Decimal(txn.quantity) * unit_price),
            "source": txn.price_source,
            "contract_nos": "",
        }
    return {
        "unit_price": None,
        "amount": None,
        "source": "缺少暂估价格",
        "contract_nos": "",
    }


@dataclass
class PurchaseAccrualRow:
    transaction: InventoryTransaction | None
    opening_item: OpeningPurchaseAccrual | None
    receipt_period: object | None
    receipt_month: object
    document_date: object
    document_no: str
    supplier_code: str
    supplier_name: str
    material_code: str
    material_name: str
    contract_nos: str
    receipt_quantity: Decimal
    invoiced_quantity: Decimal
    outstanding_quantity: Decimal
    provisional_unit_price: Decimal | None
    provisional_amount: Decimal | None
    price_source: str
    latest_invoice_period: object | None
    status: str


@dataclass
class PurchaseDifferenceRow:
    allocation: MatchAllocation
    transaction: InventoryTransaction | None
    opening_item: OpeningPurchaseAccrual | None
    invoice: object
    receipt_period: object | None
    receipt_month: object
    document_date: object
    document_no: str
    material_code: str
    material_name: str
    matched_quantity: Decimal
    provisional_unit_price: Decimal | None
    invoice_unit_price: Decimal
    provisional_amount: Decimal | None
    invoice_amount: Decimal
    difference_amount: Decimal | None
    cross_period: bool
    status: str


def _purchase_receipts_as_of(period):
    queryset = (
        InventoryTransaction.objects.filter(
            period__month__lte=period.month,
        )
        .exclude(match_mode=MatchingMode.IGNORED)
        .select_related("period")
        .prefetch_related(
            "allocations__contract_line",
            "allocations__purchase_invoice_line",
        )
        .order_by("period__month", "txn_date", "document_no", "id")
    )
    return [txn for txn in queryset if is_purchase_receipt(txn)]


def build_purchase_accrual_rows(period, include_closed=False):
    rows = []
    for txn in _purchase_receipts_as_of(period):
        invoice_allocations = txn.allocations.filter(
            kind=MatchAllocation.Kind.PURCHASE_INVOICE,
            period__month__lte=period.month,
        )
        invoice_totals = invoice_allocations.aggregate(
            quantity=Sum("matched_quantity"),
            latest_period=Max("purchase_invoice_line__invoice_period"),
        )
        receipt_quantity = abs(Decimal(txn.quantity or ZERO))
        invoiced_quantity = min(
            Decimal(invoice_totals["quantity"] or ZERO),
            receipt_quantity,
        )
        outstanding_quantity = max(
            receipt_quantity - invoiced_quantity,
            ZERO,
        )
        if outstanding_quantity == ZERO and not include_closed:
            continue
        basis = provisional_basis(txn)
        outstanding_amount = (
            money(outstanding_quantity * basis["unit_price"])
            if basis["unit_price"] is not None
            else None
        )
        if outstanding_quantity == ZERO:
            status = "全部到票"
        elif basis["unit_price"] is None:
            status = "缺少暂估价格"
        elif invoiced_quantity > ZERO:
            status = "部分到票"
        else:
            status = "全部暂估"
        rows.append(
            PurchaseAccrualRow(
                transaction=txn,
                opening_item=None,
                receipt_period=txn.period,
                receipt_month=txn.period.month,
                document_date=txn.txn_date,
                document_no=txn.document_no,
                supplier_code=txn.supplier_code,
                supplier_name=txn.supplier_name,
                material_code=txn.material_code,
                material_name=txn.material_name,
                contract_nos=basis["contract_nos"],
                receipt_quantity=receipt_quantity,
                invoiced_quantity=invoiced_quantity,
                outstanding_quantity=outstanding_quantity,
                provisional_unit_price=basis["unit_price"],
                provisional_amount=outstanding_amount,
                price_source=basis["source"],
                latest_invoice_period=invoice_totals["latest_period"],
                status=status,
            )
        )
    opening_rows = (
        OpeningPurchaseAccrual.objects.filter(
            period__month__lte=period.month,
        )
        .exclude(match_mode=MatchingMode.IGNORED)
        .prefetch_related("allocations__purchase_invoice_line")
        .order_by(
            "original_period",
            "document_date",
            "document_no",
            "line_no",
            "id",
        )
    )
    for opening in opening_rows:
        invoice_allocations = opening.allocations.filter(
            kind=MatchAllocation.Kind.PURCHASE_INVOICE,
            period__month__lte=period.month,
        ).exclude(
            purchase_invoice_line__match_mode=MatchingMode.IGNORED
        )
        totals = invoice_allocations.aggregate(
            quantity=Sum("matched_quantity"),
            latest_period=Max("purchase_invoice_line__invoice_period"),
        )
        receipt_quantity = Decimal(opening.opening_quantity or ZERO)
        invoiced_quantity = min(
            Decimal(totals["quantity"] or ZERO),
            receipt_quantity,
        )
        outstanding_quantity = max(
            receipt_quantity - invoiced_quantity,
            ZERO,
        )
        if outstanding_quantity == ZERO and not include_closed:
            continue
        unit_price = price(opening.provisional_unit_price)
        outstanding_amount = money(
            Decimal(opening.opening_amount or ZERO)
            * outstanding_quantity
            / receipt_quantity
        )
        status = (
            "全部到票"
            if outstanding_quantity == ZERO
            else "部分到票"
            if invoiced_quantity > ZERO
            else "全部暂估"
        )
        rows.append(
            PurchaseAccrualRow(
                transaction=None,
                opening_item=opening,
                receipt_period=None,
                receipt_month=opening.original_period,
                document_date=opening.document_date,
                document_no=opening.document_no,
                supplier_code=opening.supplier_code,
                supplier_name=opening.supplier_name,
                material_code=opening.material_code,
                material_name=opening.material_name,
                contract_nos=opening.contract_no,
                receipt_quantity=receipt_quantity,
                invoiced_quantity=invoiced_quantity,
                outstanding_quantity=outstanding_quantity,
                provisional_unit_price=unit_price,
                provisional_amount=outstanding_amount,
                price_source="期初暂估台账",
                latest_invoice_period=totals["latest_period"],
                status=status,
            )
        )
    rows.sort(
        key=lambda row: (
            row.receipt_month,
            row.document_date,
            row.document_no,
        )
    )
    return rows


def summarize_purchase_accruals(rows, period):
    return {
        "receipt_count": len(rows),
        "outstanding_quantity": sum(
            (row.outstanding_quantity for row in rows), ZERO
        ),
        "provisional_amount": sum(
            (
                row.provisional_amount
                for row in rows
                if row.provisional_amount is not None
            ),
            ZERO,
        ),
        "unpriced_count": sum(
            1 for row in rows if row.provisional_unit_price is None
        ),
        "prior_period_count": sum(
            1
            for row in rows
            if row.receipt_month < period.month
        ),
    }


def build_purchase_difference_rows(period):
    allocations = (
        MatchAllocation.objects.filter(
            period=period,
            kind=MatchAllocation.Kind.PURCHASE_INVOICE,
        )
        .exclude(
            purchase_invoice_line__match_mode=MatchingMode.IGNORED
        )
        .exclude(
            Q(inventory_transaction__match_mode=MatchingMode.IGNORED)
            | Q(opening_purchase_accrual__match_mode=MatchingMode.IGNORED)
        )
        .select_related(
            "inventory_transaction__period",
            "opening_purchase_accrual",
            "purchase_invoice_line",
        )
        .order_by("purchase_invoice_line__invoice_no", "id")
    )
    rows = []
    for allocation in allocations:
        txn = allocation.inventory_transaction
        opening = allocation.opening_purchase_accrual
        invoice = allocation.purchase_invoice_line
        if txn is not None:
            basis = provisional_basis(txn)
            receipt_period = txn.period
            receipt_month = txn.period.month
            document_date = txn.txn_date
            document_no = txn.document_no
            material_code = txn.material_code
            material_name = txn.material_name
        else:
            basis = {
                "unit_price": price(opening.provisional_unit_price),
                "amount": money(opening.opening_amount),
                "source": "期初暂估台账",
                "contract_nos": opening.contract_no,
            }
            receipt_period = None
            receipt_month = opening.original_period
            document_date = opening.document_date
            document_no = opening.document_no
            material_code = opening.material_code
            material_name = opening.material_name
        matched_quantity = Decimal(allocation.matched_quantity or ZERO)
        invoice_unit_price = price(
            allocation.unit_price or invoice.unit_price_excl_tax
        )
        invoice_amount = money(
            allocation.amount
            if allocation.amount is not None
            else matched_quantity * invoice_unit_price
        )
        provisional_amount = (
            money(matched_quantity * basis["unit_price"])
            if basis["unit_price"] is not None
            else None
        )
        difference_amount = (
            money(invoice_amount - provisional_amount)
            if provisional_amount is not None
            else None
        )
        rows.append(
            PurchaseDifferenceRow(
                allocation=allocation,
                transaction=txn,
                opening_item=opening,
                invoice=invoice,
                receipt_period=receipt_period,
                receipt_month=receipt_month,
                document_date=document_date,
                document_no=document_no,
                material_code=material_code,
                material_name=material_name,
                matched_quantity=matched_quantity,
                provisional_unit_price=basis["unit_price"],
                invoice_unit_price=invoice_unit_price,
                provisional_amount=provisional_amount,
                invoice_amount=invoice_amount,
                difference_amount=difference_amount,
                cross_period=receipt_month < period.month,
                status=(
                    "待确认暂估价格"
                    if difference_amount is None
                    else "有差异"
                    if difference_amount != ZERO
                    else "无差异"
                ),
            )
        )
    return rows


def summarize_purchase_differences(rows):
    totals = {
        "matched_quantity": ZERO,
        "provisional_amount": ZERO,
        "invoice_amount": ZERO,
        "difference_amount": ZERO,
        "pending_count": 0,
        "cross_period_count": 0,
    }
    grouped = defaultdict(
        lambda: {
            "supplier": "",
            "material_code": "",
            "material_name": "",
            "matched_quantity": ZERO,
            "provisional_amount": ZERO,
            "invoice_amount": ZERO,
            "difference_amount": ZERO,
            "pending_count": 0,
        }
    )
    for row in rows:
        totals["matched_quantity"] += row.matched_quantity
        totals["invoice_amount"] += row.invoice_amount
        totals["cross_period_count"] += int(row.cross_period)
        key = (
            row.invoice.supplier_code or row.invoice.supplier_name,
            row.material_code,
        )
        item = grouped[key]
        item["supplier"] = (
            row.invoice.supplier_name or row.invoice.supplier_code
        )
        item["material_code"] = row.material_code
        item["material_name"] = row.material_name
        item["matched_quantity"] += row.matched_quantity
        item["invoice_amount"] += row.invoice_amount
        if row.provisional_amount is None:
            totals["pending_count"] += 1
            item["pending_count"] += 1
        else:
            totals["provisional_amount"] += row.provisional_amount
            totals["difference_amount"] += row.difference_amount
            item["provisional_amount"] += row.provisional_amount
            item["difference_amount"] += row.difference_amount
    grouped_rows = sorted(
        grouped.values(),
        key=lambda item: (
            item["supplier"],
            item["material_code"],
        ),
    )
    return totals, grouped_rows


def build_purchase_reversal_rows(period, invoice_no="", supplier="", material="", cross_period=""):
    """发票暂估回冲明细：每张采购发票核销了哪笔入库/期初暂估，一行一条。

    基于 build_purchase_difference_rows 的行数据，补充：
    该入库单入库数量、该入库单累计已核销数量、核销后剩余未到票。
    """
    rows = build_purchase_difference_rows(period)

    # 历史累计已核销数量（截至所选期间，按入库单/期初台账聚合）
    cumulative = defaultdict(lambda: ZERO)
    history = (
        MatchAllocation.objects.filter(
            period__month__lte=period.month,
            kind=MatchAllocation.Kind.PURCHASE_INVOICE,
        )
        .exclude(purchase_invoice_line__match_mode=MatchingMode.IGNORED)
        .exclude(
            Q(inventory_transaction__match_mode=MatchingMode.IGNORED)
            | Q(opening_purchase_accrual__match_mode=MatchingMode.IGNORED)
        )
        .values_list(
            "inventory_transaction_id",
            "opening_purchase_accrual_id",
            "matched_quantity",
        )
    )
    for txn_id, opening_id, matched in history:
        if txn_id is not None:
            cumulative[("t", txn_id)] += Decimal(matched or ZERO)
        elif opening_id is not None:
            cumulative[("o", opening_id)] += Decimal(matched or ZERO)

    result = []
    for row in rows:
        if invoice_no and invoice_no.lower() not in str(row.invoice.invoice_no or "").lower():
            continue
        if supplier and supplier.lower() not in str(
            row.invoice.supplier_name or row.invoice.supplier_code or ""
        ).lower():
            continue
        if material and material.lower() not in str(row.material_code or "").lower():
            continue
        if cross_period and str(row.cross_period).lower() != cross_period.lower():
            continue

        if row.transaction is not None:
            target_type = "库存流水入库"
            receipt_quantity = abs(Decimal(row.transaction.quantity or ZERO))
            cum_key = ("t", row.transaction.id)
        else:
            target_type = "期初暂估台账"
            receipt_quantity = Decimal(
                row.opening_item.opening_quantity or ZERO
            )
            cum_key = ("o", row.opening_item.id)
        cum_quantity = min(cumulative.get(cum_key, ZERO), receipt_quantity)
        outstanding = max(receipt_quantity - cum_quantity, ZERO)

        invoice_period = getattr(row.invoice, "invoice_period", None)
        result.append(
            {
                "invoice_no": row.invoice.invoice_no,
                "invoice_date": row.invoice.invoice_date,
                "invoice_period": invoice_period,
                "target_type": target_type,
                "document_no": row.document_no,
                "receipt_month": row.receipt_month,
                "document_date": row.document_date,
                "material_code": row.material_code,
                "material_name": row.material_name,
                "supplier_name": (
                    row.invoice.supplier_name or row.invoice.supplier_code
                ),
                "matched_quantity": row.matched_quantity,
                "provisional_unit_price": row.provisional_unit_price,
                "provisional_amount": row.provisional_amount,
                "invoice_unit_price": row.invoice_unit_price,
                "invoice_amount": row.invoice_amount,
                "difference_amount": row.difference_amount,
                "receipt_quantity": receipt_quantity,
                "cumulative_quantity": cum_quantity,
                "outstanding_quantity": outstanding,
                "cross_period": row.cross_period,
                "status": row.status,
            }
        )
    result.sort(
        key=lambda r: (
            str(r["invoice_no"] or ""),
            str(r["document_no"] or ""),
        )
    )
    return result


def summarize_purchase_reversals(rows):
    totals = {
        "count": len(rows),
        "matched_quantity": ZERO,
        "provisional_amount": ZERO,
        "invoice_amount": ZERO,
        "difference_amount": ZERO,
        "cross_period_count": 0,
        "pending_count": 0,
    }
    for row in rows:
        totals["matched_quantity"] += row["matched_quantity"]
        totals["cross_period_count"] += int(row["cross_period"])
        if row["provisional_amount"] is None:
            totals["pending_count"] += 1
        else:
            totals["provisional_amount"] += row["provisional_amount"]
            totals["invoice_amount"] += row["invoice_amount"]
            totals["difference_amount"] += row["difference_amount"]
    return totals
