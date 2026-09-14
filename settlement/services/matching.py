from __future__ import annotations

from decimal import Decimal

from django.db import transaction
from django.db.models import Max, Sum

from settlement.models import (
    AccountingPeriod,
    InventoryTransaction,
    MatchAllocation,
    MatchingMode,
    OpeningGoodsIssued,
    OpeningPurchaseAccrual,
    PurchaseContractLine,
    PurchaseInvoiceLine,
    SalesInvoiceLine,
)


ZERO = Decimal("0")


def _norm(value: str) -> str:
    return "".join(str(value or "").split()).lower()


def _same_party(code_a, name_a, code_b, name_b) -> bool:
    """供应商匹配：编码相等，或名称归一化后相等/互相包含（简称对全称）。

    任一方完全没有供应商信息时返回 False，避免无供应商的流水/发票被盲目匹配。
    """
    if code_a and code_b:
        return _norm(code_a) == _norm(code_b)
    if name_a and name_b:
        name_a_norm = _norm(name_a)
        name_b_norm = _norm(name_b)
        return (
            name_a_norm == name_b_norm
            or name_a_norm in name_b_norm
            or name_b_norm in name_a_norm
        )
    return False


def _same_party_strict(code_a, name_a, code_b, name_b) -> bool:
    if code_a and code_b:
        return _norm(code_a) == _norm(code_b)
    if name_a and name_b:
        return _norm(name_a) == _norm(name_b)
    return False


def is_production_receipt(txn: InventoryTransaction) -> bool:
    """完工入库（报工入库）：业务类型为"报工"的入库。"""
    return (
        txn.direction == InventoryTransaction.Direction.IN
        and (txn.business_type or "").strip() == "报工"
    )


def is_purchase_receipt(txn: InventoryTransaction) -> bool:
    kind = _norm(txn.business_type)
    if txn.direction != InventoryTransaction.Direction.IN:
        return False
    if is_production_receipt(txn) or "退料" in kind:
        return False
    return "采购" in kind or bool(txn.supplier_code or txn.supplier_name)


def is_sales_issue(txn: InventoryTransaction) -> bool:
    kind = _norm(txn.business_type)
    if txn.direction != InventoryTransaction.Direction.OUT:
        return False
    return any(token in kind for token in ("销售", "发货", "样品")) or bool(
        txn.customer_code or txn.customer_name
    )


def is_sample_issue(txn) -> bool:
    """样品（样品免费赠送）出库判定，满足其一即为样品：

    1) 业务类型或备注含“样品”；
    2) 该行被平台生成的样品开票记录（发票号 LEDGER-SAMPLE-*）核销。

    样品出库照常参与库存计价（数量扣减、成本=月末加权价），但不参与正常销售核销，
    结算时计入销售费用（sample_cost），不进入“已开票结转/发出商品”口径。"""
    if "样品" in f"{txn.business_type or ''}{txn.remark or ''}":
        return True
    try:
        return txn.allocations.filter(
            kind=MatchAllocation.Kind.SALES_INVOICE,
            sales_invoice_line__invoice_no__startswith="LEDGER-SAMPLE-",
        ).exists()
    except Exception:
        return False


def _same_party_loose(code_a, name_a, code_b, name_b) -> bool:
    """宽松客户匹配：任一方完全没有客户信息时视为可匹配；双方都有客户时才要求严格一致。

    用于“无订单号”的开票记录（如平台台账当月结转生成的 LEDGER 发票）按物料核销当月流水，
    因为流水销售出库行通常不带客户/订单信息。"""
    has_a = bool(code_a or name_a)
    has_b = bool(code_b or name_b)
    if not has_a or not has_b:
        return True
    return _same_party_strict(code_a, name_a, code_b, name_b)


def _manual_used(source, field_name: str, kind: str) -> Decimal:
    return (
        MatchAllocation.objects.filter(
            kind=kind, is_manual=True, **{field_name: source}
        ).aggregate(total=Sum("matched_quantity"))["total"]
        or ZERO
    )


def _reset_remaining(period: AccountingPeriod):
    for contract in period.purchase_contracts.all():
        contract.remaining_quantity = max(
            contract.quantity
            - _manual_used(
                contract,
                "contract_line",
                MatchAllocation.Kind.PURCHASE_CONTRACT,
            ),
            ZERO,
        )
        contract.save(update_fields=["remaining_quantity"])
    for invoice in period.purchase_invoices.all():
        invoice.remaining_quantity = max(
            invoice.quantity
            - _manual_used(
                invoice,
                "purchase_invoice_line",
                MatchAllocation.Kind.PURCHASE_INVOICE,
            ),
            ZERO,
        )
        invoice.save(update_fields=["remaining_quantity"])
    for invoice in period.sales_invoices.all():
        invoice.remaining_quantity = max(
            invoice.quantity
            - _manual_used(
                invoice,
                "sales_invoice_line",
                MatchAllocation.Kind.SALES_INVOICE,
            ),
            ZERO,
        )
        invoice.save(update_fields=["remaining_quantity"])
    for opening in OpeningPurchaseAccrual.objects.filter(
        period__month__lte=period.month
    ):
        used = (
            opening.allocations.filter(
                kind=MatchAllocation.Kind.PURCHASE_INVOICE,
                period__month__lte=period.month,
            ).aggregate(total=Sum("matched_quantity"))["total"]
            or ZERO
        )
        opening.remaining_quantity = max(opening.opening_quantity - used, ZERO)
        opening.save(update_fields=["remaining_quantity"])
    for opening in OpeningGoodsIssued.objects.filter(
        period__month__lte=period.month
    ):
        used = (
            opening.allocations.filter(
                kind=MatchAllocation.Kind.SALES_INVOICE,
                period__month__lte=period.month,
            ).aggregate(total=Sum("matched_quantity"))["total"]
            or ZERO
        )
        opening.remaining_quantity = max(opening.opening_quantity - used, ZERO)
        opening.save(update_fields=["remaining_quantity"])


def _txn_matched_quantity(txn, kind):
    return (
        txn.allocations.filter(kind=kind).aggregate(total=Sum("matched_quantity"))[
            "total"
        ]
        or ZERO
    )


def _candidate_contracts(period, txn):
    contracts = list(
        period.purchase_contracts.filter(
            material_code=txn.material_code, remaining_quantity__gt=0
        )
    )
    if txn.supplier_code or txn.supplier_name:
        contracts = [
            item
            for item in contracts
            if _same_party(
                txn.supplier_code,
                txn.supplier_name,
                item.supplier_code,
                item.supplier_name,
            )
        ]
    else:
        # 流水无供应商：仅当该物料所有剩余合同的供应商唯一时才自动匹配，
        # 避免 R82.101（多供应商）这类流水被错误匹配到任意合同。
        suppliers = {
            (item.supplier_code or "", item.supplier_name or "")
            for item in contracts
        }
        if len(suppliers) != 1:
            return []
    # 剩余待匹配数量：一条流水可能已部分匹配到合同（人工或历史分配）
    remaining = max(
        txn.quantity
        - _txn_matched_quantity(txn, MatchAllocation.Kind.PURCHASE_CONTRACT),
        ZERO,
    )
    own_contract_no = txn.contract_no or ""

    def _contract_sort_key(item):
        if own_contract_no and item.contract_no == own_contract_no:
            # 流水自带合同号：优先匹配该合同（原逻辑）
            return (
                0,
                0,
                -(item.contract_date or period.month).toordinal(),
                -item.id,
            )
        if item.remaining_quantity <= remaining:
            # 能整块吃下：剩余数量大的先，减少拆单混价（如 R11.504 两价合同）
            return (
                1,
                -item.remaining_quantity,
                -(item.contract_date or period.month).toordinal(),
                -item.id,
            )
        # 不能整块吃下：剩余数量最小的先（最小拆分）
        return (
            2,
            item.remaining_quantity,
            -(item.contract_date or period.month).toordinal(),
            -item.id,
        )

    contracts.sort(key=_contract_sort_key)
    return contracts


def _candidate_purchase_receipts(period, invoice):
    txns = [
        txn
        for txn in InventoryTransaction.objects.filter(
            period__month__lte=period.month,
            match_mode=MatchingMode.AUTO
        ).select_related("period")
        if is_purchase_receipt(txn)
    ]
    txns = [
        txn
        for txn in txns
        if txn.material_code == invoice.material_code
        and _same_party(
            txn.supplier_code,
            txn.supplier_name,
            invoice.supplier_code,
            invoice.supplier_name,
        )
    ]
    def _receipt_sort_key(txn):
        # 流水自带合同号与发票合同号一致优先（原逻辑）
        if invoice.contract_no and (
            txn.contract_no == invoice.contract_no
            or txn.allocations.filter(
                kind=MatchAllocation.Kind.PURCHASE_CONTRACT,
                contract_line__contract_no=invoice.contract_no,
            ).exists()
        ):
            contract_bonus = 0
        else:
            contract_bonus = 1
        # 流水暂估单价与发票单价差：越小越优先，避免不同价发票核销到错价流水产生假差异
        if txn.unit_price is not None and invoice.unit_price_excl_tax is not None:
            price_diff = abs(txn.unit_price - invoice.unit_price_excl_tax)
        else:
            price_diff = Decimal("1e9")
        return (
            contract_bonus,
            price_diff,
            txn.period.month,
            txn.txn_date,
            txn.id,
        )

    txns.sort(key=_receipt_sort_key)
    return txns


def _candidate_opening_purchase_accruals(period, invoice):
    rows = list(
        OpeningPurchaseAccrual.objects.filter(
            period__month__lte=period.month,
            material_code=invoice.material_code,
            remaining_quantity__gt=0,
            match_mode=MatchingMode.AUTO,
        )
    )
    rows = [
        item
        for item in rows
        if _same_party(
            item.supplier_code,
            item.supplier_name,
            invoice.supplier_code,
            invoice.supplier_name,
        )
    ]
    rows.sort(
        key=lambda item: (
            0
            if invoice.contract_no and item.contract_no == invoice.contract_no
            else 1,
            item.original_period,
            item.document_date,
            item.id,
        )
    )
    return rows


def _candidate_sales_issues(period, invoice):
    txns = [
        txn
        for txn in InventoryTransaction.objects.filter(
            period__month__lte=period.month,
            match_mode=MatchingMode.AUTO
        ).select_related("period")
        if is_sales_issue(txn) and not is_sample_issue(txn)
    ]
    if invoice.sales_order_no:
        txns = [
            txn
            for txn in txns
            if txn.material_code == invoice.material_code
            and txn.order_no
            and _norm(txn.order_no) == _norm(invoice.sales_order_no)
            and _same_party_strict(
                txn.customer_code,
                txn.customer_name,
                invoice.customer_code,
                invoice.customer_name,
            )
        ]
    elif invoice.invoice_no.startswith("LEDGER"):
        # 无订单号：只有平台台账生成的 LEDGER 发票才按物料宽松核销（它们本就没有订单/客户）；
        # 你上传的销售发票若缺订单号，不做自动核销，留在"销售匹配"页人工核销，避免张冠李戴。
        txns = [
            txn
            for txn in txns
            if txn.material_code == invoice.material_code
            and _same_party_loose(
                txn.customer_code,
                txn.customer_name,
                invoice.customer_code,
                invoice.customer_name,
            )
        ]
    else:
        # 用户上传发票缺订单号：不自动核销（需人工指定对应出库/期初发出商品）
        txns = []
    txns.sort(
        key=lambda txn: (
            txn.period.month,
            txn.txn_date,
            txn.id,
        )
    )
    return txns


def _candidate_opening_goods_issued(period, invoice):
    if not invoice.sales_order_no:
        return []
    rows = list(
        OpeningGoodsIssued.objects.filter(
            period__month__lte=period.month,
            material_code=invoice.material_code,
            remaining_quantity__gt=0,
            match_mode=MatchingMode.AUTO,
        )
    )
    rows = [
        item
        for item in rows
        if item.sales_order_no
        and _norm(item.sales_order_no) == _norm(invoice.sales_order_no)
        and _same_party_strict(
            item.customer_code,
            item.customer_name,
            invoice.customer_code,
            invoice.customer_name,
        )
    ]
    rows.sort(
        key=lambda item: (
            item.original_period,
            item.document_date,
            item.id,
        )
    )
    return rows


def _allocate_contracts(period):
    for txn in period.inventory_transactions.all():
        if (
            not is_purchase_receipt(txn)
            or txn.match_mode != MatchingMode.AUTO
        ):
            continue
        remaining = max(
            txn.quantity
            - _txn_matched_quantity(txn, MatchAllocation.Kind.PURCHASE_CONTRACT),
            ZERO,
        )
        while remaining > 0:
            candidates = _candidate_contracts(period, txn)
            if not candidates:
                break
            contract = candidates[0]
            qty = min(remaining, contract.remaining_quantity)
            if qty <= 0:
                break
            MatchAllocation.objects.create(
                period=period,
                kind=MatchAllocation.Kind.PURCHASE_CONTRACT,
                inventory_transaction=txn,
                contract_line=contract,
                matched_quantity=qty,
                unit_price=contract.unit_price_excl_tax,
                amount=(qty * contract.unit_price_excl_tax).quantize(
                    Decimal("0.01")
                ),
                confidence=100
                if txn.contract_no and txn.contract_no == contract.contract_no
                else 85,
            )
            remaining -= qty
            contract.remaining_quantity -= qty
            contract.save(update_fields=["remaining_quantity"])


def _allocate_purchase_invoices(period):
    for invoice in period.purchase_invoices.filter(
        match_mode=MatchingMode.AUTO
    ):
        remaining_invoice = invoice.remaining_quantity
        for opening in _candidate_opening_purchase_accruals(period, invoice):
            if remaining_invoice <= 0:
                break
            qty = min(remaining_invoice, opening.remaining_quantity)
            if qty <= 0:
                continue
            MatchAllocation.objects.create(
                period=period,
                kind=MatchAllocation.Kind.PURCHASE_INVOICE,
                opening_purchase_accrual=opening,
                purchase_invoice_line=invoice,
                matched_quantity=qty,
                unit_price=invoice.unit_price_excl_tax,
                amount=(qty * invoice.unit_price_excl_tax).quantize(
                    Decimal("0.01")
                ),
                confidence=100
                if invoice.contract_no
                and opening.contract_no == invoice.contract_no
                else 85,
            )
            remaining_invoice -= qty
            opening.remaining_quantity -= qty
            opening.save(update_fields=["remaining_quantity"])
            invoice.remaining_quantity -= qty
            invoice.save(update_fields=["remaining_quantity"])
        for txn in _candidate_purchase_receipts(period, invoice):
            if remaining_invoice <= 0:
                break
            already = _txn_matched_quantity(
                txn, MatchAllocation.Kind.PURCHASE_INVOICE
            )
            remaining_txn = max(txn.quantity - already, ZERO)
            qty = min(remaining_invoice, remaining_txn)
            if qty <= 0:
                continue
            MatchAllocation.objects.create(
                period=period,
                kind=MatchAllocation.Kind.PURCHASE_INVOICE,
                inventory_transaction=txn,
                purchase_invoice_line=invoice,
                matched_quantity=qty,
                unit_price=invoice.unit_price_excl_tax,
                amount=(qty * invoice.unit_price_excl_tax).quantize(
                    Decimal("0.01")
                ),
                confidence=100
                if invoice.contract_no and txn.contract_no == invoice.contract_no
                else 85,
            )
            remaining_invoice -= qty
            invoice.remaining_quantity -= qty
            invoice.save(update_fields=["remaining_quantity"])


def _allocate_sales_invoices(period):
    for invoice in period.sales_invoices.filter(
        match_mode=MatchingMode.AUTO
    ):
        remaining_invoice = invoice.remaining_quantity
        for opening in _candidate_opening_goods_issued(period, invoice):
            if remaining_invoice <= 0:
                break
            qty = min(remaining_invoice, opening.remaining_quantity)
            if qty <= 0:
                continue
            MatchAllocation.objects.create(
                period=period,
                kind=MatchAllocation.Kind.SALES_INVOICE,
                opening_goods_issued=opening,
                sales_invoice_line=invoice,
                matched_quantity=qty,
                confidence=100,
            )
            remaining_invoice -= qty
            opening.remaining_quantity -= qty
            opening.save(update_fields=["remaining_quantity"])
            invoice.remaining_quantity -= qty
            invoice.save(update_fields=["remaining_quantity"])
        for txn in _candidate_sales_issues(period, invoice):
            if remaining_invoice <= 0:
                break
            already = _txn_matched_quantity(txn, MatchAllocation.Kind.SALES_INVOICE)
            remaining_txn = max(txn.quantity - already, ZERO)
            qty = min(remaining_invoice, remaining_txn)
            if qty <= 0:
                continue
            MatchAllocation.objects.create(
                period=period,
                kind=MatchAllocation.Kind.SALES_INVOICE,
                inventory_transaction=txn,
                sales_invoice_line=invoice,
                matched_quantity=qty,
                confidence=100,
            )
            remaining_invoice -= qty
            invoice.remaining_quantity -= qty
            invoice.save(update_fields=["remaining_quantity"])


def _refresh_purchase_transaction_prices(period):
    for txn in period.inventory_transactions.all():
        if not is_purchase_receipt(txn):
            continue
        if txn.match_mode == MatchingMode.IGNORED:
            txn.invoice_period = None
            txn.match_status = InventoryTransaction.MatchStatus.NOT_REQUIRED
            txn.price_source = "忽略匹配"
            txn.save(
                update_fields=[
                    "invoice_period",
                    "match_status",
                    "price_source",
                ]
            )
            continue
        invoice_allocations = txn.allocations.filter(
            kind=MatchAllocation.Kind.PURCHASE_INVOICE
        )
        contract_allocations = txn.allocations.filter(
            kind=MatchAllocation.Kind.PURCHASE_CONTRACT
        )
        invoice_qty = (
            invoice_allocations.aggregate(total=Sum("matched_quantity"))["total"]
            or ZERO
        )
        contract_qty = (
            contract_allocations.aggregate(total=Sum("matched_quantity"))["total"]
            or ZERO
        )
        contract_amount = (
            contract_allocations.aggregate(total=Sum("amount"))["total"] or ZERO
        )
        contract_price = contract_amount / contract_qty if contract_qty else None
        if contract_price is not None:
            txn.unit_price = contract_price.quantize(Decimal("0.00000001"))
            txn.amount = (txn.quantity * contract_price).quantize(
                Decimal("0.01")
            )
            if invoice_qty >= txn.quantity:
                txn.price_source = "合同暂估（已全部到票）"
            elif invoice_qty > 0:
                txn.price_source = "合同暂估（部分到票）"
            else:
                txn.price_source = "合同暂估"
        periods = list(
            invoice_allocations.values_list(
                "purchase_invoice_line__invoice_period", flat=True
            )
        )
        txn.invoice_period = max((item for item in periods if item), default=None)
        matched_contract_qty = contract_qty
        if matched_contract_qty >= txn.quantity:
            txn.match_status = InventoryTransaction.MatchStatus.MATCHED
        elif matched_contract_qty > 0:
            txn.match_status = InventoryTransaction.MatchStatus.PARTIAL
        else:
            txn.match_status = InventoryTransaction.MatchStatus.UNMATCHED
        txn.save(
            update_fields=[
                "unit_price",
                "amount",
                "price_source",
                "invoice_period",
                "match_status",
            ]
        )


def refresh_purchase_invoice_metadata(period):
    """Refresh receipt invoice periods without changing historical inventory values."""
    receipts = InventoryTransaction.objects.filter(
        period__month__lte=period.month
    ).select_related("period")
    for txn in receipts:
        if not is_purchase_receipt(txn):
            continue
        latest = txn.allocations.filter(
            kind=MatchAllocation.Kind.PURCHASE_INVOICE,
            period__month__lte=period.month,
        ).aggregate(
            latest=Max("purchase_invoice_line__invoice_period")
        )["latest"]
        if txn.invoice_period != latest:
            txn.invoice_period = latest
            txn.save(update_fields=["invoice_period"])
    for opening in OpeningPurchaseAccrual.objects.filter(
        period__month__lte=period.month
    ):
        if opening.match_mode == MatchingMode.IGNORED:
            opening.invoice_period = None
            opening.remaining_quantity = opening.opening_quantity
            opening.match_status = InventoryTransaction.MatchStatus.NOT_REQUIRED
            opening.save(
                update_fields=[
                    "invoice_period",
                    "remaining_quantity",
                    "match_status",
                ]
            )
            continue
        allocations = opening.allocations.filter(
            kind=MatchAllocation.Kind.PURCHASE_INVOICE,
            period__month__lte=period.month,
        )
        matched = (
            allocations.aggregate(total=Sum("matched_quantity"))["total"]
            or ZERO
        )
        opening.remaining_quantity = max(
            opening.opening_quantity - matched, ZERO
        )
        opening.invoice_period = allocations.aggregate(
            latest=Max("purchase_invoice_line__invoice_period")
        )["latest"]
        opening.match_status = (
            InventoryTransaction.MatchStatus.MATCHED
            if matched >= opening.opening_quantity
            else InventoryTransaction.MatchStatus.PARTIAL
            if matched > 0
            else InventoryTransaction.MatchStatus.UNMATCHED
        )
        opening.save(
            update_fields=[
                "invoice_period",
                "remaining_quantity",
                "match_status",
            ]
        )


def refresh_sales_invoice_metadata(period):
    """Refresh sales issue invoice metadata through the selected month."""
    transactions = InventoryTransaction.objects.filter(
        period__month__lte=period.month
    ).select_related("period")
    for txn in transactions:
        if not is_sales_issue(txn):
            continue
        if txn.match_mode == MatchingMode.IGNORED:
            txn.invoice_period = None
            txn.match_status = InventoryTransaction.MatchStatus.NOT_REQUIRED
            txn.save(update_fields=["invoice_period", "match_status"])
            continue
        allocations = txn.allocations.filter(
            kind=MatchAllocation.Kind.SALES_INVOICE,
            period__month__lte=period.month,
        )
        matched = allocations.aggregate(total=Sum("matched_quantity"))["total"] or ZERO
        periods = list(
            allocations.values_list("sales_invoice_line__invoice_period", flat=True)
        )
        txn.invoice_period = max((item for item in periods if item), default=None)
        if matched >= txn.quantity:
            txn.match_status = InventoryTransaction.MatchStatus.MATCHED
        elif matched > 0:
            txn.match_status = InventoryTransaction.MatchStatus.PARTIAL
        else:
            txn.match_status = InventoryTransaction.MatchStatus.UNMATCHED
        txn.save(update_fields=["invoice_period", "match_status"])
    for opening in OpeningGoodsIssued.objects.filter(
        period__month__lte=period.month
    ):
        if opening.match_mode == MatchingMode.IGNORED:
            opening.invoice_period = None
            opening.remaining_quantity = opening.opening_quantity
            opening.match_status = InventoryTransaction.MatchStatus.NOT_REQUIRED
            opening.save(
                update_fields=[
                    "invoice_period",
                    "remaining_quantity",
                    "match_status",
                ]
            )
            continue
        allocations = opening.allocations.filter(
            kind=MatchAllocation.Kind.SALES_INVOICE,
            period__month__lte=period.month,
        )
        matched = (
            allocations.aggregate(total=Sum("matched_quantity"))["total"]
            or ZERO
        )
        opening.remaining_quantity = max(
            opening.opening_quantity - matched, ZERO
        )
        opening.invoice_period = allocations.aggregate(
            latest=Max("sales_invoice_line__invoice_period")
        )["latest"]
        opening.match_status = (
            InventoryTransaction.MatchStatus.MATCHED
            if matched >= opening.opening_quantity
            else InventoryTransaction.MatchStatus.PARTIAL
            if matched > 0
            else InventoryTransaction.MatchStatus.UNMATCHED
        )
        opening.save(
            update_fields=[
                "invoice_period",
                "remaining_quantity",
                "match_status",
            ]
        )


@transaction.atomic
def auto_match_all(period: AccountingPeriod):
    if period.status == AccountingPeriod.Status.LOCKED:
        raise ValueError("该期间已锁账，不能重新匹配。")
    period.match_allocations.filter(is_manual=False).delete()
    _reset_remaining(period)
    _allocate_contracts(period)
    _allocate_purchase_invoices(period)
    _allocate_sales_invoices(period)
    _refresh_purchase_transaction_prices(period)
    refresh_purchase_invoice_metadata(period)
    refresh_sales_invoice_metadata(period)


@transaction.atomic
def manual_match(
    period,
    kind,
    transaction_id,
    source_id,
    quantity,
    set_source_manual=True,
):
    if period.status == AccountingPeriod.Status.LOCKED:
        raise ValueError("该期间已锁账，不能人工匹配。")
    if kind in {
        MatchAllocation.Kind.PURCHASE_INVOICE,
        MatchAllocation.Kind.SALES_INVOICE,
    }:
        txn = InventoryTransaction.objects.get(
            id=transaction_id,
            period__month__lte=period.month,
        )
        if (
            kind == MatchAllocation.Kind.PURCHASE_INVOICE
            and not is_purchase_receipt(txn)
        ):
            raise ValueError("采购发票只能匹配采购入库流水。")
        if (
            kind == MatchAllocation.Kind.SALES_INVOICE
            and not is_sales_issue(txn)
        ):
            raise ValueError("销售发票只能匹配销售出库流水。")
    else:
        txn = InventoryTransaction.objects.get(id=transaction_id, period=period)
    quantity = Decimal(quantity)
    if quantity <= 0:
        raise ValueError("匹配数量必须大于0。")
    fields = {
        MatchAllocation.Kind.PURCHASE_CONTRACT: (
            PurchaseContractLine,
            "contract_line",
        ),
        MatchAllocation.Kind.PURCHASE_INVOICE: (
            PurchaseInvoiceLine,
            "purchase_invoice_line",
        ),
        MatchAllocation.Kind.SALES_INVOICE: (
            SalesInvoiceLine,
            "sales_invoice_line",
        ),
    }
    model, field_name = fields[kind]
    source = model.objects.get(id=source_id, period=period)
    if hasattr(source, "match_mode") and source.match_mode == MatchingMode.IGNORED:
        raise ValueError("该单据已设置为忽略匹配，请先修改处理方式。")
    if set_source_manual and hasattr(source, "match_mode"):
        source.match_mode = MatchingMode.MANUAL
        source.ignore_reason = ""
        source.save(update_fields=["match_mode", "ignore_reason"])
    MatchAllocation.objects.filter(
        period=period,
        kind=kind,
        inventory_transaction=txn,
        is_manual=False,
    ).delete()
    used_on_txn = (
        MatchAllocation.objects.filter(
            kind=kind,
            inventory_transaction=txn,
        ).aggregate(total=Sum("matched_quantity"))["total"]
        or ZERO
    )
    if used_on_txn + quantity > txn.quantity:
        raise ValueError("人工匹配数量超过该库存流水的剩余数量。")
    source_capacity = getattr(source, "quantity", None)
    if source_capacity is not None:
        used_on_source = (
            MatchAllocation.objects.filter(
                kind=kind,
                **{field_name: source},
            ).aggregate(total=Sum("matched_quantity"))["total"]
            or ZERO
        )
        if used_on_source + quantity > source_capacity:
            raise ValueError("人工匹配数量超过合同、发票或订单的可匹配数量。")
    kwargs = {
        "period": period,
        "kind": kind,
        "inventory_transaction": txn,
        "matched_quantity": quantity,
        "is_manual": True,
        field_name: source,
    }
    if kind == MatchAllocation.Kind.PURCHASE_CONTRACT:
        kwargs["unit_price"] = source.unit_price_excl_tax
        kwargs["amount"] = (quantity * source.unit_price_excl_tax).quantize(
            Decimal("0.01")
        )
    elif kind == MatchAllocation.Kind.PURCHASE_INVOICE:
        kwargs["unit_price"] = source.unit_price_excl_tax
        kwargs["amount"] = (quantity * source.unit_price_excl_tax).quantize(
            Decimal("0.01")
        )
    MatchAllocation.objects.create(**kwargs)
    auto_match_all(period)


@transaction.atomic
def manual_match_opening(
    period,
    kind,
    opening_id,
    source_id,
    quantity,
    set_source_manual=True,
):
    if period.status == AccountingPeriod.Status.LOCKED:
        raise ValueError("该期间已锁账，不能人工匹配。")
    if kind == MatchAllocation.Kind.PURCHASE_INVOICE:
        opening_model = OpeningPurchaseAccrual
        opening_field = "opening_purchase_accrual"
        source_model = PurchaseInvoiceLine
        source_field = "purchase_invoice_line"
    elif kind == MatchAllocation.Kind.SALES_INVOICE:
        opening_model = OpeningGoodsIssued
        opening_field = "opening_goods_issued"
        source_model = SalesInvoiceLine
        source_field = "sales_invoice_line"
    else:
        raise ValueError("期初台账只支持采购发票或销售发票核销。")

    opening = opening_model.objects.get(
        id=opening_id,
        period__month__lte=period.month,
    )
    source = source_model.objects.get(id=source_id, period=period)
    if opening.match_mode == MatchingMode.IGNORED:
        raise ValueError("该期初明细已设置为忽略匹配。")
    if source.match_mode == MatchingMode.IGNORED:
        raise ValueError("该发票已设置为忽略匹配。")
    quantity = Decimal(quantity)
    if quantity <= 0:
        raise ValueError("匹配数量必须大于0。")

    if set_source_manual:
        source.match_mode = MatchingMode.MANUAL
        source.ignore_reason = ""
        source.save(update_fields=["match_mode", "ignore_reason"])
    MatchAllocation.objects.filter(
        period=period,
        kind=kind,
        is_manual=False,
        **{opening_field: opening},
    ).delete()
    MatchAllocation.objects.filter(
        period=period,
        kind=kind,
        is_manual=False,
        **{source_field: source},
    ).delete()
    used_on_opening = (
        MatchAllocation.objects.filter(
            kind=kind,
            period__month__lte=period.month,
            **{opening_field: opening},
        ).aggregate(total=Sum("matched_quantity"))["total"]
        or ZERO
    )
    if used_on_opening + quantity > opening.opening_quantity:
        raise ValueError("人工匹配数量超过该期初明细的剩余数量。")
    used_on_source = (
        MatchAllocation.objects.filter(
            kind=kind,
            **{source_field: source},
        ).aggregate(total=Sum("matched_quantity"))["total"]
        or ZERO
    )
    if used_on_source + quantity > source.quantity:
        raise ValueError("人工匹配数量超过该发票的剩余数量。")

    kwargs = {
        "period": period,
        "kind": kind,
        opening_field: opening,
        source_field: source,
        "matched_quantity": quantity,
        "is_manual": True,
    }
    if kind == MatchAllocation.Kind.PURCHASE_INVOICE:
        kwargs["unit_price"] = source.unit_price_excl_tax
        kwargs["amount"] = (
            quantity * source.unit_price_excl_tax
        ).quantize(Decimal("0.01"))
    MatchAllocation.objects.create(**kwargs)
    auto_match_all(period)


def _source_allocation_filter(source):
    if isinstance(source, PurchaseInvoiceLine):
        return {"purchase_invoice_line": source}
    if isinstance(source, SalesInvoiceLine):
        return {"sales_invoice_line": source}
    raise ValueError("不支持的单据类型。")


@transaction.atomic
def set_source_match_mode(period, source, mode, ignore_reason=""):
    if period.status == AccountingPeriod.Status.LOCKED:
        raise ValueError("该期间已锁账，不能修改匹配方式。")
    if mode not in MatchingMode.values:
        raise ValueError("不支持的匹配方式。")
    if source.period_id != period.id:
        raise ValueError("单据不属于当前核算期间。")
    allocation_filter = _source_allocation_filter(source)
    allocations = MatchAllocation.objects.filter(
        period=period, **allocation_filter
    )
    if mode in {MatchingMode.AUTO, MatchingMode.IGNORED}:
        allocations.delete()
    else:
        allocations.filter(is_manual=False).delete()
    source.match_mode = mode
    source.ignore_reason = (
        str(ignore_reason or "").strip()
        if mode == MatchingMode.IGNORED
        else ""
    )
    source.save(update_fields=["match_mode", "ignore_reason"])
    auto_match_all(period)


@transaction.atomic
def set_transaction_match_mode(period, txn, mode, ignore_reason=""):
    if period.status == AccountingPeriod.Status.LOCKED:
        raise ValueError("该期间已锁账，不能修改匹配方式。")
    if mode not in MatchingMode.values:
        raise ValueError("不支持的匹配方式。")
    if txn.period_id != period.id:
        raise ValueError("库存流水不属于当前核算期间。")
    allocations = MatchAllocation.objects.filter(
        period=period, inventory_transaction=txn
    )
    if mode in {MatchingMode.AUTO, MatchingMode.IGNORED}:
        allocations.delete()
    else:
        allocations.filter(is_manual=False).delete()
    txn.match_mode = mode
    txn.ignore_reason = (
        str(ignore_reason or "").strip()
        if mode == MatchingMode.IGNORED
        else ""
    )
    txn.save(update_fields=["match_mode", "ignore_reason"])
    auto_match_all(period)


@transaction.atomic
def remove_manual_allocation(period, allocation_id):
    if period.status == AccountingPeriod.Status.LOCKED:
        raise ValueError("该期间已锁账，不能删除匹配。")
    allocation = MatchAllocation.objects.get(id=allocation_id, period=period)
    allocation.delete()
    auto_match_all(period)
