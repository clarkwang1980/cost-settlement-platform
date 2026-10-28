from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, ROUND_HALF_UP

from django.db.models import Max, Q, Sum

from settlement.models import (
    InventoryTransaction,
    MatchAllocation,
    MatchingMode,
    OpeningPurchaseAccrual,
    PurchaseContractLine,
)
from settlement.services.matching import (
    _same_party,
    invoice_unit_amount,
    is_purchase_receipt,
)


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
    fallback = _contract_price_fallback(txn)
    if fallback is not None:
        return fallback
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
    if txn.unit_price is not None and Decimal(txn.unit_price) > ZERO:
        # 兜底：流水已带系统单价（如结算回填的“月末加权平均”）时直接采用，
        # 避免“缺少暂估价格”导致暂估金额空白。
        unit_price = price(txn.unit_price)
        return {
            "unit_price": unit_price,
            "amount": money(Decimal(txn.quantity) * unit_price),
            "source": txn.price_source or "系统月末加权价",
            "contract_nos": "",
        }
    return {
        "unit_price": None,
        "amount": None,
        "source": "缺少暂估价格",
        "contract_nos": "",
    }


def _contract_price_fallback(txn):
    """无合同匹配分配时，用同物料合同的"原币未税单价"作为暂估价格兜底。

    匹配不到合同数量（合同数量为 0/剩余不足）时，只要该物料存在供应商匹配、
    且单价大于 0 的合同，就按最近合同的单价暂估，避免暂估金额缺失。
    """
    contracts = list(
        PurchaseContractLine.objects.filter(
            period__month__lte=txn.period.month,
            material_code=txn.material_code,
            unit_price_excl_tax__gt=0,
        ).exclude(match_mode=MatchingMode.IGNORED)
    )
    if not contracts:
        return None
    if txn.supplier_code or txn.supplier_name:
        contracts = [
            c
            for c in contracts
            if _same_party(
                txn.supplier_code,
                txn.supplier_name,
                c.supplier_code,
                c.supplier_name,
            )
        ]
    else:
        # 流水无供应商：仅当该物料所有有单价合同的供应商唯一时才取价，
        # 多供应商时无法确定属于哪家，不兜底。
        suppliers = {
            (c.supplier_code or "", c.supplier_name or "")
            for c in contracts
        }
        if len(suppliers) != 1:
            return None
    if not contracts:
        return None
    contracts.sort(
        key=lambda c: (
            (c.contract_date or date.min).toordinal(),
            c.id,
        )
    )
    latest = contracts[-1]
    unit_price = price(latest.unit_price_excl_tax)
    return {
        "unit_price": unit_price,
        "amount": money(Decimal(txn.quantity) * unit_price),
        "source": "合同暂估（价格参考）",
        "contract_nos": latest.contract_no,
    }


def _supplier_from_contracts(txn):
    """流水未填供应商时，从已匹配到的采购合同反查供应商，保证暂估应付可按供应商挂账。"""
    contract_ids = txn.allocations.filter(
        kind=MatchAllocation.Kind.PURCHASE_CONTRACT,
        contract_line__isnull=False,
    ).values_list("contract_line_id", flat=True)
    lines = list(PurchaseContractLine.objects.filter(id__in=contract_ids))
    for c in lines:
        if c.supplier_code or c.supplier_name:
            return c.supplier_code or "", c.supplier_name or ""
    contract_nos = [c.contract_no for c in lines if c.contract_no]
    if contract_nos:
        cands = PurchaseContractLine.objects.filter(
            period__month__lte=txn.period.month,
            material_code=txn.material_code,
            contract_no__in=contract_nos,
        )
        for c in cands:
            if c.supplier_code or c.supplier_name:
                return c.supplier_code or "", c.supplier_name or ""
    return "", ""


@dataclass
class PurchaseAccrualRow:
    transaction: InventoryTransaction | None
    opening_item: OpeningPurchaseAccrual | None
    receipt_period: object | None
    # 暂估归属月份：合同收货期间优先，否则入库流水月份
    receipt_month: object
    # 入库流水的实际月份（用来看"归属"和"流水"是否不同，页面上会标注）
    flow_month: object | None
    # 若归属来自合同收货期间，这里是那个月份；否则 None
    contract_month: object | None
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
    # 这笔到票的**归属月份**（用户 2026-09-19）：发票收货期间 → 该入库的暂估归属 → 台账原月份
    attribution_month: object | None
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


def contract_receipt_month(transaction):
    """入库流水匹配到的采购合同里「收货期间」对应的月份（要求该合同有到货签收）。

    用户 2026-09-19 规则：**合同里有到货签收的，暂估按合同的收货期间归属，不按流水月份。**

    「有到货签收」判定：合同行填了「收货期间」，且
      · 模板没提供「已签收数量」列（字段为 NULL）→ 视为有（按收货期间走），或
      · 已签收数量 > 0。
    一份流水匹配到多份合同时，取「已匹配数量」最大的那一条的收货期间。

    返回 (归属月份 或 None, 合同号 或 None)。
    """
    if transaction is None:
        return None, None
    allocations = [
        item
        for item in transaction.allocations.all()
        if item.kind == MatchAllocation.Kind.PURCHASE_CONTRACT
        and item.contract_line_id
    ]
    best = None
    for item in allocations:
        line = item.contract_line
        if line is None or line.receipt_period is None:
            continue
        signed = getattr(line, "received_quantity", None)
        if signed is not None and signed <= ZERO:
            # 合同行没有到货签收 → 不用它的收货期间，回落到入库流水月份
            continue
        quantity = Decimal(item.matched_quantity or ZERO)
        if best is None or quantity > best[0]:
            best = (quantity, line.receipt_period, line.contract_no)
    if best is None:
        return None, None
    return best[1], best[2]


def receipt_attribution_month(transaction):
    """这笔入库的**暂估归属月份**：合同收货期间优先，否则入库流水所属月份。"""
    if transaction is None:
        return None
    month, _ = contract_receipt_month(transaction)
    return month or transaction.period.month


def invoice_attribution_month(invoice_line, transaction):
    """这笔到票的**归属月份**（决定算「当月到货到票」还是「本月回冲」）。

    顺序（用户 2026-09-19）：
      ① 发票行自己填的「收货期间」——**优先，且不论核销对象是入库流水还是期初台账**
         （发票上的收货期间就是"这批货哪个月收的"，与用户报表口径一致；
          2026-09-19 修正：以前核销期初台账时会忽略它，导致 6,720.00 被错分成回冲）；
      ② 没填时：核销入库流水的 → 该入库的暂估归属（合同收货期间优先，否则流水月份）；
      ③ 没填且核销的是期初暂估台账 → 返回 None，调用方沿用原规则（按台账原入库月份）。
    """
    own = getattr(invoice_line, "receipt_period", None)
    if own is not None:
        return own
    if transaction is None:
        return None
    return receipt_attribution_month(transaction)


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


#: 采购侧的入库业务类型（用户 2026-09-21）：
#: 「退货入库」是客户退货回仓（备注里写着"客户退货"），库存结转也把它计进入库数量/金额，
#: 所以在「采购入库」页提供一个可选的「含退货入库」口径，方便与库存结转按物料分类对账。
PURCHASE_INBOUND_TYPES = ("采购入库", "退货入库")


def purchase_inbound_rows_of_period(period, include_returns=False):
    """本期间录入的采购侧入库流水。

    · 默认只取业务类型「采购入库」（真正的采购到货）；
    · include_returns=True 时再带上「退货入库」（客户退货回仓）——
      这样与「库存结转」同一物料分类的入库数量/金额可以对上。
    """
    types = PURCHASE_INBOUND_TYPES if include_returns else ("采购入库",)
    return list(
        InventoryTransaction.objects.filter(
            period=period,
            direction=InventoryTransaction.Direction.IN,
            business_type__in=types,
        )
        .exclude(match_mode=MatchingMode.IGNORED)
        .select_related("period")
        .prefetch_related(
            "allocations__contract_line",
            "allocations__purchase_invoice_line",
        )
        .order_by("txn_date", "document_no", "id")
    )


def purchase_receipts_of_period(period):
    """**本期间录入的**采购入库流水（流水所属期间＝本期），按日期/单号排序。

    用于「采购入库」页面：看本期到了哪些货、每单的暂估与到票情况。
    （与「采购暂估」不同 —— 那里按暂估**归属月份**取数，这里按流水**上传期间**取数。）
    """
    return [txn for txn in _purchase_receipts_as_of(period) if txn.period_id == period.id]


def purchase_receipts_attributed_to(period):
    """暂估归属月份＝本期的采购入库流水（合同收货期间优先，否则流水月份）。"""
    return [
        txn
        for txn in _purchase_receipts_as_of(period)
        if receipt_attribution_month(txn) == period.month
    ]


def contract_amount_by_month():
    """每个期间的「合同金额」：归属到该期间的采购合同不含税金额（合同数量 × 不含税单价）。

    归属月份与暂估是**同一套概念**（用户 2026-09-19）：
      · 合同行填了「收货期间」→ 用它；
      · 否则用该合同行**所属期间**的月份。

    去重（重要）：合同会按月「自动承接」到后续期间（`is_system_generated` +
    `carried_from`），这些副本**不重复计入** —— 只认每份合同的**源头行**
    （副本的 carried_from 还在数据里就跳过副本），否则同一份合同会在多个期间被算多次。
    """
    amounts = defaultdict(lambda: Decimal("0.00"))
    # 忽略匹配的合同不计入（与自动匹配口径一致）
    lines = list(
        PurchaseContractLine.objects.exclude(match_mode=MatchingMode.IGNORED)
        .select_related("period")
    )
    present = {line.id for line in lines}
    for line in lines:
        if line.carried_from_id and line.carried_from_id in present:
            continue
        month = line.receipt_period or line.period.month
        # 金额口径（用户 2026-09-19）：优先用文件里的「原币未税总价（报表用）」列，
        # 没有这一列的老数据再退回 数量 × 不含税单价。
        if line.amount_excl_tax is not None:
            # 不逐行取整：文件里这一列是长小数，逐行四舍五入会累积出错（用户 2026-09-19）
            amount = Decimal(line.amount_excl_tax)
        else:
            amount = money(
                Decimal(line.quantity or ZERO)
                * Decimal(line.unit_price_excl_tax or ZERO)
            )
        amounts[month] += amount
    return amounts


def build_purchase_accrual_rows(period, include_closed=False):
    rows = []
    for txn in _purchase_receipts_as_of(period):
        # 归属月份：合同「收货期间」优先（用户 2026-09-19），否则入库流水月份
        contract_month, _contract_no = contract_receipt_month(txn)
        attribution_month = contract_month or txn.period.month
        if attribution_month > period.month:
            # 归属到以后的期间 → 本期不挂这笔暂估（等归属期间再挂）
            continue
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
        supplier_code = txn.supplier_code or ""
        supplier_name = txn.supplier_name or ""
        if not supplier_code and not supplier_name:
            supplier_code, supplier_name = _supplier_from_contracts(txn)
        rows.append(
            PurchaseAccrualRow(
                transaction=txn,
                opening_item=None,
                receipt_period=txn.period,
                receipt_month=attribution_month,
                flow_month=txn.period.month,
                contract_month=contract_month,
                document_date=txn.txn_date,
                document_no=txn.document_no,
                supplier_code=supplier_code,
                supplier_name=supplier_name,
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
                flow_month=None,
                contract_month=None,
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
        if txn is None and opening is None and invoice is not None:
            # 孤儿核销行（既没入库流水、也没期初暂估台账）：跳过。
            # 不跳过的话下面 `opening.provisional_unit_price` 会 NoneType 崩掉整页（2026-10-28）。
            continue
        override = invoice_attribution_month(invoice, txn)
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
            allocation.unit_price
            if allocation.unit_price is not None
            else invoice_unit_amount(invoice)
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
                attribution_month=override,
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
                # 收货期间优先：填了就按它判断跨不跨期（用户 2026-09-19）
                cross_period=(override or receipt_month) < period.month,
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
                "attribution_month": row.attribution_month,
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
