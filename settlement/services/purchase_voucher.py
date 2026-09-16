# -*- coding: utf-8 -*-
"""采购暂估 / 暂估回冲 凭证生成。

业务原理（用户 2026-09-16 确认）：
- 本月到货 + 本月到票  → 借 原材料 / 贷 应付账款—XX供应商（有价差时进材料成本差异）
- 本月到货 + 月末未到票 → 借 原材料 / 贷 应付账款—暂估应付
- 前期暂估 + 本月到票  → 借 应付账款—暂估应付（冲原暂估）/ 贷 应付账款—XX供应商，
                        差额进材料成本差异（平台库存不因发票改写，价差只能在凭证层体现）

口径与平台其它页面保持一致（不要改）：
- 入库金额 = 入库数量 × 合同暂估不含税价（发票不改写库存金额）；
- 暂估单价来源：合同暂估价 / 期初暂估台账单价 / 手工暂估价；
- 已批准忽略的流水与发票行一律排除（与「月末采购暂估」「发票暂估回冲明细」口径一致）；
- 期初暂估台账的历史余额**不自动进本次凭证**（避免与旧系统重复入账），只做提示与账龄展示。
"""
from __future__ import annotations

from collections import defaultdict
from decimal import ROUND_HALF_UP, Decimal

from django.db.models import Q

from settlement.models import MatchAllocation, MatchingMode
from settlement.services.matching import is_purchase_receipt
from settlement.services.purchase_accruals import (
    build_purchase_accrual_rows,
    provisional_basis,
)

ZERO = Decimal("0")
CENT = Decimal("0.01")

#: 默认科目（后续可换成科目表映射；页面说明里会写明）
ACCOUNT_RAW = "原材料"
ACCOUNT_ACCRUED = "应付账款—暂估应付"
ACCOUNT_SUPPLIER = "应付账款"
ACCOUNT_DIFFERENCE = "材料成本差异"
UNCLASSIFIED = "未分类"
MISSING_SUPPLIER = "（供应商缺失）"

SAME_MONTH = "same_month"
ACCRUAL = "accrual"
REVERSAL = "reversal"

#: 凭证分录口径说明（页面上展示）
VOUCHER_NOTE = (
    "借：原材料（本月到货，按合同暂估价计入存货）；"
    "贷：应付账款—XX供应商（本月到票的部分）／应付账款—暂估应付（月末仍未到票的部分）。"
    "以前月份暂估本月到票的：借 应付账款—暂估应付（冲原暂估）／贷 应付账款—XX供应商，"
    "发票价与暂估价不同时差额进材料成本差异。"
)

BLOCK_TITLES = {
    SAME_MONTH: "① 本月到货 + 本月到票",
    ACCRUAL: "② 本月到货 + 月末未到票",
    REVERSAL: "③ 前期暂估 + 本月到票",
}


def money(value) -> Decimal:
    return Decimal(value or ZERO).quantize(CENT, rounding=ROUND_HALF_UP)


def _split(total, part, whole) -> Decimal:
    """按数量比例拆金额（差额由调用方用减法补，保证合计不丢分）。"""
    whole = Decimal(whole or ZERO)
    if whole <= ZERO:
        return ZERO
    part = Decimal(part or ZERO)
    if part >= whole:
        return money(total)
    return money(Decimal(total or ZERO) * part / whole)


def _bucket():
    return {"amount": ZERO, "quantity": ZERO, "rows": 0}


def _add(bucket, amount, quantity, rows=1):
    bucket["amount"] += amount
    bucket["quantity"] += quantity
    bucket["rows"] += rows


def _receipts(period):
    return [
        txn
        for txn in period.inventory_transactions.exclude(
            match_mode=MatchingMode.IGNORED
        )
        .select_related("period")
        .prefetch_related(
            "allocations__contract_line", "allocations__purchase_invoice_line"
        )
        .order_by("txn_date", "document_no", "id")
        if is_purchase_receipt(txn)
    ]


def _current_allocations(period):
    """本期发票核销：({入库流水id: [核销行]}, [跨期/期初台账核销行])。"""
    queryset = (
        MatchAllocation.objects.filter(
            period=period, kind=MatchAllocation.Kind.PURCHASE_INVOICE
        )
        .exclude(purchase_invoice_line__match_mode=MatchingMode.IGNORED)
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
    by_receipt = defaultdict(list)
    prior = []
    for alloc in queryset:
        txn = alloc.inventory_transaction
        if txn is not None and txn.period_id == period.id:
            by_receipt[txn.id].append(alloc)
        else:
            prior.append(alloc)
    return by_receipt, prior


def _alloc_amount(alloc) -> Decimal:
    if alloc.amount is not None:
        return money(alloc.amount)
    line = alloc.purchase_invoice_line
    unit = (
        alloc.unit_price
        if alloc.unit_price is not None
        else (line.unit_price_excl_tax if line else ZERO)
    )
    return money(Decimal(alloc.matched_quantity or ZERO) * Decimal(unit or ZERO))


def _category_of(txn) -> str:
    return (txn.category_name or "").strip() or UNCLASSIFIED


def _supplier_of(*sources) -> str:
    for source in sources:
        if source is None:
            continue
        name = (getattr(source, "supplier_name", "") or "").strip()
        if name:
            return name
        code = (getattr(source, "supplier_code", "") or "").strip()
        if code:
            return code
    return MISSING_SUPPLIER


def _entry(summary, account, aux, side, bucket, rows=None):
    amount = money(bucket["amount"])
    return {
        "level": "line",
        "summary": summary,
        "account": account,
        "aux": aux,
        "debit": amount if side == "debit" else ZERO,
        "credit": amount if side == "credit" else ZERO,
        "quantity": bucket["quantity"],
        "source_rows": bucket["rows"] if rows is None else rows,
    }


def _expand(mapping, account, side, group, label, summary, aux_note):
    """按维度展开成分录行；group=single 时合成一行。"""
    if not mapping:
        return []
    if group == "single":
        total = _bucket()
        for item in mapping.values():
            _add(total, item["amount"], item["quantity"], item["rows"])
        if total["amount"] == ZERO:
            return []
        return [
            {
                "level": "line",
                "summary": summary,
                "account": account,
                "aux": aux_note,
                "debit": money(total["amount"]) if side == "debit" else ZERO,
                "credit": money(total["amount"]) if side == "credit" else ZERO,
                "quantity": total["quantity"],
                "source_rows": total["rows"],
            }
        ]
    rows = []
    for name, item in sorted(mapping.items(), key=lambda pair: -pair[1]["amount"]):
        if item["amount"] == ZERO and item["quantity"] == ZERO:
            continue
        account_label = account if (not name or name == UNCLASSIFIED) else "%s—%s" % (account, name)
        rows.append(_entry(summary, account_label, "", side, item))
    return rows


def _make_block(key, debit_map, debit_account, credit_map, credit_account, diff, group):
    rows = []
    rows += _expand(
        debit_map, debit_account, "debit", group, "物料大类",
        BLOCK_TITLES[key], "（明细见附件）",
    )
    rows += _expand(
        credit_map, credit_account, "credit", group, "供应商",
        BLOCK_TITLES[key], "（供应商明细见附件）",
    )
    diff_amount = money(diff["amount"]) if diff else ZERO
    if diff_amount != ZERO:
        rows.append(
            {
                "level": "line",
                "summary": BLOCK_TITLES[key] + "价差",
                "account": ACCOUNT_DIFFERENCE,
                "aux": "暂估价与发票价之差",
                "debit": diff_amount if diff_amount > ZERO else ZERO,
                "credit": -diff_amount if diff_amount < ZERO else ZERO,
                "quantity": diff.get("quantity", ZERO),
                "source_rows": diff.get("rows", 0),
            }
        )
    debit_total = sum((row["debit"] for row in rows), ZERO)
    credit_total = sum((row["credit"] for row in rows), ZERO)
    # 数量＝借方维度（物料大类）汇总，含金额为 0 的行；两种科目粒度下保持一致
    quantity = sum((item["quantity"] for item in debit_map.values()), ZERO)
    # 来源行数只算一侧（借贷两侧是同一批单据，不能相加）
    debit_rows = sum(row["source_rows"] for row in rows if row["debit"] > ZERO)
    credit_rows = sum(row["source_rows"] for row in rows if row["credit"] > ZERO)
    rows.append(
        {
            "level": "subtotal",
            "summary": "%s 小计" % BLOCK_TITLES[key],
            "account": "",
            "aux": "",
            "debit": money(debit_total),
            "credit": money(credit_total),
            "quantity": quantity,
            "source_rows": debit_rows or credit_rows,
        }
    )
    return {
        "key": key,
        "title": BLOCK_TITLES[key],
        "rows": rows,
        "line_count": len([row for row in rows if row["level"] == "line"]),
        "debit_total": money(debit_total),
        "credit_total": money(credit_total),
        "quantity": quantity,
        "balanced": money(debit_total - credit_total) == ZERO,
    }


def build_purchase_voucher(period, *, group="category", detail_supplier="",
                           detail_material=""):
    """生成采购暂估 / 暂估回冲凭证。

    group="category"：借方按物料大类、贷方按供应商展开（默认，便于核对辅助核算）
    group="single"  ：科目汇总（原材料 / 应付账款 / 暂估应付 / 材料成本差异）
    detail_supplier / detail_material 只筛选三张附件明细，不影响凭证金额。
    """
    group = "single" if str(group).strip() == "single" else "category"
    receipts = _receipts(period)
    by_receipt, prior_allocs = _current_allocations(period)

    same_debit = defaultdict(_bucket)      # 借 原材料（按大类）
    same_credit = defaultdict(_bucket)     # 贷 应付账款—供应商
    accrual_debit = defaultdict(_bucket)   # 借 原材料（按大类）
    accrual_credit = defaultdict(_bucket)  # 贷 应付账款—暂估应付（按供应商）
    same_diff = _bucket()

    receipt_rows = []
    receipt_amount_total = ZERO
    same_month_amount = ZERO
    accrual_amount = ZERO
    missing_price = 0
    missing_supplier = 0

    for txn in receipts:
        quantity = Decimal(txn.quantity or ZERO)
        amount = money(txn.amount)
        receipt_amount_total += amount
        allocations = by_receipt.get(txn.id, [])
        matched_quantity = min(
            sum((Decimal(a.matched_quantity or ZERO) for a in allocations), ZERO),
            max(quantity, ZERO),
        )
        matched_amount = _split(amount, matched_quantity, quantity)
        invoice_amount = sum((_alloc_amount(a) for a in allocations), ZERO)
        unmatched_quantity = max(quantity - matched_quantity, ZERO)
        unmatched_amount = money(amount - matched_amount)

        basis = provisional_basis(txn)
        supplier = _supplier_of(txn)
        category = _category_of(txn)
        if basis["unit_price"] is None:
            missing_price += 1
        if not (txn.supplier_name or txn.supplier_code):
            missing_supplier += 1

        if matched_quantity > ZERO:
            same_month_amount += matched_amount
            _add(same_debit[category], matched_amount, matched_quantity)
            _add(same_credit[supplier], invoice_amount, matched_quantity)
            delta = money(invoice_amount - matched_amount)
            if delta != ZERO:
                _add(same_diff, delta, matched_quantity)
        if unmatched_quantity > ZERO:
            accrual_amount += unmatched_amount
            _add(accrual_debit[category], unmatched_amount, unmatched_quantity)
            _add(accrual_credit[supplier], unmatched_amount, unmatched_quantity)

        if unmatched_quantity == ZERO:
            status = "全部到票"
        elif matched_quantity > ZERO:
            status = "部分到票"
        elif basis["unit_price"] is None:
            status = "缺少暂估价格"
        else:
            status = "全部暂估"
        receipt_rows.append(
            {
                "document_no": txn.document_no,
                "document_date": txn.txn_date,
                "receipt_month": txn.period.month,
                "supplier": supplier,
                "material_code": txn.material_code,
                "material_name": txn.material_name,
                "category": category,
                "quantity": quantity,
                "amount": amount,
                "provisional_unit_price": basis["unit_price"],
                "price_source": basis["source"],
                "contract_nos": basis["contract_nos"],
                "matched_quantity": matched_quantity,
                "matched_amount": matched_amount,
                "invoice_amount": invoice_amount,
                "unmatched_quantity": unmatched_quantity,
                "unmatched_amount": unmatched_amount,
                "status": status,
                "txn_id": txn.id,
            }
        )

    # ── ③ 前期暂估 + 本月到票（回冲 + 按票入账）────────────────────────
    reversal_debit = defaultdict(_bucket)    # 借 应付账款—暂估应付（按供应商）
    reversal_credit = defaultdict(_bucket)   # 贷 应付账款—供应商
    reversal_diff = _bucket()
    reversal_rows = []
    reversal_amount = ZERO
    reversal_invoice_amount = ZERO

    for alloc in prior_allocs:
        line = alloc.purchase_invoice_line
        txn = alloc.inventory_transaction
        opening = alloc.opening_purchase_accrual
        if txn is not None:
            basis = provisional_basis(txn)
            provisional_unit = basis["unit_price"]
            target_type = "库存流水入库"
            document_no = txn.document_no
            receipt_month = txn.period.month
            material_code = txn.material_code
            material_name = txn.material_name
            supplier = _supplier_of(line, txn)
            category = _category_of(txn)
        elif opening is not None:
            provisional_unit = Decimal(opening.provisional_unit_price or ZERO)
            target_type = "期初暂估台账"
            document_no = opening.document_no or "（期初台账）"
            receipt_month = opening.original_period
            material_code = opening.material_code
            material_name = opening.material_name
            supplier = _supplier_of(line, opening)
            category = UNCLASSIFIED
        else:
            continue

        matched_quantity = Decimal(alloc.matched_quantity or ZERO)
        provisional_amount = money(
            matched_quantity * Decimal(provisional_unit or ZERO)
        )
        invoice_amount = _alloc_amount(alloc)
        delta = money(invoice_amount - provisional_amount)

        reversal_amount += provisional_amount
        reversal_invoice_amount += invoice_amount
        _add(reversal_debit[supplier], provisional_amount, matched_quantity)
        _add(reversal_credit[supplier], invoice_amount, matched_quantity)
        if delta != ZERO:
            _add(reversal_diff, delta, matched_quantity)

        reversal_rows.append(
            {
                "invoice_no": line.invoice_no if line else "",
                "invoice_date": line.invoice_date if line else None,
                "invoice_period": line.invoice_period if line else None,
                "supplier": supplier,
                "target_type": target_type,
                "document_no": document_no,
                "receipt_month": receipt_month,
                "material_code": material_code,
                "material_name": material_name,
                "category": category,
                "matched_quantity": matched_quantity,
                "provisional_unit_price": provisional_unit,
                "provisional_amount": provisional_amount,
                "invoice_unit_price": (
                    alloc.unit_price
                    if alloc.unit_price is not None
                    else (line.unit_price_excl_tax if line else None)
                ),
                "invoice_amount": invoice_amount,
                "difference_amount": delta,
                "cross_period": True,
                "is_opening": opening is not None,
            }
        )

    blocks = [
        _make_block(SAME_MONTH, same_debit, ACCOUNT_RAW, same_credit,
                    ACCOUNT_SUPPLIER, same_diff, group),
        _make_block(ACCRUAL, accrual_debit, ACCOUNT_RAW, accrual_credit,
                    ACCOUNT_ACCRUED, _bucket(), group),
        _make_block(REVERSAL, reversal_debit, ACCOUNT_ACCRUED, reversal_credit,
                    ACCOUNT_SUPPLIER, reversal_diff, group),
    ]

    debit_total = sum((block["debit_total"] for block in blocks), ZERO)
    credit_total = sum((block["credit_total"] for block in blocks), ZERO)
    quantity_total = sum((block["quantity"] for block in blocks), ZERO)
    difference_total = money(same_diff["amount"] + reversal_diff["amount"])

    # ── 月末暂估余额（含期初台账；只展示与提示，不进本次凭证）──────────
    balance_map = defaultdict(_bucket)
    aging_map = defaultdict(_bucket)
    opening_balance = ZERO
    current_balance = ZERO
    unpriced_balance_rows = 0
    for row in build_purchase_accrual_rows(period):
        amount = row.provisional_amount or ZERO
        key = row.supplier_name or row.supplier_code or MISSING_SUPPLIER
        _add(balance_map[key], amount, row.outstanding_quantity)
        month_key = (
            row.receipt_month.strftime("%Y-%m") if row.receipt_month else "（无期间）"
        )
        _add(aging_map[month_key], amount, ZERO)
        if row.transaction is None:
            opening_balance += amount
        else:
            current_balance += amount
        if row.provisional_unit_price is None:
            unpriced_balance_rows += 1

    ending_balance = money(opening_balance + current_balance)
    balance_rows = [
        {
            "supplier": key,
            "amount": money(value["amount"]),
            "quantity": value["quantity"],
            "rows": value["rows"],
        }
        for key, value in sorted(
            balance_map.items(), key=lambda item: -item[1]["amount"]
        )
    ]
    aging_rows = [
        {
            "month": key,
            "amount": money(value["amount"]),
            "rows": value["rows"],
            "ratio": (
                float(value["amount"] / ending_balance * 100)
                if ending_balance
                else 0.0
            ),
        }
        for key, value in sorted(aging_map.items())
    ]

    # ── 前置检查 ─────────────────────────────────────────────────────────
    checks = []
    if missing_price:
        checks.append(
            {
                "level": "blocking",
                "text": "有 %d 行采购入库没有暂估单价（无合同价、也无手工价），"
                        "暂估金额会漏记，请先补价。" % missing_price,
            }
        )
    if missing_supplier:
        checks.append(
            {
                "level": "warning",
                "text": "有 %d 行采购入库没有供应商信息，应付账款辅助核算挂不上。"
                        % missing_supplier,
            }
        )
    if not receipts:
        checks.append({"level": "warning", "text": "本期没有采购入库流水。"})
    if opening_balance > ZERO and ending_balance > ZERO:
        checks.append(
            {
                "level": "info",
                "text": "期初暂估台账仍有 %s 未核销（占月末暂估余额 %.1f%%），本次未入账；"
                        "请确认是否已在原系统挂账，避免重复入账。"
                        % (
                            f"{opening_balance:,.2f}",
                            float(opening_balance / ending_balance * 100),
                        ),
            }
        )
    if unpriced_balance_rows:
        checks.append(
            {
                "level": "warning",
                "text": "月末暂估余额里有 %d 行缺少暂估单价，余额会偏低。"
                        % unpriced_balance_rows,
            }
        )
    ignored_invoices = list(
        period.purchase_invoices.filter(match_mode=MatchingMode.IGNORED)
    )
    if ignored_invoices:
        ignored_amount = sum(
            (Decimal(item.amount_excl_tax or ZERO) for item in ignored_invoices), ZERO
        )
        checks.append(
            {
                "level": "info",
                "text": "本月有 %d 行采购发票被「批准忽略」（不含税 %s），未参与核销，"
                        "不计入本次凭证。" % (len(ignored_invoices), f"{ignored_amount:,.2f}"),
            }
        )
    if money(debit_total - credit_total) != ZERO:
        checks.append(
            {
                "level": "blocking",
                "text": "借贷不平：借方 %s / 贷方 %s，差 %s。"
                        % (
                            f"{debit_total:,.2f}",
                            f"{credit_total:,.2f}",
                            f"{debit_total - credit_total:,.2f}",
                        ),
            }
        )

    # ── 附件筛选（只影响明细，不影响凭证）──────────────────────────────
    supplier_key = str(detail_supplier or "").strip().lower()
    material_key = str(detail_material or "").strip().lower()

    def _keep(row):
        if supplier_key and supplier_key not in (row.get("supplier") or "").lower():
            return False
        if material_key:
            haystack = "%s %s" % (
                row.get("material_code") or "",
                row.get("material_name") or "",
            )
            if material_key not in haystack.lower():
                return False
        return True

    filtered_receipts = [row for row in receipt_rows if _keep(row)]
    filtered_reversals = [row for row in reversal_rows if _keep(row)]
    filtered_balance = [
        row for row in balance_rows
        if not supplier_key or supplier_key in row["supplier"].lower()
    ]

    return {
        "month": period.month.strftime("%Y-%m"),
        "group": group,
        "blocks": blocks,
        "totals": {
            "debit": money(debit_total),
            "credit": money(credit_total),
            "difference": money(debit_total - credit_total),
            "balanced": money(debit_total - credit_total) == ZERO,
            "quantity": quantity_total,
            "line_count": sum(block["line_count"] for block in blocks),
        },
        "summary": {
            "receipt_amount": money(receipt_amount_total),
            "same_month_amount": money(same_month_amount),
            "accrual_amount": money(accrual_amount),
            "reversal_amount": money(reversal_amount),
            "reversal_invoice_amount": money(reversal_invoice_amount),
            "reversal_difference": money(reversal_diff["amount"]),
            "same_month_difference": money(same_diff["amount"]),
            "difference_total": difference_total,
            "ending_balance": ending_balance,
            "opening_balance": money(opening_balance),
            "current_balance": money(current_balance),
            "voucher_amount": money(debit_total),
            "receipt_count": len(receipts),
            "same_month_count": sum(
                1 for row in receipt_rows if row["matched_quantity"] > ZERO
            ),
            "accrual_count": sum(
                1 for row in receipt_rows if row["unmatched_quantity"] > ZERO
            ),
            "reversal_count": len(reversal_rows),
            "opening_count": sum(
                1 for row in reversal_rows if row["is_opening"]
            ),
            "balance_supplier_count": len(balance_rows),
            "unpriced_rows": missing_price,
        },
        "checks": checks,
        # 分录表列与口径说明（与销售结转成本凭证一致）
        "headers": ["摘要", "科目", "数量", "借方", "贷方"],
        "voucher_note": VOUCHER_NOTE,
        "receipt_rows": filtered_receipts,
        "reversal_rows": filtered_reversals,
        "balance_rows": filtered_balance,
        "aging_rows": aging_rows,
        "accounts": {
            "raw": ACCOUNT_RAW,
            "accrued": ACCOUNT_ACCRUED,
            "supplier": ACCOUNT_SUPPLIER,
            "difference": ACCOUNT_DIFFERENCE,
        },
        "filter": {
            "supplier": detail_supplier or "",
            "material": detail_material or "",
            "active": bool(supplier_key or material_key),
        },
    }
