# -*- coding: utf-8 -*-
"""按「期初暂估台账 + 采购合同（签收数量）+ 采购发票」出暂估 / 回冲（用户 2026-10-28 口径）。

用户口径（原话）：
「你不要按照库存流水来，按照我上传的合同和发票，我上传的合同都是货已经收到的，
  然后再根据我上传的发票来暂估和回冲；期初台账是要看的，上次的合同也要看的，
  合同也是签收数量，发票对照合同，合同也有签收数量的，该暂估还是回冲。」

所以：**不看库存流水**，只吃三张表 ——

| 表 | 作用 | 取数 |
|---|---|---|
| 期初采购暂估台账 | 系统启用前已到货、未到票的货 | 原入库期间 / 数量 / 暂估单价 / 台账金额 |
| 采购合同（本期 + 以前期间） | 到货（合同＝货已收到） | 已签收数量 / 收货期间 / 不含税单价 |
| 采购发票（本期） | 到票 | 合同号 + 物料优先，冲不上再按供应商+物料按期间先后冲 |

判定规则（该暂估还是回冲，全看「发票冲的那笔货属于哪一期」）：
1. 到货池 = 期初台账行（归属＝原入库期间）＋ 合同行（归属＝收货期间；没填用合同所属期间）。
   归属晚于本期的合同，本期不挂暂估。
2. 到货数量：期初台账＝台账数量；合同＝**已签收数量**（没填才退回合同数量）。
3. 发票按期间先后冲池子里的到货（先冲期间早的：期初台账 → 上月合同 → 本月合同），
   冲到的数量不能超过该笔到货的剩余数量。
4. 本期发票冲的是**归属本期**的到货 → 「当月到货当月到票」（直接入账，不进暂估）；
   冲的是**归属以前期间**的（期初台账、上月合同…）→ 「本月回冲」。
5. 月末暂估 = 池子里截至本期末仍未到票的数量 × 单价。
6. 价差＝发票金额 − 冲掉的暂估金额，进「材料成本差异」。
7. 首个期间额外做一张「首期建账」：借 原材料 / 贷 应付账款—暂估应付＝期初台账金额，
   否则回冲期初台账会把暂估应付冲成借方（不想要这张可以传 include_opening_entry=False）。
"""
from __future__ import annotations

from collections import defaultdict
from decimal import ROUND_HALF_UP, Decimal

from settlement.models import (
    AccountingPeriod,
    MatchingMode,
    OpeningPurchaseAccrual,
    PurchaseContractLine,
    PurchaseInvoiceLine,
)

ZERO = Decimal("0")
CENT = Decimal("0.01")

ACCOUNT_RAW = "原材料"
ACCOUNT_ACCRUED = "应付账款—暂估应付"
ACCOUNT_SUPPLIER = "应付账款"
ACCOUNT_DIFFERENCE = "材料成本差异"
MISSING_SUPPLIER = "（供应商缺失）"

KIND_OPENING = "期初暂估台账"
KIND_CONTRACT = "采购合同"

BLOCK_OPENING = "opening"
BLOCK_SAME_MONTH = "same_month"
BLOCK_ACCRUAL = "accrual"
BLOCK_REVERSAL = "reversal"

BLOCK_TITLES = {
    BLOCK_OPENING: "首期建账-期初暂估台账",
    BLOCK_SAME_MONTH: "本月到货到票",
    BLOCK_ACCRUAL: "本月货到票未到-暂估入账",
    BLOCK_REVERSAL: "前期暂估本月到票（回冲）",
}


def money(value) -> Decimal:
    return Decimal(value or ZERO).quantize(CENT, rounding=ROUND_HALF_UP)


def _norm(value) -> str:
    return (value or "").strip()


def _share(total, part, whole) -> Decimal:
    """按数量占比拆金额（保证整行分完不丢分）。"""
    total = Decimal(total or ZERO)
    part = Decimal(part or ZERO)
    whole = Decimal(whole or ZERO)
    if whole <= ZERO:
        return ZERO
    if part >= whole:
        return money(total)
    return money(total * part / whole)


# ── 到货池 ──────────────────────────────────────────────────────────────
def _opening_rows(period):
    """期初暂估台账：系统启用前已到货未到票的货（归属＝原入库期间）。"""
    rows = []
    queryset = (
        OpeningPurchaseAccrual.objects.exclude(match_mode=MatchingMode.IGNORED)
        .filter(period__month__lte=period.month)
        .select_related("period")
    )
    for item in queryset:
        quantity = Decimal(item.opening_quantity or ZERO)
        amount = money(item.opening_amount)
        rows.append(
            {
                "kind": KIND_OPENING,
                "id": item.id,
                "document_no": _norm(item.document_no) or "（期初台账）",
                "contract_no": _norm(item.contract_no),
                "month": item.original_period,
                "supplier": _norm(item.supplier_name) or _norm(item.supplier_code) or MISSING_SUPPLIER,
                "material_code": _norm(item.material_code),
                "material_name": _norm(item.material_name),
                "quantity": quantity,
                "amount_total": amount,
                "unit_price": ((amount / quantity).quantize(Decimal("0.00000001"), rounding=ROUND_HALF_UP)
                               if quantity > ZERO else Decimal(item.provisional_unit_price or ZERO)),
                "remaining": quantity,
                "invoiced": ZERO,
                "match_ways": set(),
            }
        )
    return rows


def _contract_rows(period):
    """采购合同：本期＋以前期间（承接副本去重、忽略行排除），到货量＝已签收数量。"""
    lines = list(
        PurchaseContractLine.objects.exclude(match_mode=MatchingMode.IGNORED)
        .select_related("period")
        .order_by("id")
    )
    present = {line.id for line in lines}
    rows = []
    for line in lines:
        if line.carried_from_id and line.carried_from_id in present:
            continue
        month = line.receipt_period or (line.period.month if line.period_id else None)
        if month is None or month > period.month:
            continue
        quantity = Decimal(line.received_quantity if line.received_quantity is not None else (line.quantity or ZERO))
        unit_price = Decimal(line.unit_price_excl_tax or ZERO)
        rows.append(
            {
                "kind": KIND_CONTRACT,
                "id": line.id,
                "document_no": "",
                "contract_no": _norm(line.contract_no),
                "month": month,
                "supplier": _norm(line.supplier_name) or _norm(line.supplier_code) or MISSING_SUPPLIER,
                "material_code": _norm(line.material_code),
                "material_name": _norm(line.material_name),
                "quantity": quantity,
                "amount_total": money(quantity * unit_price),
                "unit_price": unit_price,
                "remaining": quantity,
                "invoiced": ZERO,
                "match_ways": set(),
            }
        )
    return rows


def _invoice_rows(period):
    """本期及以前上传的发票（忽略行排除），按开票先后。"""
    rows = []
    queryset = (
        PurchaseInvoiceLine.objects.exclude(match_mode=MatchingMode.IGNORED)
        .select_related("period")
        .order_by("invoice_period", "invoice_date", "invoice_no", "id")
    )
    for line in queryset:
        if line.period_id and line.period.month > period.month:
            continue
        quantity = Decimal(line.quantity or ZERO)
        unit_price = Decimal(line.unit_price_excl_tax or ZERO)
        uploaded = money(line.amount_excl_tax) if line.amount_excl_tax is not None else None
        rows.append(
            {
                "id": line.id,
                "invoice_no": _norm(line.invoice_no),
                "invoice_date": line.invoice_date,
                "invoice_period": line.invoice_period,
                "month": line.receipt_period or line.invoice_period,
                "this_period": bool(line.period_id and line.period.month == period.month),
                "contract_no": _norm(line.contract_no),
                "supplier": _norm(line.supplier_name) or _norm(line.supplier_code) or MISSING_SUPPLIER,
                "material_code": _norm(line.material_code),
                "material_name": _norm(line.material_name),
                "quantity": quantity,
                "unit_price": unit_price,
                "amount_total": uploaded if uploaded is not None else money(quantity * unit_price),
                "amount_uploaded": uploaded,
                "amount_calc": money(quantity * unit_price),
                "amount_gap": money(uploaded - money(quantity * unit_price)) if uploaded is not None else None,
                "remaining": quantity,
            }
        )
    return rows


def _allocate(pool, invoices):
    """发票冲池子里的到货：合同号+物料 → 供应商+物料，按期间先后、剩余数量为准。"""
    by_no_mat = defaultdict(list)
    by_sup_mat = defaultdict(list)
    for row in pool:
        if row["kind"] == KIND_CONTRACT and row["contract_no"]:
            by_no_mat[(row["contract_no"], row["material_code"])].append(row)
        by_sup_mat[(row["supplier"], row["material_code"])].append(row)

    allocations = []
    for invoice in invoices:
        left = invoice["remaining"]
        if left <= ZERO:
            continue
        candidates = []
        way = ""
        if invoice["contract_no"]:
            candidates = [c for c in by_no_mat.get((invoice["contract_no"], invoice["material_code"]), [])
                          if c["remaining"] > ZERO]
            way = "合同号+物料"
        if not candidates:
            candidates = [c for c in by_sup_mat.get((invoice["supplier"], invoice["material_code"]), [])
                          if c["remaining"] > ZERO]
            way = "供应商+物料"
        for item in candidates:
            if left <= ZERO:
                break
            if invoice["month"] is not None and item["month"] > invoice["month"]:
                continue  # 发票期间早于这笔到货的归属期间：货还没到，不可能开票
            take = min(item["remaining"], left)
            if take <= ZERO:
                continue
            item["remaining"] -= take
            item["invoiced"] += take
            item["match_ways"].add(way)
            allocations.append(
                {
                    "pool": item,
                    "invoice": invoice,
                    "quantity": take,
                    "way": way,
                    "provisional_amount": _share(item["amount_total"], take, item["quantity"]),
                    "invoice_amount": _share(invoice["amount_total"], take, invoice["quantity"]),
                    "provisional_unit_price": item["unit_price"],
                    "invoice_unit_price": invoice["unit_price"],
                }
            )
            left -= take
        invoice["remaining"] = left
    return allocations


def _hint(invoice, pool):
    same_material = [c for c in pool if c["material_code"] == invoice["material_code"]]
    if not same_material:
        return "台账/合同里都没有这个物料（这批货没有上传合同？）"
    same_supplier = [c for c in same_material if c["supplier"] == invoice["supplier"]]
    if same_supplier:
        if all(c["remaining"] <= ZERO for c in same_supplier):
            return "同物料同供应商的到货已全部到票（发票数量超出到货数量）"
        return "同物料同供应商还有余量，但发票期间早于这笔到货的归属期间（不冲）"
    names = "、".join(sorted({"%s（%s）" % (c["supplier"], c["contract_no"] or c["kind"]) for c in same_material})[:3])
    return "有同物料、但供应商名称不同：%s" % names


def _block(key, title, debit_rows, credit_rows, diff=None, quantity=None):
    rows = []
    for label, account, amount, qty, source_rows in debit_rows:
        if money(amount) == ZERO:
            continue
        rows.append({"level": "line", "summary": "%s-%s" % (title, label) if label else title,
                     "account": account, "debit": money(amount), "credit": ZERO,
                     "quantity": qty, "source_rows": source_rows})
    for label, account, amount, qty, source_rows in credit_rows:
        if money(amount) == ZERO:
            continue
        rows.append({"level": "line", "summary": "%s-%s" % (title, label) if label else title,
                     "account": account, "debit": ZERO, "credit": money(amount),
                     "quantity": qty, "source_rows": source_rows})
    diff_amount = money(diff or ZERO)
    if diff_amount != ZERO:
        rows.append({"level": "line", "summary": "%s-价差" % title, "account": ACCOUNT_DIFFERENCE,
                     "debit": diff_amount if diff_amount > ZERO else ZERO,
                     "credit": -diff_amount if diff_amount < ZERO else ZERO,
                     "quantity": quantity or ZERO,
                     "source_rows": sum((item[4] for item in debit_rows), 0)})
    debit_total = sum((row["debit"] for row in rows), ZERO)
    credit_total = sum((row["credit"] for row in rows), ZERO)
    total_quantity = quantity if quantity is not None else sum((item[3] for item in debit_rows), ZERO)
    rows.append({"level": "subtotal", "summary": "%s 小计" % title, "account": "",
                 "debit": money(debit_total), "credit": money(credit_total),
                 "quantity": total_quantity, "source_rows": sum((item[4] for item in debit_rows), 0)})
    return {"key": key, "title": title, "rows": rows,
            "line_count": len([row for row in rows if row["level"] == "line"]),
            "debit_total": money(debit_total), "credit_total": money(credit_total),
            "quantity": total_quantity,
            "balanced": money(debit_total - credit_total) == ZERO}


def build_invoice_matched_voucher(period, *, detail_supplier="", detail_material="", include_opening_entry=True):
    """按「期初台账 + 合同（签收数量）+ 发票」生成暂估 / 回冲凭证与明细。"""
    opening_rows = _opening_rows(period)
    contract_rows = _contract_rows(period)
    pool = sorted(opening_rows + contract_rows, key=lambda row: (row["month"], row["kind"] != KIND_OPENING, row["id"]))
    invoices = _invoice_rows(period)
    allocations = _allocate(pool, invoices)

    this_period_invoices = [row for row in invoices if row["this_period"]]
    current_allocations = [a for a in allocations if a["invoice"]["this_period"]]

    same_month = [a for a in current_allocations if a["pool"]["month"] == period.month]
    reversals = [a for a in current_allocations if a["pool"]["month"] < period.month]

    def _alloc_row(alloc):
        item, invoice = alloc["pool"], alloc["invoice"]
        return {
            "invoice_no": invoice["invoice_no"],
            "invoice_date": invoice["invoice_date"],
            "invoice_period": invoice["invoice_period"],
            "target_kind": item["kind"],
            "target_label": item["contract_no"] or item["document_no"],
            "contract_no": item["contract_no"],
            "document_no": item["document_no"],
            "target_month": item["month"],
            "supplier": item["supplier"],
            "material_code": item["material_code"],
            "material_name": item["material_name"],
            "matched_quantity": alloc["quantity"],
            "provisional_unit_price": alloc["provisional_unit_price"],
            "provisional_amount": alloc["provisional_amount"],
            "invoice_unit_price": alloc["invoice_unit_price"],
            "invoice_amount": alloc["invoice_amount"],
            "difference_amount": money(alloc["invoice_amount"] - alloc["provisional_amount"]),
            "match_way": alloc["way"],
            # 冲的对象（到货池行）的规模，回冲明细要显示「该单数量 / 累计已核销 / 剩余未到票」
            "target_quantity": item["quantity"],
            "target_invoiced": item["invoiced"],
            "target_outstanding": item["remaining"],
            "target_contract_no": item["contract_no"],
            "target_document_no": item["document_no"],
        }

    same_rows = [_alloc_row(a) for a in same_month]
    reversal_rows = [_alloc_row(a) for a in reversals]

    same_amount = sum((a["provisional_amount"] for a in same_month), ZERO)
    same_invoice = sum((a["invoice_amount"] for a in same_month), ZERO)
    same_quantity = sum((a["quantity"] for a in same_month), ZERO)
    same_diff = money(same_invoice - same_amount)

    rev_amount = sum((a["provisional_amount"] for a in reversals), ZERO)
    rev_invoice = sum((a["invoice_amount"] for a in reversals), ZERO)
    rev_quantity = sum((a["quantity"] for a in reversals), ZERO)
    rev_diff = money(rev_invoice - rev_amount)

    # ── 本月暂估（归属本期的合同里、到本期末仍没到票的部分）──
    accrual_rows = []
    accrual_amount = ZERO
    accrual_quantity = ZERO
    for item in contract_rows:
        if item["month"] != period.month or item["remaining"] <= ZERO:
            continue
        amount = _share(item["amount_total"], item["remaining"], item["quantity"])
        accrual_amount += amount
        accrual_quantity += item["remaining"]
        accrual_rows.append({
            "contract_no": item["contract_no"],
            "contract_month": item["month"],
            "supplier": item["supplier"],
            "material_code": item["material_code"],
            "material_name": item["material_name"],
            "quantity": item["quantity"],
            "invoiced_quantity": item["invoiced"],
            "outstanding_quantity": item["remaining"],
            "unit_price": item["unit_price"],
            "provisional_amount": amount,
            "match_ways": "、".join(sorted(item["match_ways"])) or "未到票",
        })

    # ── 首期建账（期初暂估台账）：只在最早的核算期间做一次 ──
    opening_amount = sum((item["amount_total"] for item in opening_rows), ZERO)
    opening_quantity = sum((item["quantity"] for item in opening_rows), ZERO)
    is_first_period = not AccountingPeriod.objects.filter(month__lt=period.month).exists()
    blocks = []
    if include_opening_entry and opening_rows and is_first_period:
        blocks.append(_block(
            BLOCK_OPENING, BLOCK_TITLES[BLOCK_OPENING],
            [("", ACCOUNT_RAW, opening_amount, opening_quantity, len(opening_rows))],
            [("", ACCOUNT_ACCRUED, opening_amount, opening_quantity, len(opening_rows))],
        ))
    blocks.append(_block(
        BLOCK_SAME_MONTH, BLOCK_TITLES[BLOCK_SAME_MONTH],
        [("", ACCOUNT_RAW, same_amount, same_quantity, len({r["invoice_no"] for r in same_rows}))],
        [("", ACCOUNT_SUPPLIER, same_invoice, same_quantity, len({r["invoice_no"] for r in same_rows}))],
        diff=same_diff, quantity=same_quantity,
    ))
    blocks.append(_block(
        BLOCK_ACCRUAL, BLOCK_TITLES[BLOCK_ACCRUAL],
        [("", ACCOUNT_RAW, accrual_amount, accrual_quantity, len(accrual_rows))],
        [("", ACCOUNT_ACCRUED, accrual_amount, accrual_quantity, len(accrual_rows))],
    ))
    blocks.append(_block(
        BLOCK_REVERSAL, BLOCK_TITLES[BLOCK_REVERSAL],
        [("冲回暂估", ACCOUNT_ACCRUED, rev_amount, rev_quantity, len(reversal_rows))],
        [("按票入账", ACCOUNT_SUPPLIER, rev_invoice, rev_quantity, len(reversal_rows))],
        diff=rev_diff, quantity=rev_quantity,
    ))

    # ── 月末暂估余额（台账 + 合同）──
    balance_map = defaultdict(lambda: {"amount": ZERO, "quantity": ZERO, "rows": 0})
    aging_map = defaultdict(lambda: {"amount": ZERO, "rows": 0})
    ending_balance = ZERO
    current_balance = ZERO
    history_balance = ZERO
    for item in pool:
        if item["remaining"] <= ZERO:
            continue
        amount = _share(item["amount_total"], item["remaining"], item["quantity"])
        ending_balance += amount
        if item["month"] == period.month:
            current_balance += amount
        else:
            history_balance += amount
        balance_map[item["supplier"]]["amount"] += amount
        balance_map[item["supplier"]]["quantity"] += item["remaining"]
        balance_map[item["supplier"]]["rows"] += 1
        month_key = item["month"].strftime("%Y-%m")
        aging_map[month_key]["amount"] += amount
        aging_map[month_key]["rows"] += 1

    balance_rows = [
        {"supplier": key, "amount": money(value["amount"]), "quantity": value["quantity"], "rows": value["rows"]}
        for key, value in sorted(balance_map.items(), key=lambda pair: -pair[1]["amount"])
    ]
    aging_rows = [
        {"month": key, "amount": money(value["amount"]), "rows": value["rows"],
         "ratio": float(value["amount"] / ending_balance * 100) if ending_balance else 0.0}
        for key, value in sorted(aging_map.items())
    ]

    # ── 未匹配发票 / 发票数据异常 ──
    unmatched = [
        {"invoice_no": row["invoice_no"], "invoice_date": row["invoice_date"],
         "invoice_period": row["invoice_period"], "contract_no": row["contract_no"],
         "supplier": row["supplier"], "material_code": row["material_code"],
         "material_name": row["material_name"], "quantity": row["quantity"],
         "unmatched_quantity": row["remaining"], "unit_price": row["unit_price"],
         "amount": row["amount_total"],
         "unmatched_amount": _share(row["amount_total"], row["remaining"], row["quantity"]),
         "hint": _hint(row, pool)}
        for row in this_period_invoices if row["remaining"] > ZERO
    ]
    anomalies = [
        {"invoice_no": row["invoice_no"], "invoice_period": row["invoice_period"],
         "contract_no": row["contract_no"], "supplier": row["supplier"],
         "material_code": row["material_code"], "quantity": row["quantity"],
         "unit_price": row["unit_price"], "amount_uploaded": row["amount_uploaded"],
         "amount_calc": row["amount_calc"], "gap": row["amount_gap"]}
        for row in this_period_invoices
        if row["amount_gap"] is not None and abs(row["amount_gap"]) >= Decimal("0.01")
    ]

    match_stat = defaultdict(lambda: {"rows": 0, "quantity": ZERO, "amount": ZERO})
    for alloc in current_allocations:
        stat = match_stat[alloc["way"]]
        stat["rows"] += 1
        stat["quantity"] += alloc["quantity"]
        stat["amount"] += alloc["invoice_amount"]
    match_stat["未匹配"]["rows"] = len(unmatched)
    match_stat["未匹配"]["amount"] = sum((row["unmatched_amount"] for row in unmatched), ZERO)

    # ── 明细筛选（只影响附件，不影响凭证金额）──
    supplier_key = str(detail_supplier or "").strip().lower()
    material_key = str(detail_material or "").strip().lower()

    def _keep(row):
        if supplier_key and supplier_key not in (row.get("supplier") or "").lower():
            return False
        if material_key:
            hay = "%s %s" % (row.get("material_code") or "", row.get("material_name") or "")
            if material_key not in hay.lower():
                return False
        return True

    detail_rows = []
    for item in pool:
        amount_out = _share(item["amount_total"], item["remaining"], item["quantity"])
        detail_rows.append({
            "kind": item["kind"],
            "label": item["contract_no"] or item["document_no"],
            "document_no": item["document_no"],
            "contract_no": item["contract_no"],
            "month": item["month"],
            "supplier": item["supplier"],
            "material_code": item["material_code"],
            "material_name": item["material_name"],
            "quantity": item["quantity"],
            "unit_price": item["unit_price"],
            "amount": item["amount_total"],
            "invoiced_quantity": item["invoiced"],
            "outstanding_quantity": item["remaining"],
            "outstanding_amount": amount_out,
            "status": ("全部到票" if item["remaining"] <= ZERO
                       else ("部分到票" if item["invoiced"] > ZERO else "全部暂估")),
            "match_ways": "、".join(sorted(item["match_ways"])) or "未到票",
        })

    debit_total = sum((block["debit_total"] for block in blocks), ZERO)
    credit_total = sum((block["credit_total"] for block in blocks), ZERO)

    # ── 生成前检查 ──
    checks = []
    zero_price = [row for row in contract_rows
                  if row["month"] == period.month and row["unit_price"] <= ZERO and row["remaining"] > ZERO]
    if zero_price:
        checks.append({
            "level": "blocking",
            "text": "有 %d 行归属本期的合同没有不含税单价，这些到货的暂估金额会按 0 记，请先补价。" % len(zero_price),
        })
    unmatched_amount_total = sum((row["unmatched_amount"] for row in unmatched), ZERO)
    if unmatched:
        checks.append({
            "level": "warning",
            "text": "本期有 %d 行采购发票冲不上期初台账/合同（不含税 %s），这些票既没暂估也没回冲，"
                    "清单见「附件4：冲不上台账/合同的发票」，需补合同或补合同号。"
                    % (len(unmatched), "{:,.2f}".format(float(unmatched_amount_total))),
        })
    if anomalies:
        gap_total = sum((Decimal(row["gap"] or ZERO) for row in anomalies), ZERO)
        checks.append({
            "level": "warning",
            "text": "本期有 %d 行发票的「不含税金额」与「数量×单价」不一致（合计差 %s），"
                    "凭证按上传金额取数，价差会跟着偏，请核对发票单价列。"
                    % (len(anomalies), "{:,.2f}".format(float(gap_total))),
        })
    if not this_period_invoices:
        checks.append({"level": "info", "text": "本期没有上传采购发票：到票/回冲按 0 计算，"
                                               "归属本期的合同会全部挂在暂估里。"})
    if not contract_rows and not opening_rows:
        checks.append({"level": "warning", "text": "本期及以前既没有期初暂估台账、也没有采购合同，凭证为空。"})
    if money(debit_total - credit_total) != ZERO:
        checks.append({"level": "blocking",
                       "text": "借贷不平：借方 %s / 贷方 %s。" % (
                           "{:,.2f}".format(float(debit_total)), "{:,.2f}".format(float(credit_total)))})
    if opening_rows and include_opening_entry and is_first_period:
        checks.append({
            "level": "info",
            "text": "首期建账：期初暂估台账 %s（%d 行）已按「借 原材料 / 贷 应付账款—暂估应付」入账；"
                    "若这批历史余额已在旧系统挂过账，请告知去掉这一块。" % (
                        "{:,.2f}".format(float(opening_amount)), len(opening_rows)),
        })

    headers = ["摘要", "科目", "数量", "借方", "贷方"]
    voucher_note = (
        "到货＝期初暂估台账 ＋ 采购合同「已签收数量」；到票＝采购发票（先按合同号+物料、再按供应商+物料冲）。"
        "本期发票冲本期到货 → 本月到货到票（借 原材料 / 贷 应付账款）；"
        "本期到货未到票 → 暂估（借 原材料 / 贷 应付账款—暂估应付）；"
        "本期发票冲以前期间到货 → 回冲（借 应付账款—暂估应付 / 贷 应付账款，价差进 材料成本差异）。"
    )

    return {
        "month": period.month.strftime("%Y-%m"),
        "basis": "opening+contract+invoice",
        "blocks": blocks,
        "headers": headers,
        "voucher_note": voucher_note,
        "checks": checks,
        "filter": {
            "supplier": detail_supplier or "",
            "material": detail_material or "",
            "active": bool(supplier_key or material_key),
        },
        "totals": {"debit": money(debit_total), "credit": money(credit_total),
                   "difference": money(debit_total - credit_total),
                   "balanced": money(debit_total - credit_total) == ZERO,
                   "quantity": sum((block["quantity"] for block in blocks), ZERO),
                   "line_count": sum((block["line_count"] for block in blocks), 0)},
        "summary": {
            "opening_rows": len(opening_rows),
            "opening_quantity": opening_quantity,
            "opening_amount": money(opening_amount),
            "opening_entry_amount": money(opening_amount) if (include_opening_entry and opening_rows and is_first_period) else ZERO,
            "contract_count": len([row for row in contract_rows if row["month"] == period.month]),
            "contract_amount": money(sum((row["amount_total"] for row in contract_rows if row["month"] == period.month), ZERO)),
            "same_month_quantity": same_quantity,
            "same_month_amount": money(same_amount),
            "same_month_invoice_amount": money(same_invoice),
            "same_month_difference": money(same_diff),
            "accrual_quantity": accrual_quantity,
            "accrual_amount": money(accrual_amount),
            "reversal_quantity": rev_quantity,
            "reversal_amount": money(rev_amount),
            "reversal_invoice_amount": money(rev_invoice),
            "reversal_difference": money(rev_diff),
            "difference_total": money(same_diff + rev_diff),
            "ending_balance": money(ending_balance),
            "current_balance": money(current_balance),
            "history_balance": money(history_balance),
            "invoice_rows": len(this_period_invoices),
            "invoice_amount": money(sum((row["amount_total"] for row in this_period_invoices), ZERO)),
            "unmatched_rows": len(unmatched),
            "unmatched_amount": money(sum((row["unmatched_amount"] for row in unmatched), ZERO)),
            "balance_supplier_count": len(balance_rows),
        },
        "match_stat": {key: {"rows": value["rows"], "quantity": value["quantity"], "amount": money(value["amount"])}
                       for key, value in match_stat.items()},
        "detail_rows": [row for row in detail_rows if _keep(row)],
        "accrual_rows": [row for row in accrual_rows if _keep(row)],
        "same_month_rows": [row for row in same_rows if _keep(row)],
        "reversal_rows": [row for row in reversal_rows if _keep(row)],
        "unmatched_rows": [row for row in unmatched if _keep(row)],
        "invoice_anomaly_rows": anomalies,
        "balance_rows": balance_rows,
        "aging_rows": aging_rows,
        "accounts": {"raw": ACCOUNT_RAW, "accrued": ACCOUNT_ACCRUED,
                     "supplier": ACCOUNT_SUPPLIER, "difference": ACCOUNT_DIFFERENCE},
    }


#: 兼容旧名字
build_contract_accrual_voucher = build_invoice_matched_voucher
