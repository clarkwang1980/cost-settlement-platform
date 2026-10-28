# -*- coding: utf-8 -*-
"""月末采购暂估页（/periods/<id>/purchase-accruals/）的数据源 —— 合同口径（用户 2026-10-28）。

与凭证页同一套到货池，保证两页数字一致：

    到货 = 期初暂估台账 ＋ 采购合同「已签收数量」
    到票 = 采购发票（合同号+物料优先，冲不上再按供应商+物料按期间先后冲）
    暂估 = 到货 − 截至期末已到票；回冲 = 本期发票冲以前期间到货

本模块把凭证计算器的结果转成页面模板需要的行对象（属性名与旧的 PurchaseAccrualRow 一致），
所以模板、筛选、分页都不用改。
"""
from __future__ import annotations

from decimal import Decimal
from types import SimpleNamespace

from settlement.models import AccountingPeriod
from settlement.services.purchase_contract_voucher import (
    KIND_OPENING,
    build_invoice_matched_voucher,
    money,
)

ZERO = Decimal("0")


class Row(SimpleNamespace):
    """模板行：既能 `row.xxx` 取，也能 `row.get("xxx")` 取（页面筛选对「本月暂估」用 get）。"""

    def get(self, name, default=None):
        return getattr(self, name, default)

    def __getitem__(self, name):
        try:
            return getattr(self, name)
        except AttributeError as error:
            raise KeyError(name) from error


def _invoice(row):
    """模板只用 invoice.invoice_period / invoice_date / supplier_name / supplier_code。"""
    return SimpleNamespace(
        invoice_period=row["invoice_period"],
        invoice_date=row["invoice_date"],
        supplier_name=row["supplier"],
        supplier_code="",
        invoice_no=row["invoice_no"],
    )


def _accrual_row(item, period_month):
    """到货池一行 → 模板行（属性名与 PurchaseAccrualRow 一致）。"""
    is_opening = item["kind"] == KIND_OPENING
    return Row(
        transaction=None,
        opening_item=object() if is_opening else None,
        receipt_period=None,
        receipt_month=item["month"],
        flow_month=None,
        contract_month=None if is_opening else item["month"],
        document_date=item["month"],
        document_no=item["label"] or item["document_no"],
        supplier_code="",
        supplier_name=item["supplier"],
        material_code=item["material_code"],
        material_name=item["material_name"],
        contract_nos=item["contract_no"] or "",
        receipt_quantity=item["quantity"],
        invoiced_quantity=item["invoiced_quantity"],
        outstanding_quantity=item["outstanding_quantity"],
        provisional_unit_price=item["unit_price"],
        provisional_amount=item["outstanding_amount"],
        amount=item["outstanding_amount"],
        price_source=item["kind"] if item["kind"] == KIND_OPENING else "合同暂估",
        latest_invoice_period=None,
        status=item["status"],
        source_label=item["kind"],
        supplier=item["supplier"],
    )


def _alloc_row(row, *, cross_period):
    """发票核销分配 → 模板行（差异块 / 回冲块 / 本月到票明细共用）。"""
    difference = row["difference_amount"]
    document_no = row["target_label"] or "—"
    if row["target_kind"] == KIND_OPENING:
        document_no = "—"  # 模板在期初台账行后面会补「期初台账」小字，避免重复
    return Row(
        invoice=_invoice(row),
        invoice_no=row["invoice_no"],
        invoice_date=row["invoice_date"],
        invoice_period=row["invoice_period"],
        target_type=row["target_kind"],
        transaction=None,
        receipt_month=row["target_month"],
        attribution_month=row["target_month"],
        document_date=row["target_month"],
        document_no=document_no,
        source_label=row["target_kind"],
        supplier=row["supplier"],
        supplier_name=row["supplier"],
        supplier_code="",
        material_code=row["material_code"],
        material_name=row["material_name"],
        matched_quantity=row["matched_quantity"],
        provisional_unit_price=row["provisional_unit_price"],
        provisional_amount=row["provisional_amount"],
        invoice_unit_price=row["invoice_unit_price"],
        invoice_amount=row["invoice_amount"],
        difference_amount=difference,
        receipt_quantity=row.get("target_quantity", ZERO),
        cumulative_quantity=row.get("target_invoiced", ZERO),
        outstanding_quantity=row.get("target_outstanding", ZERO),
        cross_period=cross_period,
        status="有差异" if (difference or ZERO) != ZERO else "无差异",
    )


def build_accrual_page(period):
    """月末采购暂估页的全部数据：到票与暂估、5 张明细、主明细、差异块、回冲块。"""
    periods = list(AccountingPeriod.objects.order_by("month"))
    data = {item.id: build_invoice_matched_voucher(item) for item in periods}
    current = data[period.id]
    index = next(i for i, item in enumerate(periods) if item.id == period.id)

    # ── 到票与暂估（一行一个期间）──
    flow_rows = []
    previous_ending = None
    for i, item in enumerate(periods):
        summary = data[item.id]["summary"]
        opening = (
            data[periods[i - 1].id]["summary"]["ending_balance"] if i else summary["opening_amount"]
        )
        flow_rows.append({
            "period": item,
            "current": item.id == period.id,
            "opening_accrual": money(opening),
            "contract_amount": summary["contract_amount"],
            "invoice_matched": money(
                summary["same_month_invoice_amount"] + summary["reversal_invoice_amount"]
            ),
            # 「当月到货到票」列是**发票金额**口径（＝本月到票里核销归属本期到货的部分），
            # 这样 本月到票 ＝ 当月到货到票 ＋ 本月回冲 恒成立；暂估口径的差额由「暂估价差」解释
            "same_month": summary["same_month_invoice_amount"],
            "month_accrual": summary["accrual_amount"],
            "cross": summary["reversal_invoice_amount"],
            "difference_amount": summary["difference_total"],
            "ending_accrual": summary["ending_balance"],
            "same_month_count": len(data[item.id]["same_month_rows"]),
            "month_accrual_count": len(data[item.id]["accrual_rows"]),
            "invoice_matched_count": len(data[item.id]["same_month_rows"]) + len(data[item.id]["reversal_rows"]),
            "cross_count": len(data[item.id]["reversal_rows"]),
            "opening_accrual_source": "期初台账总额" if i == 0 else "上一期期末",
        })
        previous_ending = summary["ending_balance"]

    flow_total = {
        key: sum((row[key] for row in flow_rows), ZERO)
        for key in ("same_month", "month_accrual", "invoice_matched", "cross",
                    "contract_amount", "difference_amount")
    }
    for key in ("same_month_count", "month_accrual_count", "invoice_matched_count", "cross_count"):
        flow_total[key] = sum(row[key] for row in flow_rows)
    flow_total["opening_accrual"] = flow_rows[0]["opening_accrual"] if flow_rows else ZERO
    flow_total["ending_accrual"] = flow_rows[-1]["ending_accrual"] if flow_rows else ZERO

    # ── 5 张明细 ──
    # 期初暂估：上一期间期末仍未到票的明细；首个期间＝期初台账全部
    if index > 0:
        previous = data[periods[index - 1].id]
        opening_items = [row for row in previous["detail_rows"] if row["outstanding_quantity"] > ZERO]
        opening_from_previous = periods[index - 1]
    else:
        opening_items = [row for row in current["detail_rows"] if row["kind"] == KIND_OPENING]
        opening_from_previous = None
    opening_rows = [_accrual_row(row, period.month) for row in opening_items]

    # 本月暂估：归属本期的到货里仍未到票的（合同）
    month_accrual_items = [row for row in current["detail_rows"]
                           if row["kind"] != KIND_OPENING
                           and row["month"] == period.month
                           and row["outstanding_quantity"] > ZERO]
    month_accrual_rows = [_accrual_row(row, period.month) for row in month_accrual_items]

    invoice_rows = [_alloc_row(row, cross_period=False) for row in current["same_month_rows"]] + [
        _alloc_row(row, cross_period=True) for row in current["reversal_rows"]
    ]
    same_month_rows = [_alloc_row(row, cross_period=False) for row in current["same_month_rows"]]
    cross_rows = [_alloc_row(row, cross_period=True) for row in current["reversal_rows"]]

    # 主明细（月末暂估余额）：所有仍未到票的到货
    outstanding_items = [row for row in current["detail_rows"] if row["outstanding_quantity"] > ZERO]
    all_rows = [_accrual_row(row, period.month) for row in outstanding_items]

    # ── 暂估差异三块（＝发票金额 − 暂估金额，逐笔核销）──
    diff_rows = invoice_rows
    diff_summary_rows = []
    for item in periods:
        rows = [_alloc_row(row, cross_period=False) for row in data[item.id]["same_month_rows"]] + [
            _alloc_row(row, cross_period=True) for row in data[item.id]["reversal_rows"]
        ]
        diff_summary_rows.append({
            "period": item,
            "current": item.id == period.id,
            "matched_quantity": sum((row.matched_quantity for row in rows), ZERO),
            "provisional_amount": money(sum((row.provisional_amount for row in rows), ZERO)),
            "invoice_amount": money(sum((row.invoice_amount for row in rows), ZERO)),
            "difference_amount": money(sum((row.difference_amount for row in rows), ZERO)),
            "cross_period_count": sum(1 for row in rows if row.cross_period),
            "pending_count": 0,
        })
    diff_summary_total = {
        "matched_quantity": sum((row["matched_quantity"] for row in diff_summary_rows), ZERO),
        "provisional_amount": money(sum((row["provisional_amount"] for row in diff_summary_rows), ZERO)),
        "invoice_amount": money(sum((row["invoice_amount"] for row in diff_summary_rows), ZERO)),
        "difference_amount": money(sum((row["difference_amount"] for row in diff_summary_rows), ZERO)),
        "cross_period_count": sum(row["cross_period_count"] for row in diff_summary_rows),
        "pending_count": 0,
    }
    grouped = {}
    for row in diff_rows:
        key = (row.supplier, row.material_code, row.material_name)
        bucket = grouped.setdefault(key, {
            "supplier": row.supplier, "material_code": row.material_code,
            "material_name": row.material_name, "matched_quantity": ZERO,
            "provisional_amount": ZERO, "invoice_amount": ZERO,
            "difference_amount": ZERO, "pending_count": 0,
        })
        bucket["matched_quantity"] += row.matched_quantity
        bucket["provisional_amount"] += row.provisional_amount
        bucket["invoice_amount"] += row.invoice_amount
        bucket["difference_amount"] += row.difference_amount
    diff_grouped_rows = sorted(
        grouped.values(), key=lambda item: -abs(item["difference_amount"])
    )
    diff_totals = {
        "matched_quantity": diff_summary_total["matched_quantity"],
        "provisional_amount": diff_summary_total["provisional_amount"],
        "invoice_amount": diff_summary_total["invoice_amount"],
        "difference_amount": current["summary"]["difference_total"],
    }

    # ── 回冲明细（逐笔）──
    reversal_period_rows = []
    for item in periods:
        rows = data[item.id]["reversal_rows"]
        reversal_period_rows.append({
            "period": item,
            "current": item.id == period.id,
            "count": len(rows),
            "matched_quantity": sum((row["matched_quantity"] for row in rows), ZERO),
            "provisional_amount": money(sum((row["provisional_amount"] for row in rows), ZERO)),
            "invoice_amount": money(sum((row["invoice_amount"] for row in rows), ZERO)),
            "difference_amount": money(sum((row["difference_amount"] for row in rows), ZERO)),
            "cross_period_count": len(rows),
        })
    reversal_period_total = {
        "count": sum(row["count"] for row in reversal_period_rows),
        "matched_quantity": sum((row["matched_quantity"] for row in reversal_period_rows), ZERO),
        "provisional_amount": money(sum((row["provisional_amount"] for row in reversal_period_rows), ZERO)),
        "invoice_amount": money(sum((row["invoice_amount"] for row in reversal_period_rows), ZERO)),
        "difference_amount": money(sum((row["difference_amount"] for row in reversal_period_rows), ZERO)),
        "cross_period_count": sum(row["cross_period_count"] for row in reversal_period_rows),
    }

    return {
        "flow_rows": flow_rows,
        "flow_total": flow_total,
        "details": {
            "opening": opening_rows,
            "invoice": invoice_rows,
            "same_month": same_month_rows,
            "month_accrual": month_accrual_rows,
            "cross": cross_rows,
            "opening_from_previous": opening_from_previous,
        },
        "all_rows": all_rows,
        "diff_summary_rows": diff_summary_rows,
        "diff_summary_total": diff_summary_total,
        "diff_grouped_rows": diff_grouped_rows,
        "diff_rows": diff_rows,
        "diff_totals": diff_totals,
        "diff_pending_count": 0,
        "reversal_period_rows": reversal_period_rows,
        "reversal_period_total": reversal_period_total,
        "reversal_rows": [_alloc_row(row, cross_period=True) for row in current["reversal_rows"]],
    }
