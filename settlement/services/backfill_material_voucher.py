# -*- coding: utf-8 -*-
"""反推材料成本凭证（方案 B）。

情形③（有投料-报工超投料量）和 ④（无投料-有报工）的材料成本是按 BOM **反推**补进来的，
没有实际领料出库 → 生产成本只有贷方（完工入库）、没有借方 → 科目会变负。
这里按工单补一张凭证：
    借：生产成本—直接材料（反推补料）      X
      贷：其他应付款—暂估（反推补料）              X
摘要写明「反推材料成本」，并按工单/情形列出，一眼能看出是反推的。
"""
from decimal import Decimal

from settlement.models import AccountingPeriod, MaterialIssue
from settlement.services.costing import bom_material_prices
from settlement.services.workorder_costing import build_issue_report_rows

ZERO = Decimal("0")
DEBIT_ACCOUNT = "生产成本—直接材料"
CREDIT_ACCOUNT = "原材料"
SUMMARY_PREFIX = "反推材料成本（按BOM反推，非实际领料）"
BACKFILL_CASES = ("③", "④")


def _actual_issue_amount(period, prices):
    """按工单汇总"实际投料金额"（投料数量 × 单价表价）。"""
    out = {}
    for r in MaterialIssue.objects.filter(period=period):
        order_no = (r.work_order_no or "").strip()
        if not order_no:
            continue
        qty = Decimal(r.actual_issue_quantity or 0)
        price = Decimal(prices.get(r.material_code) or 0)
        out[order_no] = out.get(order_no, ZERO) + qty * price
    return out


def build_backfill_material_vouchers(period):
    """返回反推材料凭证行（按工单）。"""
    if not isinstance(period, AccountingPeriod):
        period = AccountingPeriod.objects.get(pk=period)
    prices = bom_material_prices(period)
    actual = _actual_issue_amount(period, prices)
    result = build_issue_report_rows(period)
    report_rows = result[0] if isinstance(result, tuple) else result
    rows = []
    for item in report_rows:
        get = item.get if isinstance(item, dict) else (lambda k, d=None: getattr(item, k, d))
        case = str(get("case_label") or get("case") or "")
        if not any(key in case for key in BACKFILL_CASES):
            continue
        order_no = str(get("order_no") or "")
        material = Decimal(get("direct_material") or 0)
        backfill = material - actual.get(order_no, ZERO)
        if backfill <= ZERO:
            continue
        rows.append(
            {
                "order_no": order_no,
                "case": case,
                "amount": backfill,
                "summary": "%s｜工单 %s｜%s" % (SUMMARY_PREFIX, order_no, case),
                "debit": DEBIT_ACCOUNT,
                "credit": CREDIT_ACCOUNT,
            }
        )
    return rows


def backfill_material_voucher_totals(period):
    rows = build_backfill_material_vouchers(period)
    return {"count": len(rows), "amount": sum((r["amount"] for r in rows), ZERO), "rows": rows}
