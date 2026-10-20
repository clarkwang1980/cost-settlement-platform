# -*- coding: utf-8 -*-
"""反推材料成本凭证（用户 2026-10-17 定：方案 B）。

情形③（有投料-报工超投料量）和 ④（无投料-有报工）的材料成本是**按 BOM 反推**补进来的，
没有实际领料出库 → 生产成本只有贷方（完工入库）、没有借方 → 科目会变负。
这里补一张凭证，把反推金额补进生产成本，贷方用"其他应付款—暂估"，不动库存。

分录：
    借：生产成本—直接材料（反推补料）      X
      贷：其他应付款—暂估（反推补料）              X

摘要明确写「反推材料成本」，并且按工单列示，一眼能看出是反推的。
"""
from decimal import Decimal

from settlement.models import AccountingPeriod, WorkOrderCostResult
from settlement.services.workorder_costing import (
    CASE_NO_REPORT,
    CASE_REPORT_EXCEED,
    CASE_REPORT_ONLY,
    CASE_ROVING_PLAN,
)

ZERO = Decimal("0")
DEBIT_ACCOUNT = "生产成本—直接材料"
CREDIT_ACCOUNT = "其他应付款—暂估"
SUMMARY_PREFIX = "反推材料成本（按BOM反推，非实际领料）"


def _backfill_amount(item):
    """材料成本 − 实际投料金额 = 反推金额（>0 才算）。"""
    material = Decimal(getattr(item, "direct_material", 0) or 0)
    actual = Decimal(getattr(item, "actual_issue_material", 0) or 0)
    value = material - actual
    return value if value > ZERO else ZERO


def build_backfill_material_vouchers(period):
    """返回按工单的反推材料凭证行：[{order_no, case, amount, summary, debit, credit}]。"""
    if not isinstance(period, AccountingPeriod):
        period = AccountingPeriod.objects.get(pk=period)
    rows = []
    for item in WorkOrderCostResult.objects.filter(period=period).select_related("work_order"):
        case = getattr(item, "case", "") or ""
        if not any(key in case for key in (CASE_REPORT_EXCEED, CASE_REPORT_ONLY, CASE_ROVING_PLAN)):
            continue
        amount = _backfill_amount(item)
        if amount <= ZERO:
            continue
        order_no = item.work_order.order_no
        rows.append(
            {
                "order_no": order_no,
                "case": case,
                "amount": amount,
                "summary": "%s｜工单 %s｜%s" % (SUMMARY_PREFIX, order_no, case),
                "debit": DEBIT_ACCOUNT,
                "credit": CREDIT_ACCOUNT,
            }
        )
    return rows


def backfill_material_voucher_totals(period):
    rows = build_backfill_material_vouchers(period)
    return {
        "count": len(rows),
        "amount": sum((r["amount"] for r in rows), ZERO),
        "rows": rows,
    }
