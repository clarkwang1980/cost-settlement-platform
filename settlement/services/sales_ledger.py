# -*- coding: utf-8 -*-
"""销售台账（出库+开票）应用逻辑 —— 方案A（2026-09-21 用户确认）：

**台账就是口径来源，平台不再把台账"核销"到库存流水。**

- 台账自带「结转类型」（当月发货当月开票结转成本 / 前期发货本期开票结转成本 /
  当月发货未开票挂发出商品 / 免费样品赠送计入销售费用）与「销售成本」，
  销售结转成本凭证、发出商品及销售成本统计表、月末未开票、损益表**全部按台账口径出数**。
- 以前这里会生成 LEDGER-* 系统发票并把台账的数量核销到「当月销售出库流水」/「期初发出商品明细」，
  那是为了按**流水口径**算未开票成本；用户 2026-09-21 明确：销售侧不需要平台去匹配
  （台账已经是匹配结果），而且那条线只匹配上 12% 的数量，把期末发出商品算虚增了 524 万。
- 现在 `apply_sales_ledger` 只做：台账入库 + 口径提示（不生成发票、不建核销）。
- `fill_sales_ledger_costs` 保留：按月末加权价回填台账行的销售成本单价/销售成本（结算时调用）。
"""
from __future__ import annotations

from collections import defaultdict
from decimal import ROUND_HALF_UP, Decimal

from settlement.models import (
    ImportBatch,
    OpeningGoodsIssued,
    SalesLedgerLine,
)

ZERO = Decimal("0")


def _fmt(value) -> str:
    return f"{Decimal(value or ZERO):g}"


def apply_sales_ledger(batch: ImportBatch) -> str:
    """销售台账导入后的处理：只统计口径分布并给出提示，不再核销库存流水。

    返回一句给导入结果用的话（含各类结转的行数与金额）。
    """
    lines = list(batch.sales_ledger_lines.order_by("row_no"))
    if not lines:
        return "销售台账：本次没有数据行。"
    by_type = defaultdict(lambda: {"count": 0, "cost": ZERO})
    for line in lines:
        key = (line.carry_type or "").strip() or "(未填结转类型)"
        by_type[key]["count"] += 1
        by_type[key]["cost"] += Decimal(line.cost_amount or ZERO)
    parts = [
        "%s %d 行 %s 元" % (key, info["count"], _fmt(info["cost"]))
        for key, info in sorted(by_type.items())
    ]
    return (
        "销售台账已入库 %d 行；销售结转与发出商品按台账口径计算（不再核销库存流水）。%s"
        % (len(lines), "；".join(parts))
    )


def fill_sales_ledger_costs(period):
    """结算后把台账行的 销售成本单价/销售成本/销售毛利/销售毛利率 回填。

    口径（台账只提供数量与收入，成本由平台算）：
    - 成本单价 = 该产品【本期月末加权平均价】（取自库存计价 InventoryValuation）；
    - 跨期行（出库期间早于本期）优先用期初发出商品明细上的历史单位成本，取不到再用本期加权价；
    - **本期没有月末加权价的物料（本期无出入库流水，取不到价）改用用户上传的「历史单价」
      兜底：出库单价优先、入库单价其次**（用户 2026-09-28：上传了单价，销售成本却没有变）；
    - 销售成本 = 销售数量 × 成本单价（四舍五入到分）；
    - 销售收入 = 终端销售收入（未填则取 射阳兰邦销售收入）；
    - 毛利 = 收入 − 成本；毛利率 = 毛利 ÷ 收入（收入为空或 0 时不计算）。
    返回回填的行数。
    """
    prices = {}
    for valuation in period.inventory_valuations.all():
        price = Decimal(valuation.weighted_average_price or ZERO)
        if price > ZERO:  # 0 视为「没有价」，让历史单价兜底
            prices[valuation.material_code] = price
    historical = {
        item.material_code: item for item in period.historical_prices.all()
    }

    def _historical_unit(code):
        """历史单价兜底：出库单价优先、入库单价其次（都为 0 时视为无价）。"""
        item = historical.get(code)
        if item is None:
            return None
        for candidate in (item.issue_price, item.receipt_price):
            if candidate is not None and Decimal(candidate) > ZERO:
                return Decimal(candidate)
        return None

    updated = 0
    for line in period.sales_ledger_lines.all():
        unit = prices.get(line.material_code)
        if line.out_month and line.out_month < period.month:
            opening = (
                OpeningGoodsIssued.objects.filter(
                    material_code=line.material_code,
                    original_period=line.out_month,
                    unit_cost__isnull=False,
                )
                .order_by("id")
                .first()
            )
            if opening and opening.unit_cost is not None:
                unit = Decimal(opening.unit_cost)
        if unit is None or unit <= ZERO:
            unit = _historical_unit(line.material_code)
        if unit is None:
            continue
        qty = Decimal(line.quantity or ZERO)
        cost = (qty * unit).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        revenue = line.terminal_amount if line.terminal_amount is not None else line.lb_amount
        line.cost_price = unit.quantize(Decimal("0.00000001"), rounding=ROUND_HALF_UP)
        line.cost_amount = cost
        if revenue is not None and Decimal(revenue) != ZERO:
            rev = Decimal(revenue)
            line.gross_profit = (rev - cost).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            line.gross_margin = ((rev - cost) / rev).quantize(
                Decimal("0.00000001"), rounding=ROUND_HALF_UP
            )
        else:
            line.gross_profit = None
            line.gross_margin = None
        line.save(
            update_fields=["cost_price", "cost_amount", "gross_profit", "gross_margin"]
        )
        updated += 1
    return updated
