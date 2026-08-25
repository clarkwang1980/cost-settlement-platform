# -*- coding: utf-8 -*-
"""历史单价自动补全：用户只需上传一次基准历史单价，
之后期间内出现的新增物料，自动用系统计算的单价补进 HistoricalPrice。

计算口径：
- 入库单价 = 入库金额 ÷ 入库数量（无入库时用月末加权价）
- 出库单价 = 出库金额 ÷ 出库数量（无出库时用月末加权价）
- 金额与数量均为 0 的物料无价格依据，不补全。

幂等：只补「该期间该物料在 HistoricalPrice 中不存在」的记录，
已存在（含用户上传）的一律不动。
"""
from decimal import Decimal

from settlement.models import HistoricalPrice, InventoryValuation

_PRICE_PLACES = Decimal("0.0001")


def _calc_prices(v: InventoryValuation):
    """根据库存计价记录计算入库/出库单价。返回 (receipt_price, issue_price) 或 None。"""
    rq = Decimal(v.receipt_quantity or 0)
    ra = Decimal(v.receipt_amount or 0)
    iq = Decimal(v.issue_quantity or 0)
    ia = Decimal(v.issue_amount or 0)
    wavg = Decimal(v.weighted_average_price or 0)

    rp = (ra / rq).quantize(_PRICE_PLACES) if rq else wavg
    ip = (ia / iq).quantize(_PRICE_PLACES) if iq else wavg
    rp = rp.quantize(_PRICE_PLACES) if rp else Decimal("0")
    ip = ip.quantize(_PRICE_PLACES) if ip else Decimal("0")
    if rp <= 0 and ip <= 0:
        return None
    return rp, ip


def autofill_historical_prices(periods):
    """对给定期间列表，把缺失的新增物料历史单价自动补全（幂等）。

    periods: AccountingPeriod 可迭代。返回补全条数。
    """
    created = 0
    for period in periods:
        existing = set(
            HistoricalPrice.objects.filter(period=period).values_list(
                "material_code", flat=True
            )
        )
        missing = (
            InventoryValuation.objects.filter(period=period)
            .exclude(material_code__in=existing)
            .order_by("material_code")
        )
        for v in missing:
            prices = _calc_prices(v)
            if prices is None:
                continue
            receipt_price, issue_price = prices
            HistoricalPrice.objects.create(
                period=period,
                material_code=v.material_code,
                material_name=v.material_name,
                receipt_price=receipt_price,
                issue_price=issue_price,
            )
            created += 1
    return created
