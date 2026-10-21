# -*- coding: utf-8 -*-
"""历史单价自动补全：用户只需上传一次基准历史单价，
之后期间内出现的新增物料，自动用系统计算的单价补进 HistoricalPrice。

计算口径（与单价表页面/导出一致）：
- 入库单价 = 入库金额 ÷ 入库数量，且**扣掉内部流转**（调拨/库存调整这类只有数量没有金额的行）
- 出库单价 = 出库金额 ÷ 出库数量，同样扣掉内部流转
- 金额与数量均为 0 的物料无价格依据，不补全。

**只给已锁账期间补价**（用户 2026-10-21）：以前每次打开单价表都会给当期的物料补一条，
而当时可能还没跑完结转（报工入库没价、调拨未计价），于是写进去一条 0.13 这种假价；
补全又是幂等的（只补不改），这条假价就一直冻在那儿，重算也不刷新。
锁账后价格才是最终值，这时补才安全。
"""
from decimal import Decimal

from settlement.models import AccountingPeriod, HistoricalPrice, InventoryValuation
from settlement.services.price_table import internal_transfer_qty

_PRICE_PLACES = Decimal("0.0001")


def _calc_prices(v: InventoryValuation, internal_in=Decimal("0"), internal_out=Decimal("0")):
    """根据库存计价记录计算入库/出库单价。返回 (receipt_price, issue_price) 或 None。"""
    rq = Decimal(v.receipt_quantity or 0) - Decimal(internal_in or 0)
    ra = Decimal(v.receipt_amount or 0)
    iq = Decimal(v.issue_quantity or 0) - Decimal(internal_out or 0)
    ia = Decimal(v.issue_amount or 0)
    wavg = Decimal(v.weighted_average_price or 0)

    rp = (ra / rq).quantize(_PRICE_PLACES) if rq > 0 else wavg
    ip = (ia / iq).quantize(_PRICE_PLACES) if iq > 0 else wavg
    rp = rp.quantize(_PRICE_PLACES) if rp else Decimal("0")
    ip = ip.quantize(_PRICE_PLACES) if ip else Decimal("0")
    if rp <= 0 and ip <= 0:
        return None
    return rp, ip


def autofill_historical_prices(periods):
    """对给定期间列表，把缺失的新增物料历史单价自动补全（幂等）。

    periods: AccountingPeriod 可迭代。返回补全条数。
    """
    periods = [p for p in periods if p.status == AccountingPeriod.Status.LOCKED]
    if not periods:
        return 0
    internal = internal_transfer_qty(periods)
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
        month = period.month.month
        for v in missing:
            int_in, int_out = internal.get((month, v.material_code), (Decimal("0"), Decimal("0")))
            prices = _calc_prices(v, int_in, int_out)
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
