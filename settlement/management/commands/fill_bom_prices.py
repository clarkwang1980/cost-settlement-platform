# -*- coding: utf-8 -*-
"""重新回填 BOM 子项单价与标准成品单价（只动 BomLine.unit_price / standard_price）。

背景（用户 2026-09-30）：
  BOM 表台账里一大堆子项**没单价**（记 0），但「单价表 / 历史单价」里明明有。
  原因：回填 BOM 单价的 `_fill_bom_prices` 只查了「库存金额结转」的**月末加权价**，
  漏了平台统一的第二层兜底 —— **历史单价（出库价优先、入库价兜底）**
  （投料单价、BOM 反推早就这么取价了）。口径已在 `costing.bom_material_prices` 补上。

本命令只重算这一处，**不动其它结算结果**（实际成本/投料/库存都不受影响）。
不重新跑整个结转也能把 BOM 表的价格补对；下次正常结转也走新口径。

用法：
  python manage.py fill_bom_prices --dry-run        # 只算不写，先看会变多少
  python manage.py fill_bom_prices                  # 全部期间
  python manage.py fill_bom_prices --period 2026-07
"""
from decimal import Decimal

from django.core.management.base import BaseCommand

from settlement.models import AccountingPeriod
from settlement.services.costing import _fill_bom_prices


def _stats(period):
    """库里现状：BOM 行 / 有价行 / 无价行 / 无价编码 / 无标准价的工单。"""
    rows = list(
        period.bom_lines.values_list(
            "sub_item_code", "unit_price", "standard_price", "work_order_no"
        )
    )
    priced = [row for row in rows if Decimal(row[1] or 0) > 0]
    unpriced = [row for row in rows if not Decimal(row[1] or 0) > 0]
    return {
        "lines": len(rows),
        "priced": len(priced),
        "unpriced": len(unpriced),
        "unpriced_codes": len({row[0] for row in unpriced}),
        "orders_no_standard_price": len(
            {row[3] for row in rows if not Decimal(row[2] or 0) > 0}
        ),
    }


class Command(BaseCommand):
    help = "按「月末加权价 → 历史单价（出库价优先）」重新回填 BOM 子项单价与标准成品单价"

    def add_arguments(self, parser):
        parser.add_argument("--period", default="", help="只处理某个期间（YYYY-MM），默认全部")
        parser.add_argument("--dry-run", action="store_true", help="只算不写库")

    def handle(self, *args, **options):
        dry = options["dry_run"]
        periods = AccountingPeriod.objects.filter(bom_lines__isnull=False).distinct()
        period_arg = (options["period"] or "").strip()
        if period_arg:
            try:
                year, month = period_arg.split("-")
            except ValueError:
                self.stderr.write(self.style.ERROR("--period 格式应为 YYYY-MM"))
                return
            periods = periods.filter(month__year=int(year), month__month=int(month))
        periods = periods.order_by("month")
        if not periods:
            self.stdout.write(self.style.WARNING("没有带 BOM 的期间。"))
            return

        for period in periods:
            before = _stats(period)
            result = _fill_bom_prices(period, save=not dry)
            if dry:
                # 没写库 → 用算出来的统计当「改后」
                after = {
                    "priced": result["priced"],
                    "unpriced": result["unpriced"],
                    "unpriced_codes": result["unpriced_codes"],
                    "orders_no_standard_price": (
                        result["orders"] - result["orders_with_standard_price"]
                    ),
                }
            else:
                after = _stats(period)
            label = f"[{period.month:%Y-%m}]"
            self.stdout.write(
                f"{label} BOM 行 {before['lines']}："
                f"有单价 {before['priced']} → {after['priced']}；"
                f"没单价 {before['unpriced']} → {after['unpriced']}"
                f"（涉及编码 {before['unpriced_codes']} → {after['unpriced_codes']}）"
            )
            self.stdout.write(
                f"{label} 标准成品单价：没有值的工单 "
                f"{before['orders_no_standard_price']} → {after['orders_no_standard_price']}"
                f"（共 {result['orders']} 张）"
                + ("　[dry-run 未写库]" if dry else "")
            )
        self.stdout.write(
            self.style.SUCCESS("完成。" if not dry else "[dry-run] 未修改数据。")
        )
