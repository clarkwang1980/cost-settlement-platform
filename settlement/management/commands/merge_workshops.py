# -*- coding: utf-8 -*-
"""一次性数据归并：染色纱线车间(RanZhengShaXian)/染色纤维车间(RanZhengXianWei)
已合并为染色车间(RanZheng)，将两个旧车间残留的工单归并到 RanZheng，
并删除对应的车间结算项。幂等，可重复执行。

用法: python manage.py merge_workshops [--dry-run]
"""
from django.core.management.base import BaseCommand

from settlement.models import WorkOrder, WorkshopSettlement

MERGE_MAP = {
    "RanZhengShaXian": "RanZheng",
    "RanZhengXianWei": "RanZheng",
}


class Command(BaseCommand):
    help = "将染色纱线/染色纤维车间(已合并)的工单归并到染色车间，并删除旧结算项"

    def add_arguments(self, parser):
        parser.add_argument(
            "--dry-run",
            action="store_true",
            help="只统计不动数据",
        )

    def handle(self, *args, **options):
        dry = options["dry_run"]
        total_orders = 0
        for old_ws, new_ws in MERGE_MAP.items():
            qs = WorkOrder.objects.filter(production_dept_code=old_ws)
            n = qs.count()
            total_orders += n
            self.stdout.write(f"工单 {old_ws} -> {new_ws}: {n} 张")
            if n and not dry:
                qs.update(production_dept_code=new_ws)

        settlements = WorkshopSettlement.objects.filter(workshop__in=list(MERGE_MAP))
        n_settle = settlements.count()
        self.stdout.write(f"旧车间结算项待删除: {n_settle} 条")
        if n_settle and not dry:
            settlements.delete()

        # 冗余保险：残留的中文车间名(染色纱线车间/染色纤维车间)也归并
        for alias in ("染色纱线车间", "染色纤维车间"):
            n = WorkOrder.objects.filter(production_dept_code=alias).count()
            if n:
                self.stdout.write(f"工单 {alias} -> RanZheng: {n} 张(冗余)")
                total_orders += n
                if not dry:
                    WorkOrder.objects.filter(production_dept_code=alias).update(
                        production_dept_code="RanZheng"
                    )

        if dry:
            self.stdout.write(self.style.WARNING(
                f"[dry-run] 将归并工单 {total_orders} 张、删除结算项 {n_settle} 条"
            ))
        else:
            self.stdout.write(self.style.SUCCESS(
                f"完成：归并工单 {total_orders} 张、删除结算项 {n_settle} 条。"
                "投料/报工/投料进度/BOM 均按工单号自动归属染色车间。"
            ))
