# -*- coding: utf-8 -*-
"""对齐正式环境车间布局到测试环境（幂等，可重复执行）：
1) 染色纱线(RanZhengShaXian)/染色纤维(RanZhengXianWei) 旧编码及中文残留工单 -> RanZheng；
2) 删除旧编码的车间结算项（已锁账期间跳过并提示）；
3) 未锁账期间的车间结转顺序重排为规范顺序：
   染色车间 -> 纺纱车间 -> 包覆车间 -> 针织车间 -> 机织车间 -> 浆粕车间(追加)。

用法: python manage.py sync_workshops [--dry-run]
"""
from django.core.management.base import BaseCommand

from settlement.models import AccountingPeriod, WorkOrder, WorkshopSettlement
from settlement.services.imports import WORKSHOP_LABELS, WORKSHOP_ORDER, normalize_workshop

OLD_CODES = ["RanZhengShaXian", "RanZhengXianWei", "\u67d3\u8272\u7eb1\u7ebf\u8f66\u95f4", "\u67d3\u8272\u7ea4\u7ef4\u8f66\u95f4"]


def _ws_label(code):
    return WORKSHOP_LABELS.get(code, code)


class Command(BaseCommand):
    help = "\u5c06\u65e7\u67d3\u8272\u7eb1\u7ebf/\u67d3\u8272\u7ea4\u7ef4\u5de5\u5355\u5f52\u5e76\u5230\u67d3\u8272\u8f66\u95f4\uff0c\u5e76\u91cd\u6392\u672a\u9501\u8d26\u671f\u95f4\u7684\u8f66\u95f4\u7ed3\u8f6c\u987a\u5e8f\u4e0e\u6d4b\u8bd5\u73af\u5883\u4e00\u81f4"

    def add_arguments(self, parser):
        parser.add_argument(
            "--dry-run",
            action="store_true",
            help="\u53ea\u7edf\u8ba1\u4e0d\u52a8\u6570\u636e",
        )

    def handle(self, *args, **options):
        dry = options["dry_run"]
        total_orders = 0
        for old in OLD_CODES:
            qs = WorkOrder.objects.filter(production_dept_code=old)
            n = qs.count()
            total_orders += n
            self.stdout.write(f"\u5de5\u5355 {old} -> RanZheng: {n} \u5f20")
            if n and not dry:
                qs.update(production_dept_code="RanZheng")

        # 删除旧车间结算项（锁账期间不动，提示）
        deleted = 0
        skipped_locked = []
        for s in WorkshopSettlement.objects.filter(workshop__in=OLD_CODES).select_related("period"):
            if s.period.status == AccountingPeriod.Status.LOCKED:
                skipped_locked.append(f"{s.period} {s.workshop}")
                continue
            deleted += 1
            if not dry:
                s.delete()
        self.stdout.write(f"\u65e7\u8f66\u95f4\u7ed3\u7b97\u9879\u5f85\u5220\u9664: {deleted} \u6761"
                          + (f"\uff08\u5df2\u9501\u8d26\u8df3\u8fc7: {skipped_locked}\uff09" if skipped_locked else ""))
        if skipped_locked and not dry:
            self.stdout.write(self.style.WARNING("\u63d0\u793a\uff1a\u5df2\u9501\u8d26\u671f\u95f4\u672a\u5904\u7406\uff0c\u9700\u8981\u7684\u8bdd\u8bf7\u5148\u91cd\u65b0\u6253\u5f00\u8be5\u671f\u95f4\u540e\u91cd\u8dd1\u672c\u547d\u4ee4\u3002"))

        # 重排未锁账期间车间结转顺序
        for period in AccountingPeriod.objects.all().order_by("month"):
            if period.status == AccountingPeriod.Status.LOCKED:
                continue
            workshops = set()
            for code in period.work_orders.exclude(production_dept_code="").values_list("production_dept_code", flat=True):
                w = normalize_workshop(code)
                if w:
                    workshops.add(w)
            for w0 in period.workshop_cost_pools.exclude(workshop="").values_list("workshop", flat=True):
                w = normalize_workshop(w0)
                if w:
                    workshops.add(w)
            if not workshops:
                continue
            ordered = [w for w in WORKSHOP_ORDER if w in workshops] + sorted(workshops - set(WORKSHOP_ORDER))
            items = {s.workshop: s for s in period.workshop_settlements.all()}
            changed = []
            for idx, w in enumerate(ordered, 1):
                if w in items and items[w].sequence != idx:
                    if not dry:
                        items[w].sequence = idx
                        items[w].save(update_fields=["sequence"])
                    changed.append(w)
            self.stdout.write(
                f"{period}: \u8f66\u95f4\u987a\u5e8f -> "
                + " \u2192 ".join(f"{_ws_label(w)}({i})" for i, w in enumerate(ordered, 1))
                + (f"\uff1b\u5df2\u8c03\u6574: {changed}" if changed else "\uff1b\u5df2\u4e00\u81f4")
            )

        if dry:
            self.stdout.write(self.style.WARNING(f"[dry-run] \u5c06\u5f52\u5e76\u5de5\u5355 {total_orders} \u5f20\u3001\u5220\u9664\u7ed3\u7b97\u9879 {deleted} \u6761\u3001\u8c03\u6574\u7ed3\u8f6c\u987a\u5e8f\u3002"))
        else:
            self.stdout.write(self.style.SUCCESS("\u5b8c\u6210\u3002\u5982\u679c\u6d89\u53ca\u5df2\u7ed3\u8f6c\u7684\u671f\u95f4\uff0c\u8bf7\u5728\u8be5\u671f\u95f4\u8f66\u95f4\u7ed3\u8f6c\u9875\u70b9\u201c\u91cd\u7f6e\u8f66\u95f4\u7ed3\u8f6c\u201d\u540e\u91cd\u65b0\u7ed3\u8f6c\u3002"))
