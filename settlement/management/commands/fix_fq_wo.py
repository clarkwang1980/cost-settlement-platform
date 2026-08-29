# -*- coding: utf-8 -*-
"""修正工单号：FQ31.295-20260320 -> 31.295-20260320（FQ 前缀为 ERP 误加）。

幂等，可重复执行：
- 若某期间仅存在 FQ 工单：直接改名；
- 若 FQ 与目标工单同时存在：把 FQ 的 BOM/工单明细/期初在产品等改挂到目标工单，
  补全字段后删除 FQ 工单；
- 若不存在 FQ 工单：跳过。

用法: python manage.py fix_fq_wo [--dry-run]
"""
from django.core.management.base import BaseCommand

from settlement.models import (
    BomLine, IssueProgress, MaterialIssue, OpeningWip,
    WorkOrder, WorkOrderDetail, WorkReport,
)

OLD = "FQ31.295-20260320"
NEW = "31.295-20260320"
MERGE_FIELDS = [
    "identifier", "plan_quantity", "planned_start", "planned_finish",
    "planning_dept_code", "planner_code", "specified_material", "status",
    "use_routing", "process_plan_type", "production_mode",
    "source_order_no", "material_color",
]


def _has(v):
    return v not in (None, "", 0, False)


class Command(BaseCommand):
    help = "把工单号 FQ31.295-20260320 统一为 31.295-20260320（幂等）"

    def add_arguments(self, parser):
        parser.add_argument("--dry-run", action="store_true", help="只统计不动数据")

    def handle(self, *args, **options):
        dry = options["dry_run"]
        fq_list = list(WorkOrder.objects.filter(order_no=OLD))
        if not fq_list:
            self.stdout.write(self.style.SUCCESS("未发现工单号 %s，无需处理。" % OLD))
            return

        for wo_fq in fq_list:
            period = wo_fq.period
            wo_a = WorkOrder.objects.filter(period=period, order_no=NEW).first()
            self.stdout.write(
                f"[{period.month:%Y-%m}] {OLD} id={wo_fq.id}"
                + (f" -> 已存在 {NEW} id={wo_a.id}，执行合并" if wo_a else f" -> 无 {NEW}，直接改名")
            )

            # 改挂关联数据（字符串外键类）
            rels = [
                (BomLine, "work_order_no"),
                (WorkOrderDetail, "order_no"),
                (OpeningWip, "order_no"),
                (MaterialIssue, "work_order_no"),
                (WorkReport, "work_order_no"),
                (IssueProgress, "work_order_no"),
            ]
            for model, col in rels:
                n = model.objects.filter(period=period, **{col: OLD}).count()
                if n:
                    self.stdout.write(f"  {model.__name__} 改挂 {n} 条")
                    if not dry:
                        model.objects.filter(period=period, **{col: OLD}).update(**{col: NEW})

            if wo_a:
                merged = [f for f in MERGE_FIELDS if _has(getattr(wo_fq, f))]
                self.stdout.write(f"  补全字段: {', '.join(merged)}")
                if not dry:
                    for f in merged:
                        setattr(wo_a, f, getattr(wo_fq, f))
                    wo_a.save()
                    wo_fq.delete()
            else:
                if not dry:
                    wo_fq.order_no = NEW
                    wo_fq.save()

        self.stdout.write(self.style.SUCCESS("完成。" if not dry else "[dry-run] 未修改数据。"))
