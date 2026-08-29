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
        parser.add_argument("--workshop", default="",
                            help="补建主档时指定车间编码（如 ZhenZhi/针织）；仅在台账无法推断车间时使用")

    def handle(self, *args, **options):
        dry = options["dry_run"]
        ws_arg = (options.get("workshop") or "").strip()
        fq_list = list(WorkOrder.objects.filter(order_no=OLD))
        if not fq_list:
            self._backfill_missing_master(dry, ws_arg)
            self.stdout.write(self.style.SUCCESS("完成（未发现 FQ 工单，已检查是否需要补建主档）。"))
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



    def _backfill_missing_master(self, dry, ws_arg=""):
        """台账中出现了 31.295-20260320，但工单主档完全没有（也无 FQ 可改名）时自动补建。"""
        from collections import defaultdict
        from settlement.services.imports import normalize_workshop

        refs = defaultdict(set)  # period_id -> 车间候选集合
        for model, col in [(OpeningWip, "order_no"), (MaterialIssue, "work_order_no"),
                           (WorkReport, "work_order_no"), (IssueProgress, "work_order_no")]:
            for r in model.objects.filter(**{col: NEW}).select_related("period"):
                refs[r.period_id].add("")
                dept = getattr(r, "department", "") or ""
                if dept:
                    refs[r.period_id].add(normalize_workshop(dept))

        for period_id in sorted(refs):
            period = WorkOrder._meta.get_field("period").related_model.objects.get(id=period_id)
            if WorkOrder.objects.filter(period_id=period_id, order_no=NEW).exists():
                continue  # 主档已有，无需补建
            depts = {d for d in refs[period_id] if d}
            dept = sorted(depts)[0] if depts else ""
            if not dept and ws_arg:
                from settlement.services.imports import WORKSHOP_ALIASES
                dept = WORKSHOP_ALIASES.get(ws_arg, ws_arg)
            if not dept:
                self.stdout.write(self.style.WARNING(
                    f"[{period.month:%Y-%m}] 台账存在 {NEW} 但主档缺失，且无法从台账推断车间，"
                    f"请确认该工单归属车间后手动补建工单主档。"))
                continue
            if dry:
                self.stdout.write(f"[dry-run] [{period.month:%Y-%m}] 将补建主档工单 {NEW}（车间 {dept}）")
                continue
            WorkOrder.objects.get_or_create(
                period=period, order_no=NEW,
                defaults={
                    "production_dept_code": dept,
                    "order_type": "期初承接",
                    "remark": f"由 fix_fq_wo 自动补建（台账引用工单 {NEW}，车间 {dept}）。",
                },
            )
            self.stdout.write(self.style.SUCCESS(
                f"[{period.month:%Y-%m}] 已补建主档工单 {NEW}（车间 {dept}）。"))

        self.stdout.write(self.style.SUCCESS("完成。" if not dry else "[dry-run] 未修改数据。"))
