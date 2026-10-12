# -*- coding: utf-8 -*-
"""补齐 BOM 行上的「损耗率(%) / 标准用量 / 需求数」三列（重读当初上传的那份 BOM 文件）。

背景（用户 2026-09-30）：
  这三列是后来才加到平台上的（migration `0025_bomline_required_qty`）。
  2026-08-28 导入的那批 BOM 行当时没有这三列，所以库里全是 NULL ——
  页面上显示成「—」，但用户上传的文件里这三列是有数的。

做法：
  * 按**上传批次**重读原始文件（`settlement.services.imports::read_rows`，与导入同一套表头别名）；
  * 与库里该批次的 BOM 行按 **(工单编号, 子项物料编号, 项次)** 对应（同一 key 多行按文件顺序取）；
  * **只填 NULL 的值**：已有的值不覆盖，也不碰结算回填的「单价 / 标准成品单价」。

幂等，可重复执行（已补过的行不会再变）。

用法：
  python manage.py backfill_bom_columns                 # 全部期间的 BOM 批次
  python manage.py backfill_bom_columns --period 2026-07
  python manage.py backfill_bom_columns --dry-run       # 只统计不动数据
"""
from decimal import Decimal, InvalidOperation
from pathlib import Path

from django.core.management.base import BaseCommand

from settlement.models import BomLine, ImportBatch
from settlement.services.imports import read_rows

FIELDS = ("loss_rate", "standard_qty", "required_qty")
LABELS = {"loss_rate": "损耗率(%)", "standard_qty": "标准用量", "required_qty": "需求数"}


def _key(work_order_no, sub_item_code, line_no):
    """与上传文件对应的行键：(工单编号, 子项物料编号, 项次)。"""
    return (
        str(work_order_no or "").strip(),
        str(sub_item_code or "").strip(),
        str(line_no if line_no is not None else "").strip(),
    )


def _dec(value):
    if value in (None, ""):
        return None
    try:
        return Decimal(str(value).replace(",", "").strip())
    except (InvalidOperation, ValueError):
        return None


class Command(BaseCommand):
    help = "用当初上传的 BOM 文件补齐「损耗率(%) / 标准用量 / 需求数」三列（只填空值）"

    def add_arguments(self, parser):
        parser.add_argument("--period", default="", help="只处理某个期间（YYYY-MM），默认全部")
        parser.add_argument("--dry-run", action="store_true", help="只统计不动数据")

    def handle(self, *args, **options):
        dry = options["dry_run"]
        batches = (
            ImportBatch.objects.filter(kind=ImportBatch.Kind.BOM)
            .exclude(file="")
            .select_related("period")
            .order_by("period__month", "created_at", "id")
        )
        period_arg = (options["period"] or "").strip()
        if period_arg:
            try:
                year, month = period_arg.split("-")
                batches = batches.filter(
                    period__month__year=int(year), period__month__month=int(month)
                )
            except ValueError:
                self.stderr.write(self.style.ERROR("--period 格式应为 YYYY-MM"))
                return
        if not batches:
            self.stdout.write(self.style.WARNING("没有带原始文件的 BOM 上传批次。"))
            return

        grand_lines = grand_filled = grand_present = grand_unmatched = 0
        for batch in batches:
            label = f"[{batch.period.month:%Y-%m}] 批次#{batch.id} {batch.original_name}"
            path = Path(batch.file.path)
            if not path.is_file():
                self.stdout.write(self.style.WARNING(f"{label}：原始文件已丢失，跳过"))
                continue
            try:
                source_rows = read_rows(path, "bom")
            except Exception as exc:  # 解析失败（文件损坏/格式异常）不中断其它批次
                self.stdout.write(self.style.WARNING(f"{label}：解析失败 {exc}，跳过"))
                continue

            # 文件里这三列的取值，按行键排队（同一 key 多行按文件顺序对应）
            # 兼容历史修正：`fix_fq_wo` 把工单号 FQ31.295-20260320 改成了 31.295-20260320，
            # 源文件里还带着 "FQ" 前缀，这里两种写法都登记一份。
            source = {}
            for row in source_rows:
                key = _key(
                    row.get("work_order_no"), row.get("sub_item_code"), row.get("line_no")
                )
                values = tuple(_dec(row.get(field)) for field in FIELDS)
                source.setdefault(key, []).append(values)
                if key[0].startswith("FQ"):
                    source.setdefault((key[0][2:], key[1], key[2]), []).append(values)

            lines = list(batch.bom_lines.order_by("id"))
            updates = []
            filled = present = unmatched = 0
            unmatched_samples = []
            for line in lines:
                key = _key(line.work_order_no, line.sub_item_code, line.line_no)
                queue = source.get(key)
                if not queue:
                    unmatched += 1
                    if len(unmatched_samples) < 10:
                        unmatched_samples.append(
                            f"id={line.id} 工单={key[0]!r} 子项={key[1]!r} 项次={key[2]!r}"
                        )
                    continue
                values = queue.pop(0)
                changed = {}
                for field, value in zip(FIELDS, values):
                    if getattr(line, field) is None and value is not None:
                        changed[field] = value
                if changed:
                    updates.append((line.id, changed))
                    filled += 1
                else:
                    present += 1

            if updates and not dry:
                for line_id, changed in updates:
                    BomLine.objects.filter(id=line_id).update(**changed)

            self.stdout.write(
                f"{label}：库内 {len(lines)} 行 / 源文件 {len(source_rows)} 行 → "
                f"补齐 {filled} 行、本来就有值 {present} 行、源文件对不上 {unmatched} 行"
                + ("（dry-run 未写库）" if dry else "")
            )
            for sample in unmatched_samples:
                self.stdout.write(f"    · 源文件没有对应行：{sample}")
            grand_lines += len(lines)
            grand_filled += filled
            grand_present += present
            grand_unmatched += unmatched

        self.stdout.write(
            self.style.SUCCESS(
                f"合计：{grand_lines} 行 → 补齐 {grand_filled}、已有值 {grand_present}、"
                f"对不上 {grand_unmatched}"
                + ("（dry-run 未写库）" if dry else "")
            )
        )
        if not dry:
            left = BomLine.objects.filter(
                loss_rate__isnull=True, standard_qty__isnull=True, required_qty__isnull=True
            ).count()
            self.stdout.write(f"补完后三列仍全空的行：{left} 行")
