# -*- coding: utf-8 -*-
"""结转诊断命令：在正式环境跑完整结算，捕获并打印完整 traceback（用于排查 500）。"""
import traceback

from django.core.management.base import BaseCommand

from settlement.models import AccountingPeriod
from settlement.services.costing import run_settlement


class Command(BaseCommand):
    help = "试跑完整结算并打印完整错误堆栈（诊断 500 用）"

    def handle(self, *args, **options):
        periods = list(AccountingPeriod.objects.order_by("-month"))
        if not periods:
            self.stdout.write("无核算期间。")
            return
        period = periods[0]
        self.stdout.write(f"试跑期间: {period.month:%Y-%m} 当前状态={period.status}")
        try:
            run = run_settlement(period, None)
            self.stdout.write(
                f"结算成功: iterations={run.iterations} converged={run.converged}"
            )
        except Exception:
            self.stdout.write("=" * 60)
            self.stdout.write("结转失败，完整堆栈如下：")
            self.stdout.write("=" * 60)
            self.stdout.write(traceback.format_exc())
            self.stdout.write("=" * 60)
            self.stdout.write(
                "请把以上输出（含 Traceback 到最后一行）完整复制发回。"
            )
