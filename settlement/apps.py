# -*- coding: utf-8 -*-
"""应用启动钩子。

* 本地 sqlite（`DB_ENGINE=sqlite`）连接建立时统一设置 PRAGMA：
  `settings.py` 里 sqlite 分支写的 `OPTIONS["init_command"]` 在当前 Django 的
  sqlite 后端上**不会执行**（实测 `PRAGMA cache_size` 没生效），所以放到
  `connection_created` 信号里执行，保证服务进程与命令行都一致。
* 为什么必须调大页缓存（2026-09-30 实测）：WAL 模式下缓存默认 2MB，车间结转这种
  分钟级长写事务要"换出脏页"时会去开临时文件，本机环境拒绝 → 报
  `OperationalError: unable to open database file`（同一句 bulk_update：
  2MB 必失败、64MB 必成功）。本库才 ~14MB，给 128MB 缓存后整库常驻内存。
"""
from django.apps import AppConfig
from django.db.backends.signals import connection_created
from django.dispatch import receiver

#: sqlite 连接建立时执行的 PRAGMA（本地调试库用；其它后端忽略）
SQLITE_PRAGMAS = (
    "PRAGMA journal_mode=WAL",
    "PRAGMA synchronous=NORMAL",
    "PRAGMA busy_timeout=120000",
    "PRAGMA cache_size=-131072",  # 128MB 页缓存，长事务不换页
    "PRAGMA temp_store=MEMORY",
)


@receiver(connection_created)
def _setup_sqlite_connection(sender, connection, **kwargs):  # noqa: ARG001
    if connection.vendor != "sqlite":
        return
    if connection.in_atomic_block:
        return  # journal_mode 不能在事务里改
    with connection.cursor() as cursor:
        for pragma in SQLITE_PRAGMAS:
            try:
                cursor.execute(pragma)
            except Exception:  # noqa: BLE001 - PRAGMA 不支持时不影响功能
                continue


class SettlementConfig(AppConfig):
    default_auto_field = "django.db.models.BigAutoField"
    name = "settlement"
    verbose_name = "成本结算"
