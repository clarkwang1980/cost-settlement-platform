# -*- coding: utf-8 -*-
"""本地开发专用 settings。

只在 localhost 开发时用（`python dev.py up` 会自动带
`--settings=cost_platform.dev_settings`），Railway 正式部署完全不受影响
（entrypoint.sh 用的是 gunicorn + cost_platform.settings）。

作用：
  * 显式指定非缓存模板加载器，杜绝"改了 templates 但服务端还返回旧模板"
    的怀疑空间 —— 改完模板刷新浏览器即生效，不需要重启服务。
  * 固定 DEBUG=True，方便看报错页。
"""
from .settings import *  # noqa: F401,F403

DEBUG = True
ALLOWED_HOSTS = ["*"]

TEMPLATES[0].setdefault("OPTIONS", {})  # noqa: F405
# Django 规定：显式指定 loaders 时 APP_DIRS 必须为 False，
# 所以这里用 app_directories.Loader 手写等价行为。
TEMPLATES[0]["APP_DIRS"] = False  # noqa: F405
TEMPLATES[0]["OPTIONS"]["loaders"] = [  # noqa: F405
    "django.template.loaders.filesystem.Loader",
    "django.template.loaders.app_directories.Loader",
]
