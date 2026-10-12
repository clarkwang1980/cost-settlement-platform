@echo off
rem 成本结算平台 本地开发加速器（cmd 包装，等价于 .venv\Scripts\python.exe dev.py ...）
rem   dev up / dev restart -w / dev status / dev log / dev down
rem 注意：不要加 -X utf8，实测会让输出丢失。
setlocal
"%~dp0.venv\Scripts\python.exe" "%~dp0dev.py" %*
