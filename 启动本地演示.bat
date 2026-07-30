@echo off
chcp 65001 >nul
cd /d "%~dp0"
set DB_ENGINE=sqlite
set PYTHONDONTWRITEBYTECODE=1

if not exist ".venv\Scripts\python.exe" (
  echo 正在建立本地运行环境...
  python -m venv .venv
  .venv\Scripts\python.exe -m pip install -r requirements.txt
)

.venv\Scripts\python.exe manage.py migrate --noinput
.venv\Scripts\python.exe manage.py seed_demo

echo.
echo 成本结算平台正在启动...
echo 访问地址：http://127.0.0.1:8000
echo 演示账号：admin
echo 演示密码：admin123
echo.
.venv\Scripts\python.exe manage.py runserver 127.0.0.1:8000

