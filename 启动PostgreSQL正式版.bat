@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist ".env" (
  copy ".env.example" ".env" >nul
  echo 已生成 .env，请先修改其中的数据库密码、管理员密码和密钥。
  notepad ".env"
  pause
)

docker compose up --build

