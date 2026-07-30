# 成本结算平台 — Railway 自动部署方案

## 现状分析

你的项目已经**完全准备好了**，只需在 Railway 上一键连接即可：

| 组件 | 状态 | 说明 |
|------|------|------|
| GitHub 仓库 | ✅ | `clarkwang1980/cost-settlement-platform` (private) |
| Dockerfile | ✅ | 基于 `python:3.12-slim`，Railway 会自动检测 |
| entrypoint.sh | ✅ | 自动执行 migrate → collectstatic → 启动 gunicorn |
| Gunicorn | ✅ | 已经配置好，生产级 WSGI 服务器 |
| WhiteNoise | ✅ | 静态文件已配置好，不需要额外的反向代理 |
| PostgreSQL | ✅ | settings.py 通过环境变量连接 |
| .env.example | ✅ | 所有需要设置的变量都已列出 |

不需要任何代码修改！可以直接部署。

---

## 部署步骤

### 第 1 步：在 Railway 创建项目

1. 打开 [Railway](https://railway.com) 并登录（用 GitHub 账号登录最方便）
2. 点击 **New Project**
3. 选择 **Deploy from GitHub repo**
4. 授权 Railway 访问 GitHub，选择仓库 `clarkwang1980/cost-settlement-platform`
5. Railway 会自动检测到 `Dockerfile`，开始构建和部署

### 第 2 步：添加 PostgreSQL 数据库

1. 在项目画布上，点击 **+ Create** → **Database** → **Add PostgreSQL**
2. Railway 会自动部署一个 PostgreSQL 实例
3. 部署完成后，PostgreSQL 服务上会显示环境变量：`PGHOST`、`PGPORT`、`PGUSER`、`PGPASSWORD`、`PGDATABASE`

### 第 3 步：配置环境变量

在 **Web 服务**（不是数据库服务）的 **Variables** 选项卡中，添加以下变量（用 Raw Editor 一次添加最快）：

```env
DJANGO_SECRET_KEY=在命令行执行 python -c "import secrets; print(secrets.token_urlsafe(50))" 生成一个
DJANGO_DEBUG=0
DJANGO_ALLOWED_HOSTS=.railway.app,yourdomain.com

# 数据库连接（用 Railway 的变量引用语法）
POSTGRES_DB=${{Postgres.PGDATABASE}}
POSTGRES_USER=${{Postgres.PGUSER}}
POSTGRES_PASSWORD=${{Postgres.PGPASSWORD}}
POSTGRES_HOST=${{Postgres.PGHOST}}
POSTGRES_PORT=${{Postgres.PGPORT}}

# 管理员账号
DJANGO_SUPERUSER_USERNAME=admin
DJANGO_SUPERUSER_EMAIL=admin@example.com
DJANGO_SUPERUSER_PASSWORD=设置一个强密码
```

> `${{Postgres.PGxxx}}` 是 Railway 的**变量引用**语法，会自动填入数据库服务的实际连接信息。

### 第 4 步：设置公网域名

1. 切换到 Web 服务的 **Settings** 选项卡
2. 找到 **Networking** 部分
3. 点击 **Generate Domain**
4. Railway 会分配一个 `*.railway.app` 的域名给你的应用

### 第 5 步：验证部署

1. 点击 **Deploy** 按钮
2. 查看 **Logs** 选项卡，确认日志显示：
   - `Operations to perform: Apply all migrations... OK`
   - `Superuser created successfully`
   - `Listening at: http://0.0.0.0:8000`
3. 打开分配的域名，应该能看到登录页面

---

## 自动更新机制

部署完成后，每次你**推送代码到 GitHub 的 `main` 分支**，Railway 会自动：

1. 检测到新提交（通过 GitHub Webhook）
2. 重新构建 Docker 镜像
3. 停止旧容器，启动新容器
4. 自动执行 `python manage.py migrate`（entrypoint.sh 中定义）

**也就是说**，你以后的工作流就是：

```
本地修改代码 → git push → Railway 自动部署 ✅
```

**完全自动化，无需任何手动操作。**

---

## 注意事项

### 1. 上传文件（media 目录）

Django 默认使用本地文件系统存储上传文件。Railway 的磁盘是**临时**的，重启后会丢失。
如果你需要持久化上传的 CSV/Excel 文件，可以：

- **(推荐)** 使用 [Railway Volumes](https://docs.railway.com/data/volumes) 挂载持久化存储
- 或者使用云存储如阿里云 OSS / AWS S3（需要额外配置 `django-storages`）

### 2. SQLite vs PostgreSQL

Railway 上必须使用 PostgreSQL。你的 `settings.py` 已经通过 `DB_ENGINE` 环境变量区分：
- 未设置时默认走 PostgreSQL（生产环境）
- 设置 `DB_ENGINE=sqlite` 时才用 SQLite（仅本地测试）

你已经在 Railway 上使用 PostgreSQL，这部分**无需任何操作**。

### 3. 白名单域名

如果后续绑定了自己的域名，记得同步更新 `DJANGO_ALLOWED_HOSTS`。

---

## 工作流总结

```
你本地开发 (PyCharm/VSCode)
    │
    ├── 修改代码
    ├── git add → git commit → git push
    │
    ▼
GitHub (main 分支)
    │
    ├── Railway Webhook 自动触发
    │
    ▼
Railway 自动构建部署
    ├── 构建 Docker 镜像
    ├── 启动新容器
    │   ├── python manage.py migrate
    │   ├── python manage.py ensure_admin
    │   ├── python manage.py collectstatic
    │   └── gunicorn 启动
    │
    ▼
应用上线 🚀
```
