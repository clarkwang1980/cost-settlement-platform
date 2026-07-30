# 成本结算平台 — Railway 自动部署方案

## 一键部署（3 步搞定）

### 第 1 步：在 Railway 创建项目

1. 打开 [Railway](https://railway.com) 用 GitHub 登录
2. **New Project** → **Deploy from GitHub repo** → 选择 `cost-settlement-platform`

### 第 2 步：添加 PostgreSQL

1. 在画布上点 **+ Create** → **Database** → **Add PostgreSQL**

### 第 3 步：设置环境变量

数据库连接**无需配置**（代码已自动兼容 Railway）。只需在 Web 服务的 Variables 里添加这些：

```env
DJANGO_SECRET_KEY=执行: python -c "import secrets; print(secrets.token_urlsafe(50))" 生成
DJANGO_DEBUG=0
DJANGO_ALLOWED_HOSTS=.railway.app
DJANGO_SUPERUSER_USERNAME=admin
DJANGO_SUPERUSER_PASSWORD=你的强密码
```

**然后再点一次 Deploy，搞定。**

---

## 自动更新

每次 `git push` 到 `main` 分支，Railway 自动：

1. 检测新提交 → 重建 Docker 镜像
2. 启动容器 → 自动执行 `migrate` → `collectstatic` → `gunicorn`

```
本地改代码 → git push → 自动部署 🚀
```

---

## 原理

Railway 的 PostgreSQL 服务会自动注入 `PGHOST`/`PGDATABASE`/`PGUSER`/`PGPASSWORD`/`PGPORT`。
`settings.py` 已修改为优先读 `POSTGRES_*`，fallback 到 Railway 的 `PG*`，实现零配置。

## 注意事项

- 上传文件（media 目录）在 Railway 默认磁盘上不持久，建议使用 [Railway Volumes](https://docs.railway.com/data/volumes)
- 生产环境始终用 PostgreSQL（`DB_ENGINE` 默认为 postgres，不做任何改动）
- 绑定自定义域名后，记得更新 `DJANGO_ALLOWED_HOSTS`
