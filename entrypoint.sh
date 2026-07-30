#!/bin/sh
set -e

echo "=== DEBUG: Database-related environment variables ==="
env | grep -iE 'PG|POSTGRES|DATABASE' | sort || true
echo "=== DEBUG: End of env dump ==="

python manage.py migrate --noinput
python manage.py ensure_admin
python manage.py collectstatic --noinput
exec gunicorn cost_platform.wsgi:application \
  --bind 0.0.0.0:8000 \
  --workers "${GUNICORN_WORKERS:-3}" \
  --timeout "${GUNICORN_TIMEOUT:-120}" \
  --access-logfile -
