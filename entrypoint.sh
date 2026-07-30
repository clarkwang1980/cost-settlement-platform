#!/bin/sh
set -e

python manage.py migrate --noinput
python manage.py ensure_admin
python manage.py collectstatic --noinput
exec gunicorn cost_platform.wsgi:application \
  --bind 0.0.0.0:8000 \
  --workers "${GUNICORN_WORKERS:-3}" \
  --timeout "${GUNICORN_TIMEOUT:-120}" \
  --access-logfile -
