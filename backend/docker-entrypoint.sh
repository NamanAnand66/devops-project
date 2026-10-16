#!/bin/sh
# Compose waits for Postgres to be healthy. Retrying here also makes startup
# robust against a short connection gap while the database is becoming ready.
set -eu

max_attempts="${DB_MIGRATION_MAX_ATTEMPTS:-30}"
retry_delay="${DB_MIGRATION_RETRY_DELAY_SECONDS:-2}"
attempt=1

until alembic upgrade head; do
  if [ "$attempt" -ge "$max_attempts" ]; then
    echo "Database migrations failed after ${max_attempts} attempts." >&2
    exit 1
  fi

  echo "Database is not ready; retrying migrations in ${retry_delay}s (attempt ${attempt}/${max_attempts})..." >&2
  attempt=$((attempt + 1))
  sleep "$retry_delay"
done

exec uvicorn app.main:app --host 0.0.0.0 --port 8000
