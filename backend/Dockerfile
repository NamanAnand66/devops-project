FROM python:3.12-slim
WORKDIR /app
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt && useradd --create-home --uid 10001 appuser
COPY alembic.ini ./
COPY alembic ./alembic
COPY app ./app
COPY docker-entrypoint.sh /usr/local/bin/docker-entrypoint
RUN chmod 755 /usr/local/bin/docker-entrypoint
USER 10001
EXPOSE 8000
ENTRYPOINT ["/usr/local/bin/docker-entrypoint"]
