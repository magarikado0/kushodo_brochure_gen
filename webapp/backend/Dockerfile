FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

COPY webapp/backend/requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

COPY tools /repo/tools
COPY templates /repo/templates
COPY webapp/backend/app /app/app

ENV PYTHONPATH=/app
ENV KUSHODO_GENERATOR_PATH=/repo/tools/generate.py
ENV KUSHODO_TEMPLATE_PATH=/repo/templates/パンフ鋳型.docx

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
