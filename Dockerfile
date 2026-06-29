FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY lustia_api.py .

RUN mkdir -p /app/data

EXPOSE 8000

CMD uvicorn lustia_api:app --host 0.0.0.0 --port ${PORT:-8000}
