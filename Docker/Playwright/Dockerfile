FROM mcr.microsoft.com/playwright/python:v1.60.0-jammy

WORKDIR /app

ENV PYTHONPATH=/app

COPY requirements.txt .

RUN apt-get update && \
    apt-get install -y tesseract-ocr

RUN pip install --no-cache-dir -r requirements.txt

COPY . .