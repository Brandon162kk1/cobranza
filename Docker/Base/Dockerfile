FROM python:3.11-slim

WORKDIR /app

# Variables
ENV PYTHONUNBUFFERED=1
ENV PYTHONPATH=/app/Codigo

# Dependencias Linux comunes
RUN apt-get update && apt-get install -y \
    curl \
    unzip \
    wget \
    xvfb \
    && rm -rf /var/lib/apt/lists/*

# Copiar requirements primero
COPY requirements.txt .

# Instalar dependencias Python
RUN pip install --no-cache-dir -r requirements.txt

# Copiar proyecto
COPY Codigo /app/Codigo

# Carpetas comunes
RUN mkdir -p /app/Downloads

USER user1