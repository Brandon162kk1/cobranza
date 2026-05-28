# Usar tu imagen base personalizada
FROM chromedriver:stable

ENV PYTHONUNBUFFERED=1

COPY requirements.txt .

USER root

RUN pip install --no-cache-dir -r requirements.txt

COPY Supervisord /app/Supervisord

USER user1