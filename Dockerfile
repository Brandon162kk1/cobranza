# Usar tu imagen base personalizada
FROM chromedriver:stable

WORKDIR /app

# Copiar solo dependencias primero (mejor cache)
COPY requirements.txt .

# Volver temporalmente a root para instalar dependencias
USER root
RUN pip install --no-cache-dir -r requirements.txt

# Crear carpetas necesarias
RUN mkdir -p /app/Downloads \
    && mkdir -p /codigo_mapfre \
    && chown -R user1:user1 /app /codigo_mapfre

# Volver a usuario sin privilegios
USER user1

# Variables de entorno
ENV PYTHONUNBUFFERED=1
ENV PYTHONPATH=/app/Codigo