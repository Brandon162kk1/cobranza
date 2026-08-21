# Usar tu imagen base personalizada
FROM chromedriver:stable

# Hace que los logs de Python se envíen directamente a la consola en tiempo real sin almacenarse en bufer
ENV PYTHONUNBUFFERED=1

# Es el directorio raíz para buscar módulos y paquetes
ENV PYTHONPATH=/app

# Copiar requirements.txt
COPY requirements.txt .

# Instalar dependencias Python
RUN pip install --no-cache-dir -r requirements.txt