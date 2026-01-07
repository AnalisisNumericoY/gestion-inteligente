# Dockerfile para Frontera App con R + Python
FROM rocker/r-ver:4.3.0

# Instalar Python y dependencias del sistema
RUN apt-get update && \
    apt-get install -y \
    python3 \
    python3-pip \
    python3-venv \
    libcurl4-openssl-dev \
    libssl-dev \
    libxml2-dev \
    libfontconfig1-dev \
    libharfbuzz-dev \
    libfribidi-dev \
    libfreetype6-dev \
    libpng-dev \
    libtiff5-dev \
    libjpeg-dev \
    && rm -rf /var/lib/apt/lists/*

# Instalar paquetes R necesarios
RUN R -e "install.packages(c('openxlsx', 'lightgbm', 'here', 'data.table', 'dplyr'), repos='https://cloud.r-project.org/')"

# Crear directorios de trabajo
RUN mkdir -p /app /content/input /content/src /content/output /frontend-simple

# Copiar requirements.txt primero (para cache)
COPY requirements.txt /app/requirements.txt

# Copiar scripts R y archivos de datos
COPY content/ /content/

# Copiar código Python
COPY app/ /app/

# Instalar dependencias Python
RUN pip3 install --no-cache-dir -r /app/requirements.txt

# Exponer puerto
EXPOSE 8000

# Comando por defecto (será sobreescrito por docker-compose en desarrollo)
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]