# Costes Artículo — imagen con ODBC Driver 18 for SQL Server (necesario para pyodbc)
FROM python:3.12-slim

# --- Microsoft ODBC Driver 18 (Debian 12 / bookworm) ---
RUN apt-get update \
 && apt-get install -y --no-install-recommends \
        curl gnupg ca-certificates apt-transport-https unixodbc-dev \
 && curl -sSL https://packages.microsoft.com/keys/microsoft.asc \
        | gpg --dearmor -o /usr/share/keyrings/microsoft-prod.gpg \
 && echo "deb [signed-by=/usr/share/keyrings/microsoft-prod.gpg] https://packages.microsoft.com/debian/12/prod bookworm main" \
        > /etc/apt/sources.list.d/mssql-release.list \
 && apt-get update \
 && ACCEPT_EULA=Y apt-get install -y --no-install-recommends msodbcsql18 \
 && apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 5000

# Servidor WSGI de producción (waitress). Las credenciales se pasan por variables
# de entorno en tiempo de ejecución (--env-file .env), NUNCA se copian a la imagen.
CMD ["waitress-serve", "--host=0.0.0.0", "--port=5000", "app:app"]
