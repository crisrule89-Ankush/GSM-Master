FROM python:3.12-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

# Install Microsoft SQL Server ODBC Driver 18 required by pyodbc.
RUN apt-get update \
    && apt-get install -y --no-install-recommends curl ca-certificates gnupg unixodbc-dev libgssapi-krb5-2 \
    && curl -sSL -o /tmp/packages-microsoft-prod.deb https://packages.microsoft.com/config/debian/12/packages-microsoft-prod.deb \
    && dpkg -i /tmp/packages-microsoft-prod.deb \
    && rm /tmp/packages-microsoft-prod.deb \
    && apt-get update \
    && ACCEPT_EULA=Y apt-get install -y --no-install-recommends msodbcsql18 \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY . ./

# Render supplies PORT at runtime.
CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:${PORT:-10000} --workers 2 --threads 4 --timeout 120 app:app"]