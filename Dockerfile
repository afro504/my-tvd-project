# =========================
# Base image commune
# =========================
FROM python:3.14-slim AS base

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PATH="/opt/mssql-tools/bin:${PATH}"

WORKDIR /app

# Installer dépendances système
RUN apt-get update && apt-get install -y \
    curl gnupg apt-transport-https \
    unixodbc unixodbc-dev \
    gcc g++ make \
    && rm -rf /var/lib/apt/lists/*

# Ajouter la clé Microsoft
RUN curl https://packages.microsoft.com/keys/microsoft.asc | gpg --dearmor > /usr/share/keyrings/microsoft.gpg

# Ajouter le repo Microsoft (Debian 11 bullseye)
RUN echo "deb [arch=amd64 signed-by=/usr/share/keyrings/microsoft.gpg] https://packages.microsoft.com/debian/11/prod bullseye main" \
    > /etc/apt/sources.list.d/mssql-release.list

# Installer le driver ODBC 17 et outils SQL
RUN apt-get update && ACCEPT_EULA=Y apt-get install -y msodbcsql17 mssql-tools \
    && rm -rf /var/lib/apt/lists/*

# Ajouter sqlcmd au PATH
ENV PATH="$PATH:/opt/mssql-tools/bin"

# Créer le dossier staticfiles
RUN mkdir -p /app/staticfiles



# Copier requirements
COPY requirements.txt .

# =========================
# Développement
# =========================
FROM base AS development

RUN pip install --upgrade pip && pip install -r requirements.txt \
    && pip install debugpy ipython django-debug-toolbar

COPY . .



COPY entrypoint.sh /app/entrypoint.sh
RUN chmod +x /app/entrypoint.sh

# Collecter les fichiers statiques au build
RUN python manage.py collectstatic --noinput

EXPOSE 8000
ENTRYPOINT ["/app/entrypoint.sh"]
CMD ["python", "manage.py", "runserver", "0.0.0.0:8000"]

# =========================
# Production
# =========================
FROM base AS production

RUN pip install --upgrade pip && pip install -r requirements.txt

COPY . .

COPY entrypoint.sh /app/entrypoint.sh
RUN chmod +x /app/entrypoint.sh

# Collecter les fichiers statiques au build
RUN python manage.py collectstatic --noinput

EXPOSE 8080
ENTRYPOINT ["/app/entrypoint.sh"]
CMD ["gunicorn", "TVDdata.wsgi:application", "--bind", "0.0.0.0:8080", "--workers=4"]

