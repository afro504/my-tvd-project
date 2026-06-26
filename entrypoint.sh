#!/bin/sh

#!/bin/sh
set -e

echo "🚀 Initialisation du conteneur Django..."
# Attente que SQL Server accepte les connexions
until /opt/mssql-tools/bin/sqlcmd -S db -U sa -P "Reine2009#" -Q "SELECT 1" > /dev/null 2>&1; do
  >&2 echo "SQL Server n'est pas encore prêt..."
  sleep 5
done

end_time=$(date +%s)
elapsed=$((end_time - start_time))

echo "✅ SQL Server est prêt (attente totale : ${elapsed}s)"

# Vérifier si la base existe, sinon la créer
echo "📂 Vérification de la base TVDdata..."
/opt/mssql-tools/bin/sqlcmd -S db -U sa -P "Reine2009#" -Q "IF DB_ID('TVDdata') IS NULL CREATE DATABASE TVDdata;"

# Appliquer les migrations
echo "📦 Application des migrations..."
python manage.py migrate --noinput


# Collecter les fichiers statiques (utile en prod)
echo "🎨 Collecte des fichiers statiques..."
python manage.py collectstatic --noinput

# Créer le superutilisateur si les variables sont définies
if [ -n "$DJANGO_SUPERUSER_USERNAME" ] && [ -n "$DJANGO_SUPERUSER_EMAIL" ] && [ -n "$DJANGO_SUPERUSER_PASSWORD" ]; then
  echo "👤 Création du superutilisateur..."
  python manage.py createsuperuser \
    --noinput \
    --username "$DJANGO_SUPERUSER_USERNAME" \
    --email "$DJANGO_SUPERUSER_EMAIL" || true
  # ⚠️ Django ne permet pas de passer le mot de passe en argument, il doit être défini via variable d’environnement
  # django-mssql-backend utilisera DJANGO_SUPERUSER_PASSWORD automatiquement
fi

# Lancer le serveur (Gunicorn en prod, runserver en dev)
if [ "$DJANGO_ENV" = "prod" ]; then
  echo "🚀 Lancement de Gunicorn..."
  exec gunicorn TVDdata.wsgi:application --bind 0.0.0.0:8080
else
  echo "🚀 Lancement du serveur Django (dev)..."
  exec python manage.py runserver 0.0.0.0:8000
fi