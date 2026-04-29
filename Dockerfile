FROM python:3.12-slim

# Ustaw katalog roboczy
WORKDIR /app

# Skopiuj plik aplikacji i zasoby statyczne
COPY app.py .
COPY static/ /app/static/

# Zainstaluj zależności
RUN pip install --no-cache-dir \
    flask \
    psycopg2-binary \
    openpyxl \
    gunicorn

# Utwórz katalog na pliki tymczasowe
RUN mkdir -p /tmp/basketkolcz_pending

# Port aplikacji
ENV PORT=5000

# Uruchom przez gunicorn
CMD gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --timeout 120
