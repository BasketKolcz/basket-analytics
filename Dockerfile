FROM python:3.12-slim

# Ustaw katalog roboczy
WORKDIR /app

# Skopiuj plik aplikacji i zasoby statyczne
COPY app.py .
COPY static/ /app/static/

# Zainstaluj zależności Pythona
RUN pip install --no-cache-dir \
    flask \
    psycopg2-binary \
    openpyxl \
    gunicorn \
    playwright \
    anthropic

# Zainstaluj system deps wymagane przez Chromium + pobierz binarkę przeglądarki
# (playwright install-deps używa apt-get, dlatego musimy mieć go w slim)
RUN apt-get update && apt-get install -y --no-install-recommends \
        wget gnupg ca-certificates fonts-liberation \
    && playwright install --with-deps chromium \
    && rm -rf /var/lib/apt/lists/*

# Utwórz katalog na pliki tymczasowe
RUN mkdir -p /tmp/basketkolcz_pending

# Port aplikacji
ENV PORT=5000

# Uruchom przez gunicorn (workers=2 — jeden worker obsługuje request użytkownika,
# drugi obsługuje wewnętrzny request od Playwright Chromium do tej samej aplikacji)
CMD gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 180
