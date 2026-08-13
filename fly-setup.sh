#!/bin/bash
# One-time Fly.io setup.
# Run once to provision everything, then use "fly deploy" for all future deploys.
#
# Prerequisites:
#   brew install flyctl   (or: curl -L https://fly.io/install.sh | sh)
#   fly auth login

set -euo pipefail

APP="basket-analytics"   # must match the app name in fly.toml; change if taken
REGION="fra"             # Frankfurt — nearest available Fly.io region to Poland

echo "==> Creating app..."
fly apps create "$APP" --machines

echo "==> Creating Postgres cluster..."
fly postgres create \
  --name "${APP}-db" \
  --region "$REGION" \
  --initial-cluster-size 1 \
  --vm-size shared-cpu-1x \
  --volume-size 1

echo "==> Attaching Postgres (sets DATABASE_URL secret automatically)..."
fly postgres attach "${APP}-db" --app "$APP"

echo "==> Creating persistent volume for match files..."
fly volumes create match_files --region "$REGION" --size 1 --app "$APP"

echo "==> Setting secrets..."
fly secrets set \
  SECRET_KEY="$(python3 -c 'import secrets; print(secrets.token_hex(32))')" \
  PORTAL_LOGIN="ZAWODNICY_GTK" \
  PORTAL_PASS="change-me-before-running" \
  --app "$APP"

echo ""
echo "Done. Edit PORTAL_LOGIN / PORTAL_PASS above, then run:"
echo "  fly deploy"
