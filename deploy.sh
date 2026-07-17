#!/usr/bin/env bash
#
# deploy.sh — UNICO comando para publicar el backend de Maleu (Apps Script).
#
#   Uso:  ./deploy.sh "v440: que cambiaste"
#
# Fuente UNICA de verdad: .clasp-src/Code.js   (NO existe mas apps-script.gs)
#
# Hace TODO el ciclo seguro, sin pasos manuales:
#   1. Chequea sintaxis.
#   2. Verifica que nadie haya publicado sin commitear (deploy_check.py).
#   3. Publica con clasp (push + deploy al deployment de PRODUCCION).
#   4. Commitea y pushea a git  -> nunca queda "codigo suelto" que otra sesion pise.
#
set -euo pipefail

DESC="${1:-}"
if [ -z "$DESC" ]; then
  echo "ERROR: falta la descripcion."
  echo 'Uso:  ./deploy.sh "v440: que cambiaste"'
  exit 1
fi

cd "$(dirname "$0")"
SRC=".clasp-src/Code.js"
PROD_DEPLOY="AKfycbxmrG5YVSshcYezk8lXFx_uxb7NFGcb9EfTXc7dsIN4rZyj73CET4mk_aKPFPDY2wNi"

echo "→ [1/4] Sintaxis…"
node --check "$SRC"

echo "→ [2/4] Verificando que la base no este pisada por otra sesion…"
python deploy_check.py

echo "→ [3/4] Publicando (clasp push + deploy)…"
clasp push -f
clasp deploy -i "$PROD_DEPLOY" -d "$DESC"

echo "→ [4/4] Guardando en git (commit + push)…"
git add "$SRC" .clasp-src/appsscript.json
if git diff --cached --quiet; then
  echo "  (sin cambios para commitear)"
else
  git commit -m "$DESC

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
  git push || echo "  (push a git fallo — NO critico: el deploy ya esta vivo; pushear despues)"
fi

echo "✅ Deploy OK — $DESC"
echo "   Verificar en vivo si hace falta (esperar 1-2 min de propagacion GCP)."
