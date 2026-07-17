# Backend Maleu — Cómo se edita y se publica (LEER PRIMERO)

## Regla de oro: UNA sola fuente de verdad
Todo el backend (Apps Script) vive en **`.clasp-src/Code.js`** y en ningún otro lado.

> ⛔ Ya **NO existe** `apps-script.gs`. Era una fotocopia que había que sincronizar
> a mano y causaba que se chocaran códigos viejos. Se eliminó el 17/07/2026.
> No la vuelvas a crear. No hay mirror. Un archivo, y listo.

## Cómo publicar: UN solo comando
```bash
cd /c/Users/tadeu/estancias
# 1. Editar .clasp-src/Code.js
# 2. Publicar (hace check + verifica base + push + deploy + commit + push git):
./deploy.sh "v440: que cambiaste"
```
Eso es todo. El script `deploy.sh`:
1. Chequea sintaxis (`node --check`).
2. Corre `deploy_check.py`: **frena** si otra sesión publicó algo sin commitear
   (te avisa antes de pisarlo).
3. `clasp push -f` + `clasp deploy` al deployment de PRODUCCIÓN.
4. `git add/commit/push` — nunca queda "código suelto" para que otra sesión pise.

## Antes de EDITAR (si venís de una sesión vieja)
Si no estás seguro de tener la última base, traé lo publicado primero:
```bash
cd /c/Users/tadeu/estancias && clasp pull   # trae el código realmente vivo a .clasp-src/Code.js
```
Y recién ahí editás.

## Datos
- **Deployment de PRODUCCIÓN** (la URL que usan tienda/panel/ruta):
  `AKfycbxmrG5YVSshcYezk8lXFx_uxb7NFGcb9EfTXc7dsIN4rZyj73CET4mk_aKPFPDY2wNi`
- **Sólo lectura** del código vivo: service account (`maleu-service-account.json`).
- Los `.py` sueltos viejos (`deploy_script.py`, `deploy_deposito.py`, etc.) quedaron
  obsoletos: publicaban leyendo la fotocopia. Usar SIEMPRE `./deploy.sh`.
