"""
deploy_check.py — Guarda contra "codigos viejos que se chocan".

Compara el codigo Apps Script REALMENTE PUBLICADO (live) contra lo que tiene
git en HEAD para .clasp-src/Code.js.

- Si son IGUALES  -> exit 0 (base sana, seguro deployar).
- Si DIFIEREN     -> exit 2 (ALGUIEN publico sin commitear en una sesion
                     paralela; si publicas ahora le pisas ese cambio).

No usa OAuth: lee con el service account (permiso de LECTURA en Apps Script).
Lo invoca deploy.sh antes de publicar. No modifica nada.
"""
import subprocess
import sys

from google.oauth2 import service_account
from googleapiclient.discovery import build

SCRIPT_ID = "1wtfgJFESRbD1llGX39zOhWIWm0v047CGWqdYXkHJIlS0cDt3Ove3cSza"
SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"


def live_source():
    creds = service_account.Credentials.from_service_account_file(
        SA_KEY, scopes=["https://www.googleapis.com/auth/script.projects"]
    )
    svc = build("script", "v1", credentials=creds)
    content = svc.projects().getContent(scriptId=SCRIPT_ID).execute()
    for f in content.get("files", []):
        if f["name"] == "Code" and f["type"] == "SERVER_JS":
            return f["source"]
    raise SystemExit("No se encontro el archivo 'Code' en el proyecto Apps Script.")


def head_source():
    # Codigo commiteado en git HEAD (el ultimo deploy que quedo guardado).
    out = subprocess.run(
        ["git", "show", "HEAD:.clasp-src/Code.js"],
        capture_output=True, cwd=r"C:\Users\tadeu\estancias",
    )
    if out.returncode != 0:
        raise SystemExit("No pude leer HEAD:.clasp-src/Code.js — " + out.stderr.decode("utf-8", "replace").strip())
    return out.stdout.decode("utf-8", "replace")


def norm(s):
    # Normaliza fin de linea para no falsear el diff por CRLF/LF.
    return s.replace("\r\n", "\n").replace("\r", "\n")


def main():
    live = norm(live_source())
    head = norm(head_source())
    if live == head:
        print("OK: el codigo publicado coincide con git HEAD. Base sana.")
        sys.exit(0)

    # Diferencia -> reportar para que Claude reconcilie antes de pisar.
    ll, hl = live.split("\n"), head.split("\n")
    print("ALERTA: el codigo PUBLICADO difiere de git HEAD.")
    print("  -> Otra sesion publico cambios que NO estan commiteados en git.")
    print("  -> Si deployas ahora, esos cambios vivos se PISAN.")
    print(f"  live={len(ll)} lineas  |  HEAD={len(hl)} lineas")
    print("  Reconciliar: traer lo vivo a git ANTES de deployar:")
    print("     cd /c/Users/tadeu/estancias && clasp pull && cp .clasp-src/Code.js .clasp-src/Code.js")
    print("     git add .clasp-src/Code.js && git commit -m 'sync: traer deploy vivo a git'")
    print("     y recien ahi re-aplicar tu cambio sobre esa base.")
    sys.exit(2)


if __name__ == "__main__":
    main()
