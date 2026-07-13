"""Ejecutar setupProductosFormulas() en Apps Script."""
import google.auth
from googleapiclient.discovery import build

SCRIPT_ID = "1wtfgJFESRbD1llGX39zOhWIWm0v047CGWqdYXkHJIlS0cDt3Ove3cSza"

creds, _ = google.auth.default(scopes=[
    "https://www.googleapis.com/auth/script.projects",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/script.scriptapp",
])
svc = build("script", "v1", credentials=creds)

print("Ejecutando setupProductosFormulas()...")
resp = svc.scripts().run(
    scriptId=SCRIPT_ID,
    body={"function": "setupProductosFormulas", "devMode": True},
).execute()

if "error" in resp:
    print("ERROR:", resp["error"])
else:
    print("OK — fórmulas regeneradas.")
    print(resp.get("response", {}))
