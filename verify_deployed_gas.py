"""Read the currently deployed Apps Script code from the API and confirm no 'Depósito'."""
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import google.auth
from googleapiclient.discovery import build

SCRIPT_ID = "1wtfgJFESRbD1llGX39zOhWIWm0v047CGWqdYXkHJIlS0cDt3Ove3cSza"

creds, project = google.auth.default(scopes=[
    "https://www.googleapis.com/auth/script.projects",
    "https://www.googleapis.com/auth/script.deployments",
])
service = build("script", "v1", credentials=creds)

content = service.projects().getContent(scriptId=SCRIPT_ID).execute()
for f in content.get("files", []):
    if f["type"] == "SERVER_JS":
        source = f["source"]
        print(f"Archivo: {f['name']} | {len(source)} chars")
        print(f"  'Depósito' (con acento): {source.count('Depósito')}")
        print(f"  'Deposito' (sin acento): {source.count('Deposito')}")

deployments = service.projects().deployments().list(scriptId=SCRIPT_ID).execute()
print("\nDeployments actuales:")
for d in deployments.get("deployments", []):
    cfg = d.get("deploymentConfig", {})
    print(f"  {d['deploymentId'][:30]}... | ver={cfg.get('versionNumber', 'HEAD')} | desc={cfg.get('description', '')}")
