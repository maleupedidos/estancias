import json
import google.auth
from googleapiclient.discovery import build

SCRIPT_ID = "1wtfgJFESRbD1llGX39zOhWIWm0v047CGWqdYXkHJIlS0cDt3Ove3cSza"
VERSION_DESC = "v59 - cobrosPendientes devuelve r (fila); marcarEntregado fallback busca ultimo match"

# Auth
creds, project = google.auth.default(scopes=[
    "https://www.googleapis.com/auth/script.projects",
    "https://www.googleapis.com/auth/script.deployments",
])
service = build("script", "v1", credentials=creds)

# 1. Read current project to get appsscript.json
print("Reading current project content...")
content = service.projects().getContent(scriptId=SCRIPT_ID).execute()
files = content.get("files", [])

# Find appsscript.json
appscript_json = None
for f in files:
    if f["name"] == "appsscript" and f["type"] == "JSON":
        appscript_json = f
        print(f"  Found appsscript.json")
    else:
        print(f"  Found file: {f['name']} ({f['type']})")

if not appscript_json:
    raise Exception("Could not find appsscript.json in project")

# 2. Read the new .gs file
with open(r"C:\Users\tadeu\estancias\apps-script.gs", "r", encoding="utf-8") as fh:
    new_code = fh.read()

print(f"\nNew code: {len(new_code)} chars, {new_code.count(chr(10))} lines")

# 3. Update project with new code + existing appsscript.json
print("\nUpdating project content...")
body = {
    "files": [
        appscript_json,
        {
            "name": "Código",
            "type": "SERVER_JS",
            "source": new_code,
        }
    ]
}
service.projects().updateContent(scriptId=SCRIPT_ID, body=body).execute()
print("  Project updated!")

# 4. Create new version
print(f"\nCreating version: {VERSION_DESC}")
version = service.projects().versions().create(
    scriptId=SCRIPT_ID,
    body={"description": VERSION_DESC}
).execute()
version_number = version["versionNumber"]
print(f"  Created version #{version_number}")

# 5. List deployments and find web app deployment
print("\nListing deployments...")
deployments = service.projects().deployments().list(scriptId=SCRIPT_ID).execute()
web_deploy = None
for d in deployments.get("deployments", []):
    deploy_id = d["deploymentId"]
    config = d.get("deploymentConfig", {})
    desc = config.get("description", "")
    ver = config.get("versionNumber", "HEAD")
    print(f"  {deploy_id[:30]}... | ver={ver} | desc={desc}")
    # HEAD deployment has no versionNumber or versionNumber=0
    if config.get("versionNumber") and config["versionNumber"] > 0:
        web_deploy = d

if not web_deploy:
    raise Exception("No web app deployment found (non-HEAD)")

deploy_id = web_deploy["deploymentId"]
print(f"\nUpdating deployment {deploy_id[:30]}... to version #{version_number}")

service.projects().deployments().update(
    scriptId=SCRIPT_ID,
    deploymentId=deploy_id,
    body={
        "deploymentConfig": {
            "versionNumber": version_number,
            "description": VERSION_DESC,
        }
    }
).execute()
print(f"  Deployment updated to version #{version_number}!")
print("\nDone!")
