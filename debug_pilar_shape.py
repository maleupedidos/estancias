"""Verificar la estructura de Pilar y fecha/estado de Gloria en detalle."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

# Metadata
meta = svc.get(spreadsheetId=SID, fields="sheets(properties(title,gridProperties))").execute()
for sh in meta["sheets"]:
    name = sh["properties"]["title"]
    if name in ("Home", "Pilar", "Clubes", "Red"):
        gp = sh["properties"]["gridProperties"]
        print(f"{name:10} rows={gp.get('rowCount'):5} cols={gp.get('columnCount'):4}")

# Headers de Pilar
hdr = svc.values().get(spreadsheetId=SID, range="'Pilar'!1:1").execute().get("values", [[]])[0]
print(f"\nPilar tiene {len(hdr)} headers:")
for i, h in enumerate(hdr, 1):
    print(f"  col {i:3} ({chr(64+i) if i<=26 else 'A'+chr(64+i-26)}): {h!r}")

# Gloria: fila 10 en Pilar - dump completo
print("\n=== Gloria fila 10 completa ===")
r = svc.values().get(spreadsheetId=SID, range="'Pilar'!A10:BD10",
                     valueRenderOption="FORMATTED_VALUE").execute().get("values", [[]])[0]
for i, v in enumerate(r, 1):
    if v:
        col = chr(64+i) if i<=26 else 'A'+chr(64+i-27)
        hdrName = hdr[i-1] if i-1 < len(hdr) else "?"
        print(f"  col {i:3} ({col:3}) {hdrName[:25]:25} = {v!r}")
