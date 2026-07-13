"""Agrega cols 'Estado Pago a Maleu' y 'Fecha Pago a Maleu' a hoja Red si no existen.
Default Estado Pago a Maleu = 'No' para todos los pedidos existentes."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA = r"C:\Users\tadeu\maleu-service-account.json"
SID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
c = Credentials.from_service_account_file(SA, scopes=["https://www.googleapis.com/auth/spreadsheets"])
svc = build("sheets", "v4", credentials=c).spreadsheets()

meta = svc.get(spreadsheetId=SID).execute()
sheet_id = None
cols_cur = None
for s in meta["sheets"]:
    if s["properties"]["title"] == "Red":
        sheet_id = s["properties"]["sheetId"]
        cols_cur = s["properties"]["gridProperties"]["columnCount"]
        break

hdr = svc.values().get(spreadsheetId=SID, range="'Red'!1:1").execute().get("values", [[]])[0]
existing = {str(h).strip() for h in hdr}

to_add = []
if "Estado Pago a Maleu" not in existing: to_add.append("Estado Pago a Maleu")
if "Fecha Pago a Maleu" not in existing: to_add.append("Fecha Pago a Maleu")

if not to_add:
    print("Ambas cols ya existen."); exit()

def col_letter(n):
    s=""
    while n>0: n,r=divmod(n-1,26); s=chr(65+r)+s
    return s

needed = len(hdr) + len(to_add)
if needed > cols_cur:
    svc.batchUpdate(spreadsheetId=SID, body={"requests":[{"appendDimension":{"sheetId":sheet_id,"dimension":"COLUMNS","length":needed - cols_cur}}]}).execute()

for i, name in enumerate(to_add):
    col = len(hdr) + 1 + i
    svc.values().update(spreadsheetId=SID, range=f"'Red'!{col_letter(col)}1", valueInputOption="RAW", body={"values": [[name]]}).execute()
    print(f"  + {name} en col {col_letter(col)}")
    # Formato header
    svc.batchUpdate(spreadsheetId=SID, body={"requests": [{
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": col-1, "endColumnIndex": col},
            "cell": {"userEnteredFormat": {"backgroundColor": {"red": 0.365, "green": 0.263, "blue": 0.216}, "textFormat": {"foregroundColor": {"red": 1, "green": 1, "blue": 1}, "bold": True}}},
            "fields": "userEnteredFormat(backgroundColor,textFormat)",
        }
    }]}).execute()

# Default "No" en col Estado Pago a Maleu para filas existentes
rows = svc.values().get(spreadsheetId=SID, range="'Red'!A:A").execute().get("values", [])
n_rows = len(rows) - 1  # excluye header
if n_rows > 0 and "Estado Pago a Maleu" in to_add:
    col_est = len(hdr) + 1 + to_add.index("Estado Pago a Maleu")
    col_letter_est = col_letter(col_est)
    svc.values().update(
        spreadsheetId=SID,
        range=f"'Red'!{col_letter_est}2:{col_letter_est}{n_rows+1}",
        valueInputOption="RAW",
        body={"values": [["No"] for _ in range(n_rows)]},
    ).execute()
    print(f"  Default 'No' en {n_rows} filas")

print("Listo.")
