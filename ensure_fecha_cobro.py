"""Crear col 'Fecha de Cobro' en Home/Pilar/Capital Federal/Clubes si no existe."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SPREADSHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"

creds = Credentials.from_service_account_file(
    SA_KEY,
    scopes=["https://www.googleapis.com/auth/spreadsheets"],
)
svc = build("sheets", "v4", credentials=creds).spreadsheets()

HOJAS = ["Home", "Pilar", "Capital Federal", "Clubes"]
HEADER_NEW = "Fecha de Cobro"

# Obtener metadata con sheet IDs y gridProperties
meta = svc.get(spreadsheetId=SPREADSHEET_ID).execute()
sheet_info = {}
for s in meta["sheets"]:
    p = s["properties"]
    sheet_info[p["title"]] = {
        "sheetId": p["sheetId"],
        "cols": p["gridProperties"]["columnCount"],
        "rows": p["gridProperties"]["rowCount"],
    }

def col_letter(n):
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s

for hoja in HOJAS:
    if hoja not in sheet_info:
        print(f"\n--- {hoja}: no existe, salteo ---")
        continue
    # Leer headers
    r = svc.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{hoja}'!1:1",
    ).execute()
    headers = r.get("values", [[]])[0] if r.get("values") else []
    print(f"\n--- {hoja} (header fila: {len(headers)} cols, grid: {sheet_info[hoja]['cols']} cols) ---")
    if HEADER_NEW in [str(h).strip() for h in headers]:
        idx = [str(h).strip() for h in headers].index(HEADER_NEW) + 1
        print(f"  Ya existe en col {idx} ({col_letter(idx)})")
        continue
    new_col = len(headers) + 1
    # Expandir grid si hace falta
    if new_col > sheet_info[hoja]["cols"]:
        delta = new_col - sheet_info[hoja]["cols"]
        svc.batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": [{
                "appendDimension": {
                    "sheetId": sheet_info[hoja]["sheetId"],
                    "dimension": "COLUMNS",
                    "length": delta,
                }
            }]},
        ).execute()
        print(f"  Grid expandido +{delta} cols")
    col = col_letter(new_col)
    svc.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{hoja}'!{col}1",
        valueInputOption="RAW",
        body={"values": [[HEADER_NEW]]},
    ).execute()
    # Formato del header (marron + blanco + bold)
    svc.batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{
            "repeatCell": {
                "range": {
                    "sheetId": sheet_info[hoja]["sheetId"],
                    "startRowIndex": 0, "endRowIndex": 1,
                    "startColumnIndex": new_col - 1, "endColumnIndex": new_col,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.365, "green": 0.263, "blue": 0.216},
                        "textFormat": {
                            "foregroundColor": {"red": 1, "green": 1, "blue": 1},
                            "bold": True,
                        },
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat)",
            }
        }]},
    ).execute()
    print(f"  Creada en col {col} (#{new_col})")

print("\nListo.")
