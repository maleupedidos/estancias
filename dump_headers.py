"""Dump headers (1-based index) de Home, Pilar, Clubes."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SPREADSHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

def col_letter(n):
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s

# List all sheets
meta = svc.get(spreadsheetId=SPREADSHEET_ID).execute()
all_sheets = [s["properties"]["title"] for s in meta["sheets"]]
print("HOJAS EXISTENTES:", all_sheets)
print()

for hoja in ["Home", "Pilar", "Clubes"]:
    if hoja not in all_sheets:
        print(f"\n!!! {hoja} no existe")
        continue
    r = svc.values().get(spreadsheetId=SPREADSHEET_ID, range=f"'{hoja}'!1:1").execute()
    headers = r.get("values", [[]])[0] if r.get("values") else []
    print(f"=== {hoja} ({len(headers)} columnas) ===")
    for i, h in enumerate(headers, 1):
        print(f"  {col_letter(i):>3} ({i:>2}): {h}")
    print()

    # sample row 2
    r2 = svc.values().get(spreadsheetId=SPREADSHEET_ID, range=f"'{hoja}'!2:2").execute()
    row2 = r2.get("values", [[]])[0] if r2.get("values") else []
    if row2:
        print(f"  -- fila 2 (sample) --")
        for i, v in enumerate(row2, 1):
            print(f"  {col_letter(i):>3} ({i:>2}): {v!r}")
    print()
