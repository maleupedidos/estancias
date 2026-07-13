"""Dump headers exactos para diseñar generarResumenSemanal.
Hojas que necesito: Home, Pilar, Clubes, Red, Egresos, Ingresos, Productos.
"""
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

meta = svc.get(spreadsheetId=SPREADSHEET_ID).execute()
all_sheets = [s["properties"]["title"] for s in meta["sheets"]]
print("HOJAS:", all_sheets)
print("=" * 80)

for hoja in ["Home", "Pilar", "Clubes", "Red", "Egresos", "Ingresos", "Productos", "Saldo Base", "Pagos Red Liq"]:
    if hoja not in all_sheets:
        print(f"\n!!! NO EXISTE: {hoja}")
        continue
    r = svc.values().get(spreadsheetId=SPREADSHEET_ID, range=f"'{hoja}'!1:1").execute()
    headers = r.get("values", [[]])[0] if r.get("values") else []
    print(f"\n=== {hoja} ({len(headers)} cols) ===")
    for i, h in enumerate(headers, 1):
        print(f"  {col_letter(i):>3} ({i:>2}): {h}")
