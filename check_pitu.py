"""Diagnóstico fila 264 Home (pedido Pitu) — descuento 10% por +$100k no aplicado."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SPREADSHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

# Headers
headers = svc.values().get(spreadsheetId=SPREADSHEET_ID, range="'Home'!A1:BA1").execute().get("values", [[]])[0]
# Row 264
row = svc.values().get(spreadsheetId=SPREADSHEET_ID, range="'Home'!A264:BA264").execute().get("values", [[]])[0]

print("=== HOME fila 264 ===")
def col_letter(i):
    return chr(65+i) if i<26 else 'A'+chr(65+i-26)
for i, h in enumerate(headers):
    val = row[i] if i < len(row) else ""
    if val not in ("", 0, "0"):
        print(f"  {col_letter(i)} ({i+1}) {h}: {val}")

print()
print("=== Resumen monetario ===")
def get(name):
    try:
        idx = [h.strip().lower() for h in headers].index(name.lower())
        return row[idx] if idx < len(row) else ""
    except ValueError:
        return "(no col)"
print(f"  Subtotal sin Descuento: {get('Subtotal sin Descuento')}")
print(f"  Descuento:              {get('Descuento')}")
print(f"  Total ($):              {get('Total ($)')}")
print(f"  Envio:                  {get('Envio')}")
print(f"  Efectivo:               {get('Efectivo')}")
print(f"  Transferencia:          {get('Transferencia')}")
print(f"  Forma de Pago:          {get('Forma de Pago')}")
print(f"  Facturado:              {get('Facturado')}")
