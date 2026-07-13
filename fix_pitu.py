"""Corrige fila 264 Home (Pitu): aplicar 10% OFF retroactivo.
Subtotal $166.000 - Descuento $16.600 = Total $149.400 (lo que cobró Tadeo)."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SPREADSHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

# Layout Home v2 — fila 264:
#   N(14) Subtotal Producto = 166000  (no cambia)
#   O(15) Envío             = 0       (no cambia)
#   P(16) Descuento         = 0  → 16600
#   Q(17) Total a cobrar    = 166000 → 149400
#   R(18) Efectivo          = 0       (no cambia)
#   S(19) Transferencia     = 166000 → 149400
#   V(22) Facturado         = fórmula =Q+T+U → recalcula sola

updates = [
    {"range": "'Home'!P264", "values": [[16600]]},
    {"range": "'Home'!Q264", "values": [[149400]]},
    {"range": "'Home'!S264", "values": [[149400]]},
]

result = svc.values().batchUpdate(
    spreadsheetId=SPREADSHEET_ID,
    body={"valueInputOption": "USER_ENTERED", "data": updates},
).execute()
print("Updated cells:", result.get("totalUpdatedCells"))

# Verificar
row = svc.values().get(spreadsheetId=SPREADSHEET_ID, range="'Home'!N264:V264").execute().get("values", [[]])[0]
labels = ["Subtotal", "Envio", "Descuento", "Total", "Efectivo", "Transferencia", "Propina Ef", "Propina Tr", "Facturado"]
print()
print("=== Fila 264 Pitu (verificacion) ===")
for lbl, val in zip(labels, row):
    print(f"  {lbl:<15}: {val}")
