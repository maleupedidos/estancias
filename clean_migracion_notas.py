"""Limpia las notas que dicen '(migración ...)' en hoja Pagos Proveedores."""
import re
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SPREADSHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
SHEET = "Pagos Proveedores"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

# Headers
headers = svc.values().get(spreadsheetId=SPREADSHEET_ID, range=f"'{SHEET}'!1:1").execute().get("values", [[]])[0]
print("Headers:", headers)

# Find the notes column
notas_idx = None
for i, h in enumerate(headers):
    if "nota" in h.lower() or "concepto" in h.lower() or "detalle" in h.lower():
        notas_idx = i
        print(f"Columna notas detectada: {h} (índice {i}, col {chr(65+i)})")
        break
if notas_idx is None:
    print("ERROR: no encontre columna de notas. Frenando.")
    exit(1)

# Read all data
data = svc.values().get(spreadsheetId=SPREADSHEET_ID, range=f"'{SHEET}'!A:Z").execute().get("values", [])
print(f"Filas totales: {len(data)-1}")

# Patron: " (migración X)" o " (migracion X)" — case insensitive, con o sin espacio extra
pat = re.compile(r"\s*\(migraci[óo]n[^)]*\)\s*", re.IGNORECASE)

updates = []
for r_idx in range(1, len(data)):
    row = data[r_idx]
    if len(row) <= notas_idx:
        continue
    val = row[notas_idx]
    if not val:
        continue
    new_val = pat.sub("", val).strip()
    if new_val != val:
        cell = f"'{SHEET}'!{chr(65+notas_idx)}{r_idx+1}"
        updates.append({"range": cell, "values": [[new_val]]})
        print(f"  Fila {r_idx+1}: '{val}' -> '{new_val}'")

if not updates:
    print("Nada para limpiar.")
else:
    body = {"valueInputOption": "RAW", "data": updates}
    svc.values().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()
    print(f"\nOK {len(updates)} celdas actualizadas.")
