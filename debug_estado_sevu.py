"""Estado actual: qué filas Sevuchitas tienen Pagado=Sí + pagos en Egresos."""
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"

creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
sheets = build("sheets", "v4", credentials=creds).spreadsheets()

def parse_money(v):
    if v in (None, ""):
        return 0.0
    s = str(v).replace("$", "").replace(" ", "").strip()
    # En AR el "." es separador de miles, "," puede ser decimal. Quito puntos.
    s = s.replace(".", "")
    s = s.replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

# OC: filas Sevuchitas marcadas Pagado=Sí
print("=== OC Sevuchitas con Pagado Proveedor = Sí ===")
resp = sheets.values().get(spreadsheetId=SHEET_ID, range="Orden de Compra!A1:Y").execute()
data = resp.get("values", [])
total_pagadas = 0
filas_si = []
for r, row in enumerate(data[1:], start=2):
    while len(row) < 25:
        row.append("")
    if (row[9] or "").strip().lower() != "sevuchitas":
        continue
    if (row[23] or "").strip() not in ("Sí", "Si"):
        continue
    sem = row[2]
    estado = row[20]
    costo = parse_money(row[14])
    prod = row[10]
    qty = row[12]
    fecha_rec = row[22]
    print(f"  Fila {r:>3} | OC {row[0]:<7} | Sem {sem:<3} | Est {estado:<10} | Rec {fecha_rec:<12} | {prod[:30]:<30} {qty:>3}u ${costo:>10,.0f}")
    total_pagadas += costo
    filas_si.append(r)
print(f"  TOTAL marcado Pagado=Sí: ${total_pagadas:,.0f}")
print(f"  Filas: {filas_si}")

# Egresos: pagos a Sevuchitas
print("\n=== Egresos a Sevuchitas ===")
resp = sheets.values().get(spreadsheetId=SHEET_ID, range="Egresos!A1:H").execute()
data = resp.get("values", [])
header = data[0]
print(f"  Header: {header}")
for r, row in enumerate(data[1:], start=2):
    while len(row) < 8:
        row.append("")
    concepto = (row[4] or "")
    if "sevuchit" not in concepto.lower():
        continue
    print(f"  Fila {r:>3} | {row[0]:<12} | Sem {row[1]:<3} | {row[3]:<10} | {row[4][:35]:<35} | {row[5]:<14} | ${parse_money(row[6]):>10,.0f} | {row[7]}")
