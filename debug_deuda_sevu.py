"""Lista filas de la hoja Orden de Compra con Sevuchitas, Estado=Recibido, no pagadas."""
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"

creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
sheets = build("sheets", "v4", credentials=creds).spreadsheets()

resp = sheets.values().get(spreadsheetId=SHEET_ID, range="Orden de Compra!A1:Y").execute()
data = resp.get("values", [])
header = data[0]
print("Header cols (idx -> name):")
for i, h in enumerate(header):
    print(f"  {i}: {h}")
print()

print(f"{'Fila':>5} {'OC':<8} {'Sem':<6} {'Estado':<12} {'Pagado':<8} {'Origen':<14} {'Producto':<28} {'Cant':>5} {'CostoTot':>12}")
total = 0
total_w17 = 0
total_w18 = 0
for r, row in enumerate(data[1:], start=2):
    while len(row) < 25:
        row.append("")
    proveedor = (row[9] or "").strip()
    estado = (row[20] or "").strip()
    pagado = (row[23] or "").strip()
    origen = (row[19] or "").strip()
    if proveedor.lower() != "sevuchitas":
        continue
    if estado != "Recibido":
        continue
    if pagado in ("Sí", "Si"):
        continue
    if origen.startswith("Dep"):
        continue
    sem = (row[2] or "").strip()
    raw_costo = (row[14] or "0")
    costo = float(str(raw_costo).replace("$", "").replace(".", "").replace(",", ".").strip() or 0)
    prod = (row[10] or "").strip()
    qty = (row[12] or "")
    oc = (row[0] or "").strip()
    print(f"{r:>5} {oc:<8} {sem:<6} {estado:<12} {pagado:<8} {origen:<14} {prod[:28]:<28} {qty:>5} ${costo:>11,.0f}")
    total += costo
    if sem == "17":
        total_w17 += costo
    elif sem == "18":
        total_w18 += costo
print()
print(f"Total Sevuchitas pendiente: ${total:,.0f}")
print(f"  Semana 17: ${total_w17:,.0f}")
print(f"  Semana 18: ${total_w18:,.0f}")
