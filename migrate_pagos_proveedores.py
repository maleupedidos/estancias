"""
Migración a ledger Pagos Proveedores.
1. Crea hoja "Pagos Proveedores" si no existe.
2. Inserta pago de Sevuchitas del 27/04 ($459.100 efectivo).
3. Revierte filas Sevuchitas mal marcadas Pagado=Sí (las que el último pago marcó por error).

Filas a revertir: 4-19 + 30-34 (semana 17). Fila 2 (sem 16, OC-REC-SEVU-16, $804.300) queda Sí
porque corresponde al pago real del 24/04 que sí cubrió esa deuda completa.
"""
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from datetime import datetime, timezone, timedelta

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"

creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets"])
sheets = build("sheets", "v4", credentials=creds).spreadsheets()

# 1) Crear hoja Pagos Proveedores si no existe
meta = sheets.get(spreadsheetId=SHEET_ID).execute()
existing = {s["properties"]["title"] for s in meta["sheets"]}
if "Pagos Proveedores" in existing:
    print("Hoja Pagos Proveedores ya existe — skip creación")
else:
    print("Creando hoja Pagos Proveedores...")
    req = {"requests": [{"addSheet": {"properties": {"title": "Pagos Proveedores", "gridProperties": {"frozenRowCount": 1}}}}]}
    sheets.batchUpdate(spreadsheetId=SHEET_ID, body=req).execute()
    sheets.values().update(
        spreadsheetId=SHEET_ID,
        range="Pagos Proveedores!A1:F1",
        valueInputOption="USER_ENTERED",
        body={"values": [["Fecha", "Proveedor", "Efectivo", "Mercado Pago", "Total", "Notas"]]},
    ).execute()
    print("  → Hoja creada con cabeceras")

# 2) Insertar pago histórico Sevuchitas del 27/04 — $459.100 efectivo
print("\nInsertando pago histórico Sevuchitas 27/04 $459.100...")
ar_now = datetime.now(timezone(timedelta(hours=-3)))
fecha_pago = ar_now.strftime("%d/%m/%Y %H:%M")
sheets.values().append(
    spreadsheetId=SHEET_ID,
    range="Pagos Proveedores!A:F",
    valueInputOption="USER_ENTERED",
    insertDataOption="INSERT_ROWS",
    body={"values": [[fecha_pago, "Sevuchitas", 459100, 0, 459100, "Pago parcial semana 17 (migración 27/04)"]]},
).execute()
print(f"  → Insertado: {fecha_pago} | Sevuchitas | Ef $459.100 | MP $0 | Total $459.100")

# 3) Revertir Pagado Proveedor de filas mal marcadas
filas_a_revertir = [4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 18, 19, 30, 31, 32, 33, 34]
print(f"\nRevirtiendo {len(filas_a_revertir)} filas de Sevuchitas col X (Pagado Proveedor) → 'No'...")
data_to_update = []
for r in filas_a_revertir:
    data_to_update.append({"range": f"Orden de Compra!X{r}", "values": [["No"]]})
sheets.values().batchUpdate(
    spreadsheetId=SHEET_ID,
    body={"valueInputOption": "USER_ENTERED", "data": data_to_update},
).execute()
print(f"  → Revertidas filas: {filas_a_revertir}")
print("\nMigración OK.")
print("\nValidación esperada:")
print("  - Sevuchitas total: $1.231.400 - $459.100 = $772.300")
print("  - Sem 17: $509.400 + $469.700 (revertidas) = $979.100  -  $459.100 (pago) = $520.000")
print("  - Sem 18: $722.000")
print("  - Total $520.000 + $722.000 = $1.242.000")
print()
print("⚠️ El total será $1.242.000 (no $772.300) porque al revertir las filas, vuelven al cálculo")
print("   de costos. La deuda real es: costos $1.701.100 - pagos ledger $459.100 = $1.242.000")
print("   ($1.701.100 = $979.100 sem 17 + $722.000 sem 18)")
