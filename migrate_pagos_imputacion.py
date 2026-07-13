"""
Migración v2 hoja Pagos Proveedores:
- Agregar columna F "Semana Imputada" (entre E Total y F Notas viejas).
- Reasignar el pago de hoy ($459.100 Sevuchitas) → Semana Imputada = 17.
"""
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"

creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets"])
sheets = build("sheets", "v4", credentials=creds).spreadsheets()

# Leer estado actual
resp = sheets.values().get(spreadsheetId=SHEET_ID, range="Pagos Proveedores!A1:G").execute()
data = resp.get("values", [])
print("Estado actual hoja Pagos Proveedores:")
for r, row in enumerate(data, start=1):
    print(f"  {r}: {row}")

# Header actual: ['Fecha', 'Proveedor', 'Efectivo', 'Mercado Pago', 'Total', 'Notas']
# Necesitamos: ['Fecha', 'Proveedor', 'Efectivo', 'Mercado Pago', 'Total', 'Semana Imputada', 'Notas']
# Hay que insertar una columna entre E (Total) y F (Notas) actual.

if len(data[0]) == 6 and data[0][5] == "Notas":
    print("\nReestructurando: insertando columna F 'Semana Imputada' antes de 'Notas'...")
    # Estrategia: reescribir todas las filas con la nueva estructura.
    new_rows = [["Fecha", "Proveedor", "Efectivo", "Mercado Pago", "Total", "Semana Imputada", "Notas"]]
    for r, row in enumerate(data[1:], start=2):
        while len(row) < 6:
            row.append("")
        # Si es el pago de hoy ($459.100 a Sevuchitas) → imputar a 17
        proveedor = row[1]
        total = str(row[4]).replace("$", "").replace(".", "").replace(",", "").strip()
        try:
            total_num = float(total)
        except:
            total_num = 0
        sem_imp = ""
        notas_old = row[5]
        if proveedor == "Sevuchitas" and abs(total_num - 459100) < 1:
            sem_imp = "17"
            if "semana 17" not in notas_old.lower():
                notas_old = (notas_old + " (imputado a sem 17)").strip()
        new_row = [row[0], row[1], row[2], row[3], row[4], sem_imp, notas_old]
        new_rows.append(new_row)

    # Escribir
    sheets.values().clear(spreadsheetId=SHEET_ID, range="Pagos Proveedores!A:G").execute()
    sheets.values().update(
        spreadsheetId=SHEET_ID,
        range=f"Pagos Proveedores!A1:G{len(new_rows)}",
        valueInputOption="USER_ENTERED",
        body={"values": new_rows},
    ).execute()
    print(f"  -> Reescritas {len(new_rows)} filas con nueva columna F")
elif len(data[0]) >= 7 and data[0][5] == "Semana Imputada":
    print("\nLa hoja ya tiene la columna 'Semana Imputada'. Skip migración.")
else:
    print(f"\nHeader inesperado: {data[0]}. Abortando.")
    raise SystemExit(1)

print("\nEstado final:")
resp = sheets.values().get(spreadsheetId=SHEET_ID, range="Pagos Proveedores!A1:G").execute()
for r, row in enumerate(resp.get("values", []), start=1):
    print(f"  {r}: {row}")
