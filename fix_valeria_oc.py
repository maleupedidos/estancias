"""Fix: OCs de Valeria Fernandez Marin (OC-1 y OC-2) tienen N° Pedido Origen
mal (3 y 4). El N° correcto es 232 (fila 233 de Home, col B).

Verifica antes y después. Aborta si los datos no coinciden con lo esperado.
"""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SPREADSHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(
    SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets"]
)
svc = build("sheets", "v4", credentials=creds).spreadsheets()

# ── 1. Verificar N° Pedido real en Home fila 233 ──
r = svc.values().get(spreadsheetId=SPREADSHEET_ID, range="'Home'!B233:H233").execute()
home_row = r.get("values", [[]])[0]
home_nped = home_row[0] if len(home_row) > 0 else ""
home_cliente = home_row[6] if len(home_row) > 6 else ""
print(f"Home fila 233: N°Ped={home_nped!r} Cliente={home_cliente!r}")
assert home_nped == "232", f"Home fila 233 no tiene N°Ped=232 (tiene {home_nped!r}). Abort."
assert "valeria" in home_cliente.lower(), f"Home fila 233 no es Valeria ({home_cliente!r}). Abort."

# ── 2. Verificar OC filas 4 y 5 son Valeria con F=3 y F=4 ──
r = svc.values().get(spreadsheetId=SPREADSHEET_ID, range="'Orden de Compra'!A4:G5").execute()
oc_rows = r.get("values", [])
print(f"OC fila 4: {oc_rows[0] if len(oc_rows) > 0 else 'vacia'}")
print(f"OC fila 5: {oc_rows[1] if len(oc_rows) > 1 else 'vacia'}")

assert len(oc_rows) == 2, "Se esperaban 2 filas de OC (4 y 5)"
for i, r_ in enumerate(oc_rows):
    cliente = r_[6] if len(r_) > 6 else ""
    nped = r_[5] if len(r_) > 5 else ""
    assert "valeria" in cliente.lower(), f"OC fila {4+i} no es Valeria ({cliente!r})"
    assert nped in ("3", "4"), f"OC fila {4+i} tiene N°Ped={nped!r}, esperaba 3 o 4"

# ── 3. Update col F (N° Pedido Origen) a "232" para ambas filas ──
print("\nActualizando OC F4:F5 -> 232 ...")
res = svc.values().update(
    spreadsheetId=SPREADSHEET_ID,
    range="'Orden de Compra'!F4:F5",
    valueInputOption="USER_ENTERED",
    body={"values": [["232"], ["232"]]},
).execute()
print(f"Update OK: {res.get('updatedCells')} celdas")

# ── 4. Verificar ──
r = svc.values().get(spreadsheetId=SPREADSHEET_ID, range="'Orden de Compra'!F4:F5").execute()
print(f"Despues: {r.get('values')}")
