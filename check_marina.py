"""Diagnóstico pedido de Marina Donati (Red): origen mixto DEP/OC, TLC en Depósito."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SPREADSHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

# 1) Buscar Marina Donati en hoja Red
print("=== HOJA RED — buscar Marina Donati ===")
red_data = svc.values().get(spreadsheetId=SPREADSHEET_ID, range="'Red'!A1:BC").execute().get("values", [])
red_headers = red_data[0]
print(f"Cols Red: {len(red_headers)}")
def col_letter(i):
    if i<26: return chr(65+i)
    return chr(65 + i//26 - 1) + chr(65 + i%26)

# Localizar columnas relevantes
hdr_low = [h.strip().lower() for h in red_headers]
def gi(name):
    try: return hdr_low.index(name.lower())
    except: return -1

i_n = gi("N°")
i_cli = gi("Cliente")
i_origen = gi("Origen")
i_origen_det = gi("Origen Detalle")
i_estado = gi("Estado de Entrega")
i_pago = gi("Estado de Pago")
i_total = gi("Total ($)")
i_fecha = gi("Fecha")

target_row = None
for idx, r in enumerate(red_data[1:], start=2):
    cli = (r[i_cli] if i_cli>=0 and i_cli<len(r) else "").lower()
    if "donati" in cli:
        target_row = idx
        print(f"  Fila {idx}: N°={r[i_n] if i_n<len(r) else ''} | Cliente={r[i_cli]} | Origen={r[i_origen] if i_origen<len(r) else ''} | Estado={r[i_estado] if i_estado<len(r) else ''}")
        print(f"  Origen Detalle (col {col_letter(i_origen_det)}): {r[i_origen_det] if i_origen_det>=0 and i_origen_det<len(r) else '(vacío)'}")
        print()
        # Mostrar todos los productos con qty > 0
        print("  Productos con qty:")
        # Las cols de producto están entre 22 y 44 según docs (Red), o detectar por header
        for ci, h in enumerate(red_headers):
            if ci<i_origen_det and ci>=15:  # rango aprox de productos
                val = r[ci] if ci<len(r) else ""
                if val and str(val) not in ("0","",""):
                    try:
                        q = int(str(val).replace(",","").replace(".",""))
                        if q > 0:
                            print(f"    {col_letter(ci)} {h}: {q}")
                    except:
                        pass
        break

if not target_row:
    print("  No se encontró fila de Marina Donati en Red")

# 2) Buscar OCs vinculadas
print()
print("=== HOJA ORDEN DE COMPRA — OCs de Marina Donati ===")
oc_data = svc.values().get(spreadsheetId=SPREADSHEET_ID, range="'Orden de Compra'!A1:Y").execute().get("values", [])
oc_headers = oc_data[0]
print(f"Cols OC: {len(oc_headers)}")
print(f"Total filas OC: {len(oc_data)-1}")
print()
print("  Buscando filas con Cliente = Marina Donati y/o Canal=Red...")
for idx, r in enumerate(oc_data[1:], start=2):
    cli = (r[6] if len(r)>6 else "").lower()  # G = Cliente
    canal = r[4] if len(r)>4 else ""           # E = Canal
    pedidoOrig = r[5] if len(r)>5 else ""      # F = N° Pedido Origen
    if "donati" in cli:
        print(f"  Fila {idx} OC: N°={r[0]} | Canal={canal} | Pedido={pedidoOrig} | Cliente={r[6]} | Producto={r[10] if len(r)>10 else ''} ({r[11] if len(r)>11 else ''}) | Cant={r[12] if len(r)>12 else ''} | Estado={r[20] if len(r)>20 else ''}")
