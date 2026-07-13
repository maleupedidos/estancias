"""Debug: listar pedidos que suman al Efectivo cobrado en el panel."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SPREADSHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"

creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

def read(name):
    r = svc.values().get(spreadsheetId=SPREADSHEET_ID, range=f"'{name}'!A:BA", valueRenderOption="UNFORMATTED_VALUE").execute()
    return r.get("values", [])

def to_num(v):
    if isinstance(v, (int, float)): return v
    try: return float(str(v).replace("$","").replace(",","").replace(".","")) if v else 0
    except: return 0

print("="*80)
print("EFECTIVO COBRADO — DESGLOSE POR PEDIDO")
print("="*80)

total_ef = 0
total_pef = 0

# Home / Pilar / Capital Federal: P=Efectivo (15), Q=Transferencia (16), R=PropEf (17), S=PropTr (18), T=Facturado (19), L=FormaPago (11), M=EstadoPago (12)
for hoja in ["Home", "Pilar", "Capital Federal"]:
    rows = read(hoja)
    if not rows: continue
    print(f"\n--- {hoja} ---")
    print(f"{'N°':<7} {'Cliente':<25} {'FormaPago':<14} {'EstadoPago':<12} {'Efectivo':>10} {'PropEf':>8} {'Total EF':>10}")
    for i, row in enumerate(rows[1:], start=2):
        row = row + [""]*(53-len(row))
        estado_pago = str(row[12]).strip()
        if estado_pago != "Cobrado": continue
        n = row[1]; cliente = str(row[7])[:24]; fp = str(row[11])
        ef = to_num(row[15]); pef = to_num(row[17])
        if ef == 0 and pef == 0: continue
        print(f"{str(n):<7} {cliente:<25} {fp:<14} {estado_pago:<12} {ef:>10,.0f} {pef:>8,.0f} {ef+pef:>10,.0f}")
        total_ef += ef
        total_pef += pef

# Clubes: S=Efectivo (18), T=Transferencia (19), U=PropEf (20), V=PropTr (21), O=FormaPago (14), P=EstadoPago (15)
rows = read("Clubes")
if rows:
    print(f"\n--- Clubes ---")
    print(f"{'N°':<7} {'Cliente':<25} {'FormaPago':<14} {'EstadoPago':<12} {'Efectivo':>10} {'PropEf':>8} {'Total EF':>10}")
    for i, row in enumerate(rows[1:], start=2):
        row = row + [""]*(34-len(row))
        estado_pago = str(row[15]).strip()
        if estado_pago != "Cobrado": continue
        n = row[1]; cliente = str(row[7])[:24]; fp = str(row[14])
        ef = to_num(row[18]); pef = to_num(row[20])
        if ef == 0 and pef == 0: continue
        print(f"{str(n):<7} {cliente:<25} {fp:<14} {estado_pago:<12} {ef:>10,.0f} {pef:>8,.0f} {ef+pef:>10,.0f}")
        total_ef += ef
        total_pef += pef

# Red: Q=Efectivo (16), R=Transferencia (17), S=PropEf (18), T=PropTr (19), M=FormaPago (12), N=EstadoPago (13)
rows = read("Red")
if rows:
    print(f"\n--- Red ---")
    print(f"{'N°':<7} {'Cliente':<25} {'FormaPago':<14} {'EstadoPago':<12} {'Efectivo':>10} {'PropEf':>8} {'Total EF':>10}")
    for i, row in enumerate(rows[1:], start=2):
        row = row + [""]*(55-len(row))
        estado_pago = str(row[13]).strip()
        if estado_pago != "Cobrado": continue
        n = row[1]; cliente = str(row[8])[:24]; fp = str(row[12])
        ef = to_num(row[16]); pef = to_num(row[18])
        if ef == 0 and pef == 0: continue
        print(f"{str(n):<7} {cliente:<25} {fp:<14} {estado_pago:<12} {ef:>10,.0f} {pef:>8,.0f} {ef+pef:>10,.0f}")
        total_ef += ef
        total_pef += pef

print("\n" + "="*80)
print(f"TOTAL Efectivo (col Efectivo): ${total_ef:,.0f}")
print(f"TOTAL Propina Ef:              ${total_pef:,.0f}")
print(f"TOTAL EF + PropEf (panel):     ${total_ef+total_pef:,.0f}")
print(f"Dicho por Tadeo:               $372,800")
print(f"Diferencia:                    ${total_ef+total_pef-372800:,.0f}")
print("="*80)
