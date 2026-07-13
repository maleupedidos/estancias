"""Listar pedidos cobrados cuya fecha relevante cae en Semana ISO 17 del 2026
y que suman Transferencia + Propina Transferencia al total del Panel."""
from datetime import datetime, date
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SPREADSHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

def week_iso_js(d):
    """Replica getWeek() del panel.html (no es ISO pura, usa epiphany del 1/1)."""
    o = date(d.year, 1, 1)
    days = (d - o).days
    import math
    return math.ceil((days + o.weekday() + 1 + 1) / 7)
# El JS usa `o.getDay()` que es 0=Sunday en JS, pero en Python weekday() es 0=Monday.
# Convertir: JS getDay()==0 Sunday → Python weekday() Sunday=6. 1/1/2026 JS getDay=Thursday=4, Py weekday=3.
# Mejor replicar literal: JS hace Math.ceil(((d-o)/864e5 + o.getDay() + 1) / 7)

def week_panel(d):
    o = date(d.year, 1, 1)
    days = (d - o).days
    js_day = (o.weekday() + 1) % 7  # Python Mon=0 → JS: Mon=1, Sun=0. weekday() Mon=0 → +1 mod 7 = 1. Sun=6 → 0.
    import math
    return math.ceil((days + js_day + 1) / 7)

def read(name):
    r = svc.values().get(spreadsheetId=SPREADSHEET_ID, range=f"'{name}'!A:BZ", valueRenderOption="FORMATTED_VALUE").execute()
    return r.get("values", [])

def to_num(v):
    if isinstance(v, (int, float)): return v
    try: return float(str(v).replace("$","").replace(",","").replace(".","")) if v else 0
    except: return 0

def parse_fecha(s):
    if not s: return None
    s = str(s).split(" ")[0]
    parts = s.split("/")
    if len(parts) < 2: return None
    try:
        d = int(parts[0]); m = int(parts[1])
        y = int(parts[2]) if len(parts) == 3 else 2026
        if y < 100: y += 2000
        return date(y, m, d)
    except: return None

print(f"Hoy {date.today()} es semana {week_panel(date.today())}")
print("="*90)
print(f"{'Hoja':<8} {'N°':<5} {'Cliente':<25} {'FechaRel':<12} {'Sem':<4} {'Tr':>10} {'PropTr':>8} {'TotTR':>10}")
print("="*90)

total_tr = 0
found = []

def proc(hoja, rows, idx_estPago, idx_tr, idx_ptr, fc_header="Fecha de Cobro", idx_fe=None, idx_f=None):
    global total_tr
    if not rows: return
    headers = rows[0]
    idx_fc = next((i for i,h in enumerate(headers) if str(h).strip() == fc_header), None)
    print(f"\n-- {hoja}: {len(rows)-1} filas --")
    for i, row in enumerate(rows[1:], start=2):
        if not row: continue
        row = row + [""]*(len(headers)-len(row))
        ep = str(row[idx_estPago]).strip()
        if ep != "Cobrado": continue
        tr = to_num(row[idx_tr])
        ptr = to_num(row[idx_ptr])
        ef = to_num(row[idx_tr-1])
        pef = to_num(row[idx_ptr-1])
        total_row = tr + ptr
        fp = str(row[11]).strip() if hoja != 'Clubes' else str(row[14]).strip()
        # show all cobrados
        fc_raw = row[idx_fc] if idx_fc is not None and idx_fc < len(row) else ""
        fe_raw = row[idx_fe] if idx_fe is not None and idx_fe < len(row) else ""
        f_raw = row[idx_f] if idx_f is not None else ""
        fc = parse_fecha(fc_raw) or parse_fecha(fe_raw) or parse_fecha(f_raw)
        sem = week_panel(fc) if fc else 0
        n = row[1]; cliente = str(row[7])[:20]
        fac = to_num(row[19]) if hoja != 'Clubes' else to_num(row[22])
        print(f"  {hoja} N°{n} {cliente:<20} fp={fp:<14} EF={ef:>8,.0f} pEF={pef:>5,.0f} TR={tr:>8,.0f} pTR={ptr:>5,.0f} Fact={fac:>8,.0f} Sem={sem} fc={fc}")
        if total_row == 0: continue
        # Fecha relevante: fc > fe > f
        fc_raw = row[idx_fc] if idx_fc is not None else ""
        fe_raw = row[idx_fe] if idx_fe is not None and idx_fe < len(row) else ""
        f_raw = row[idx_f] if idx_f is not None else ""
        fc = parse_fecha(fc_raw) or parse_fecha(fe_raw) or parse_fecha(f_raw)
        if not fc: continue
        sem = week_panel(fc)
        n = row[1]; cliente = str(row[7])[:24]
        print(f"{hoja:<8} {str(n):<5} {cliente:<25} {fc.strftime('%d/%m/%Y'):<12} {sem:<4} {tr:>10,.0f} {ptr:>8,.0f} {total_row:>10,.0f}")
        total_tr += total_row
        found.append((hoja, n, cliente, total_row))

# Home: P=15 Ef, Q=16 Tr, R=17 pEf, S=18 pTr, M=12 estPago, D=3 fecha, AV=47 fechaEntrega
rows = read("Home")
proc("Home", rows, idx_estPago=12, idx_tr=16, idx_ptr=18, idx_fe=47, idx_f=3)

# Pilar NEW: M=12 estPago, Q=16 Tr, S=18 pTr, AY=50 fechaEntrega, D=3 fecha
rows = read("Pilar")
proc("Pilar", rows, idx_estPago=12, idx_tr=16, idx_ptr=18, idx_fe=50, idx_f=3)

# Clubes: P=15 estPago, T=19 Tr, V=21 pTr, D=3 fecha
rows = read("Clubes")
# Clubes no tiene fechaEntrega
proc("Clubes", rows, idx_estPago=15, idx_tr=19, idx_ptr=21, idx_fe=None, idx_f=3)

print("="*90)
print(f"TOTAL TR + PropTr (Semana 17): ${total_tr:,.0f}")
print(f"Dicho por Panel: $188,570")
print(f"Diferencia: ${total_tr - 188570:,.0f}")
