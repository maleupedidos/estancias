"""Replica exacta de la logica del backend para calcular saldo vivo.
Imprime cobros, ingresos, gastos post-saldoBase y verifica matematica."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime
import re

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets","v4",credentials=creds).spreadsheets()

def parse_date_any(v):
    if not v: return None
    s = str(v).strip()
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?', s)
    if not m: return None
    yr,mo,dd = int(m.group(3)), int(m.group(2)), int(m.group(1))
    hh = int(m.group(4)) if m.group(4) else 0
    mi = int(m.group(5)) if m.group(5) else 0
    try: return datetime(yr,mo,dd,hh,mi)
    except: return None

def num(v):
    if not v: return 0
    s = str(v).replace("$","").replace(" ","").replace(",","").replace(".","")
    try: return int(s or 0)
    except: return 0

def num_f(v):
    # Para Saldo Base (soporta decimales)
    if not v: return 0
    s = str(v).replace("$","").replace(" ","").replace(".","").replace(",",".")
    try: return float(s)
    except: return 0

# Saldo Base
r = svc.values().get(spreadsheetId=SID, range="'Saldo Base'!A:C").execute()
sb = r.get("values",[])
saldoBase_ef = num_f(sb[-1][1]) if len(sb)>1 else 0
saldoBase_mp = num_f(sb[-1][2]) if len(sb)>1 else 0
saldoBase_fecha_str = sb[-1][0] if len(sb)>1 else ""
saldoBase_date = parse_date_any(saldoBase_fecha_str)
print(f"=== SALDO BASE: {saldoBase_fecha_str}  ef={saldoBase_ef}  mp={saldoBase_mp} ===")
print(f"    parsed date: {saldoBase_date}")
print()

def after_saldo(v):
    if not saldoBase_date: return True
    d = parse_date_any(v)
    if not d: return False
    return d > saldoBase_date

# Cobros
print("=== COBROS POST-SALDO ===")
cobradoEf = cobradoMP = 0
def sumar_cobrado(hoja, cols):
    global cobradoEf, cobradoMP
    r = svc.values().get(spreadsheetId=SID, range=f"'{hoja}'!A1:ZZ").execute()
    vals = r.get("values",[]);
    if not vals: return
    hdr = vals[0]
    idxFc = -1
    for i,h in enumerate(hdr):
        if str(h).strip()=="Fecha de Cobro": idxFc=i; break
    print(f"  [{hoja}] idxFc={idxFc}")
    for row in vals[1:]:
        def g(i): return row[i] if 0<=i<len(row) else ""
        if str(g(cols['ep'])).strip() != 'Cobrado': continue
        if saldoBase_date:
            if idxFc<0: continue
            fc_val = g(idxFc)
            if not after_saldo(fc_val):
                continue
        ef = num(g(cols['ef'])); tr = num(g(cols['tr']))
        pef = num(g(cols['pef'])); ptr = num(g(cols['ptr']))
        cl = g(7)
        nped = g(1)
        print(f"    {hoja} N°{nped} {cl[:25]:<25} fc={g(idxFc) if idxFc>=0 else '?'} ef+pef={ef+pef} tr+ptr={tr+ptr}")
        cobradoEf += ef+pef; cobradoMP += tr+ptr

# Home/Pilar 0-based: ep=12, ef=17, tr=18, pef=19, ptr=20
sumar_cobrado('Home', {'ep':12,'ef':17,'tr':18,'pef':19,'ptr':20})
sumar_cobrado('Pilar', {'ep':12,'ef':17,'tr':18,'pef':19,'ptr':20})
# Clubes 0-based: ep=15, ef=18, tr=19, pef=20, ptr=21
sumar_cobrado('Clubes', {'ep':15,'ef':18,'tr':19,'pef':20,'ptr':21})

print(f"\n  cobradoEf={cobradoEf}  cobradoMP={cobradoMP}")

# Ingresos post-saldo
print("\n=== INGRESOS POST-SALDO ===")
ingEf=ingMP=0
r = svc.values().get(spreadsheetId=SID, range="'Ingresos'!A:H").execute()
for row in r.get("values",[])[1:]:
    if not row: continue
    f = row[0] if len(row)>0 else ""
    if not after_saldo(f): continue
    met = (row[5] if len(row)>5 else "").strip()
    mto = num(row[6] if len(row)>6 else 0)
    cat = row[3] if len(row)>3 else ""
    con = row[4] if len(row)>4 else ""
    print(f"    {f} {met:<15} ${mto} {cat} · {con}")
    if met=="Efectivo": ingEf+=mto
    else: ingMP+=mto

# Gastos post-saldo
print("\n=== GASTOS POST-SALDO ===")
gasEf=gasMP=0
r = svc.values().get(spreadsheetId=SID, range="'Egresos'!A:H").execute()
for row in r.get("values",[])[1:]:
    if not row: continue
    f = row[0] if len(row)>0 else ""
    if not after_saldo(f): continue
    met = (row[5] if len(row)>5 else "").strip()
    mto = num(row[6] if len(row)>6 else 0)
    cat = row[3] if len(row)>3 else ""
    con = row[4] if len(row)>4 else ""
    print(f"    {f} {met:<15} ${mto} {cat} · {con}")
    if met=="Efectivo": gasEf+=mto
    else: gasMP+=mto

# Saldo final
saldoEf = saldoBase_ef + cobradoEf + ingEf - gasEf
saldoMP = saldoBase_mp + cobradoMP + ingMP - gasMP
print(f"\n=== SALDO FINAL ===")
print(f"  Efectivo: {saldoBase_ef} + {cobradoEf} + {ingEf} - {gasEf} = {saldoEf}")
print(f"  MP:       {saldoBase_mp} + {cobradoMP} + {ingMP} - {gasMP} = {saldoMP}")
print(f"  Total: {saldoEf+saldoMP}")
