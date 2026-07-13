"""Leer Egresos col A con valores UNFORMATTED para ver el serial real de la fecha."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime, timedelta

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets","v4",credentials=creds).spreadsheets()

# UNFORMATTED_VALUE devuelve seriales para fechas
r = svc.values().get(spreadsheetId=SID, range="'Egresos'!A:H", valueRenderOption="UNFORMATTED_VALUE").execute()
vals = r.get("values",[])

# Sheets serial date: day number since 1899-12-30
EPOCH = datetime(1899,12,30)
def serial_to_dt(s):
    try:
        d = float(s)
        return EPOCH + timedelta(days=d)
    except: return None

saldo_dt = datetime(2026,4,23,19,30)
print(f"Saldo Base fecha: {saldo_dt}\n")
print(f"{'idx':<4} {'fecha_raw':<20} {'parsed':<22} {'after?':<8} {'met':<15} {'monto':<10} concepto")
for i,row in enumerate(vals):
    if i==0: continue
    f_raw = row[0] if len(row)>0 else ""
    parsed = serial_to_dt(f_raw) if isinstance(f_raw,(int,float)) else None
    if not parsed and isinstance(f_raw,str):
        # Try parse string
        import re
        m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?', f_raw)
        if m:
            try: parsed = datetime(int(m.group(3)),int(m.group(2)),int(m.group(1)),
                int(m.group(4)) if m.group(4) else 0, int(m.group(5)) if m.group(5) else 0)
            except: pass
    after = "YES" if (parsed and parsed>saldo_dt) else "no"
    met = row[5] if len(row)>5 else ""
    mto = row[6] if len(row)>6 else 0
    con = row[4] if len(row)>4 else ""
    print(f"{i:<4} {str(f_raw)[:20]:<20} {str(parsed)[:22]:<22} {after:<8} {str(met)[:15]:<15} ${mto!s:<10} {con}")
