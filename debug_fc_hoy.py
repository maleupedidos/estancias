"""Ver Fecha de Cobro y estado de los 3 cobros de hoy + saldo base."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets","v4",credentials=creds).spreadsheets()

# Saldo Base
r = svc.values().get(spreadsheetId=SID, range="'Saldo Base'!A:C").execute()
sb = r.get("values",[])
print("=== SALDO BASE (último) ===")
if len(sb)>1:
    ult = sb[-1]
    print(f"  fecha={ult[0]}  ef={ult[1]}  mp={ult[2]}")
print()

# Home: Iñaki (N°236 23/04), Carolina Llorente (buscar)
# Pilar: Carolina Llorente N°10 (24/04)?  -- usuario dijo Home; verificar
# Clubes: Lucila Blanco N°29 (23/04)
for hoja in ["Home","Pilar","Clubes"]:
    r = svc.values().get(spreadsheetId=SID, range=f"'{hoja}'!A1:ZZ").execute()
    vals = r.get("values",[])
    if not vals: continue
    hdr = vals[0]
    # indice col Fecha de Cobro
    try: idxFc = hdr.index("Fecha de Cobro")
    except ValueError:
        # buscar variante con tildes
        idxFc = -1
        for i,h in enumerate(hdr):
            if "Fecha de Cobro" in str(h) or "Fecha Cobro" in str(h):
                idxFc = i; break
    try: idxEp = hdr.index("Estado de Pago")
    except: idxEp = -1
    try: idxCl = hdr.index("Cliente")
    except: idxCl = -1
    print(f"=== {hoja} — idxFc={idxFc} ({chr(65+idxFc) if idxFc<26 else '?'}) idxEp={idxEp} ===")
    data = vals[1:]
    # últimas 10 filas
    for i in range(max(0,len(data)-10), len(data)):
        row = data[i]
        def g(j): return row[j] if 0<=j<len(row) else ""
        cl = g(idxCl); ep = g(idxEp); fc = g(idxFc) if idxFc>=0 else "?"
        nped = g(1) # col B
        fecha = g(3)
        print(f"  fila{i+2:>3} N°{nped:<5} {fecha:<12} {str(cl)[:25]:<25} ep={ep:<12} fc={fc}")
    print()
