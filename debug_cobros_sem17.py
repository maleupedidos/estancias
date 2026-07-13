"""Diagnostico: pedidos sem 17 (21/04-27/04) Entregados y estado de pago en Home/Pilar/Clubes/Red."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

# sem 17 = 21/04 al 27/04
def parse_fecha(s):
    if not s: return None
    s = str(s).strip()
    for fmt in ("%d/%m/%Y","%d/%m/%y","%Y-%m-%d"):
        try: return datetime.strptime(s, fmt)
        except: pass
    return None

def money_to_num(s):
    if not s: return 0
    try:
        return int(str(s).replace("$","").replace(".","").replace(" ","").replace(",","") or 0)
    except: return 0

hojas = {
    # hoja: (col_fecha_entrega_idx_1based, col_estado_entrega, col_estado_pago, col_total, col_fecha_cobro, col_cliente, col_id, col_fp, col_ef, col_tr)
    "Home":   (50, 11, 13, 22, 55, 8, 2, 12, 18, 19),  # AX=50 Fecha Entrega, V=22 Facturado, BC=55 Fecha de Cobro
    "Pilar":  (52, 11, 13, 22, 58, 8, 2, 12, 18, 19),  # BA=52, BF=58
    "Clubes": (0,  14, 16, 0, 0, 8, 2, 15, 19, 20),
    "Red":    (0,  12, 14, 0, 0, 8, 2, 13, 17, 18),
}

for hoja, idx in hojas.items():
    fe_i, ee_i, ep_i, tot_i, fc_i, cl_i, id_i, fp_i, ef_i, tr_i = idx
    r = svc.values().get(spreadsheetId=SID, range=f"'{hoja}'!A1:ZZ").execute()
    vals = r.get("values", [])
    if not vals: continue
    headers = vals[0]
    data = vals[1:]
    print(f"\n=== {hoja} — {len(data)} filas ===")

    # Mes = Abril, últimos 20 pedidos
    printed = 0
    for i in range(len(data)-1, -1, -1):
        row = data[i]
        def g(j): return row[j-1] if j-1 < len(row) and j>0 else ""
        est = str(g(ee_i)).strip() if ee_i else ""
        ep = str(g(ep_i)).strip() if ep_i else ""
        cliente = str(g(cl_i)).strip()
        nped = str(g(id_i)).strip()
        if not cliente: continue
        # Solo Abril
        fecha = parse_fecha(g(4))  # D=4 Fecha
        if not fecha or fecha.month != 4 or fecha.year != 2026: continue
        fp = str(g(fp_i)).strip() if fp_i else ""
        ef = g(ef_i) if ef_i else ""
        tr = g(tr_i) if tr_i else ""
        fc = g(fc_i) if fc_i else ""
        tot = g(tot_i) if tot_i else ""
        flag = ""
        if est == "Entregado" and ep != "Cobrado":
            flag = "  <<< ENTREGADO SIN COBRAR"
        elif est == "Entregado" and ep == "Cobrado" and not fc:
            flag = "  <<< COBRADO SIN FECHA DE COBRO"
        print(f"  fila {i+2:>3} N°{nped:<6} {fecha.strftime('%d/%m'):<6} {cliente[:22]:<22} est={est:<10} pago={ep:<12} fp={fp:<14} ef={str(ef)[:10]:<10} tr={str(tr)[:10]:<10} fc={str(fc)[:16]}{flag}")
        printed += 1
        if printed >= 15: break
