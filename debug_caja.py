"""Debug saldo Caja: saldoBase + cobrado pedidos + ingresos - gastos."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

def read(name):
    return svc.values().get(spreadsheetId=SID, range=f"'{name}'!A:Z", valueRenderOption="UNFORMATTED_VALUE").execute().get("values", [])

def num(v):
    if isinstance(v,(int,float)): return v
    try: return float(str(v).replace("$","").replace(",","")) if v else 0
    except: return 0

# Saldo Base (última fila)
sb = read("Saldo Base")
if len(sb) > 1:
    last = sb[-1]
    print(f"Saldo Base (última fila): EF=${num(last[1]):,.0f}  MP=${num(last[2]):,.0f}  fecha={last[0]}")
else:
    print("Saldo Base vacío")
print()

# Cobrado en hojas operativas (ahora mismo)
print("--- COBRADO en hojas operativas (después de archivar) ---")
cob_ef_op = 0; cob_mp_op = 0
for hoja in ["Home","Pilar","Clubes","Red"]:
    rows = read(hoja)
    if not rows: continue
    ef_h=0; mp_h=0; n_ef=0; n_mp=0
    if hoja == "Clubes":
        idxFp, idxEp, idxFact = 14, 15, 22
    elif hoja == "Red":
        idxFp, idxEp, idxFact = 12, 13, 20
    else:
        idxFp, idxEp, idxFact = 11, 12, 19
    for r in rows[1:]:
        r = r + [""]*(max(idxFact,idxEp,idxFp)+1-len(r))
        if str(r[idxEp]).strip() != "Cobrado": continue
        fac = num(r[idxFact])
        if str(r[idxFp]).strip() == "Efectivo":
            ef_h += fac; n_ef += 1
        else:
            mp_h += fac; n_mp += 1
    print(f"  {hoja:<8} EF=${ef_h:>12,.0f} ({n_ef} ped)  MP=${mp_h:>12,.0f} ({n_mp} ped)")
    cob_ef_op += ef_h
    cob_mp_op += mp_h

print(f"\n  TOTAL cobrado operativo: EF=${cob_ef_op:,.0f}  MP=${cob_mp_op:,.0f}")

# Ingresos
ing = read("Ingresos")
ing_ef=0; ing_mp=0
for r in (ing or [])[1:]:
    r = r + [""]*10
    met = str(r[5]).strip(); mt = num(r[6])
    if met == "Efectivo": ing_ef += mt
    else: ing_mp += mt
print(f"\nIngresos: EF=${ing_ef:,.0f}  MP=${ing_mp:,.0f}")

# Egresos
eg = read("Egresos")
eg_ef=0; eg_mp=0
for r in (eg or [])[1:]:
    r = r + [""]*10
    met = str(r[5]).strip(); mt = num(r[6])
    if met == "Efectivo": eg_ef += mt
    else: eg_mp += mt
print(f"Egresos:  EF=${eg_ef:,.0f}  MP=${eg_mp:,.0f}")

# Cálculo final
sb_ef = num(sb[-1][1]) if len(sb)>1 else 0
sb_mp = num(sb[-1][2]) if len(sb)>1 else 0
saldo_ef = sb_ef + cob_ef_op + ing_ef - eg_ef
saldo_mp = sb_mp + cob_mp_op + ing_mp - eg_mp
print(f"\n--- SALDO (fórmula panel) ---")
print(f"EF: base {sb_ef:,.0f} + cobrado {cob_ef_op:,.0f} + ingresos {ing_ef:,.0f} - gastos {eg_ef:,.0f} = ${saldo_ef:,.0f}")
print(f"MP: base {sb_mp:,.0f} + cobrado {cob_mp_op:,.0f} + ingresos {ing_mp:,.0f} - gastos {eg_mp:,.0f} = ${saldo_mp:,.0f}")
print(f"Total: ${saldo_ef+saldo_mp:,.0f}")
