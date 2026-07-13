"""Convierte valores '$12.345' a números en Historico Home/Pilar/Clubes.
Se ejecuta después de la migración de layout."""
import re
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

MONEY_RE = re.compile(r'^-?\$?\s*-?[\d.,]+$')

def parse_money(v):
    """Convierte '$58.000' -> 58000. 'no-es-plata' -> v. None -> ''."""
    if v is None or v == "":
        return ""
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    if not s: return ""
    # Sheets format $#,##0 en es-AR usa '.' para miles y ',' para decimales
    # ej: '$58.000' -> 58000; '$1.234,56' -> 1234.56
    if MONEY_RE.match(s):
        # quitar $, espacios; convertir '.' miles a nada; ',' decimal a '.'
        num = s.replace('$', '').replace(' ', '').strip()
        # Si tiene tanto '.' como ',' asumimos '.' miles y ',' decimal
        if '.' in num and ',' in num:
            num = num.replace('.', '').replace(',', '.')
        elif ',' in num:
            num = num.replace(',', '.')
        elif num.count('.') == 1 and len(num.split('.')[1]) <= 2:
            pass  # puede ser decimal '.': dejarlo
        else:
            # múltiples '.' o '.' con 3 dígitos → separadores de miles
            num = num.replace('.', '')
        try: return float(num) if '.' in num else int(num)
        except: return s
    return s

for hoja in ["Historico Home", "Historico Pilar", "Historico Clubes"]:
    print(f"\n{hoja}...")
    r = svc.values().get(spreadsheetId=SID, range=f"'{hoja}'!A:BZ", valueRenderOption="UNFORMATTED_VALUE").execute()
    rows = r.get("values", [])
    if not rows: continue
    header = rows[0]
    # Identificar columnas monetarias por nombre
    money_headers = {'Total ($)','Envío','Env\u00edo','Efectivo','Transferencia','Propina Efectivo','Propina Transferencia','Facturado','Costo','Margen Bruto','Subtotal sin Descuento','Descuento'}
    money_idx = set()
    for i, h in enumerate(header):
        if str(h).strip() in money_headers:
            money_idx.add(i)
    # Procesar cada fila
    changed = 0
    new_data = [header]
    for row in rows[1:]:
        r2 = list(row) + [""] * (len(header) - len(row))
        for i in money_idx:
            if i >= len(r2): continue
            v = r2[i]
            if isinstance(v, str) and v:
                nv = parse_money(v)
                if nv != v:
                    r2[i] = nv
                    changed += 1
        new_data.append(r2)
    # Re-escribir todo
    svc.values().clear(spreadsheetId=SID, range=f"'{hoja}'!A:BZ").execute()
    svc.values().update(spreadsheetId=SID, range=f"'{hoja}'!A1", valueInputOption="USER_ENTERED", body={"values": new_data}).execute()
    print(f"  {changed} celdas convertidas de string a número. {len(new_data)-1} filas re-escritas.")
print("\nListo.")
