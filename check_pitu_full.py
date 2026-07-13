"""Diagnóstico fila 264 Home — leer TODAS las columnas incluso las de descuento."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SPREADSHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

# Headers + row 264 hasta col BB para incluir Subtotal sin Descuento y Descuento
headers = svc.values().get(spreadsheetId=SPREADSHEET_ID, range="'Home'!A1:BB1").execute().get("values", [[]])[0]
row = svc.values().get(spreadsheetId=SPREADSHEET_ID, range="'Home'!A264:BB264").execute().get("values", [[]])[0]

print(f"Total cols header: {len(headers)}, en fila 264: {len(row)}")
print()
def col_letter(i):
    if i<26: return chr(65+i)
    return chr(65 + i//26 - 1) + chr(65 + i%26)
print("=== TODAS las cols (header → valor) ===")
for i, h in enumerate(headers):
    val = row[i] if i < len(row) else "(vacío)"
    marker = " ⚠️" if "descuento" in h.lower() or "subtotal" in h.lower() else ""
    print(f"  {col_letter(i)} ({i+1:>2}) {h:<35}: {val}{marker}")

# También: comparar contra otra fila reciente (pedido +$100k transferencia) si existe
print()
print("=== Buscando otros pedidos +$100k Transferencia recientes ===")
all_data = svc.values().get(spreadsheetId=SPREADSHEET_ID, range="'Home'!A1:BB").execute().get("values", [])
hdr_low = [h.strip().lower() for h in all_data[0]]
def gi(name):
    try: return hdr_low.index(name.lower())
    except: return -1
i_total = gi("Total ($)")
i_pago = gi("Forma de Pago")
i_desc = gi("Descuento")
i_subt = gi("Subtotal sin Descuento")
i_cli  = gi("Cliente")
i_fecha = gi("Fecha")
print(f"  cols: total={i_total}, pago={i_pago}, desc={i_desc}, subt={i_subt}, cli={i_cli}, fecha={i_fecha}")
for idx, r in enumerate(all_data[1:], start=2):
    try:
        total_str = r[i_total] if i_total>=0 and i_total<len(r) else ""
        total_num = int(str(total_str).replace("$","").replace(".","").replace(",","").strip() or 0)
        if total_num >= 100000 and i_pago>=0 and i_pago<len(r) and r[i_pago]=="Transferencia":
            cli = r[i_cli] if i_cli>=0 and i_cli<len(r) else ""
            fec = r[i_fecha] if i_fecha>=0 and i_fecha<len(r) else ""
            desc = r[i_desc] if i_desc>=0 and i_desc<len(r) else ""
            subt = r[i_subt] if i_subt>=0 and i_subt<len(r) else ""
            print(f"  fila {idx}: {fec} | {cli:<25} | Total={total_str} | Subt={subt} | Desc={desc}")
    except Exception as e:
        pass
