"""Diagnóstico: fila 233 de Home + todas las OC del cliente Valeria Fernandez Marin."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SPREADSHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

print("=== HOME fila 233 (primeras 20 cols) ===")
r = svc.values().get(spreadsheetId=SPREADSHEET_ID, range="'Home'!A233:T233").execute()
row = r.get("values", [[]])[0] if r.get("values") else []
home_headers = svc.values().get(spreadsheetId=SPREADSHEET_ID, range="'Home'!A1:T1").execute().get("values", [[]])[0]
for i, h in enumerate(home_headers):
    val = row[i] if i < len(row) else ""
    print(f"  {chr(65+i) if i<26 else 'A'+chr(65+i-26)} {h}: {val}")

print()
print("=== Buscar Valeria en hoja Home (ped num) ===")
r = svc.values().get(spreadsheetId=SPREADSHEET_ID, range="'Home'!A1:H").execute()
rows = r.get("values", [])
for idx, rr in enumerate(rows, 1):
    nombre = rr[7] if len(rr) > 7 else ""
    if "valeria" in nombre.lower() and "fernandez" in nombre.lower():
        print(f"  Fila {idx}: N°Ped={rr[1] if len(rr)>1 else ''} | Cliente={nombre}")

print()
print("=== OC con cliente Valeria ===")
r = svc.values().get(spreadsheetId=SPREADSHEET_ID, range="'Orden de Compra'!A1:Y").execute()
oc_rows = r.get("values", [])
hdr = oc_rows[0] if oc_rows else []
for idx, rr in enumerate(oc_rows[1:], 2):
    cliente = rr[6] if len(rr) > 6 else ""
    if "valeria" in cliente.lower() and "fernandez" in cliente.lower():
        ocnum = rr[0] if len(rr) > 0 else ""
        fCre = rr[1] if len(rr) > 1 else ""
        sem = rr[2] if len(rr) > 2 else ""
        canal = rr[4] if len(rr) > 4 else ""
        nped = rr[5] if len(rr) > 5 else ""
        prod = rr[10] if len(rr) > 10 else ""
        qty = rr[12] if len(rr) > 12 else ""
        origen = rr[19] if len(rr) > 19 else ""
        est = rr[20] if len(rr) > 20 else ""
        print(f"  Fila OC {idx}: {ocnum} | {fCre} sem{sem} | {canal} N°Ped={nped} | {prod} x{qty} | Origen={origen} | Est={est}")
