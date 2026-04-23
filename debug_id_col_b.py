"""Debug profundo: mapear toda la hoja Pilar y Home para ver IDs, estados, y la
situacion real del bug Gloria Pavlovsky."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

def dump(hoja, cols="A:AY"):
    uf = svc.values().get(spreadsheetId=SID, range=f"'{hoja}'!{cols}",
                          valueRenderOption="UNFORMATTED_VALUE").execute().get("values", [])
    fm = svc.values().get(spreadsheetId=SID, range=f"'{hoja}'!{cols}",
                          valueRenderOption="FORMATTED_VALUE").execute().get("values", [])
    return uf, fm

# PILAR: col B=N°, H=Cliente, I=Origen, J=Dia Entrega, K=Est Entrega, L=Fp, M=Est Pago, N=Total, V=Facturado
print("=" * 90)
print("PILAR - Todas las filas")
print("=" * 90)
uf, fm = dump("Pilar")
print(f"{'fila':5} {'N°':4} {'Cliente':25} {'Origen':10} {'Dia':10} {'EstEnt':12} {'EstPago':12}")
print("-" * 90)
for i, row in enumerate(fm):
    if i == 0: continue
    b = uf[i][1] if len(uf[i]) > 1 else ""
    cli = row[7] if len(row) > 7 else ""
    origen = row[8] if len(row) > 8 else ""
    dia = row[9] if len(row) > 9 else ""
    ent = row[10] if len(row) > 10 else ""
    pago = row[12] if len(row) > 12 else ""
    mark = " <-- GLORIA" if "Gloria" in str(cli) else ""
    print(f"{i+1:5} {str(b):4} {str(cli)[:24]:25} {str(origen)[:9]:10} {str(dia)[:9]:10} {str(ent)[:11]:12} {str(pago)[:11]:12}{mark}")

print()
print("=" * 90)
print("HOME - Últimas 10 filas")
print("=" * 90)
uf, fm = dump("Home")
print(f"{'fila':5} {'N°':4} {'Cliente':25} {'Origen':10} {'Dia':10} {'EstEnt':12} {'EstPago':12}")
print("-" * 90)
for i, row in enumerate(fm[-10:]):
    idx = len(fm) - 10 + i
    b = uf[idx][1] if len(uf[idx]) > 1 else ""
    cli = row[7] if len(row) > 7 else ""
    origen = row[8] if len(row) > 8 else ""
    dia = row[9] if len(row) > 9 else ""
    ent = row[10] if len(row) > 10 else ""
    pago = row[12] if len(row) > 12 else ""
    print(f"{idx+1:5} {str(b):4} {str(cli)[:24]:25} {str(origen)[:9]:10} {str(dia)[:9]:10} {str(ent)[:11]:12} {str(pago)[:11]:12}")
