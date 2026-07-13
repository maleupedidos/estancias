"""Crea la hoja 'Resumenes Semanales' para histórico de resúmenes."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SPREADSHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

meta = svc.get(spreadsheetId=SPREADSHEET_ID).execute()
existing = [s["properties"]["title"] for s in meta["sheets"]]
if "Resumenes Semanales" in existing:
    print("Ya existe.")
else:
    body = {"requests":[{"addSheet":{"properties":{"title":"Resumenes Semanales","tabColor":{"red":0.94,"green":0.49,"blue":0.28}}}}]}
    svc.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()
    headers = [["Fecha generación","Semana","Periodo","Pedidos","Facturado","Cobrado","CMV","Margen Bruto","Gastos","Resultado Neto","Clientes únicos","Caras nuevas","Top cliente","Top producto","Meta alcanzada","JSON completo"]]
    svc.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range="'Resumenes Semanales'!A1",
        valueInputOption="RAW",
        body={"values":headers}
    ).execute()
    print("Hoja creada con headers.")
