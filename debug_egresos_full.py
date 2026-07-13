from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets","v4",credentials=creds).spreadsheets()

for hoja in ["Egresos","Ingresos","Saldo Base"]:
    print(f"=== {hoja} ===")
    r = svc.values().get(spreadsheetId=SID, range=f"'{hoja}'!A:Z").execute()
    for i,row in enumerate(r.get("values",[])):
        print(f"  {i:>2}: {row}")
    print()
