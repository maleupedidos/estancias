"""Dump headers detallado Clubes + Red."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

def col_letter(n):
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26); s = chr(65 + rem) + s
    return s

for hoja in ["Clubes","Red"]:
    r = svc.values().get(spreadsheetId=SID, range=f"'{hoja}'!1:1").execute()
    headers = r.get("values", [[]])[0] if r.get("values") else []
    print(f"=== {hoja} ({len(headers)} cols) ===")
    for i, h in enumerate(headers, 1):
        print(f"  {col_letter(i):>3} ({i:>2}): {h}")
    print()
