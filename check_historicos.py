"""Comparar headers: Hoja operativa vs Historico."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SPREADSHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

def col_letter(n):
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s

pares = [("Home", "Historico Home"), ("Pilar", "Historico Pilar"), ("Clubes", "Historico Clubes"), ("Red", "Historico Red")]

meta = svc.get(spreadsheetId=SPREADSHEET_ID).execute()
existing = {s["properties"]["title"] for s in meta["sheets"]}

for op, hist in pares:
    print(f"\n{'='*70}")
    print(f"  {op}  vs  {hist}")
    print("="*70)
    if op not in existing:
        print(f"  !!! {op} no existe"); continue
    if hist not in existing:
        print(f"  !!! {hist} no existe"); continue
    ro = svc.values().get(spreadsheetId=SPREADSHEET_ID, range=f"'{op}'!1:1").execute().get("values", [[]])[0]
    rh = svc.values().get(spreadsheetId=SPREADSHEET_ID, range=f"'{hist}'!1:1").execute().get("values", [[]])[0]
    # Last row con data en hist
    lr = svc.values().get(spreadsheetId=SPREADSHEET_ID, range=f"'{hist}'!A:A").execute()
    last = len(lr.get("values", [])) if lr.get("values") else 0
    n = max(len(ro), len(rh))
    diff = 0
    for i in range(n):
        a = str(ro[i]).strip() if i < len(ro) else ""
        b = str(rh[i]).strip() if i < len(rh) else ""
        mark = "  " if a == b else "XX"
        if a != b: diff += 1
        print(f"  {mark} {col_letter(i+1):>3}  op:{a!r:<40}  hist:{b!r}")
    print(f"\n  Op: {len(ro)} cols | Hist: {len(rh)} cols | Histórico tiene {last} filas (incl header)")
    print(f"  Diferencias: {diff}")
