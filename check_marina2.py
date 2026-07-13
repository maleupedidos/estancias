"""Diagnóstico v2 — todas las filas Red de Marina Donati con detalle completo."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SPREADSHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

red_data = svc.values().get(spreadsheetId=SPREADSHEET_ID, range="'Red'!A1:BC").execute().get("values", [])
red_headers = red_data[0]

def col_letter(i):
    if i<26: return chr(65+i)
    return chr(65 + i//26 - 1) + chr(65 + i%26)

# Headers para referencia
print("=== HEADERS RED ===")
for i, h in enumerate(red_headers):
    print(f"  {col_letter(i):>3} ({i+1:>2}) {h}")
print()

# Buscar Marina Donati en col Cliente — necesito saber qué col es Cliente
hdr_low = [h.strip().lower() for h in red_headers]
i_cli = hdr_low.index("cliente") if "cliente" in hdr_low else -1
print(f"Col Cliente = {col_letter(i_cli)} (idx {i_cli+1})")
print()

print("=== FILAS RED de Marina Donati ===")
for idx, r in enumerate(red_data[1:], start=2):
    cli = (r[i_cli] if i_cli>=0 and i_cli<len(r) else "")
    if "donati" in cli.lower():
        print(f"--- Fila {idx} ---")
        for ci, h in enumerate(red_headers):
            v = r[ci] if ci < len(r) else ""
            if v not in ("", "0", 0):
                print(f"  {col_letter(ci):>3} {h:<35}: {v}")
        print()
