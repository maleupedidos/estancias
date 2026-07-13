"""Verify no 'Depósito' remains in any sheet, and re-run residual updates if needed."""
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SPREADSHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
OLD = "Depósito"
NEW = "Deposito"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file(SA_KEY, scopes=SCOPES)
svc = build("sheets", "v4", credentials=creds)
SHEETS = svc.spreadsheets()

meta = SHEETS.get(spreadsheetId=SPREADSHEET_ID, includeGridData=False).execute()
all_sheets_names = [s["properties"]["title"] for s in meta["sheets"]]

print(f"Escaneo total en {len(all_sheets_names)} hojas")
total_residuals = 0
residual_updates = []

for sheet_name in all_sheets_names:
    try:
        res = SHEETS.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{sheet_name}'",
            valueRenderOption="FORMATTED_VALUE",
        ).execute()
    except Exception as e:
        print(f"  {sheet_name}: error → {e}")
        continue
    rows = res.get("values", [])

    found = 0
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            if val == OLD:
                col_letter = ""
                n = c_idx
                while True:
                    col_letter = chr(ord("A") + (n % 26)) + col_letter
                    n = n // 26 - 1
                    if n < 0:
                        break
                residual_updates.append({
                    "range": f"'{sheet_name}'!{col_letter}{r_idx+1}",
                    "values": [[NEW]],
                })
                found += 1
    if found:
        print(f"  {sheet_name}: {found} celdas con '{OLD}'")
        total_residuals += found

print(f"\nTotal residuales encontrados: {total_residuals}")

if residual_updates:
    # Apply in chunks of 1000
    for i in range(0, len(residual_updates), 1000):
        chunk = residual_updates[i:i+1000]
        resp = SHEETS.values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"valueInputOption": "RAW", "data": chunk}
        ).execute()
        print(f"  Chunk {i//1000+1}: {resp.get('totalUpdatedCells', 0)} celdas actualizadas")

print("\nDone.")
