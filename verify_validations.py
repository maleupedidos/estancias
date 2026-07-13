"""Verify data validations and conditional formats no longer contain 'Depósito'."""
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

TARGET_SHEETS = ["Home", "Pilar", "Capital Federal", "Clubes", "Orden de Compra"]

# Check conditional formats
print("── Formatos condicionales ──")
meta2 = SHEETS.get(
    spreadsheetId=SPREADSHEET_ID,
    fields="sheets(properties(sheetId,title),conditionalFormats)",
).execute()

total_old = 0
for sheet_data in meta2["sheets"]:
    title = sheet_data["properties"]["title"]
    if title not in TARGET_SHEETS:
        continue
    cfs = sheet_data.get("conditionalFormats", [])
    for idx, rule in enumerate(cfs):
        boolean_rule = rule.get("booleanRule")
        if not boolean_rule:
            continue
        condition = boolean_rule.get("condition", {})
        values = condition.get("values", [])
        for v in values:
            uv = v.get("userEnteredValue", "")
            if OLD in uv:
                print(f"  {title} regla #{idx}: AÚN CONTIENE '{OLD}': {uv}")
                total_old += 1
            elif uv == NEW:
                print(f"  {title} regla #{idx}: OK ({NEW})")

print(f"\nReglas con '{OLD}' pendientes: {total_old}")

# Check data validations by reading rows
print("\n── Data validations (filas 2-5 como muestra) ──")
meta_full = SHEETS.get(
    spreadsheetId=SPREADSHEET_ID,
    includeGridData=True,
    ranges=[f"'{s}'!A2:Z5" for s in TARGET_SHEETS],
).execute()

cols = {"Home": 8, "Pilar": 8, "Capital Federal": 8, "Clubes": 11, "Orden de Compra": 19}

for sheet_data in meta_full["sheets"]:
    title = sheet_data["properties"]["title"]
    if title not in cols:
        continue
    col_idx = cols[title]
    grid = sheet_data.get("data", [{}])[0]
    rowData = grid.get("rowData", [])
    for rnum, row in enumerate(rowData):
        cells = row.get("values", [])
        if col_idx < len(cells):
            dv = cells[col_idx].get("dataValidation")
            if dv:
                vals = [v.get("userEnteredValue", "") for v in dv["condition"].get("values", [])]
                has_old = OLD in vals
                print(f"  {title} fila {rnum+2} col {col_idx+1}: {vals} {'❌ OLD' if has_old else '✓'}")
                break  # just one sample per sheet
print("\nDone.")
