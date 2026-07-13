"""
Update Google Sheets to replace 'Depósito' (con acento) with 'Deposito' (sin acento):
- Cell values
- Data validation (dropdowns) in Origen columns (+ Canal in OC)
- Conditional formatting rules

Sheets afectadas: Home (I), Pilar (I), Capital Federal (I), Clubes (L), Orden de Compra (T + E + I)
"""
import json
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

# ── 1. Get spreadsheet metadata (sheet IDs, titles) ───────────────────────────
meta = SHEETS.get(
    spreadsheetId=SPREADSHEET_ID,
    includeGridData=False
).execute()

sheet_info = {}
for sh in meta["sheets"]:
    props = sh["properties"]
    sheet_info[props["title"]] = {
        "sheetId": props["sheetId"],
        "rowCount": props["gridProperties"]["rowCount"],
        "colCount": props["gridProperties"]["columnCount"],
    }

print("Hojas encontradas:")
for name, info in sheet_info.items():
    print(f"  {name}: sheetId={info['sheetId']} rows={info['rowCount']} cols={info['colCount']}")

# Configuración de hojas a procesar
# (sheet_name, origen_col_index_0based, canal_col_index_0based or None)
TARGETS = [
    ("Home", 8, None),              # I (col 9 = idx 8)
    ("Pilar", 8, None),             # I
    ("Capital Federal", 8, None),   # I
    ("Clubes", 11, None),           # L (col 12 = idx 11)
    ("Orden de Compra", 19, 4),     # T (col 20 = idx 19), Canal col E (col 5 = idx 4)
]

# ── 2. Update cell VALUES: find every cell == "Depósito" and replace ──────────
print("\n── PASO 1: Actualizar valores de celda (Depósito → Deposito) ──")

value_updates = []  # list of (range, new_values)
total_value_changes = 0

for sheet_name, *_ in TARGETS:
    if sheet_name not in sheet_info:
        print(f"  ADVERTENCIA: hoja '{sheet_name}' no existe, skip")
        continue

    # Read whole sheet content with FORMATTED_VALUE
    res = SHEETS.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'",
        valueRenderOption="FORMATTED_VALUE",
    ).execute()
    rows = res.get("values", [])

    changed_cells = []
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            if val == OLD:
                # A1 notation column letter
                col_letter = ""
                n = c_idx
                while True:
                    col_letter = chr(ord("A") + (n % 26)) + col_letter
                    n = n // 26 - 1
                    if n < 0:
                        break
                a1 = f"'{sheet_name}'!{col_letter}{r_idx+1}"
                changed_cells.append((a1, NEW))

    print(f"  {sheet_name}: {len(changed_cells)} celdas con valor '{OLD}'")
    for a1, newval in changed_cells:
        value_updates.append({"range": a1, "values": [[newval]]})
    total_value_changes += len(changed_cells)

# Apply with batchUpdate
if value_updates:
    resp = SHEETS.values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"valueInputOption": "RAW", "data": value_updates}
    ).execute()
    print(f"  ACTUALIZADAS: {resp.get('totalUpdatedCells', 0)} celdas")
else:
    print("  No había celdas con 'Depósito' para cambiar.")

# ── 3. Get ALL sheets with grid data to analyze validations + conditional formats ─
print("\n── PASO 2: Actualizar data validations (dropdowns) ──")

meta_full = SHEETS.get(
    spreadsheetId=SPREADSHEET_ID,
    includeGridData=True,
    ranges=[f"'{s}'" for s, *_ in TARGETS],
).execute()

requests_batch = []

for sheet_data in meta_full["sheets"]:
    props = sheet_data["properties"]
    sheet_name = props["title"]
    sheet_id = props["sheetId"]

    target = None
    for t in TARGETS:
        if t[0] == sheet_name:
            target = t
            break
    if not target:
        continue

    _, origen_col, canal_col = target
    max_rows = props["gridProperties"]["rowCount"]

    # Process Origen column
    cols_to_process = [("Origen", origen_col)]
    if canal_col is not None:
        cols_to_process.append(("Canal", canal_col))

    for col_label, col_idx in cols_to_process:
        # Inspect first-data-row validation (row 2 = idx 1) — assume uniform
        sample_validation = None
        grid_data = sheet_data.get("data", [])
        if grid_data and "rowData" in grid_data[0]:
            for rowData in grid_data[0]["rowData"][1:200]:  # check first 200 data rows
                cells = rowData.get("values", [])
                if col_idx < len(cells):
                    dv = cells[col_idx].get("dataValidation")
                    if dv:
                        sample_validation = dv
                        break

        if sample_validation:
            condition = sample_validation.get("condition", {})
            ctype = condition.get("type")
            values = condition.get("values", [])
            literals = [v.get("userEnteredValue", "") for v in values]
            print(f"  {sheet_name}.{col_label} (col {col_idx+1}): validación actual {ctype} = {literals}")

            if OLD in literals:
                new_literals = [NEW if v == OLD else v for v in literals]
                print(f"    → Reemplazo por: {new_literals}")

                # Build new validation with full column range rows 2..max
                new_rule = {
                    "condition": {
                        "type": ctype,
                        "values": [{"userEnteredValue": v} for v in new_literals],
                    },
                    "strict": sample_validation.get("strict", True),
                    "showCustomUi": sample_validation.get("showCustomUi", True),
                }
                if "inputMessage" in sample_validation:
                    new_rule["inputMessage"] = sample_validation["inputMessage"]

                requests_batch.append({
                    "setDataValidation": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 1,          # skip header
                            "endRowIndex": max_rows,
                            "startColumnIndex": col_idx,
                            "endColumnIndex": col_idx + 1,
                        },
                        "rule": new_rule,
                    }
                })
            else:
                print(f"    (no contiene '{OLD}', skip)")
        else:
            print(f"  {sheet_name}.{col_label} (col {col_idx+1}): sin data validation en filas examinadas")

# ── 4. Update conditional formatting rules ────────────────────────────────────
print("\n── PASO 3: Actualizar reglas de formato condicional ──")

# Need per-sheet conditionalFormats
for sheet_data in meta_full["sheets"]:
    props = sheet_data["properties"]
    sheet_name = props["title"]
    sheet_id = props["sheetId"]

    if sheet_name not in [t[0] for t in TARGETS]:
        continue

    cond_formats = sheet_data.get("conditionalFormats", [])
    # Alternative location
    if not cond_formats:
        # conditionalFormats are actually at sheet level, not in data — re-fetch
        pass

print("  Re-obteniendo formatos condicionales…")

# Fetch conditional formats per-sheet (they're at the sheet level, not within grid data in the above call)
meta2 = SHEETS.get(
    spreadsheetId=SPREADSHEET_ID,
    fields="sheets(properties(sheetId,title),conditionalFormats)",
).execute()

rules_to_update = []  # (sheetId, ruleIndex, newRule)

for sheet_data in meta2["sheets"]:
    props = sheet_data["properties"]
    sheet_name = props["title"]
    sheet_id = props["sheetId"]

    if sheet_name not in [t[0] for t in TARGETS]:
        continue

    cfs = sheet_data.get("conditionalFormats", [])
    print(f"  {sheet_name}: {len(cfs)} reglas de formato condicional")

    for idx, rule in enumerate(cfs):
        boolean_rule = rule.get("booleanRule")
        if not boolean_rule:
            continue
        condition = boolean_rule.get("condition", {})
        ctype = condition.get("type", "")
        values = condition.get("values", [])

        # Only text-equal type that matches "Depósito"
        for v in values:
            if v.get("userEnteredValue") == OLD:
                print(f"    Regla #{idx}: type={ctype} matches '{OLD}' — se actualizará")
                # Build updated rule
                new_condition = dict(condition)
                new_condition["values"] = [
                    {"userEnteredValue": NEW} if vv.get("userEnteredValue") == OLD else vv
                    for vv in values
                ]
                new_boolean_rule = dict(boolean_rule)
                new_boolean_rule["condition"] = new_condition
                new_rule = dict(rule)
                new_rule["booleanRule"] = new_boolean_rule
                rules_to_update.append((sheet_id, idx, new_rule))
                break

# Apply updates — use updateConditionalFormatRule
for sheet_id, idx, new_rule in rules_to_update:
    requests_batch.append({
        "updateConditionalFormatRule": {
            "sheetId": sheet_id,
            "index": idx,
            "rule": new_rule,
        }
    })

print(f"\n  Total reglas cond. format. a actualizar: {len(rules_to_update)}")

# ── 5. Execute all batch requests ─────────────────────────────────────────────
print(f"\n── PASO 4: Ejecutar batchUpdate con {len(requests_batch)} requests ──")

if requests_batch:
    resp = SHEETS.batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": requests_batch}
    ).execute()
    print(f"  OK — {len(resp.get('replies', []))} respuestas")
else:
    print("  No hay requests que aplicar.")

# ── 6. Also check Productos sheet and any other sheet for 'Depósito' ──────────
print("\n── PASO 5: Buscar 'Depósito' en TODAS las hojas (incl. Productos) ──")

all_sheets_names = list(sheet_info.keys())
residual_updates = []
for sheet_name in all_sheets_names:
    try:
        res = SHEETS.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{sheet_name}'",
            valueRenderOption="FORMATTED_VALUE",
        ).execute()
    except Exception as e:
        print(f"  {sheet_name}: error al leer → {e}")
        continue
    rows = res.get("values", [])

    found = []
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
                found.append((r_idx + 1, col_letter))
                residual_updates.append({
                    "range": f"'{sheet_name}'!{col_letter}{r_idx+1}",
                    "values": [[NEW]],
                })
    if found:
        print(f"  {sheet_name}: {len(found)} celdas restantes con '{OLD}': {found[:10]}")
    else:
        print(f"  {sheet_name}: sin '{OLD}' (OK)")

if residual_updates:
    resp = SHEETS.values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"valueInputOption": "RAW", "data": residual_updates}
    ).execute()
    print(f"  Actualizadas {resp.get('totalUpdatedCells', 0)} celdas residuales")

print("\nDone!")
