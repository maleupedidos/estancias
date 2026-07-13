"""
Diagnóstico de hojas auxiliares en Maleu - Pedidos
Hojas: Kardex, Log Errores, Resumen Diario, Alertas Stock, Productos
"""
import json
import re
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SPREADSHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

creds = Credentials.from_service_account_file(SA_KEY, scopes=SCOPES)
service = build("sheets", "v4", credentials=creds)
sheets = service.spreadsheets()

SHEETS_TO_READ = ["Kardex", "Log Errores", "Resumen Diario", "Alertas Stock", "Productos"]

ERROR_PATTERNS = re.compile(r"#(NAME\?|REF!|VALUE!|DIV/0!|N/A|NULL!|ERROR!|NUM!)")

def read_sheet(name):
    try:
        result = sheets.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{name}'!A:Z",
            valueRenderOption="FORMULA"
        ).execute()
        formula_values = result.get("values", [])
    except Exception as e:
        return None, None, str(e)

    try:
        result2 = sheets.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{name}'!A:Z",
            valueRenderOption="FORMATTED_VALUE"
        ).execute()
        display_values = result2.get("values", [])
    except:
        display_values = formula_values

    return formula_values, display_values, None

def print_separator():
    print("=" * 80)

# Read all sheets
all_data = {}
for sheet_name in SHEETS_TO_READ:
    formulas, display, error = read_sheet(sheet_name)
    all_data[sheet_name] = {"formulas": formulas, "display": display, "error": error}

    print_separator()
    print(f"HOJA: {sheet_name}")
    print_separator()

    if error:
        print(f"  ERROR al leer: {error}")
        continue

    if not display or len(display) == 0:
        print("  VACÍA - sin datos")
        continue

    headers = display[0] if display else []
    data_rows = display[1:] if len(display) > 1 else []

    print(f"  Headers: {headers}")
    print(f"  Filas de datos: {len(data_rows)}")

    # Show all rows (or first 50 if too many)
    max_show = 50
    for i, row in enumerate(data_rows[:max_show]):
        print(f"  Fila {i+1}: {row}")
    if len(data_rows) > max_show:
        print(f"  ... ({len(data_rows) - max_show} filas más)")

    # Check for formula errors in display values
    errors_found = []
    for r, row in enumerate(display):
        for c, cell in enumerate(row):
            if ERROR_PATTERNS.search(str(cell)):
                errors_found.append((r, c, cell))

    if errors_found:
        print(f"\n  ERRORES DE FÓRMULA DETECTADOS ({len(errors_found)}):")
        for r, c, val in errors_found[:20]:
            print(f"    Fila {r}, Col {c}: {val}")
    else:
        print("\n  Sin errores de fórmula en valores mostrados.")

    # Also check formulas themselves for errors
    if formulas:
        formula_errors = []
        for r, row in enumerate(formulas):
            for c, cell in enumerate(row):
                if ERROR_PATTERNS.search(str(cell)):
                    formula_errors.append((r, c, cell))
        if formula_errors:
            print(f"\n  ERRORES EN FÓRMULAS RAW ({len(formula_errors)}):")
            for r, c, val in formula_errors[:20]:
                print(f"    Fila {r}, Col {c}: {val}")

    # Show formulas for first few rows
    if formulas and len(formulas) > 0:
        print(f"\n  Fórmulas (headers): {formulas[0]}")
        for i, row in enumerate(formulas[1:5]):
            has_formula = any(str(c).startswith("=") for c in row)
            if has_formula:
                print(f"  Fórmulas fila {i+1}: {row}")

print_separator()
print("\nANÁLISIS CRUZADO: Productos vs Alertas Stock")
print_separator()

prod = all_data["Productos"]
alertas = all_data["Alertas Stock"]

if prod["display"] and len(prod["display"]) > 1:
    prod_headers = prod["display"][0]
    print(f"Productos headers: {prod_headers}")
    print(f"Productos filas: {len(prod["display"]) - 1}")

    # Show all product rows with stock info
    for i, row in enumerate(prod["display"][1:]):
        print(f"  Producto {i+1}: {row}")

    # Show formulas
    if prod["formulas"] and len(prod["formulas"]) > 1:
        print(f"\nProductos fórmulas (header): {prod["formulas"][0]}")
        for i, row in enumerate(prod["formulas"][1:3]):
            print(f"  Fórmulas producto {i+1}: {row}")

if alertas["display"] and len(alertas["display"]) > 1:
    print(f"\nAlertas headers: {alertas['display'][0]}")
    for i, row in enumerate(alertas["display"][1:]):
        print(f"  Alerta {i+1}: {row}")

# Summary
print_separator()
print("RESUMEN DE DIAGNÓSTICO")
print_separator()
for name in SHEETS_TO_READ:
    d = all_data[name]
    if d["error"]:
        status = f"ERROR: {d['error']}"
    elif not d["display"] or len(d["display"]) == 0:
        status = "VACÍA"
    elif len(d["display"]) == 1:
        status = "Solo headers, sin datos"
    else:
        status = f"OK - {len(d['display'])-1} filas de datos"
    print(f"  {name}: {status}")
