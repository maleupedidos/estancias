"""Alinea Historico Home/Pilar/Clubes con el layout operativo.
Mapea cada fila por nombre de header (con aliases) y la re-escribe en el orden correcto.
Preserva los datos existentes, solo reordena."""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

# Alias: nombres equivalentes entre historico (viejo) y operativo (nuevo)
ALIAS = {
    "Hora Pedido": "Hora",
    "N° Pedido": "N°",
    "Día Pedido": "Día",
    "Fecha Pedido": "Fecha",
    "Mes Pedido": "Mes",
    "Semana Pedido": "Semana",
    "Año Pedido": "Año",
}

def canon(h):
    h = str(h).strip()
    return ALIAS.get(h, h)

def col_letter(n):
    s=""
    while n>0:
        n,rem=divmod(n-1,26); s=chr(65+rem)+s
    return s

def read_vals(name, mode="UNFORMATTED_VALUE"):
    r = svc.values().get(spreadsheetId=SID, range=f"'{name}'!A:BZ", valueRenderOption=mode).execute()
    return r.get("values", [])

def ensure_size(sheet_id, cols_needed, rows_needed):
    meta = svc.get(spreadsheetId=SID).execute()
    for s in meta["sheets"]:
        if s["properties"]["sheetId"] == sheet_id:
            g = s["properties"]["gridProperties"]
            reqs = []
            if g["columnCount"] < cols_needed:
                reqs.append({"appendDimension": {"sheetId": sheet_id, "dimension": "COLUMNS", "length": cols_needed - g["columnCount"]}})
            if g["rowCount"] < rows_needed:
                reqs.append({"appendDimension": {"sheetId": sheet_id, "dimension": "ROWS", "length": rows_needed - g["rowCount"]}})
            if reqs:
                svc.batchUpdate(spreadsheetId=SID, body={"requests": reqs}).execute()
            return

# Obtener meta para sheetIds
meta = svc.get(spreadsheetId=SID).execute()
sheet_ids = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}

def migrar(op_name, hist_name):
    print(f"\n{'='*70}\n  {hist_name} <- alineando a layout de {op_name}\n{'='*70}")
    op_rows = read_vals(op_name)
    if not op_rows:
        print(f"  !!! {op_name} vacía, skip"); return
    op_headers = [str(h).strip() for h in op_rows[0]]
    hist_rows = read_vals(hist_name)
    if not hist_rows:
        print(f"  !!! {hist_name} vacía, solo escribo headers")
        hist_headers_old = []
        hist_data = []
    else:
        hist_headers_old = [canon(h) for h in hist_rows[0]]
        hist_data = hist_rows[1:]
    # Construir mapeo: para cada col operativa, buscar su posición en hist
    mapping = []  # idx en hist_data (None si no existe)
    for h in op_headers:
        idx = hist_headers_old.index(h) if h in hist_headers_old else None
        mapping.append(idx)
    # Reconstruir datos
    new_data = []
    for row in hist_data:
        new_row = []
        for idx in mapping:
            v = ""
            if idx is not None and idx < len(row):
                v = row[idx]
            new_row.append(v if v is not None else "")
        new_data.append(new_row)
    # Asegurar tamaño de grid
    ensure_size(sheet_ids[hist_name], len(op_headers), max(len(new_data)+10, 100))
    # Limpiar TODO el contenido
    svc.values().clear(spreadsheetId=SID, range=f"'{hist_name}'!A:BZ").execute()
    # Escribir headers + data
    body = {"values": [op_headers] + new_data}
    svc.values().update(spreadsheetId=SID, range=f"'{hist_name}'!A1", valueInputOption="RAW", body=body).execute()
    print(f"  OK: {len(new_data)} filas re-escritas en {len(op_headers)} columnas")
    # Formato header (marrón/blanco/bold)
    svc.batchUpdate(spreadsheetId=SID, body={"requests": [{
        "repeatCell": {
            "range": {"sheetId": sheet_ids[hist_name], "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": len(op_headers)},
            "cell": {"userEnteredFormat": {"backgroundColor": {"red": 0.365, "green": 0.263, "blue": 0.216}, "textFormat": {"foregroundColor": {"red": 1, "green": 1, "blue": 1}, "bold": True}}},
            "fields": "userEnteredFormat(backgroundColor,textFormat)",
        }
    }, {"updateSheetProperties": {"properties": {"sheetId": sheet_ids[hist_name], "gridProperties": {"frozenRowCount": 1}}, "fields": "gridProperties.frozenRowCount"}}]}).execute()

migrar("Home", "Historico Home")
migrar("Clubes", "Historico Clubes")
migrar("Pilar", "Historico Pilar")
print("\nListo. Red se hace aparte por layout muy distinto.")
