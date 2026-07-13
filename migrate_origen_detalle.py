"""
Migración: Origen Detalle formato viejo → nuevo.

Formato viejo:  {"PPM":"D"}  o  {"PPM":"OC"}
Formato nuevo:  {"PPM":{"d":N,"oc":M}}  donde N+M = qty del producto en la fila.

Recorre Home, Pilar, Clubes, Red. Cada fila con Origen ∈ {Deposito, Mixto, Orden de Compra}
y Origen Detalle no vacío que sea formato viejo se reescribe en formato nuevo.

CONFIG: poner DRY_RUN = False para escribir.
"""
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import json

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SPREADSHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
DRY_RUN = True  # cambiar a False para aplicar

creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

def col_letter(i):
    if i < 26: return chr(65+i)
    return chr(65 + i//26 - 1) + chr(65 + i%26)

# Layout por hoja:
#   col_origen (1-based), col_origen_detalle (1-based o "auto"), col_estado (1-based),
#   prod_start_col_letter, prod_end_col_letter, abbr_to_col_idx mapping (computed from headers)
HOJAS = {
    "Home":   {"col_origen": 9,  "col_estado": 11},
    "Pilar":  {"col_origen": 9,  "col_estado": 11},
    "Clubes": {"col_origen": 12, "col_estado": 14},
    "Red":    {"col_origen": 10, "col_estado": 12},
}

# Productos: la abreviatura está en el header de cada hoja. Sacamos los rangos de cantidad
# leyendo las columnas que coinciden con abbreviaturas conocidas.
KNOWN_ABBRS = {
    "PPM","PPJyQ","PPCyQ","PMu","PMa","PJyQ","PCC","PJyM",
    "SQB","SL","SCo","SPyP","SJyQ","SE","SCa",
    "ECaC","EJyQ","ECyQ","EV",
    "TG","TLC","TC","F",
}

stats = {"checked":0, "viejo":0, "nuevo":0, "vacio":0, "migrados":0, "errores":0}
print(f"=== MIGRACIÓN ORIGEN DETALLE  (DRY_RUN={DRY_RUN}) ===\n")

for hoja, cfg in HOJAS.items():
    print(f"\n--- Hoja: {hoja} ---")
    data = svc.values().get(spreadsheetId=SPREADSHEET_ID, range=f"'{hoja}'!A1:BC").execute().get("values", [])
    if not data:
        print(f"  (vacía)")
        continue
    headers = data[0]

    # Localizar Origen Detalle (puede no existir en algunas hojas)
    try:
        idx_origen_detalle = [h.strip().lower() for h in headers].index("origen detalle")
        col_origen_detalle = idx_origen_detalle + 1  # 1-based
    except ValueError:
        print(f"  ⚠ Sin columna 'Origen Detalle'. Skipping.")
        continue

    # Mapeo abbr → col_idx (1-based) para la hoja
    abbr_to_col = {}
    for i, h in enumerate(headers):
        if h and h.strip() in KNOWN_ABBRS:
            abbr_to_col[h.strip()] = i + 1

    if not abbr_to_col:
        print(f"  ⚠ Sin columnas de productos detectadas. Skipping.")
        continue

    print(f"  Origen Detalle col: {col_letter(col_origen_detalle-1)} ({col_origen_detalle}) | productos detectados: {len(abbr_to_col)}")

    updates = []
    for r_idx, row in enumerate(data[1:], start=2):
        if r_idx == 2 and not row: continue  # tolerar fila vacía
        get = lambda i: (row[i-1] if (i-1) < len(row) else "")

        origen = str(get(cfg["col_origen"])).strip()
        if origen not in ("Deposito", "Mixto", "Orden de Compra"):
            continue

        det_str = str(get(col_origen_detalle)).strip()
        stats["checked"] += 1

        if not det_str:
            stats["vacio"] += 1
            continue

        try:
            det = json.loads(det_str)
        except Exception:
            print(f"  Fila {r_idx}: JSON inválido — {det_str[:80]}")
            stats["errores"] += 1
            continue

        # Detectar formato
        es_viejo = False
        es_nuevo = True
        for v in det.values():
            if isinstance(v, str):
                es_viejo = True
                es_nuevo = False
                break
            elif isinstance(v, dict) and ("d" in v or "oc" in v):
                continue
            else:
                es_nuevo = False
                break

        if es_nuevo and not es_viejo:
            stats["nuevo"] += 1
            continue

        if not es_viejo:
            print(f"  Fila {r_idx}: formato desconocido — {det_str[:80]}")
            stats["errores"] += 1
            continue

        # Migrar viejo → nuevo
        stats["viejo"] += 1
        nuevo = {}
        for abbr, marker in det.items():
            col_idx = abbr_to_col.get(abbr)
            if not col_idx:
                # producto no presente en la hoja — preservar como mejor se pueda
                m = str(marker).strip().upper()
                nuevo[abbr] = {"d": 0, "oc": 0}
                continue
            qty = 0
            try:
                qty = int(str(get(col_idx)).replace(",","").strip() or "0")
            except Exception:
                qty = 0
            m = str(marker).strip().upper()
            if m == "OC":
                nuevo[abbr] = {"d": 0, "oc": qty}
            else:  # 'D' o cualquier otro = todo Depósito
                nuevo[abbr] = {"d": qty, "oc": 0}

        nuevo_str = json.dumps(nuevo, ensure_ascii=False)
        cli = get(8) if hoja != "Red" else get(9)  # H=Cliente Home/Pilar/Clubes, I=Cliente Red
        print(f"  Fila {r_idx} ({cli}): {det_str[:60]}  →  {nuevo_str[:60]}")
        updates.append({
            "range": f"'{hoja}'!{col_letter(col_origen_detalle-1)}{r_idx}",
            "values": [[nuevo_str]]
        })
        stats["migrados"] += 1

    if updates and not DRY_RUN:
        result = svc.values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"valueInputOption": "RAW", "data": updates}
        ).execute()
        print(f"  ✓ {result.get('totalUpdatedCells')} celdas actualizadas")
    elif updates:
        print(f"  (dry-run) {len(updates)} celdas se actualizarían")

print(f"\n=== RESUMEN ===")
print(f"  Filas con origen confirmado: {stats['checked']}")
print(f"  Vacías:                      {stats['vacio']}")
print(f"  Formato viejo:               {stats['viejo']}")
print(f"  Formato nuevo (ya migradas): {stats['nuevo']}")
print(f"  Errores:                     {stats['errores']}")
print(f"  Migradas en esta corrida:    {stats['migrados']}{'  (dry-run)' if DRY_RUN else ''}")
