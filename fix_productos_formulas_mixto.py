"""Reescribe Reservado (G) y Vendidos Semana (E) en hoja Productos para
que cuenten Origen=Deposito Y Origen=Mixto (con producto en flag "D" del JSON)."""
import google.auth
from googleapiclient.discovery import build
from datetime import date

creds, _ = google.auth.default(scopes=['https://www.googleapis.com/auth/spreadsheets'])
svc = build('sheets', 'v4', credentials=creds)
SID = '1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY'

HOME = {'PPM':'W','PPJyQ':'X','PPCyQ':'Y','SCo':'Z','SJyQ':'AA','SCa':'AB','ECaC':'AC','EJyQ':'AD','ECyQ':'AE','EV':'AF','TG':'AG','TLC':'AH','TC':'AI','F':'AJ','PMu':'AK','PMa':'AL','PJyQ':'AM','PCC':'AN','PJyM':'AO'}
PILAR = {'PPM':'W','PPJyQ':'X','PPCyQ':'Y','SQB':'Z','SL':'AA','SCo':'AB','SPyP':'AC','SJyQ':'AD','SE':'AE','SCa':'AF','ECaC':'AG','EJyQ':'AH','ECyQ':'AI','EV':'AJ','TG':'AK','TLC':'AL','TC':'AM','F':'AN','PMu':'AO','PMa':'AP','PJyQ':'AQ','PCC':'AR','PJyM':'AS'}
CLUB = {'PMu':'X','PMa':'Y','PJyQ':'Z','PCC':'AA','PJyM':'AB','PPM':'AC','PPJyQ':'AD','PPCyQ':'AE'}
# Red v3: V(22)→AR(44) según RED_COL_TO_ABBR del backend
RED  = {'PPM':'V','PPJyQ':'W','PPCyQ':'X','SQB':'Y','SL':'Z','SCo':'AA','SPyP':'AB','SJyQ':'AC','SE':'AD','SCa':'AE','ECaC':'AF','EJyQ':'AG','ECyQ':'AH','EV':'AI','TG':'AJ','TLC':'AK','TC':'AL','F':'AM','PMu':'AN','PMa':'AO','PJyQ':'AP','PCC':'AQ','PJyM':'AR'}

hoy = date.today()
sem = hoy.isocalendar()[1]
anio = hoy.year

# Locale español usa ';' como separador. La cantidad reservada/vendida desde Depósito
# para Mixto se extrae del JSON "Origen Detalle":
#   Formato nuevo: {"PMu":{"d":5,"oc":3},"PJyQ":{"d":0,"oc":2}}
#   Formato viejo: {"PMu":"D","PJyQ":"OC"}  (ya descontinuado pero se mantiene compat)
#
# Usamos REGEXEXTRACT para sacar el "d":N. Si no matchea, IFERROR devuelve "0" y VALUE→0.
# Para Origen=Deposito: cuenta la cantidad TOTAL del producto en el pedido (sin parsear JSON).
def dep_term(orig_full, det_full, abbr, qty_full):
    # Patrón regex: "ABBR":{"d":(\d+) → captura el número
    pat = f'""{abbr}"":\\{{""d"":(\\d+)'
    return (
        f'(({orig_full}="Deposito")*{qty_full}'
        f'+({orig_full}="Mixto")*IFERROR(VALUE(REGEXEXTRACT({det_full};"{pat}"));0))'
    )

# Leer abbrs por fila
res = svc.spreadsheets().values().get(spreadsheetId=SID, range='Productos!A1:C').execute()
rows = res.get('values', [])
prod_abbrs = []
for i, row in enumerate(rows[1:], start=2):
    if len(row) >= 3:
        ab = str(row[2]).strip()
        if ab:
            prod_abbrs.append((i, ab))

# Helper para construir la cantidad a sumar para una hoja+abbr.
# Soporta ambos formatos de Origen Detalle:
#   Nuevo: {"abbr":{"d":N,"oc":Y}}  → REGEXEXTRACT extrae N
#   Viejo: {"abbr":"D"}              → REGEXMATCH detecta "D" y cuenta cantidad total
def _cant_sumar(orig, det, qty_col, abbr):
    # Nuevo formato JSON: {"abbr":{"d":N,"oc":Y}} — REGEXEXTRACT extrae N
    pat_new = f'""{abbr}"":\\{{""d"":(\\d+)'
    # Viejo formato JSON: {"abbr":"D"} — REGEXMATCH detecta "abbr":"D"
    pat_old = f'""{abbr}"":""D""'
    return (
        f'(({orig}="Deposito")*{qty_col}'
        f'+({orig}="Mixto")*('
        f'IFERROR(VALUE(REGEXEXTRACT({det};"{pat_new}"));0)'
        f'+IFERROR(REGEXMATCH({det};"{pat_old}");FALSE)*{qty_col}'
        f'))'
    )

def vendidos(ab):
    parts = []
    if ab in HOME:
        c = HOME[ab]; q = f'Home!{c}$2:{c}$10000'
        cs = _cant_sumar('Home!$I$2:$I$10000', 'Home!$BB$2:$BB$10000', q, ab)
        parts.append(f'SUMPRODUCT({cs}*(Home!$K$2:$K$10000="Entregado")*(Home!$AZ$2:$AZ$10000={sem})*(Home!$BA$2:$BA$10000={anio}))')
    if ab in PILAR:
        c = PILAR[ab]; q = f'Pilar!{c}$2:{c}$10000'
        cs = _cant_sumar('Pilar!$I$2:$I$10000', 'Pilar!$BE$2:$BE$10000', q, ab)
        parts.append(f'SUMPRODUCT({cs}*(Pilar!$K$2:$K$10000="Entregado")*(Pilar!$BC$2:$BC$10000={sem})*(Pilar!$BD$2:$BD$10000={anio}))')
    if ab in CLUB:
        c = CLUB[ab]; q = f'Clubes!{c}$2:{c}$10000'
        cs = _cant_sumar('Clubes!$L$2:$L$10000', 'Clubes!$AJ$2:$AJ$10000', q, ab)
        parts.append(f'SUMPRODUCT({cs}*(Clubes!$N$2:$N$10000="Entregado"))')
    if ab in RED:
        c = RED[ab]; q = f'Red!{c}$2:{c}$10000'
        cs = _cant_sumar('Red!$J$2:$J$10000', 'Red!$BD$2:$BD$10000', q, ab)
        parts.append(f'SUMPRODUCT({cs}*(Red!$L$2:$L$10000="Entregado"))')
    return '=' + '+'.join(parts) if parts else ''

def reservado(ab):
    parts = []
    if ab in HOME:
        c = HOME[ab]; q = f'Home!{c}$2:{c}$10000'
        cs = _cant_sumar('Home!$I$2:$I$10000', 'Home!$BB$2:$BB$10000', q, ab)
        parts.append(f'SUMPRODUCT({cs}*(Home!$K$2:$K$10000="Reservado"))')
    if ab in PILAR:
        c = PILAR[ab]; q = f'Pilar!{c}$2:{c}$10000'
        cs = _cant_sumar('Pilar!$I$2:$I$10000', 'Pilar!$BE$2:$BE$10000', q, ab)
        parts.append(f'SUMPRODUCT({cs}*(Pilar!$K$2:$K$10000="Reservado"))')
    if ab in CLUB:
        c = CLUB[ab]; q = f'Clubes!{c}$2:{c}$10000'
        cs = _cant_sumar('Clubes!$L$2:$L$10000', 'Clubes!$AJ$2:$AJ$10000', q, ab)
        parts.append(f'SUMPRODUCT({cs}*(Clubes!$N$2:$N$10000="Reservado"))')
    if ab in RED:
        c = RED[ab]; q = f'Red!{c}$2:{c}$10000'
        cs = _cant_sumar('Red!$J$2:$J$10000', 'Red!$BD$2:$BD$10000', q, ab)
        parts.append(f'SUMPRODUCT({cs}*(Red!$L$2:$L$10000="Reservado"))')
    return '=' + '+'.join(parts) if parts else ''

data = []
for r, ab in prod_abbrs:
    fv = vendidos(ab); fr = reservado(ab)
    if fv: data.append({'range': f'Productos!E{r}', 'values': [[fv]]})
    if fr: data.append({'range': f'Productos!G{r}', 'values': [[fr]]})

resp = svc.spreadsheets().values().batchUpdate(
    spreadsheetId=SID,
    body={'valueInputOption': 'USER_ENTERED', 'data': data}
).execute()
print(f'OK: {resp.get("totalUpdatedCells")} celdas actualizadas con REGEXMATCH+IFERROR.')
print(f'   {len(prod_abbrs)} productos. Semana ISO {sem} / Año {anio}.')
