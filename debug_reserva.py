"""Audita estado de pedidos y fórmulas de Reservado en hoja Productos."""
import google.auth
from googleapiclient.discovery import build
import json

creds, _ = google.auth.default(scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
svc = build('sheets', 'v4', credentials=creds)
SID = '1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY'

HOME_ABBRS = ['PPM','PPJyQ','PPCyQ','SCo','SJyQ','SCa','ECaC','EJyQ','ECyQ','EV','TG','TLC','TC','F','PMu','PMa','PJyQ','PCC','PJyM']
RED_ABBRS = ['PPM','PPJyQ','PPCyQ','SQB','SL','SCo','SPyP','SJyQ','SE','SCa','ECaC','EJyQ','ECyQ','EV','TG','TLC','TC','F','PMu','PMa','PJyQ','PCC','PJyM']

# Buscar todos los Reservados activos en Home, Pilar, Clubes, Red
for sh, dee_col, est_col, det_col, abbrs, prod_start in [
    ('Home',   'I', 'K', 'BB', HOME_ABBRS, 22),  # cols I=Origen, K=Estado, BB=OrigenDetalle, V-AO=productos (idx 22+)
    ('Pilar',  'I', 'K', 'BE', RED_ABBRS,  22),
    ('Clubes', 'L', 'N', 'AJ', None,       23),
    ('Red',    'J', 'L', 'BD', RED_ABBRS,  21),
]:
    r = svc.spreadsheets().values().get(spreadsheetId=SID, range=f'{sh}!A2:BZ', valueRenderOption='UNFORMATTED_VALUE').execute()
    rows = r.get('values', [])
    print(f'\n=== {sh}: pedidos Reservados ===')
    for i, row in enumerate(rows, start=2):
        # Indices de origen/estado/detalle
        ICOL = ord(dee_col)-65 if len(dee_col)==1 else (ord(dee_col[0])-64)*26+ord(dee_col[1])-65
        ECOL = ord(est_col)-65 if len(est_col)==1 else (ord(est_col[0])-64)*26+ord(est_col[1])-65
        DCOL = (ord(det_col[0])-64)*26+ord(det_col[1])-65 if len(det_col)==2 else ord(det_col)-65
        if len(row) <= max(ICOL, ECOL): continue
        est = str(row[ECOL]).strip() if len(row)>ECOL else ''
        if est != 'Reservado': continue
        org = str(row[ICOL]).strip() if len(row)>ICOL else ''
        det = str(row[DCOL]).strip() if len(row)>DCOL else ''
        cli = str(row[7]).strip() if len(row)>7 else ''  # H = cliente (Home/Pilar/Clubes); en Red I=cliente, H=vendedor
        if sh == 'Red':
            cli = str(row[8]).strip() if len(row)>8 else ''
        nped = row[1] if len(row)>1 else ''
        print(f'  fila {i}  N={nped}  cli={cli[:30]:30}  origen={org}  detalle={det[:80]}')

# Reservado actual por producto (fórmulas calculadas)
print('\n=== Productos: Reservado actual (col G) ===')
r2 = svc.spreadsheets().values().get(spreadsheetId=SID, range='Productos!A1:H', valueRenderOption='UNFORMATTED_VALUE').execute()
prods = r2.get('values', [])
for prow in prods[1:]:
    if len(prow) >= 7:
        ab = str(prow[2]).strip() if len(prow)>2 else ''
        nm = str(prow[1]).strip() if len(prow)>1 else ''
        st = prow[5] if len(prow)>5 else 0
        rs = prow[6] if len(prow)>6 else 0
        ds = prow[7] if len(prow)>7 else 0
        if rs and rs != 0:
            print(f'  {ab:7} ({nm[:25]:25}) Stock={st!s:>4} Reservado={rs!s:>4} Disponible={ds!s:>4}')
