import google.auth
from googleapiclient.discovery import build

creds, _ = google.auth.default(scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
svc = build('sheets', 'v4', credentials=creds)
SID = '1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY'

r = svc.spreadsheets().values().get(spreadsheetId=SID, range='Home!258:258', valueRenderOption='UNFORMATTED_VALUE').execute()
row = r.get('values', [[]])[0]
def safe(v): return v if v is not None else ''
print('Home row 258:')
print(f"  N pedido (B): {safe(row[1]) if len(row)>1 else ''}")
print(f"  Cliente (H):  {safe(row[7]) if len(row)>7 else ''}")
print(f"  Origen (I):   {safe(row[8]) if len(row)>8 else ''}")
print(f"  Estado (K):   {safe(row[10]) if len(row)>10 else ''}")
print(f"  Origen Detalle (BB): {safe(row[53]) if len(row)>53 else ''}")

HOME_ABBRS = ['PPM','PPJyQ','PPCyQ','SCo','SJyQ','SCa','ECaC','EJyQ','ECyQ','EV','TG','TLC','TC','F','PMu','PMa','PJyQ','PCC','PJyM']
print('  Productos:')
for i, ab in enumerate(HOME_ABBRS):
    q = row[22+i] if len(row) > 22+i else 0
    if q:
        print(f'    {ab}: {q}')

print()
r2 = svc.spreadsheets().values().get(spreadsheetId=SID, range='Productos!A1:H', valueRenderOption='UNFORMATTED_VALUE').execute()
prods = r2.get('values', [])
print('Productos del pedido en hoja Productos (Reservado col G):')
import json
od_str = safe(row[53]) if len(row) > 53 else ''
od = {}
try: od = json.loads(od_str) if od_str else {}
except: od = {}
prods_pedido = [ab for ab in HOME_ABBRS if (row[22+HOME_ABBRS.index(ab)] if len(row)>22+HOME_ABBRS.index(ab) else 0)]
for prow in prods[1:]:
    if len(prow) < 3: continue
    ab = str(prow[2]).strip()
    if ab in prods_pedido:
        stock = prow[5] if len(prow) > 5 else 0
        res = prow[6] if len(prow) > 6 else 0
        disp = prow[7] if len(prow) > 7 else 0
        flag = od.get(ab, '?')
        print(f'  {ab:7} flag={flag}  Stock={stock!s:>4}  Reservado={res!s:>4}  Disponible={disp!s:>4}')
