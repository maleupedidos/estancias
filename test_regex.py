"""Test si REGEXMATCH funciona dentro de SUMPRODUCT en Sheets."""
import google.auth
from googleapiclient.discovery import build
creds, _ = google.auth.default(scopes=['https://www.googleapis.com/auth/spreadsheets'])
svc = build('sheets', 'v4', credentials=creds)
SID = '1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY'

# Probar varias fórmulas en celdas auxiliares
tests = [
    ('Productos!N1', '=SUMPRODUCT(--(IFERROR(REGEXMATCH(Home!$BB$2:$BB$10000,"""PPJyQ"":""D"""),FALSE)))'),
    ('Productos!N2', '=SUMPRODUCT(IFERROR(REGEXMATCH(Home!$BB$2:$BB$10000,"""PPJyQ"":""D""")*1,0))'),
    ('Productos!N3', '=ARRAYFORMULA(SUM(IFERROR(REGEXMATCH(Home!$BB$2:$BB$10000,"""PPJyQ"":""D""")*1,0)))'),
    ('Productos!N4', '=SUMPRODUCT(--(IFERROR(SEARCH("""PPJyQ"":""D""",Home!$BB$2:$BB$10000),0)>0))'),
]
data = [{'range': c, 'values': [[f]]} for c, f in tests]
svc.spreadsheets().values().batchUpdate(spreadsheetId=SID, body={
    'valueInputOption': 'USER_ENTERED', 'data': data
}).execute()

# Esperar y leer resultados
import time; time.sleep(2)
res = svc.spreadsheets().values().get(spreadsheetId=SID, range='Productos!N1:N4', valueRenderOption='UNFORMATTED_VALUE').execute()
for i, row in enumerate(res.get('values', [])):
    print(f'N{i+1}: {row[0] if row else "(vacio)"}')
print('Esperado: 1 (hay un solo pedido con PPJyQ flag D)')

# Cleanup
svc.spreadsheets().values().batchUpdate(spreadsheetId=SID, body={
    'valueInputOption': 'USER_ENTERED',
    'data': [{'range': c, 'values': [['']]} for c, _ in tests]
}).execute()
