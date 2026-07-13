"""Test usando CHAR(34) en lugar de escape de comillas."""
import google.auth
from googleapiclient.discovery import build
creds, _ = google.auth.default(scopes=['https://www.googleapis.com/auth/spreadsheets'])
svc = build('sheets', 'v4', credentials=creds)
SID = '1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY'

tests = [
    ('Productos!N1', '=REGEXMATCH("hola","ho")'),  # baseline simple
    ('Productos!N2', '=SUMPRODUCT(--(IFERROR(REGEXMATCH(Home!$BB$2:$BB$10000,"PPJyQ"),FALSE)))'),
    ('Productos!N3', '=SUMPRODUCT(--(IFERROR(REGEXMATCH(Home!$BB$2:$BB$10000,CHAR(34)&"PPJyQ"&CHAR(34)&":"&CHAR(34)&"D"&CHAR(34)),FALSE)))'),
    ('Productos!N4', '=Home!$BB$258'),  # Ver el JSON de Tadeo
]
data = [{'range': c, 'values': [[f]]} for c, f in tests]
svc.spreadsheets().values().batchUpdate(spreadsheetId=SID, body={
    'valueInputOption': 'USER_ENTERED', 'data': data
}).execute()
import time; time.sleep(2)
res = svc.spreadsheets().values().get(spreadsheetId=SID, range='Productos!N1:N4', valueRenderOption='UNFORMATTED_VALUE').execute()
for i, row in enumerate(res.get('values', [])):
    print(f'N{i+1}: {row[0] if row else "(vacio)"}')
# cleanup
svc.spreadsheets().values().batchUpdate(spreadsheetId=SID, body={
    'valueInputOption': 'USER_ENTERED',
    'data': [{'range': c, 'values': [['']]} for c, _ in tests]
}).execute()
