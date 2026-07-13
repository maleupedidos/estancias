"""Inspecciona la fórmula actual de E2 y G2 en Productos para ver el escape exacto."""
import google.auth
from googleapiclient.discovery import build
creds, _ = google.auth.default(scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
svc = build('sheets', 'v4', credentials=creds)
SID = '1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY'
r = svc.spreadsheets().values().get(
    spreadsheetId=SID, range='Productos!E2:G2',
    valueRenderOption='FORMULA'
).execute()
vals = r.get('values', [[]])[0]
for i, v in enumerate(vals):
    print(f'Col {chr(69+i)}:')
    print(v)
    print()
