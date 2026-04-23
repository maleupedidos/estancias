"""Fix: continuar con stock + kardex de Gloria (estado/fecha ya grabados en intento anterior)."""
import sys
sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

# Verificar estado actual de Gloria
curr = svc.values().get(spreadsheetId=SID, range="'Pilar'!A10:BD10",
                       valueRenderOption="FORMATTED_VALUE").execute().get("values", [[]])[0]
print(f"Gloria: EstEntrega={curr[10]} EstPago={curr[12]} FechaEnt={curr[52] if len(curr)>52 else ''}")

# Descontar stock: Gloria pidió PPM(1), PPJyQ(1), PPCyQ(1), PMa(1), PCC(1)
prod_rows = svc.values().get(spreadsheetId=SID, range="'Productos'!A:J",
                              valueRenderOption="UNFORMATTED_VALUE").execute().get("values", [])
abbrs_a_descontar = {"PPM": 1, "PPJyQ": 1, "PPCyQ": 1, "PMa": 1, "PCC": 1}
updates = []
kardex_rows = []
now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
for i, r in enumerate(prod_rows[1:], start=2):
    if len(r) < 6: continue
    ab = str(r[2]).strip()
    if ab in abbrs_a_descontar:
        qty = abbrs_a_descontar[ab]
        try: stock_ant = int(r[5]) if r[5] != "" else 0
        except: stock_ant = 0
        stock_new = max(0, stock_ant - qty)
        updates.append({"range": f"'Productos'!F{i}", "values": [[stock_new]]})
        kardex_rows.append([now_str, ab, "-SAL", qty, stock_ant, stock_new, "Pilar", "P-1-Gloria-fix-manual"])
        print(f"  {ab}: {stock_ant} -> {stock_new}")

if updates:
    svc.values().batchUpdate(spreadsheetId=SID, body={"valueInputOption":"USER_ENTERED","data":updates}).execute()
    svc.values().append(spreadsheetId=SID, range="'Kardex'!A:H",
                        valueInputOption="USER_ENTERED",
                        insertDataOption="INSERT_ROWS",
                        body={"values": kardex_rows}).execute()
    print(f"Stock y Kardex OK ({len(kardex_rows)} movs)")
else:
    print("Nada para actualizar")
