"""Escanear hojas operativas y contar pedidos con Día de Entrega Elegido en semana 17
(lunes 20/04/2026 a domingo 26/04/2026). Replica la lógica del panel:
- Excluye Cancelado, sin cliente, total 0
- Excluye Red (no cuenta en métrica semanal)
- Usa dee (Día de Entrega Elegido, ISO yyyy-MM-dd)
"""
from datetime import date, datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SPREADSHEET_ID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()

LUN = date(2026, 4, 20)
DOM = date(2026, 4, 26)

def read(name):
    r = svc.values().get(spreadsheetId=SPREADSHEET_ID, range=f"'{name}'!A:BZ", valueRenderOption="FORMATTED_VALUE").execute()
    return r.get("values", [])

def parse_dee(v):
    if not v:
        return None
    s = str(v).strip()
    # ISO yyyy-MM-dd
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except Exception:
            pass
    return None

def scan(hoja):
    data = read(hoja)
    if not data:
        print(f"[{hoja}] vacía")
        return []
    header = [h.strip() for h in data[0]]
    # Buscar columna "Día de Entrega Elegido" (case-insensitive, con o sin acentos)
    def norm(s):
        return s.lower().replace("í","i").replace("á","a").replace("é","e").replace("ó","o").replace("ú","u")
    dee_idx = None
    for i, h in enumerate(header):
        nh = norm(h)
        if "dia de entrega elegido" in nh or "fecha entrega elegida" in nh or nh == "dee":
            dee_idx = i
            break
    # Buscar Total, Estado Entrega, Cliente, N° Pedido, Fecha
    idx = {}
    for i, h in enumerate(header):
        nh = norm(h)
        if nh == "cliente": idx["cli"] = i
        if "n pedido" in nh or "n\u00b0 pedido" in nh or nh.startswith("n pedido") or "nº pedido" in nh or "n° pedido" in nh: idx.setdefault("ped", i)
        if nh.startswith("total"): idx.setdefault("tot", i)
        if "estado de entrega" in nh: idx.setdefault("est", i)
        if nh == "fecha": idx.setdefault("fecha", i)
    print(f"[{hoja}] header dee_idx={dee_idx} ({header[dee_idx] if dee_idx is not None else 'NO'}), claves={idx}")

    rows = []
    for r in data[1:]:
        def g(k):
            i = idx.get(k)
            if i is None or i >= len(r): return ""
            return r[i]
        cli = str(g("cli")).strip()
        est = str(g("est")).strip()
        totraw = str(g("tot")).replace("$","").replace(".","").replace(",",".").strip()
        try:
            tot = float(totraw) if totraw else 0
        except:
            tot = 0
        if not cli or tot <= 0 or est.lower() == "cancelado":
            continue
        deeval = r[dee_idx] if (dee_idx is not None and dee_idx < len(r)) else ""
        d = parse_dee(deeval)
        if d and LUN <= d <= DOM:
            rows.append({
                "ped": str(g("ped")),
                "cli": cli,
                "fecha": str(g("fecha")),
                "dee": str(deeval),
                "tot": tot,
                "est": est,
            })
    return rows

total = 0
for hoja in ["Home", "Pilar", "Clubes", "B2B", "Catering"]:
    try:
        rows = scan(hoja)
    except Exception as e:
        print(f"[{hoja}] ERROR {e}")
        continue
    print(f"\n=== {hoja}: {len(rows)} pedidos con entrega en semana 17 ===")
    for r in rows:
        print(f"  {r['ped']:<10} {r['cli']:<30} fecha={r['fecha']} dee={r['dee']} ${r['tot']:>10,.0f} est={r['est']}")
    total += len(rows)

print(f"\n>>> TOTAL (Home+Pilar+Clubes+B2B+Catering, sin Red): {total} ventas")
