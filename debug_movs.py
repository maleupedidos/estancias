"""Verifica que el nuevo endpoint admin devuelva movimientos[]."""
import urllib.request, json, time

URL = "https://script.google.com/macros/s/AKfycbxmrG5YVSshcYezk8lXFx_uxb7NFGcb9EfTXc7dsIN4rZyj73CET4mk_aKPFPDY2wNi/exec?action=admin&t=" + str(int(time.time()))
req = urllib.request.Request(URL)
resp = urllib.request.urlopen(req, timeout=60)
data = json.loads(resp.read())
movs = data.get("movimientos", None)
print("movimientos field present?", movs is not None)
print("total movimientos:", len(movs) if movs else 0)
print("\nBreakdown por tipo:")
if movs:
    from collections import Counter
    c = Counter([m["tipo"] for m in movs])
    for t,n in c.items(): print(f"  {t}: {n}")
    print("\nPrimeros 15 (orden actual del backend):")
    for m in movs[:15]:
        sig = "+" if m["tipo"]!="gasto" else "-"
        ts = m.get("ts",0)
        print(f"  ts={ts:<14} [{m['tipo']:<7}] {m.get('f','')[:18]:<18} {m.get('con','')[:30]:<30} {sig}${m.get('$',0)}")
