"""Aplica las etiquetas tipo_de_contacto=Home + zona=<X> en WATI.

Previamente ejecutar wati_plan.py para generar wati_plan.json.
Zona usa la ortografia vigente en WATI (sin tilde en 'Rio') para NO crear duplicados.
"""
import sys, io, json, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import requests

WATI_URL = "https://live-mt-server.wati.io/1034656"
WATI_TOKEN = "wati_6cac1b8c-07cc-4946-b954-5f52df8ba948.iRUrSg_H28yY_zWU3jyMYFu96ErdgwhsnhNA-1_yHN5simg3-rUejn_ROEAGRhIOp2ulVLp4t-7g5VCyD2mMwXqqWGYn0_SahlRTLVoPczz3xwIH8bXV5NkyJob-dPKn"
HEADERS = {"Authorization": f"Bearer {WATI_TOKEN}", "Content-Type": "application/json-patch+json"}

ZONA_MAP = {
    "Estancias del Pilar": "Estancias del Pilar",
    "Estancias del Río": "Estancias del Rio",   # sin tilde para matchear lo existente
    "Los Alcanfores": "Los Alcanfores",
}

with open(r"C:\Users\tadeu\estancias\wati_plan.json", "r", encoding="utf-8") as f:
    plan = json.load(f)

updates = plan["updates"]
print(f"Aplicando {len(updates)} actualizaciones...")

ok, fail = 0, 0
errores = []

for i, u in enumerate(updates, 1):
    phone = u["phone"]
    zona = ZONA_MAP.get(u["zona_objetivo"], u["zona_objetivo"])
    body = {
        "customParams": [
            {"name": "tipo_de_contacto", "value": "Home"},
            {"name": "zona", "value": zona},
        ]
    }
    try:
        r = requests.post(f"{WATI_URL}/api/v1/updateContactAttributes/{phone}",
                          headers=HEADERS, json=body, timeout=15)
        if r.status_code == 200:
            ok += 1
        else:
            fail += 1
            errores.append({"phone": phone, "status": r.status_code, "body": r.text[:200]})
    except Exception as e:
        fail += 1
        errores.append({"phone": phone, "error": str(e)})
    if i % 20 == 0:
        print(f"  {i}/{len(updates)}  ok={ok} fail={fail}")
    time.sleep(0.1)

print(f"\nTerminado. OK={ok}  FAIL={fail}")
if errores:
    print("\nPrimeros errores:")
    for e in errores[:10]:
        print(" ", e)
    with open(r"C:\Users\tadeu\estancias\wati_apply_errors.json","w",encoding="utf-8") as f:
        json.dump(errores, f, ensure_ascii=False, indent=2)
