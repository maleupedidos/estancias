"""Plan de etiquetado WATI (dry-run).

Reglas:
  - Para cada cliente del Sheet Home, buscar telefono en WATI.
  - Asignar tipo_de_contacto = 'Home'.
  - Asignar zona segun Sub Barrio del Sheet:
      * 'Estancias del Río'   si Sub Barrio in {'Estancias Del Rio','Estancias del Rio'}
      * 'Los Alcanfores'      si Sub Barrio in {'Alcanfores','Los Alcanfores'}
      * 'Estancias del Pilar' para el resto de sub-barrios (Champagnat Alto/Bajo, Golf,
        La Pionera, El Recuerdo, Argentina 1-4, La Paz, etc.)

Genera JSON con:
  - contactos a actualizar (match ok)
  - contactos Home sin match en WATI (a crear o investigar)
  - contactos WATI ya con tipo=Home pero sin zona asignada o con zona distinta

NO APLICA NADA. Solo reporta."""
import sys, io, json, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import requests
from collections import Counter, defaultdict
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

WATI_URL = "https://live-mt-server.wati.io/1034656"
WATI_TOKEN = "wati_6cac1b8c-07cc-4946-b954-5f52df8ba948.iRUrSg_H28yY_zWU3jyMYFu96ErdgwhsnhNA-1_yHN5simg3-rUejn_ROEAGRhIOp2ulVLp4t-7g5VCyD2mMwXqqWGYn0_SahlRTLVoPczz3xwIH8bXV5NkyJob-dPKn"
HEADERS = {"Authorization": f"Bearer {WATI_TOKEN}", "Content-Type": "application/json"}

SA_KEY = r"C:\Users\tadeu\maleu-service-account.json"
SID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"

ZONAS_VALIDAS = {"Estancias del Pilar", "Estancias del Río", "Los Alcanfores"}

def last10(p):
    s = "".join(c for c in str(p or "") if c.isdigit())
    return s[-10:] if len(s) >= 10 else s

def full_phone(p):
    """Devuelve el numero tal como WATI lo guarda (549...)."""
    s = "".join(c for c in str(p or "") if c.isdigit())
    # Garantizar prefijo 549
    if s.startswith("549"): return s
    if s.startswith("54") and len(s) == 12: return s  # sin 9
    if s.startswith("9") and len(s) == 11: return "54" + s
    if len(s) == 10: return "549" + s
    return s

def clasificar_zona(sub_barrio):
    sb = (sub_barrio or "").strip().lower()
    if not sb or sb == "-": return None
    if "rio" in sb or "río" in sb:
        return "Estancias del Río"
    if "alcanfor" in sb:
        return "Los Alcanfores"
    # Todo lo demás en Home se considera Estancias del Pilar
    return "Estancias del Pilar"

# ---- WATI ----
print("Descargando contactos WATI...")
contacts = []
page = 0
while True:
    r = requests.get(f"{WATI_URL}/api/v1/getContacts", headers=HEADERS,
                     params={"pageSize": 100, "pageNumber": page})
    r.raise_for_status()
    data = r.json()
    batch = data.get("contact_list") or data.get("contacts") or []
    if not batch: break
    contacts.extend(batch)
    page += 1
    if page > 50: break
    time.sleep(0.15)
print(f"  total: {len(contacts)}")

wati_by_last10 = {}
for c in contacts:
    ph = last10(c.get("wAid") or c.get("phone") or "")
    if ph: wati_by_last10[ph] = c

def get_cp(c, key):
    for cp in (c.get("customParams") or []):
        if cp.get("name","").strip().lower() == key.lower():
            return cp.get("value","").strip()
    return ""

# ---- Sheet Home ----
print("Leyendo Sheet Home...")
creds = Credentials.from_service_account_file(SA_KEY, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
svc = build("sheets", "v4", credentials=creds).spreadsheets()
home = svc.values().get(spreadsheetId=SID, range="'Home'!A1:BZ3000").execute().get("values", [])
headers = home[0]
def col(n):
    for i,h in enumerate(headers):
        if h.strip().lower() == n.lower(): return i
    return -1

i_cli = col("Cliente")
i_tel = [i for i,h in enumerate(headers) if "tel" in h.lower()][0]
i_sb  = [i for i,h in enumerate(headers) if "sub barrio" in h.lower()][0]

home_clients = {}  # last10 -> {cliente, sub_barrio, telefono_full}
for row in home[1:]:
    if len(row) <= max(i_cli,i_tel,i_sb):
        # extender
        row = row + [""]*(max(i_cli,i_tel,i_sb)+1-len(row))
    tel_raw = row[i_tel]
    p10 = last10(tel_raw)
    if len(p10) < 8: continue
    cli = (row[i_cli] or "").strip()
    sb = (row[i_sb] or "").strip()
    if p10 not in home_clients:
        home_clients[p10] = {"cliente": cli, "sub_barrio": sb, "tel_full": full_phone(tel_raw)}
    else:
        # si el registro previo no tenia sub_barrio, usar este
        if sb and not home_clients[p10]["sub_barrio"]:
            home_clients[p10]["sub_barrio"] = sb

print(f"  clientes unicos Home: {len(home_clients)}")

# ---- Matching ----
updates = []        # {phone, cliente, sub_barrio, zona, wati_id, tipo_actual, zona_actual}
no_match = []       # clientes Home sin contacto en WATI
sub_barrio_desconocido = []

for p10, info in home_clients.items():
    zona = clasificar_zona(info["sub_barrio"])
    if zona is None:
        sub_barrio_desconocido.append({"tel": info["tel_full"], "cliente": info["cliente"], "sub_barrio": info["sub_barrio"]})
        continue
    c = wati_by_last10.get(p10)
    if not c:
        no_match.append({"tel": info["tel_full"], "cliente": info["cliente"], "sub_barrio": info["sub_barrio"], "zona_objetivo": zona})
        continue
    tipo_actual = get_cp(c, "tipo_de_contacto")
    zona_actual = get_cp(c, "zona")
    updates.append({
        "phone": c.get("wAid") or c.get("phone"),
        "cliente_sheet": info["cliente"],
        "wati_name": c.get("fullName") or c.get("firstName"),
        "sub_barrio": info["sub_barrio"],
        "zona_objetivo": zona,
        "tipo_actual_wati": tipo_actual,
        "zona_actual_wati": zona_actual,
        "cambia_tipo": tipo_actual != "Home",
        "cambia_zona": zona_actual != zona,
    })

# Resumen
print("\n=== PLAN ===")
print(f"Clientes Home con match en WATI y clasificables: {len(updates)}")
cambia_tipo = sum(1 for u in updates if u["cambia_tipo"])
cambia_zona = sum(1 for u in updates if u["cambia_zona"])
sin_cambio  = sum(1 for u in updates if not u["cambia_tipo"] and not u["cambia_zona"])
print(f"  - cambian tipo_de_contacto a Home: {cambia_tipo}")
print(f"  - cambian zona: {cambia_zona}")
print(f"  - ya correctos (sin cambio): {sin_cambio}")

print(f"\nDistribucion por zona objetivo:")
for z, n in Counter(u["zona_objetivo"] for u in updates).items():
    print(f"  {z}: {n}")

print(f"\nClientes Home SIN match en WATI: {len(no_match)}")
print(f"Clientes con Sub Barrio desconocido o vacio: {len(sub_barrio_desconocido)}")
if sub_barrio_desconocido:
    print("  (estos quedan fuera del plan):")
    for x in sub_barrio_desconocido[:10]:
        print(f"    {x['tel']} - {x['cliente']} - sub='{x['sub_barrio']}'")

# Guardar plan
with open(r"C:\Users\tadeu\estancias\wati_plan.json","w",encoding="utf-8") as f:
    json.dump({
        "updates": updates,
        "no_match": no_match,
        "sub_barrio_desconocido": sub_barrio_desconocido,
    }, f, ensure_ascii=False, indent=2)
print("\nPlan guardado en wati_plan.json")

# Mostrar 10 ejemplos de cambios
print("\nEjemplos de cambios (primeros 10):")
for u in updates[:10]:
    print(f"  {u['phone']} {u['cliente_sheet']:<25} sb='{u['sub_barrio']}' -> tipo={u['tipo_actual_wati']!r}->Home | zona={u['zona_actual_wati']!r}->{u['zona_objetivo']!r}")
