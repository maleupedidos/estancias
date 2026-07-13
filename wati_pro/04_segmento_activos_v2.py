"""Segmento Activos Home 2026 — versión robusta.

Estrategia:
1. Lee Sheet Home. Filtra Año=2026 y Estado != Cancelado.
2. Para cada cliente, normaliza el teléfono con heurística mejorada.
3. Busca el contacto en WATI por teléfono normalizado.
4. Si no matchea por teléfono, busca por NOMBRE (fullName fuzzy match en snapshot WATI).
5. Si matchea por nombre, usa ese wAid para taggear.
6. Aplica activo_home_2026=Si + pedidos_home_2026=N.
"""
import sys, io, json, re, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import requests
from collections import Counter
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

WATI_TOKEN = "wati_6cac1b8c-07cc-4946-b954-5f52df8ba948.iRUrSg_H28yY_zWU3jyMYFu96ErdgwhsnhNA-1_yHN5simg3-rUejn_ROEAGRhIOp2ulVLp4t-7g5VCyD2mMwXqqWGYn0_SahlRTLVoPczz3xwIH8bXV5NkyJob-dPKn"
WATI_URL = "https://live-mt-server.wati.io/1034656"
HG = {"Authorization": f"Bearer {WATI_TOKEN}"}
HP = {"Authorization": f"Bearer {WATI_TOKEN}", "Content-Type": "application/json-patch+json"}

SA = r"C:\Users\tadeu\maleu-service-account.json"
SID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"

def normalize_phone(raw):
    """Devuelve lista de candidatos plausibles para un teléfono argentino crudo."""
    s = re.sub(r'\D', '', str(raw or ''))
    if not s: return []
    cands = set()
    # Caso ya bien formado
    if s.startswith('549') and len(s) == 13:
        cands.add(s); return list(cands)
    # 54 + 12 dígitos (sin 9 móvil): agregar 9 después del 54
    if s.startswith('54') and len(s) == 12:
        cands.add('549' + s[2:])
    # 0 al principio (formato fijo AR con 0): sacar 0, agregar 549
    if s.startswith('0'):
        cands.add('549' + s[1:])
    # 15 al inicio (formato móvil viejo): sacar 15, agregar 5491
    if s.startswith('15'):
        cands.add('5491' + s[2:])
    # 9 al inicio + 11 dígitos: agregar 54
    if s.startswith('9') and len(s) == 11:
        cands.add('54' + s)
    # 9 al inicio + 12 dígitos: agregar 54 (algunos clientes ponen 9 y código área completo)
    if s.startswith('9') and len(s) == 12:
        cands.add('54' + s)
    # 11 al inicio (área CABA/GBA) + 10 dígitos: agregar 549
    if s.startswith('11') and len(s) == 10:
        cands.add('549' + s)
    # 11 al inicio + 11 dígitos (cliente puso un dígito de más, típico typo): probar quitando uno
    if s.startswith('11') and len(s) == 11:
        # candidato 1: 549 + s (queda 14 dig, raro pero por las dudas)
        cands.add('549' + s)
        # candidato 2: probar quitando dígitos repetidos en el medio (typos comunes)
        for i in range(2, len(s)-1):
            if s[i] == s[i+1]:
                cands.add('549' + s[:i] + s[i+1:])
    # 10 dígitos puros: 549 + s
    if len(s) == 10:
        cands.add('549' + s)
    # 8 dígitos: 54911 + s
    if len(s) == 8:
        cands.add('54911' + s)
    # 12 dígitos sin nada conocido: probar agregar 5
    if len(s) == 12 and s.startswith('491'):
        cands.add('5' + s)
    # Último recurso: el crudo
    cands.add(s)
    return list(cands)

def normalize_name(n):
    """Normaliza nombre para fuzzy match: lowercase, sin acentos, una palabra."""
    import unicodedata
    n = str(n or '').strip().lower()
    n = unicodedata.normalize('NFKD', n).encode('ascii', 'ignore').decode('ascii')
    return ' '.join(n.split())

# === BAJAR SHEET HOME ===
print('Leyendo Sheet Home...')
creds = Credentials.from_service_account_file(SA, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
svc = build('sheets', 'v4', credentials=creds).spreadsheets().values()
home = svc.get(spreadsheetId=SID, range='Home').execute().get('values', [])
COL_ANIO=6; COL_CLI=7; COL_ESTADO=10; COL_TEL=46
clientes_2026 = {}  # cliente_normalizado -> {nombres, telefonos_raw, pedidos}
clientes_por_tel = {}  # tel_raw -> info
total_filas = 0
for r in home[1:]:
    total_filas += 1
    if len(r) <= COL_TEL: continue
    if str(r[COL_ANIO]).strip() != '2026': continue
    if str(r[COL_ESTADO]).strip().lower().startswith('cancel'): continue
    tel_raw = r[COL_TEL]
    cli = r[COL_CLI] if COL_CLI < len(r) else ''
    if not tel_raw: continue
    key = tel_raw  # agrupo por teléfono raw
    if key not in clientes_por_tel:
        clientes_por_tel[key] = {'cliente': cli, 'pedidos': 0, 'tel_raw': tel_raw}
    clientes_por_tel[key]['pedidos'] += 1
    # Si hay variaciones de nombre del mismo tel, prefiero el más largo
    if len(cli) > len(clientes_por_tel[key]['cliente']):
        clientes_por_tel[key]['cliente'] = cli

print(f'Pedidos Home 2026 (no cancelados): {sum(c["pedidos"] for c in clientes_por_tel.values())}')
print(f'Clientes únicos por tel raw: {len(clientes_por_tel)}')

# === BAJAR WATI SNAPSHOT ===
print('\nBajando snapshot WATI...')
all_contacts = []
page = 0
while True:
    for a in range(5):
        r = requests.get(f'{WATI_URL}/api/v1/getContacts?pageSize=100&pageNumber={page}', headers=HG, timeout=30)
        if r.status_code == 429: time.sleep(2**a); continue
        r.raise_for_status(); break
    d = r.json()
    items = d.get('contact_list', d.get('contacts', []))
    if not items: break
    all_contacts.extend(items)
    if len(items) < 100: break
    page += 1
    time.sleep(0.4)
# Dedup
seen = {}
for c in all_contacts:
    p = c.get('wAid') or c.get('phone', '')
    if p and p not in seen: seen[p] = c
wati = list(seen.values())
print(f'WATI únicos: {len(wati)}')

# Index WATI por phone y por nombre normalizado
wati_by_phone = {c.get('wAid') or c.get('phone'): c for c in wati}
wati_by_name = {}  # nombre_norm -> [contactos]
for c in wati:
    fn = c.get('fullName') or c.get('firstName') or ''
    nn = normalize_name(fn)
    if nn:
        wati_by_name.setdefault(nn, []).append(c)

# === MATCH ===
matched = []  # {phone_wati, cliente, pedidos, via}
unmatched = []
stats = Counter()

for info in clientes_por_tel.values():
    cli = info['cliente']
    pedidos = info['pedidos']
    tel_raw = info['tel_raw']
    cands = normalize_phone(tel_raw)
    found = None
    via = None
    for c in cands:
        if c in wati_by_phone:
            found = c; via = 'tel_directo' if c == cands[0] else 'tel_normalizado'; break
    if not found:
        # Buscar por nombre exacto
        nn = normalize_name(cli)
        if nn and nn in wati_by_name and len(wati_by_name[nn]) == 1:
            wc = wati_by_name[nn][0]
            found = wc.get('wAid') or wc.get('phone')
            via = 'nombre_exacto'
        elif nn:
            # Buscar por contains parcial (apellido o nombre completo)
            partes = nn.split()
            cand_contactos = []
            for nname, wcs in wati_by_name.items():
                if all(p in nname for p in partes):  # todas las palabras del cliente están en el nombre WATI
                    cand_contactos.extend(wcs)
            cand_contactos = list({c.get('wAid'): c for c in cand_contactos}.values())
            if len(cand_contactos) == 1:
                found = cand_contactos[0].get('wAid') or cand_contactos[0].get('phone')
                via = 'nombre_parcial_unico'
            elif len(cand_contactos) > 1:
                via = f'nombre_ambiguo({len(cand_contactos)})'
    if found:
        matched.append({'phone': found, 'cliente': cli, 'pedidos': pedidos, 'tel_raw': tel_raw, 'via': via})
        stats[via] += 1
    else:
        unmatched.append({'cliente': cli, 'pedidos': pedidos, 'tel_raw': tel_raw, 'via': via or 'sin_match'})
        stats[via or 'sin_match'] += 1

print(f'\n=== Match results ===')
print(f'Total clientes únicos: {len(clientes_por_tel)}')
print(f'Matched: {len(matched)}')
print(f'Unmatched: {len(unmatched)}')
print('Por método:')
for k, v in stats.most_common(): print(f'  {k}: {v}')
print('\nUnmatched detail:')
for u in unmatched:
    print(f'  - {u["cliente"]:35} tel={u["tel_raw"]:20} pedidos={u["pedidos"]} via={u["via"]}')

# === BUILD PLAN ===
plan = []
for m in matched:
    plan.append({'phone': m['phone'], 'name': m['cliente'], 'current': {},
                 'set': [{'name': 'activo_home_2026', 'value': 'Si'},
                         {'name': 'pedidos_home_2026', 'value': str(m['pedidos'])}],
                 'reasons': [f'compró {m["pedidos"]} vez/veces en Home 2026 — match via {m["via"]}']})
with open('C:/Users/tadeu/estancias/wati_pro/plan_segmento_activos_v2.json', 'w', encoding='utf-8') as f:
    json.dump(plan, f, ensure_ascii=False, indent=2)
with open('C:/Users/tadeu/estancias/wati_pro/segmento_activos_unmatched.json', 'w', encoding='utf-8') as f:
    json.dump(unmatched, f, ensure_ascii=False, indent=2)
print(f'\nPlan: {len(plan)} contactos a actualizar')
print(f'Unmatched: {len(unmatched)} (en wati_pro/segmento_activos_unmatched.json)')
