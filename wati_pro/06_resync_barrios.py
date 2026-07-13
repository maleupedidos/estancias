"""Resincroniza barrio_privado en WATI tomando el ÚLTIMO Sub Barrio del Sheet Home
para cada cliente. Excluye los 3 visitantes que Tadeo marcó como zona=Capital
(Belen Garcia, Lautaro Mari, Tomas Lanfranchi) y a Pauline de Ocampo (sin info).
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

# Excluir: visitantes con zona=Capital + Pauline (sin info)
EXCLUIR = {'5492901408422','5491139379131','5491144017589','5491158945401'}

SA = r"C:\Users\tadeu\maleu-service-account.json"
SID = "1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY"

def normalize_phone(raw):
    s=re.sub(r'\D','',str(raw or ''))
    if not s: return []
    cs=set()
    if s.startswith('549') and len(s)==13: cs.add(s); return list(cs)
    if s.startswith('54') and len(s)==12: cs.add('549'+s[2:])
    if s.startswith('0'): cs.add('549'+s[1:])
    if s.startswith('15'): cs.add('5491'+s[2:])
    if s.startswith('9') and len(s)==11: cs.add('54'+s)
    if s.startswith('11') and len(s)==10: cs.add('549'+s)
    if s.startswith('11') and len(s)==11:
        for i in range(2,len(s)-1):
            if s[i]==s[i+1]: cs.add('549'+s[:i]+s[i+1:])
    if len(s)==10: cs.add('549'+s)
    if len(s)==8: cs.add('54911'+s)
    cs.add(s)
    return list(cs)

def barrio_canon(b):
    if not b: return ''
    M={'estancias del rio':'Estancias del Rio','estancias del río':'Estancias del Rio',
       'los alcanfores':'Los Alcanfores','alcanfores':'Los Alcanfores',
       'champagnat alto':'Champagnat Alto','champagnat bajo':'Champagnat Bajo',
       'golf':'Golf','la pionera':'La Pionera','el recuerdo':'El Recuerdo',
       'la paz':'La Paz','argentina 1':'Argentina 1','argentina 2':'Argentina 2',
       'argentina 3':'Argentina 3','argentina 4':'Argentina 4','pilara':'Pilara'}
    return M.get(b.strip().lower(), b.strip())

def zona_from_sub(sub):
    sb=str(sub or '').strip().lower()
    if not sb: return None
    if 'rio' in sb or 'río' in sb: return 'Estancias del Rio'
    if 'alcanfor' in sb: return 'Los Alcanfores'
    return 'Estancias del Pilar'

# Leer Sheet Home
print('Leyendo Sheet Home...')
creds = Credentials.from_service_account_file(SA, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
svc = build('sheets', 'v4', credentials=creds).spreadsheets().values()
home = svc.get(spreadsheetId=SID, range='Home').execute().get('values', [])
COL_TEL=46; COL_CLI=7; COL_BARRIO=43; COL_SUB=44

# Recorrer en orden NORMAL pero ir actualizando — el último gana (más reciente al final del Sheet)
sheet_by_tel = {}
for r in home[1:]:
    if len(r) <= COL_TEL: continue
    tel = r[COL_TEL]
    sub = str(r[COL_SUB] if COL_SUB < len(r) else '').strip()
    barrio = str(r[COL_BARRIO] if COL_BARRIO < len(r) else '').strip()
    cli = r[COL_CLI] if COL_CLI < len(r) else ''
    if not tel: continue
    # Solo guardo si tiene sub no vacío (siempre piso con el más reciente con sub)
    if sub:
        sheet_by_tel[tel] = {'cliente': cli, 'sub': sub, 'barrio': barrio}
    elif tel not in sheet_by_tel:
        sheet_by_tel[tel] = {'cliente': cli, 'sub': '', 'barrio': barrio}

# WATI
print('Bajando WATI snapshot...')
all_=[]; page=0
while True:
    for a in range(5):
        r=requests.get(f'{WATI_URL}/api/v1/getContacts?pageSize=100&pageNumber={page}', headers=HG, timeout=30)
        if r.status_code==429: time.sleep(2**a); continue
        r.raise_for_status(); break
    d=r.json(); items=d.get('contact_list',d.get('contacts',[]))
    if not items: break
    all_.extend(items)
    if len(items)<100: break
    page+=1; time.sleep(0.4)
wati_by_phone={}
for c in all_:
    p=c.get('wAid') or c.get('phone','')
    if p: wati_by_phone[p]=c

# Match + plan
plan = []
for tel, info in sheet_by_tel.items():
    if not info['sub']: continue  # solo proceso los que tienen sub en Sheet
    expected_barrio = barrio_canon(info['sub'])
    expected_zona = zona_from_sub(info['sub'])

    cands = normalize_phone(tel)
    found = None; phone_wati = None
    for c in cands:
        if c in wati_by_phone:
            found = wati_by_phone[c]; phone_wati = c; break
    if not found: continue
    if phone_wati in EXCLUIR:
        continue

    cps = {cp['name']: cp['value'] for cp in found.get('customParams', [])}
    sets = []
    if cps.get('barrio_privado','') != expected_barrio:
        sets.append({'name':'barrio_privado','value':expected_barrio})
    # zona solo se actualiza si tipo=Home
    if cps.get('tipo_de_contacto') == 'Home' and cps.get('zona','') != expected_zona:
        sets.append({'name':'zona','value':expected_zona})
    if sets:
        plan.append({'phone': phone_wati, 'name': info['cliente'],
                     'current': {'barrio': cps.get('barrio_privado'), 'zona': cps.get('zona'), 'tipo': cps.get('tipo_de_contacto')},
                     'expected': {'barrio': expected_barrio, 'zona': expected_zona},
                     'set': sets})

print(f'\nDesincronizados a actualizar: {len(plan)}')
for p in plan:
    print(f'  {p["name"][:30]:30} {p["current"]["barrio"] or "VACIO":18} -> {p["expected"]["barrio"]:18}  (tipo={p["current"]["tipo"]})')

# Aplicar
ok=fail=0
print('\nAplicando...')
for p in plan:
    body={'customParams': p['set']}
    r=requests.post(f'{WATI_URL}/api/v1/updateContactAttributes/{p["phone"]}', headers=HP, json=body, timeout=30)
    rj=r.json() if r.text else {}
    if r.status_code==200 and (rj.get('result') is True or rj=={}):
        ok+=1
    else:
        fail+=1
        print(f'  FAIL {p["name"]}: {r.status_code} {r.text[:100]}')
    time.sleep(0.3)
print(f'\nOK: {ok} | FAIL: {fail}')
