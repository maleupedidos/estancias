"""Crea los 7 contactos faltantes en WATI con la categorización que dio Tadeo
y les aplica los tags activo_home_2026 + tipo_de_contacto + zona donde corresponda.
"""
import sys, io, json, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import requests

WATI_TOKEN = "wati_6cac1b8c-07cc-4946-b954-5f52df8ba948.iRUrSg_H28yY_zWU3jyMYFu96ErdgwhsnhNA-1_yHN5simg3-rUejn_ROEAGRhIOp2ulVLp4t-7g5VCyD2mMwXqqWGYn0_SahlRTLVoPczz3xwIH8bXV5NkyJob-dPKn"
WATI_URL = "https://live-mt-server.wati.io/1034656"

# Decisiones según contexto que dio Tadeo (29/04/2026):
contactos = [
    {'phone':'5492901408422','name':'Belen Garcia Basualdo','tipo':'Prospecto','zona':None,
     'nota':'prima de Tadeo, sin casa propia en Estancias. Compra cuando viene de visita'},
    {'phone':'5491139379131','name':'Lautaro Mari','tipo':'Prospecto','zona':None,
     'nota':'amigo, compró pasando un finde con un amigo en Estancias'},
    {'phone':'5491126857061','name':'Dionisio Quesada','tipo':'Home','zona':'Estancias del Pilar',
     'nota':'casa de fin de semana en Estancias'},
    {'phone':'5491144017589','name':'Tomas Lanfranchi','tipo':'Prospecto','zona':None,
     'nota':'amigo, compró pasando un finde con un amigo en Estancias'},
    {'phone':'5491140269813','name':'Sofia Goldaracena','tipo':'Home','zona':'Estancias del Pilar',
     'nota':'vive en Estancias'},
    {'phone':'5493462383262','name':'Rufino de Bary','tipo':'Home','zona':'Estancias del Pilar',
     'nota':'vive en Estancias'},
    {'phone':'5491158945401','name':'Pauline de Ocampo','tipo':'Prospecto','zona':None,
     'nota':'sin info — pedido sin contexto claro'},
]

print(f'Procesando {len(contactos)} contactos...\n')
ok, fail = [], []

for c in contactos:
    phone = c['phone']
    print(f'[{c["name"]}] phone={phone}')
    # Paso 1: crear contacto (addContact)
    add_url = f'{WATI_URL}/api/v1/addContact/{phone}'
    add_body = {'name': c['name'], 'firstName': c['name'].split()[0], 'lastName': ' '.join(c['name'].split()[1:])}
    add_h = {'Authorization': f'Bearer {WATI_TOKEN}', 'Content-Type': 'application/json'}
    try:
        r = requests.post(add_url, headers=add_h, json=add_body, timeout=30)
        print(f'  addContact: {r.status_code} → {r.text[:150]}')
    except Exception as e:
        print(f'  addContact ERROR: {e}')
        fail.append({'name':c['name'],'step':'addContact','error':str(e)})
        continue
    time.sleep(0.5)

    # Paso 2: setear customParams
    custom = [
        {'name':'tipo_de_contacto','value':c['tipo']},
        {'name':'activo_home_2026','value':'Si'},
        {'name':'pedidos_home_2026','value':'1'},
    ]
    if c['zona']:
        custom.append({'name':'zona','value':c['zona']})
    upd_url = f'{WATI_URL}/api/v1/updateContactAttributes/{phone}'
    upd_h = {'Authorization': f'Bearer {WATI_TOKEN}', 'Content-Type': 'application/json-patch+json'}
    try:
        r = requests.post(upd_url, headers=upd_h, json={'customParams':custom}, timeout=30)
        rj = r.json() if r.text else {}
        if r.status_code == 200 and (rj.get('result') is True or rj == {}):
            print(f'  update: OK')
            ok.append(c['name'])
        else:
            print(f'  update FAIL: {r.status_code} → {r.text[:200]}')
            fail.append({'name':c['name'],'step':'update','code':r.status_code,'body':r.text[:200]})
    except Exception as e:
        print(f'  update ERROR: {e}')
        fail.append({'name':c['name'],'step':'update','error':str(e)})
    time.sleep(0.4)

print(f'\n=== Resultado ===')
print(f'OK: {len(ok)}')
print(f'FAIL: {len(fail)}')
if fail:
    for f in fail: print(f'  - {f}')
