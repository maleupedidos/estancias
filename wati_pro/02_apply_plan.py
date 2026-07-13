"""WATI BBDD — Apply Plan.

Lee plan.json y aplica vía POST /api/v1/updateContactAttributes/{phone}
Header: Content-Type: application/json-patch+json
Body: {"customParams":[{"name":"x","value":"y"}, ...]}

Maneja rate limits (429) con backoff exponencial.
Guarda errores en apply_errors.json y resumen en apply_log.txt.
"""
import sys, io, json, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import requests

WATI_TOKEN = "wati_6cac1b8c-07cc-4946-b954-5f52df8ba948.iRUrSg_H28yY_zWU3jyMYFu96ErdgwhsnhNA-1_yHN5simg3-rUejn_ROEAGRhIOp2ulVLp4t-7g5VCyD2mMwXqqWGYn0_SahlRTLVoPczz3xwIH8bXV5NkyJob-dPKn"
WATI_URL = "https://live-mt-server.wati.io/1034656"
HEADERS = {"Authorization": f"Bearer {WATI_TOKEN}", "Content-Type": "application/json-patch+json"}

PLAN = sys.argv[1] if len(sys.argv) > 1 else 'C:/Users/tadeu/estancias/wati_pro/plan.json'
ERRORS_OUT = PLAN.replace('.json','_errors.json')
LOG_OUT = PLAN.replace('.json','_log.txt')

plan = json.load(open(PLAN, encoding='utf-8'))
print(f'Aplicando {len(plan)} cambios...')

ok, fail = [], []
for i, item in enumerate(plan, 1):
    phone = item['phone']
    body = {'customParams': item['set']}
    url = f'{WATI_URL}/api/v1/updateContactAttributes/{phone}'
    success = False
    last_err = None
    for attempt in range(5):
        try:
            r = requests.post(url, headers=HEADERS, json=body, timeout=30)
            if r.status_code == 429:
                wait = 2 ** attempt
                print(f'  [{i}/{len(plan)}] {phone} 429, esperando {wait}s')
                time.sleep(wait); continue
            if r.status_code == 200:
                rj = r.json() if r.text else {}
                if rj.get('result') is True or rj.get('success') is True or rj == {}:
                    success = True
                    break
                last_err = f'200 pero result={rj}'
            else:
                last_err = f'HTTP {r.status_code}: {r.text[:200]}'
        except Exception as e:
            last_err = str(e)
        time.sleep(1 + attempt)
    if success:
        ok.append({'phone':phone,'name':item['name']})
        if i % 20 == 0:
            print(f'  [{i}/{len(plan)}] OK ({len(ok)} acumulados)')
    else:
        fail.append({'phone':phone,'name':item['name'],'error':last_err,'set':item['set']})
        print(f'  [{i}/{len(plan)}] FAIL {phone}: {last_err}')
    time.sleep(0.3)  # rate limit prevent

with open(ERRORS_OUT,'w',encoding='utf-8') as f:
    json.dump(fail,f,ensure_ascii=False,indent=2)

with open(LOG_OUT,'w',encoding='utf-8') as f:
    f.write(f'Total: {len(plan)}\n')
    f.write(f'OK: {len(ok)}\n')
    f.write(f'FAIL: {len(fail)}\n')

print(f'\n=== Aplicado ===')
print(f'  OK: {len(ok)}')
print(f'  FAIL: {len(fail)}')
print(f'\nErrores: {ERRORS_OUT}')
