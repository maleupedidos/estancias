"""WATI BBDD — Fase 2.

Limpia:
  1. Sin tipo_de_contacto -> Prospecto (default)
  2. tipo "Contacto" residual -> Prospecto
  3. Home sin zona -> 'Estancias del Pilar' (default seguro, mayoría)

Excluye: el dueño de la cuenta (5491155038905) y contactos con números
no estándar (cortos, internos).
"""
import sys, io, json, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from collections import Counter

WATI_SNAP = 'C:/Users/tadeu/estancias/wati_pro/wati_snapshot.json'
PLAN_OUT = 'C:/Users/tadeu/estancias/wati_pro/plan_fase2.json'
REPORT_OUT = 'C:/Users/tadeu/estancias/wati_pro/plan_fase2_report.txt'

OWNER = '5491155038905'
TIPOS_CANONICOS = {'Home','Pilar','Clubes','Red','Prospecto','NO_INTERESADO','Proveedor'}

wati = json.load(open(WATI_SNAP, encoding='utf-8'))
plan = []
stats = Counter()

for c in wati:
    phone = c.get('wAid') or c.get('phone') or ''
    if not phone or phone == OWNER: continue
    # Filtrar números no estándar (menos de 10 dígitos, o internos)
    digits = re.sub(r'\D','', phone)
    if len(digits) < 10:
        stats['skip_telefono_corto'] += 1
        continue

    name = c.get('fullName') or c.get('firstName') or ''
    cps = {cp['name']:cp['value'] for cp in c.get('customParams', [])}
    tipo = cps.get('tipo_de_contacto')
    zona = cps.get('zona')

    sets = []; reasons = []

    # 1) Sin tipo o no canónico residual -> Prospecto
    if not tipo:
        sets.append({'name':'tipo_de_contacto','value':'Prospecto'})
        reasons.append('sin tipo -> Prospecto (default)')
        stats['default_Prospecto'] += 1
        tipo = 'Prospecto'
    elif tipo not in TIPOS_CANONICOS:
        sets.append({'name':'tipo_de_contacto','value':'Prospecto'})
        reasons.append(f'"{tipo}" no canónico residual -> Prospecto')
        stats['residual_Prospecto'] += 1
        tipo = 'Prospecto'

    # 2) Home sin zona -> Estancias del Pilar (default mayoritario)
    if tipo == 'Home' and not zona:
        sets.append({'name':'zona','value':'Estancias del Pilar'})
        reasons.append('Home sin zona -> default Estancias del Pilar')
        stats['home_zona_default'] += 1

    if sets:
        plan.append({
            'phone': phone,
            'name': name,
            'current': {'tipo':cps.get('tipo_de_contacto'),'zona':cps.get('zona')},
            'set': sets,
            'reasons': reasons,
        })

with open(PLAN_OUT,'w',encoding='utf-8') as f:
    json.dump(plan,f,ensure_ascii=False,indent=2)

with open(REPORT_OUT,'w',encoding='utf-8') as f:
    f.write(f'Plan Fase 2: {len(plan)} cambios\n')
    for k,v in stats.most_common(): f.write(f'  {k}: {v}\n')

print(f'Plan Fase 2: {len(plan)} contactos')
for k,v in stats.most_common(): print(f'  {k}: {v}')
