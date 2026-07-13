"""WATI BBDD — Build Plan (dry-run profesional).

Genera plan.json con todas las normalizaciones a aplicar:
  1. Tipos no canónicos -> mapeo a canónicos
  2. Grafías inconsistentes en barrio_privado -> grafía canónica
  3. Sin tipo + telefono en Sheets -> tipo correspondiente (Home/Pilar/Clubes/Red)
  4. Home sin zona -> zona derivada de Sub Barrio en Sheet Home
  5. Proveedores en Sheet Proveedores -> tipo=Proveedor

NO aplica nada. Solo genera plan + reporte.
"""
import sys, io, json, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from collections import Counter, defaultdict

WATI_SNAP = 'C:/Users/tadeu/estancias/wati_pro/wati_snapshot.json'
SHEETS_SNAP = 'C:/Users/tadeu/estancias/wati_pro/sheets_snapshot.json'
PLAN_OUT = 'C:/Users/tadeu/estancias/wati_pro/plan.json'
REPORT_OUT = 'C:/Users/tadeu/estancias/wati_pro/plan_report.txt'

# Valores canónicos
TIPOS_CANONICOS = {'Home','Pilar','Clubes','Red','Prospecto','NO_INTERESADO','Proveedor'}
ZONAS_CANONICAS = {'Estancias del Pilar','Estancias del Rio','Los Alcanfores'}

# Mapeos de tipos no canónicos -> canónico
MAP_TIPO_NO_CANON = {
    'Cliente Directo':'Home',          # 4 contactos: clientes que compraron directo
    'Capital Federal':'Pilar',         # 1 contacto: CABA pide por Pilar (decisión 21/04)
    'Prospecto Activo':'Prospecto',    # 1 contacto
    'Ocasional':'Prospecto',           # 1 contacto
    # 'Contacto' (2) lo dejamos para clasificar por Sheets, no es claro qué es
}

# Grafía canónica de barrio_privado (case-insensitive lookup)
MAP_BARRIO_CANON = {
    'estancias del rio':'Estancias del Rio',
    'estancias del río':'Estancias del Rio',
    'los alcanfores':'Los Alcanfores',
    'alcanfores':'Los Alcanfores',
    'champagnat alto':'Champagnat Alto',
    'champagnat bajo':'Champagnat Bajo',
    'golf':'Golf',
    'la pionera':'La Pionera',
    'el recuerdo':'El Recuerdo',
    'la paz':'La Paz',
    'argentina 1':'Argentina 1',
    'argentina 2':'Argentina 2',
    'argentina 3':'Argentina 3',
    'argentina 4':'Argentina 4',
    'pilara':'Pilara',
}

def last10(p):
    s = re.sub(r'\D','', str(p or ''))
    return s[-10:] if len(s)>=10 else s

def clasificar_zona(sub_barrio):
    sb = (sub_barrio or '').strip().lower()
    if 'rio' in sb or 'río' in sb: return 'Estancias del Rio'
    if 'alcanfor' in sb: return 'Los Alcanfores'
    if not sb: return None
    return 'Estancias del Pilar'

def normalizar_barrio_canon(b):
    if not b: return None
    key = b.strip().lower()
    return MAP_BARRIO_CANON.get(key, b.strip())

# ============== LOAD ==============
wati = json.load(open(WATI_SNAP, encoding='utf-8'))
sheets = json.load(open(SHEETS_SNAP, encoding='utf-8'))

# WATI: index by last10
wati_by_l10 = {}
for c in wati:
    phone = c.get('wAid') or c.get('phone') or ''
    l10 = last10(phone)
    if l10:
        wati_by_l10[l10] = c

# Sheets: extract tel -> contexto
def col(headers, name_partial):
    """Find column index by partial header match (case-insensitive)."""
    name_partial = name_partial.lower()
    for i, h in enumerate(headers):
        if name_partial in h.lower(): return i
    return None

home_tels = {}      # l10 -> {'sub_barrio':..., 'cliente':...}
for r in sheets['Home'][1:]:
    if len(r) < 47: continue
    h = sheets['Home'][0]
    ic = col(h,'cliente'); itel = col(h,'tel'); isb = col(h,'sub barrio'); ib = col(h,'barrio')
    tel = r[itel] if itel is not None and itel < len(r) else ''
    l10 = last10(tel)
    if l10:
        home_tels[l10] = {
            'cliente': r[ic] if ic is not None and ic < len(r) else '',
            'sub_barrio': r[isb] if isb is not None and isb < len(r) else '',
            'barrio': r[ib] if ib is not None and ib < len(r) else '',
        }

pilar_tels = {}
h = sheets['Pilar'][0] if sheets['Pilar'] else []
for r in sheets['Pilar'][1:]:
    itel = col(h,'tel'); ic = col(h,'cliente'); ib = col(h,'barrio')
    tel = r[itel] if itel is not None and itel < len(r) else ''
    l10 = last10(tel)
    if l10:
        pilar_tels[l10] = {
            'cliente': r[ic] if ic is not None and ic < len(r) else '',
            'barrio': r[ib] if ib is not None and ib < len(r) else '',
        }

clubes_tels = {}
h = sheets['Clubes'][0] if sheets['Clubes'] else []
for r in sheets['Clubes'][1:]:
    itel = col(h,'tel'); ic = col(h,'cliente'); icl = col(h,'club')
    tel = r[itel] if itel is not None and itel < len(r) else ''
    l10 = last10(tel)
    if l10:
        clubes_tels[l10] = {
            'cliente': r[ic] if ic is not None and ic < len(r) else '',
            'club': r[icl] if icl is not None and icl < len(r) else '',
        }

red_tels = {}
h = sheets['Red'][0] if sheets['Red'] else []
for r in sheets['Red'][1:]:
    itel = col(h,'tel'); ic = col(h,'cliente'); ibp = col(h,'barrio')
    tel = r[itel] if itel is not None and itel < len(r) else ''
    l10 = last10(tel)
    if l10:
        red_tels[l10] = {
            'cliente': r[ic] if ic is not None and ic < len(r) else '',
            'barrio': r[ibp] if ibp is not None and ibp < len(r) else '',
        }

# Proveedores: telefonos
prov_tels = {}
h_p = sheets['Proveedores'][0] if sheets['Proveedores'] else []
itel_p = None
for i,h in enumerate(h_p):
    if 'tel' in h.lower() or 'whatsapp' in h.lower() or 'wpp' in h.lower():
        itel_p = i; break
in_p = col(h_p,'nombre') if h_p else None
if itel_p is not None:
    for r in sheets['Proveedores'][1:]:
        if itel_p < len(r):
            tel = r[itel_p]
            l10 = last10(tel)
            if l10:
                prov_tels[l10] = {'nombre': r[in_p] if in_p is not None and in_p < len(r) else ''}

print(f'Sheets indexados: Home={len(home_tels)} Pilar={len(pilar_tels)} Clubes={len(clubes_tels)} Red={len(red_tels)} Proveedores={len(prov_tels)}')

# ============== PLAN ==============
plan = []  # cada item: {phone, contact_name, current, set: [{name,value}], reasons:[]}

stats = Counter()

for c in wati:
    phone = c.get('wAid') or c.get('phone') or ''
    if not phone: continue
    l10 = last10(phone)
    name = c.get('fullName') or c.get('firstName') or ''
    cps = {cp['name']:cp['value'] for cp in c.get('customParams', [])}
    tipo = cps.get('tipo_de_contacto')
    zona = cps.get('zona')
    barrio = cps.get('barrio_privado')

    sets = []
    reasons = []

    # 1) Tipo no canónico -> mapear
    if tipo and tipo not in TIPOS_CANONICOS:
        if tipo in MAP_TIPO_NO_CANON:
            new_tipo = MAP_TIPO_NO_CANON[tipo]
            sets.append({'name':'tipo_de_contacto','value':new_tipo})
            reasons.append(f'tipo no canónico "{tipo}" -> "{new_tipo}"')
            stats['tipo_no_canon_mapeado'] += 1
            tipo = new_tipo
        else:
            # 'Contacto' u otros: limpiar para reclasificar abajo
            reasons.append(f'tipo no canónico "{tipo}" sin mapeo, intento reclasificar')
            tipo = None

    # 2) Sin tipo o reset -> cruzar contra Sheets
    if not tipo:
        if l10 in home_tels:
            sets.append({'name':'tipo_de_contacto','value':'Home'})
            reasons.append(f'matchea Home Sheet (cliente={home_tels[l10]["cliente"]})')
            tipo = 'Home'; stats['asignado_Home'] += 1
        elif l10 in clubes_tels:
            sets.append({'name':'tipo_de_contacto','value':'Clubes'})
            reasons.append(f'matchea Clubes Sheet (club={clubes_tels[l10]["club"]})')
            tipo = 'Clubes'; stats['asignado_Clubes'] += 1
        elif l10 in pilar_tels:
            sets.append({'name':'tipo_de_contacto','value':'Pilar'})
            reasons.append(f'matchea Pilar Sheet')
            tipo = 'Pilar'; stats['asignado_Pilar'] += 1
        elif l10 in red_tels:
            sets.append({'name':'tipo_de_contacto','value':'Red'})
            reasons.append(f'matchea Red Sheet')
            tipo = 'Red'; stats['asignado_Red'] += 1
        elif l10 in prov_tels:
            sets.append({'name':'tipo_de_contacto','value':'Proveedor'})
            reasons.append(f'matchea Proveedores Sheet ({prov_tels[l10]["nombre"]})')
            tipo = 'Proveedor'; stats['asignado_Proveedor'] += 1

    # 3) Si tipo=Home, asignar/normalizar zona y barrio_privado
    if tipo == 'Home' and l10 in home_tels:
        sb = home_tels[l10]['sub_barrio']
        new_zona = clasificar_zona(sb)
        if new_zona and new_zona != zona:
            sets.append({'name':'zona','value':new_zona})
            reasons.append(f'zona derivada de Sub Barrio="{sb}" -> {new_zona}')
            stats['zona_actualizada'] += 1
        # barrio_privado canonico
        if sb:
            new_barrio = normalizar_barrio_canon(sb)
            if new_barrio and new_barrio != barrio:
                sets.append({'name':'barrio_privado','value':new_barrio})
                reasons.append(f'barrio_privado normalizado "{barrio}" -> "{new_barrio}"')
                stats['barrio_normalizado'] += 1

    # 4) Aún sin tipo Home y barrio_privado mal escrito -> normalizar
    if barrio and not any(s['name']=='barrio_privado' for s in sets):
        new_barrio = normalizar_barrio_canon(barrio)
        if new_barrio != barrio:
            sets.append({'name':'barrio_privado','value':new_barrio})
            reasons.append(f'barrio_privado grafía "{barrio}" -> "{new_barrio}"')
            stats['barrio_normalizado'] += 1

    if sets:
        plan.append({
            'phone': phone,
            'name': name,
            'current': {'tipo':cps.get('tipo_de_contacto'),'zona':cps.get('zona'),'barrio':cps.get('barrio_privado')},
            'set': sets,
            'reasons': reasons,
        })

# ============== SAVE ==============
with open(PLAN_OUT,'w',encoding='utf-8') as f:
    json.dump(plan,f,ensure_ascii=False,indent=2)

with open(REPORT_OUT,'w',encoding='utf-8') as f:
    f.write(f'WATI BBDD — Plan de profesionalización\n')
    f.write(f'Total cambios planificados: {len(plan)}\n\n')
    f.write('Por categoría:\n')
    for k,v in stats.most_common():
        f.write(f'  {k}: {v}\n')
    f.write('\n--- Detalle (primeros 50) ---\n')
    for p in plan[:50]:
        f.write(f'\n📱 {p["phone"]} ({p["name"]})\n')
        f.write(f'   Antes: tipo={p["current"]["tipo"]} zona={p["current"]["zona"]} barrio={p["current"]["barrio"]}\n')
        for s in p['set']:
            f.write(f'   SET {s["name"]} = {s["value"]}\n')
        for r in p['reasons']:
            f.write(f'   • {r}\n')

print(f'\nPlan generado: {len(plan)} contactos a actualizar')
print('Stats:')
for k,v in stats.most_common():
    print(f'  {k}: {v}')
print(f'\nArchivos:\n  {PLAN_OUT}\n  {REPORT_OUT}')
