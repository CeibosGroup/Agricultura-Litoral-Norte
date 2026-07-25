"""
build_finn.py
Regenera finn.js a partir de Reporte.xlsx y Plan_de_Siembra_26-27_-_Uruguay.xlsx
Output: finn.js con FINN_FULL comprimido + CROP_LOTE

Uso:
  python3 build_finn.py

Requiere:
  pip install pandas openpyxl
  Archivos en el mismo directorio:
    - Reporte.xlsx
    - Plan_de_Siembra_26-27_-_Uruguay.xlsx
"""

import pandas as pd, json, datetime, re

# ── CONFIG ────────────────────────────────────────────────────
REPORTE = "Reporte.xlsx"
PLAN    = "Plan_de_Siembra_26-27_-_Uruguay.xlsx"
OUTPUT  = "finn.js"

ZONA_MAP = {
    'UY_Litoral Norte':'Litoral Norte','UY_Salto':'Salto',
    'UY_Ruta 20':'Ruta 20','UY_Durazno':'Durazno','UY_Tacuarembo':'Tacuarembó'
}
ACT_CULT = {
    'COLZA':'Cz','TRIGO':'Tr','CEBADA':'Ceb','AVENA':'Cobertura',
    'CARINATA':'Carinata','COBERTURA':'Cobertura',
    'COLZA RIEGO':'Cz','TRIGO RIEGO':'Tr','CEBADA RIEGO':'Ceb',
    'VERDEOS INVIERNO':'Cobertura','PASTURA':'Cobertura',
    'RYE GRASS':'Cobertura','FESTUCA':'Cobertura','TREBOL':'Cobertura',
    'LOTUS':'Cobertura','ALFALFA':'Cobertura','DACTYLIS':'Cobertura',
}

# ── HELPERS ───────────────────────────────────────────────────
def clean_estab(e):
    s = str(e or '').replace('UY_','')
    for o,n in [('EL ','El '),('LA ','La '),('LOS ','Los '),('LAS ','Las ')]:
        s = s.replace(o, n)
    return s.replace('_TECNO',' Tecno').strip()

def clean_lote(estab, lote):
    s = str(lote or '').strip()
    if s.upper().startswith('UY_'): s = s[3:]
    ec = str(estab or '').replace('UY_','').strip()
    ec_parts = [re.escape(p) for p in re.split(r'[_\-\s]+', ec) if p]
    if ec_parts:
        pat = r'^' + r'[_\-\s]*'.join(ec_parts) + r'[_\-\s]+'
        s = re.sub(pat, '', s, flags=re.IGNORECASE)
    for p in ['El Malacara-','El Juncal-','El Rincon-','El Trebol-','Los Tilos-',
              'Monte Cristo-','La Campana-','El Laurel-','El Arbolito-','El Tajamar-',
              'El Totoral-','Los Nogales-','Torre Alta_','Don Venancio-','La Macarena-',
              'Manantiales-','Bichadero-','Buena Vista-','Cinco Sauce-','Laguna Negra-',
              'La Cautiva-','Las Margaritas-','San Juan-']:
        if s.lower().startswith(p.lower()): s=s[len(p):]; break
    s = s.strip()
    if not s or 'generico' in s.lower(): return None
    if 'SANTA RAFAELA' in str(estab).upper(): return 'L'+s
    m = re.match(r'^0*(\d+[A-Za-z]*)$', s)
    return m.group(1).upper() if m else s

# ── KEYWORD LISTS ─────────────────────────────────────────────
SKIP = {'GASOIL (STOCK)','ACIDO FOSFORICO 23 G','ALCOHOL LINEAL ETOXILADO 32.7%',
        'BONIFICACION SIEMBRA','SEMILLA RINDE','BONIFICACION'}
SKIP_P = ['CAMPO FULL SIL','ADHERENTE','COADYUVANTE','SURFACTANTE','HAMMER S',
          'ACEITE DE SOJA','ACEITE VEGETAL','SILICONADO','HORAS TRACTOR',
          'RIEGO (MILIMETROS)','ROTATIVA','CINCEL']

HERB = ['GLIFOSATO','GLYPHOSATE','CLETODIM','PARAQUAT','DIQUAT','2.4D','2,4D',
        'ATRAZINA','LONTREL','AXIAL','HUSSAR','KORVETTO','TORDON','TRONADOR',
        'AMINA','STARANE','CLOPIRALID','FLUROXYPYR','METSULFURON','AMINOPYRALID',
        'PINOXADEM','PICLORAM','IMAZAPIR','MCPA','DICAMBA','FLUAZIFOP','HALAUX',
        'HALAUXIFEN','SAL AMINA','FLUMETSULAN','2-4DB','IODOSULFURON','CLOPYRALID',
        'NICOSULFURON','FLUMIOXAZIM','SAFLUFENACIL','OXIFLUORFEN','DIURON',
        'PIROXASULFONA','THIENCARBAZONE','ISOXAFLUTOLE','S-METOLACLOR','GLUFOSINATO',
        'FLORASULAM','FLUROXIPIR','2,4-D','HALAUXYFEN']

FUNG = ['TRIFLOX','TEBUCON','AZOXIST','EPOXIC','BIXAFEN','PROTHIO','CARBENDAZIM',
        'IPRODIONE','MANCOZEB','PROPICONAZOL','FLUXAPYR','PIRACLOSTROBIN',
        'PROTIOCONAZOL','VITAVAX']

INSEC = ['LAMBDA','CIPERMETRINA','IMIDACLOPRID','CLORPIRIFOS','DELTAMETRINA',
         'BIFENTRIN','CIPERM','ACETAMIPRID','TIAMETOXAM','PIRIMICARB','DIMETOATO',
         'CLORANTRANILIPROL','SPINOSAD','FIPRONIL','ABAMECTINA','ISOCYCLOSERAM']

FERT_N = ['UREA AZUFRADA','UREA (46','UREA (0','UAN ','AZUFRADA','ENTEC',
          'CAN ','SULFATO DE AMONIO','VITANIT','N SOL']

FERT_BAS = ['SPT','SPS','FOSFATO','MAP','DAP','FOSZINC','CLORURO DE POTASIO',
            '0-0-60','YESO','SULFATO DE CALCIO','AZUFERTIL','KCL',
            '12-45','0-46','46-0','12-46','SUPERFOSFATO','SUPERTRIPLE','MICROSTAR',
            'NITROGENO 120GR']

HUM = ['HUMICO','HÚMICO','HUME','BIOESTIM','HYDRA','FULVIC','ALGA','BIOSTIM',
       'VITABALANCE','VITAFERTIL','VITAROOT','VITAGROW','BIOAVANZA']

FIELDS = ['barb','dg','rep','siem','fert_bas','fert_n','fung','insec','znf','hum','re1']

def get_field(tipo, labor, fecha_dt, siem_dt):
    l = str(labor or '').upper().strip()
    for s in SKIP:
        if s.upper() in l: return None
    for s in SKIP_P:
        if s.upper() in l: return None
    if tipo == 'Laboreo':
        if 'RESIEMBRA' in l: return 're1'
        if 'SIEMBRA' in l: return 'siem'
        if any(x in l for x in ['DISQUERA','DISCO ACONDICIONADOR']): return 'barb'
        return None
    before = siem_dt is None or (fecha_dt is not None and fecha_dt < siem_dt)
    if 'SEMILLA' in l: return 'siem'
    if any(x in l for x in HERB):
        if 'DIQUAT' in l: return 'dg'
        return 'barb' if before else 'rep'
    if any(x in l for x in FERT_N): return 'fert_n'
    if any(x in l for x in FERT_BAS): return 'fert_bas'
    if any(x in l for x in FUNG): return 'fung'
    if any(x in l for x in INSEC): return 'insec'
    if 'ZINC' in l and not any(x in l for x in ['12-45','FOSZINC','12-46','MICROSTAR','SULFATO']): return 'znf'
    if any(x in l for x in HUM): return 'hum'
    return None

# ── PARSE PLAN DE SIEMBRA ─────────────────────────────────────
print("Leyendo Plan de Siembra...")
ps_df = pd.read_excel(PLAN, sheet_name="Plan de Siembra", header=2)
plan_siem = {}
for _, row in ps_df.iterrows():
    estab = str(row.get('Establecimiento','') or '')
    if not estab or estab=='nan': continue
    zona_raw = str(row.get('Zona','') or '')
    if zona_raw not in ZONA_MAP: continue
    zona = ZONA_MAP[zona_raw]
    campo = clean_estab(estab)
    lote = clean_lote(estab, row.get('Lote Ceres',''))
    if not lote: continue
    fecha_raw = row.get('Fecha Siembra')
    fecha = pd.Timestamp(fecha_raw).strftime('%d/%m/%Y') if pd.notna(fecha_raw) else None
    var = str(row.get('Variedad/Hibrido','') or '').strip()
    var = None if var in ('','nan') else var
    cult_raw = str(row.get('Cultivo','') or '').upper().strip()
    cult = next((v for k,v in ACT_CULT.items() if k in cult_raw), None)
    has = row.get('Has CG')
    has = float(has) if pd.notna(has) else None
    key = f"{zona}|{campo}|{lote}"
    if key not in plan_siem:
        plan_siem[key] = {'fecha_siem':fecha,'var':var,'cult':cult,'has':has}
    elif fecha and not plan_siem[key]['fecha_siem']:
        plan_siem[key]['fecha_siem'] = fecha

print(f"  {len(plan_siem)} lotes en plan de siembra")

# ── PARSE REPORTE ─────────────────────────────────────────────
print("Leyendo Reporte...")
df = pd.read_excel(REPORTE, sheet_name="hoja1")
print(f"  {len(df)} filas")

# Fecha real de siembra por lote (leída del propio Reporte), prioridad sobre Plan de Siembra
actual_siem = {}
for _, row in df.iterrows():
    zona_raw = str(row.get('Zona','') or '')
    if zona_raw not in ZONA_MAP: continue
    zona = ZONA_MAP[zona_raw]
    estab = str(row.get('Establecimiento','') or '')
    campo = clean_estab(estab)
    lote = clean_lote(estab, row.get('Lote',''))
    if not lote: continue
    tipo = str(row.get('Tipo','') or '')
    labor_u = str(row.get('Labor/producto','') or '').strip().upper()
    is_siem = (tipo == 'Laboreo' and 'RESIEMBRA' not in labor_u and 'SIEMBRA' in labor_u) or \
              (tipo == 'Insumo' and 'SEMILLA' in labor_u)
    if not is_siem: continue
    fecha_raw = row.get('Fecha')
    if pd.isna(fecha_raw): continue
    fecha_dt = pd.Timestamp(fecha_raw)
    key = f"{zona}|{campo}|{lote}"
    if key not in actual_siem or fecha_dt < actual_siem[key]:
        actual_siem[key] = fecha_dt

print(f"  {len(actual_siem)} lotes con fecha de siembra real detectada en el Reporte")

finn_full = {}
crop_lote = {}

for _, row in df.iterrows():
    zona_raw = str(row.get('Zona','') or '')
    if zona_raw not in ZONA_MAP: continue
    zona = ZONA_MAP[zona_raw]
    estab = str(row.get('Establecimiento','') or '')
    campo = clean_estab(estab)
    lote = clean_lote(estab, row.get('Lote',''))
    if not lote: continue

    fecha_raw = row.get('Fecha')
    fecha_dt = pd.Timestamp(fecha_raw) if pd.notna(fecha_raw) else None
    fecha = fecha_dt.strftime('%d/%m/%Y') if fecha_dt else None
    tipo = str(row.get('Tipo','') or '')
    labor = str(row.get('Labor/producto','') or '').strip()
    cant_ha = row.get('Cant./ha')
    unidad = str(row.get('Unidad','') or '').strip()
    actividad = str(row.get('Actividad','') or '').upper().strip()

    key = f"{zona}|{campo}|{lote}"
    ps = plan_siem.get(key, {})
    siem_dt = actual_siem.get(key)
    if siem_dt is None:
        siem_str = ps.get('fecha_siem')
        if siem_str:
            try:
                p = siem_str.split('/')
                siem_dt = pd.Timestamp(f"{p[2]}-{p[1]}-{p[0]}")
            except: pass

    field = get_field(tipo, labor, fecha_dt, siem_dt)
    if field is None: continue

    cult = next((v for k,v in ACT_CULT.items() if k in actividad), None)
    if cult and key not in crop_lote: crop_lote[key] = cult

    if zona not in finn_full: finn_full[zona] = {}
    if campo not in finn_full[zona]: finn_full[zona][campo] = {}
    if lote not in finn_full[zona][campo]:
        finn_full[zona][campo][lote] = {
            'cult': crop_lote.get(key) or ps.get('cult'),
            'var': ps.get('var'),
            'has': ps.get('has'),
            **{f:{'done':False,'fecha':None,'items':[]} for f in FIELDS}
        }

    d = finn_full[zona][campo][lote]
    if not d['cult'] and cult: d['cult'] = cult

    if field in d:
        if fecha:
            d[field]['done'] = True
            if not d[field]['fecha']: d[field]['fecha'] = fecha
        if labor and field != 'siem':
            item = {'l': labor}
            if pd.notna(cant_ha) and float(cant_ha) not in (0,1):
                item['q'] = round(float(cant_ha), 2)
            if unidad and unidad not in ('Ha.',''):
                item['u'] = unidad
            if not any(x.get('l')==labor for x in d[field]['items']):
                d[field]['items'].append(item)
        elif field == 'siem' and pd.notna(cant_ha) and float(cant_ha) > 1:
            d['siem']['kg_ha'] = round(float(cant_ha), 0)

# Merge plan dates/vars
for zona, campos in finn_full.items():
    for campo, lotes in campos.items():
        for lote, d in lotes.items():
            k = f"{zona}|{campo}|{lote}"
            ps = plan_siem.get(k, {})
            if ps.get('var') and not d.get('var'): d['var'] = ps['var']
            if ps.get('fecha_siem') and not d['siem']['fecha']:
                d['siem']['fecha'] = ps['fecha_siem']
                d['siem']['done'] = True
            if ps.get('cult') and not d.get('cult'): d['cult'] = ps['cult']
            if ps.get('has') and not d.get('has'): d['has'] = ps['has']

# Add lotes from plan with no reporte data
for key, ps in plan_siem.items():
    zona, campo, lote = key.split('|', 2)
    if zona not in finn_full: finn_full[zona] = {}
    if campo not in finn_full[zona]: finn_full[zona][campo] = {}
    if lote not in finn_full[zona][campo]:
        finn_full[zona][campo][lote] = {
            'cult': ps.get('cult'), 'var': ps.get('var'), 'has': ps.get('has'),
            **{f:{'done':False,'fecha':None,'items':[]} for f in FIELDS}
        }

# Clean empty fields
def clean_obj(obj):
    if isinstance(obj, dict):
        return {k:clean_obj(v) for k,v in obj.items()
                if v is not None and v is not False and v!=[] and v!={} and str(v)!='nan'}
    if isinstance(obj, list): return [clean_obj(i) for i in obj if i]
    return obj

finn_clean = clean_obj(finn_full)

# Compress keys
FM = {'barb':'b','dg':'dg','rep':'r','siem':'s','fert_bas':'fb','fert_n':'fn',
      'fung':'fu','insec':'ins','znf':'z','hum':'h','re1':'re',
      'cult':'c','var':'v','done':'d','fecha':'f','items':'i','kg_ha':'k','has':'ha'}

def compress(d):
    if isinstance(d, dict):
        out = {}
        for k,v in d.items():
            nk = FM.get(k, k)
            nv = compress(v)
            if nv is not None and nv is not False and nv != [] and nv != {}:
                out[nk] = nv
        return out or None
    if isinstance(d, list): return [compress(i) for i in d if compress(i)] or None
    return d

finn_c = {}
for zona, campos in finn_clean.items():
    finn_c[zona] = {}
    for campo, lotes in campos.items():
        finn_c[zona][campo] = {}
        for lote, d in lotes.items():
            cd = compress(d)
            if cd: finn_c[zona][campo][lote] = cd

finn_js = json.dumps(finn_c, ensure_ascii=True, separators=(',',':'))
crop_js = json.dumps(crop_lote, ensure_ascii=True, separators=(',',':'))

content = f"var FINN_FULL={finn_js};\nvar CROP_LOTE={crop_js};\n"
with open(OUTPUT, 'w') as f:
    f.write(content)

# Stats
total = sum(sum(len(l) for l in c.values()) for c in finn_c.values())
print(f"\nfinn.js generado: {len(content)//1024} KB")
print(f"Total lotes: {total}")
for zona, campos in sorted(finn_c.items()):
    tl = sum(len(l) for l in campos.values())
    print(f"  {zona}: {tl} lotes")
print("\nListo. Subí finn.js al repo.")
