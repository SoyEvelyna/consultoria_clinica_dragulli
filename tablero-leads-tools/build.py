"""Procesa la planilla DRA_GULLI Seguimiento de leads y genera los datos del tablero.
Uso: python3 -I build.py PLANILLA.xlsx CORTE(YYYY-MM-DD) SALIDA_DIR
Genera leads.json, turnos.json, leads_dra_gulli_sin_datos_personales.xlsx,
excel_anonimo.json y excel_original.json en SALIDA_DIR."""
import sys, os, re, json, base64, unicodedata
import pandas as pd

SRC, CORTE, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
os.makedirs(OUT, exist_ok=True)
CUT = pd.Timestamp(CORTE + ' 23:59:59')
x = pd.read_excel(SRC, sheet_name=None)


def norm(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return None
    s = re.sub(r'\s+', ' ', str(s)).strip()
    return s or None


def strip(s):
    return unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode().lower()


def k8(p):
    if p is None or (isinstance(p, float) and pd.isna(p)):
        return None
    s = str(p).strip()
    if s.endswith('.0'):
        s = s[:-2]
    s = re.sub(r'\D', '', s)
    return s[-8:] if len(s) >= 8 else None


rows = []


def add(hoja, d, fecha, trat=None, estado=None, motivo=None, soporte=None, asesora=None,
        coord=None, servicio=None, conocio=None, stray_turno=False):
    for _, r in d.iterrows():
        fe = pd.to_datetime(r.get(fecha), errors='coerce')
        if pd.isna(fe):
            continue
        est = norm(r.get(estado)) if estado else None
        if stray_turno and len(r) > 13 and norm(r.iloc[13]) == 'Turno':
            est = 'Turno'
        rows.append(dict(
            hoja=hoja, fecha=fe, canal=norm(r.get('Canal')),
            trat=norm(r.get(trat)) if trat else None, estado=est,
            motivo=norm(r.get(motivo)) if motivo else None,
            soporte=norm(r.get(soporte)) if soporte else None,
            asesora=norm(r.get(asesora)) if asesora else None,
            coord=(norm(r.iloc[coord]) if isinstance(coord, int) else norm(r.get(coord)) if coord else None),
            servicio=norm(r.get(servicio)) if servicio else None,
            conocio=norm(r.get(conocio)) if conocio else None,
            pk=k8(r.get('Numero_WA'))))


def sheet(name):
    return x[name] if name in x else None


if sheet('ASESORIA-MD-DANI') is not None:
    add('IG-Dani DM', x['ASESORIA-MD-DANI'], 'Fecha_Entrada', 'Objetivo_Consulta', 'Estado', 'Observaciones')
if sheet('No COMPLETARON-MD-DANI') is not None:
    add('IG-Dani No completaron', x['No COMPLETARON-MD-DANI'], 'Fecha_Entrada', None, None, 'Observaciones')
add('Turnos', x['TURNOS'], 'Fecha_Entrada', 'Servicio_Interes', 'Estado', coord=24, servicio='Servicio_Interes', conocio='Como nos conocio')
add('Asesoría', x['ASESORIA'], 'Fecha_Entrada', 'Tratamiento', 'Estado', 'Observaciones', asesora='Asesora')
add('Pauta', x['PAUTA'], 'Fecha_Entrada', 'Tratamiento', 'Estado', 'Observaciones', 'Soporte', 'Asesora', stray_turno=True)
add('No completaron', x['NO COMPLETARON'], 'Fecha_Entrada', 'Tratamiento', None, 'Observaciones', 'Soporte')
for nm, sh in [('Dermatología', 'DERMATOLOGIA'), ('Estética', 'ESTETICA')]:
    if sheet(sh) is not None:
        add(nm, x[sh], 'Fecha_Entrada', 'TRATAMIENTO ', 'Estado', coord='Asesora', servicio='TRATAMIENTO ')
add('Mail', x['MAIL'], 'Fecha_Entrada')
for sh, col, nm in [('REMARKETING24', 'Fecha_Remarketing24hs', 'Remarketing 24hs'),
                    ('Remarketing24hs', 'Fecha_Remarketing24hs', 'Remarketing 24hs'),
                    ('REMARKETING48', 'Fecha_Remarketing48hs', 'Remarketing 48hs')]:
    if sheet(sh) is not None:
        add(nm, x[sh], col, 'Tratamiento', 'Estado', 'Observaciones', 'Soporte', asesora='ASESORA')

df = pd.DataFrame(rows)
df['_k'] = [(p, f.strftime('%Y-%m-%d %H:%M:%S')) for p, f in zip(df.pk, df.fecha)]
nuevas = set(df.loc[df.hoja.isin(['Dermatología', 'Estética']), '_k'])
dup = (df.hoja == 'Turnos') & df._k.isin(nuevas) & df.pk.notna()
rmk = df.hoja.str.startswith('Remarketing')
dup2 = rmk & df.duplicated(subset=['hoja', '_k'])
print('TURNOS repetidos en DERMATOLOGIA/ESTETICA:', int(dup.sum()), '| remarketing repetidos:', int(dup2.sum()))
df = df[~dup & ~dup2].drop(columns='_k')
df = df[df.fecha <= CUT].copy()
for c in ['canal', 'trat', 'estado', 'motivo', 'soporte', 'asesora', 'coord', 'servicio', 'conocio', 'pk']:
    df[c] = df[c].astype(object).where(df[c].notna(), None)

# --- normalizaciones ---
df['cuenta'] = ['Dani' if c and 'DANI' in c.upper() else 'Clínica' for c in df.canal]


def via(h, c):
    c = (c or '').upper()
    if h == 'Mail' or c == 'MAIL':
        return 'Mail'
    if c.startswith('PAUTA'):
        return 'Pauta'
    if c == 'WHATSAPP':
        return 'WhatsApp'
    return 'Instagram DM'


df['via'] = [via(h, c) for h, c in zip(df.hoja, df.canal)]
mt = df.trat.isin(['Precio', 'Distancia de la clinica'])
df.loc[mt & df.motivo.isna(), 'motivo'] = df.loc[mt & df.motivo.isna(), 'trat']
df.loc[mt, 'trat'] = None
df.loc[df.hoja == 'Mail', 'trat'] = None


def trat(t):
    if not t:
        return 'Sin dato'
    s = strip(t)
    if 'endolift' in s:
        return 'Endolift corporal' if ('corporal' in s or 'celulitis' in s) else ('Endolift facial' if 'facial' in s else 'Endolift (sin especificar)')
    if 'co2' in s or 'laser facial' in s or 'cicatri' in s:
        return 'Láser CO2'
    if 'ultra' in s:
        return 'Ultraformer'
    if 'cool' in s or 'adiposidad' in s or 'papada' in s:
        return 'Cooltech'
    if 'sculptra' in s or 'scultra' in s:
        return 'Sculptra corporal' if 'corporal' in s else 'Sculptra facial'
    if 'mancha' in s or 'nordlys' in s:
        return 'Manchas'
    if any(k in s for k in ['inyect', 'toxina', 'botox', 'relleno', 'ah ', 'radiesse', 'armoniz']):
        return 'Inyectables'
    if 'dermato' in s or 'acne' in s:
        return 'Consulta dermatológica'
    if 'estetica' in s or 'rejuven' in s or 'corporal' in s:
        return 'Consulta estética'
    return 'Otros'


df['tratamiento'] = df.trat.map(trat)
ABAND = ('No completaron', 'IG-Dani No completaron')


def estado(e, h):
    if h in ABAND:
        return 'Abandonó el formulario'
    if not e:
        return 'Sin estado cargado'
    s = strip(e)
    if s.startswith('turno'):
        return 'Turno'
    if 'presupuesto' in s or s.startswith('plan'):
        return 'Presupuesto enviado'
    if 'pendiente' in s:
        return 'Pendiente'
    if 'escribio' in s:
        return 'No escribió al WhatsApp'
    if 'contest' in s:
        return 'Dejó de contestar'
    if 'rechaz' in s:
        return 'Rechazado'
    return 'Sin estado cargado'


df['estado_g'] = [estado(e, h) for e, h in zip(df.estado, df.hoja)]


def motivo(m, h):
    if not m:
        return None
    s = strip(m)
    if h in ABAND or h.startswith('Remarketing'):
        for k, v in [('nombre', 'No dio su nombre'), ('email', 'No dio su email'), ('telefono', 'No dio su teléfono'),
                     ('objetivo', 'No contó qué busca'), ('subcategoria', 'No eligió tratamiento')]:
            if k in s:
                return v
        return 'Otro dato faltante'
    if 'precio' in s:
        return 'Precio'
    if 'deribo' in s or 'derivo' in s:
        return 'No pasó al WhatsApp'
    if 'contesto' in s:
        return 'No contestó más'
    if 'adelante' in s:
        return 'Para más adelante'
    if 'distancia' in s:
        return 'Distancia'
    return None


df['motivo_g'] = [motivo(m, h) for m, h in zip(df.motivo, df.hoja)]
ALIAS = {'Ailu': 'Ailen', 'Ailén': 'Ailen'}
df['asesora_n'] = [ALIAS.get(a.strip().capitalize(), a.strip().capitalize()) if a else None for a in df.asesora]

rm = df[df.hoja.str.startswith('Remarketing')].copy()
L = df[~df.hoja.str.startswith('Remarketing')].copy()
comp = L[~L.hoja.isin(ABAND)]
rec, tpost = [], []
for r in rm.itertuples():
    if not r.pk:
        rec.append(False); tpost.append(False); continue
    m = comp[(comp.pk == r.pk) & (comp.fecha > r.fecha)]
    rec.append(len(m) > 0); tpost.append(bool((m.estado_g == 'Turno').any()))
rm['recuperado'] = rec
rm['turno_post'] = tpost

allr = pd.concat([L.assign(tipo='lead'), rm.assign(tipo='remarketing')]).sort_values('fecha', kind='stable')
def nn(v):
    return None if v is None or (isinstance(v, float) and pd.isna(v)) else v


pid = {}
out = []
for i, r in enumerate(allr.itertuples()):
    if nn(r.pk):
        if r.pk not in pid:
            pid[r.pk] = f'p{len(pid) + 1:04d}'
        per = pid[r.pk]
    else:
        per = f'x{i}'
    d = dict(fecha=r.fecha.strftime('%Y-%m-%d'), tipo=r.tipo, cuenta=r.cuenta, via=r.via, hoja=r.hoja,
             tratamiento=r.tratamiento, estado=r.estado_g, motivo=nn(r.motivo_g))
    if r.tipo == 'remarketing':
        d.update(remarketing=r.hoja.replace('Remarketing ', ''), recuperado=bool(r.recuperado),
                 turno_post=bool(r.turno_post), estado_rmk=(nn(r.estado) or 'Sin estado'))
    d.update(asesora=nn(r.asesora_n),
             formato=((nn(r.soporte) or 'Sin dato').capitalize() if r.via == 'Pauta' else None), persona=per)
    out.append(d)
json.dump(out, open(os.path.join(OUT, 'leads.json'), 'w'), ensure_ascii=False)

# --- turnos ---
DISP = {'dermatológica': 'Consulta dermatológica', 'consulta dermatológica': 'Consulta dermatológica',
        'consulta estética': 'Consulta estética', 'consulta corporal': 'Consulta corporal', 'endolift': 'Endolift',
        'co2': 'Láser CO2', 'láser co2': 'Láser CO2', 'láser nordlys': 'Manchas (Nordlys)',
        'nordlys - manchas': 'Manchas (Nordlys)', 'tratamiento de láser facial': 'Láser facial'}


def disp(t):
    if not isinstance(t, str) or not t.strip():
        return 'Sin dato en la planilla'
    s = t.strip()
    return DISP.get(s.lower(), s[0].upper() + s[1:])


HO = {'Turnos': 'Hoja TURNOS', 'Dermatología': 'Hoja DERMATOLOGIA', 'Estética': 'Hoja ESTETICA',
      'Pauta': 'Pauta', 'Asesoría': 'Asesoría', 'IG-Dani DM': 'IG-Dani', 'Mail': 'Mail'}
T = L[L.estado_g == 'Turno'].sort_values('fecha', kind='stable')
tl = []
for i, r in enumerate(T.itertuples(), 1):
    resp = nn(r.coord) or nn(r.asesora)
    resp = ALIAS.get(resp.strip().capitalize(), resp.strip().capitalize()) if resp else 'Sin asignar'
    tl.append(dict(clave=f'T-{i:03d}', fecha=r.fecha.strftime('%Y-%m-%d'), cuenta=r.cuenta, via=r.via,
                   canal=('Pauta ' if r.via == 'Pauta' else 'Instagram ' if r.via == 'Instagram DM' else r.via + ' ') + r.cuenta,
                   origen=HO.get(r.hoja, r.hoja), tratamiento=disp(nn(r.servicio) or nn(r.trat)),
                   responsable=resp, como_nos_conocio=nn(r.conocio) or 'Sin dato'))
json.dump(tl, open(os.path.join(OUT, 'turnos.json'), 'w'), ensure_ascii=False)

# --- Excel sin datos personales ---
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

wb = Workbook(); ws = wb.active; ws.title = 'Leads'


def hoja_xl(ws, rs, cols):
    ws.append([c[1] for c in cols])
    for r in rs:
        ws.append([r.get(k) for k, _ in cols])
    for c in ws[1]:
        c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='2F6FD6')
    for i, (k, h) in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(i)].width = max(12, min(30, len(h) + 6))
    ws.freeze_panes = 'A2'; ws.auto_filter.ref = ws.dimensions


cols = [('fecha', 'Fecha'), ('cuenta', 'Cuenta'), ('via', 'Vía'), ('hoja', 'Hoja de origen'), ('tratamiento', 'Tratamiento'),
        ('estado', 'Estado'), ('motivo', 'Motivo'), ('formato', 'Formato de anuncio'), ('asesora', 'Comercial'), ('persona', 'Persona (código)')]
hoja_xl(ws, [r for r in out if r['tipo'] == 'lead'], cols)
rcols = [('fecha', 'Fecha del mensaje'), ('remarketing', 'Tipo'), ('cuenta', 'Cuenta'), ('via', 'Vía'), ('tratamiento', 'Tratamiento'),
         ('motivo', 'Dato que faltaba'), ('estado_rmk', 'Respuesta'), ('recuperado', 'Volvió y completó'), ('turno_post', 'Turno después'),
         ('formato', 'Formato de anuncio'), ('persona', 'Persona (código)')]
hoja_xl(wb.create_sheet('Remarketing'), [dict(r, recuperado='Sí' if r.get('recuperado') else 'No', turno_post='Sí' if r.get('turno_post') else 'No')
                                         for r in out if r['tipo'] == 'remarketing'], rcols)
hoja_xl(wb.create_sheet('Turnos'), tl, [('clave', 'N.º'), ('fecha', 'Fecha'), ('cuenta', 'Cuenta'), ('via', 'Vía'), ('canal', 'Canal'), ('origen', 'Hoja'),
                                         ('tratamiento', 'Tratamiento'), ('responsable', 'Responsable'), ('como_nos_conocio', 'Cómo nos conoció')])
n = wb.create_sheet('Notas'); dd = CORTE[8:10] + '/' + CORTE[5:7] + '/' + CORTE[:4]
for line in [f'Leads Dra. Gulli · datos hasta el domingo {dd} inclusive',
             'Sin nombres, teléfonos, mails ni DNI: cada persona es un código.',
             'Consulta = cada entrada en PAUTA, ASESORIA, NO COMPLETARON, TURNOS, DERMATOLOGIA, ESTETICA, MAIL e IG-Dani.',
             'Turno = Estado "Turno" en cualquier hoja.', 'Comercial = columna Asesora (Gaby / Ailen).',
             'Remarketing = mensajes a 24 y 48 hs a quien no terminó el formulario; no suman consultas. Volvió = misma persona completó sus datos después del mensaje.']:
    n.append([line])
n.column_dimensions['A'].width = 90
anon = os.path.join(OUT, 'leads_dra_gulli_sin_datos_personales.xlsx'); wb.save(anon)
json.dump([{'archivo': f'leads_dra_gulli_sin_datos_personales_{CORTE}.xlsx', 'b64': base64.b64encode(open(anon, 'rb').read()).decode()}],
          open(os.path.join(OUT, 'excel_anonimo.json'), 'w'))
json.dump([{'archivo': f'DRA_GULLI_Seguimiento_de_leads_{CORTE}.xlsx', 'b64': base64.b64encode(open(SRC, 'rb').read()).decode()}],
          open(os.path.join(OUT, 'excel_original.json'), 'w'))

Lr = [r for r in out if r['tipo'] == 'lead']
print('consultas', len(Lr), 'turnos', sum(r['estado'] == 'Turno' for r in Lr), 'remarketing', len(out) - len(Lr), 'max', max(r['fecha'] for r in out))
