"""Genera index.html a partir de la planilla (xlsx exportado) y data/fotos.json.
Uso: python3 build/build.py planilla.xlsx "27-09-2026"
"""
import json, sys, html, os
from openpyxl import load_workbook
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
xlsx, fecha = sys.argv[1], sys.argv[2]
fotos = json.load(open(os.path.join(ROOT, 'data', 'fotos.json'), encoding='utf-8'))
ws = load_workbook(xlsx, data_only=True)['Productos']
hdr = [c.value for c in ws[1]]
col = {h: i for i, h in enumerate(hdr)}
def num(v):
    if v in (None, ''): return None
    try: return round(float(str(v).replace('$', '').replace('.', '').replace(',', '.')) if isinstance(v, str) else float(v))
    except ValueError: return None
items = []
for r in ws.iter_rows(min_row=2, values_only=True):
    nombre = r[col['Producto']]
    if not nombre: continue
    it = {'n': str(nombre).strip(), 'c': str(r[col['Categoría']]).strip()}
    p = num(r[col['Precio venta unidad ($)']]); k = num(r[col['Pack venta (unidades)']]); kp = num(r[col['Precio venta pack ($)']])
    if p: it['p'] = p
    if k and kp: it['k'] = int(k); it['kp'] = kp
    if it['n'] in fotos and os.path.exists(os.path.join(ROOT, fotos[it['n']])): it['img'] = fotos[it['n']]
    items.append(it)
con = sum(1 for i in items if 'p' in i)
note = '' if con else 'Aún no hay precios cargados. Los precios aparecerán aquí a medida que se completen.'
t = open(os.path.join(ROOT, 'build', 'template.html'), encoding='utf-8').read()
js = json.dumps(items, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
head = ('<!doctype html><html lang="es"><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
        '<meta name="theme-color" content="#8B1A1A"><meta name="robots" content="noindex">'
        '<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 64 64%22%3E%3Crect width=%2264%22 height=%2264%22 rx=%2214%22 fill=%22%238B1A1A%22/%3E%3Ctext x=%2232%22 y=%2244%22 font-size=%2232%22 font-family=%22Arial%22 font-weight=%22bold%22 text-anchor=%22middle%22 fill=%22%23FFD84D%22%3E$%3C/text%3E%3C/svg%3E">'
        '<link rel="apple-touch-icon" href="img/icono.png">'
        '<style>html{-webkit-text-size-adjust:100%}body{margin:0}[hidden]{display:none!important}img{max-width:100%}</style></head><body>')
out = head + t.replace('__DATA__', js).replace('__FECHA__', html.escape(fecha)).replace('__NOTE__', f'<p class="note">{html.escape(note)}</p>' if note else '') + '</body></html>'
open(os.path.join(ROOT, 'index.html'), 'w', encoding='utf-8').write(out)
print(f'{len(items)} productos, {con} con precio, {sum(1 for i in items if "img" in i)} con foto')
