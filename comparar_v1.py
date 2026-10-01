import zipfile
import xml.etree.ElementTree as ET

NS = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def leer_shared_strings(z):
    if 'xl/sharedStrings.xml' not in z.namelist():
        return []
    root = ET.fromstring(z.read('xl/sharedStrings.xml'))
    strings = []
    for si in root.findall('x:si', NS):
        t = si.find('x:t', NS)
        if t is not None:
            strings.append(t.text or '')
        else:
            partes = si.findall('.//x:t', NS)
            strings.append(''.join(p.text or '' for p in partes))
    return strings

def leer_hoja(z, sheet_name):
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rid_target = {r.get('Id'): r.get('Target') for r in rels}
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    for sh in wb.findall('.//x:sheet', NS):
        if sh.get('name') == sheet_name:
            rid = sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            path = rid_target[rid].lstrip('/')
            shared = leer_shared_strings(z)
            root = ET.fromstring(z.read(path))
            filas = {}
            for row in root.findall('.//x:row', NS):
                rnum = int(row.get('r'))
                celdas = {}
                for c in row.findall('x:c', NS):
                    ref = c.get('r', '')
                    col = ''.join(ch for ch in ref if ch.isalpha())
                    tipo = c.get('t', 'n')
                    v = c.find('x:v', NS)
                    if v is not None and v.text is not None:
                        if tipo == 's':
                            try: valor = shared[int(v.text)]
                            except: valor = v.text
                        else:
                            try: valor = int(v.text) if '.' not in v.text else float(v.text)
                            except: valor = v.text
                        celdas[col] = valor
                if celdas:
                    filas[rnum] = celdas
            return filas
    return {}

dol = {}
sol = {}
with zipfile.ZipFile('Asiento_ITF_DOL_1.xlsm') as z:
    dol = leer_hoja(z, 'CONTABILIDAD')
with zipfile.ZipFile('Asiento_ITF_SOL_1.xlsm') as z:
    sol = leer_hoja(z, 'CONTABILIDAD')

print(f'DOL filas: {len(dol)}, SOL filas: {len(sol)}')
print()

# Mostrar primeras 4 filas de cada uno con TODAS las columnas
print('=== DOL primeras 4 filas con datos ===')
for i, (fnum, row) in enumerate(sorted(dol.items())):
    print(f'  fila {fnum}: {dict(sorted(row.items()))}')
    if i >= 3: break

print()
print('=== SOL primeras 4 filas con datos ===')
for i, (fnum, row) in enumerate(sorted(sol.items())):
    print(f'  fila {fnum}: {dict(sorted(row.items()))}')
    if i >= 3: break

# Comparar columnas presentes
cols_dol = set()
cols_sol = set()
for row in dol.values():
    cols_dol.update(row.keys())
for row in sol.values():
    cols_sol.update(row.keys())

print()
print('Columnas en DOL:', sorted(cols_dol))
print('Columnas en SOL:', sorted(cols_sol))
print('Solo en DOL:', sorted(cols_dol - cols_sol))
print('Solo en SOL:', sorted(cols_sol - cols_dol))

# Ver todos los valores únicos por columna en las primeras 10 filas de datos
print()
print('=== Valores columna D (Correlativo) ===')
print('DOL:', [dol[f].get('D') for f in sorted(dol.keys())[:10]])
print('SOL:', [sol[f].get('F') for f in sorted(sol.keys())[:10]])

# Buscar qué columna es Correlativo en SOL (puede ser distinta)
print()
print('=== Todas las columnas fila 9 DOL:', dict(sorted(dol.get(9, {}).items())))
print('=== Todas las columnas fila 9 SOL:', dict(sorted(sol.get(9, {}).items())))
