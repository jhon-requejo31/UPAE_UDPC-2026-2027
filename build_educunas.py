#!/usr/bin/env python3
"""Genera educunas.js desde EDUCUNAS_2026.xlsx (hoja 'EDUCUNAS 2026').
Uso: python build_educunas.py EDUCUNAS_2026.xlsx educunas.js
Normaliza respuestas (Si/SI/NO/0/NO APLICA...) y guarda las columnas en formato columnar compacto.
Codificación Sí/No: 1 = Sí, 0 = No, 2 = No aplica, 3 = Sin dato."""
import sys, json, re, unicodedata, pandas as pd

d = pd.read_excel(sys.argv[1], sheet_name='EDUCUNAS 2026', dtype=str).fillna('')
d = d.apply(lambda s: s.str.strip())


def nz(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').lower().strip()


def yn(x):
    k = nz(x)
    if k in ('si', '1'): return 1
    if k == 'no': return 0
    if 'no aplica' in k or k == 'no usa': return 2
    return 3  # '0' u otro valor no interpretable


def est(x):
    k = nz(x)
    if k in ('adecuado', 'apto', 'bueno', 'optimo'): return 'Adecuado'
    if k in ('regular', 'aceptable', 'condicionado'): return 'Regular'
    if k in ('critico', 'malo', 'defectuoso', 'inoperativo'): return 'Crítico'
    if 'no aplica' in k: return 'No aplica'
    return 'Sin dato'


def via(x):
    k = nz(x)
    if 'no pavimentada' in k: return 'Vía no pavimentada'
    if 'pavimentada' in k: return 'Vía pavimentada'
    if 'mixta' in k: return 'Mixta'
    if 'trocha' in k: return 'Trocha carrozable'
    if 'peatonal' in k: return 'Camino peatonal'
    return 'Sin dato'


def comb(x):
    k = nz(x)
    if 'balon' in k: return 'Balón GLP'
    if 'gnv' in k: return 'Red GNV'
    if 'red glp' in k: return 'Red GLP'
    if 'electric' in k: return 'Eléctrica'
    return 'No aplica / no usa'


def alm(x):
    k = nz(x)
    if k in ('0', ''): return 'Sin dato'
    if k.startswith('estantes'): return 'Estantes, anaqueles o mesas'
    if k == 'mesas': return 'Mesas'
    if k.startswith('en materiales'): return 'Materiales aislantes (estibas / recipientes)'
    if 'parihuela' in k: return 'Parihuelas (pallets)'
    if 'no cuenta' in k: return 'No cuenta'
    return 'Otros'


def eq(x):
    k = nz(x)
    if k == 'optimo': return 'Óptimo'
    if k == 'aceptable': return 'Aceptable'
    if k in ('malo', 'defectuoso', 'inoperativo'): return 'Malo / inoperativo'
    return 'Sin dato'


def forma(x):
    if 'CONCES' in x.upper(): return 'CONCESIÓN'
    return x.upper() if x not in ('', '-') else 'SIN FORMA (LISTO PARA CONSUMO)'


YN = ['cuenta_con_cocina', 'cocina_ventilada', 'cocina_distante_sshh', 'cuenta_fuente_agua', 'cuenta_lavadero',
      'ventanas_mosquiteras_cocina', 'proteccion_acceso_cocina', 'cuenta_electricidad', 'cuenta_artefacto_cocina',
      'cuenta_mesa_trabajo', 'cuenta_ollas', 'cuenta_utensilios_cocina', 'cuenta_con_almacen', 'almacen_distante_sshh',
      'ventanas_mosquiteras_almacen', 'proteccion_acceso_almacen', 'cuenta_equipamiento_almacen', 'cuenta_sistema_frio']
ES = ['estado_paredes_cocina', 'estado_techo_cocina', 'estado_piso_cocina',
      'estado_paredes_almacen', 'estado_techo_almacen', 'estado_piso_almacen']
o = {k: [] for k in ['ut', 'dep', 'prov', 'dis', 'cm', 'nom', 'comite', 'item', 'mod', 'forma', 'area', 'via', 'sec',
                     'ben', 'comb', 'alm', 'ref', 'mic', 'com'] + YN + ES}
for _, r in d.iterrows():
    o['ut'].append(r.unidad_territorial.upper()); o['dep'].append(r.departamento.upper())
    o['prov'].append(r.provincia.upper()); o['dis'].append(r.distrito.upper())
    o['cm'].append(r.cod_modular); o['nom'].append(r.nombre_ie); o['comite'].append(r.propuesta_comite_2027.upper())
    o['item'].append(re.sub(r'\s+', ' ', r.propuesta_item_2027.replace('\\n', ' ')).strip().upper())
    o['mod'].append(r.propuesta_modalidad_2027.upper()); o['forma'].append(forma(r.propuesta_forma_atencion_2027))
    o['area'].append(r.area); o['via'].append(via(r.via_acceso_ie))
    o['sec'].append(int(r.n_secciones or 0)); o['ben'].append(int(r.beneficiarios or 0))
    o['comb'].append(comb(r.tipo_combustible_cocina)); o['alm'].append(alm(r.tipo_almacenamiento))
    o['ref'].append(eq(r.estado_refrigeradora)); o['mic'].append(eq(r.estado_microondas))
    o['com'].append(1 if len(r.comentarios_sugerencias.strip()) > 2 else 0)
    for c in YN: o[c].append(yn(r[c]))
    for c in ES: o[c].append(est(r[c]))
open(sys.argv[2], 'w', encoding='utf8').write(
    'window.EDUCUNAS=' + json.dumps({'n': len(d), 'year': 2026, 'c': o}, ensure_ascii=False, separators=(',', ':')) + ';')
print(len(d), 'filas')
