# -*- coding: utf-8 -*-
"""Diagnostico: por que 'Estado de la sede' y 'ESTADO UPS' no se encuentran."""
import sys
from pathlib import Path

BASE = Path(r'C:\Users\darobles\Documents\PROYECTOS\CONTROL INVENTARIO MANTENIMIENTOS 3')
sys.path.insert(0, str(BASE))
import ingesta  # noqa: E402
import pandas as pd  # noqa: E402

p = BASE / 'Ingesta de datos diaria' / 'Campos dashborad 16 sept.xlsx'
d = pd.read_excel(p, sheet_name='Dashboard_KPI')
print('columnas 40..43:', [repr(c) for c in d.columns[40:44]])
print('buscar_columna(Estado de la sede) ->', repr(ingesta.buscar_columna(d, 'Estado de la sede')))
print('buscar_columna(ESTADO UPS)       ->', repr(ingesta.buscar_columna(d, 'ESTADO UPS')))
print()
print('--- ajustes de campos ---')
for nombre, regla in ingesta.ANALISIS_COLUMNAS["campos"].items():
    print(f'  {nombre!r:<26} -> {ingesta.buscar_columna(d, nombre)!r}')
print()
print('--- que pasa dentro de aplicar_ajustes ---')
r = ingesta.aplicar_ajustes(d.copy(), ingesta.ANALISIS_COLUMNAS["campos"])
print('tras ajustar, ESTADO UPS =', r['ESTADO UPS'].value_counts(dropna=False).to_dict())
print()
print('--- orden de pasos en preparar_hoja ---')
import inspect
print(inspect.getsource(ingesta.preparar_hoja))
