# -*- coding: utf-8 -*-
"""Traza exacta de aplicar_ajustes."""
import sys
from pathlib import Path

BASE = Path(r'C:\Users\darobles\Documents\PROYECTOS\CONTROL INVENTARIO MANTENIMIENTOS 3')
sys.path.insert(0, str(BASE))
import ingesta  # noqa: E402
import pandas as pd  # noqa: E402

p = BASE / 'Ingesta de datos diaria' / 'Campos dashborad 16 sept.xlsx'
d = pd.read_excel(p, sheet_name='Dashboard_KPI')
print('type(d):', type(d), '| shape:', d.shape)
print('d.columns is Index:', isinstance(d.columns, pd.Index), '| len:', len(d.columns))
print("'Estado de la sede' in d.columns ->", 'Estado de la sede' in d.columns)
print("normalizar_nombre('Estado de la sede') ->", repr(ingesta.normalizar_nombre('Estado de la sede')))
print("claves reales (ultimas 6):", list({ingesta.normalizar_nombre(c): c for c in d.columns}.keys())[-6:])

ajustes = dict(ingesta.ANALISIS_COLUMNAS["campos"])
print('\nclaves de ajustes:', list(ajustes.keys()))
for nombre, regla in ajustes.items():
    encontrada = ingesta.buscar_columna(d, nombre)
    print(f'  buscar_columna({nombre!r}) -> {encontrada!r}')
    if encontrada is None:
        print('     normalizado ->', repr(ingesta.normalizar_nombre(nombre)))
