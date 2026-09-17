# -*- coding: utf-8 -*-
"""Inspecciona hojas/columnas de los dos archivos de la ingesta (previo al script final)."""
import sys
from pathlib import Path

import pandas as pd

ING = Path(r'C:\Users\darobles\Documents\PROYECTOS\CONTROL INVENTARIO MANTENIMIENTOS 3\Ingesta de datos diaria')
for p in sorted(ING.glob('*.xlsx'), key=lambda x: x.stat().st_mtime, reverse=True):
    if p.name.startswith('~$'):
        continue
    print('=' * 90)
    print('ARCHIVO:', p.name, '|', round(p.stat().st_size / 1e6, 1), 'MB')
    xl = pd.ExcelFile(p)
    for hoja in xl.sheet_names:
        d = xl.parse(hoja, nrows=200000)
        print(f'  HOJA "{hoja}": {len(d)} filas x {len(d.columns)} columnas')
        for i, c in enumerate(d.columns):
            nn = d[c].notna().sum()
            muestra = [str(v)[:26] for v in d[c].dropna().unique()[:3]]
            print(f'     [{i:>2}] {str(c)[:38]:<38} no-nulos={nn:<7} ej={muestra}')
