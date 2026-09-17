# -*- coding: utf-8 -*-
"""Instrumenta preparar_hoja para ver con que ajustes y columnas trabaja."""
import sys
from pathlib import Path

BASE = Path(r'C:\Users\darobles\Documents\PROYECTOS\CONTROL INVENTARIO MANTENIMIENTOS 3')
sys.path.insert(0, str(BASE))
import ingesta  # noqa: E402
import pandas as pd  # noqa: E402

_original = ingesta.aplicar_ajustes


def con_traza(df, ajustes):
    print(f'>>> aplicar_ajustes: {df.shape} | claves ajustes = {list((ajustes or {}).keys())}')
    print(f'    columnas del df: {list(df.columns)[:6]} ... total {len(df.columns)}')
    for nombre in (ajustes or {}):
        print(f'    buscar({nombre!r}) = {ingesta.buscar_columna(df, nombre)!r}')
    return _original(df, ajustes)


ingesta.aplicar_ajustes = con_traza

p = BASE / 'Ingesta de datos diaria' / 'Campos dashborad 16 sept.xlsx'
hojas = ingesta.leer_y_preparar(p, 'campos')
print('RESULTADO:', {k: v.shape for k, v in hojas.items()})
