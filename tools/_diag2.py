# -*- coding: utf-8 -*-
"""Reproduce exactamente leer_y_preparar para ver de donde salen los avisos."""
import sys
from pathlib import Path

BASE = Path(r'C:\Users\darobles\Documents\PROYECTOS\CONTROL INVENTARIO MANTENIMIENTOS 3')
sys.path.insert(0, str(BASE))
import ingesta  # noqa: E402

p = BASE / 'Ingesta de datos diaria' / 'Campos dashborad 16 sept.xlsx'
print('=== leer_y_preparar(campos) ===')
hojas = ingesta.leer_y_preparar(p, 'campos')
for hoja, df in hojas.items():
    print(f'  -> {hoja}: {df.shape} | cols={list(df.columns)}')

print()
print('=== tipo_de_archivo ===')
for q in sorted((BASE / 'Ingesta de datos diaria').glob('*.xlsx')):
    print(f'  {q.name:<40} -> {ingesta.tipo_de_archivo(q)!r}')
