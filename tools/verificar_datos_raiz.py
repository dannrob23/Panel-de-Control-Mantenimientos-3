# -*- coding: utf-8 -*-
"""VERIFICADOR DE DATOS del panel MT3  (no publica nada, solo lee y reporta).

Uso:   python tools\\verificar_datos.py

Revisa los Excel vigentes (deploy_panel\\data, o data\\ si aun no se ha publicado) y muestra:
  · los totales de control de las tarjetas de oficinas: sedes, finalizadas y UPS finalizadas;
  · el consolidado por jefatura EXACTAMENTE como lo calcula el panel (misma funcion).
Al final dice si los totales cuadran o si hay que revisar algo.

Sirve para validar los archivos del dia SIN abrir el panel y SIN gastar asistencia:
si el consolidado y los totales coinciden con el archivo, los datos estan bien.
"""
import os
import sys
import logging
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))
os.environ["INGESTA_SIN_APP"] = "1"  # importar app.py sin arrancar Streamlit
# app.py importa Streamlit, que escupe avisos de "no runtime" en modo consola.
logging.getLogger("streamlit").setLevel(logging.CRITICAL)
for _f in (sys.stdout, sys.stderr):
    try:
        _f.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import contextlib  # noqa: E402
import io  # noqa: E402
import pandas as pd  # noqa: E402

# `app.py` importa Streamlit, que en modo consola escupe avisos de "no runtime".
# Se silencian durante el import para que la salida quede limpia y legible.
_buf = io.StringIO()
with contextlib.redirect_stderr(_buf):
    import app  # noqa: E402


def buscar(nombre: str):
    """Primero lo que ve la nube (deploy_panel/data); si no, la copia local."""
    for base in ("deploy_panel/data", "data"):
        p = RAIZ / base / nombre
        if p.exists():
            return p
    return None


def main() -> int:
    print("=" * 88)
    print(" VERIFICADOR DE DATOS - PANEL MT3")
    print("=" * 88)

    p = buscar("Campos_actual.xlsx")
    if p is None:
        print("\n[ERROR] No encuentro Campos_actual.xlsx en data/ ni en deploy_panel/data/.")
        print("        Ejecuta la ingesta primero (actualizar_panel.bat).")
        return 1
    meta = buscar("ultima_actualizacion.json")
    print(f"\n Archivo  : {p.relative_to(RAIZ)}")
    if meta is not None:
        import json
        try:
            info = json.loads(meta.read_text(encoding="utf-8"))
            print(f" Actualizado: {info.get('fecha_hora', '?')}  "
                  f"(data: {info.get('filas_data', '?')} filas)")
        except Exception:  # noqa: BLE001
            pass

    kpi = pd.read_excel(p, sheet_name="Dashboard_KPI")
    faltan = [c for c in ("Estado de la sede", "ESTADO UPS",
                          "Jefaturas  Operaciones Regional") if c not in kpi.columns]
    if faltan:
        print(f"\n[ERROR] Al archivo le faltan columnas: {faltan}")
        print("        Vuelve a generar los datos con actualizar_panel.bat.")
        return 1

    # --- Totales "tal cual lo reporta la oficina" (control directo del archivo) ---
    est = kpi["Estado de la sede"].fillna("").astype(str).str.strip()
    fin = int(est.eq("Finalizada").sum())
    rfin = int(est.eq("Reprogramada_Finalizada").sum())
    upsf = int(kpi["ESTADO UPS"].fillna("").astype(str).str.strip()
               .str.lower().str.startswith("finaliz").sum())
    print(f"\n TOTALES DEL ARCHIVO (lo que reportan las oficinas)")
    print(f"   Sedes (filas)                : {len(kpi)}")
    print(f"   Finalizadas                  : {fin}")
    print(f"   Reprogramada_Finalizada      : {rfin}")
    print(f"   Total finalizadas            : {fin + rfin}")
    print(f"   UPS finalizadas (col AP)     : {upsf}")

    # --- Consolidado por jefatura (misma funcion que usa el panel) ---
    t = app.tabla_jefaturas_campos(kpi)
    print(f"\n CONSOLIDADO POR JEFATURA (como lo calcula el panel)")
    print(t.to_string(index=False))

    # --- Comprobacion: el consolidado debe cuadrar con el archivo ---
    ok = (len(t) > 0
          and int(t["Sedes"].sum()) == len(kpi)
          and int(t["Finalizadas"].sum()) == fin
          and int(t["Reprog. finalizadas"].sum()) == rfin
          and int(t["UPS finalizadas"].sum()) == upsf)
    print("\n" + "-" * 88)
    if ok:
        print(f" [OK] El consolidado cuadra con el archivo: {len(t)} jefaturas · "
              f"{len(kpi)} sedes · {fin + rfin} finalizadas · {upsf} UPS finalizadas.")
    else:
        print(" [REVISAR] El consolidado NO cuadra con el archivo. Revisa las columnas "
              "«Jefaturas  Operaciones Regional» y «Estado de la sede».")
    print("-" * 88)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
