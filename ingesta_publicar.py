# -*- coding: utf-8 -*-
"""
Ingesta diaria -> publica los archivos en data/ y actualiza la fecha de última
actualización. Opcionalmente hace commit y push al repositorio de GitHub.

Uso:
    python ingesta_publicar.py            # copia + metadatos + git push
    python ingesta_publicar.py --no-push  # solo copia + metadatos
"""
import hashlib
import json
import shutil
import subprocess
import sys
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent
ORIGEN = BASE.parent / "Ingesta de datos diaria"
DATA = BASE / "data"
DATA.mkdir(exist_ok=True)

DESTINOS = {"data": "Data_actual.xlsx",
            "campos": "Campos_actual.xlsx",
            "crono": "Cronograma_actual.xlsx",
            "crono_ups": "Cronograma_UPS_actual.xlsx"}


def _tipo(path: Path) -> str:
    try:
        xl = pd.ExcelFile(path)
    except Exception:  # noqa: BLE001
        return ""
    if "Dashboard_KPI" in xl.sheet_names:
        return "campos"
    if "UPS" in xl.sheet_names or "UPS" in path.name.upper():
        return "crono_ups"
    try:
        cols = set(map(str, xl.parse(xl.sheet_names[0], nrows=0).columns))
    except Exception:  # noqa: BLE001
        return ""
    if {"Consecutivo mantenimiento 3", "Facturable"} <= cols:
        return "data"
    if {"ESTADO", "Fecha Inicio", "Fecha Fin", "tipo"} <= cols:
        return "crono"
    if "Consecutivo mantenimiento 3" in cols:
        return "data"
    return ""


def _elegir() -> dict:
    elegidos = {}
    if not ORIGEN.exists():
        return elegidos
    for p in sorted(ORIGEN.glob("*.xlsx"), key=lambda x: x.stat().st_mtime, reverse=True):
        t = _tipo(p)
        if t and t not in elegidos:
            elegidos[t] = p
    return elegidos


def _git(args) -> int:
    return subprocess.call(["git"] + args, cwd=str(BASE))


def main(push: bool = True) -> int:
    tipos = _elegir()
    if "data" not in tipos:
        print("[ERROR] No encontré una Data válida (hoja Hoja1) en:", ORIGEN)
        return 1
    info = {}
    for t, destino in DESTINOS.items():
        if t not in tipos:
            continue
        origen = Path(tipos[t])
        destino_path = DATA / destino
        shutil.copy2(origen, destino_path)
        sha = hashlib.sha256(destino_path.read_bytes()).hexdigest()
        info[t] = {"archivo_origen": origen.name, "archivo": destino,
                   "sha256": sha, "bytes": destino_path.stat().st_size}
        print(f"OK {t}: {origen.name} -> {destino}")

    meta = {"fecha_hora": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
            "archivos": info}
    try:
        meta["filas_data"] = int(len(pd.read_excel(DATA / "Data_actual.xlsx",
                                                   sheet_name="Hoja1")))
    except Exception:  # noqa: BLE001
        pass
    try:
        if (DATA / "Campos_actual.xlsx").exists():
            meta["filas_novedades"] = int(len(pd.read_excel(DATA / "Campos_actual.xlsx",
                                                            sheet_name="Novedades Equipos")))
    except Exception:  # noqa: BLE001
        pass
    (DATA / "ultima_actualizacion.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print("Última actualización:", meta["fecha_hora"])

    if not push:
        return 0
    _git(["add", "-A"])
    _git(["commit", "-m", f"Ingesta diaria {meta['fecha_hora']}"])
    rc = _git(["push"])
    if rc != 0:
        print("\n[AVISO] No se pudo hacer push. Verifique primero:")
        print("  git remote add origin https://github.com/dannrob23/Panel-de-Control-Mantenimientos-3.git")
        print("  git branch -M main")
        print("  git push -u origin main")
    else:
        print("Push realizado. Streamlit Cloud se redepliega automáticamente.")
    return rc


if __name__ == "__main__":
    sys.exit(main(push="--no-push" not in sys.argv))
