# -*- coding: utf-8 -*-
"""Despliega en deploy_panel/ la app y la documentación de la raíz del proyecto.

Se ejecuta desde deploy_panel/ingesta_publicar.bat antes de publicar, para que el
panel de Streamlit Cloud reciba siempre la misma versión de `app.py` que se prueba
en local. Normaliza los finales de línea a LF (evita diffs falsos en git).
"""
import shutil
import sys
from pathlib import Path

# La consola de Windows usa cp1252: se fuerza UTF-8 para no caer por un carácter.
for _flujo in (sys.stdout, sys.stderr):
    try:
        _flujo.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # noqa: BLE001
        pass

BASE = Path(__file__).resolve().parent
ORIGEN = BASE.parent

ARCHIVOS = {
    "app.py": "app.py",
    "requirements.txt": "requirements.txt",
    "README.md": "README.md",
    "TUTORIAL.md": "TUTORIAL.md",
}

# Scripts de trabajo que viven en la raiz del proyecto. La raiz NO es un repositorio
# git (el repo es esta carpeta), asi que se copian aqui para quedar respaldados en
# GitHub: si se pierden en el PC, se recuperan del repositorio.
# Se guardan en tools/ y con el sufijo "_raiz" para que se vea que se editan ALLA.
HERRAMIENTAS = {
    "ingesta.py": "tools/ingesta_raiz.py",
    "ingesta.bat": "tools/ingesta_raiz.bat",
    "vigilar_ingesta.py": "tools/vigilar_ingesta_raiz.py",
    "actualizacion_automatica.bat": "tools/actualizacion_automatica_raiz.bat",
}


def main() -> int:
    cambios = 0
    for origen_nombre, destino_nombre in {**ARCHIVOS, **HERRAMIENTAS}.items():
        origen = ORIGEN / origen_nombre
        destino = BASE / destino_nombre
        if not origen.exists():
            print(f"[AVISO] No existe {origen}; se omite.")
            continue
        texto = origen.read_text(encoding="utf-8").replace("\r\n", "\n")
        actual = (destino.read_text(encoding="utf-8").replace("\r\n", "\n")
                  if destino.exists() else None)
        if actual == texto:
            print(f"= {destino_nombre}: ya estaba al día")
            continue
        respaldo = None
        if destino.exists():
            respaldo = destino.with_suffix(destino.suffix + ".bak")
            shutil.copy2(destino, respaldo)
        destino.parent.mkdir(parents=True, exist_ok=True)
        destino.write_text(texto, encoding="utf-8", newline="\n")
        cambios += 1
        print(f"[OK] {origen_nombre} -> deploy_panel/{destino_nombre}"
              + (f" (respaldo: {respaldo.name})" if respaldo else ""))
    print(f"Sincronización terminada: {cambios} archivo(s) actualizado(s).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
