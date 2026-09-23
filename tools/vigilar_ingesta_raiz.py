# -*- coding: utf-8 -*-
"""
=============================================================================
 ACTUALIZACION AUTOMATICA DEL PANEL MT3
=============================================================================
Vigila la carpeta "Ingesta de datos diaria". Cuando guardas los documentos
(Data + Campos dashboard), espera a que el archivo termine de copiarse,
actualiza el dashboard y hace el commit + push a GitHub. Todo solo.

COMO SE USA:

    Doble clic en `actualizacion_automatica.bat`   (uso normal: se queda vigilando)
    python vigilar_ingesta.py                     (igual, desde la terminal)
    python vigilar_ingesta.py --una-vez           (revisa UNA vez y termina)
    python vigilar_ingesta.py --sin-publicar      (actualiza pero NO sube a GitHub)

Deja la ventana abierta y olvídate: cada vez que guardes los Excel en la
carpeta, el panel se actualiza y se publica solo.

NO duplica lógica: usa `ingesta.py`, que es el que ajusta columnas, valida,
anonimiza, escribe `data/` y `deploy_panel/data`, y hace el push.
=============================================================================
"""
from __future__ import annotations

import argparse
import hashlib
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

# =============================================================================
# 1) CONFIGURACION  <<<<<<  EDITA SOLO ESTA SECCION
# =============================================================================
BASE = Path(__file__).resolve().parent
CARPETA = BASE / "Ingesta de datos diaria"   # carpeta que se vigila
PUBLICAR = True                              # True = además hace commit + push

CADA_SEGUNDOS = 5        # cada cuánto revisa la carpeta
ESPERA_SEGUNDOS = 25     # espera a que el archivo deje de cambiar (copias grandes)
MAX_ESPERA = 900         # si tras 15 min sigue cambiando, avisa y no publica
LOG = BASE / "tools" / "ingesta_auto.log"

# La consola de Windows usa cp1252: se fuerza UTF-8 para no caer por un carácter.
for _flujo in (sys.stdout, sys.stderr):
    try:
        _flujo.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # noqa: BLE001
        pass


# =============================================================================
# 2) UTILIDADES
# =============================================================================
def ahora() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(mensaje: str, archivo: bool = True):
    """Imprime en pantalla y (opcionalmente) agrega la línea al archivo de log."""
    linea = f"[{ahora()}] {mensaje}"
    print(linea, flush=True)
    if not archivo:
        return
    try:
        LOG.parent.mkdir(parents=True, exist_ok=True)
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(linea + "\n")
    except Exception:  # noqa: BLE001
        pass


def xlsx_validos(carpeta: Path) -> list:
    """Los .xlsx de la carpeta, sin los temporales que deja Excel abierto (~$)."""
    if not carpeta.exists():
        return []
    return sorted((p for p in carpeta.glob("*.xlsx") if not p.name.startswith("~$")),
                  key=lambda p: p.name)


def huella(archivos: list) -> str:
    """Huella de la carpeta: nombre, tamaño y fecha de cada archivo.

    Sirve para no volver a publicar si nada cambió, y para detectar si el
    archivo todavía se está copiando (el tamaño/fecha cambia).
    """
    partes = []
    for p in archivos:
        try:
            st = p.stat()
            partes.append(f"{p.name}|{st.st_size}|{int(st.st_mtime)}")
        except OSError:
            partes.append(f"{p.name}|?|?")
    return hashlib.sha256("::".join(partes).encode("utf-8")).hexdigest()[:16]


def clasificar(archivos: list) -> tuple:
    """Separa en reconocibles y no reconocidos, reutilizando la lógica de ingesta.py."""
    sys.path.insert(0, str(BASE))
    import ingesta  # noqa: PLC0415
    reconocidos, otros = {}, []
    for p in archivos:
        tipo = ingesta.tipo_de_archivo(p)
        if tipo:
            reconocidos.setdefault(tipo, p)
        else:
            otros.append(p.name)
    return reconocidos, otros


def esperar_archivos_quietos(archivos: list) -> bool:
    """Espera a que los archivos dejen de cambiar (terminó la copia). True si están listos."""
    inicio = time.time()
    anterior = None
    estable_desde = None
    while True:
        actual = huella(archivos)
        if actual != anterior:
            anterior = actual
            estable_desde = time.time()
            log(f"   … los archivos están cambiando todavía (espero a que terminen)")
        elif estable_desde and (time.time() - estable_desde) >= ESPERA_SEGUNDOS:
            return True
        if (time.time() - inicio) > MAX_ESPERA:
            log(f"   ⚠️ Los archivos siguen cambiando tras {MAX_ESPERA//60} min. "
                "No publico para no subir algo a medias.")
            return False
        time.sleep(CADA_SEGUNDOS)


def ejecutar_ingesta() -> int:
    """Llama a ingesta.py (el publicador real) y devuelve su código de salida."""
    cmd = [sys.executable, str(BASE / "ingesta.py")]
    cmd.append("--publicar" if PUBLICAR else "--seco")
    log(f"   ▶ Ejecutando: python ingesta.py {'--publicar' if PUBLICAR else '--seco'}")
    try:
        r = subprocess.run(cmd, cwd=str(BASE))
        return int(r.returncode)
    except Exception as exc:  # noqa: BLE001
        log(f"   ❌ No pude ejecutar la ingesta: {exc}")
        return 1


SECUENCIA = 0     # contador de ciclos, para el latido de estado
LATIDO = 60       # cada cuántos segundos avisa que sigue vivo (0 = sin latido)


# =============================================================================
# 3) VIGILANCIA
# =============================================================================
def main() -> int:
    ap = argparse.ArgumentParser(description="Actualiza y publica el panel solo")
    ap.add_argument("--una-vez", action="store_true",
                    help="revisa una vez y termina (no se queda vigilando)")
    ap.add_argument("--sin-publicar", action="store_true",
                    help="actualiza los datos pero no sube a GitHub")
    args = ap.parse_args()
    global PUBLICAR
    if args.sin_publicar:
        PUBLICAR = False

    print("=" * 74)
    print(" ACTUALIZACION AUTOMATICA DEL PANEL MT3")
    print("=" * 74)
    log(f" Carpeta vigilada : {CARPETA}")
    log(f" Publicar en GitHub: {'SÍ' if PUBLICAR else 'NO (solo actualiza)'}")
    log(f" Revisa cada      : {CADA_SEGUNDOS} s · espera {ESPERA_SEGUNDOS} s de calma")
    if not CARPETA.exists():
        log(f" ❌ No existe la carpeta: {CARPETA}")
        return 2

    # --- Revisión inicial: ¿ya hay archivos nuevos sin publicar? ---
    archivos = xlsx_validos(CARPETA)
    if not archivos:
        log(" ⚠️ La carpeta está vacía. Guarda ahí los Excel (Data y Campos dashboard).")
    else:
        reconocidos, otros = clasificar(archivos)
        log(f" Archivos en la carpeta: {', '.join(reconocidos.get(t, Path('?')).name for t in reconocidos)}")
        if otros:
            log(f" (ignorados por no ser de la ingesta: {', '.join(otros)})")
        if not reconocidos.get("data"):
            log(" ❌ Falta la Data de ejecución (hoja Hoja1). No publico todavía.")
        else:
            log(" ▶ Hay archivos para procesar. Actualizando…")
            if esperar_archivos_quietos(archivos):
                codigo = ejecutar_ingesta()
                if codigo == 0:
                    log(" ✅ Listo: dashboard actualizado" +
                        (" y publicado en GitHub." if PUBLICAR else "."))
                else:
                    log(f" ❌ La ingesta terminó con código {codigo}. Revisa los mensajes de arriba.")

    if args.una_vez:
        log(" Revisión única terminada.")
        return 0

    # --- Bucle de vigilancia ---
    log(" 👀 Vigilando… (deja esta ventana abierta; Ctrl+C para salir)")
    ultima = huella(xlsx_validos(CARPETA))
    t_inicial = time.time()
    while True:
        try:
            time.sleep(CADA_SEGUNDOS)
            archivos = xlsx_validos(CARPETA)
            actual = huella(archivos)
            if actual == ultima:
                if LATIDO and (time.time() - t_inicial) >= LATIDO:
                    log(f" 💤 Sin cambios · {len(archivos)} archivo(s) en la carpeta · "
                        "sigo vigilando")
                    t_inicial = time.time()
                continue
            log("─" * 70)
            log(" 📥 Detecté un cambio en la carpeta")
            reconocidos, otros = clasificar(archivos)
            for t, p in reconocidos.items():
                log(f"   · {t:<10} {p.name}")
            if otros:
                log(f"   (ignorados: {', '.join(otros)})")
            if not reconocidos.get("data"):
                log("   ⚠️ Aún no hay Data válida; espero a que termine de copiarse.")
                continue
            ultima = actual
            t_inicial = time.time()
            if esperar_archivos_quietos(archivos):
                # Se recalcula: pudo cambiar mientras esperaba.
                ultima = huella(xlsx_validos(CARPETA))
                codigo = ejecutar_ingesta()
                if codigo == 0:
                    log(" ✅ Listo: dashboard actualizado" +
                        (" y publicado en GitHub." if PUBLICAR else "."))
                else:
                    log(f" ❌ La ingesta terminó con código {codigo}. Revisa los mensajes de arriba.")
                t_inicial = time.time()
        except KeyboardInterrupt:
            log(" 👋 Vigilancia detenida por el usuario.")
            return 0
        except Exception as exc:  # noqa: BLE001
            log(f" ⚠️ Error inesperado en la vigilancia: {exc}")
            time.sleep(CADA_SEGUNDOS)


if __name__ == "__main__":
    sys.exit(main())
