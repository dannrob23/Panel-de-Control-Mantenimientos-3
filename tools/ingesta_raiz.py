# -*- coding: utf-8 -*-
"""
=============================================================================
 INGESTA DIARIA - PANEL MT3 (Banco Agrario / COLSOF)
=============================================================================
Un solo script para: leer los Excel del dia -> ajustar columnas -> anonimizar
(opcional) -> publicar en data/ -> hacer commit y push a GitHub.

COMO SE USA (doble clic en `ingesta.bat` o desde la terminal):

    python ingesta.py                 Copia, valida y muestra el resumen (NO publica)
    python ingesta.py --publicar      Todo lo anterior + commit y push a GitHub
    python ingesta.py --analisis      Solo analiza las columnas (no escribe nada)
    python ingesta.py --seco          Igual que --publicar pero sin subir a GitHub

TODO LO EDITABLE ESTA EN LA SECCION "1) CONFIGURACION" (mas abajo).
Para analizar una columna nueva: ejecuta `--analisis`, mira el reporte
(tools\\analisis_ingesta.md) y agrega el ajuste en ANALISIS_COLUMNAS.
=============================================================================
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import unicodedata
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

# La consola de Windows usa cp1252 y revienta con flechas/emojis en los print().
# Se fuerza UTF-8 (con reemplazo) para que la ingesta nunca se caiga por un carácter.
for _flujo in (sys.stdout, sys.stderr):
    try:
        _flujo.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # noqa: BLE001
        pass

# ----------------------------------------------------------------------------
# Zona horaria oficial: América/Bogotá (UTC-5) - Bogotá, Lima, Quito
# ----------------------------------------------------------------------------
try:
    import zoneinfo
    TZ_BO = zoneinfo.ZoneInfo("America/Bogota")
except Exception:
    TZ_BO = timezone.utc

def ahora() -> datetime:
    """datetime actual en zona horaria de Bogotá (UTC-5), naive."""
    return datetime.now(tz=TZ_BO).replace(tzinfo=None)

# =============================================================================
# 1) CONFIGURACION  <<<<<<  EDITA SOLO ESTA SECCION
# =============================================================================
BASE = Path(__file__).resolve().parent                 # carpeta del proyecto
ORIGEN = BASE / "Ingesta de datos diaria"              # (1) de donde se leen los Excel
DATA_LOCAL = BASE / "data"                             # (2) copia local para ver en tu PC
DATA_DEPLOY = BASE / "deploy_panel" / "data"           # (3) datos que ve la web publica
APP_DEPLOY = BASE / "deploy_panel"                     # (4) repositorio git que se publica

ANONIMIZAR = True          # True = oculta datos personales antes de publicar
                           # (seriales -> H_xxx, placas -> ****1234, sin nombres/cedulas)

SHEETS_POR_TIPO = {        # hoja(s) que se conservan de cada archivo, en orden
    "data": ["Hoja1"],                              # Data de ejecucion (equipos)
    "campos": ["Dashboard_KPI", "Novedades Equipos"],  # estado de sede + novedades
    "crono": ["Hoja1"],                             # cronograma de equipos
    "crono_ups": ["UPS"],                           # cronograma de UPS
}

NOMBRES_DESTINO = {        # nombre con el que se publica cada archivo
    "data": "Data_actual.xlsx",
    "campos": "Campos_actual.xlsx",
    "crono": "Cronograma_actual.xlsx",
    "crono_ups": "Cronograma_UPS_actual.xlsx",
}

# Datos personales: red de seguridad. Todo lo que este aqui se elimina SIEMPRE,
# aunque aparezca en COLUMNAS_CONSERVAR (esa lista nunca gana sobre esta).
# Lo normal es que no haga falta tocarla: COLUMNAS_CONSERVAR ya deja fuera lo demas.
COLUMNAS_PERSONALES = {
    "data": ["No. Identificación", "Nombre de usuario", "Usuario de dominio",
             "Correo electrónico", "Cargo", "Dependencia", "Generada por",
             "Hostname", "IP", "MAC"],
    "campos": ["Nonbre Tecnico", "No. Documento",
               "Responsible visto bueno PMU Banco", "Tecnico de Calidad",
               "Observaciones actas PCT", "Observaciones equipos linea base"],
    "crono": [],
    "crono_ups": [],
}

# Columnas que se CONSERVAN de cada hoja (vacio = conservar todas).
# Esta es la lista que manda: lo que NO este aqui se descarta (asi la web carga rapido
# y no se publican datos que el panel no usa).
# ⚠️ NO quites las columnas marcadas con [CLAVE]: el panel las necesita para funcionar.
COLUMNAS_CONSERVAR = {
    "data": ["Serial", "Placa", "Categoría", "Modelo", "SBAN", "Oficina",     # [CLAVE]
             "Regional", "Consecutivo mantenimiento 3",                       # [CLAVE]
             "Fecha de mantenimiento 3", "Facturable"],                       # [CLAVE]
    "campos": {
        "Dashboard_KPI": ["SBAN", "Nombre Oficina", "Fecha Inicio", "Fecha Fin",
                          "Estado de la sede", "ESTADO UPS", "OBSERVACIONES", "UPS"],
        "Novedades Equipos": ["SBAN", "SBAN PCT", "Fecha", "Serial ",         # [CLAVE SBAN,
                              "Oficina", "Estado", "Facturable",              #  Fecha, Categoria]
                              "Nombre Oficina", "Departamento", "ALIADO",
                              "Categoria  Novedad", "Observaciones"],
    },
    "crono": ["SBAN", "tipo", "Nombre Oficina", "Jefaturas  Operaciones Regional",
              "ESTADO", "Fecha Inicio", "Fecha Fin"],
    "crono_ups": ["SBAN", "TIPO", "Nombre Oficina", "Jefaturas  Operaciones Regional",
                  "FECHA", "UPS"],
}

# Ajustes por COLUMNA: quitar espacios, unificar mayusculas, convertir a fecha/numero,
# reemplazar textos escritos de varias formas, etc.
#   tipo: "texto" (quita espacios) · "fecha" · "numero" · "mayus" · "titulo"
#   reemplazar: {valor_que_viene: valor_que_quiero}
# Si el tipo tiene VARIAS hojas, usa "_comunes" (aplica a todas) y el nombre de la hoja
# para lo particular. Los ajustes nunca se aplican a hojas que no los necesitan.
ANALISIS_COLUMNAS = {
    "data": {
        "Oficina": {"tipo": "texto"},
        "Regional": {"tipo": "texto"},
        "Serial": {"tipo": "texto"},
        "Categoría": {"tipo": "texto"},
        "Facturable": {"tipo": "texto",
                       "reemplazar": {"SI": "Si", "NO": "No", "si": "Si", "no": "No"}},
        "Consecutivo mantenimiento 3": {"tipo": "texto"},
        "Fecha de mantenimiento 3": {"tipo": "fecha"},
    },
    "campos": {
        "_comunes": {
            "Nombre Oficina": {"tipo": "texto"},
        },
        "Dashboard_KPI": {
            "Estado de la sede": {"tipo": "texto",
                                  "reemplazar": {"finalizada": "Finalizada",
                                                 "reprogramada_finalizada": "Reprogramada_Finalizada",
                                                 "en proceso": "En proceso",
                                                 "programada": "Programada"}},
            "ESTADO UPS": {"tipo": "texto",
                           "reemplazar": {"finalizada": "Finalizado",
                                          "finalizado": "Finalizado",
                                          "FINALIZADO": "Finalizado",
                                          "en proceso": "En proceso"}},
        },
        "Novedades Equipos": {
            "Categoria  Novedad": {"tipo": "texto"},
            "Fecha": {"tipo": "fecha"},
            "SBAN": {"tipo": "texto"},
        },
    },
}

# Publicacion en GitHub (se usa al elegir la opcion 2 / --publicar)
CONFIG_GIT = {
    "usuario": "dannrob23",                          # autor de los commits
    "correo": "dannrob23@users.noreply.github.com",
    "ssl_backend": "openssl",   # 'openssl' o 'schannel'. En este equipo schannel da
                                # error SEC_E_NO_CREDENTIALS al conectar con GitHub.
    "prompt": False,            # True = permite pedir credenciales por consola
    # Ruta del Git Credential Manager (para reintentar el push sin consola).
    # Dejar "" para no usarlo.
    "credential_manager": r"C:\Users\darobles\AppData\Local\hermes\git\mingw64\bin\git-credential-manager.exe",
}

# Columnas SIN las cuales el panel no funciona: se avisa y se detiene.
COLUMNAS_OBLIGATORIAS = {
    "data": ["Serial", "Placa", "Categoría", "SBAN", "Oficina", "Facturable",
             "Consecutivo mantenimiento 3"],
    "campos": ["SBAN", "Estado de la sede"],
}

# Texto que se escribe en el campo "anonimizado" de ultima_actualizacion.json
ORGANIZACION = "BANCO AGRARIO / COLSOF - Mantenimiento Preventivo 3"

# =============================================================================
# 2) DETECCION DEL TIPO DE ARCHIVO
# =============================================================================
FIRMAS = {
    # tipo buscado -> función que decide si el archivo es de ese tipo
    "campos": lambda hojas, cols, nombre: "Dashboard_KPI" in hojas,
    "crono_ups": lambda hojas, cols, nombre: "UPS" in hojas or "UPS" in nombre,
    "crono": lambda hojas, cols, nombre: {"ESTADO", "Fecha Inicio", "Fecha Fin", "tipo"} <= cols,
    "data": lambda hojas, cols, nombre: {"Consecutivo mantenimiento 3", "Facturable"} <= cols,
}


def sin_acentos(texto: str) -> str:
    """'Categoría' -> 'categoria' (para comparar nombres sin depender de tildes)."""
    base = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in base if not unicodedata.combining(c)).lower().strip()


def normalizar_nombre(valor) -> str:
    """Nombre de columna comparable: sin tildes, sin espacios dobles, en minusculas."""
    return " ".join(sin_acentos(valor).split())


def buscar_columna(df: pd.DataFrame, *nombres):
    """Devuelve el nombre real de una columna tolerando tildes, espacios y mayusculas."""
    if df is None or not len(getattr(df, "columns", [])):
        return None
    reales = {normalizar_nombre(c): c for c in df.columns}
    for nombre in nombres:
        if nombre in df.columns:
            return nombre
        clave = normalizar_nombre(nombre)
        if clave in reales:
            return reales[clave]
    return None


def es_xlsx_util(p: Path) -> bool:
    """Descarta los temporales que Excel deja abiertos (`~$archivo.xlsx`)."""
    return p.suffix.lower() == ".xlsx" and not p.name.startswith("~$")


def hojas_de(path: Path) -> list:
    try:
        return pd.ExcelFile(path).sheet_names
    except Exception:  # noqa: BLE001
        return []


def columnas_de(path: Path, hoja: str) -> set:
    try:
        return set(map(str, pd.read_excel(path, sheet_name=hoja, nrows=0).columns))
    except Exception:  # noqa: BLE001
        return set()


def tipo_de_archivo(path: Path) -> str:
    """Clasifica un .xlsx en 'campos', 'data', 'crono' o 'crono_ups' ('' si no lo reconoce).

    El orden importa: los tipos mas especificos se revisan primero para no confundir
    el cronograma de equipos con la Data (comparten columnas como 'Facturable').
    """
    hojas = hojas_de(path)
    if not hojas:
        return ""
    cols = columnas_de(path, hojas[0])
    nombre = path.name.upper()
    for tipo, firma in FIRMAS.items():
        try:
            if firma(hojas, cols, nombre):
                return tipo
        except Exception:  # noqa: BLE001
            continue
    return ""


def elegir_archivos(origen: Path) -> dict:
    """Toma, por tipo, el .xlsx mas reciente de la carpeta de ingesta."""
    elegidos, ignorados = {}, []
    for p in sorted(origen.glob("*.xlsx"), key=lambda x: x.stat().st_mtime, reverse=True):
        if not es_xlsx_util(p):
            ignorados.append(p.name)
            continue
        tipo = tipo_de_archivo(p)
        if tipo and tipo not in elegidos:
            elegidos[tipo] = p
    if ignorados:
        print(f"  (se ignoran {len(ignorados)} temporal(es) de Excel: {', '.join(ignorados)})")
    return elegidos


# =============================================================================
# 3) LECTURA Y AJUSTE DE COLUMNAS
# =============================================================================
def ajustes_de(tipo: str, hoja: str) -> dict:
    """Devuelve los ajustes que aplican a una hoja concreta.

    Admite dos formatos en ANALISIS_COLUMNAS:
      - plano:            {"columna": {...}}                  (una sola hoja)
      - por hoja:         {"_comunes": {...}, "Hoja1": {...}} (varias hojas)
    """
    config = ANALISIS_COLUMNAS.get(tipo) or {}
    if not config:
        return {}
    if any(isinstance(v, dict) and ("tipo" in v or "reemplazar" in v)
           for v in config.values()):
        return config                                  # formato plano
    comunes = config.get("_comunes") or {}
    particulares = config.get(hoja) or {}
    return {**comunes, **particulares}


def aplicar_ajustes(df: pd.DataFrame, ajustes: dict, etiqueta: str = "") -> pd.DataFrame:
    """Aplica la configuracion de columnas a un DataFrame."""
    for nombre, regla in (ajustes or {}).items():
        col = buscar_columna(df, nombre)
        if col is None:
            print(f"    · aviso: no encontre la columna '{nombre}'"
                  f"{f' en {etiqueta}' if etiqueta else ''} (se omite el ajuste)")
            continue
        tipo = str(regla.get("tipo", "texto")).lower()
        if tipo == "fecha":
            df[col] = pd.to_datetime(df[col], errors="coerce")
        elif tipo == "numero":
            df[col] = pd.to_numeric(df[col], errors="coerce")
        else:
            serie = df[col]
            if serie.dtype == object or str(serie.dtype).startswith("str"):
                serie = serie.astype(str).str.strip()
                serie = serie.mask(serie.isin(["", "nan", "None", "NaT"]))
            if tipo == "mayus":
                serie = serie.str.upper()
            elif tipo == "titulo":
                serie = serie.str.title()
            df[col] = serie
        for viejo, nuevo in (regla.get("reemplazar") or {}).items():
            objetivo = buscar_columna(df, nombre)
            serie = df[objetivo].astype(str).str.strip()
            df.loc[serie.str.lower().eq(str(viejo).strip().lower()), objetivo] = nuevo
    return df


def conservar_columnas(df: pd.DataFrame, tipo: str, hoja: str) -> pd.DataFrame:
    """Deja solo las columnas pedidas en COLUMNAS_CONSERVAR (si hay lista)."""
    pedido = COLUMNAS_CONSERVAR.get(tipo)
    if isinstance(pedido, dict):
        pedido = pedido.get(hoja)
    if not pedido:
        return df
    presentes = [c for c in pedido if c in df.columns]
    faltan = [c for c in pedido if c not in df.columns]
    if faltan:
        print(f"    · aviso: columnas pedidas que no existen en {hoja}: {faltan}")
    return df[presentes] if presentes else df


def quitar_personales(df: pd.DataFrame, tipo: str, hoja: str = "") -> pd.DataFrame:
    """Elimina columnas con datos personales (siempre manda sobre COLUMNAS_CONSERVAR)."""
    blandas = list(COLUMNAS_PERSONALES.get(tipo) or [])
    if isinstance(COLUMNAS_PERSONALES.get(tipo), dict):
        blandas = list((COLUMNAS_PERSONALES[tipo].get(hoja) or []))
    quitar = [c for c in df.columns if any(normalizar_nombre(c) == normalizar_nombre(b)
                                           for b in blandas)]
    if quitar:
        print(f"    · datos personales fuera: {', '.join(map(str, quitar))}")
    return df.drop(columns=quitar)


def anonimizar_valor(valor) -> str:
    texto = str(valor).strip()
    if not texto or texto.lower() in ("nan", "none", "nat"):
        return ""
    return "H_" + hashlib.sha256(texto.encode("utf-8")).hexdigest()[:12]


def enmascarar(valor) -> str:
    texto = str(valor).strip()
    return "****" + texto[-4:] if len(texto) > 4 else texto


def anonimizar(df: pd.DataFrame, tipo: str) -> pd.DataFrame:
    """Serial -> hash · Placa -> ****1234. Solo cuando ANONIMIZAR es True."""
    if not ANONIMIZAR:
        return df
    col_serial = buscar_columna(df, "Serial", "Serial ")
    if col_serial is not None:
        df[col_serial] = df[col_serial].apply(anonimizar_valor)
    col_placa = buscar_columna(df, "Placa")
    if col_placa is not None:
        df[col_placa] = df[col_placa].apply(enmascarar)
    return df


def preparar_hoja(path: Path, tipo: str, hoja: str) -> pd.DataFrame:
    """Lee una hoja y le aplica todo el flujo de ajuste."""
    df = pd.read_excel(path, sheet_name=hoja)
    print(f"    · {hoja}: {len(df):,} filas x {len(df.columns)} columnas".replace(",", "."))
    df = aplicar_ajustes(df, ajustes_de(tipo, hoja), f"{tipo}/{hoja}")
    df = conservar_columnas(df, tipo, hoja)
    df = quitar_personales(df, tipo, hoja)
    df = anonimizar(df, tipo)
    return df


def leer_y_preparar(path: Path, tipo: str) -> dict:
    """Devuelve {hoja: DataFrame} ya ajustado para un archivo."""
    hojas_pedidas = SHEETS_POR_TIPO.get(tipo) or hojas_de(path)
    disponibles = hojas_de(path)
    salida = {}
    for hoja in hojas_pedidas:
        if hoja not in disponibles:
            # La hoja 'UPS' a veces viene con otro nombre: se usa la primera.
            if tipo == "crono_ups" and disponibles:
                alternativa = disponibles[0]
                print(f"    · aviso: no hay hoja '{hoja}'; uso '{alternativa}'")
                hoja = alternativa
            else:
                print(f"    · aviso: el archivo no trae la hoja '{hoja}' (se omite)")
                continue
        salida[hoja] = preparar_hoja(path, tipo, hoja)
    return salida


# =============================================================================
# 4) VALIDACION (usa las mismas reglas del panel, app.py)
# =============================================================================
def validar_con_el_panel(tipo: str, hojas: dict) -> list:
    """Reutiliza funciones de app.py para avisar si el panel tendria problemas.

    Para la Data se usa el MISMO cargador del panel (`_cargar_y_preparar`), escribiendo
    la hoja ya preparada en un archivo temporal: asi la validacion es identica a la real.
    """
    os.environ["INGESTA_SIN_APP"] = "1"
    avisos: list = []
    try:
        sys.path.insert(0, str(BASE))
        import app  # noqa: PLC0415  (import tardio a proposito)
    except Exception as exc:  # noqa: BLE001
        return [f"no pude reutilizar las validaciones de app.py ({exc})"]

    origen = list(hojas.values())[0]

    if tipo == "data":
        faltan = [c for c in COLUMNAS_OBLIGATORIAS["data"]
                  if app._col(origen, c) is None]
        if faltan:
            avisos.append(f"le faltan columnas que el panel usa: {faltan}")
            return avisos
        temporal = BASE / "tools" / "_validacion_data.xlsx"
        try:
            origen.to_excel(temporal, index=False, sheet_name="Hoja1")
            df = app._cargar_y_preparar(str(temporal), os.path.getmtime(temporal))
            df["_componente"] = app.clasificar_componente(df)   # igual que hace el panel
            avisos.append(f"{len(df):,} equipos · componentes "
                          f"{df['_componente'].value_counts().to_dict()}".replace(",", "."))
            dup = int(df["Serial"].duplicated().sum())
            if dup:
                avisos.append(f"{dup} Serial(es) duplicado(s) - 1 fila debe ser 1 equipo")
            sin_sban = int(df["_SBAN"].astype(str).str.len().eq(0).sum())
            if sin_sban:
                avisos.append(f"{sin_sban} fila(s) sin SBAN (no se pueden ubicar por oficina)")
        except Exception as exc:  # noqa: BLE001
            avisos.append(f"no pude validar la Data como lo hace el panel: {exc}")
        finally:
            temporal.unlink(missing_ok=True)

    elif tipo == "campos":
        kpi = hojas.get("Dashboard_KPI")
        if kpi is None:
            avisos.append("falta la hoja Dashboard_KPI (el estado de sede no se integraria)")
        else:
            faltan = [c for c in COLUMNAS_OBLIGATORIAS["campos"]
                      if app._col(kpi, c) is None]
            if faltan:
                avisos.append(f"faltan columnas clave en Dashboard_KPI: {faltan}")
            if app._col(kpi, "ESTADO UPS") is None:
                avisos.append("falta la columna AP 'ESTADO UPS' "
                              "(la pestana de avance UPS quedaria sin estado)")
            else:
                estado = kpi[app._col(kpi, "ESTADO UPS")].dropna()
                if len(estado):
                    avisos.append(f"ESTADO UPS (col AP): {len(estado)} sede(s) reportadas "
                                  f"-> {estado.value_counts().to_dict()}")
        nov = hojas.get("Novedades Equipos")
        if nov is not None:
            avisos.append(f"{len(nov.dropna(how='all'))} novedad(es) de oficina con datos")
    elif tipo in ("crono", "crono_ups"):
        avisos.append(f"{len(origen):,} fila(s)".replace(",", "."))
    return avisos


# =============================================================================
# 5) ANALISIS DE COLUMNAS (reporte para decidir ajustes)
# =============================================================================
def analizar(origen: Path, archivos: dict) -> None:
    """Genera tools/analisis_ingesta.md con columnas, llenado y valores frecuentes."""
    lineas = [f"# Analisis de la ingesta - {ahora():%Y-%m-%d %H:%M}", ""]
    lineas.append(f"Carpeta de origen: `{origen}`")
    lineas.append("")
    for tipo, path in archivos.items():
        lineas.append(f"## {tipo} · `{path.name}`")
        lineas.append("")
        for hoja in hojas_de(path):
            try:
                d = pd.read_excel(path, sheet_name=hoja)
            except Exception as exc:  # noqa: BLE001
                lineas.append(f"### {hoja} (no se pudo leer: {exc})")
                continue
            obligatorias = COLUMNAS_OBLIGATORIAS.get(tipo) or []
            lineas.append(f"### {hoja} · {len(d):,} filas x {len(d.columns)} columnas"
                          .replace(",", "."))
            lineas.append("")
            lineas.append("| # | Columna | Con dato | % | Valores frecuentes |")
            lineas.append("|---|---|---|---|---|")
            for i, c in enumerate(d.columns):
                n = int(d[c].notna().sum())
                pct = (n / len(d) * 100) if len(d) else 0
                vals = [str(v)[:24] for v in d[c].dropna().unique()[:5]]
                marca = " ⚠️ OBLIGATORIA" if str(c) in obligatorias else ""
                lineas.append(f"| {i} | {c}{marca} | {n} | {pct:.0f}% | {vals} |")
            lineas.append("")
        lineas.append("")
    destino = BASE / "tools" / "analisis_ingesta.md"
    destino.write_text("\n".join(lineas), encoding="utf-8")
    print(f"  Reporte escrito en: {destino}")
    print("  Abrilo para ver cada columna, cuanta data trae y sus valores mas repetidos.")


# =============================================================================
# 6) ESCRITURA Y PUBLICACION
# =============================================================================
def escribir(destino: Path, tipo: str, hojas: dict) -> Path:
    """Guarda las hojas preparadas en un unico .xlsx (destino)."""
    destino.parent.mkdir(parents=True, exist_ok=True)
    tmp = destino.with_suffix(".tmp.xlsx")
    with pd.ExcelWriter(tmp, engine="openpyxl") as w:
        for hoja, df in hojas.items():
            df.to_excel(w, sheet_name=hoja[:31], index=False)
    tmp.replace(destino)
    print(f"  OK {tipo} -> {destino.relative_to(BASE)} "
          f"({destino.stat().st_size/1e6:.1f} MB)")
    return destino


def escribir_metadatos(destino_data: Path, detalle: dict) -> None:
    """Escribe ultima_actualizacion.json (fecha/hora, hashes y conteos)."""
    meta = {
        "fecha_hora": ahora().strftime("%Y-%m-%d %H:%M:%S"),
        "organizacion": ORGANIZACION,
        "anonimizado": bool(ANONIMIZAR),
        "archivos": {},
    }
    for tipo, nombre in NOMBRES_DESTINO.items():
        p = destino_data / nombre
        if not p.exists():
            continue
        meta["archivos"][nombre] = {
            "origen": detalle.get(tipo, {}).get("archivo", ""),
            "sha256": hashlib.sha256(p.read_bytes()).hexdigest(),
            "bytes": p.stat().st_size,
        }
    data_xlsx = destino_data / NOMBRES_DESTINO["data"]
    if data_xlsx.exists():
        try:
            meta["filas_data"] = int(len(pd.read_excel(data_xlsx, sheet_name=0)))
        except Exception:  # noqa: BLE001
            pass
    campos_xlsx = destino_data / NOMBRES_DESTINO["campos"]
    if campos_xlsx.exists():
        try:
            hojas = hojas_de(campos_xlsx)
            hoja_nov = "Novedades Equipos" if "Novedades Equipos" in hojas else hojas[-1]
            nov = pd.read_excel(campos_xlsx, sheet_name=hoja_nov)
            meta["filas_novedades"] = int(len(nov.dropna(how="all")))
        except Exception:  # noqa: BLE001
            pass
    (destino_data / "ultima_actualizacion.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  OK metadatos -> {destino_data.relative_to(BASE)}/ultima_actualizacion.json")
    print(f"     fecha_hora={meta['fecha_hora']} · filas_data={meta.get('filas_data')} "
          f"· filas_novedades={meta.get('filas_novedades')}")


def sincronizar_app() -> None:
    """Copia app.py y la documentacion a deploy_panel/ (lo que se publica)."""
    sincronizador = APP_DEPLOY / "sincronizar_app.py"
    if sincronizador.exists():
        codigo = subprocess.run([sys.executable, str(sincronizador)],
                                capture_output=True, text=True, cwd=str(APP_DEPLOY))
        print((codigo.stdout or "").rstrip() or "  (sin salida)")
        if codigo.returncode != 0:
            print("  aviso: no se pudo sincronizar app.py")
        return
    for nombre in ("app.py", "README.md", "TUTORIAL.md"):
        origen, destino = BASE / nombre, APP_DEPLOY / nombre
        if origen.exists():
            texto = origen.read_text(encoding="utf-8").replace("\r\n", "\n")
            if not destino.exists() or destino.read_text(encoding="utf-8") != texto:
                destino.write_text(texto, encoding="utf-8", newline="\n")
                print(f"  ↑ {nombre}")


def _entorno_git() -> dict:
    """Entorno para git: sin prompts y con TLS OpenSSL.

    En este equipo el backend TLS `schannel` de git falla con
    'SEC_E_NO_CREDENTIALS' al conectar con GitHub; con OpenSSL funciona.
    Se puede cambiar en CONFIG_GIT['ssl_backend'].
    """
    entorno = {**os.environ, "GIT_TERMINAL_PROMPT": "0"}
    backend = CONFIG_GIT.get("ssl_backend", "openssl")
    if backend:
        entorno["GIT_CONFIG_COUNT"] = "1"
        entorno["GIT_CONFIG_KEY_0"] = "http.sslBackend"
        entorno["GIT_CONFIG_VALUE_0"] = backend
    if not CONFIG_GIT.get("prompt", False):
        entorno["GCM_INTERACTIVE"] = "never"
    return entorno


def _git(args: list) -> int:
    codigo = subprocess.run(["git"] + args, cwd=str(APP_DEPLOY),
                            capture_output=True, text=True,
                            encoding="utf-8", errors="replace",
                            env=_entorno_git())
    salida = (codigo.stdout or "") + (codigo.stderr or "")
    if salida.strip():
        print("    " + salida.strip()[:900])
    return codigo.returncode


def _credencial_github():
    """Toma usuario y token guardados por Git Credential Manager (si existen).

    Devuelve (usuario, token) o (None, None). El token NUNCA se imprime.
    """
    gcm = CONFIG_GIT.get("credential_manager", "")
    if not gcm or not Path(gcm).exists():
        return None, None
    try:
        p = subprocess.run([gcm, "get"], input="protocol=https\nhost=github.com\n\n",
                           capture_output=True, text=True, timeout=30,
                           env={**os.environ, "GCM_INTERACTIVE": "never"})
        usuario = re.search(r"^username=(.+)$", p.stdout or "", re.M)
        token = re.search(r"^password=(.+)$", p.stdout or "", re.M)
        if token:
            return (usuario.group(1).strip() if usuario else "x"), token.group(1).strip()
    except Exception:  # noqa: BLE001
        pass
    return None, None


def _push(args: list, usuario: str, token: str) -> int:
    """Push inyectando la credencial en la cabecera (evita el helper externo)."""
    basico = base64.b64encode(f"{usuario}:{token}".encode()).decode()
    entorno = _entorno_git()
    entorno.update({"GIT_CONFIG_COUNT": "2",
                    "GIT_CONFIG_KEY_0": "http.sslBackend",
                    "GIT_CONFIG_VALUE_0": CONFIG_GIT.get("ssl_backend", "openssl"),
                    "GIT_CONFIG_KEY_1": "http.https://github.com/.extraHeader",
                    "GIT_CONFIG_VALUE_1": f"Authorization: Basic {basico}"})
    codigo = subprocess.run(["git"] + args, cwd=str(APP_DEPLOY), capture_output=True,
                            text=True, encoding="utf-8", errors="replace", env=entorno)
    salida = ((codigo.stdout or "") + (codigo.stderr or "")).replace(token, "***")
    if salida.strip():
        print("    " + salida.strip()[:900])
    return codigo.returncode


def publicar_en_github(mensaje: str) -> bool:
    """git add + commit + push desde deploy_panel/. Devuelve True si subio."""
    if not (APP_DEPLOY / ".git").exists():
        print("  ERROR: no encuentro el repositorio git en deploy_panel/")
        return False
    _git(["add", "-A"])
    estado = subprocess.run(["git", "status", "--porcelain"], cwd=str(APP_DEPLOY),
                            capture_output=True, text=True,
                            env=_entorno_git()).stdout.strip()
    if estado:
        _git(["-c", f"user.name={CONFIG_GIT['usuario']}",
              "-c", f"user.email={CONFIG_GIT['correo']}",
              "commit", "-m", mensaje])
    else:
        print("    (no hay cambios nuevos que commitear)")
    print("  Subiendo a GitHub...")
    if _git(["push", "origin", "main"]) == 0:
        print("  [OK] Push realizado. Streamlit Cloud se actualiza en 1-2 minutos.")
        return True
    # Plan B: usar la credencial guardada directamente (util cuando el helper de git
    # no puede ejecutarse, p. ej. en entornos sin consola interactiva).
    print("  Reintentando con la credencial guardada...")
    usuario, token = _credencial_github()
    if token and _push(["push", "origin", "main"], usuario, token) == 0:
        print("  [OK] Push realizado. Streamlit Cloud se actualiza en 1-2 minutos.")
        return True
    print("  [ERROR] No se pudo subir a GitHub.")
    print("     Abre CMD o Git Bash y ejecuta una vez: git -C deploy_panel push origin main")
    print("     Los commits ya quedaron hechos; solo falta subirlos.")
    return False


# =============================================================================
# 7) FLUJO PRINCIPAL
# =============================================================================
def main() -> int:
    ap = argparse.ArgumentParser(description="Ingesta diaria del panel MT3")
    ap.add_argument("--publicar", action="store_true",
                    help="commit + push a GitHub al terminar")
    ap.add_argument("--seco", action="store_true",
                    help="escribe los archivos pero NO sube a GitHub")
    ap.add_argument("--analisis", action="store_true",
                    help="solo genera el reporte de columnas (no escribe datos)")
    ap.add_argument("--solo-local", action="store_true",
                    help="escribe solo en data/ local (no toca deploy_panel)")
    args = ap.parse_args()

    print("=" * 74)
    print(" INGESTA DIARIA - PANEL MANTENIMIENTO PREVENTIVO 3")
    print("=" * 74)
    print(f" Origen : {ORIGEN}")
    if not ORIGEN.exists():
        print(f"\n[ERROR] No existe la carpeta de ingesta: {ORIGEN}")
        return 1

    archivos = elegir_archivos(ORIGEN)
    if not archivos:
        print("\n[ERROR] No encontre ningun .xlsx reconocible en la carpeta de ingesta.")
        print("        Debe haber al menos la 'Data' (hoja Hoja1 con 'Consecutivo "
              "mantenimiento 3' y 'Facturable').")
        return 1
    print(f"\n Archivos elegidos ({len(archivos)}):")
    for tipo, path in archivos.items():
        print(f"   · {tipo:<10} {path.name}  ({datetime.fromtimestamp(path.stat().st_mtime):%Y-%m-%d %H:%M})")
    if "data" not in archivos:
        print("\n[ERROR] Falta la Data de ejecucion (es obligatoria).")
        return 1

    if args.analisis:
        print("\n[ANALISIS] Revisando columnas (no se escribe nada)...")
        analizar(ORIGEN, archivos)
        return 0

    # --- Leer y preparar cada archivo ---
    print("\n[1/4] Leyendo y ajustando columnas...")
    preparados, detalle = {}, {}
    for tipo, path in archivos.items():
        hojas = leer_y_preparar(path, tipo)
        if not hojas:
            print(f"  aviso: {path.name} no aporto ninguna hoja util")
            continue
        preparados[tipo] = hojas
        detalle[tipo] = {"archivo": path.name}

    # --- Validacion con las reglas del panel ---
    print("\n[2/4] Validando contra las reglas del panel...")
    problemas = []
    for tipo, hojas in preparados.items():
        faltan = [c for c in (COLUMNAS_OBLIGATORIAS.get(tipo) or [])
                  if buscar_columna(list(hojas.values())[0], c) is None]
        if faltan:
            problemas.append(f"{tipo}: faltan columnas obligatorias {faltan}")
        for aviso in validar_con_el_panel(tipo, hojas):
            print(f"  · {tipo}: {aviso}")
    if problemas:
        print("\n[ERROR] Detengo la publicacion para no romper el panel:")
        for p in problemas:
            print("   ·", p)
        return 2

    # --- Escribir los datos ---
    print("\n[3/4] Escribiendo datos...")
    destinos = [DATA_LOCAL] + ([] if args.solo_local else [DATA_DEPLOY])
    for carpeta in destinos:
        print(f"  En {carpeta.relative_to(BASE)}/")
        for tipo, hojas in preparados.items():
            escribir(carpeta / NOMBRES_DESTINO[tipo], tipo, hojas)
        escribir_metadatos(carpeta, detalle)

    if not args.solo_local:
        # Si este script es la COPIA de respaldo que vive en deploy_panel/tools/,
        # no se sincroniza nada (deploy_panel es su propia raíz).
        if BASE.name.lower() == "deploy_panel":
            print("\n  (copia de respaldo en el repositorio: no se sincroniza nada)")
        else:
            print("\n  Sincronizando app.py, documentacion y scripts con deploy_panel/...")
            sincronizar_app()

    # --- Publicar ---
    print("\n[4/4] Publicacion")
    if args.publicar and not args.solo_local:
        ok = publicar_en_github(f"Ingesta diaria {ahora():%Y-%m-%d %H:%M}")
        print("=" * 74)
        return 0 if ok else 3
    print("  (modo sin push: los archivos quedaron listos en data/ y deploy_panel/data/)")
    print("  Para subir a GitHub:  python ingesta.py --publicar")
    print("=" * 74)
    return 0


if __name__ == "__main__":
    sys.exit(main())
