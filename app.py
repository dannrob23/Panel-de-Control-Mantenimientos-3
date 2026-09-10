# -*- coding: utf-8 -*-
"""
Dashboard Mantenimiento Preventivo 3 (MT3) - Banco Agrario / COLSOF
App Streamlit de un solo archivo. Fuente: Data_PCTriage.xlsx (hoja 'Hoja1').
"""
import hashlib
import io
import os
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ----------------------------------------------------------------------------
# Configuración general de la página
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Control Mantenimiento Preventivo 3",
    page_icon="🛠️",
    layout="wide",
    initial_sidebar_state="expanded",
)

RUTA_XLSX = Path(__file__).resolve().parent / "Data_PCTriage.xlsx"
RUTA_CRONO = Path(__file__).resolve().parent / "Cronograma Mto Preventivo 3 _ Equipos 1.xlsx"
CARPETA_INGESTA = Path(__file__).resolve().parent / "Ingesta de datos diaria"
CARPETA_DATA = Path(__file__).resolve().parent / "data"
CARPETAS_INGESTA = (CARPETA_INGESTA, CARPETA_DATA)
DIR_CARGA = Path(__file__).resolve().parent / "uploads"
DIR_CARGA.mkdir(exist_ok=True)
META_ACTUALIZACION = CARPETA_DATA / "ultima_actualizacion.json"

# Modo público (despliegue sin login): enmascara datos personales.
# Se puede desactivar con el secreto `modo_publico = false` (uso interno).
try:
    MODO_PUBLICO = bool(st.secrets.get("modo_publico", True))
except Exception:  # noqa: BLE001
    MODO_PUBLICO = True


# ----------------------------------------------------------------------------
# Descubrimiento de archivos de la ingesta diaria (data / campos / cronograma)
# ----------------------------------------------------------------------------
def _tipo_archivo(path) -> str:
    """Clasifica un .xlsx: 'data', 'campos' o 'crono'. Devuelve '' si no lo reconoce."""
    try:
        xl = pd.ExcelFile(path)
    except Exception:  # noqa: BLE001
        return ""
    if "Dashboard_KPI" in xl.sheet_names:
        return "campos"
    if "UPS" in xl.sheet_names or "UPS" in Path(path).name.upper():
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


def _tipo_archivo_bytes(data: bytes) -> str:
    try:
        xl = pd.ExcelFile(io.BytesIO(data))
    except Exception:  # noqa: BLE001
        return ""
    if "Dashboard_KPI" in xl.sheet_names:
        return "campos"
    if "UPS" in xl.sheet_names:
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


def _mtime_carpeta(p: Path) -> float:
    try:
        return os.path.getmtime(p)
    except OSError:
        return 0.0


@st.cache_data(show_spinner=False)
def _descubrir_ingesta(clave: str, mtime: float, carpetas: tuple) -> dict:
    """Toma el .xlsx más reciente de cada tipo dentro de las carpetas de ingesta."""
    out = {}
    for cp in carpetas:
        carpeta = Path(cp)
        if not carpeta.exists():
            continue
        archivos = sorted(carpeta.glob("*.xlsx"),
                          key=lambda p: p.stat().st_mtime, reverse=True)
        for p in archivos:
            t = _tipo_archivo(p)
            if t and t not in out:
                out[t] = str(p)
    out.setdefault("data", str(RUTA_XLSX) if RUTA_XLSX.exists() else "")
    out.setdefault("crono", str(RUTA_CRONO) if RUTA_CRONO.exists() else "")
    out.setdefault("crono_ups", "")
    out.setdefault("campos", "")
    return out


def rutas_activas() -> dict:
    carpetas = tuple(str(p) for p in CARPETAS_INGESTA)
    clave = "|".join(carpetas)
    mtime = sum(_mtime_carpeta(Path(c)) for c in carpetas)
    rutas = _descubrir_ingesta(clave, mtime, carpetas)
    for clave_sesion, sesion in (("data", "ruta_data"), ("crono", "ruta_crono"),
                                 ("crono_ups", "ruta_crono_ups"),
                                 ("campos", "ruta_campos")):
        if st.session_state.get(sesion):
            rutas[clave_sesion] = st.session_state[sesion]
    return rutas


def leer_ultima_actualizacion() -> str:
    """Fecha/hora de la última ingesta (archivo de metadatos) con respaldo en mtimes."""
    try:
        if META_ACTUALIZACION.exists():
            import json
            meta = json.loads(META_ACTUALIZACION.read_text(encoding="utf-8"))
            fecha = meta.get("fecha_hora")
            if fecha:
                return str(fecha).replace("T", " ")[:16]
    except Exception:  # noqa: BLE001
        pass
    rutas = [Path(v) for v in rutas_activas().values() if v]
    marcas = [_mtime_carpeta(p) for p in rutas if Path(p).exists()]
    if marcas:
        return pd.Timestamp.fromtimestamp(max(marcas)).strftime("%Y-%m-%d %H:%M")
    return "sin datos"


def _mask_serial(v) -> str:
    s = str(v).strip()
    return "****" + s[-4:] if len(s) > 4 else s


def _mask_placa(v) -> str:
    s = str(v).strip()
    return "****" + s[-4:] if len(s) > 4 else s


def banner_publico():
    if MODO_PUBLICO:
        st.info("🔒 Modo público: datos personales (seriales, placas y observaciones) enmascarados.")


# ----------------------------------------------------------------------------
# Archivos en uso: permite subir versiones actualizadas (auto-detección)
# ----------------------------------------------------------------------------
def ruta_data_act() -> Path:
    return Path(rutas_activas()["data"])


def ruta_crono_act() -> Path:
    return Path(rutas_activas()["crono"])


def ruta_crono_ups_act():
    ruta = rutas_activas().get("crono_ups", "")
    return Path(ruta) if ruta else None


def ruta_campos_act():
    ruta = rutas_activas()["campos"]
    return Path(ruta) if ruta else None


def es_local(p: Path) -> bool:
    try:
        r = p.resolve()
    except Exception:  # noqa: BLE001
        return False
    return r in (RUTA_XLSX.resolve(), RUTA_CRONO.resolve())


def _registrar_carga(tipo: str, archivo_original: str, destino: Path, sha: str, filas: int = 0):
    """Bitácora de auditoría: cada archivo recibido queda registrado (integridad)."""
    reg = DIR_CARGA / "bitacora_cargas.csv"
    import csv
    nuevo = not reg.exists()
    with open(reg, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if nuevo:
            w.writerow(["fecha_hora", "tipo", "archivo_original", "destino",
                        "sha256", "filas"])
        w.writerow([pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
                    tipo, archivo_original, destino.name, sha, filas])


def _guardar_subida(uploaded, tipo: str, sufijo: str) -> Path:
    destino = DIR_CARGA / f"{pd.Timestamp.now():%Y%m%d_%H%M%S}_{sufijo}.xlsx"
    destino.write_bytes(uploaded.getbuffer())
    sha = hashlib.sha256(uploaded.getbuffer()).hexdigest()
    _registrar_carga(tipo, uploaded.name, destino, sha)
    return destino


SUFIJOS = {"data": ("Data_PCTriage", "ruta_data"),
           "campos": ("Campos_dashboard", "ruta_campos"),
           "crono": ("Cronograma_Equipos", "ruta_crono"),
           "crono_ups": ("Cronograma_UPS", "ruta_crono_ups")}
ETIQUETAS = {"data": "Data de ejecución", "campos": "Campos dashboard",
             "crono": "Cronograma Equipos", "crono_ups": "Cronograma UPS"}


def aplicar_subida(uploaded):
    """Detecta el tipo del archivo subido, lo guarda, lo activa y recarga."""
    if uploaded is None:
        return
    datos = uploaded.getbuffer()
    tipo = _tipo_archivo_bytes(bytes(datos))
    if not tipo:
        st.sidebar.error(f"⚠️ No reconocí el archivo `{uploaded.name}`. "
                         "Debe ser Data (Hoja1), Campos dashboard o Cronograma.")
        return
    marca = (uploaded.name, uploaded.size, tipo)
    if st.session_state.get("last_up") == marca:
        return
    with st.spinner(f"Aplicando {uploaded.name} ({ETIQUETAS[tipo]})…"):
        sufijo, sesion = SUFIJOS[tipo]
        destino = _guardar_subida(uploaded, ETIQUETAS[tipo], sufijo)
        st.session_state[sesion] = str(destino)
        st.session_state["last_up"] = marca
        st.cache_data.clear()
        st.rerun()

COLUMNAS_TECNICO = [
    "Serial",
    "Placa",
    "Categoría",
    "_SBAN",
    "Oficina",
    "Consecutivo mantenimiento 3",
    "Facturable",
]


# ----------------------------------------------------------------------------
# Carga y preparación de datos (con caché)
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner="Leyendo Data_PCTriage.xlsx…")
def _cargar_y_preparar(path: str, mtime: float) -> pd.DataFrame:
    """Lee la hoja Hoja1 y normaliza el dataframe. mtime invalida la caché si el archivo cambia."""
    df = pd.read_excel(path, sheet_name="Hoja1")

    # --- Manejo de nulos / normalización de tipos ---
    df["Oficina"] = df["Oficina"].fillna("SIN NOMBRE").astype(str).str.strip()
    df["Categoría"] = df["Categoría"].fillna("SIN CATEGORÍA").astype(str)
    df["Serial"] = df["Serial"].astype(str).str.strip()
    df["Facturable"] = df["Facturable"].fillna("").astype(str).str.strip()
    df["Consecutivo mantenimiento 3"] = (
        df["Consecutivo mantenimiento 3"].fillna("").astype(str).str.strip()
    )

    # Placa: se muestra con ceros a la izquierda cuando es numérica (8 dígitos).
    def _formato_placa(v: str) -> str:
        try:
            return str(int(float(v))).zfill(8)
        except (ValueError, TypeError):
            return v if v and v != "nan" else ""

    df["Placa"] = df["Placa"].apply(
        lambda v: _formato_placa(v) if pd.notna(v) else ""
    ).astype(str)

    # --- SBAN: código de oficina (visible y buscable, formato 5 dígitos) ---
    def _sban5(v):
        try:
            return str(int(float(v))).zfill(5)
        except (ValueError, TypeError):
            return str(v) if pd.notna(v) and str(v) != "nan" else ""

    df["_SBAN"] = df["SBAN"].apply(_sban5)
    df["_ofi_key"] = df["_SBAN"] + " - " + df["Oficina"]

    # --- Lógica MT3: mantenimiento realizado solo si el consecutivo empieza por "MT" ---
    df["_mt"] = df["Consecutivo mantenimiento 3"].str.upper().str.startswith("MT")
    df["_pendiente"] = ~df["_mt"]

    # Orden para que el técnico encuentre rápido lo pendiente
    df = df.sort_values(["Oficina", "_pendiente", "Categoría"]).reset_index(drop=True)
    return df


def cargar_datos() -> pd.DataFrame:
    """Envuelve la carga pasando el mtime del archivo para refrescar caché."""
    p = ruta_data_act()
    if not p.exists():
        raise FileNotFoundError(f"No se encontró el archivo: {p}")
    return _cargar_y_preparar(str(p), os.path.getmtime(p))


@st.cache_data(show_spinner="Leyendo estados del cronograma…")
def _cargar_estados_crono(path: str, mtime: float) -> dict:
    """Lee la columna ESTADO (I) del cronograma y devuelve {SBAN_int: estados}."""
    c = pd.read_excel(path, sheet_name="Hoja1")
    c["SBAN"] = pd.to_numeric(c["SBAN"], errors="coerce")
    c = c.dropna(subset=["SBAN"])
    c["ESTADO"] = c["ESTADO"].astype(str).str.strip()
    c = c[c["ESTADO"].notna() & c["ESTADO"].ne("") & c["ESTADO"].ne("nan")]
    out = {}
    for sban, sub in c.groupby("SBAN"):
        estados = sorted(set(sub["ESTADO"].tolist()))
        out[int(sban)] = " / ".join(estados)
    return out


def cargar_estados_crono() -> dict:
    p = ruta_crono_act()
    if not p.exists():
        return {}
    return _cargar_estados_crono(str(p), os.path.getmtime(p))


# ----------------------------------------------------------------------------
# Cronograma UPS (hoja 'UPS'): fecha planificada y UPS programadas por sede
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner="Leyendo cronograma UPS…")
def _cargar_ups(path: str, mtime: float) -> pd.DataFrame:
    xl = pd.ExcelFile(path)
    hoja = "UPS" if "UPS" in xl.sheet_names else xl.sheet_names[0]
    return xl.parse(hoja)


def cargar_crono_ups():
    p = ruta_crono_ups_act()
    if p is None or not p.exists():
        return None
    return _cargar_ups(str(p), os.path.getmtime(p))


def preparar_crono_ups(u):
    """Consolida el cronograma UPS por SBAN (fecha mínima y UPS programadas)."""
    vacio = pd.DataFrame(columns=["_SBAN", "Oficina", "Regional", "Fecha plan", "UPS plan"])
    if u is None or getattr(u, "empty", True):
        return vacio
    c_sban = _col(u, "SBAN")
    c_fecha = _col(u, "FECHA", "Fecha")
    c_ups = _col(u, "UPS")
    c_nom = _col(u, "Nombre Oficina")
    c_reg = _col(u, "Jefaturas  Operaciones Regional",
                 "Jefaturas\u00a0 Operaciones Regional", "Gerencia Regional", "Gerencia Zonal")
    if c_sban is None:
        return vacio
    tmp = pd.DataFrame({
        "_SBAN": u[c_sban].apply(_pad5),
        "Oficina": u[c_nom].astype(str) if c_nom else "",
        "Regional": u[c_reg].astype(str) if c_reg else "",
        "Fecha plan": pd.to_datetime(u[c_fecha], errors="coerce") if c_fecha else pd.NaT,
        "UPS plan": pd.to_numeric(u[c_ups], errors="coerce").fillna(0) if c_ups else 0,
    })
    tmp = tmp[tmp["_SBAN"].astype(str).str.len() > 0]
    g = tmp.groupby("_SBAN", as_index=False).agg(
        Oficina=("Oficina", lambda s: s.mode().iloc[0] if len(s) else ""),
        Regional=("Regional", lambda s: s.mode().iloc[0] if len(s) else ""),
        **{"Fecha plan": ("Fecha plan", "min"), "UPS plan": ("UPS plan", "sum")})
    return g


def _comp_por_texto(txt) -> str:
    t = str(txt).lower()
    if "ups" in t:
        return "UPS"
    if any(k in t for k in ("impresora", "laser", "láser", "mfp", "multifuncional")):
        return "C2"
    return "C1"


def estados_ups(df: pd.DataFrame, cup: pd.DataFrame) -> dict:
    """Estado de la sede UPS derivado de la FECHA del cronograma UPS y el avance en MT."""
    if cup is None or cup.empty:
        return {}
    ups = df[df["_componente"].eq("UPS")].copy()
    if "_SBAN" not in ups.columns:
        ups["_SBAN"] = ups["SBAN"].apply(_pad5)
    agg = ups.groupby("_SBAN").agg(Total=("Serial", "size"), MT=("_mt", "sum"))
    hoy = pd.Timestamp.now().normalize()
    out = {}
    for _, r in cup.iterrows():
        s = r["_SBAN"]
        fecha = r["Fecha plan"]
        tot = int(agg.loc[s, "Total"]) if s in agg.index else 0
        mt = int(agg.loc[s, "MT"]) if s in agg.index else 0
        if pd.notna(fecha) and hoy < fecha:
            out[s] = "Programada"
        elif tot > 0 and mt >= tot:
            out[s] = "Finalizada"
        elif pd.notna(fecha) and fecha <= hoy:
            out[s] = "En proceso"
        else:
            out[s] = "Programada"
    return out


# ----------------------------------------------------------------------------
# Campos dashboard: estado de la sede, Categoría Novedad (Y) y contador diario
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner="Leyendo Campos dashboard…")
def _cargar_campos(path: str, mtime: float):
    kpi = pd.read_excel(path, sheet_name="Dashboard_KPI")
    try:
        nov = pd.read_excel(path, sheet_name="Novedades Equipos")
    except Exception:  # noqa: BLE001
        nov = pd.DataFrame()
    return kpi, nov


def cargar_campos():
    p = ruta_campos_act()
    if p is None or not p.exists():
        return None, None
    return _cargar_campos(str(p), os.path.getmtime(p))


def _pad5(v):
    try:
        return str(int(float(v))).zfill(5)
    except (ValueError, TypeError):
        return ""


def _col(df, *nombres):
    """Busca una columna tolerando espacios y mayúsculas/minúsculas."""
    mapa = {str(c).strip().lower(): c for c in df.columns}
    for nom in nombres:
        if nom in df.columns:
            return nom
        clave = str(nom).strip().lower()
        if clave in mapa:
            return mapa[clave]
    return None


def preparar_campos(kpi, nov, df=None):
    """Normaliza Dashboard_KPI y Novedades Equipos por SBAN (5 dígitos).

    Asigna `_comp` (C1/C2/UPS) para que las novedades no se trasladen entre módulos:
    - Con Serial: componente del elemento según la Data.
    - Sin Serial: C1 por defecto, salvo que el texto mencione UPS o impresora/láser.
    """
    k = pd.DataFrame(columns=["_SBAN", "Categoría Novedad", "Estado sede", "Obs. novedad"])
    if kpi is not None and not kpi.empty:
        cat = _col(kpi, "Categoria Novedad", "Categoría Novedad")
        est = _col(kpi, "Estado de la sede")
        obs = _col(kpi, "Observaciones actas PCT")
        if cat is None:
            cat = next((c for c in kpi.columns if "ategoria" in str(c) and "ovedad" in str(c)), None)
        sban = kpi[_col(kpi, "SBAN")] if _col(kpi, "SBAN") else pd.Series(dtype=str)
        k = pd.DataFrame({
            "_SBAN": sban.apply(_pad5),
            "Categoría Novedad": (kpi[cat].fillna("").astype(str).str.strip()
                                  if cat else pd.Series([""] * len(kpi))),
            "Estado sede": (kpi[est].fillna("").astype(str).str.strip()
                            if est else pd.Series([""] * len(kpi))),
            "Obs. novedad": (kpi[obs].fillna("").astype(str).str.strip()
                             if obs else pd.Series([""] * len(kpi))),
        })
        k = k[k["_SBAN"].astype(str).str.len() > 0]
        k = k.drop_duplicates("_SBAN", keep="first")

    n = pd.DataFrame()
    if nov is not None and not nov.empty:
        n = nov.copy()
        c_sban = _col(n, "SBAN PCT", "SBAN")
        n["_SBAN"] = n[c_sban].apply(_pad5) if c_sban else ""
        n["_FECHA"] = pd.to_datetime(n[_col(n, "Fecha")], errors="coerce") if _col(n, "Fecha") else pd.NaT
        c_serial = _col(n, "Serial")
        n["_Serial"] = n[c_serial].astype(str).str.strip() if c_serial else ""
        n = n[n["_SBAN"].astype(str).str.len() > 0]
        tiene_serial = n["_Serial"].notna() & ~n["_Serial"].str.lower().isin(["", "nan"])
        clave = (n["_SBAN"].astype(str) + "|" + n["_Serial"].astype(str) + "|"
                 + n["_FECHA"].astype(str))
        clave = clave.where(tiene_serial, "R" + n.index.astype(str))
        n = n[~clave.duplicated(keep="first")]
        for origen, destino in (("Estado", "Estado novedad"),
                                ("Facturable", "Facturable novedad"),
                                ("Observaciones", "Obs. novedad detalle")):
            col = _col(n, origen)
            if col:
                n[destino] = n[col].astype(str).fillna("")
        # Componente por Serial (o por texto si no hay serial)
        mapa = {}
        if df is not None and not df.empty and "Serial" in df.columns \
                and "_componente" in df.columns:
            mapa = dict(zip(df["Serial"].astype(str).str.strip(),
                            df["_componente"].astype(str)))

        texto_n = (n["Obs. novedad detalle"].astype(str) if "Obs. novedad detalle" in n.columns
                   else pd.Series([""] * len(n), index=n.index))

        def _comp_novedad(serial, texto):
            comp = mapa.get(str(serial).strip())
            return comp if comp in ("C1", "C2", "UPS") else _comp_por_texto(texto)

        if len(n):
            n["_comp"] = [_comp_novedad(s, t) for s, t in zip(n["_Serial"], texto_n)]
        else:
            n["_comp"] = []

    if "_comp" not in n.columns:
        n["_comp"] = pd.Series(dtype=str)
    if len(k):
        texto_k = (k["Obs. novedad"].astype(str) if "Obs. novedad" in k.columns
                   else pd.Series([""] * len(k), index=k.index))
        k["_comp"] = [_comp_por_texto(t) for t in texto_k]
    else:
        k["_comp"] = []
    return k, n


def resumen_novedades(k, n):
    """Contador por SBAN: categoría novedad, del día y acumulado (dedupe por Serial+Fecha)."""
    if n is None or n.empty:
        base = (k[["_SBAN", "Categoría Novedad", "Obs. novedad", "Estado sede"]].copy()
                if k is not None and not k.empty else pd.DataFrame(columns=["_SBAN"]))
        base["Novedades del día"] = 0
        base["Novedades acumuladas"] = 0
        base["Última novedad"] = pd.NaT
        base["Fecha de corte"] = pd.NaT
        return base
    corte = n["_FECHA"].max()
    acum = n.groupby("_SBAN").size().rename("Novedades acumuladas")
    dia = n[n["_FECHA"].eq(corte)].groupby("_SBAN").size().rename("Novedades del día")
    ult = n.groupby("_SBAN")["_FECHA"].max().rename("Última novedad")
    res = pd.concat([acum, dia, ult], axis=1).reset_index()
    res["Novedades del día"] = res["Novedades del día"].fillna(0).astype(int)
    res["Novedades acumuladas"] = res["Novedades acumuladas"].fillna(0).astype(int)
    res["Fecha de corte"] = corte
    if k is not None and not k.empty:
        res = res.merge(k, on="_SBAN", how="outer")
        res["Categoría Novedad"] = res["Categoría Novedad"].fillna("")
        res["Obs. novedad"] = res["Obs. novedad"].fillna("")
        res["Estado sede"] = res["Estado sede"].fillna("")
    res["Novedades del día"] = (pd.to_numeric(res["Novedades del día"], errors="coerce")
                                .fillna(0).astype(int))
    res["Novedades acumuladas"] = (pd.to_numeric(res["Novedades acumuladas"], errors="coerce")
                                   .fillna(0).astype(int))
    return res


def clasificar_componente(df: pd.DataFrame) -> pd.Series:
    """C1 = resto · C2 = Impresoras láser · UPS = UPS."""
    cat = df["Categoría"].astype(str).str.strip()
    modelo = df["Modelo"].astype(str)
    laser = cat.str.lower().eq("impresora") & modelo.str.contains("laser", case=False, na=False)
    ups = cat.str.upper().eq("UPS")
    comp = pd.Series("C1", index=df.index)
    comp[laser] = "C2"
    comp[ups] = "UPS"
    return comp


# ----------------------------------------------------------------------------
# Validación de integridad: estructura, consistencia y cruce Data <-> Cronograma
# ----------------------------------------------------------------------------
REQ_DATA = {"Serial", "Placa", "Categoría", "SBAN", "Oficina", "Facturable",
            "Consecutivo mantenimiento 3"}
REQ_CRONO = {"SBAN", "tipo", "Nombre Oficina", "ESTADO", "Fecha Inicio", "Fecha Fin"}
ESTADOS_OK = {"Programada", "En proceso", "Finalizada", "Reprogramada"}
CATS_CRONO = ["Equipo de escritorio", "Equipo portátil", "Escáner", "Impresora",
              "Lector biométrico", "Lector de banda pin pad", "Lector de código",
              "Monitor", "PAD de firmas", "Servidor", "Tablet"]


@st.cache_data(show_spinner="Leyendo cronograma para validar…")
def _cronograma_validacion(path: str, mtime: float) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name="Hoja1")


def validar_integridad(d: pd.DataFrame, c: pd.DataFrame, kpi=None, nov=None, cup=None):
    """Devuelve (criticos, avisos). Criticos detienen la app; avisos se muestran."""
    criticos, avisos = [], []

    falt_d = REQ_DATA - set(d.columns)
    if falt_d:
        criticos.append("❌ Data_PCTriage: faltan columnas: " + ", ".join(sorted(falt_d)))
    if c is None or getattr(c, "empty", True):
        avisos.append("⚠️ Cronograma no disponible: se omite la verificación de fechas/estados.")
        c = pd.DataFrame(columns=sorted(REQ_CRONO))
    falt_c = REQ_CRONO - set(c.columns)
    if falt_c:
        criticos.append("❌ Cronograma: faltan columnas: " + ", ".join(sorted(falt_c)))
    if criticos:
        return criticos, avisos

    # --- Data ---
    dup = int(d["Serial"].duplicated().sum())
    if dup:
        avisos.append(f"⚠️ Data: {dup:,} Serial(es) duplicado(s) — 1 fila debe ser 1 elemento".replace(",", "."))
    si = int((d["Facturable"] == "Si").sum())
    no = int((d["Facturable"] == "No").sum())
    otros = len(d) - si - no
    if otros:
        avisos.append(f"⚠️ Data: {otros} filas sin 'Si'/'No' en Facturable (revisar).")
    g = d["Consecutivo mantenimiento 3"].astype(str).str.strip()
    anomalo = int((g.notna() & g.ne("") & g.ne("nan") & ~g.str.upper().str.startswith("MT")).sum())
    if anomalo:
        avisos.append(f"⚠️ Data: {anomalo} consecutivos MT3 que no empiezan por 'MT' (se cuentan como no realizados).")

    # --- Cronograma ---
    est = c["ESTADO"].dropna().astype(str).str.strip()
    inval = est[~est.isin(ESTADOS_OK)]
    if len(inval):
        avisos.append(f"⚠️ Cronograma: {len(inval)} filas con ESTADO fuera de "
                      f"Programada/En proceso/Finalizada/Reprogramada: {sorted(set(inval))[:6]}.")
    ini = pd.to_datetime(c["Fecha Inicio"], errors="coerce")
    fin = pd.to_datetime(c["Fecha Fin"], errors="coerce")
    n_fecha = int(ini.isna().sum() + fin.isna().sum())
    if n_fecha:
        avisos.append(f"⚠️ Cronograma: {n_fecha} fecha(s) de inicio/fin inválidas.")
    n_inv = int(((ini.notna() & fin.notna() & (fin < ini)).sum()))
    if n_inv:
        avisos.append(f"⚠️ Cronograma: {n_inv} fila(s) con Fecha Fin anterior a Fecha Inicio.")

    # --- Cruce Data <-> Cronograma ---
    d_sb = {int(x) for x in pd.to_numeric(d["SBAN"], errors="coerce").dropna().unique()}
    ofi = c[c["tipo"].isin(["Oficina", "DG"])]
    c_sb = {int(x) for x in pd.to_numeric(ofi["SBAN"], errors="coerce").dropna().unique()}
    sin_data = sorted(c_sb - d_sb)
    if sin_data:
        avisos.append(f"⚠️ Cronograma: {len(sin_data)} oficina(s) sin registro en Data_PCTriage "
                      f"(ej. {sin_data[:5]}).")
    cats = [k for k in CATS_CRONO if k in c.columns]
    if cats:
        crono_tot = pd.to_numeric(ofi[cats].stack(), errors="coerce").groupby(level=0).sum()
        tmp = pd.DataFrame({"SBAN": ofi["SBAN"].values,
                            "Crono": crono_tot.reindex(ofi.index).fillna(0).values})
        tmp["SBAN"] = pd.to_numeric(tmp["SBAN"], errors="coerce")
        ele = d.groupby(pd.to_numeric(d["SBAN"], errors="coerce"))["Serial"].size()
        tmp["Data"] = tmp["SBAN"].map(ele).fillna(0)
        sob = tmp[tmp["Data"] == 0]
        if len(sob):
            avisos.append(f"⚠️ {len(sob)} oficina(s) programadas sin elementos en Data "
                          f"(revisar si el cronograma incluye equipos no cargados).")

    # --- Campos dashboard / Novedades ---
    if kpi is None or getattr(kpi, "empty", True):
        avisos.append("⚠️ Campos dashboard no disponible: no se integran Categoría Novedad "
                      "ni el contador diario de novedades.")
    else:
        cat = next((x for x in kpi.columns if "ategoria" in str(x) and "ovedad" in str(x)), None)
        if cat is None:
            avisos.append("⚠️ Campos dashboard: no se encontró la columna 'Categoria Novedad'.")
        else:
            vals = set(kpi[cat].dropna().astype(str).str.strip())
            vals.discard("")
            raros = sorted(vals - {"DENUNCIA", "UBICACION ERRADA", "ADICIONAL", "Ninguna novedad"})
            if raros:
                avisos.append(f"⚠️ Categoria Novedad con valores no esperados: {raros[:6]}.")
    if nov is not None and not getattr(nov, "empty", True):
        faltan_n = {"Fecha"} - set(nov.columns)
        if _col(nov, "SBAN PCT", "SBAN") is None:
            faltan_n.add("SBAN PCT/SBAN")
        if _col(nov, "Serial") is None:
            faltan_n.add("Serial")
        if faltan_n:
            avisos.append("⚠️ Novedades Equipos: faltan columnas " + ", ".join(sorted(faltan_n)) + ".")
        else:
            col_sb = _col(nov, "SBAN PCT", "SBAN")
            sb = nov[col_sb] if col_sb else pd.Series(dtype=str)
            sb = {_pad5(v) for v in sb.dropna()}
            sb.discard("")
            en_data = {_pad5(x) for x in pd.to_numeric(d["SBAN"], errors="coerce").dropna().unique()}
            fuera = sorted(sb - en_data)
            if fuera:
                avisos.append(f"⚠️ Novedades Equipos: {len(fuera)} SBAN sin registro en la Data "
                              f"(ej. {fuera[:5]}).")

    # --- Cronograma UPS ---
    if cup is None or getattr(cup, "empty", True):
        avisos.append("⚠️ Cronograma UPS no disponible: la pestaña UPS quedará sin fecha/estado.")
    else:
        c_sban = _col(cup, "SBAN")
        c_fecha = _col(cup, "FECHA", "Fecha")
        c_ups = _col(cup, "UPS")
        faltan_u = set()
        if c_sban is None:
            faltan_u.add("SBAN")
        if c_fecha is None:
            faltan_u.add("FECHA")
        if c_ups is None:
            faltan_u.add("UPS")
        if faltan_u:
            avisos.append("⚠️ Cronograma UPS: faltan columnas " + ", ".join(sorted(faltan_u)) + ".")
        else:
            sb = {_pad5(v) for v in cup[c_sban].dropna()}
            sb.discard("")
            d_ups = d[d["Categoría"].astype(str).str.upper().eq("UPS")]
            en_data = {_pad5(v) for v in d_ups["SBAN"].dropna()}
            fuera = sorted(sb - en_data)
            if fuera:
                avisos.append(f"⚠️ Cronograma UPS: {len(fuera)} SBAN sin UPS en la Data "
                              f"(ej. {fuera[:5]}).")
            plan = int(pd.to_numeric(cup[c_ups], errors="coerce").fillna(0).sum())
            real = int(len(d_ups))
            if plan and plan != real:
                avisos.append(f"⚠️ UPS programadas {plan} vs UPS en Data {real} "
                              f"(diferencia {plan - real:+d}).")
    return criticos, avisos


# ----------------------------------------------------------------------------
# Almacén de observaciones del técnico (por Serial, persiste entre filtros)
# ----------------------------------------------------------------------------
def obs_estado() -> dict:
    if "obs" not in st.session_state:
        st.session_state["obs"] = {}
    return st.session_state["obs"]


# ----------------------------------------------------------------------------
# Sidebar: filtros
# ----------------------------------------------------------------------------
def render_sidebar(df: pd.DataFrame):
    with st.sidebar:
        st.markdown("## 📂 Ingesta diaria")
        st.caption("Sube los .xlsx del día: **Data**, **Campos dashboard** y/o **Cronograma**. "
                   "El tipo se detecta solo y queda bitácora en `uploads\\bitacora_cargas.csv`.")
        subidos = st.file_uploader(
            "Archivos (.xlsx)", type=["xlsx"], accept_multiple_files=True, key="up_files",
            help="Data (Hoja1) · Campos dashboard (Dashboard_KPI) · Cronograma (ESTADO).")
        for archivo in subidos or []:
            aplicar_subida(archivo)

        rutas = rutas_activas()
        p_data, p_crono = Path(rutas["data"]), Path(rutas["crono"])
        p_campos = rutas["campos"]
        local = es_local(p_data) and es_local(p_crono)
        fuentes = (f"Data: `{p_data.name}` · Cronograma: `{p_crono.name}` · "
                   + (f"Campos: `{Path(p_campos).name}`" if p_campos else "Campos: —"))
        st.caption(fuentes)
        st.caption(f"🕒 Última actualización: **{leer_ultima_actualizacion()}**")
        if not local:
            st.info("Usando archivos subidos (carpeta `uploads`).")
        if st.button("↩️ Volver a archivos locales",
                     disabled=local, key="reset_files"):
            for k in ("ruta_data", "ruta_crono", "ruta_crono_ups", "ruta_campos", "last_up"):
                st.session_state.pop(k, None)
            st.cache_data.clear()
            st.rerun()

        st.divider()
        st.markdown("## 🛠️ Filtros")
        opciones = sorted(df["_ofi_key"].dropna().unique().tolist())

        seleccion = st.multiselect(
            "Oficina (SBAN - Nombre)",
            options=opciones,
            default=[],
            placeholder="Todas las oficinas (vista global)",
            help="Busque por SBAN o por nombre. Sin selección se muestran TODAS las oficinas. Use Ctrl+clic para varias.",
        )

        solo_pendientes = st.checkbox(
            "Mostrar solo pendientes ⚠️",
            value=False,
            help="Filtra la tabla y métricas a los equipos que aún no tienen consecutivo MT3.",
        )

        f_attr_label = st.radio(
            "Distinguir por facturación (AN)",
            options=["Todos", "BANCO (Facturable Si)", "COLSOF (No facturable)"],
            index=0,
            help="Según columna Facturable: Si = atribuible a BANCO · No = gestión COLSOF.",
        )
        f_attr = {"Todos": "T",
                  "BANCO (Facturable Si)": "Si",
                  "COLSOF (No facturable)": "No"}[f_attr_label]

        st.divider()
        if st.button("🔄 Recargar datos", width="stretch"):
            st.cache_data.clear()
            st.rerun()

        with st.expander("🗂️ Auditoría y fuentes", expanded=False):
            rutas2 = rutas_activas()
            p_data2, p_crono2 = Path(rutas2["data"]), Path(rutas2["crono"])
            p_campos2, p_ups2 = rutas2["campos"], rutas2.get("crono_ups", "")
            st.caption(
                f"**Data en uso:** `{p_data2.name}`\n"
                f"· Modif.: {pd.Timestamp.fromtimestamp(os.path.getmtime(p_data2)):%Y-%m-%d %H:%M}\n"
                f"**Cronograma Equipos:** `{p_crono2.name}`\n"
                f"· Modif.: {pd.Timestamp.fromtimestamp(os.path.getmtime(p_crono2)):%Y-%m-%d %H:%M}\n"
                + (f"**Campos dashboard:** `{Path(p_campos2).name}`\n"
                   f"· Modif.: {pd.Timestamp.fromtimestamp(os.path.getmtime(p_campos2)):%Y-%m-%d %H:%M}\n"
                   if p_campos2 else "**Campos dashboard:** —\n")
                + (f"**Cronograma UPS:** `{Path(p_ups2).name}`\n"
                   f"· Modif.: {pd.Timestamp.fromtimestamp(os.path.getmtime(p_ups2)):%Y-%m-%d %H:%M}"
                   if p_ups2 else "**Cronograma UPS:** —")
            )
            st.caption(f"🕒 Última actualización: **{leer_ultima_actualizacion()}**")
            bitacora = DIR_CARGA / "bitacora_cargas.csv"
            if bitacora.exists():
                with open(bitacora, "rb") as fh:
                    st.download_button("⬇️ Bitácora de cargas (CSV)", data=fh.read(),
                                       file_name="bitacora_cargas.csv", mime="text/csv",
                                       width="stretch")
        return seleccion, solo_pendientes, f_attr


# ----------------------------------------------------------------------------
# Métricas KPI
# ----------------------------------------------------------------------------
def render_kpis(datos: pd.DataFrame):
    total = len(datos)
    realizados = int(datos["_mt"].sum())
    pendientes = int(datos["_pendiente"].sum())
    fact_si = int((datos["Facturable"] == "Si").sum())
    fact_no = int((datos["Facturable"] == "No").sum())

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("📦 Total de elementos", f"{total:,}".replace(",", "."),
              help="Equipos en la(s) oficina(s) seleccionada(s)")
    c2.metric("✅ Mantenimientos realizados (MT)", f"{realizados:,}".replace(",", "."),
              help="Filas cuyo 'Consecutivo mantenimiento 3' empieza por MT")
    c3.metric("⚠️ Mantenimientos pendientes", f"{pendientes:,}".replace(",", "."),
              delta=f"{pendientes/total*100:.1f}%" if total else "0%",
              delta_color="inverse", help="Total de elementos − Realizados")
    c4.metric("🏦 Facturables (Si)", f"{fact_si:,}".replace(",", "."),
              help="Elementos atribuibles a BANCO")
    c5.metric("🏢 No facturables (No)", f"{fact_no:,}".replace(",", "."),
              help="Elementos de gestión COLSOF")
    return total, realizados, pendientes, fact_si, fact_no


# ----------------------------------------------------------------------------
# Gráficas Plotly
# ----------------------------------------------------------------------------
def grafico_dona(realizados: int, pendientes: int):
    total = realizados + pendientes
    pct = (realizados / total * 100) if total else 0.0
    fig = go.Figure(go.Pie(
        labels=["✅ Subsanados (MT)", "⚠️ Pendientes"],
        values=[realizados, pendientes],
        hole=0.68,
        marker=dict(colors=["#2E9E5B", "#D64541"], line=dict(color="#FFFFFF", width=2)),
        textinfo="none",
        hovertemplate="<b>%{label}</b><br>%{value:,} equipos (%{percent})<extra></extra>",
    ))
    fig.update_layout(
        title=dict(text="<b>Avance de Mantenimiento Preventivo 3</b>",
                   font=dict(size=15), x=0.02),
        annotations=[dict(
            text=(f"<span style='font-size:28px;font-weight:700'>{pct:.1f}%</span>"
                  f"<br><span style='font-size:11px;letter-spacing:.5px'>AVANCE GLOBAL</span>"),
            x=0.5, y=0.5, showarrow=False, align="center")],
        margin=dict(l=15, r=15, t=45, b=10),
        height=360,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.08, xanchor="center", x=0.5),
    )
    return fig


def grafico_top_pendientes(datos: pd.DataFrame, top_n: int = 15):
    """Ranking horizontal de sedes con más pendientes (Realizados vs Pendientes)."""
    ag = (datos.assign(_sb=datos["_SBAN"] + " " + datos["Oficina"].str[:22])
          .groupby("_sb").agg(Realizados=("_mt", "sum"), Pendientes=("_pendiente", "sum")))
    ag["Total"] = ag["Realizados"] + ag["Pendientes"]
    ag = ag[ag["Total"] > 0].sort_values("Pendientes", ascending=False).head(top_n)
    ag = ag.sort_values("Pendientes", ascending=True)
    fig = go.Figure()
    fig.add_bar(y=ag.index, x=ag["Realizados"], orientation="h", name="✅ Subsanados (MT)",
                marker_color="#2E9E5B",
                hovertemplate="<b>%{y}</b><br>Subsanados: %{x:,}<extra></extra>")
    fig.add_bar(y=ag.index, x=ag["Pendientes"], orientation="h", name="⚠️ Pendientes",
                marker_color="#D64541",
                hovertemplate="<b>%{y}</b><br>Pendientes: %{x:,}<extra></extra>")
    fig.update_layout(
        title=dict(text=f"<b>Top {top_n} sedes con más pendientes de MT3</b>",
                   font=dict(size=15), x=0.02),
        barmode="stack",
        height=360,
        margin=dict(l=10, r=15, t=45, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(title="Equipos", gridcolor="rgba(128,128,128,.25)"),
        yaxis=dict(title=""),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
    )
    return fig


def grafico_barras(datos: pd.DataFrame, limite: int = 30):
    """Facturables (Si) vs No facturables (No) por Oficina."""
    tabla = (
        datos.assign(Facturable=datos["Facturable"].replace({"Si": "Facturables (Si)", "No": "No facturables (No)"}))
        .groupby(["Oficina", "Facturable"], dropna=False)
        .size()
        .unstack(fill_value=0)
    )
    for col in ["Facturables (Si)", "No facturables (No)"]:
        if col not in tabla.columns:
            tabla[col] = 0
    tabla["Total"] = tabla.sum(axis=1)
    tabla = tabla.sort_values("Total", ascending=False).head(limite).drop(columns="Total")

    fig = go.Figure()
    fig.add_bar(x=tabla.index, y=tabla["Facturables (Si)"],
                name="Facturables (Si)", marker_color="#2E75B6",
                hovertemplate="%{x}<br>Facturables: %{y:,}<extra></extra>")
    fig.add_bar(x=tabla.index, y=tabla["No facturables (No)"],
                name="No facturables (No)", marker_color="#E67E22",
                hovertemplate="%{x}<br>No facturables: %{y:,}<extra></extra>")
    n_oficinas = len(datos["Oficina"].unique())
    if n_oficinas > limite:
        st.caption(f"⚠️ Se muestran las {limite} oficinas con más equipos (de {n_oficinas}). Ajuste el filtro de Oficina para ver todas.")
    fig.update_layout(
        title=dict(text="Facturables vs No facturables por Oficina", x=0.02),
        barmode="group",
        height=360,
        margin=dict(l=10, r=10, t=50, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(title="", tickangle=-45),
        yaxis=dict(title="Elementos"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
    )
    return fig


# ----------------------------------------------------------------------------
# Colores por estado del cronograma (web y Excel)
# ----------------------------------------------------------------------------
ESTADO_COLORES = {
    "Finalizada": ("C6EFCE", "006100"),      # verde
    "En proceso": ("FFEB9C", "7F6000"),      # amarillo
    "Programada": ("FFFFFF", "1A1A1A"),      # blanco
    "Reprogramada": ("E4DFEC", "3B2A63"),    # morado suave
    "Sin cronograma": ("EDEDED", "333333"),  # gris
}


def _color_de_estado(valor):
    texto = str(valor)
    for clave in ("Finalizada", "En proceso", "Reprogramada", "Programada"):
        if clave in texto:
            return ESTADO_COLORES[clave]
    return ESTADO_COLORES.get(texto, ("EDEDED", "333333"))


COLS_NOVEDAD = ("Categoría Novedad", "Novedades del día", "Novedades acumuladas")


def _css_fila(row):
    total = str(row.get("SBAN", "")) == "TOTAL"
    if total:
        base = "font-weight:bold;background-color:#DDEBF7;color:#000000;"
    else:
        bg, fg = _color_de_estado(row.get("Estado cronograma", ""))
        base = f"background-color:#{bg};color:#{fg};"
    css = []
    for col in row.index:
        if not total and col in COLS_NOVEDAD:
            val = str(row.get(col, "")).strip()
            if col == "Categoría Novedad":
                tiene = bool(val) and val.lower() not in ("", "nan", "sin novedad")
            elif col == "Novedades acumuladas":
                tiene = False
            else:
                try:
                    tiene = float(val) > 0
                except ValueError:
                    tiene = False
            if tiene:
                css.append("background-color:#FCE4D6;color:#833C00;font-weight:bold;")
                continue
        css.append(base)
    return css


def generar_excel_resumen(res: pd.DataFrame) -> bytes:
    """Devuelve un .xlsx con la vista Resumen por oficina (coloreada) + gráficos."""
    from openpyxl import Workbook
    from openpyxl.chart import DoughnutChart, BarChart, Reference
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    buf = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen por oficina"
    hdrs = list(res.columns)
    for j, h in enumerate(hdrs, 1):
        c = ws.cell(1, j, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="FF1F4E78")
    ancho = {"SBAN": 10, "Oficina": 40, "Estado cronograma": 22, "Total elementos": 14,
             "Subsanados (MT)": 15, "Pendientes": 12, "% Avance": 10,
             "Categoría Novedad": 20, "Novedades del día": 14,
             "Novedades acumuladas": 16, "Obs. novedad": 45}
    for j, h in enumerate(hdrs, 1):
        ws.column_dimensions[get_column_letter(j)].width = ancho.get(h, 14)
    for i, row in enumerate(res.itertuples(index=False), start=2):
        es_total = str(row[0]) == "TOTAL"
        for j, val in enumerate(row, 1):
            c = ws.cell(i, j, val)
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                c.number_format = "#,##0"
        if es_total:
            bg, fg = "DDEBF7", "000000"
        else:
            bg, fg = _color_de_estado(row[2])
        for cell in ws[i]:
            cell.fill = PatternFill("solid", fgColor="FF" + bg)
            cell.font = Font(bold=es_total, color="FF" + fg)
    ws.freeze_panes = "A2"

    # Hoja con gráficos
    ws2 = wb.create_sheet("Gráficos")
    sin_total = res[res.iloc[:, 0].astype(str) != "TOTAL"]
    ws2["A1"] = "Subsanados (MT)"; ws2["B1"] = int(sin_total["Subsanados (MT)"].sum())
    ws2["A2"] = "Pendientes"; ws2["B2"] = int(sin_total["Pendientes"].sum())

    cab = ["Estado cronograma", "Subsanados (MT)", "Pendientes"]
    orden = ["Programada", "En proceso", "Finalizada", "Reprogramada", "Sin cronograma"]
    rows = res[res["SBAN"] != "TOTAL"].copy()
    ag = rows.groupby("Estado cronograma")[["Subsanados (MT)", "Pendientes"]].sum()
    for j, h in enumerate(cab, 1):
        ws2.cell(4, j, h).font = Font(bold=True)
    fila = 5
    for estado in orden:
        if estado in ag.index:
            ws2.cell(fila, 1, estado)
            ws2.cell(fila, 2, int(ag.loc[estado, "Subsanados (MT)"]))
            ws2.cell(fila, 3, int(ag.loc[estado, "Pendientes"]))
            fila += 1

    dona = DoughnutChart()
    dona.title = "Avance: Subsanados vs Pendientes"
    dona.add_data(Reference(ws2, min_col=2, min_row=1, max_row=2))
    dona.set_categories(Reference(ws2, min_col=1, min_row=1, max_row=2))
    dona.width, dona.height = 13, 9

    bar = BarChart()
    bar.type = "col"
    bar.grouping = "stacked"
    bar.overlap = 100
    bar.title = "Subsanados vs Pendientes por Estado del cronograma"
    bar.add_data(Reference(ws2, min_col=2, min_row=4, max_col=3, max_row=fila - 1),
                 titles_from_data=True)
    bar.set_categories(Reference(ws2, min_col=1, min_row=5, max_row=fila - 1))
    bar.width, bar.height = 18, 9
    if bar.series:
        bar.series[0].graphicalProperties.solidFill = "2E9E5B"
        bar.series[1].graphicalProperties.solidFill = "D64541"
    ws2.add_chart(dona, "A12")
    ws2.add_chart(bar, "I12")

    # Gráfico de novedades del día por sede (si aplica)
    if "Novedades del día" in sin_total.columns:
        top = sin_total.sort_values("Novedades del día", ascending=False).head(15)
        if int(top["Novedades del día"].sum()) > 0:
            r0 = 22
            ws2.cell(r0, 1, "Sede").font = Font(bold=True)
            ws2.cell(r0, 2, "Novedades del día").font = Font(bold=True)
            r = r0 + 1
            for _, x in top.iterrows():
                if int(x["Novedades del día"]) <= 0:
                    break
                ws2.cell(r, 1, f"{x['SBAN']} {str(x['Oficina'])[:18]}")
                ws2.cell(r, 2, int(x["Novedades del día"]))
                r += 1
            if r > r0 + 1:
                nov = BarChart()
                nov.type = "bar"
                nov.title = "Top sedes con novedades del día"
                nov.add_data(Reference(ws2, min_col=2, min_row=r0, max_row=r - 1),
                             titles_from_data=True)
                nov.set_categories(Reference(ws2, min_col=1, min_row=r0 + 1, max_row=r - 1))
                nov.series[0].graphicalProperties.solidFill = "F4B183"
                nov.width, nov.height = 16, 9
                ws2.add_chart(nov, "A32")

    wb.save(buf)
    return buf.getvalue()


def render_resumen(base: pd.DataFrame, seleccion: list, tipo: str = "T", nov_res=None):
    """Resumen por oficina: Total, Subsanados, Pendientes, % Avance y Novedades.

    tipo: 'T' todos · 'Si' solo Facturables BANCO · 'No' solo No facturables COLSOF.
    nov_res: tabla por SBAN con Categoría Novedad y contadores diarios/acumulados.
    """
    etiqueta = {"T": "Todos", "Si": "BANCO (Facturables Si)", "No": "COLSOF (No facturables)"}.get(tipo, "Todos")
    st.markdown("---")
    with st.expander(
        "📊 Resumen por oficina: Total de elementos · Subsanados · Pendientes",
        expanded=bool(seleccion),
    ):
        if base.empty:
            st.info("Sin oficinas para los filtros seleccionados.")
            return
        if tipo in ("Si", "No"):
            base = base[base["Facturable"] == tipo]
            if base.empty:
                st.info(f"No hay elementos **{etiqueta}** en las oficinas seleccionadas.")
                return
        nov_lookup = {}
        if nov_res is not None and not nov_res.empty and "_SBAN" in nov_res.columns:
            def _n0(v):
                try:
                    return 0 if pd.isna(v) else int(v)
                except (TypeError, ValueError):
                    return 0
            for _, nr in nov_res.iterrows():
                nov_lookup[str(nr["_SBAN"])] = (
                    str(nr.get("Categoría Novedad", "") or "").strip(),
                    _n0(nr.get("Novedades del día", 0)),
                    _n0(nr.get("Novedades acumuladas", 0)),
                    str(nr.get("Obs. novedad", "") or "").strip(),
                )
        filas = []
        for key, sub in base.groupby("_ofi_key", sort=True):
            total = len(sub)
            subsanados = int(sub["_mt"].sum())
            pendientes = total - subsanados
            avance = (subsanados / total * 100) if total else 0.0
            estados = sorted({str(x) for x in sub["_est_crono"].dropna().tolist() if str(x).strip()})
            sban_ofi = str(sub["_SBAN"].iloc[0])
            cat_nov, n_dia, n_acum, obs_nov = nov_lookup.get(sban_ofi, ("", 0, 0, ""))
            filas.append({
                "SBAN": sban_ofi,
                "Oficina": sub["Oficina"].iloc[0],
                "Estado cronograma": " / ".join(estados) if estados else "Sin cronograma",
                "Total elementos": total,
                "Subsanados (MT)": subsanados,
                "Pendientes": pendientes,
                "% Avance": f"{avance:.1f}%".replace(".", ","),
                "Categoría Novedad": cat_nov if cat_nov else "Sin novedad",
                "Novedades del día": n_dia,
                "Novedades acumuladas": n_acum,
                "Obs. novedad": "" if MODO_PUBLICO else obs_nov,
            })
        res = pd.DataFrame(filas)
        datos_resumen = res.copy()  # sin fila TOTAL

        # ---------------- Filtros intuitivos de la vista ----------------
        estados_posibles = ["Programada", "En proceso", "Finalizada", "Reprogramada", "Sin cronograma"]
        presentes = sorted({e for v in datos_resumen["Estado cronograma"] for e in estados_posibles if e in v})
        f1, f2, f3 = st.columns([2.2, 1.6, 1])
        with f1:
            sel_est = st.multiselect("🗂️ Estado cronograma", options=presentes, default=presentes,
                                     key="res_estado", placeholder="Todos los estados")
        with f2:
            min_av = st.slider("🎯 Avance mínimo (%)", 0, 100, 0, step=5,
                               key="res_min_av", help="Muestra oficinas con avance igual o mayor")
        with f3:
            solo_pend = st.checkbox("Solo pendientes\n(< 100%)", key="res_solo_pend")

        vis = datos_resumen.copy()
        vis["_av"] = vis["% Avance"].str.replace(",", ".").str.rstrip("%").astype(float)
        if sel_est:
            vis = vis[vis["Estado cronograma"].apply(lambda e: any(s in e for s in sel_est))]
        vis = vis[vis["_av"] >= min_av]
        if solo_pend:
            vis = vis[vis["_av"] < 100.0]
        vis = vis.drop(columns="_av").reset_index(drop=True)

        # ---------------- Filtros de novedades ----------------
        if "Novedades del día" in datos_resumen.columns:
            g1, g2 = st.columns([2.2, 1])
            cats_nov = sorted({str(c) for c in datos_resumen["Categoría Novedad"]
                               if str(c).strip() and str(c).strip().lower() not in ("nan", "sin novedad")})
            with g1:
                sel_cat = st.multiselect("🏷️ Categoría Novedad", options=cats_nov,
                                         default=cats_nov, key="res_catnov",
                                         placeholder="Todas las categorías")
            with g2:
                solo_nov = st.checkbox("Solo novedades del día", key="res_solo_nov")
            if sel_cat:
                vis = vis[vis["Categoría Novedad"].astype(str).isin(sel_cat)]
            if solo_nov:
                vis = vis[pd.to_numeric(vis["Novedades del día"], errors="coerce").fillna(0) > 0]
            vis = vis.reset_index(drop=True)

        if vis.empty:
            st.info("Sin oficinas que cumplan los filtros del resumen.")
            return

        n_vis = len(vis)
        vt = int(vis["Total elementos"].sum())
        vs = int(vis["Subsanados (MT)"].sum())
        vp = vt - vs
        va = (vs / vt * 100) if vt else 0.0
        nd = (int(pd.to_numeric(vis["Novedades del día"], errors="coerce").fillna(0).sum())
              if "Novedades del día" in vis.columns else 0)
        na = (int(pd.to_numeric(vis["Novedades acumuladas"], errors="coerce").fillna(0).sum())
              if "Novedades acumuladas" in vis.columns else 0)
        vis.loc[n_vis] = ["TOTAL", f"{n_vis} oficina(s)", "", vt, vs, vp,
                          f"{va:.1f}%".replace(".", ","), "", nd, na, ""]
        if MODO_PUBLICO and "Obs. novedad" in vis.columns:
            vis = vis.drop(columns=["Obs. novedad"])

        st.caption(
            f"Atribución: **{etiqueta}** · Mostrando **{n_vis}** de {len(datos_resumen)} oficinas · "
            "Cada fila agrupa SBAN + nombre de oficina."
        )
        st.dataframe(
            vis.style.apply(_css_fila, axis=1),
            width="stretch",
            hide_index=True,
            height=min(420, 38 + 34 * len(vis)),
        )
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Oficinas", f"{n_vis:,}".replace(",", "."))
        m2.metric("Total elementos", f"{vt:,}".replace(",", "."))
        m3.metric("Subsanados", f"{vs:,}".replace(",", "."))
        m4.metric("Pendientes", f"{vp:,}".replace(",", "."))
        m5.metric("Avance", f"{va:.1f}%".replace(".", ","),
                  delta=f"{vp:,} pendientes".replace(",", "."),
                  delta_color="inverse")
        if "Novedades del día" in vis.columns:
            sedes_nov = int((pd.to_numeric(vis.loc[:n_vis - 1, "Novedades del día"],
                                           errors="coerce").fillna(0) > 0).sum())
            n1, n2, n3 = st.columns(3)
            n1.metric("🏷️ Novedades del día", f"{nd:,}".replace(",", "."))
            n2.metric("🏢 Sedes con novedad hoy", f"{sedes_nov:,}".replace(",", "."))
            n3.metric("📚 Novedades acumuladas", f"{na:,}".replace(",", "."))
        st.download_button(
            "⬇️ Descargar Excel (vista filtrada + gráficos + novedades)",
            data=generar_excel_resumen(vis),
            file_name="Resumen_por_oficina_MT3.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
        st.caption("🎨 Verde = Finalizada · Amarillo = En proceso · Blanco = Programada · "
                   "Morado = Reprogramada · Gris = Sin cronograma · "
                   "Naranja = sede con novedad del día.")


# ----------------------------------------------------------------------------
# Tabla de gestión de novedades (st.data_editor)
# ----------------------------------------------------------------------------
def render_tabla(datos: pd.DataFrame, seleccion: list, solo_pendientes: bool):
    st.markdown("---")
    st.markdown("## 📋 Gestión de novedades")
    st.caption(
        "Escriba directamente en la columna **Observaciones / Novedades**. "
        "Sus anotaciones se conservan al cambiar de filtro y puede exportarlas abajo."
    )

    obs = obs_estado()

    vista = datos[COLUMNAS_TECNICO].copy()
    vista = vista.rename(columns={"_SBAN": "SBAN"})
    vista["Estado cronograma"] = datos["_est_crono"].astype(str)
    vista["Estado"] = np_where(datos["_mt"].loc[vista.index], "✅ Subsanado", "⚠️ Pendiente")
    vista["Observaciones / Novedades"] = vista["Serial"].map(obs).fillna("")
    vista = vista[["Serial", "Placa", "Categoría", "SBAN", "Oficina", "Estado cronograma",
                   "Consecutivo mantenimiento 3", "Facturable", "Estado",
                   "Observaciones / Novedades"]]
    if MODO_PUBLICO:
        vista["Serial"] = vista["Serial"].apply(_mask_serial)
        vista["Placa"] = vista["Placa"].apply(_mask_placa)

    clave_editor = "editor_" + hashlib.md5(
        (str(sorted(datos["_ofi_key"].unique().tolist())) + "|" + str(solo_pendientes)).encode()
    ).hexdigest()[:10]

    column_config = {
        "Serial": st.column_config.TextColumn("Serial", disabled=True, width="medium"),
        "Placa": st.column_config.TextColumn("Placa", disabled=True, width="small"),
        "Categoría": st.column_config.TextColumn("Categoría", disabled=True, width="medium"),
        "SBAN": st.column_config.TextColumn("SBAN", disabled=True, width="small",
            help="Código de la oficina (5 dígitos)"),
        "Oficina": st.column_config.TextColumn("Oficina", disabled=True, width="medium"),
        "Estado cronograma": st.column_config.TextColumn(
            "Estado cronograma", disabled=True, width="medium",
            help="ESTADO de la oficina según Cronograma Mto Preventivo 3 (columna I): "
                 "Programada / En proceso / Finalizada / Reprogramada"),
        "Consecutivo mantenimiento 3": st.column_config.TextColumn(
            "Consecutivo MT3 (G)", disabled=True, width="small",
            help="Empieza por MT = mantenimiento realizado"),
        "Facturable": st.column_config.TextColumn("Facturable", disabled=True, width="small"),
        "Estado": st.column_config.TextColumn("Estado", disabled=True, width="small"),
        "Observaciones / Novedades": st.column_config.TextColumn(
            "Observaciones / Novedades", required=False, width="large",
            help="Registre novedades, alertas o equipos faltantes por mantener."),
    }

    editada = st.data_editor(
        vista,
        key=clave_editor,
        hide_index=True,
        num_rows="fixed",
        width="stretch",
        height=520,
        disabled=False,
        column_config=column_config,
    )

    # Guardar observaciones por Serial (persisten entre filtros)
    for _, fila in editada.iterrows():
        texto = fila["Observaciones / Novedades"]
        if pd.notna(texto) and str(texto).strip():
            obs[str(fila["Serial"]).strip()] = str(texto).strip()

    con_obs = sum(1 for v in obs.values() if v.strip())
    st.caption(f"💾 Observaciones registradas en esta sesión: **{con_obs}**.")

    exportar_excel(editada)
    return editada


# ----------------------------------------------------------------------------
# Exportaciones
# ----------------------------------------------------------------------------
def exportar_excel(vista: pd.DataFrame):
    c1, c2 = st.columns(2)
    csv = vista.to_csv(index=False).encode("utf-8-sig")
    c1.download_button(
        "⬇️ Descargar CSV (con observaciones)",
        data=csv,
        file_name="novedades_mt3.csv",
        mime="text/csv",
        width="stretch",
    )
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        vista.to_excel(writer, index=False, sheet_name="Novedades")
    c2.download_button(
        "⬇️ Descargar Excel (con observaciones)",
        data=buf.getvalue(),
        file_name="novedades_mt3.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )


# ----------------------------------------------------------------------------
# Helper numpy-safe
# ----------------------------------------------------------------------------
def np_where(cond: pd.Series, si: str, no: str) -> pd.Series:
    return pd.Series(si, index=cond.index).where(cond, no)


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
def main():
    try:
        df = cargar_datos()
    except FileNotFoundError as exc:
        st.error(f"❌ {exc}")
        st.stop()
    except Exception as exc:  # noqa: BLE001 - capturar errores de lectura
        st.error(f"❌ No se pudo cargar el archivo Excel: {exc}")
        st.stop()

    estados_crono = cargar_estados_crono()

    # --- Campos dashboard: Estado de la sede, Categoría Novedad (Y) y novedades ---
    kpi_campos, nov_campos = cargar_campos()
    kpi_norm, nov_norm = preparar_campos(kpi_campos, nov_campos, df)
    crono_ups_raw = cargar_crono_ups()
    crono_ups = preparar_crono_ups(crono_ups_raw)
    fecha_corte = (nov_norm["_FECHA"].max() if (nov_norm is not None and not nov_norm.empty)
                   else None)

    # Estado por SBAN: prioriza la fuente diaria (Dashboard_KPI, col N) y usa el
    # cronograma como respaldo.
    estados_sede = {}
    if kpi_norm is not None and not kpi_norm.empty and "Estado sede" in kpi_norm.columns:
        for _, fila_k in kpi_norm.iterrows():
            est = str(fila_k.get("Estado sede", "") or "").strip()
            if est:
                estados_sede[str(fila_k["_SBAN"])] = est

    def _estado_sban(v):
        if pd.isna(v):
            return None
        sban = _pad5(v)
        return estados_sede.get(sban) or estados_crono.get(int(v))

    df["_est_crono"] = df["SBAN"].map(_estado_sban).fillna("Sin cronograma")

    df["_componente"] = clasificar_componente(df)

    seleccion, solo_pendientes, f_attr = render_sidebar(df)

    # --- Selector de módulo (pestañas) ---
    opciones_mod = ["Componente 1", "Componente 2 (Impresoras láser)", "UPS"]
    if hasattr(st, "segmented_control"):
        modulo = st.segmented_control("Módulo", opciones_mod,
                                      default=opciones_mod[0], key="modulo")
    else:
        modulo = st.radio("Módulo", opciones_mod, horizontal=True, key="modulo")
    modulo = modulo or opciones_mod[0]
    cod_mod = {"Componente 1": "C1", "Componente 2 (Impresoras láser)": "C2",
               "UPS": "UPS"}[modulo]

    # UPS: estado de sede derivado del cronograma UPS (fecha + avance en MT)
    if cod_mod == "UPS" and crono_ups is not None and not crono_ups.empty:
        est_ups = estados_ups(df, crono_ups)
        df["_est_crono"] = df["_SBAN"].map(est_ups).fillna("Sin cronograma")

    df_mod = df[df["_componente"].eq(cod_mod)].copy()
    conteos = df["_componente"].value_counts().to_dict()

    # Novedades aisladas por componente (no se trasladan entre módulos)
    def _novedades_de(comp):
        if kpi_norm is not None and "_comp" in kpi_norm.columns:
            kc = kpi_norm[kpi_norm["_comp"].eq(comp)]
        else:
            kc = kpi_norm.iloc[0:0] if kpi_norm is not None else None
        if nov_norm is not None and "_comp" in nov_norm.columns:
            nc = nov_norm[nov_norm["_comp"].eq(comp)]
        else:
            nc = nov_norm.iloc[0:0] if nov_norm is not None else None
        return resumen_novedades(kc, nc)

    nov_resumen = _novedades_de(cod_mod)

    # --- Filtrado dinámico: oficina → atribución (Facturable) → solo pendientes ---
    mascara_office = pd.Series(True, index=df_mod.index)
    if seleccion:
        mascara_office &= df_mod["_ofi_key"].isin(seleccion)
    base_oficina = df_mod[mascara_office]
    if f_attr != "T":
        base_oficina = base_oficina[base_oficina["Facturable"] == f_attr]

    filtrado = base_oficina
    if solo_pendientes:
        filtrado = base_oficina[base_oficina["_pendiente"]].copy()

    if filtrado.empty:
        st.warning("Sin datos para los módulos/filtros seleccionados.")
        st.stop()

    # --- Título / KPIs ---
    st.title("🛠️ Control Mantenimiento Preventivo 3")
    scope = "Todas las oficinas" if not seleccion else f"{len(seleccion)} oficina(s)"
    atrib = {"T": "Todos", "Si": "BANCO (Facturable Si)", "No": "COLSOF (No facturable)"}[f_attr]
    st.markdown(
        f":blue-badge[**Módulo: {modulo}**] "
        f":green-badge[🟢 Operativo] "
        f":gray-badge[🕒 Última actualización: {leer_ultima_actualizacion()}]"
    )
    st.caption(
        f"Vista: **{scope}** · Atribución: **{atrib}** · "
        f"{len(filtrado):,} equipos".replace(",", ".")
    )
    st.caption(
        f"Totales base: Componente 1 = {conteos.get('C1', 0):,} · "
        f"Componente 2 (láser) = {conteos.get('C2', 0):,} · UPS = {conteos.get('UPS', 0):,}".replace(",", ".")
        + (f" · Novedades del día: **{int(nov_resumen['Novedades del día'].sum())}** "
           f"(corte {fecha_corte:%Y-%m-%d})"
           if (nov_resumen is not None and not nov_resumen.empty
               and 'Novedades del día' in nov_resumen.columns and pd.notna(fecha_corte))
           else " · Novedades: sin datos")
    )
    banner_publico()

    # --- Validación de integridad de los archivos en uso ---
    p_crono_val = ruta_crono_act()
    if p_crono_val.exists():
        criticos, avisos = validar_integridad(
            df, _cronograma_validacion(str(p_crono_val), os.path.getmtime(p_crono_val)),
            kpi_campos, nov_campos, crono_ups_raw)
    else:
        criticos, avisos = validar_integridad(df, None, kpi_campos, nov_campos, crono_ups_raw)
        avisos.insert(0, "⚠️ Cronograma no disponible: el estado de la sede se toma de "
                         "Campos dashboard (col N).")
    if criticos:
        for msg in criticos:
            st.error(msg)
        st.stop()
    if avisos:
        with st.expander(f"🔍 Validación de integridad — {len(avisos)} aviso(s)",
                         expanded=False):
            for msg in avisos:
                st.warning(msg)

    render_kpis(filtrado)
    avance_filtro = float(filtrado["_mt"].mean() * 100) if len(filtrado) else 0.0
    st.progress(min(max(avance_filtro / 100.0, 0.0), 1.0),
                text=f"Avance MT3 del filtro actual: {avance_filtro:.1f}%".replace(".", ","))

    # --- Gráficas ---
    col_izq, col_der = st.columns([1, 1.4])
    with col_izq:
        st.plotly_chart(
            grafico_dona(int(filtrado["_mt"].sum()), int(filtrado["_pendiente"].sum())),
            width="stretch",
        )
    with col_der:
        st.plotly_chart(grafico_top_pendientes(filtrado, top_n=15), width="stretch")
    st.plotly_chart(grafico_barras(filtrado), width="stretch")

    # --- Resumen por oficina (Total / Subsanados / Pendientes / % Avance / Novedades) ---
    render_resumen(base_oficina, seleccion, f_attr, nov_resumen)

    # --- Tabla de gestión ---
    render_tabla(filtrado, seleccion, solo_pendientes)

    st.markdown("---")
    st.caption(f"Mantenimiento Preventivo 3 · Módulo {modulo} · 1 fila = 1 elemento (Serial único).")


if __name__ == "__main__":
    main()
