# -*- coding: utf-8 -*-
"""
Dashboard Mantenimiento Preventivo 3 (MT3) - Banco Agrario / COLSOF
App Streamlit de un solo archivo. Fuente: Data_PCTriage.xlsx (hoja 'Hoja1').
"""
import hashlib
import io
import os
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ----------------------------------------------------------------------------
# Zona horaria oficial: América/Bogotá (UTC-5) — Bogotá, Lima, Quito
# ----------------------------------------------------------------------------
try:
    import zoneinfo
    TZ_BO = zoneinfo.ZoneInfo("America/Bogota")
except Exception:  # fallback si no está disponible zoneinfo
    TZ_BO = timezone.utc  # se usa UTC como respaldo, pero se nota en logs

def ahora() -> pd.Timestamp:
    """Timestamp actual en zona horaria de Bogotá (UTC-5)."""
    return pd.Timestamp.now(tz=TZ_BO).tz_convert(None)

def hoy() -> pd.Timestamp:
    """Fecha actual (medianoche) en zona horaria de Bogotá."""
    return ahora().normalize()

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

# Regla de negocio: las impresoras láser (Componente 2) SE FACTURAN al Banco.
# ⚠️ ACTIVADA por decisión del negocio (COLSOF). El Excel de origen (Data) las trae como
# 'No', así que el panel aplica esta regla y la anuncia con un chip en el encabezado.
IMPRESORAS_FACTURABLES = True

# Modo administrador: habilita la carga de datos (módulo "Ingesta diaria").
# En la app PÚBLICA (Streamlit Cloud) NO se define este secreto → el público solo consulta.
# Para activarlo en tu PC, crea `.streamlit/secrets.toml` con:   modo_admin = true
# (ese archivo NO se sube al repositorio).
try:
    MODO_ADMIN = bool(st.secrets.get("modo_admin", False))
except Exception:  # noqa: BLE001
    MODO_ADMIN = False


# ----------------------------------------------------------------------------
# Tema visual: OSCURO premium (por defecto) y CLARO institucional
# El usuario alterna con el toggle de la barra superior derecha.
# ----------------------------------------------------------------------------
TEMAS = {
    "oscuro": {
        "bg": "#0A0E14", "bg_soft": "#0F1522", "card": "#141C2D",
        "border": "rgba(255,255,255,.08)", "ink": "#E6EAF2", "muted": "#7C8AA5",
        "grid": "rgba(255,255,255,.08)", "primario": "#22D3EE", "btn_prim_ink": "#04121A",
        "verde": "#34D399", "rojo": "#FB7185", "ambar": "#FBBF24", "morado": "#A78BFA",
        "celeste": "#22D3EE", "sombra": "0 12px 30px rgba(0,0,0,.35)",
        "pend": "#FBBF24", "pista": "rgba(255,255,255,.08)",
        "heat": ["#2B1B0E", "#8C4A0F", "#E8850F", "#FFD166"], "heat_border": "#0A0E14",
    },
    "claro": {
        "bg": "#F2F5F9", "bg_soft": "#FFFFFF", "card": "#FFFFFF",
        "border": "#E3E8EF", "ink": "#1B2430", "muted": "#64748B",
        "grid": "#EDF1F6", "primario": "#1F4E78", "btn_prim_ink": "#FFFFFF",
        "verde": "#2E9E5B", "rojo": "#D64541", "ambar": "#D97706", "morado": "#6D28D9",
        "celeste": "#2E75B6", "sombra": "0 1px 3px rgba(16,24,40,.08)",
        "pend": "#EAB308", "pista": "#EDF1F6",
        "heat": ["#FFF7E6", "#FEE391", "#FE9929", "#D7301F"], "heat_border": "#FFFFFF",
    },
}

_CSS_TEMA = '''
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, .stApp, [class*="css"] { font-family: 'Inter', system-ui, sans-serif !important; }
.stApp { background: __BG__ !important; }
[data-testid="stHeader"] { background: transparent !important; }
[data-testid="stSidebar"] { background: __BGSOFT__ !important; border-right: 1px solid __BORDER__ !important; }
[data-testid="stSidebar"] p, [data-testid="stSidebar"] span, [data-testid="stSidebar"] label,
[data-testid="stSidebar"] li, [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 { color: __INK__ !important; }
h1, h2, h3, h4 { color: __INK__ !important; font-weight: 800 !important; letter-spacing: -.4px; }
p, li, label { color: __INK__ !important; }
[data-testid="stCaptionContainer"] p, .stCaption, small { color: __MUTED__ !important; }
/* --- Tarjetas KPI (detalle y resumen): mismo lenguaje visual que las destacadas --- */
[data-testid="stMetric"] { background: __CARD__ !important; border: 1px solid __BORDER__ !important;
  border-radius: 16px; padding: 14px 15px 13px; box-shadow: __SOMBRA__; }
[data-testid="stMetricValue"] { color: __INK__ !important; font-weight: 800 !important;
  font-size: 1.45rem !important; letter-spacing: -1.1px;
  white-space: nowrap !important; overflow: visible !important; text-overflow: clip !important; }
[data-testid="stMetricValue"] > div, [data-testid="stMetricValue"] > span {
  white-space: nowrap !important; overflow: visible !important; text-overflow: clip !important; }
[data-testid="stMetricLabel"] p { color: __MUTED__ !important; font-weight: 700 !important;
  font-size: .70rem !important; letter-spacing: .7px !important; text-transform: uppercase;
  white-space: normal !important; overflow: visible !important;
  text-overflow: clip !important; line-height: 1.2 !important; }
.stButton > button { border-radius: 10px !important; border: 1px solid __BORDER__ !important;
  background: __CARD__ !important; color: __INK__ !important; font-weight: 600 !important; }
.stButton > button p { color: inherit !important; }
.stButton > button[kind="primary"], .stButton > button[data-testid="baseButton-primary"] {
  background: __PRIMARY__ !important; border-color: __PRIMARY__ !important; }
.stButton > button[kind="primary"] p, .stButton > button[data-testid="baseButton-primary"] p {
  color: __BTNINK__ !important; }
[data-testid="stExpander"] { border: 1px solid __BORDER__ !important; border-radius: 12px !important;
  background: __CARD__ !important; }
[data-testid="stExpander"] summary p { color: __INK__ !important; font-weight: 600; }
[data-testid="stDataFrame"], [data-testid="stDataEditor"] { border: 1px solid __BORDER__ !important;
  border-radius: 12px !important; }
[data-testid="stAlert"] { border-radius: 12px !important; }
hr { border-color: __BORDER__ !important; }
.stProgress > div > div > div > div { background: __VERDE__ !important; }
[data-baseweb="select"] > div, [data-baseweb="input"] > div { background: __CARD__ !important;
  border-color: __BORDER__ !important; }
[data-baseweb="select"] div, [data-baseweb="input"] input { color: __INK__ !important; }
[data-testid="stFileUploaderDropzone"] { background: __CARD__ !important; border-color: __BORDER__ !important; }
/* Controles tipo botón (segmented control / pills / uploader) en ambos temas */
[data-testid="stBaseButton-secondary"], [data-testid="stBaseButton-primary"],
[data-testid="stBaseButton-segmented_control"], [data-testid="stBaseButton-pills"] {
  background: __CARD__ !important; color: __INK__ !important; border-color: __BORDER__ !important; }
[data-testid="stBaseButton-segmented_control"] p, [data-testid="stBaseButton-pills"] p,
[data-testid="stBaseButton-secondary"] p { color: __INK__ !important; }
[data-testid="stBaseButton-segmented_control"][aria-checked="true"],
[data-testid="stBaseButton-pills"][aria-checked="true"],
[data-testid="stBaseButton-secondary"][aria-pressed="true"] {
  background: __PRIMARY__ !important; border-color: __PRIMARY__ !important; }
[data-testid="stBaseButton-segmented_control"][aria-checked="true"] p,
[data-testid="stBaseButton-pills"][aria-checked="true"] p { color: __BTNINK__ !important; }
[data-testid="stSegmentedControl"] { background: transparent !important; }
[data-testid="stSegmentedControl"] > div { background: __CARD__ !important; border-radius: 12px; }
/* Segmented control / pills (selector de módulo, ranking, etc.) — se estilizan por rol ARIA */
button[role="radio"] { background: __CARD__ !important; color: __INK__ !important;
  border: 1px solid __BORDER__ !important; }
button[role="radio"] p, button[role="radio"] span { color: __INK__ !important; font-weight: 600 !important; }
button[role="radio"][aria-checked="true"] { background: __PRIMARY__ !important;
  border-color: __PRIMARY__ !important; }
button[role="radio"][aria-checked="true"] p, button[role="radio"][aria-checked="true"] span {
  color: __BTNINK__ !important; }
/* Pestañas (Gráficos / Resumen / Gestión) */
[data-testid="stTabs"] button[role="tab"] { border-radius: 10px 10px 0 0 !important; }
[data-testid="stTabs"] button[role="tab"] p { color: __MUTED__ !important; font-weight: 600 !important;
  font-size: .95rem !important; }
[data-testid="stTabs"] button[role="tab"][aria-selected="true"] { background: __CARD__ !important; }
[data-testid="stTabs"] button[role="tab"][aria-selected="true"] p { color: __INK__ !important; }
[data-testid="stTabs"] [data-baseweb="tab-highlight"] { background: __PRIMARY__ !important; }
[data-testid="stTabs"] [data-baseweb="tab-border"] { background: __BORDER__ !important; }
[data-testid="stFileUploader"] button { background: __CARD__ !important; color: __INK__ !important;
  border: 1px solid __BORDER__ !important; border-radius: 10px !important; }
[data-testid="stFileUploader"] button * { color: __INK__ !important; }
[data-testid="stFileUploader"] span, [data-testid="stFileUploader"] small { color: __MUTED__ !important; }
[data-testid="stToolbar"] button, [data-testid="stToolbar"] span, [data-testid="stToolbar"] a {
  color: __MUTED__ !important; }
[data-testid="stMainMenu"] { color: __MUTED__ !important; }

/* ===== KPIs destacados (oficinas · avance · equipos impactados) ===== */
.kd { background: __CARD__ !important; border: 1px solid __BORDER__ !important;
  border-radius: 16px; padding: 15px 18px 14px; box-shadow: __SOMBRA__; height: 100%; }
.kd-lbl { font-size: 11px; font-weight: 700; letter-spacing: .7px; text-transform: uppercase;
  color: __MUTED__ !important; }
.kd-num { font-size: 40px; font-weight: 800; letter-spacing: -1.8px; color: __INK__ !important;
  line-height: 1.05; margin-top: 6px; }
.kd-sub { font-size: 11.5px; color: __MUTED__ !important; margin-top: 6px; }
.kd-bar { height: 7px; border-radius: 99px; background: __BORDER__; margin-top: 10px; overflow: hidden; }
.kd-bar > i { display: block; height: 100%; border-radius: 99px; }

/* ================= RESPONSIVE (tablet y móvil) ================= */
@media (max-width: 1000px) {
  [data-testid="stHorizontalBlock"] { flex-wrap: wrap !important; gap: .55rem !important; }
  [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] { flex: 1 1 46% !important; min-width: 46% !important; }
  [data-testid="stMainBlockContainer"] { padding: 1rem 1rem 2.5rem !important; }
}
@media (max-width: 680px) {
  [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] { flex: 1 1 100% !important; min-width: 100% !important; }
  #mt3-hero-titulo { font-size: 19px !important; line-height: 1.2 !important; }
  #mt3-hero-sub { font-size: 12px !important; }
  [data-testid="stMetricValue"] { font-size: 1.45rem !important; }
  .kd-num { font-size: 30px !important; letter-spacing: -1px !important; }
  [data-testid="stMetric"] { padding: 11px 13px !important; }
  [data-testid="stMainBlockContainer"] { padding: .8rem .7rem 2.5rem !important; }
  p, li, label { font-size: .93rem !important; }
  h1, h2, h3 { font-size: 1.05rem !important; }
  [data-testid="stSidebar"] { min-width: 265px !important; }
  .stButton > button, [data-testid="stBaseButton-secondary"], [data-testid="stBaseButton-primary"] {
    min-height: 40px !important; }
}
</style>
'''


def tema_actual() -> dict:
    """Tokens del tema activo (oscuro por defecto)."""
    if "tema" not in st.session_state:
        st.session_state["tema"] = "oscuro"
    return TEMAS.get(st.session_state["tema"], TEMAS["oscuro"])


def es_oscuro() -> bool:
    return st.session_state.get("tema", "oscuro") == "oscuro"


def aplicar_css():
    """Inyecta el CSS del tema activo (fondos, tarjetas, tipografía, controles)."""
    t = tema_actual()
    css = (_CSS_TEMA
           .replace("__BG__", t["bg"]).replace("__BGSOFT__", t["bg_soft"])
           .replace("__CARD__", t["card"]).replace("__BORDER__", t["border"])
           .replace("__INK__", t["ink"]).replace("__MUTED__", t["muted"])
           .replace("__PRIMARY__", t["primario"]).replace("__VERDE__", t["verde"])
           .replace("__SOMBRA__", t["sombra"]).replace("__BTNINK__", t["btn_prim_ink"]))
    st.markdown(css, unsafe_allow_html=True)


def _guardar_tema(nuevo: str):
    """Callback: fija el tema antes de que se dibujen los botones.

    Se usa con `on_click` porque escribir en `st.session_state` DESPUÉS de crear el
    widget lanza StreamlitWidgetAlreadyInstantiatedError.
    """
    st.session_state["tema"] = nuevo


def barra_tema():
    """Toggle de vista (oscura/clara) — píldoras arriba a la derecha."""
    _vacio, c_osc, c_cla = st.columns([6.6, 1.5, 1.5], vertical_alignment="center")
    with c_osc:
        st.button("🌙 Oscuro", key="btn_tema_osc", width="stretch",
                  type="primary" if es_oscuro() else "secondary",
                  on_click=_guardar_tema, args=("oscuro",))
    with c_cla:
        st.button("☀️ Claro", key="btn_tema_cla", width="stretch",
                  type="secondary" if es_oscuro() else "primary",
                  on_click=_guardar_tema, args=("claro",))


def plotly_base() -> dict:
    """Colores base de los gráficos según el tema activo (Plotly afinado)."""
    t = tema_actual()
    return dict(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, system-ui, sans-serif", color=t["ink"], size=12),
        colorway=[t["celeste"], t["verde"], t["ambar"], t["rojo"], t["morado"]],
    )


def dias_desde_actualizacion():
    """Días transcurridos desde la última actualización de datos (None si no se sabe)."""
    try:
        f = pd.to_datetime(leer_ultima_actualizacion(), errors="coerce")
        if pd.isna(f):
            return None
        return int((hoy() - f.normalize()).days)
    except Exception:  # noqa: BLE001
        return None


def encabezado_hero(modulo: str, scope: str, atrib: str, n_equipos: int,
                    conteos: dict, nov_resumen=None, fecha_corte=None):
    """Encabezado tipo 'hero' con chips de contexto (estilo Propuesta 2)."""
    t = tema_actual()

    def chip(txt: str, tono: str = "neutro") -> str:
        colores = {
            "neutro": (f"{t['card']}", f"{t['border']}", f"{t['muted']}"),
            "ok": (f"{t['verde']}22", f"{t['verde']}55", f"{t['verde']}"),
            "aviso": (f"{t['ambar']}22", f"{t['ambar']}55", f"{t['ambar']}"),
            "prim": (f"{t['primario']}22", f"{t['primario']}55", f"{t['primario']}"),
        }[tono]
        return (f'<span style="display:inline-flex;align-items:center;gap:6px;font-size:12px;font-weight:600;'
                f'padding:5px 11px;border-radius:999px;background:{colores[0]};border:1px solid {colores[1]};'
                f'color:{colores[2]};white-space:nowrap">{txt}</span>')

    nov_txt = ""
    if (nov_resumen is not None and not nov_resumen.empty
            and "Novedades del día" in nov_resumen.columns and pd.notna(fecha_corte)):
        nov_txt = (f" · Novedades del día: <b>{int(nov_resumen['Novedades del día'].sum())}</b> "
                   f"(corte {fecha_corte:%d/%m/%Y})")

    _dias = dias_desde_actualizacion()
    _chip_viejo = (chip(f"⚠️ Datos de hace {_dias} días", "aviso")
                   if (_dias is not None and _dias >= 3) else "")
    _chip_regla = (chip("ℹ️ Impresoras: facturables (Si)", "aviso")
                   if IMPRESORAS_FACTURABLES else "")
    _chip_fact = ""
    if str(atrib).startswith("BANCO"):
        _chip_fact = chip("🏦 Viendo: Solo facturables", "prim")
    elif str(atrib).startswith("COLSOF"):
        _chip_fact = chip("🏢 Viendo: Solo no facturables", "prim")

    st.markdown(
        f'<div style="display:flex;justify-content:space-between;align-items:flex-end;gap:16px;'
        f'flex-wrap:wrap;padding:2px 0 14px;border-bottom:1px solid {t["border"]};margin:4px 0 12px">'
        f'  <div>'
        f'    <div id="mt3-hero-titulo" style="font-size:26px;font-weight:800;letter-spacing:-.7px;'
        f'color:{t["ink"]};line-height:1.15">🛠️ Control Mantenimiento Preventivo 3</div>'
        f'    <div id="mt3-hero-sub" style="font-size:13px;color:{t["muted"]};margin-top:5px">'
        f'Vista: <b>{scope}</b> · Atribución: <b>{atrib}</b> · <b>{f"{n_equipos:,}".replace(",", ".")}</b> equipos'
        f'    </div>'
        f'  </div>'
        f'  <div style="display:flex;gap:8px;flex-wrap:wrap;justify-content:flex-end">'
        f'    {chip("📦 " + modulo, "prim")}{chip("Operativo", "ok")}'
        f'    {chip("🕒 " + str(leer_ultima_actualizacion()))}'
        + _chip_regla
        + _chip_fact
        + _chip_viejo
        + (chip("🔒 Modo público", "aviso") if MODO_PUBLICO else "")
        + f'  </div>'
        f'</div>'
        f'<div style="font-size:12px;color:{t["muted"]};margin:-4px 0 10px">'
        f'Totales base: Componente 1 = {conteos.get("C1", 0):,} · '
        f'Componente 2 (láser) = {conteos.get("C2", 0):,} · UPS = {conteos.get("UPS", 0):,}'
        f'{nov_txt}</div>'.replace(",", "."),
        unsafe_allow_html=True,
    )


def _reset_filtros_por_modulo():
    """on_change del selector de módulo: reinicia los filtros que lo dejarían vacío.

    Se ejecuta como callback (antes del rerun), así que puede limpiar claves de widgets
    sin riesgo de StreamlitWidgetAlreadyInstantiatedError.
    """
    st.session_state.pop("seg_fact", None)   # el selector de facturación se reinicia
    st.session_state.pop("click_sban", None)
    for clave, valor in VALORES_FILTRO.items():
        if clave in ("seg_fact", "click_sban"):
            continue
        if valor is None:
            st.session_state.pop(clave, None)
        else:
            st.session_state[clave] = valor


def f_attr_seleccionado() -> str:
    """Traduce la selección del selector VISIBLE de facturación a 'T' / 'Si' / 'No'."""
    sel = st.session_state.get("seg_fact") or ""
    if "Facturables (BANCO)" in sel:
        return "Si"
    if "No facturables" in sel:
        return "No"
    return "T"


def barra_facturacion(df_mod: pd.DataFrame, mostrar_conteos: bool = True):
    """Barra VISIBLE de facturación (Opción 1 de usabilidad).

    Píldoras siempre a la vista en el área principal (no escondidas en el sidebar),
    con la opción activa resaltada. En módulos de una sola categoría se informa con
    una etiqueta en lugar de ofrecer opciones inútiles.

    `mostrar_conteos=False` cuando las tarjetas KPI ya muestran esos conteos (evita
    repetir la misma cifra dos veces en pantalla).
    """
    if df_mod is None or getattr(df_mod, "empty", True) or "Facturable" not in df_mod.columns:
        return
    n_tot = int(len(df_mod))
    n_si = int((df_mod["Facturable"] == "Si").sum())
    n_no = int((df_mod["Facturable"] == "No").sum())

    def fmt(n: int) -> str:
        return f"{n:,}".replace(",", ".")

    if n_si == 0 or n_no == 0:
        st.session_state.pop("seg_fact", None)   # nunca deja la vista vacía
        cual = "Facturable (BANCO)" if n_si else "No facturable (COLSOF)"
        extra = f" — {fmt(n_tot)} equipos" if mostrar_conteos else ""
        st.info(f"ℹ️ Este módulo es **100% {cual}**{extra}. No requiere filtro.")
        return

    st.segmented_control(
        "Facturación",
        ["Todos", "🏦 Facturables (BANCO)", "🏢 No facturables (COLSOF)"],
        default="Todos", key="seg_fact",
        help="Filtra el tablero por atribución de facturación (columna 'Facturable').",
    )
    if mostrar_conteos:
        st.caption(f"📊 Del módulo ({fmt(n_tot)} equipos): 🏦 Facturables {fmt(n_si)} · "
                   f"🏢 No facturables {fmt(n_no)}")


VALORES_FILTRO = {
    # valor por defecto de cada filtro (clave de session_state)
    "f_oficinas": [],
    "f_solo_pend": False,
    "f_atrib": "Todos",
    "click_sban": None,
    "seg_fact": None,
    "res_estado": None,
    "res_catnov": None,
    "res_min_av": 0,
    "res_solo_pend": False,
    "res_solo_nov": False,
}


def _quitar_filtro(clave: str):
    """Callback: devuelve un filtro a su valor por defecto y recarga.

    Siempre por callback (`on_click`) para no tocar `st.session_state` después de
    crear los widgets (causa de StreamlitWidgetAlreadyInstantiatedError).
    """
    valor = VALORES_FILTRO.get(clave)
    if valor is None:
        st.session_state.pop(clave, None)
    else:
        st.session_state[clave] = valor
    # La atribución se maneja con el selector visible 'seg_fact'; si se quita el chip,
    # hay que limpiar también ese selector o seguiría filtrando al usuario.
    if clave == "f_atrib":
        st.session_state.pop("seg_fact", None)


def _limpiar_todos_los_filtros():
    """Callback del botón «Limpiar todos los filtros»."""
    for clave, valor in VALORES_FILTRO.items():
        if valor is None:
            st.session_state.pop(clave, None)
        else:
            st.session_state[clave] = valor


def _limpiar_filtros_resumen():
    """Callback del botón «Restablecer filtros del resumen»."""
    for clave in ("res_estado", "res_catnov", "res_min_av",
                  "res_solo_pend", "res_solo_nov"):
        valor = VALORES_FILTRO.get(clave)
        if valor is None:
            st.session_state.pop(clave, None)
        else:
            st.session_state[clave] = valor


def barra_filtros_activos(seleccion, solo_pendientes, f_attr, click_sban):
    """Chips de filtros activos, cada uno con botón para quitarlo de un toque."""
    activos = []
    if seleccion:
        if len(seleccion) == 1:
            txt = seleccion[0][:26] + ("…" if len(seleccion[0]) > 26 else "")
        else:
            txt = f"{len(seleccion)} oficinas"
        activos.append(("🏢 " + txt, "f_oficinas"))
    if solo_pendientes:
        activos.append(("⚠️ Solo pendientes", "f_solo_pend"))
    if f_attr != "T":
        activos.append(("🏦 BANCO" if f_attr == "Si" else "🏢 COLSOF", "f_atrib"))
    if click_sban:
        activos.append((f"🎯 SBAN {click_sban}", "click_sban"))
    if not activos:
        return
    cols = st.columns([1.05] + [1.15] * len(activos), vertical_alignment="center")
    with cols[0]:
        st.caption("Filtros activos:")
    for col, (etiqueta, clave) in zip(cols[1:], activos):
        with col:
            st.button("✖ " + etiqueta, key=f"quitar_{clave}", width="stretch",
                      on_click=_quitar_filtro, args=(clave,))


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


def _es_xlsx_util(p: Path) -> bool:
    """Descarta los temporales que Excel deja abiertos (`~$archivo.xlsx`)."""
    return not p.name.startswith("~$")


def _archivo_mas_reciente(carpeta: Path, tipo: str) -> Path | None:
    """Devuelve el .xlsx más reciente del tipo dado en la carpeta."""
    if not carpeta.exists():
        return None
    archivos = [p for p in carpeta.glob("*.xlsx")
                if _es_xlsx_util(p) and _tipo_archivo(p) == tipo]
    if not archivos:
        return None
    return max(archivos, key=lambda p: p.stat().st_mtime)


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


@st.cache_data(show_spinner="Descubriendo ingesta diaria…")
def _descubrir_ingesta(clave: str, mtime: float, carpetas: tuple) -> dict:
    """Toma el .xlsx más reciente de cada tipo dentro de las carpetas de ingesta."""
    out = {}
    for cp in carpetas:
        carpeta = Path(cp)
        if not carpeta.exists():
            continue
        for tipo in ("data", "crono", "crono_ups", "campos"):
            arch = _archivo_mas_reciente(carpeta, tipo)
            if arch and tipo not in out:
                out[tipo] = str(arch)
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
        return pd.Timestamp.fromtimestamp(max(marcas), tz=TZ_BO).tz_convert(None).strftime("%Y-%m-%d %H:%M")
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
    """¿El archivo en uso es LOCAL del proyecto (no una carga subida por la web)?

    Locales: carpeta `data/`, carpeta 'Ingesta de datos diaria' y las rutas canónicas.
    NO local: cualquier archivo dentro de `uploads/` (subido desde el panel).
    """
    try:
        r = p.resolve()
    except Exception:  # noqa: BLE001
        return False
    cargas = str(DIR_CARGA.resolve()).lower()
    rl = str(r).lower()
    if rl == cargas or rl.startswith(cargas + os.sep):
        return False                                    # proviene de uploads/
    if r in (RUTA_XLSX.resolve(), RUTA_CRONO.resolve()):
        return True
    for carpeta in (CARPETA_DATA, CARPETA_INGESTA):
        try:
            base = str(carpeta.resolve()).lower()
        except Exception:  # noqa: BLE001
            continue
        if rl == base or rl.startswith(base + os.sep):
            return True
    return False


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
        w.writerow([ahora().strftime("%Y-%m-%d %H:%M:%S"),
                    tipo, archivo_original, destino.name, sha, filas])


def _guardar_subida(uploaded, tipo: str, sufijo: str) -> Path:
    destino = DIR_CARGA / f"{ahora():%Y%m%d_%H%M%S}_{sufijo}.xlsx"
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
    hoy_local = hoy()
    out = {}
    for _, r in cup.iterrows():
        s = r["_SBAN"]
        fecha = r["Fecha plan"]
        tot = int(agg.loc[s, "Total"]) if s in agg.index else 0
        mt = int(agg.loc[s, "MT"]) if s in agg.index else 0
        if pd.notna(fecha) and hoy_local < fecha:
            out[s] = "Programada"
        elif tot > 0 and mt >= tot:
            out[s] = "Finalizada"
        elif pd.notna(fecha) and fecha <= hoy_local:
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
    """Busca una columna tolerando espacios (dobles/sobrantes) y mayúsculas."""
    def norm(x):
        return " ".join(str(x).strip().lower().split())
    mapa = {norm(c): c for c in df.columns}
    for nom in nombres:
        if nom in df.columns:
            return nom
        clave = norm(nom)
        if clave in mapa:
            return mapa[clave]
    return None


_PRIO_UPS = ["Finalizado", "Finalizada", "En proceso", "En ejecución", "Programado",
             "Programada", "No aplica", "Pendiente"]
# Prioridad del ESTADO DE LA SEDE (col N). Regla del negocio: una sede con varias
# filas (Regional / Jefatura / Oficina) se resume por la de MAYOR prioridad.
_PRIO_SEDE = ["Finalizada", "Reprogramada_Finalizada", "En proceso", "Programada"]


def _mejor_por_sban(valores, prioridad):
    """Resume varias filas por SBAN quedándose con el estado de MAYOR prioridad.

    `prioridad` es la lista ordenada de mayor a menor; lo que no esté en ella queda
    al final. A diferencia de «gana la última fila», esto no depende del orden del
    archivo: la sede conserva siempre el estado más avanzado.
    """
    idx = {v: i for i, v in enumerate(prioridad)}
    mejor: dict = {}
    for sban, valor in valores:
        v = str(valor).strip()
        if not v or v.lower() in ("nan", "none"):
            continue
        actual = mejor.get(sban)
        if actual is None or idx.get(v, len(prioridad)) < idx.get(actual, len(prioridad)):
            mejor[sban] = v
    return mejor


def preparar_campos(kpi, nov, df=None):
    """Normaliza Dashboard_KPI (estado de la sede, ESTADO UPS, observaciones) y Novedades Equipos (por SBAN).

    Las novedades viven en la hoja 'Novedades Equipos' e incluyen su propia
    columna 'Categoria Novedad' (ya no viene en Dashboard_KPI).
    El avance de UPS se toma TAL CUAL de la columna AP «ESTADO UPS» de
    Dashboard_KPI (Banco Agrario + COLSOF en el mismo documento): no se deriva.
    Amarre por SBAN (5 dígitos) con respaldo en SBAN PCT.
    Asigna `_comp` (C1/C2/UPS) para que las novedades no se trasladen entre módulos.
    """
    k = pd.DataFrame(columns=["_SBAN", "Estado sede", "Estado UPS", "Obs. novedad",
                              "Obs. UPS", "Fuente UPS"])
    if kpi is not None and not kpi.empty:
        est = _col(kpi, "Estado de la sede")
        ups = _col(kpi, "ESTADO UPS")
        obs = _col(kpi, "Observaciones actas PCT")
        obs_ups = _col(kpi, "OBSERVACIONES")
        sban = kpi[_col(kpi, "SBAN")] if _col(kpi, "SBAN") else pd.Series(dtype=str)
        k = pd.DataFrame({
            "_SBAN": sban.apply(_pad5),
            "Estado sede": (kpi[est].fillna("").astype(str).str.strip()
                            if est else pd.Series([""] * len(kpi))),
            "Estado UPS": (kpi[ups].fillna("").astype(str).str.strip()
                           if ups else pd.Series([""] * len(kpi))),
            "Obs. novedad": (kpi[obs].fillna("").astype(str).str.strip()
                             if obs else pd.Series([""] * len(kpi))),
            "Obs. UPS": (kpi[obs_ups].fillna("").astype(str).str.strip()
                         if obs_ups else pd.Series([""] * len(kpi))),
        })
        # Declara el origen del avance UPS: la columna AP del archivo (no se deriva).
        k["Fuente UPS"] = ("Campos dashboard · col AP «ESTADO UPS»" if ups
                           else "Campos dashboard sin columna «ESTADO UPS»")
        k = k[k["_SBAN"].astype(str).str.len() > 0]

        # Un SBAN puede tener varias filas (Regional/Jefatura/sede): se resume con la
        # MISMA regla de prioridad que usan las tarjetas de oficinas (`_mejor_por_sban`),
        # para que ninguna cifra difiera entre vistas.
        def _ultimo_lleno(serie):
            """Última observación no vacía (evita perder texto en filas secundarias)."""
            for x in reversed(list(serie)):
                if str(x).strip() and str(x).strip().lower() != "nan":
                    return str(x).strip()
            return ""

        mapa_sede = _mejor_por_sban(zip(k["_SBAN"].astype(str), k["Estado sede"]), _PRIO_SEDE)
        # El estado UPS se NORMALIZA antes de agregar, igual que en `oficinas_campos`:
        # así ambas rutas alimentan el helper con los mismos valores y no divergen.
        mapa_ups = _mejor_por_sban(
            zip(k["_SBAN"].astype(str), k["Estado UPS"].apply(normalizar_estado_ups)),
            _PRIO_UPS)

        k = k.groupby("_SBAN", as_index=False).agg(
            {"Obs. novedad": _ultimo_lleno, "Obs. UPS": _ultimo_lleno,
             "Fuente UPS": _ultimo_lleno})
        k["Estado sede"] = k["_SBAN"].map(mapa_sede).fillna("")
        k["Estado UPS"] = k["_SBAN"].map(mapa_ups).fillna("")

    n = pd.DataFrame()
    if nov is not None and not nov.empty:
        n = nov.copy()
        c_sban = _col(n, "SBAN")
        c_pct = _col(n, "SBAN PCT")
        base = (n[c_sban].astype(str) if c_sban else pd.Series([""] * len(n), index=n.index))
        alt = (n[c_pct].astype(str) if c_pct else pd.Series([""] * len(n), index=n.index))
        # SBAN es la llave principal; SBAN PCT solo si SBAN viene vacío
        llave = base.where(base.str.strip().replace("nan", "").ne(""), alt)
        n["_SBAN"] = llave.apply(_pad5)
        c_fecha = _col(n, "Fecha")
        n["_FECHA"] = pd.to_datetime(n[c_fecha], errors="coerce") if c_fecha else pd.NaT
        c_serial = _col(n, "Serial")
        n["_Serial"] = n[c_serial].astype(str).str.strip() if c_serial else ""
        c_cat = _col(n, "Categoria Novedad", "Categoría Novedad")
        n["Categoría Novedad"] = (n[c_cat].fillna("").astype(str).str.strip()
                                  if c_cat else "")
        c_obs = _col(n, "Observaciones")
        n["_Obs"] = n[c_obs].astype(str).str.strip() if c_obs else ""

        n = n[n["_SBAN"].astype(str).str.len() > 0]
        # Descartar filas vacías (sin categoría, fecha, serial ni observación)
        tiene_dato = (n["Categoría Novedad"].str.strip().ne("")
                      | n["_FECHA"].notna()
                      | n["_Serial"].str.strip().replace("nan", "").ne("")
                      | n["_Obs"].str.strip().replace("nan", "").ne(""))
        n = n[tiene_dato]

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
    """Contador por SBAN: categoría (última novedad), del día y acumulado."""
    if n is None or n.empty:
        base = (k.copy() if (k is not None and not k.empty)
                else pd.DataFrame(columns=["_SBAN", "Estado sede", "Estado UPS",
                                           "Obs. novedad", "Obs. UPS"]))
        if "Categoría Novedad" not in base.columns:
            base["Categoría Novedad"] = ""
        if "Obs. novedad" not in base.columns:
            base["Obs. novedad"] = ""
        base["Novedades del día"] = 0
        base["Novedades acumuladas"] = 0
        base["Última novedad"] = pd.NaT
        base["Fecha de corte"] = pd.NaT
        return base

    corte = n["_FECHA"].max()
    acum = n.groupby("_SBAN").size().rename("Novedades acumuladas")
    dia = n[n["_FECHA"].eq(corte)].groupby("_SBAN").size().rename("Novedades del día")
    ult = n.groupby("_SBAN")["_FECHA"].max().rename("Última novedad")

    orden = n.sort_values("_FECHA", na_position="first")
    if "Categoría Novedad" in n.columns:
        cc = orden[orden["Categoría Novedad"].astype(str).str.strip().ne("")]
        cat = cc.groupby("_SBAN")["Categoría Novedad"].last().rename("_cat_det")
    else:
        cat = pd.Series(dtype=str)
    if "Obs. novedad detalle" in n.columns:
        oo = orden[orden["Obs. novedad detalle"].astype(str).str.strip()
                   .replace("nan", "").ne("")]
        obs = oo.groupby("_SBAN")["Obs. novedad detalle"].last().rename("_obs_det")
    else:
        obs = pd.Series(dtype=str)

    res = pd.concat([s for s in (acum, dia, ult, cat, obs)
                     if isinstance(s, pd.Series) and len(s) > 0], axis=1)
    res.index.name = "_SBAN"
    res = res.reset_index()
    if k is not None and not k.empty:
        res = res.merge(k, on="_SBAN", how="outer")

    base_obs = res["Obs. novedad"].fillna("") if "Obs. novedad" in res.columns else ""
    det = res["_obs_det"].fillna("") if "_obs_det" in res.columns else ""
    res["Categoría Novedad"] = (res["_cat_det"].fillna("")
                                if "_cat_det" in res.columns else "")
    if isinstance(det, pd.Series):
        res["Obs. novedad"] = det.where(det.astype(str).str.strip().ne(""), base_obs)
    elif "Obs. novedad" not in res.columns:
        res["Obs. novedad"] = ""
    for c in ("Categoría Novedad", "Obs. novedad", "Estado sede", "Estado UPS", "Obs. UPS"):
        if c not in res.columns:
            res[c] = ""
        res[c] = res[c].fillna("")
    res = res.drop(columns=[c for c in ("_cat_det", "_obs_det") if c in res.columns])

    res["Novedades del día"] = (pd.to_numeric(res["Novedades del día"], errors="coerce")
                                .fillna(0).astype(int))
    res["Novedades acumuladas"] = (pd.to_numeric(res["Novedades acumuladas"], errors="coerce")
                                   .fillna(0).astype(int))
    res["Fecha de corte"] = corte
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
# Avance UPS: columna AP «ESTADO UPS» de Campos dashboard (fuente autoritativa)
# ----------------------------------------------------------------------------
def normalizar_estado_ups(v) -> str:
    """Lleva los textos de la columna AP a los 4 estados del panel."""
    t = str(v).strip()
    if not t or t.lower() in ("nan", "none"):
        return ""
    tl = t.lower()
    if tl.startswith("finaliz"):
        return "Finalizada"
    if tl.startswith("reprogram"):
        return "Reprogramada"
    if tl.startswith("en proceso") or tl.startswith("en ejec"):
        return "En proceso"
    if tl.startswith("program"):
        return "Programada"
    return t


def tabla_avance_ups(k, df) -> pd.DataFrame:
    """Cruza el ESTADO UPS de Campos (col AP) con el conteo de UPS de la Data.

    Columnas: SBAN · Oficina · Regional · UPS en Data · UPS con MT3 · Avance MT3 ·
    ESTADO UPS (col AP) · Observación UPS · Alerta.
    """
    vacio = pd.DataFrame(columns=["SBAN", "Oficina", "Regional", "UPS en Data",
                                  "UPS con MT3", "Avance MT3 %", "Estado UPS",
                                  "Estado UPS normalizado", "Observación UPS", "Alerta"])
    if k is None or getattr(k, "empty", True):
        return vacio
    ups = (df[df["_componente"].eq("UPS")].copy() if "_componente" in df.columns
           else df.iloc[0:0].copy())
    if "_SBAN" not in ups.columns and "SBAN" in ups.columns:
        ups["_SBAN"] = ups["SBAN"].apply(_pad5)
    if "_mt" not in ups.columns:
        ups["_mt"] = False
    if "Oficina" not in ups.columns:
        ups["Oficina"] = ""
    if len(ups):
        def _moda(s):
            v = s.dropna()
            m = v.mode() if len(v) else v
            return m.iloc[0] if len(m) else ""

        g = (ups.groupby("_SBAN")
                .agg(**{"UPS en Data": ("Serial", "size"), "UPS con MT3": ("_mt", "sum")})
                .reset_index())
        g["Oficina"] = (ups.groupby("_SBAN")["Oficina"].agg(_moda)
                        .reindex(g["_SBAN"]).values)
        reg = _col(ups, "Regional")
        g["Regional"] = (ups.groupby("_SBAN")[reg].agg(_moda).reindex(g["_SBAN"]).values
                         if reg else "")
    else:
        g = pd.DataFrame(columns=["_SBAN", "UPS en Data", "UPS con MT3", "Oficina", "Regional"])

    t = k.copy()
    t["_SBAN"] = t["_SBAN"].apply(_pad5)
    if "Estado UPS" not in t.columns:
        t["Estado UPS"] = ""
    if "Obs. UPS" not in t.columns:
        t["Obs. UPS"] = ""
    for c in ("Oficina", "Regional"):
        if c not in t.columns:
            t[c] = ""
    t = t[["_SBAN", "Oficina", "Regional", "Estado UPS", "Obs. UPS"]]
    if len(g):
        g = g.rename(columns={"Oficina": "Oficina Data", "Regional": "Regional Data"})
    res = t.merge(g, on="_SBAN", how="left")
    if "Oficina Data" in res.columns:
        ofi_k = res["Oficina"].fillna("").astype(str).str.strip()
        res["Oficina"] = ofi_k.where(ofi_k.ne(""), res["Oficina Data"].fillna(""))
        reg_k = res["Regional"].fillna("").astype(str).str.strip()
        res["Regional"] = reg_k.where(reg_k.ne(""), res["Regional Data"].fillna(""))
        res = res.drop(columns=["Oficina Data", "Regional Data"])
    elif "Oficina" not in res.columns:
        res["Oficina"] = ""
    if "Regional" not in res.columns:
        res["Regional"] = ""
    for c in ("UPS en Data", "UPS con MT3"):
        res[c] = pd.to_numeric(res.get(c), errors="coerce").fillna(0).astype(int)
    res["Avance MT3 %"] = ((res["UPS con MT3"] / res["UPS en Data"] * 100)
                           .where(res["UPS en Data"] > 0, 0.0).round(1))
    res["Estado UPS normalizado"] = res["Estado UPS"].apply(normalizar_estado_ups)

    fin = res["Estado UPS normalizado"].eq("Finalizada")
    con_datos = res["UPS en Data"] > 0
    res["Alerta"] = ""
    res.loc[fin & con_datos & res["UPS con MT3"].lt(res["UPS en Data"]), "Alerta"] = (
        "⚠️ Marca Finalizado pero hay UPS sin MT3 en la Data")
    res.loc[~fin & con_datos & res["UPS con MT3"].ge(res["UPS en Data"]), "Alerta"] = (
        "ℹ️ Todos los UPS con MT3 pero la col AP no dice Finalizado")
    res["SBAN"] = res["_SBAN"]
    res = res.rename(columns={"Obs. UPS": "Observación UPS"})
    return res[["SBAN", "Oficina", "Regional", "UPS en Data", "UPS con MT3", "Avance MT3 %",
                "Estado UPS", "Estado UPS normalizado", "Observación UPS", "Alerta"]]


def render_avance_ups(k, df, seleccion=None):
    """Panel del avance UPS con el ESTADO UPS (col AP) como fuente principal.

    Solo se listan las sedes que **tienen UPS en la Data**, para que el total, la
    cobertura y el control de MT3 hablen del mismo conjunto de oficinas.
    """
    st.markdown("#### 🔋 Avance de UPS (columna AP «ESTADO UPS» de Campos dashboard)")
    if k is None or getattr(k, "empty", True):
        st.info("Este archivo de **Campos dashboard** todavía no trae la hoja `Dashboard_KPI` "
                "con la columna **ESTADO UPS**. Cárgala para ver el avance de las UPS aquí.")
        return
    t = tabla_avance_ups(k, df)
    t = t[t["UPS en Data"] > 0]          # mismo universo que el resto de indicadores UPS
    if seleccion:
        sbans = {str(s).split(" - ")[0] for s in seleccion}
        t = t[t["SBAN"].isin(sbans)]
    if t.empty:
        st.info("Sin sedes con UPS para los filtros seleccionados.")
        return

    total = len(t)
    reportadas = int(t["Estado UPS normalizado"].astype(str).str.strip().ne("").sum())
    fin = int(t["Estado UPS normalizado"].eq("Finalizada").sum())
    proc = int(t["Estado UPS normalizado"].eq("En proceso").sum())
    prog = int(t["Estado UPS normalizado"].eq("Programada").sum())
    sin_rep = total - reportadas

    # --- Conciliación con el control de MT3 (fuente distinta) ---
    con_ups = t["UPS en Data"] > 0
    cien_mt3 = con_ups & t["UPS con MT3"].ge(t["UPS en Data"])
    fin_ap = t["Estado UPS normalizado"].eq("Finalizada")
    ambos = int((fin_ap & cien_mt3).sum())
    solo_ap = int((fin_ap & ~cien_mt3).sum())
    solo_mt3 = int((~fin_ap & cien_mt3).sum())
    n_cien = int(cien_mt3.sum())
    n_con_ups = int(con_ups.sum())

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🔋 Oficinas que ya reportaron", f"{reportadas} de {total}",
              delta=f"{sin_rep} sin dato", delta_color="off", border=True,
              help="Oficinas **con UPS en la Data** que ya traen algún valor en la columna "
                   "«ESTADO UPS» del archivo. Las demás están en blanco: la oficina todavía no "
                   "marca su avance.")
    # NOTA: los conteos de «reportado por la oficina» y «completo según la Data» ya están en
    # las tarjetas de indicadores de oficinas; aquí NO se repiten. Esta tarjeta mide la
    # CONCILIACIÓN (cuántas oficinas coinciden), que es el propósito de esta pestaña.
    c2.metric("🧩 Coinciden el reporte y la Data", f"{ambos}",
              delta=f"de {n_con_ups} oficinas con UPS", delta_color="off", border=True,
              help="Oficinas donde el reporte dice `Finalizado` Y todos sus UPS de la Data "
                   "tienen mantenimiento hecho. Los dos conteos por separado están en las "
                   "tarjetas «🔋 UPS finalizadas (reportado)» y «🏢 Oficinas con UPS completo».")
    c3.metric("🟨 En proceso · ⬜ Programadas", f"{proc} · {prog}",
              help="Conteo del estado reportado en los demás casos. " + (f"Sin dato: {sin_rep}."
                   if sin_rep else "Todas las oficinas tienen dato."), border=True)
    c4.metric("⚠️ Diferencias reporte vs Data", f"{solo_ap + solo_mt3}",
              delta_color="inverse" if (solo_ap + solo_mt3) else "off",
              border=True,
              help="Oficinas donde lo que reportó la oficina y lo que muestra la Data no "
                   "coinciden. El detalle está en el desplegable de conciliación.")

    with st.expander("❓ Conciliación: lo reportado por la oficina vs la Data",
                     expanded=False):
        ejemplos_ap = ", ".join(t.loc[fin_ap & ~cien_mt3, "SBAN"].astype(str).tolist()[:6]) or "—"
        st.markdown(
            f"""
Las dos cifras vienen de **fuentes distintas** y no miden lo mismo: una es **lo que reportó la
oficina** y la otra es **lo que el panel revisa en la Data de ejecución**.

| Fuente | Qué cuenta | Valor |
|---|---|---|
| **Reportado por la oficina** (archivo Campos) | Oficinas que reportaron sus UPS como `Finalizado` | **{fin}** |
| **Reportado por la oficina** (avance del reporte) | Oficinas que ya traen algún dato (de {total}) | **{reportadas}** |
| **Revisado por el panel** (Data de ejecución) | Oficinas donde **todos** sus UPS ya tienen mantenimiento hecho | **{n_cien}** |

Al cruzar oficina por oficina ({n_con_ups} oficinas con UPS en la Data):

- ✅ Coinciden en **{ambos}** oficinas.
- ⚠️ La oficina reportó `Finalizado` pero **falta mantenimiento** en **{solo_ap}** oficina(s){f" (ej. {ejemplos_ap})" if solo_ap else ""}.
- ℹ️ Todos sus UPS ya tienen mantenimiento pero **la oficina no lo reportó** como `Finalizado`: **{solo_mt3}** oficina(s) (aparecen en la columna **Alerta** de la tabla).
- ⚪ **{sin_rep}** oficina(s) todavía **no reportaron** ese dato: siguen en blanco en el archivo.

> 💡 Por eso el indicador principal es **{fin}** (lo reportado) y el chequeo con la Data
> (**{n_cien}**) solo sirve para detectar estas diferencias, no para reemplazar lo reportado.
"""
        )
        detalle = t.loc[(fin_ap & ~cien_mt3) | (~fin_ap & cien_mt3),
                        ["SBAN", "Oficina", "UPS en Data", "UPS con MT3",
                         "Estado UPS normalizado"]].copy()
        detalle["Diferencia"] = [
            "La oficina reportó Finalizado, pero falta mantenimiento"
            if f else "Todo con mantenimiento, pero la oficina no lo reportó como Finalizado"
            for f in detalle["Estado UPS normalizado"].eq("Finalizada")
        ]
        st.dataframe(detalle.rename(columns={"Estado UPS normalizado": "Estado reportado"}),
                     hide_index=True, width="stretch",
                     column_config={
                         "SBAN": st.column_config.TextColumn("SBAN", width="small"),
                         "Oficina": st.column_config.TextColumn("Oficina", width="medium"),
                         "UPS en Data": st.column_config.NumberColumn("UPS en Data", width="small"),
                         "UPS con MT3": st.column_config.NumberColumn("UPS con MT3", width="small"),
                         "Estado reportado": st.column_config.TextColumn(
                             "Estado reportado", width="small"),
                         "Diferencia": st.column_config.TextColumn("Diferencia", width="large"),
                     })

    orden = {"Finalizada": 0, "En proceso": 1, "Programada": 2, "Reprogramada": 3}
    t = t.assign(_o=t["Estado UPS normalizado"].map(orden).fillna(9),
                 _rep=t["Estado UPS normalizado"].astype(str).str.strip().eq(""))
    t = t.sort_values(["_rep", "_o", "SBAN"]).drop(columns=["_o", "_rep"]).reset_index(drop=True)

    vista = t.copy()
    vista["Estado reportado"] = vista["Estado UPS"].apply(
        lambda v: _badge_estado(normalizar_estado_ups(v)) if str(v).strip() else "⚪ Sin reporte")
    vista["Oficina"] = vista["Oficina"].fillna("").astype(str)
    cols_vista = ["SBAN", "Oficina", "UPS en Data", "UPS con MT3", "Avance MT3 %",
                  "Estado reportado", "Observación UPS", "Alerta"]
    if MODO_PUBLICO:
        vista["Observación UPS"] = ""
    st.dataframe(
        vista[cols_vista], hide_index=True, width="stretch", height=430,
        column_config={
            "SBAN": st.column_config.TextColumn("SBAN", width="small"),
            "Oficina": st.column_config.TextColumn("Oficina", width="medium"),
            "UPS en Data": st.column_config.NumberColumn("UPS en Data", width="small"),
            "UPS con MT3": st.column_config.NumberColumn("UPS con MT3", width="small"),
            "Avance MT3 %": st.column_config.ProgressColumn(
                "Avance MT3 %", min_value=0, max_value=100, format="%.0f%%", width="small"),
            "Estado reportado": st.column_config.TextColumn(
                "Estado reportado", width="small",
                help="Lo que la oficina escribió para esa sede en el archivo Campos dashboard."),
            "Observación UPS": st.column_config.TextColumn(
                "Observación UPS", width="large",
                help="Observación que la oficina dejó para esas UPS en el archivo Campos."),
            "Alerta": st.column_config.TextColumn(
                "Alerta", width="medium",
                help="Diferencia entre lo reportado por la oficina y lo que muestra la Data."),
        },
    )
    c1, c2 = st.columns(2)
    c1.download_button(
        "⬇️ Descargar avance UPS (CSV)", width="stretch",
        data=vista[cols_vista].to_csv(index=False).encode("utf-8-sig"),
        file_name="avance_ups_estado_colAP.csv", mime="text/csv")
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        vista[cols_vista].to_excel(writer, index=False, sheet_name="Avance UPS")
    c2.download_button(
        "⬇️ Descargar avance UPS (Excel)", width="stretch",
        data=buf.getvalue(), file_name="avance_ups_estado_colAP.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("El estado de cada sede se lee **tal cual** de la columna AP "
               "«ESTADO UPS» del archivo Campos dashboard. Las columnas de conteo y "
               "% vienen de la Data y sirven de control cruzado (no cambian el estado).")


def render_novedades_oficinas(n, seleccion=None, fecha_corte=None):
    """Pestaña de novedades de las oficinas (hoja 'Novedades Equipos')."""
    st.markdown("## 📰 Novedades de las oficinas")
    st.caption("Novedades reportadas por las oficinas, tomadas de la hoja "
               "**Novedades Equipos** del archivo Campos dashboard. "
               "Una fila = una novedad (serial, estado y observación).")
    if n is None or getattr(n, "empty", True):
        st.info("El archivo **Campos dashboard** no trae novedades de oficina "
                "(hoja `Novedades Equipos`) para este módulo. "
                "Cárgalo desde **📂 Ingesta diaria** en la barra lateral.")
        return

    d = n.copy()
    if seleccion and "_SBAN" in d.columns:
        sbans = {str(s).split(" - ")[0] for s in seleccion}
        d = d[d["_SBAN"].astype(str).isin(sbans)]
    if d.empty:
        st.info("Sin novedades para las oficinas seleccionadas.")
        return

    d["_Fecha"] = pd.to_datetime(d["_FECHA"], errors="coerce")
    d["_Cat"] = (d["Categoría Novedad"].fillna("").astype(str).str.strip()
                 if "Categoría Novedad" in d.columns else "")

    cats = sorted({c for c in d["_Cat"] if c and c.lower() != "nan"})
    k1, k2, k3 = st.columns([2.2, 1.4, 1])
    sel_cat = k1.multiselect("🏷️ Categoría de la novedad", options=cats, default=cats,
                             key="nov_ofi_cat", placeholder="Todas las categorías")
    fechas = d["_Fecha"].dropna()
    if len(fechas):
        f_min, f_max = fechas.min().date(), fechas.max().date()
        rango = k2.date_input("📅 Rango de fechas", value=(f_min, f_max),
                              min_value=f_min, max_value=f_max, key="nov_ofi_fecha")
        if isinstance(rango, tuple) and len(rango) == 2:
            ini, fin = pd.Timestamp(rango[0]), pd.Timestamp(rango[1])
            d = d[d["_Fecha"].isna() | d["_Fecha"].between(ini, fin)]
    k3.metric("Novedades mostradas", f"{len(d):,}".replace(",", "."), border=True,
              help="Novedades que cumplen los filtros de categoría y fechas de esta pestaña. "
                   "Los totales por oficina están en la pestaña 🏢 Resumen por oficina.")

    if sel_cat:
        d = d[d["_Cat"].isin(sel_cat)]
    if d.empty:
        st.info("Ninguna novedad coincide con los filtros.")
        return

    acomodo = {"DENUNCIO": 0, "DENUNCIA": 0, "RECOLECCION": 1,
               "CAMBIO DE ESTADO A FACTURABLE": 2}
    d = d.assign(_o=d["_Cat"].map(acomodo).fillna(5),
                 _f=d["_Fecha"].fillna(pd.Timestamp("1900-01-01")))
    d = d.sort_values(["_f", "_o"], ascending=[False, True]).drop(columns=["_o", "_f"])

    vista = pd.DataFrame({
        "SBAN": d["_SBAN"].astype(str),
        "Oficina": (d["Nombre Oficina"].fillna("").astype(str)
                    if "Nombre Oficina" in d.columns else ""),
        "Fecha": d["_Fecha"].dt.strftime("%Y-%m-%d").fillna(""),
        "Regional": (d["Departamento"].fillna("").astype(str)
                     if "Departamento" in d.columns else ""),
        "Aliado": (d["ALIADO"].fillna("").astype(str) if "ALIADO" in d.columns else ""),
        "Categoría": d["_Cat"],
        "Estado del equipo": (d["Estado novedad"].fillna("").astype(str)
                              if "Estado novedad" in d.columns else ""),
        "Facturable": (d["Facturable novedad"].fillna("").astype(str)
                       if "Facturable novedad" in d.columns else ""),
        "Serial": (d["_Serial"].astype(str) if "_Serial" in d.columns else ""),
        "Observación": (d["Obs. novedad detalle"].fillna("").astype(str)
                        if "Obs. novedad detalle" in d.columns else ""),
    })
    if MODO_PUBLICO:
        vista["Serial"] = vista["Serial"].apply(_mask_serial)
        vista["Observación"] = ""
    vista = vista[~vista["Serial"].astype(str).str.lower().isin(["", "nan"]) | vista["Categoría"].ne("")]
    vista = vista.reset_index(drop=True)

    t = tema_actual()
    por_cat = (vista.assign(_u=1).groupby("Categoría", dropna=False)["_u"].sum()
               .sort_values(ascending=False))
    tot_hoja = len(n) if n is not None else 0
    txt_nov = f"**{len(vista):,}** novedad(es) con datos"
    if tot_hoja:
        txt_nov += f" de **{tot_hoja:,}** filas de la hoja `Novedades Equipos`"
    txt_nov += (" (las filas sin categoría, fecha, serial ni observación se descartan).")
    st.caption(txt_nov.replace(",", "."))
    fig = go.Figure(go.Bar(
        x=[int(v) for v in por_cat.values], y=[str(i) for i in por_cat.index],
        orientation="h", marker=dict(color=t["primario"]),
        text=[int(v) for v in por_cat.values], textposition="outside",
        cliponaxis=False,
        hovertemplate="<b>%{y}</b><br>%{x} novedad(es)<extra></extra>",
    ))
    fig.update_layout(
        title=dict(text="<b>Novedades por categoría</b>", font=dict(size=15, color=t["ink"]),
                   x=0.02),
        height=max(240, 34 * len(por_cat) + 110), margin=dict(l=10, r=30, t=55, b=10),
        showlegend=False, **plotly_base())
    st.plotly_chart(fig, width="stretch",
                    config={"displaylogo": False,
                            "modeBarButtonsToRemove": ["lasso2d", "select2d"]})

    st.markdown(f"##### Detalle de novedades ({len(vista):,})".replace(",", "."))
    st.dataframe(vista, hide_index=True, width="stretch", height=460,
                 column_config={
                     "SBAN": st.column_config.TextColumn("SBAN", width="small"),
                     "Oficina": st.column_config.TextColumn("Oficina", width="medium"),
                     "Fecha": st.column_config.TextColumn("Fecha", width="small"),
                     "Regional": st.column_config.TextColumn("Departamento", width="small"),
                     "Aliado": st.column_config.TextColumn("Aliado", width="small"),
                     "Categoría": st.column_config.TextColumn("Categoría", width="medium"),
                     "Estado del equipo": st.column_config.TextColumn("Estado del equipo",
                                                                     width="small"),
                     "Facturable": st.column_config.TextColumn("Facturable", width="small"),
                     "Serial": st.column_config.TextColumn("Serial", width="medium"),
                     "Observación": st.column_config.TextColumn("Observación", width="large"),
                 })
    c1, c2 = st.columns(2)
    c1.download_button(
        "⬇️ Descargar novedades (CSV)", width="stretch",
        data=vista.to_csv(index=False).encode("utf-8-sig"),
        file_name="novedades_oficinas.csv", mime="text/csv")
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        vista.to_excel(writer, index=False, sheet_name="Novedades oficinas")
        por_cat.rename("Novedades").to_frame().to_excel(writer, sheet_name="Por categoría")
    c2.download_button(
        "⬇️ Descargar novedades (Excel)", width="stretch",
        data=buf.getvalue(), file_name="novedades_oficinas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if fecha_corte is not None and pd.notna(fecha_corte):
        st.caption(f"Fecha de corte de las novedades: **{pd.Timestamp(fecha_corte):%Y-%m-%d}** "
                   f"· {len(vista):,} novedad(es) mostradas.".replace(",", "."))


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

    # --- Campos dashboard (estado de sede) y Novedades Equipos (hoja propia) ---
    if kpi is None or getattr(kpi, "empty", True):
        avisos.append("⚠️ Campos dashboard no disponible: no se integra el estado de la sede "
                      "ni el avance de UPS.")
    elif _col(kpi, "ESTADO UPS") is None:
        avisos.append("⚠️ Campos dashboard: falta la columna **ESTADO UPS** (col AP); "
                      "la pestaña de avance UPS quedará sin estado por sede.")
    else:
        c_ups = _col(kpi, "ESTADO UPS")
        vals_ups = kpi[c_ups].fillna("").astype(str).str.strip()
        n_ups = int(vals_ups.ne("").sum())
        raros_ups = sorted({v for v in vals_ups.unique() if v and v not in
                            ("Finalizado", "Finalizada", "En proceso", "Programada",
                             "Programado", "Reprogramada", "No aplica")})
        avisos.append(f"ℹ️ Campos dashboard · ESTADO UPS (col AP): {n_ups} sede(s) con estado "
                      f"reportado de {len(kpi)} fila(s).")
        if raros_ups:
            avisos.append(f"⚠️ ESTADO UPS: valores nuevos/no esperados: {raros_ups[:6]}.")
    if nov is None or getattr(nov, "empty", True):
        avisos.append("⚠️ Novedades Equipos no disponible: la pestaña «Novedades de las "
                      "oficinas» quedará vacía.")
    else:
        faltan_n = set()
        if _col(nov, "Fecha") is None:
            faltan_n.add("Fecha")
        if _col(nov, "SBAN") is None and _col(nov, "SBAN PCT") is None:
            faltan_n.add("SBAN/SBAN PCT")
        if _col(nov, "Serial") is None:
            faltan_n.add("Serial")
        if _col(nov, "Categoria Novedad") is None:
            faltan_n.add("Categoria Novedad")
        if faltan_n:
            avisos.append("⚠️ Novedades Equipos: faltan columnas " + ", ".join(sorted(faltan_n)) + ".")
        else:
            c_cat = _col(nov, "Categoria Novedad")
            c_fec = _col(nov, "Fecha")
            c_ser = _col(nov, "Serial")
            c_obs = _col(nov, "Observaciones")
            c_sb = _col(nov, "SBAN") or _col(nov, "SBAN PCT")
            con_datos = (nov[c_cat].astype(str).str.strip().replace("nan", "").ne("")
                         | pd.to_datetime(nov[c_fec], errors="coerce").notna()
                         | nov[c_ser].astype(str).str.strip().replace("nan", "").ne("")
                         | (nov[c_obs].astype(str).str.strip().replace("nan", "").ne("")
                            if c_obs else False))
            n_datos = int(con_datos.sum())
            vals = set(nov[c_cat].dropna().astype(str).str.strip())
            vals.discard("")
            conocidos = {"DENUNCIO", "DENUNCIA", "RECOLECCION", "CAMBIO DE ESTADO A FACTURABLE",
                         "OPERACION", "UBICACION ERRADA", "ADICIONAL"}
            raros = sorted(vals - conocidos)
            if raros:
                avisos.append(f"⚠️ Categorías de novedad nuevas/no esperadas: {raros[:6]}.")
            sb = {_pad5(v) for v in nov[c_sb].dropna()}
            sb.discard("")
            en_data = {_pad5(x) for x in pd.to_numeric(d["SBAN"], errors="coerce").dropna().unique()}
            fuera = sorted(sb - en_data)
            if fuera:
                avisos.append(f"⚠️ Novedades Equipos: {len(fuera)} SBAN sin registro en la Data "
                              f"(ej. {fuera[:5]}).")
            avisos.append(f"ℹ️ Novedades Equipos: {n_datos} fila(s) con datos de {len(nov)} "
                          f"(se ignoran {len(nov) - n_datos} vacías).")

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
def render_sidebar(df: pd.DataFrame, cod_mod: str = "C1"):
    with st.sidebar:
        t = tema_actual()
        _logo_txt = "#04121A" if es_oscuro() else "#FFFFFF"
        st.markdown(
            f'<div style="display:flex;align-items:center;gap:11px;padding:2px 0 12px;'
            f'border-bottom:1px solid {t["border"]};margin-bottom:12px">'
            f'  <div style="width:38px;height:38px;border-radius:11px;display:grid;place-items:center;'
            f'font-weight:800;font-size:14px;color:{_logo_txt};background:{t["primario"]}">MT3</div>'
            f'  <div><div style="font-size:14.5px;font-weight:700;color:{t["ink"]}">Panel MT3</div>'
            f'  <div style="font-size:11.5px;color:{t["muted"]}">Banco Agrario · COLSOF</div></div>'
            f'</div>',
            unsafe_allow_html=True,
        )
        # ─── INGESTA DIARIA: solo visible en MODO ADMINISTRADOR ───
        # En la app pública (Cloud) no existe el secreto `modo_admin` → el público solo
        # consulta; la carga de datos queda restringida al administrador.
        if MODO_ADMIN:
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
        else:
            st.markdown("## 🔒 Datos")
            st.caption(f"🕒 Última actualización: **{leer_ultima_actualizacion()}**")
            st.caption("Panel de **consulta**. La carga y actualización de datos está "
                       "restringida al administrador.")

        st.divider()
        st.markdown("## 🛠️ Filtros")
        base_mod = df[df["_componente"].eq(cod_mod)] if "_componente" in df.columns else df
        opciones = sorted(base_mod["_ofi_key"].dropna().unique().tolist())
        # Limpieza (no asignación nueva) antes de crear el widget: si una oficina ya no
        # existe en el módulo, se quita de la selección.
        st.session_state["f_oficinas"] = [
            o for o in st.session_state.get("f_oficinas", []) if o in opciones]

        seleccion = st.multiselect(
            "Oficina (SBAN - Nombre)",
            options=opciones,
            key="f_oficinas",
            placeholder="Todas las oficinas (vista global)",
            help="Opciones según el módulo activo. Sin selección se muestran TODAS.",
        )
        if not opciones:
            st.warning("Este módulo no tiene oficinas con datos.")

        solo_pendientes = st.checkbox(
            "Mostrar solo pendientes ⚠️",
            key="f_solo_pend",
            help="Filtra la tabla y métricas a los equipos que aún no tienen consecutivo MT3.",
        )

        # El filtro de facturación vive ahora en la BARRA VISIBLE del área principal
        # (arriba, siempre a la vista). Aquí solo se sincroniza su valor.
        f_attr = f_attr_seleccionado()
        st.caption("💡 El filtro de **Facturación** está arriba, en la barra visible.")

        st.divider()
        st.button("🧹 Limpiar todos los filtros", width="stretch", key="limpiar_filtros",
                  on_click=_limpiar_todos_los_filtros)
        if st.button("🔄 Recargar datos", width="stretch"):
            st.cache_data.clear()
            st.rerun()

        # 🗂️ Auditoría y fuentes: información interna → solo en modo administrador
        if MODO_ADMIN:
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
def _sparkline_svg(valores, color: str, w: int = 118, h: int = 30) -> str:
    """Sparkline SVG inline (sin dependencias) a partir de una serie de valores."""
    vals = [float(v) for v in valores if v is not None]
    if len(vals) < 2:
        return ""
    vmin, vmax = min(vals), max(vals)
    rango = (vmax - vmin) or 1.0
    n = len(vals)
    pts = [(2 + i * (w - 4) / (n - 1), (h - 3) - ((v - vmin) / rango) * (h - 9))
           for i, v in enumerate(vals)]
    linea = " ".join(f"{x:.1f},{y:.1f}" for x, y in pts)
    area = f"{pts[0][0]:.1f},{h - 2} " + linea + f" {pts[-1][0]:.1f},{h - 2}"
    cx, cy = pts[-1]
    return (f'<svg width="{w}" height="{h}" viewBox="0 0 {w} {h}" style="display:block">'
            f'<polygon points="{area}" fill="{color}" opacity=".15"/>'
            f'<polyline points="{linea}" fill="none" stroke="{color}" stroke-width="2" '
            f'stroke-linejoin="round" stroke-linecap="round"/>'
            f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="3" fill="{color}"/></svg>')


def oficinas_campos(kpi_campos, seleccion=None, click_sban=None):
    """Fuente ÚNICA de los indicadores oficiales de Campos dashboard.

    Devuelve (universo, finalizadas_N, reportadas_AP, finalizadas_AP), todo contado por
    **SBAN único** (no por fila) para que cualquier tarjeta que use este dato muestre
    exactamente la misma cifra. Lo usan las tarjetas destacadas y las de indicadores.
    """
    if kpi_campos is None or getattr(kpi_campos, "empty", True):
        return set(), 0, 0, 0
    c_sb = _col(kpi_campos, "SBAN")
    if c_sb is None:
        return set(), 0, 0, 0
    kk = kpi_campos.copy()
    kk["_SBAN"] = kk[c_sb].apply(_pad5)
    kk = kk[kk["_SBAN"].astype(str).str.len() > 0]
    if seleccion:
        sbans = {str(s).split(" - ")[0] for s in seleccion}
        kk = kk[kk["_SBAN"].isin(sbans)]
    if click_sban:
        kk = kk[kk["_SBAN"].eq(click_sban)]
    universo = set(kk["_SBAN"].astype(str))

    finalizadas_n = 0
    c_est = _col(kk, "Estado de la sede")
    if c_est is not None:
        estados = _mejor_por_sban(zip(kk["_SBAN"].astype(str),
                                      kk[c_est].fillna("").astype(str)), _PRIO_SEDE)
        finalizadas_n = sum(1 for v in estados.values() if v == "Finalizada")
    reportadas_ap = finalizadas_ap = 0
    c_ap = _col(kk, "ESTADO UPS")
    if c_ap is not None:
        ups = _mejor_por_sban(zip(kk["_SBAN"].astype(str),
                                  kk[c_ap].apply(normalizar_estado_ups)), _PRIO_UPS)
        reportadas_ap = len(ups)
        finalizadas_ap = sum(1 for v in ups.values() if v == "Finalizada")
    return universo, finalizadas_n, reportadas_ap, finalizadas_ap


def tarjetas_destacadas(datos: pd.DataFrame):
    """KPIs destacados: oficinas intervenidas (revisado con la Data) y avance MT3.

    «Oficinas intervenidas» = oficinas con al menos 1 equipo con mantenimiento hecho en el
    módulo activo: es una revisión que hace el panel con la Data, distinta de lo que la
    oficina reporta en el archivo Campos, que se muestra en «🏁 Oficinas reportadas como
    finalizadas». Así las dos cifras no comparten etiqueta ni se confunden.

    `datos` debe llegar SIN el filtro de tabla «solo pendientes»: ese filtro deja solo
    filas sin MT3, así que aplicarlo aquí haría caer el numerador a 0. El numerador y el
    denominador se calculan sobre el MISMO conjunto (`datos`), para que la tarjeta
    siempre sea coherente consigo misma.

    Los conteos de EQUIPOS (Total / Realizados / Pendientes) no se repiten aquí: viven
    en `render_kpis()`. Esta cabecera solo muestra el % de avance con su barra.
    """
    t = tema_actual()
    if datos is None or len(datos) == 0:
        return

    def fmt(n) -> str:
        return f"{int(n):,}".replace(",", ".")

    # Numerador y denominador sobre el mismo conjunto: oficinas del módulo activo que
    # cumplen los filtros de sidebar (Oficina · Facturación · clic del ranking).
    n_ofi = int(datos["_SBAN"].nunique())
    n_reg = int(datos["Regional"].nunique()) if "Regional" in datos.columns else 0
    try:
        g_of = datos.groupby("_SBAN")["_mt"].agg(["size", "sum"])
        n_ofi_int = int((g_of["sum"] > 0).sum())
    except Exception:  # noqa: BLE001
        n_ofi_int = 0
    total = int(len(datos))
    realizados = int(datos["_mt"].sum())
    avance = (realizados / total * 100) if total else 0.0
    av_txt = f"{avance:.1f}".replace(".", ",") + " %"

    serie: list = []
    if "Fecha de mantenimiento 3" in datos.columns:
        f = pd.to_datetime(datos.loc[datos["_mt"], "Fecha de mantenimiento 3"],
                           errors="coerce").dropna()
        if not f.empty:
            serie = f.dt.date.value_counts().sort_index().values.tolist()
    spark = _sparkline_svg(serie, t["verde"])

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(
            f'<div class="kd"><div class="kd-lbl">🏢 Oficinas intervenidas</div>'
            f'<div class="kd-num">{fmt(n_ofi_int)}</div>'
            f'<div class="kd-sub">de {fmt(n_ofi)} oficinas del módulo'
            + (f' · {n_reg} regionales' if n_reg else '') + '</div></div>',
            unsafe_allow_html=True)
    with c2:
        st.markdown(
            f'<div class="kd"><div class="kd-lbl">📈 Avance MT3</div>'
            f'<div class="kd-num">{av_txt}</div>'
            f'<div class="kd-bar"><i style="width:{max(0.0, min(100.0, avance)):.1f}%;'
            f'background:{t["verde"]}"></i></div>'
            f'<div class="kd-sub">{fmt(realizados)} de {fmt(total)} equipos</div></div>',
            unsafe_allow_html=True)


def render_kpis(datos: pd.DataFrame):
    """Tarjetas KPI del filtro activo (módulo + sidebar + clic del ranking).

    Es la ÚNICA fuente de los conteos de equipos: el resto de vistas no los repite.
    El % de avance vive en la tarjeta destacada «📈 Avance MT3» (arriba), así que
    aquí NO se repite como delta.
    """
    total = len(datos)
    realizados = int(datos["_mt"].sum())
    pendientes = int(datos["_pendiente"].sum())
    fact_si = int((datos["Facturable"] == "Si").sum())
    fact_no = int((datos["Facturable"] == "No").sum())

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("📦 Total de elementos", f"{total:,}".replace(",", "."),
              help="Equipos del módulo activo que cumplen los filtros (Oficina · Facturación · "
                   "solo pendientes · clic del ranking).", border=True)
    c2.metric("✅ Mantenimientos realizados (MT)", f"{realizados:,}".replace(",", "."),
              help="Filas cuyo 'Consecutivo mantenimiento 3' empieza por MT. El % de avance del "
                   "filtro está en la tarjeta destacada «📈 Avance MT3» (no se repite aquí).",
              border=True)
    c3.metric("⚠️ Mantenimientos pendientes", f"{pendientes:,}".replace(",", "."),
              delta=f"{pendientes/total*100:.1f}%" if total else "0%",
              delta_color="inverse", help="Total de elementos − Realizados", border=True)
    c4.metric("🏦 Facturables (Si)", f"{fact_si:,}".replace(",", "."),
              help="Elementos atribuibles a BANCO", border=True)
    c5.metric("🏢 No facturables", f"{fact_no:,}".replace(",", "."),
              help="Elementos de gestión COLSOF", border=True)
    return total, realizados, pendientes, fact_si, fact_no


def render_kpis_oficinas(df: pd.DataFrame, seleccion: list, f_attr: str,
                         click_sban=None, kpi_campos=None, cod_mod="C1",
                         solo_pendientes=False):
    """Tarjetas de OFICINAS, todas con el mismo universo y los mismos filtros.

    Universo = oficinas del archivo Campos dashboard (una por SBAN). Las 5 tarjetas
    aplican los MISMOS filtros (Oficina · Facturación 🏦/🏢 · clic del ranking ·
    solo pendientes) y **no** dependen del módulo activo, para que sus cifras sean
    siempre comparables y no cambien al cambiar de pestaña de módulo.

    Nomenclatura pensada para el usuario final (sin jerga interna):

    - `🏢 Oficinas con <componente> completo` = el panel revisa la Data de ejecución y
      cuenta las oficinas donde TODOS los equipos de ese componente tienen consecutivo
      MT3. El valor se muestra como "completas de total" y el delta dice cuántas
      oficinas del archivo faltan. El tooltip explica qué queda fuera por los filtros.
    - `🏁 Oficinas reportadas como finalizadas` = lo que reportó la oficina en el archivo
      Campos dashboard (columna «Estado de la sede»), contado por SBAN único.
    - `🔋 UPS finalizadas (reportado)` = lo que reportó la oficina en la columna
      «ESTADO UPS» del archivo, mostrado sobre el total de oficinas (no sobre las
      reportadas: un "100% de las reportadas" ocultaría las que aún no reportan).

    `cod_mod` se recibe para documentar el módulo activo, pero las cifras se calculan
    para los 3 componentes siempre (así no cambian al cambiar de módulo).
    """
    b = df
    if seleccion:
        b = b[b["_ofi_key"].isin(seleccion)]
    if f_attr != "T":
        b = b[b["Facturable"] == f_attr]
    if click_sban:
        b = b[b["_SBAN"].eq(click_sban)]
    if solo_pendientes and "_pendiente" in b.columns:
        b = b[b["_pendiente"]]

    # Universo y cifras oficiales: fuente ÚNICA compartida con tarjetas_destacadas()
    universo, fin_n, rep_ap, fin_ap = oficinas_campos(kpi_campos, seleccion, click_sban)
    if not universo:
        universo = set(df["_SBAN"].astype(str))
    n_universo = len(universo)
    n_filtro = len({s for s in b["_SBAN"].astype(str)} & universo) if len(b) else 0
    n_datos = len(set(df["_SBAN"].astype(str)))

    etiqueta_filtro = {"T": "Todas las oficinas",
                       "Si": "filtro 🏦 Facturables (BANCO)",
                       "No": "filtro 🏢 No facturables (COLSOF)"}.get(f_attr, "Todas las oficinas")
    if seleccion:
        etiqueta_filtro += f" · {len(seleccion)} oficina(s) seleccionada(s)"
    if click_sban:
        etiqueta_filtro += f" · sede {click_sban}"
    if solo_pendientes:
        etiqueta_filtro += " · solo pendientes"

    st.markdown("##### 🏢 Indicadores de oficinas")
    st.caption(f"{etiqueta_filtro} · se están contando **{n_filtro}** de las "
               f"**{n_universo}** oficinas del archivo.")
    cols = st.columns(5)

    for col, (cod, nombre) in zip(cols[:3], [("C1", "Componente 1"),
                                             ("C2", "Componente 2 (láser)"),
                                             ("UPS", "UPS")]):
        # Cada tarjeta se calcula con SU componente, aplicando los mismos filtros.
        sub = b[b["_componente"].eq(cod)] if "_componente" in b.columns else b.iloc[0:0]
        sub_global = (df[df["_componente"].eq(cod)] if "_componente" in df.columns
                      else df.iloc[0:0])
        if sub.empty:
            col.metric(f"🏢 Oficinas con {nombre} completo", "0 de 0",
                       delta=f"{n_universo} pendientes",
                       help=f"Ninguna oficina con {nombre} cumple los filtros actuales.",
                       border=True)
            continue
        g = sub.groupby("_SBAN")["_mt"].agg(["size", "sum"])
        visibles = g.index.astype(str).isin(universo)
        total = int(visibles.sum())
        completas = int((g["sum"] >= g["size"])[visibles].sum())
        pct = (completas / total * 100) if total else 0.0
        ayuda = (f"**Completo** = la oficina ya tiene el mantenimiento hecho en TODOS sus "
                 f"equipos de {nombre} (consecutivo MT3). Calculado por el panel con la Data "
                 f"de ejecución: {completas} de {total} oficinas que cumplen "
                 f"{etiqueta_filtro.lower()}.")
        if len(b) < len(df):
            # Cuántas oficinas con ese componente quedan FUERA por los filtros activos.
            fuera = len((set(sub_global["_SBAN"].astype(str)) & universo)
                        - (set(sub["_SBAN"].astype(str)) & universo))
            if fuera:
                ayuda += (f" Los filtros dejan fuera {fuera} oficina(s) que sí tienen {nombre} "
                          f"(el archivo tiene {n_universo} oficinas en total).")
        col.metric(f"🏢 Oficinas con {nombre} completo", f"{completas} de {total}",
                   delta=f"{n_universo - completas} pendientes en el archivo",
                   delta_color="off", help=ayuda, border=True)

    with cols[3]:
        if universo and _col(kpi_campos, "Estado de la sede") is not None:
            tot = n_universo
            pct = (fin_n / tot * 100) if tot else 0.0
            cols[3].metric(
                "🏁 Oficinas reportadas como finalizadas", f"{fin_n} de {tot}",
                delta=f"{tot - fin_n} pendientes por reportar",
                delta_color="off",
                help="Estado de la sede SEGÚN LA OFICINA: columna «Estado de la sede» del archivo "
                     "Campos dashboard, contada por SBAN único (una sede con varias filas se "
                     "resume por la de mayor prioridad: Finalizada > Reprogramada_Finalizada > "
                     "En proceso > Programada). Es distinto de «Oficinas con … completo», que el "
                     "panel calcula con la Data. 'Reprogramada_Finalizada' se cuenta aparte.",
                border=True,
            )
        else:
            cols[3].metric("🏁 Oficinas reportadas como finalizadas", "n/d",
                           help="Campos dashboard no disponible", border=True)

    # Estado OFICIAL de las UPS: columna AP «ESTADO UPS» de Campos dashboard.
    # El conteo calculado con la Data (equipos completos) NO se repite aquí: vive como
    # alerta y en el desplegable de conciliación de la pestaña 🔋 Avance UPS.
    with cols[4]:
        if rep_ap or fin_ap:
            cols[4].metric(
                "🔋 UPS finalizadas (reportado)", f"{fin_ap} de {n_universo}",
                delta=f"{n_universo - fin_ap} sin reportar",
                delta_color="off",
                help="Lo que la oficina reportó en la columna «ESTADO UPS» del archivo Campos "
                     f"dashboard, contada por SBAN único: {fin_ap} oficinas con 'Finalizado' de "
                     f"{n_universo}. Las otras {n_universo - fin_ap} todavía no traen ese dato.",
                border=True,
            )
        else:
            cols[4].metric("🔋 UPS finalizadas (reportado)", "n/d",
                           help="Campos dashboard sin la columna de estado de UPS", border=True)

    st.caption("**Cómo leer estas tarjetas.** Todas usan el mismo universo (las "
               f"{n_universo} oficinas del archivo) y los mismos filtros: Oficina · Facturación · "
               "clic del ranking · solo pendientes. "
               "**«Oficinas con … completo»** = el panel revisa la Data de ejecución y cuenta las "
               "oficinas donde ya se le hizo mantenimiento a TODOS sus equipos. "
               "**«Oficinas reportadas como finalizadas»** y **«UPS finalizadas (reportado)»** = "
               "lo que la propia oficina reportó en el archivo Campos dashboard. "
               "Son fuentes distintas: por eso pueden no coincidir (las diferencias se detallan en "
               "la pestaña 🔋 Avance UPS)."
               + (f" La Data trae {n_datos} oficinas: {n_datos - n_universo} de ellas no está(n) "
                  "en el archivo y no entran en las tarjetas." if n_datos > n_universo else ""))


# ----------------------------------------------------------------------------
# Gráficas Plotly
# ----------------------------------------------------------------------------
def grafico_dona(realizados: int, pendientes: int):
    t = tema_actual()
    total = realizados + pendientes
    pct = (realizados / total * 100) if total else 0.0
    fig = go.Figure(go.Pie(
        labels=["✅ Subsanados (MT)", "⚠️ Pendientes"],
        values=[realizados, pendientes],
        hole=0.68,
        marker=dict(colors=[t["verde"], t["pend"]], line=dict(color=t["card"], width=3)),
        textinfo="none",
        hovertemplate="<b>%{label}</b><br>%{value:,} equipos (%{percent})<extra></extra>",
    ))
    fig.update_layout(
        title=dict(text="<b>Avance de Mantenimiento Preventivo 3</b>",
                   font=dict(size=15, color=t["ink"]), x=0.02),
        annotations=[dict(
            text=(f"<span style='font-size:28px;font-weight:700;color:{t['ink']}'>{pct:.1f}%</span>"
                  f"<br><span style='font-size:11px;letter-spacing:.5px;color:{t['muted']}'>AVANCE GLOBAL</span>"),
            x=0.5, y=0.5, showarrow=False, align="center")],
        margin=dict(l=15, r=15, t=45, b=10),
        height=360,
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.08, xanchor="center", x=0.5,
                    font=dict(color=t["muted"])),
        **plotly_base(),
    )
    return fig


def grafico_tendencia(datos: pd.DataFrame):
    """Tendencia diaria y acumulada de MT3 (usa Fecha de mantenimiento 3)."""
    t = tema_actual()
    if "Fecha de mantenimiento 3" not in datos.columns:
        return None
    d = datos[datos["_mt"]].copy()
    d["_f"] = pd.to_datetime(d["Fecha de mantenimiento 3"], errors="coerce")
    d = d[d["_f"].notna()]
    if d.empty:
        return None
    serie = d.groupby(d["_f"].dt.date).size().sort_index()
    acum = serie.cumsum()
    x = [f"{f:%d/%m}" for f in serie.index]
    cel = str(t["celeste"]).lstrip("#")
    try:
        rr, gg, bb = int(cel[0:2], 16), int(cel[2:4], 16), int(cel[4:6], 16)
        relleno = f"rgba({rr},{gg},{bb},0.15)"
    except (ValueError, IndexError):
        relleno = "rgba(34,211,238,0.15)"
    fig = go.Figure()
    fig.add_scatter(x=x, y=serie.values, name="Por día", mode="lines+markers+text",
                    line=dict(color=t["celeste"], width=2.5), marker=dict(size=7),
                    text=[str(int(v)) for v in serie.values], textposition="top center",
                    textfont=dict(size=10, color=t["ink"]), cliponaxis=False,
                    fill="tozeroy", fillcolor=relleno,
                    hovertemplate="%{x}<br>MT3 del día: %{y}<extra></extra>")
    fig.add_scatter(x=x, y=acum.values, name="Acumulado", mode="lines",
                    line=dict(color=t["verde"], width=2, dash="dot"), yaxis="y2",
                    hovertemplate="%{x}<br>Acumulado: %{y}<extra></extra>")
    total_d = int(serie.sum())
    fig.update_layout(
        title=dict(text=f"<b>Tendencia MT3 (realizados por día)</b>"
                        f"<br><span style='font-size:11px'>Total: {total_d} MT3 · "
                        f"último día: {x[-1]} ({int(serie.iloc[-1])})</span>",
                   font=dict(size=15, color=t["ink"]), x=0.02),
        height=360, margin=dict(l=10, r=10, t=65, b=10),
        xaxis=dict(title="", gridcolor=t["grid"], zeroline=False,
                   tickfont=dict(color=t["muted"])),
        yaxis=dict(title="Por día", rangemode="tozero", gridcolor=t["grid"],
                   tickfont=dict(color=t["muted"])),
        yaxis2=dict(title="Acumulado", overlaying="y", side="right", showgrid=False,
                    tickfont=dict(color=t["muted"])),
        legend=dict(orientation="h", y=1.02, x=0, font=dict(color=t["muted"])),
        **plotly_base(),
    )
    return fig


def grafico_mini_componentes(df: pd.DataFrame, seleccion: list, f_attr: str,
                             click_sban=None):
    """Barras de cobertura de oficinas con MT3 por componente (solo visual)."""
    t = tema_actual()
    b = df
    if seleccion:
        b = b[b["_ofi_key"].isin(seleccion)]
    if f_attr != "T":
        b = b[b["Facturable"] == f_attr]
    if click_sban:
        b = b[b["_SBAN"].eq(click_sban)]
    nombres, vals, hovers = [], [], []
    for cod, nombre in (("C1", "Componente 1"), ("C2", "Comp. 2 láser"), ("UPS", "UPS")):
        sub = b[b["_componente"].eq(cod)] if "_componente" in b.columns else b.iloc[0:0]
        if sub.empty:
            nombres.append(nombre); vals.append(0.0); hovers.append("sin datos")
            continue
        g = sub.groupby("_SBAN")["_mt"].agg(["size", "sum"])
        tot = int(len(g))
        inter = int((g["sum"] > 0).sum())
        nombres.append(nombre)
        vals.append(round(inter / tot * 100, 1) if tot else 0.0)
        hovers.append(f"{inter} de {tot} oficinas")
    fig = go.Figure(go.Bar(
        x=vals, y=nombres, orientation="h",
        text=[f"{v:.0f}%" for v in vals], textposition="inside",
        marker_color=[t["celeste"], t["ambar"], t["verde"]],
        customdata=hovers,
        hovertemplate="<b>%{y}</b><br>Cobertura oficinas: %{x}%<br>%{customdata}<extra></extra>",
    ))
    fig.update_layout(
        title=dict(text="<b>Oficinas con MT3 por componente</b>",
                   font=dict(size=15, color=t["ink"]), x=0.02),
        height=250, margin=dict(l=10, r=10, t=45, b=10), showlegend=False,
        xaxis=dict(range=[0, 100], visible=False),
        yaxis=dict(autorange="reversed", tickfont=dict(color=t["ink"], size=12)),
        **plotly_base(),
    )
    return fig


def grafico_bullet_oficinas(datos: pd.DataFrame):
    """Barra 100% de oficinas por estado de cobertura del módulo activo."""
    t = tema_actual()
    if datos.empty or "_SBAN" not in datos.columns:
        return None
    g = datos.groupby("_SBAN")["_mt"].agg(["size", "sum"])
    comp = int((g["sum"] >= g["size"]).sum())
    parcial = int(((g["sum"] > 0) & (g["sum"] < g["size"])).sum())
    cero = int((g["sum"] == 0).sum())
    fig = go.Figure()
    for nombre, valor, color in (("Completas", comp, t["verde"]),
                                 ("En proceso", parcial, t["ambar"]),
                                 ("Sin iniciar", cero, t["rojo"])):
        fig.add_bar(y=["Oficinas"], x=[valor], orientation="h", name=nombre,
                    marker_color=color, text=[f"{nombre}: {valor}"],
                    textposition="inside", insidetextanchor="middle",
                    hovertemplate=f"{nombre}: %{{x}} oficinas<extra></extra>")
    fig.update_layout(
        title=dict(text="<b>Oficinas por estado de cobertura</b>",
                   font=dict(size=15, color=t["ink"]), x=0.02),
        barmode="stack", height=250, margin=dict(l=10, r=10, t=45, b=10),
        xaxis=dict(visible=False),
        yaxis=dict(visible=False),
        legend=dict(orientation="h", y=-0.05, x=0, font=dict(color=t["muted"])),
        **plotly_base(),
    )
    return fig


def grafico_sparkline_tendencia(datos: pd.DataFrame):
    """Tendencia MT3 por día con valores y ejes visibles (lectura a primera vista)."""
    t = tema_actual()
    if "Fecha de mantenimiento 3" not in datos.columns:
        return None
    d = datos[datos["_mt"]].copy()
    d["_f"] = pd.to_datetime(d["Fecha de mantenimiento 3"], errors="coerce")
    d = d[d["_f"].notna()]
    if d.empty:
        return None
    serie = d.groupby(d["_f"].dt.date).size().sort_index()
    cel = str(t["celeste"]).lstrip("#")
    try:
        rr, gg, bb = int(cel[0:2], 16), int(cel[2:4], 16), int(cel[4:6], 16)
        relleno = f"rgba({rr},{gg},{bb},0.15)"
    except (ValueError, IndexError):
        relleno = "rgba(34,211,238,0.15)"
    etiquetas = [str(int(v)) for v in serie.values] if len(serie) <= 18 else None
    fig = go.Figure(go.Scatter(
        x=[f"{f:%d/%m}" for f in serie.index], y=serie.values,
        mode="lines+markers+text" if etiquetas else "lines+markers",
        text=etiquetas, textposition="top center",
        textfont=dict(size=10, color=t["ink"]),
        line=dict(color=t["celeste"], width=2), marker=dict(size=7),
        fill="tozeroy", fillcolor=relleno, cliponaxis=False,
        hovertemplate="%{x}<br>%{y} MT3<extra></extra>"))
    total_d = int(serie.sum())
    fig.update_layout(
        title=dict(text=f"<b>Tendencia MT3 por día</b>"
                        f"<br><span style='font-size:11px'>Total: {total_d} MT3 · "
                        f"último día: {serie.index[-1]:%d/%m} ({int(serie.iloc[-1])})</span>",
                   font=dict(size=15, color=t["ink"]), x=0.02),
        height=280, margin=dict(l=10, r=10, t=62, b=10), showlegend=False,
        xaxis=dict(title="", tickfont=dict(size=10, color=t["muted"]),
                   gridcolor=t["grid"], zeroline=False),
        yaxis=dict(title="", rangemode="tozero", tickfont=dict(size=10, color=t["muted"]),
                   gridcolor=t["grid"]),
        **plotly_base(),
    )
    return fig


def render_vista_ejecutiva(filtrado: pd.DataFrame, df: pd.DataFrame, seleccion: list,
                           f_attr: str, click_sban=None):
    """Vista solo de KPIs visuales: 5 gráficos, sin tablas ni cifras duras de KPIs."""
    st.markdown("#### 🎯 Vista ejecutiva · indicadores visuales")
    cfg = {"displaylogo": False, "modeBarButtonsToRemove": ["lasso2d", "select2d"]}
    if filtrado.empty:
        st.info("Sin datos con los filtros actuales.")
        return
    realizados = int(filtrado["_mt"].sum())
    pendientes = int(filtrado["_pendiente"].sum())

    # Un solo gráfico de avance: la dona (el taquímetro mostraba la misma cifra).
    st.plotly_chart(grafico_dona(realizados, pendientes), width="stretch", config=cfg)

    f2a, f2b = st.columns([1.4, 1])
    with f2a:
        st.plotly_chart(grafico_mini_componentes(df, seleccion, f_attr, click_sban),
                        width="stretch", config=cfg)
    with f2b:
        figb = grafico_bullet_oficinas(filtrado)
        if figb is not None:
            st.plotly_chart(figb, width="stretch", config=cfg)
        else:
            st.info("Sin oficinas para mostrar.")

    figs = grafico_sparkline_tendencia(filtrado)
    if figs is not None:
        st.plotly_chart(figs, width="stretch", config=cfg)
    else:
        st.info("Sin fechas de MT3 para la tendencia.")


def _color_en_escala(escala, frac):
    """Interpola un color (RGB) en la escala, en la posición 0..1 (igual que Plotly)."""
    def rgb(h):
        h = str(h).lstrip("#")
        return tuple(int(h[k:k + 2], 16) for k in (0, 2, 4))
    n = len(escala) - 1
    pos = max(0.0, min(1.0, float(frac))) * n
    i = int(pos)
    f = pos - i
    c1, c2 = rgb(escala[min(i, n)]), rgb(escala[min(i + 1, n)])
    return tuple(round(a + (b - a) * f) for a, b in zip(c1, c2))


def _texto_contraste(escala, frac):
    """Color de texto legible sobre el color de la escala (blanco u oscuro)."""
    r, g, b = _color_en_escala(escala, frac)
    lum = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0
    return "#0A0E14" if lum > 0.55 else "#FFFFFF"


def grafico_heatmap(datos: pd.DataFrame):
    """Mapa de calor Regional × Estado de la sede (pendientes, total y avance MT3).

    No oculta información:
      * incluye TODOS los estados presentes, aunque no estén en la lista conocida
        (p. ej. «Reprogramada_Finalizada»), y
      * agrupa las filas sin Regional bajo «SIN REGIONAL» en lugar de perderlas
        (antes `pivot_table` descartaba esas filas por tener índice vacío).
    La escala de color va de 0 al máximo real, para que el tono sea comparable.
    """
    t = tema_actual()
    if "Regional" not in datos.columns or "_est_crono" not in datos.columns:
        return None
    d = datos.copy()
    d["_est"] = (d["_est_crono"].astype(str)
                 .str.replace(r"^[^A-Za-zÁÉÍÓÚáéíóúñ]+", "", regex=True).str.strip())
    d.loc[d["_est"].isin(["", "nan", "None"]), "_est"] = "Sin cronograma"
    reg = d["Regional"].astype(str).str.strip()
    d["_reg"] = reg.where(~reg.isin(["", "nan", "None"]), "SIN REGIONAL")
    d["_mt_int"] = d["_mt"].astype(int)

    orden = ["Programada", "En proceso", "Finalizada", "Reprogramada",
             "Reprogramada_Finalizada", "Sin cronograma"]
    presentes = set(d["_est"])
    cols = ([c for c in orden if c in presentes]
            + [c for c in sorted(presentes) if c not in orden])
    if not cols:
        return None

    def _pivote(valores, agg):
        return (d.pivot_table(index="_reg", columns="_est", values=valores,
                              aggfunc=agg, fill_value=0)
                .reindex(columns=cols, fill_value=0))

    pend = _pivote("_pendiente", "sum")
    total = _pivote("_pendiente", "size")
    mt = _pivote("_mt_int", "sum")
    avance = (mt / total * 100).where(total > 0, 0).round(0)

    custom = [[[int(total.iloc[i, j]), f"{avance.iloc[i, j]:.0f}%"]
               for j in range(len(cols))] for i in range(len(pend.index))]
    zmax = float(pend.values.max()) if pend.values.size else 1.0
    zmax = zmax if zmax > 0 else 1.0
    fig = go.Figure(go.Heatmap(
        z=pend.values, x=cols, y=list(pend.index),
        zmin=0, zmax=zmax,                       # escala desde 0 hasta el máximo real
        customdata=custom, colorscale=t["heat"], xgap=3, ygap=3,
        colorbar=dict(title=dict(text="Pendientes", font=dict(color=t["ink"])),
                      tickfont=dict(color=t["muted"])),
        hovertemplate="<b>%{y}</b> · %{x}<br>Pendientes: %{z}<br>"
                      "Total: %{customdata[0]}<br>Avance MT3: %{customdata[1]}<extra></extra>",
    ))
    # Etiquetas con color calculado (contraste garantizado sobre cualquier tono del calor)
    for i, reg_n in enumerate(pend.index):
        for j, est in enumerate(cols):
            v = int(pend.values[i][j])
            if v <= 0:
                continue
            fig.add_annotation(x=est, y=reg_n, text=f"{v:,}".replace(",", "."),
                               showarrow=False, xref="x", yref="y",
                               font=dict(size=11, color=_texto_contraste(t["heat"], v / zmax)))
    _total_equipos = int(total.values.sum())
    fig.update_layout(
        title=dict(text="<b>Mapa de calor: pendientes por Regional y Estado de la sede</b>"
                        "<br><span style='font-size:11px'>«SIN REGIONAL» agrupa bodegas/stock "
                        "sin regional asignada. "
                        f"Equipos cubiertos: {_total_equipos:,}".replace(",", ".")
                        + "</span>",
                   font=dict(size=15, color=t["ink"]), x=0.02),
        height=max(320, 42 * len(pend.index) + 120),
        margin=dict(l=10, r=10, t=70, b=10),
        xaxis=dict(title="", color=t["muted"]),
        yaxis=dict(title="", autorange="reversed", color=t["muted"]),
        **plotly_base(),
    )
    return fig


def grafico_top_pendientes(datos: pd.DataFrame, top_n: int = 12):
    """Ranking horizontal de sedes con más pendientes (evita amontonamiento)."""
    t = tema_actual()
    tmp = datos.assign(
        _lbl=datos["_SBAN"] + " · " + datos["Oficina"].str.slice(0, 30),
        _full=datos["_SBAN"] + " - " + datos["Oficina"])
    ag = tmp.groupby(["_lbl", "_full"]).agg(Realizados=("_mt", "sum"),
                                            Pendientes=("_pendiente", "sum")).reset_index()
    ag["Total"] = ag["Realizados"] + ag["Pendientes"]
    ag = ag[ag["Total"] > 0].sort_values("Pendientes", ascending=False).head(top_n)
    ag = ag.sort_values("Pendientes", ascending=True)          # mayor arriba
    total_pend = int(ag["Pendientes"].sum())

    fig = go.Figure()
    fig.add_bar(
        y=ag["_lbl"], x=ag["Realizados"], orientation="h", name="✅ Subsanados (MT)",
        marker=dict(color=t["verde"], line=dict(width=0), cornerradius=4),
        customdata=ag["_full"],
        hovertemplate="<b>%{customdata}</b><br>Subsanados: %{x:,}<extra></extra>")
    fig.add_bar(
        y=ag["_lbl"], x=ag["Pendientes"], orientation="h", name="⚠️ Pendientes",
        marker=dict(color=t["pend"], line=dict(width=0), cornerradius=4),
        text=ag["Pendientes"], textposition="outside", textfont=dict(size=10, color=t["muted"]),
        cliponaxis=False, customdata=ag["_full"],
        hovertemplate="<b>%{customdata}</b><br>Pendientes: %{x:,}<extra></extra>")
    alto = max(340, 30 * len(ag) + 120)
    fig.update_layout(
        title=dict(text=f"<b>Top {len(ag)} sedes con más pendientes de MT3</b>"
                        f"<br><span style='font-size:11px;color:{t['muted']}'>{total_pend:,} pendientes "
                        f"en este ranking</span>".replace(",", "."),
                   font=dict(size=15, color=t["ink"]), x=0.02),
        barmode="stack", bargap=0.35, barcornerradius=5,
        height=alto,
        margin=dict(l=10, r=40, t=60, b=10),
        xaxis=dict(title="Equipos", gridcolor=t["grid"],
                   zeroline=False, automargin=True),
        yaxis=dict(title="", tickfont=dict(size=11, color=t["ink"]), automargin=True),
        legend=dict(orientation="h", yanchor="bottom", y=1.0, x=0, font=dict(color=t["muted"])),
        **plotly_base(),
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
    if "Reprogramada_Finalizada" in texto:
        return "E4DFEC", "3B2A63"
    for clave in ("Finalizada", "En proceso", "Reprogramada", "Programada"):
        if clave in texto:
            return ESTADO_COLORES[clave]
    return ESTADO_COLORES.get(texto, ("EDEDED", "333333"))


def _badge_estado(valor) -> str:
    """Estado como badge accesible (símbolo + texto), sin depender solo del color."""
    t = str(valor)
    if "Reprogramada_Finalizada" in t or "Reprogramada / Finalizada" in t:
        return "🟪 Reprogramada_Finalizada"
    if "Finalizada" in t:
        return "🟩 Finalizada"
    if "En proceso" in t:
        return "🟨 En proceso"
    if "Reprogramada" in t:
        return "🟪 Reprogramada"
    if "Programada" in t:
        return "⬜ Programada"
    return "⬜ " + (t if t.strip() else "Sin cronograma")


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
    ancho = {"SBAN": 10, "Oficina": 40, "Estado": 20, "Estado cronograma": 22,
             "Total elementos": 14, "Subsanados (MT)": 15, "Pendientes": 12,
             "% Avance": 12, "Categoría Novedad": 22, "Novedades del día": 14,
             "Novedades acumuladas": 16, "Obs. novedad": 45}
    for j, h in enumerate(hdrs, 1):
        ws.column_dimensions[get_column_letter(j)].width = ancho.get(h, 14)
    for i, row in enumerate(res.itertuples(index=False), start=2):
        es_total = str(row[0]) == "TOTAL"
        for j, val in enumerate(row, 1):
            c = ws.cell(i, j, val)
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                c.number_format = '0.0"%"' if hdrs[j - 1] == "% Avance" else "#,##0"
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

    cab = ["Estado", "Subsanados (MT)", "Pendientes"]
    orden = ["Programada", "En proceso", "Finalizada", "Reprogramada", "Sin cronograma"]
    rows = res[res.iloc[:, 0].astype(str) != "TOTAL"].copy()
    col_est = "Estado" if "Estado" in rows.columns else "Estado cronograma"
    rows["_estado"] = (rows[col_est].astype(str)
                       .str.replace(r"^[^A-Za-zÁÉÍÓÚáéíóúñ]+", "", regex=True).str.strip())
    ag = rows.groupby("_estado")[["Subsanados (MT)", "Pendientes"]].sum()
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
        for key, sub in base.groupby("_SBAN", sort=True):
            total = len(sub)
            subsanados = int(sub["_mt"].sum())
            pendientes = total - subsanados
            avance = (subsanados / total * 100) if total else 0.0
            estados = sorted({str(x) for x in sub["_est_crono"].dropna().tolist() if str(x).strip()})
            sban_ofi = str(sub["_SBAN"].iloc[0])
            nombres_ofi = [str(x).strip() for x in dict.fromkeys(sub["Oficina"].dropna())
                           if str(x).strip()]
            etiqueta_ofi = (" / ".join(nombres_ofi[:2])
                            + (" …" if len(nombres_ofi) > 2 else "")) or "SIN NOMBRE"
            cat_nov, n_dia, n_acum, obs_nov = nov_lookup.get(sban_ofi, ("", 0, 0, ""))
            filas.append({
                "SBAN": sban_ofi,
                "Oficina": etiqueta_ofi,
                "Estado": _badge_estado(" / ".join(estados) if estados else "Sin cronograma"),
                "Estado cronograma": " / ".join(estados) if estados else "Sin cronograma",
                "Total elementos": total,
                "Subsanados (MT)": subsanados,
                "Pendientes": pendientes,
                "% Avance": round(avance, 1),
                "Categoría Novedad": ("🟠 " + cat_nov) if cat_nov else "—",
                "Novedades del día": n_dia,
                "Novedades acumuladas": n_acum,
                "Obs. novedad": "" if MODO_PUBLICO else obs_nov,
            })
        res = pd.DataFrame(filas)
        datos_resumen = res.copy()  # sin fila TOTAL

        # ---------------- Filtros intuitivos de la vista ----------------
        # Valores por defecto ANTES de crear los widgets: así el slider siempre tiene
        # un valor definido y el reinicio por callback no choca con su instanciación.
        for _k, _v in VALORES_FILTRO.items():
            if _k.startswith("res_") and _v is not None:
                st.session_state.setdefault(_k, _v)
        estados_posibles = ["Programada", "En proceso", "Finalizada", "Reprogramada",
                            "Reprogramada_Finalizada", "Sin cronograma"]
        tokens = {tok.strip() for v in datos_resumen["Estado cronograma"]
                  for tok in str(v).split("/") if tok.strip()}
        presentes = ([e for e in estados_posibles if e in tokens]
                     + [e for e in sorted(tokens) if e not in estados_posibles])
        f1, f2, f3 = st.columns([2.2, 1.6, 1])
        with f1:
            sel_est = st.multiselect("🗂️ Estado cronograma", options=presentes, default=presentes,
                                     key="res_estado", placeholder="Todos los estados")
        with f2:
            min_av = st.slider("🎯 Avance mínimo (%)", 0, 100, step=5,
                               key="res_min_av", help="Muestra oficinas con avance igual o mayor")
        with f3:
            solo_pend = st.checkbox("Solo pendientes\n(< 100%)", key="res_solo_pend")

        vis = datos_resumen.copy()
        vis["% Avance"] = pd.to_numeric(vis["% Avance"], errors="coerce").fillna(0)
        if sel_est:
            vis = vis[vis["Estado cronograma"].apply(
                lambda e: any(t.strip() in sel_est for t in str(e).split("/")))]
        vis = vis[vis["% Avance"] >= min_av]
        if solo_pend:
            vis = vis[vis["% Avance"] < 100.0]
        vis = vis.reset_index(drop=True)

        # ---------------- Filtros de novedades ----------------
        if "Novedades del día" in datos_resumen.columns:
            g1, g2 = st.columns([2.2, 1])
            cats_nov = sorted({str(c) for c in datos_resumen["Categoría Novedad"]
                               if str(c).strip() and str(c).strip().lower()
                               not in ("nan", "sin novedad", "—")})
            with g1:
                sel_cat = st.multiselect("🏷️ Categoría Novedad", options=cats_nov,
                                         default=[], key="res_catnov",
                                         placeholder="Todas las categorías (sin filtrar)")
            with g2:
                solo_nov = st.checkbox("Solo novedades del día", key="res_solo_nov")
            if sel_cat:
                vis = vis[vis["Categoría Novedad"].astype(str).isin(sel_cat)]
            if solo_nov:
                vis = vis[pd.to_numeric(vis["Novedades del día"], errors="coerce").fillna(0) > 0]
            vis = vis.reset_index(drop=True)

        info_f, btn_f = st.columns([3, 1])
        with info_f:
            st.caption("ℹ️ Los filtros de **este resumen** son independientes de los KPIs "
                       "superiores (que usan el módulo y los filtros del sidebar).")
        with btn_f:
            # on_click: el reinicio corre ANTES de crear los widgets en el siguiente
            # rerun. Escribir en session_state aquí dentro lanzaría
            # StreamlitWidgetAlreadyInstantiatedError (los filtros de arriba ya existen).
            st.button("♻️ Restablecer filtros del resumen", key=f"reset_res_{tipo}",
                      width="stretch", on_click=_limpiar_filtros_resumen)

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
        fila_total = {c: "" for c in vis.columns}
        fila_total.update({"SBAN": "TOTAL", "Oficina": f"{n_vis} oficina(s)", "Estado": "—",
                           "Total elementos": vt, "Subsanados (MT)": vs, "Pendientes": vp,
                           "% Avance": round(va, 1), "Categoría Novedad": "—",
                           "Novedades del día": nd, "Novedades acumuladas": na})
        vis.loc[n_vis] = fila_total
        if "Estado cronograma" in vis.columns:
            vis = vis.drop(columns=["Estado cronograma"])
        if MODO_PUBLICO and "Obs. novedad" in vis.columns:
            vis = vis.drop(columns=["Obs. novedad"])

        st.caption(
            f"Atribución: **{etiqueta}** · Mostrando **{n_vis}** de {len(datos_resumen)} oficinas · "
            "Cada fila agrupa SBAN + nombre de oficina."
        )
        st.dataframe(
            vis,
            width="stretch",
            hide_index=True,
            height=min(460, 38 + 34 * len(vis)),
            column_config={
                "SBAN": st.column_config.TextColumn("SBAN", width="small"),
                "Oficina": st.column_config.TextColumn("Oficina", width="large"),
                "Estado": st.column_config.TextColumn(
                    "Estado", width="small",
                    help="🟩 Finalizada · 🟨 En proceso · ⬜ Programada · 🟪 Reprogramada"),
                "Total elementos": st.column_config.NumberColumn("Total", format="%d", width="small"),
                "Subsanados (MT)": st.column_config.NumberColumn("Subsanados", format="%d", width="small"),
                "Pendientes": st.column_config.NumberColumn("Pendientes", format="%d", width="small"),
                "% Avance": st.column_config.ProgressColumn(
                    "Avance MT3", help="Porcentaje de equipos con consecutivo MT",
                    format="%.1f%%", min_value=0, max_value=100),
                "Categoría Novedad": st.column_config.TextColumn("Categoría novedad", width="medium"),
                "Novedades del día": st.column_config.NumberColumn(
                    "Nov. del día", format="%d", width="medium",
                    help="Novedades registradas en la fecha de corte (hoy)"),
                "Novedades acumuladas": st.column_config.NumberColumn(
                    "Nov. acumuladas", format="%d", width="medium",
                    help="Novedades acumuladas desde el inicio del registro"),
                "Obs. novedad": st.column_config.TextColumn("Obs. novedad", width="large"),
            },
        )
        m1, m2, m3 = st.columns(3)
        m1.metric("🏢 Oficinas mostradas", f"{n_vis:,}".replace(",", "."), border=True,
                  help="Filas de la tabla (una por SBAN) que cumplen los filtros del resumen. "
                       "El detalle por sede está en la tabla de abajo.")
        m2.metric("🏷️ Novedades del día", f"{nd:,}".replace(",", "."), border=True,
                  help="Novedades de las oficinas con la fecha más reciente del archivo "
                       "(hoja 'Novedades Equipos').")
        m3.metric("📚 Novedades acumuladas", f"{na:,}".replace(",", "."), border=True,
                  help="Total histórico de novedades registradas para esas oficinas.")
        st.caption("Los conteos de **equipos** (Total · Subsanados · Pendientes · Avance) están en "
                   "las tarjetas KPI de arriba; aquí solo se repiten **por fila** en la tabla.")
        st.download_button(
            "⬇️ Descargar Excel (vista filtrada + gráficos + novedades)",
            data=generar_excel_resumen(vis),
            file_name="Resumen_por_oficina_MT3.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
        st.caption("Estado en badges: 🟩 Finalizada · 🟨 En proceso · ⬜ Programada · "
                   "🟪 Reprogramada · 🟠 Categoría de novedad presente · "
                   "barra de progreso en la columna Avance MT3.")


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
    aplicar_css()
    barra_tema()
    try:
        df = cargar_datos()
    except FileNotFoundError as exc:
        st.error(f"❌ {exc}")
        st.stop()
    except Exception as exc:  # noqa: BLE001 - capturar errores de lectura
        st.error(f"❌ No se pudo cargar el archivo Excel: {exc}")
        st.stop()

    estados_crono = cargar_estados_crono()

    # Componente (C1 / C2 láser / UPS) ANTES de preparar Campos: así las novedades se
    # amarran por Serial al componente REAL del equipo y no solo por el texto de la
    # observación (antes se calculaba después y el mapa por Serial quedaba vacío).
    df["_componente"] = clasificar_componente(df)

    # --- Campos dashboard: Estado de la sede, Categoría Novedad (Y) y novedades ---
    kpi_campos, nov_campos = cargar_campos()
    kpi_norm, nov_norm = preparar_campos(kpi_campos, nov_campos, df)
    crono_ups_raw = cargar_crono_ups()
    crono_ups = preparar_crono_ups(crono_ups_raw)
    fecha_corte = (nov_norm["_FECHA"].max() if (nov_norm is not None and not nov_norm.empty)
                   else None)

    # Estado por SBAN desde Campos dashboard (col N). Se usa la MISMA regla de prioridad
    # que el resto del panel (`_mejor_por_sban`), para que no exista una segunda
    # implementación del criterio que pueda desincronizarse.
    estados_sede = {}
    if kpi_norm is not None and not kpi_norm.empty and "Estado sede" in kpi_norm.columns:
        estados_sede = _mejor_por_sban(
            zip(kpi_norm["_SBAN"].astype(str), kpi_norm["Estado sede"]), _PRIO_SEDE)

    def _estado_sban(v):
        if pd.isna(v):
            return None
        sban = _pad5(v)
        return estados_sede.get(sban) or estados_crono.get(int(v))

    df["_est_crono"] = df["SBAN"].map(_estado_sban).fillna("Sin cronograma")

    # Regla de negocio opcional: impresoras láser (C2) tratadas como facturables
    if IMPRESORAS_FACTURABLES:
        df.loc[df["_componente"].eq("C2"), "Facturable"] = "Si"

    # --- Selector de módulo (se elige antes que los filtros) ---
    # El reinicio de filtros al cambiar de módulo se hace con on_change (callback),
    # nunca escribiendo en session_state después de crear los widgets de filtro.
    opciones_mod = ["Componente 1", "Componente 2 (Impresoras láser)", "UPS"]
    if hasattr(st, "segmented_control"):
        modulo = st.segmented_control("Módulo", opciones_mod,
                                      default=opciones_mod[0], key="modulo",
                                      on_change=_reset_filtros_por_modulo)
    else:
        modulo = st.radio("Módulo", opciones_mod, horizontal=True, key="modulo",
                          on_change=_reset_filtros_por_modulo)
    modulo = modulo or opciones_mod[0]
    cod_mod = {"Componente 1": "C1", "Componente 2 (Impresoras láser)": "C2",
               "UPS": "UPS"}[modulo]

    seleccion, solo_pendientes, f_attr = render_sidebar(df, cod_mod)

    # UPS: el avance por sede viene de la columna AP «ESTADO UPS» del archivo Campos
    # dashboard (Banco Agrario + COLSOF en el mismo documento). Solo si ese archivo no
    # trae la columna, se usa como respaldo el cronograma UPS (fecha + avance en MT).
    ups_fuente = ""
    if cod_mod == "UPS":
        if kpi_norm is not None and not kpi_norm.empty and "Estado UPS" in kpi_norm.columns:
            mapa_ups = {str(r["_SBAN"]): normalizar_estado_ups(r.get("Estado UPS"))
                        for _, r in kpi_norm.iterrows()}
            con_dato = {s: e for s, e in mapa_ups.items() if str(e).strip()}
            if con_dato:
                df["_est_ups"] = df["_SBAN"].map(mapa_ups).fillna("")
                df["_est_crono"] = df["_SBAN"].map(mapa_ups).fillna("")
                sin_rep = df["_est_crono"].astype(str).str.strip().eq("")
                df.loc[sin_rep, "_est_crono"] = "Sin cronograma"
                ups_fuente = "lo reportado por la oficina en el archivo Campos dashboard"
        if not ups_fuente and crono_ups is not None and not crono_ups.empty:
            est_ups = estados_ups(df, crono_ups)
            df["_est_crono"] = df["_SBAN"].map(est_ups).fillna("Sin cronograma")
            df["_est_ups"] = df["_SBAN"].map(est_ups).fillna("")
            ups_fuente = "el cronograma de UPS (respaldo: fecha + avance)"

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
    click_sban = st.session_state.get("click_sban")
    if click_sban:
        base_oficina = base_oficina[base_oficina["_SBAN"].eq(click_sban)]

    filtrado = base_oficina
    if solo_pendientes:
        filtrado = base_oficina[base_oficina["_pendiente"]].copy()

    if filtrado.empty:
        base_mod = df[df["_componente"].eq(cod_mod)]
        if seleccion:
            st.warning(f"Las oficinas seleccionadas no tienen elementos de **{modulo}**. "
                       "Quita el filtro de Oficina para ver el módulo completo.")
            st.button("🧹 Quitar filtro de oficinas", key="limpiar_oficinas",
                      on_click=_quitar_filtro, args=("f_oficinas",))
        elif solo_pendientes and not base_mod.empty and bool(base_mod["_mt"].all()):
            st.success(f"🎉 Todos los elementos de **{modulo}** ya tienen MT3.")
        elif f_attr != "T" and not base_mod.empty:
            nombre_atr = "BANCO (Facturable Si)" if f_attr == "Si" else "COLSOF (No facturable)"
            st.warning(f"El módulo **{modulo}** no tiene elementos **{nombre_atr}**. "
                       "Cambia la atribución a **Todos** para verlo.")
            st.button("🧹 Ver todos (atribución Todos)", key="limpiar_atrib",
                      on_click=_quitar_filtro, args=("f_atrib",))
        else:
            st.warning(f"Sin datos para **{modulo}** con los filtros seleccionados. "
                       "Revisa el filtro de Oficina o la atribución.")
        st.stop()

    # --- Validación de integridad (siempre; detiene si hay críticos) ---
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

    # --- Selector de vista: Ejecutiva (solo visuales) u Operativa (detalle) ---
    if hasattr(st, "segmented_control"):
        vista = st.segmented_control("Vista", ["Operativa", "Ejecutiva"],
                                     default="Operativa", key="vista_v2")
    else:
        vista = st.radio("Vista", ["Operativa", "Ejecutiva"], horizontal=True,
                         key="vista_v2")
    vista = vista or "Operativa"
    if vista == "Ejecutiva":
        # Mismo bloque de KPIs principales que la vista operativa (lenguaje visual unificado).
        # Se pasa `base_oficina` (sin el filtro de tabla "solo pendientes") para que el
        # numerador de "Oficinas intervenidas" no quede en 0.
        tarjetas_destacadas(base_oficina)
        render_vista_ejecutiva(filtrado, df, seleccion, f_attr,
                               st.session_state.get("click_sban"))
        st.markdown("---")
        st.caption(f"Vista Ejecutiva · Módulo {modulo} · 1 fila = 1 elemento (Serial único).")
        return

    # --- Operativa: encabezado hero + KPIs ---
    scope = "Todas las oficinas" if not seleccion else f"{len(seleccion)} oficina(s)"
    atrib = {"T": "Todos", "Si": "BANCO (Facturable Si)", "No": "COLSOF (No facturable)"}[f_attr]
    encabezado_hero(modulo, scope, atrib, len(filtrado), conteos, nov_resumen, fecha_corte)

    # --- Barra VISIBLE de facturación (filtro a la vista) ---
    # Los conteos ya están en las tarjetas KPI de abajo: aquí solo va el filtro.
    barra_facturacion(df_mod, mostrar_conteos=False)

    if avisos:
        with st.expander(f"🔍 Validación de integridad — {len(avisos)} aviso(s)",
                         expanded=False):
            for msg in avisos:
                st.warning(msg)

    # --- KPIs PRINCIPALES (de primera): oficinas intervenidas y avance MT3 ---
    # `base_oficina` (sin "solo pendientes"): ese filtro es de tabla y dejaría el
    # numerador en 0; los KPIs de cabecera describen el módulo, no la tabla.
    tarjetas_destacadas(base_oficina)

    # --- KPIs de detalle (cada bloque se dibuja UNA sola vez) ---
    render_kpis(filtrado)
    render_kpis_oficinas(df, seleccion, f_attr, st.session_state.get("click_sban"),
                         kpi_campos, cod_mod=cod_mod, solo_pendientes=solo_pendientes)

    if cod_mod == "UPS" and ups_fuente:
        st.caption(f"🔋 Avance de UPS leído de: **{ups_fuente}**.")

    # --- Filtros activos (chips con X) ---
    barra_filtros_activos(seleccion, solo_pendientes, f_attr,
                          st.session_state.get("click_sban"))

    # --- Pestañas: Gráficos · Resumen por oficina · Avance UPS · Novedades · Gestión ---
    cfg_chart = {"displaylogo": False,
                 "modeBarButtonsToRemove": ["lasso2d", "select2d"]}
    cfg_sel = {"displaylogo": False,
               "modeBarButtonsToRemove": ["lasso2d", "select2d", "zoom2d", "pan2d"]}

    nombres_tabs = ["📈 Gráficos y avance", "🏢 Resumen por oficina"]
    if cod_mod == "UPS":
        nombres_tabs.append("🔋 UPS: avance y diferencias")
    nombres_tabs += ["📰 Novedades de las oficinas", "📋 Gestión de novedades"]
    tabs = st.tabs(nombres_tabs)
    tab_graf, tab_res = tabs[0], tabs[1]
    i = 2
    tab_ups = None
    if cod_mod == "UPS":
        tab_ups = tabs[i]
        i += 1
    tab_nov_ofi, tab_nov = tabs[i], tabs[i + 1]

    with tab_graf:
        # Un solo gráfico de avance: la dona (el taquímetro repetía la misma cifra).
        st.plotly_chart(
            grafico_dona(int(filtrado["_mt"].sum()), int(filtrado["_pendiente"].sum())),
            width="stretch", config=cfg_chart)

        fila2a, fila2b = st.columns([1, 1.3])
        with fila2a:
            top_n = st.segmented_control("Sedes a mostrar en el ranking",
                                         options=[5, 10, 12, 15, 20, 25], default=12,
                                         key="top_n_pend") or 12
            sel = st.plotly_chart(
                grafico_top_pendientes(filtrado, top_n=top_n), width="stretch",
                config=cfg_sel, key="chart_top", on_select="rerun", selection_mode="points")
            try:
                pts = sel.selection.points if (sel is not None and getattr(sel, "selection", None)) else []
            except Exception:  # noqa: BLE001
                pts = []
            if pts:
                etiqueta = str(pts[0].get("y") or pts[0].get("x") or "")
                candidato = etiqueta[:5]
                if candidato.isdigit() and st.session_state.get("click_sban") != candidato:
                    # El widget de oficinas ya se creó en este run: si la sede no está
                    # en su lista de opciones, asignar 'click_sban' directamente lanzaría
                    # StreamlitWidgetAlreadyInstantiatedError. En ese caso se limpia
                    # primero la selección (el widget se recrea en el rerun siguiente).
                    oficinas = st.session_state.get("f_oficinas") or []
                    if oficinas and not any(str(o).startswith(f"{candidato} - ")
                                            for o in oficinas):
                        st.session_state["f_oficinas"] = []
                    st.session_state["click_sban"] = candidato
                    st.rerun()
            st.caption("💡 Haz clic en una barra para filtrar el panel por esa sede.")
        with fila2b:
            fig_t = grafico_tendencia(filtrado)
            if fig_t is not None:
                st.plotly_chart(fig_t, width="stretch", config=cfg_chart)
            else:
                st.info("Sin fechas de MT3 para la selección actual.")

        fig_h = grafico_heatmap(filtrado)
        if fig_h is not None:
            st.plotly_chart(fig_h, width="stretch", config=cfg_chart)

    with tab_res:
        # Resumen por oficina (Total / Subsanados / Pendientes / % Avance / Novedades)
        render_resumen(base_oficina, seleccion, f_attr, nov_resumen)

    if tab_ups is not None:
        with tab_ups:
            render_avance_ups(kpi_norm, df, seleccion)

    with tab_nov_ofi:
        # Novedades de las oficinas: TODAS (no se aísla por módulo, porque una novedad
        # de oficina —denuncio, recolección, cambio a facturable— es transversal).
        render_novedades_oficinas(nov_norm, seleccion, fecha_corte)

    with tab_nov:
        # Tabla de gestión de novedades
        render_tabla(filtrado, seleccion, solo_pendientes)

    st.markdown("---")
    st.caption(f"Mantenimiento Preventivo 3 · Módulo {modulo} · 1 fila = 1 elemento (Serial único).")


if __name__ == "__main__":
    # INGESTA_SIN_APP=1 permite importar este módulo (para reutilizar sus validaciones
    # desde ingesta.py) sin arrancar el panel de Streamlit.
    if os.environ.get("INGESTA_SIN_APP") != "1":
        main()
