# -*- coding: utf-8 -*-
"""
Genera en data/ una versión ANONIMIZADA de los archivos para repositorios públicos:
- Serial -> hash (se conserva la unicidad, no el dato)
- Placa -> enmascarada (****1234)
- Se eliminan columnas con datos personales (nombres, cédulas, correos, teléfonos,
  direcciones, observaciones con nombres).
El panel funciona igual: KPIs, estados, novedades y conteos.
"""
import hashlib
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent
ORIGEN = BASE.parent / "Ingesta de datos diaria"
CRONO_FALLBACK = BASE.parent / "Cronograma Mto Preventivo 3 _ Equipos 1.xlsx"
DATA = BASE / "data"
DATA.mkdir(exist_ok=True)


def _hash(v) -> str:
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return ""
    return "H_" + hashlib.sha256(s.encode("utf-8")).hexdigest()[:12]


def _mask(v) -> str:
    s = str(v).strip()
    return "****" + s[-4:] if len(s) > 4 else s


def _elegir() -> dict:
    out = {}
    for p in sorted(ORIGEN.glob("*.xlsx"), key=lambda x: x.stat().st_mtime, reverse=True):
        try:
            xl = pd.ExcelFile(p)
        except Exception:  # noqa: BLE001
            continue
        if "Dashboard_KPI" in xl.sheet_names:
            out.setdefault("campos", p)
            continue
        if "UPS" in xl.sheet_names or "UPS" in p.name.upper():
            out.setdefault("crono_ups", p)
            continue
        cols = set(map(str, xl.parse(xl.sheet_names[0], nrows=0).columns))
        if {"Consecutivo mantenimiento 3", "Facturable"} <= cols:
            out.setdefault("data", p)
        elif {"ESTADO", "Fecha Inicio", "Fecha Fin", "tipo"} <= cols:
            out.setdefault("crono", p)
    if "crono" not in out and CRONO_FALLBACK.exists():
        out["crono"] = CRONO_FALLBACK
    return out


def sanear_data(df: pd.DataFrame) -> pd.DataFrame:
    cols = ["Categoría", "Modelo", "Serial", "Placa", "SBAN", "Oficina", "Tipo ubicación",
            "Regional", "Consecutivo mantenimiento 3", "Estado", "Estado interno",
            "Facturable", "Fecha de mantenimiento 3"]
    df = df[[c for c in cols if c in df.columns]].copy()
    if "Serial" in df:
        df["Serial"] = df["Serial"].apply(_hash)
    if "Placa" in df:
        df["Placa"] = df["Placa"].apply(_mask)
    return df


def sanear_campos(path: Path):
    kpi = pd.read_excel(path, sheet_name="Dashboard_KPI")
    try:
        nov = pd.read_excel(path, sheet_name="Novedades Equipos")
    except Exception:  # noqa: BLE001
        nov = pd.DataFrame()
    # 'ESTADO UPS' (col AP) se conserva: es el avance de las UPS que muestra el panel.
    kcols = ["SBAN", "Tipo", "Nombre Oficina", "Fecha Inicio", "Fecha Fin",
             "Estado de la sede", "ESTADO UPS", "Categoria  Novedad"]
    kpi = kpi[[c for c in kcols if c in kpi.columns]].copy()
    if "Categoria  Novedad" in kpi.columns:
        kpi["Categoria  Novedad"] = kpi["Categoria  Novedad"].fillna("").astype(str).str.strip()
    ncols = {"SBAN": "SBAN", "SBAN PCT": "SBAN PCT", "Fecha": "Fecha",
             "Serial ": "Serial ", "Estado": "Estado", "Facturable": "Facturable",
             "Nombre Oficina": "Nombre Oficina", "Departamento": "Departamento",
             "ALIADO": "ALIADO", "Categoria  Novedad": "Categoria  Novedad"}
    nov = nov[[c for c in ncols if c in nov.columns]].copy()
    for c in list(nov.columns):
        if str(c).strip().lower() == "serial":
            nov[c] = nov[c].apply(_hash)
    return kpi, nov


def main():
    t = _elegir()
    if "data" in t:
        d = pd.read_excel(t["data"], sheet_name="Hoja1")
        sanear_data(d).to_excel(DATA / "Data_actual.xlsx", index=False, sheet_name="Hoja1")
        print("Data anonimizada:", len(d), "filas")
    if "campos" in t:
        kpi, nov = sanear_campos(t["campos"])
        with pd.ExcelWriter(DATA / "Campos_actual.xlsx", engine="openpyxl") as w:
            kpi.to_excel(w, sheet_name="Dashboard_KPI", index=False)
            nov.to_excel(w, sheet_name="Novedades Equipos", index=False)
        print("Campos anonimizado: KPI", len(kpi), "| Novedades", len(nov))
    if "crono" in t:
        c = pd.read_excel(t["crono"], sheet_name="Hoja1")
        keep = ["SBAN", "tipo", "Nombre Oficina", "Jefaturas  Operaciones Regional",
                "ESTADO", "Fecha Inicio", "Fecha Fin", "Equipo de escritorio",
                "Equipo portátil", "Escáner", "Impresora", "Lector biométrico",
                "Lector de banda pin pad", "Lector de código", "Monitor",
                "PAD de firmas", "Servidor", "Tablet"]
        c[[k for k in keep if k in c.columns]].to_excel(DATA / "Cronograma_actual.xlsx",
                                                        index=False, sheet_name="Hoja1")
        print("Cronograma anonimizado:", len(c))
    if "crono_ups" in t:
        xl = pd.ExcelFile(t["crono_ups"])
        hoja = "UPS" if "UPS" in xl.sheet_names else xl.sheet_names[0]
        u = xl.parse(hoja)
        keep_u = ["SBAN", "TIPO", "Nombre Oficina", "Jefaturas  Operaciones Regional",
                  "FECHA", "UPS"]
        u[[k for k in keep_u if k in u.columns]].to_excel(
            DATA / "Cronograma_UPS_actual.xlsx", index=False, sheet_name="UPS")
        print("Cronograma UPS anonimizado:", len(u))
    _escribir_metadata()
    print("Listo. Revisar data/ y hacer commit/push.")


def _escribir_metadata():
    import json
    info = {}
    for nombre in ("Data_actual.xlsx", "Campos_actual.xlsx", "Cronograma_actual.xlsx",
                   "Cronograma_UPS_actual.xlsx"):
        p = DATA / nombre
        if p.exists():
            info[nombre] = {"sha256": hashlib.sha256(p.read_bytes()).hexdigest(),
                            "bytes": p.stat().st_size}
    meta = {"fecha_hora": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
            "anonimizado": True, "archivos": info}
    try:
        meta["filas_data"] = int(len(pd.read_excel(DATA / "Data_actual.xlsx",
                                                   sheet_name="Hoja1")))
    except Exception:  # noqa: BLE001
        pass
    try:
        meta["filas_novedades"] = int(len(pd.read_excel(DATA / "Campos_actual.xlsx",
                                                        sheet_name="Novedades Equipos").dropna(how="all")))
    except Exception:  # noqa: BLE001
        pass
    (DATA / "ultima_actualizacion.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print("Última actualización:", meta["fecha_hora"])


if __name__ == "__main__":
    main()
