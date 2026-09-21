# -*- coding: utf-8 -*-
"""
Genera dashboard_web/data.js (JSON) para la app web del control MT3.
Origen: Data_PCTriage.xlsx (Hoja1). 1 fila = 1 elemento (Serial unico).
Reglas: AN='Si' -> BANCO | AN='No' -> COLSOF | Pendientes = Elementos - Con MT3
"""
import json
import os
from datetime import timezone

import pandas as pd

try:
    import zoneinfo
    TZ_BO = zoneinfo.ZoneInfo("America/Bogota")
except Exception:
    TZ_BO = timezone.utc

def ahora() -> pd.Timestamp:
    return pd.Timestamp.now(tz=TZ_BO).tz_convert(None)

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, "Data_PCTriage.xlsx")
OUT_DIR = os.path.join(BASE, "dashboard_web")
os.makedirs(OUT_DIR, exist_ok=True)

df = pd.read_excel(SRC, sheet_name="Hoja1")
df["Serial"] = df["Serial"].astype(str)
df["Regional"] = df["Regional"].fillna("SIN REGIONAL")
df["Oficina"] = df["Oficina"].fillna("SIN NOMBRE")
df["G"] = df["Consecutivo mantenimiento 3"]
df["H"] = pd.to_datetime(df["Fecha de mantenimiento 3"], errors="coerce")
df["F"] = df["Facturable"].fillna("")
df["tiene"] = df["G"].notna()

def s5(x):
    return str(int(x)).zfill(5)

# Tickets MT3 que cruzan mas de un SBAN
tk_sban = df[df["G"].notna()].groupby("G")["SBAN"].nunique()
shared = set(tk_sban[tk_sban > 1].index)
sban_por_tk = df[df["G"].isin(shared)].groupby("G")["SBAN"].apply(
    lambda s: sorted(set(int(x) for x in s)))

# ---- Control por SBAN ----
control = []
for sban, sub in df.groupby("SBAN"):
    e = len(sub)
    c = int(sub["G"].notna().sum())
    p = e - c
    eleS = int((sub["F"] == "Si").sum())
    eleN = int((sub["F"] == "No").sum())
    cS = int(((sub["F"] == "Si") & sub["G"].notna()).sum())
    cN = int(((sub["F"] == "No") & sub["G"].notna()).sum())
    tks = int(sub["G"].dropna().nunique())
    val = float(sub.loc[sub["F"] == "Si", "Valor facturación"].sum())
    reg = sub["Regional"].mode().iloc[0]
    ofi = sub["Oficina"].mode().iloc[0]
    sh = sorted({t for t in sub["G"].dropna() if t in shared})
    obs = []
    if p > 0:
        obs.append("FALTAN %d por MT3" % p)
    elif c > 0:
        obs.append("Completo MT3")
    if sh:
        obs.append("SOBRA? tickets %s en otros SBAN" % ",".join(str(t) for t in sh[:3]))
    if ofi == "SIN NOMBRE":
        obs.append("Oficina sin nombre")
    if sub["Oficina"].nunique() > 1:
        obs.append("SBAN con varios nombres de oficina")
    control.append({
        "s": s5(sban), "o": ofi, "r": reg, "e": e, "c": c, "p": p,
        "k": tks, "v": val, "obs": "; ".join(obs),
        "eS": eleS, "cS": cS, "eN": eleN, "cN": cN,
    })
control.sort(key=lambda x: (x["r"], -x["p"], x["s"]))

# ---- Faltantes (elementos sin MT3) ----
falt_hdr = ["Regional", "SBAN", "Oficina", "Tipo ubicacion", "Categoria",
            "Modelo", "Serial", "Placa", "Estado", "Estado interno", "Facturable"]
falt = df[~df["tiene"]].copy()
falt = falt.sort_values(["Regional", "SBAN"])
faltantes = []
for _, r in falt.iterrows():
    pv = r["Placa"]
    if pd.isna(pv):
        pla = ""
    else:
        try:
            pla = str(int(float(str(pv)))).zfill(8)
        except ValueError:
            pla = str(pv)
    faltantes.append([r["Regional"], s5(r["SBAN"]), r["Oficina"], r["Tipo ubicación"],
                      r["Categoría"], str(r["Modelo"]), r["Serial"], pla,
                      str(r["Estado"]), str(r["Estado interno"]) if pd.notna(r["Estado interno"]) else "",
                      r["F"]])

# ---- Diario MT3 ----
dia = df[df["H"].notna()].groupby(df["H"].dt.date).size().sort_index()
diario = {"labels": [d.strftime("%d/%m/%Y") for d in dia.index],
          "datos": [int(v) for v in dia.values]}

# ---- Alertas: tickets compartidos ----
alertas = []
for tk in sorted(shared):
    sb = sban_por_tk[tk]
    nfilas = int((df["G"] == tk).sum())
    alertas.append({"g": tk, "sbans": [s5(x) for x in sb],
                    "n": len(sb), "filas": nfilas})

# ---- Resumen ----
resumen = {
    "totalEle": int(len(df)), "totalCon": int(df["tiene"].sum()),
    "totalPend": int((~df["tiene"]).sum()),
    "banco": int((df["F"] == "Si").sum()), "colsof": int((df["F"] == "No").sum()),
    "conBanco": int(((df["F"] == "Si") & df["tiene"]).sum()),
    "conColsof": int(((df["F"] == "No") & df["tiene"]).sum()),
    "tickets": int(df["G"].dropna().nunique()),
    "valor": float(df.loc[df["F"] == "Si", "Valor facturación"].sum()),
    "generado": ahora().strftime("%Y-%m-%d %H:%M"),
}

data = {"resumen": resumen, "control": control, "faltantes": faltantes,
        "falt_hdr": falt_hdr, "diario": diario, "alertas": alertas}

with open(os.path.join(OUT_DIR, "data.js"), "w", encoding="utf-8") as f:
    f.write("window.DATA = " + json.dumps(data, ensure_ascii=False,
            separators=(",", ":")) + ";")

print("OK data.js | control:", len(control), "| faltantes:", len(faltantes),
      "| alertas:", len(alertas))
