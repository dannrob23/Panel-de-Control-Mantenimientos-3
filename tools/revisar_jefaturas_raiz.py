# -*- coding: utf-8 -*-
"""REVISION DE LA COLUMNA AA (SEDES-JEFATURA) del Excel de ejecucion.

Uso:   python tools\\revisar_jefaturas.py [ruta.xlsx]

Revisa la columna AA («SBAN») del archivo de Data mas reciente y reporta:

  · las sedes-jefatura que trae, con su codigo, equipos, MT3 hechos y pendientes;
  · AVISOS a mirar antes de publicar:
      - nombres mal escritos que el panel corrige al mostrar (BQUILLA -> BARRANQUILLA),
      - nombres que no corresponden a ninguna sede-jefatura conocida (posible error nuevo),
      - la misma sede escrita de dos formas distintas (p. ej. BQUILLA y BAQUILLA),
      - que la columna AA no traiga ninguna fila de jefatura.

Codigos de salida:  0 = sin avisos · 1 = hay avisos · 2 = no se pudo revisar.
No escribe ni publica nada: solo lee y reporta.
"""
import contextlib
import io
import os
import sys
from difflib import SequenceMatcher
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))
os.environ["INGESTA_SIN_APP"] = "1"  # importar app.py sin arrancar Streamlit
for _f in (sys.stdout, sys.stderr):
    try:
        _f.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import pandas as pd  # noqa: E402

_buf = io.StringIO()  # app.py importa Streamlit y avisa en modo consola: se silencia
with contextlib.redirect_stderr(_buf):
    import app  # noqa: E402
import ingesta as ing  # noqa: E402

# Sedes-jefatura conocidas, escritas CORRECTAMENTE. Si el archivo trae un nombre que no
# está aquí (y se parece a uno de estos), se avisa: puede ser un error de escritura o una
# sede nueva. Para una sede nueva de verdad, agregarla a esta lista.
SEDES_JEFATURA = [
    "BARRANQUILLA", "BOGOTA", "BUCARAMANGA", "CALI", "CUCUTA", "IBAGUE", "MANIZALES",
    "MEDELLIN", "MONTERIA", "NEIVA", "PASTO", "POPAYAN", "SINCELEJO", "TUNJA",
    "VALLEDUPAR", "VILLAVICENCIO",
]


def parecido(nombre: str) -> str:
    """Sede conocida mas parecida al nombre ('' si no se parece a ninguna)."""
    mejor, largo = "", 0
    for cand in SEDES_JEFATURA:
        m = SequenceMatcher(None, nombre, cand).find_longest_match(
            0, len(nombre), 0, len(cand))
        if m.size > largo:
            mejor, largo = cand, m.size
    # Se exige un bloque comun relevante: al menos 4 letras y la mitad del nombre.
    if mejor and largo >= 4 and largo >= 0.5 * len(nombre):
        return mejor
    return ""


def main() -> int:
    if len(sys.argv) > 1:
        ruta = Path(sys.argv[1])
    else:
        archivos = ing.elegir_archivos(ing.ORIGEN)
        ruta = archivos.get("data")
    print("=" * 84)
    print(" REVISION DE LA COLUMNA AA (SEDES-JEFATURA)")
    print("=" * 84)
    if ruta is None or not Path(ruta).exists():
        print("\n[ERROR] No encontre el Excel de Data en la carpeta de ingesta.")
        print("        Copia el archivo del dia y vuelve a intentar.")
        return 2
    print(f" Archivo: {Path(ruta).name}")

    try:
        hojas = ing.leer_y_preparar(Path(ruta), "data")
    except Exception as e:  # noqa: BLE001
        print(f"\n[ERROR] No pude leer el archivo: {e}")
        return 2
    if not hojas:
        print("\n[ERROR] El archivo no aportó ninguna hoja util.")
        return 2
    df = list(hojas.values())[0]

    cod = df["SBAN"].astype(str).str.strip()
    es = cod.apply(app.es_jefatura)
    suf = cod.apply(app.sufijo_jefatura)
    nom_ab = df["Oficina"].apply(app.normalizar_nombre_jefatura)
    suf_nombre = suf.ne("") & ~suf.str.upper().eq("J")
    # Mismo criterio del panel: el sufijo de la columna AA manda si trae el nombre.
    jef = suf.where(suf_nombre, nom_ab).where(es, "")
    mt = (df["Consecutivo mantenimiento 3"].notna()
          & df["Consecutivo mantenimiento 3"].astype(str).str.strip().ne(""))

    n_filas, n_jef = len(df), int(es.sum())
    print(f" Filas del archivo : {n_filas:,}".replace(",", "."))
    print(f" Filas de jefatura : {n_jef:,}".replace(",", "."))
    if n_jef == 0:
        print("\n[REVISAR] La columna AA no trae ninguna fila de jefatura.")
        print("          (Con el archivo del 28-sept no venian: el filtro «Jefatura»")
        print("           del panel aparece vacio. Revisa el archivo de origen.)")
        return 1

    # --- Tabla por sede-jefatura (agrupando por la region de la que depende) ---
    cod_j = suf.copy()          # el sufijo identifica la sede dentro de su regional
    cod_reg = cod.str.extract(r"^\s*(\d{5})", expand=False).fillna("")
    tabla = (pd.DataFrame({"_reg": cod_reg[es], "_jef": jef[es].astype(str),
                           "_mt": mt[es]})
             .groupby(["_reg", "_jef"], as_index=False)
             .agg(Equipos=("_mt", "size"), MT3=("_mt", "sum")))
    tabla["Pendientes"] = tabla["Equipos"] - tabla["MT3"]
    tabla = tabla.sort_values(["_reg", "_jef"])
    print(f"\n Sedes-jefatura: {tabla['_jef'].nunique()} · equipos: {int(tabla['Equipos'].sum())}"
          f" · con MT3: {int(tabla['MT3'].sum())} · pendientes: {int(tabla['Pendientes'].sum())}")
    print(f"\n   {'COD':<7}{'SEDE-JEFATURA':<20}{'EQUIPOS':>8}{'MT3':>6}{'PEND':>7}   OBSERVACION")
    for _, r in tabla.iterrows():
        obs = ""
        if r["_jef"] not in SEDES_JEFATURA:
            sug = parecido(r["_jef"])
            obs = f"⚠ revisar (¿{sug}?)" if sug else "⚠ nombre no reconocido"
        print(f"   {r['_reg']:<7}{r['_jef']:<20}{int(r['Equipos']):>8}{int(r['MT3']):>6}"
              f"{int(r['Pendientes']):>7}   {obs}")

    # --- Avisos ---
    avisos: list = []
    nombres = sorted(x for x in tabla["_jef"].unique().tolist() if x)

    # 1) Nombres que el panel ya corrige al mostrar.
    corregidos = sorted({n for n in nombres
                         if app.normalizar_nombre_jefatura(n) != n})
    for c in corregidos:
        avisos.append(f"'{c}' esta mal escrito en el archivo: el panel lo muestra como "
                      f"'{app.normalizar_nombre_jefatura(c)}'. Conviene corregirlo en el Excel de origen.")

    # 2) Nombres que no son sedes conocidas (posible error nuevo o sede nueva).
    for n in nombres:
        if n not in SEDES_JEFATURA and n not in corregidos:
            sug = parecido(n)
            avisos.append(f"'{n}' no es una sede-jefatura conocida"
                          + (f"; se parece a '{sug}'." if sug else "."))

    # 3) La misma sede escrita de dos formas distintas.
    por_sede: dict = {}
    for n in nombres:
        clave = app.normalizar_nombre_jefatura(n)
        if clave not in SEDES_JEFATURA:
            clave = parecido(clave) or clave
        por_sede.setdefault(clave, []).append(n)
    for sede, variantes in sorted(por_sede.items()):
        if len(variantes) > 1:
            avisos.append(f"La sede '{sede}' viene escrita de {len(variantes)} formas: "
                          + " y ".join(f"'{v}'" for v in variantes)
                          + ". El panel las une al mostrar, pero el Excel de origen tiene el error.")

    print()
    if avisos:
        print(f" [REVISAR] {len(avisos)} aviso(s) en la columna AA:")
        for a in avisos:
            print(f"   · {a}")
        print("\n Los avisos NO detienen la publicacion: el panel corrige lo que puede al mostrar.")
    else:
        print(f" [OK] {len(nombres)} sedes-jefatura, todos los nombres correctos.")
    return 1 if avisos else 0


if __name__ == "__main__":
    sys.exit(main())
