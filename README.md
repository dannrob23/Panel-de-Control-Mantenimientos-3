# Panel de Control Mantenimiento Preventivo 3

Dashboard **Streamlit** del mantenimiento preventivo MT3 (Banco Agrario · gestión COLSOF), con
3 módulos (**Componente 1** · **Componente 2 Impresoras láser** · **UPS**), avance por oficina (SBAN),
novedades por sede y fecha/hora de última actualización.

> ⚠️ **IMPORTANTE (datos personales)**: este repositorio debe ser **PRIVADO**.
> La app puede ser pública sin login, pero los archivos de datos NO deben quedar en un repositorio público.
> En GitHub: *Settings → General → Danger Zone → Change visibility → **Private***.

📘 **¿Eres usuario del panel?** Lee el **[TUTORIAL.md](TUTORIAL.md)** (guía paso a paso, sin tecnicismos).

---

## ✨ Funcionalidades

- **Tema oscuro y claro** con interruptor arriba a la derecha (paletas propias, no solo el tema de Streamlit).
- **Barra de "Facturación" siempre visible** (Todos · 🏦 Facturables BANCO · 🏢 No facturables COLSOF) con los
  conteos de cada grupo — antes vivía escondida en la barra lateral.
- **Chips de estado** en el encabezado: módulo, operatividad, fecha de datos, filtro de facturación activo,
  regla de impresoras, datos desactualizados y modo público.
- **Encabezado (hero)** con la vista actual y los totales base; **tarjetas KPI** con desglose de facturación.
- **Indicadores uniformes (una sola fuente por concepto)**: los equipos se cuentan en un único bloque de
  5 tarjetas (Total · MT con el % de avance · Pendientes · Facturables · No facturables) y las oficinas en
  un bloque de 5 tarjetas con **el mismo universo** (las 806 oficinas del archivo Campos dashboard, una por
  SBAN) y **los mismos filtros** (Oficina · Facturación · clic del ranking · solo pendientes):
  `Componente 1/2/UPS: oficinas 100% MT3`, `🏁 Oficinas finalizadas (Campos N)` y
  `🔋 UPS finalizadas (col AP)`. Las cifras no cambian al cambiar de módulo y el tooltip explica qué queda
  fuera del filtro.
- **Pestañas**: 📈 Gráficos y avance · 🏢 Resumen por oficina · 📋 Gestión de novedades; en el módulo **UPS**
  se agrega **🔋 Avance UPS (col AP)** y siempre está disponible **📰 Novedades de las oficinas**.
- **Avance de las UPS desde la columna AP «ESTADO UPS»** del archivo *Campos dashboard* (Banco Agrario y
  COLSOF integrados en el mismo documento): indicadores de sedes reportadas / finalizadas / en proceso /
  programadas / sin reporte, tabla por sede con el avance MT3 de las UPS y **alertas de control cruzado**
  (por ejemplo, sede marcada `Finalizado` con UPS sin MT3). Descarga CSV/Excel.
- **Novedades de las oficinas** en una pestaña propia (hoja `Novedades Equipos`), con gráfico por categoría,
  filtros de categoría y rango de fechas, y exportación; no se aíslan por componente porque son transversales.
- **Gráficos (Plotly)**: dona, taquímetro de avance, ranking de sedes con más pendientes (clic para filtrar),
  tendencia diaria y mapa de calor Regional × Estado.
- **Resumen por oficina** con filtros propios, fila de totales y **descarga a Excel** (vista + gráficos + novedades).
- **Gestión de novedades**: tabla editable (las anotaciones persisten) y exportación CSV/Excel.
- **Filtros** de oficina (multiselección con búsqueda), "solo pendientes", filtro por clic en el ranking,
  **chips de filtros activos con ✖** y botón **🧹 Limpiar todos los filtros**.
- **Aviso de datos desactualizados** (chip ⚠️ cuando los datos tienen 3 días o más).
- **Diseño responsivo** (móvil/tablet): tarjetas apiladas, gráficos a ancho completo y controles táctiles.
- **Validación de integridad** de los archivos en uso y expander **🗂️ Auditoría y fuentes** (rutas, fechas, bitácora).
- **Modo público** con enmascarado de seriales/placas.

---

## 🔒 Seguridad: vista pública vs. administrador

El panel tiene **dos modos**, controlados por secretos (`.streamlit/secrets.toml` local o
*Settings → Secrets* en Streamlit Cloud). `secrets.toml` **está en `.gitignore`** (nunca se sube).

| Secreto | Efecto | Recomendado |
|---|---|---|
| *(ninguno)* | **Vista pública**: solo consulta. Sin uploader, sin carga posible | ✅ **Streamlit Cloud** |
| `modo_admin = true` | Habilita el módulo **📂 Ingesta diaria** (subir `.xlsx`) | ✅ Solo en el PC del administrador |
| `modo_publico = false` | Muestra seriales y placas **sin enmascarar** | Solo en el PC del administrador |

> 🛡️ **En Streamlit Cloud NO definir `modo_admin`**: así el público nunca puede subir archivos
> (el `st.file_uploader` no se dibuja) y el panel queda 100 % de consulta. Los datos se actualizan
> por `git push` (ver más abajo) y Streamlit Cloud redepliega automáticamente.

**Ejemplo de `.streamlit/secrets.toml` (solo en el PC del administrador):**
```toml
modo_admin = true
# modo_publico = false   # descomentar para ver seriales/placas reales en local
```

---

## 🚀 Ejecutar en local

```bash
# 1) Entorno virtual e instalación de dependencias
python -m venv .venv
.venv\Scripts\activate          # Windows (bash: source .venv/Scripts/activate)
pip install -r requirements.txt

# 2) (Opcional) habilitar la carga de datos en tu PC
#    Crear .streamlit\secrets.toml con:  modo_admin = true

# 3) Arrancar
streamlit run app.py            # http://localhost:8501
                                 # red local: http://<IP-del-PC>:8501
```

---

## ☁️ Despliegue en Streamlit Community Cloud

1. Repositorio **privado** en GitHub (este).
2. En https://share.streamlit.io → **New app** → repositorio, branch `main`, main file `app.py` → **Deploy**.
3. Queda una URL pública tipo `https://<nombre>.streamlit.app` (**sin login**).
4. **Settings → Secrets**: dejar **vacío** de `modo_admin` (ver sección de seguridad).
5. Tras cambios en el repo, la app se redespliega sola; si no, usa **Reboot app** en el menú ⋮.

---

## 📥 Actualización diaria de datos

El flujo **que publica** en la web es el publicador (no el uploader del panel):

1. Copiar los `.xlsx` del día (Data, Cronograma, Campos) en la carpeta hermana
   **`..\Ingesta de datos diaria`**. El **Campos dashboard** debe traer la hoja `Dashboard_KPI`
   con la columna **AP «ESTADO UPS»** y la hoja `Novedades Equipos`.
2. Doble clic en **`ingesta_publicar.bat`** (dentro de `deploy_panel/`) → sincroniza `app.py` y la
   documentación desde la raíz, **anonimiza** los datos a `data/`, escribe la **fecha/hora** en
   `ultima_actualizacion.json` y hace `git commit/push`.
3. Streamlit Cloud redespliega (1-2 min) y el panel muestra el nuevo `🕒 Última actualización`.

> ⚠️ El **uploader** de la barra lateral (modo admin) sirve para *ver* datos en local, **no publica**.
> Además, en Streamlit Cloud el disco es efímero: lo subido por esa vía se pierde en el próximo reinicio.

> 🔎 Si el `push` falla por credenciales de GitHub (`SEC_E_NO_CREDENTIALS`), abre Git Bash/CMD,
> ejecuta `git -C deploy_panel push origin main` e introduce tu usuario y token una vez; los commits
> ya quedan hechos y solo falta subirlos.

---

## 🔄 Actualizar el proyecto en otro equipo

Doble clic en **`actualizar_panel.bat`** (equivale a `git pull origin main`). Si aparece un conflicto
(habitualmente en `data/*.xlsx` por publicaciones simultáneas), **no borrar nada**: resolver manualmente
o pedir ayuda.

---

## 🧩 Estructura del repositorio

```
app.py                     Aplicación principal (Streamlit)
TUTORIAL.md                Guía de uso para usuarios y administradores
ingesta_publicar.py/.bat   Publicador diario (copia a data/ + commit/push)
actualizar_panel.bat       Trae los últimos cambios del repositorio (git pull)
sanitizar_publico.py       Utilidad de saneado de datos para vistas públicas
data/                      Excel vigentes + ultima_actualizacion.json
.streamlit/config.toml     Tema base y configuración del servidor
.streamlit/secrets.toml    🔒 Local (ignorado por git): modo_admin / modo_publico
uploads/                   (ignorado por git) archivos subidos por el uploader + bitácora
requirements.txt           streamlit · pandas · openpyxl · plotly
```

`deploy_panel/` es el repositorio que se publica en Streamlit Cloud (tiene su propio git y su
`data/`). Ahí vive **`sincronizar_app.py`**, que copia `app.py`, `README.md` y `TUTORIAL.md` desde la
raíz del proyecto antes de publicar, para que la nube reciba exactamente la versión probada en local.

---

## ⚙️ Reglas de negocio

- **Impresoras láser (Componente 2) = facturables (Si).** El Excel de origen las trae como "No", por lo que
  el panel aplica la regla con la constante `IMPRESORAS_FACTURABLES` (en `app.py`) y lo anuncia con el chip
  `ℹ️ Impresoras: facturables (Si)`. **Pendiente:** corregirlo también en el archivo de origen.
- **Avance de UPS = columna AP «ESTADO UPS» de Campos dashboard.** Es la fuente autoritativa: el panel la
  muestra tal cual (normalizando `Finalizado → Finalizada`) y **no la deriva** del cronograma. Solo si el
  archivo no trae esa columna se usa el respaldo del cronograma UPS (`FECHA` + avance MT3, `estados_ups()`).
- **Novedades de oficina = hoja `Novedades Equipos`** (categoría, fecha, serial, estado y observación).
  El cruce con la Data es por **Serial** y, si no aparece, por texto de la observación; el resumen por sede
  se amarra por **SBAN** (respaldo `SBAN PCT`).
- **Subsanado** = tiene consecutivo MT3 · **Pendiente** = no lo tiene.
- **Un solo denominador para las oficinas (806).** Las tarjetas de oficinas usan el universo de
  **Campos dashboard** (una fila por SBAN); las 3 oficinas que están en la Data y no en Campos (bodegas)
  quedan fuera y el panel lo advierte en la nota al pie. Antes cada tarjeta usaba un universo distinto
  (808 de la Data / 824 filas de Campos / 805 de UPS), lo que hacía ver cifras incoherentes.
- **% de avance:** un único indicador, como delta de la tarjeta de MT realizados (antes se repetía en una
  barra de progreso aparte).
- Al cambiar de módulo se reinician los filtros que podrían dejar la vista vacía.

---

## 🎨 Motor de temas (para desarrollo)

- `TEMAS` (diccionario) define las paletas `oscuro` y `claro`: fondos, tarjetas, bordes, textos y colores
  semánticos (`verde` avance, `pend` pendientes, `pista` fondo del taquímetro).
- `aplicar_css()` inyecta un CSS con marcadores (`__INK__`, `__CARD__`, …) sustituidos según el tema activo.
- `plotly_base()` construye los gráficos con la paleta del tema; `grafico_dona/gauge/tendencia/heatmap/top_pendientes`.
- El toggle se lee de `st.session_state` **antes** de aplicar el CSS (evita el efecto "hay que pulsar dos veces").
- Detalles y trampas conocidas: ver skill `streamlit-tema-claro-oscuro`.

---

## 📚 Documentación

- **[TUTORIAL.md](TUTORIAL.md)** — guía de uso (usuarios y administradores).
- `README.md` — este documento (operación, despliegue y mantenimiento).
