# 📘 Tutorial del Panel MT3 — Control Mantenimiento Preventivo 3

Guía práctica para usar el panel. No necesitas conocimientos técnicos.

---

## 1. ¿Qué es y para qué sirve?

Es un **tablero de control** del mantenimiento preventivo (MT3) del Banco Agrario (gestión COLSOF).
Muestra, en tiempo real y de forma visual:

- El **avance** de los mantenimientos por módulo (Componente 1, Componente 2 — impresoras láser — y UPS)
- El **detalle por oficina (SBAN)**: cuántos equipos hay, cuántos están subsanados y cuántos pendientes
- Las **novedades** por sede y fecha
- Permite **descargar** la información filtrada a Excel

---

## 2. Cómo entrar

| Desde dónde | Dirección | Qué puedes hacer |
|---|---|---|
| **Navegador (URL pública)** | `https://<tu-app>.streamlit.app` | **Solo consulta** (nadie puede subir archivos) |
| **Tu PC** | `http://localhost:8501` | Consulta + carga de datos (si eres administrador) |
| **Celular (misma red WiFi que el PC)** | `http://192.168.1.6:8501` | Consulta + carga (si el PC está encendido) |

> 🔒 En la **URL pública** el panel es **100 % de consulta**: la opción de subir archivos está bloqueada
> por diseño. La actualización de datos la hace el administrador desde su PC.

---

## 3. Las partes de la pantalla

```
┌──────────────────────────────────────────────────────────────────────┐
│                                        [🌙 Oscuro] [☀️ Claro]  ← tema │
│  Módulo:  [Componente 1] [Componente 2] [UPS]                        │
│                                                                       │
│  🛠️ Control Mantenimiento Preventivo 3                                │
│  Vista: Todas las oficinas · Atribución: Todos · 29.860 equipos       │
│  [📦 Componente 1] [Operativo] [🕒 2026-09-10 12:07] [🏦 Viendo: …]  │  ← chips de estado
│                                                                       │
│  Facturación  [Todos] [🏦 Facturables (BANCO)] [🏢 No facturables]    │  ← filtro visible
│  📊 Del módulo (29.860 equipos): 🏦 22.230 · 🏢 7.630                 │
│                                                                       │
│  [📦 Total] [✅ MT] [⚠️ Pendientes] [🏦 Facturables] [🏢 No fact.]   │  ← tarjetas KPI
│                                                                       │
│  📈 Gráficos y avance │ 🏢 Resumen por oficina │ 📋 Gestión novedades │  ← pestañas
│  (con el módulo UPS se agregan: 🔋 Avance UPS (col AP) · 📰 Novedades de las oficinas)
└──────────────────────────────────────────────────────────────────────┘
   ☰ Barra lateral: Datos · Filtros · Auditoría y fuentes
```

### Los "chips" del encabezado (siempre te dicen el estado)

| Chip | Significa |
|---|---|
| `📦 Componente 1/2/UPS` | Módulo que estás viendo |
| `Operativo` | El panel está funcionando correctamente |
| `🕒 fecha/hora` | Última actualización de los datos |
| `🏦 Viendo: Solo facturables` | Tienes activo el filtro de facturación |
| `ℹ️ Impresoras: facturables (Si)` | Regla del negocio: las impresoras láser se facturan al Banco |
| `⚠️ Datos de hace N días` | Los datos tienen 3 días o más (conviene actualizar) |
| `🔒 Modo público` | Los seriales y placas se muestran enmascarados |

---

## 4. Los 3 módulos

| Módulo | Qué incluye | Equipos |
|---|---|---|
| **Componente 1** | Equipos del componente 1 | 29.860 |
| **Componente 2 (Impresoras láser)** | Impresoras láser | 1.064 |
| **UPS** | Equipos UPS | 867 |

- Se cambia con el selector **Módulo** (arriba).
- Al cambiar de módulo, **los filtros se reinician solos** (para no dejarte una vista vacía sin explicación).
- **Regla de negocio:** las impresoras láser (Componente 2) se **facturan al Banco**, aunque el archivo de
  origen venga marcado como "No". El panel lo aplica y te lo avisa con el chip `ℹ️ Impresoras: facturables (Si)`.
- **Avance de las UPS:** en el módulo **UPS** el estado de cada sede se toma **tal cual** de la columna
  **AP «ESTADO UPS»** del archivo **Campos dashboard** (`Finalizado`, `En proceso`, `Programada`…).
  El panel **no lo calcula ni lo estima**: solo lo muestra, junto al conteo de UPS y su avance MT3
  (que sirve de control cruzado).

---

## 5. Las tarjetas KPI (indicadores)

| Indicador | Qué mide |
|---|---|
| 📦 **Total de elementos** | Equipos que cumplen los filtros actuales |
| ✅ **Mantenimientos realizados (MT)** | Equipos con consecutivo MT3 asignado (subsanados) |
| ⚠️ **Mantenimientos pendientes** | Equipos sin consecutivo MT3 (y su % sobre el total) |
| 🏦 **Facturables (Si)** | Equipos atribuibles al Banco |
| 🏢 **No facturables (No)** | Equipos de gestión COLSOF |

Debajo verás el **desglose**: `BANCO (Si): 22.230 · COLSOF (No): 7.630 · Total: 29.860`
y la barra de **Avance MT3 del filtro actual**.

> **Subsanado** = ya tiene mantenimiento hecho (consecutivo MT3). **Pendiente** = aún no lo tiene.

---

## 6. Filtros (cómo dejar la vista en lo que te interesa)

| Filtro | Dónde está | Para qué sirve |
|---|---|---|
| **Módulo** | Arriba | Ver Componente 1, impresoras o UPS |
| **Facturación** | Barra visible bajo el encabezado | Ver solo facturables (BANCO) o solo no facturables (COLSOF) |
| **Oficina (SBAN - Nombre)** | Barra lateral ☰ | Elegir una o varias oficinas (escribe para buscar) |
| **Mostrar solo pendientes ⚠️** | Barra lateral ☰ | Ocultar lo que ya está subsanado |
| **Clic en el ranking** | Pestaña Gráficos | Haz clic en una barra del ranking para filtrar esa sede |
| **Chips de "Filtros activos"** | Bajo las tarjetas KPI | Quita un filtro puntual con su ✖ |
| **🧹 Limpiar todos los filtros** | Barra lateral ☰ | Vuelve todo a cero de una vez |

💡 **Consejo:** si los números no coinciden con lo que esperas, mira primero los **chips de filtros activos**.

---

## 7. Las pestañas

### 📈 Gráficos y avance
- **Dona**: subsanados vs. pendientes
- **Taquímetro**: % de avance global
- **Ranking de sedes**: las que tienen más pendientes (puedes mostrar 5, 10, 12, 15, 20 o 25)
- **Tendencia MT3**: mantenimientos realizados por día
- **Mapa de calor**: Regional × Estado

### 🏢 Resumen por oficina
Tabla con una fila por sede: `SBAN · Oficina · Estado · Total · Subsanados · Pendientes · % Avance ·
Categoría novedad · Nov. del día · Nov. acumuladas`, más una fila de **totales**.
Tiene sus **propios filtros** (estado del cronograma, avance mínimo, categoría de novedad, solo novedades del día)
y el botón **⬇️ Descargar Excel** (vista filtrada + gráficos + novedades).

### 🔋 Avance UPS (col AP) — solo en el módulo UPS
Panel del avance de las UPS tomado de la **columna AP «ESTADO UPS»** del archivo Campos dashboard:

| Indicador | Qué mide |
|---|---|
| 🔋 **Sedes con UPS (col AP)** | Sedes que traen algún valor en esa columna |
| ✅ **UPS finalizadas (col AP)** | Sedes marcadas `Finalizado` (y su % sobre el total) |
| 🟨 **En proceso · ⬜ Programadas** | Conteo de los demás estados de la columna AP |
| ⚪ **Sin reporte en col AP** | Sedes sin dato + número de **alertas** |

La tabla muestra `SBAN · Oficina · UPS en Data · UPS con MT3 · Avance MT3 % · Estado UPS (col AP) ·
Observación UPS · Alerta`. Las **alertas** cruzan la columna AP con la Data (por ejemplo, una sede que dice
`Finalizado` pero todavía tiene UPS sin MT3). Puedes **descargar** el panel en CSV o Excel.

### 📰 Novedades de las oficinas — en todos los módulos
Novedades reportadas por las oficinas (hoja **Novedades Equipos** del archivo Campos dashboard):
`SBAN · Oficina · Fecha · Departamento · Aliado · Categoría · Estado del equipo · Facturable · Serial · Observación`.
Trae un gráfico de **novedades por categoría**, filtros de **categoría** y **rango de fechas**, y descarga en CSV/Excel.
A diferencia de la tabla de gestión, aquí las novedades **no se separan por componente** (una novedad de oficina
es transversal: denuncio, recolección, cambio a facturable…).

### 📋 Gestión de novedades
Tabla **editable**: escribe directamente en la columna **Observaciones / Novedades**.
Tus anotaciones se conservan al cambiar de filtro y puedes **exportarlas a CSV/Excel**.

---

## 8. Tema claro/oscuro y uso en el celular

- **Tema:** botones `🌙 Oscuro` / `☀️ Claro` arriba a la derecha. Tu elección se mantiene mientras navegas.
- **Celular:** el panel se adapta (tarjetas en 2 columnas, gráficos apilados, botones táctiles).
  Las tablas se deslizan horizontalmente con el dedo.

---

## 9. Avisos que puedes ver (y qué significan)

| Aviso | Qué hacer |
|---|---|
| `🔒 Modo público (datos enmascarados)` | Es normal en la vista pública: seriales/placas salen como `***1234` |
| `⚠️ Datos de hace N días` | Avisar al administrador para que publique la data nueva |
| `🔍 Validación de integridad — N aviso(s)` | Revisa qué archivo está desactualizado o inconsistente (se abre el detalle) |
| `🔒 Datos — Panel de consulta` | Estás en la vista pública: la carga está restringida al administrador |

---

## 10. 🛠️ Para el administrador: cargar y publicar datos

Hay **dos cosas distintas** que conviene no confundir:

| | **Uploader** (barra lateral) | **Publicador** `ingesta_publicar.bat` |
|---|---|---|
| Qué hace | Sube un `.xlsx` y el panel **local** muestra esos datos | Copia los Excel a `data/` y hace **commit + push** |
| ¿Actualiza la web pública? | ❌ **NO** | ✅ **SÍ** (la nube se actualiza sola en 1-2 min) |
| De dónde toma los archivos | Lo que subes | Carpeta hermana `Documents\Ingesta de datos diaria\` |

### Para ver el uploader en tu PC
Crea el archivo `.streamlit\secrets.toml` dentro del proyecto con:
```toml
modo_admin = true
```
Sin ese archivo, el panel se ve como la vista pública (solo consulta).

### Flujo correcto de publicación (paso a paso)
1. Guarda los Excel del día (Data, Cronograma, Campos) en la carpeta **`Ingesta de datos diaria`**
   (hermana del proyecto). El archivo **Campos dashboard** debe traer la hoja `Dashboard_KPI` con la
   columna **AP «ESTADO UPS»** y la hoja `Novedades Equipos` (novedades de las oficinas).
2. Doble clic en **`ingesta_publicar.bat`** → copia, actualiza la fecha/hora y sube a GitHub.
3. Espera 1-2 minutos y recarga la página pública con **Ctrl+F5**: la fecha del chip `🕒` debe cambiar.

### Actualizar el proyecto en otro equipo
Doble clic en **`actualizar_panel.bat`** (trae los últimos cambios del repositorio).

---

## 11. Problemas comunes

| Síntoma | Causa probable | Solución |
|---|---|---|
| No veo la opción de subir archivos | Estás en modo consulta (o es la URL pública) | Es lo correcto. Para cargar, hazlo en el PC del administrador |
| Subí un archivo y la web no cambia | El uploader **no publica** | Ejecuta `ingesta_publicar.bat` |
| La fecha de actualización no avanza | No se ejecutó el publicador o el push falló | Ejecuta el `.bat` y revisa que diga "push" sin errores |
| Veo seriales como `***1234` | Modo público (enmascarado) | Es normal en la vista pública |
| Los totales no cuadran | Hay filtros activos | Mira los chips de "Filtros activos" o pulsa 🧹 Limpiar filtros |
| "Datos de hace N días" | La data no se ha actualizado | Publicar la data nueva con el `.bat` |

---

## 12. Glosario

| Término | Significado |
|---|---|
| **SBAN** | Código de la oficina (5 dígitos) |
| **MT3 / MT** | Mantenimiento preventivo 3 (el consecutivo asignado cuando se realiza) |
| **Subsanado** | Equipo que ya tiene su mantenimiento hecho |
| **Pendiente** | Equipo que aún no tiene mantenimiento |
| **Cronograma** | Programación de las visitas/actividades por sede |
| **Novedad** | Situación particular reportada de un equipo o sede |
| **Facturable** | Equipo cuyo mantenimiento se factura al Banco (Si) o es gestión COLSOF (No) |
| **UPS** | Sistema de alimentación ininterrumpida (baterías/respaldo) |
| **Modo público** | Vista de consulta con datos personales enmascarados |
| **Modo administrador** | Vista con permiso para cargar datos (solo en el PC del administrador) |

---

*Panel MT3 · Banco Agrario · COLSOF — documento de uso interno.*
