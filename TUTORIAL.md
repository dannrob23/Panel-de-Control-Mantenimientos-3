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
│  [🏢 Oficinas intervenidas] [📈 Avance MT3]                           │  ← destacadas (titular)
│  [📦 Total] [✅ MT] [⚠️ Pendientes] [🏦 Facturables] [🏢 No fact.]   │  ← KPI de equipos
│  [C1/C2/UPS completo] [🏁 Reportadas finalizadas] [🔋 UPS reportado]  │  ← KPI de oficinas
│                                                                       │
│  📈 Gráficos y avance │ 🏢 Resumen por oficina │ 📋 Gestión novedades │  ← pestañas
│  (con el módulo UPS se agregan: 🔋 UPS: avance y diferencias · 📰 Novedades de las oficinas)
└──────────────────────────────────────────────────────────────────────┘
   ☰ Barra lateral: Datos · Filtros · Auditoría y fuentes
```

### Ruta rápida: cómo leer el panel en 4 pasos

1. **Elige el módulo** arriba: `Componente 1`, `Componente 2 (impresoras láser)` o `UPS`.
2. **Mira las 2 tarjetas destacadas**: *Oficinas intervenidas* (¿cuántas oficinas empezamos?) y
   *Avance MT3* (¿qué % del módulo está hecho?).
3. **Revisa los KPI**: el bloque de **equipos** (total / realizados / pendientes) y el de **oficinas**
   (cuántas tienen el componente **completo** y cómo va lo **reportado por la oficina**).
4. **Baja a las pestañas** para entender el *por qué*: en **📈 Gráficos y avance** ves el ranking
   (¿quién acumula pendientes?), la **tendencia** (¿a qué ritmo vamos?) y el **mapa de calor**
   (¿dónde se concentra el atraso?).

> 🗣️ **Dos fuentes, dos lenguajes.** El panel distingue siempre entre:
> - **Lo que revisa el panel** (con la Data de ejecución): «*Oficinas con… completo*», «*Oficinas
>   intervenidas*», «*Avance MT3*».
> - **Lo que reporta la oficina** (en el archivo Campos dashboard): «*Oficinas reportadas como
>   finalizadas*», «*UPS finalizadas (reportado)*».
>
> Si las dos cifras no coinciden, **no es un error**: son fuentes distintas y el panel te muestra
> las diferencias en la pestaña **🔋 UPS: avance y diferencias**.

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

Todas las tarjetas del panel están pensadas para que **cada dato se vea una sola vez**. Si un número
no coincide con lo que esperabas, revisa primero el filtro 🏦/🏢 y los chips de filtros activos.

### Tarjetas destacadas (las 2 grandes de arriba)

Son el "titular" del panel: responden *¿cuántas oficinas ya trabajamos?* y *¿cuánto llevamos?*

| Indicador | Qué mide | Cómo se calcula |
|---|---|---|
| 🏢 **Oficinas intervenidas** | Oficinas del **módulo activo** donde ya se empezó (al menos **1** MT3) | SBAN distintos con `≥1 MT3` ÷ SBAN distintos del módulo. **No** es el estado oficial de la sede |
| 📈 **Avance MT3** | % de equipos ya subsanados en el filtro activo (con su barra) | `realizados ÷ total` |

> ⚠️ **No confundir** *intervenida* con *completa*:
> **intervenida** = ya empezaron (aunque sea 1 equipo) · **completa** = terminaron **todos** los equipos.
> Ejemplo real (Componente 1): **381 de 808** oficinas intervenidas, pero solo **265 de 806** están completas.
> Por eso la primera cifra siempre es **mayor** que la segunda.

### Equipos del módulo (5 tarjetas)

| Indicador | Qué mide |
|---|---|
| 📦 **Total de elementos** | Equipos que cumplen los filtros actuales |
| ✅ **Mantenimientos realizados (MT)** | Equipos con consecutivo MT3 asignado (subsanados). El **% de avance** está en la tarjeta destacada 📈 (no se repite aquí) |
| ⚠️ **Mantenimientos pendientes** | Equipos sin consecutivo MT3 (y su % sobre el total) |
| 🏦 **Facturables (Si)** | Equipos atribuibles al Banco |
| 🏢 **No facturables (No)** | Equipos de gestión COLSOF |

### Oficinas (5 tarjetas)

Todas usan **el mismo universo** (las oficinas del archivo Campos dashboard, una por SBAN) y
los **mismos filtros** (Oficina · Facturación · clic del ranking · solo pendientes). Por eso las cifras
**no cambian al cambiar de módulo** y son comparables entre sí.

Los nombres están escritos **sin jerga interna**: dicen *quién* afirma el dato.

| Indicador | Quién lo dice | Qué mide |
|---|---|---|
| 🏢 **Oficinas con C1 / C2 / UPS completo** | **El panel** (revisa la Data de ejecución) | Oficinas donde **todos** los equipos de ese componente ya tienen mantenimiento hecho. Elijo el número como «completas de total» y el delta dice cuántas oficinas del archivo faltan |
| 🏁 **Oficinas reportadas como finalizadas** | **La oficina** (archivo Campos dashboard) | Oficinas que la propia oficina marcó como `Finalizada` |
| 🔋 **UPS finalizadas (reportado)** | **La oficina** (archivo Campos dashboard) | Oficinas que reportaron sus UPS como `Finalizado`, **sobre el total de oficinas** |

> **Completo** = el panel comprobó que ya se le hizo mantenimiento a todos los equipos de esa oficina
> (consecutivo MT3). **Reportado** = lo que la oficina escribió en el archivo; puede no coincidir con
> la Data, y esa diferencia es justamente lo que el panel vigila.
>
> Los conteos de equipos (Total, Realizados, Pendientes) viven **solo** en el bloque de Equipos; en las
> tablas se repiten únicamente **por fila** (por oficina), no como total general.

---

## 6. Filtros (cómo dejar la vista en lo que te interesa)

| Filtro | Dónde está | Para qué sirve |
|---|---|---|
| **Módulo** | Arriba | Ver Componente 1, impresoras o UPS |
| **Facturación** | Barra visible bajo el encabezado | Ver solo facturables (BANCO) o solo no facturables (COLSOF). **Afecta TODAS las tarjetas**, incluidas las de oficinas |
| **Oficina (SBAN - Nombre)** | Barra lateral ☰ | Elegir una o varias oficinas (escribe para buscar) |
| **Mostrar solo pendientes ⚠️** | Barra lateral ☰ | Ocultar lo que ya está subsanado |
| **Clic en el ranking** | Pestaña Gráficos | Haz clic en una barra del ranking para filtrar esa sede |
| **Chips de "Filtros activos"** | Bajo las tarjetas KPI | Quita un filtro puntual con su ✖ |
| **🧹 Limpiar todos los filtros** | Barra lateral ☰ | Vuelve todo a cero de una vez |

💡 **Consejo:** si los números no coinciden con lo que esperas, mira primero los **chips de filtros activos**.

---

## 7. Las pestañas

### 📈 Gráficos y avance

Esta pestaña responde a tres preguntas: **¿cuánto llevamos?**, **¿quién va atrasado?** y
**¿dónde se concentra el pendiente?** Todos los gráficos reaccionan a los filtros activos.

#### 🍩 Dona — "Avance de Mantenimiento Preventivo 3"
- **Qué muestra:** dos porciones sobre el total de equipos del filtro: verde = **subsanados** (con MT3)
  y naranja = **pendientes**.
- **El número del centro** es el **% de avance global** (`realizados ÷ total`).
- **Para qué sirve:** es la foto de un vistazo del avance. Si la porción verde crece, el plan avanza.
- **Cómo leerla:** pasa el mouse por cada porción para ver el conteo exacto y el porcentaje.

#### 🏆 Ranking de sedes (pendientes)
- **Qué muestra:** barras horizontales con las sedes que tienen **más pendientes**. Puedes elegir
  mostrar 5, 10, 12, 15, 20 o 25.
- Cada barra combina **subsanados + pendientes**; la etiqueta es `SBAN · nombre de la oficina`.
- **Para qué sirve:** **priorizar**. Las barras de arriba son las que más trabajo acumulan.
- **Truco:** **haz clic en una barra** y todo el panel se filtra por esa sede (aparece un chip arriba).
  Quítalo con **✖ Quitar filtro** o con 🧹 Limpiar filtros.

#### 📈 Tendencia MT3 (realizados por día)
- **Qué muestra:** dos series sobre la misma línea de tiempo:
  - **Área celeste con puntos** = MT3 realizados **cada día** (según `Fecha de mantenimiento 3`).
  - **Línea punteada verde** = **acumulado** (suma corrida, eje derecho).
- **Para qué sirve:** ver el **ritmo**. Si el área diaria baja o el acumulado se aplana, la operación
  se está frenando; si el acumulado sube en línea recta, el ritmo es constante.
- **Cómo leerla:** eje izquierdo = MT3 del día · eje derecho = acumulado. El subtítulo indica el total
  y el último día con registros.

#### 🗺️ Mapa de calor (Regional × Estado de la sede)
- **Qué muestra:** una cuadrícula:
  - **Filas** = **Regional** (más una fila **SIN REGIONAL**, que agrupa bodegas/stock sin regional asignada).
  - **Columnas** = **estado de la sede**: `Programada · En proceso · Finalizada · Reprogramada_Finalizada · Sin cronograma`.
  - **Color y número de cada celda** = **equipos pendientes** de esa combinación. Cuanto más claro/amarillo,
    más pendientes (el rango de color va de **0** al máximo real).
- **Al pasar el mouse** por una celda ves además el **Total** de equipos y el **% de avance MT3**.
- **Para qué sirve:** localizar **dónde** está el atraso, no solo cuánto.
  - Una mancha fuerte en **Programada** de una regional = mucho trabajo aún por ejecutar.
  - Una mancha en **Finalizada** = equipos pendientes en sedes que la oficina ya marcó como cerradas:
    **revisar**, porque hay incoherencia entre el estado oficial y la Data.
  - Una mancha en **SIN REGIONAL** = stock/bodegas sin regional; no es una sede operativa.
- **Cómo validarlo:** el subtítulo del gráfico dice cuántos equipos cubre. Ese número debe **coincidir**
  con el "📦 Total de elementos" de las tarjetas KPI para el mismo filtro.

#### 📊 Novedades por categoría (dentro de 📰 Novedades de las oficinas)
- **Qué muestra:** cuántas novedades hay por **categoría** (DENUNCIO, RECOLECCIÓN, CAMBIO DE ESTADO A
  FACTURABLE, OPERACIÓN, ASEGURAMIENTO…).
- **Para qué sirve:** ver **qué tipo de problema domina** y a quién escalarlo (seguridad, logística, facturación).


### 🏢 Resumen por oficina
Tabla con una fila por sede: `SBAN · Oficina · Estado · Total · Subsanados · Pendientes · % Avance ·
Categoría novedad · Nov. del día · Nov. acumuladas`, más una fila de **totales**.
Tiene sus **propios filtros** (estado del cronograma, avance mínimo, categoría de novedad, solo novedades del día)
y el botón **⬇️ Descargar Excel** (vista filtrada + gráficos + novedades).
Los **totales de equipos** no se repiten como tarjetas aquí: están en las tarjetas KPI de arriba
(esta vista solo los muestra **por fila**).

### 🔋 UPS: avance y diferencias — solo en el módulo UPS
Pestaña para comparar **lo que reportó la oficina** (columna «ESTADO UPS» del archivo Campos) con
**lo que el panel revisa en la Data**. Solo lista las oficinas **con UPS en la Data**:

| Indicador | Qué mide |
|---|---|
| 🔋 **Oficinas que ya reportaron** | Oficinas con UPS que ya traen algún dato en el archivo |
| 🧩 **Coinciden el reporte y la Data** | Oficinas donde el reporte dice `Finalizado` **y** además todos sus UPS tienen mantenimiento hecho |
| 🟨 **En proceso · ⬜ Programadas** | Conteo de los demás casos reportados |
| ⚠️ **Diferencias reporte vs Data** | Oficinas donde lo reportado y lo que muestra la Data no cuadran |

> Los dos conteos generales (**🔋 UPS finalizadas (reportado)** y **🏢 Oficinas con UPS completo**)
> están en las tarjetas de indicadores de oficinas, arriba: **no se repiten** aquí. Esta pestaña
> existe para ver **dónde no coinciden**.

La tabla muestra `SBAN · Oficina · UPS en Data · UPS con MT3 · Avance MT3 % · Estado reportado ·
Observación UPS · Alerta`, y el desplegable **❓ Conciliación** explica la diferencia entre lo reportado
por la oficina y lo que el panel calcula con la Data —que puede no coincidir, justamente porque vienen de
fuentes distintas—. Puedes **descargar** el panel en CSV o Excel.

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

| | **Uploader** (barra lateral) | **Ingesta** `ingesta.bat` |
|---|---|---|
| Qué hace | Sube un `.xlsx` y el panel **local** muestra esos datos | Ajusta columnas, valida, copia a `data/` y hace **commit + push** |
| ¿Actualiza la web pública? | ❌ **NO** | ✅ **SÍ** (la nube se actualiza sola en 1-2 min) |
| De dónde toma los archivos | Lo que subes | Carpeta hermana `Documents\Ingesta de datos diaria\` |

### Para ver el uploader en tu PC
Crea el archivo `.streamlit\secrets.toml` dentro del proyecto con:
```toml
modo_admin = true
```
Sin ese archivo, el panel se ve como la vista pública (solo consulta).

### 🤖 Lo más fácil: que se actualice solo

1. Doble clic en **`actualizacion_automatica.bat`** y **deja la ventana abierta**.
2. Guarda los Excel del día en la carpeta **`Ingesta de datos diaria`** (como siempre).
3. Ya está: el panel detecta los archivos, espera a que terminen de copiarse, actualiza los datos
   y **publica solo en GitHub**. En 1-2 minutos la web está al día (recarga con **Ctrl+F5**).

La ventana muestra lo que va pasando (`📥 Detecté un cambio`, `✅ Listo…`) y queda registrado en
`tools\ingesta_auto.log`. Si la cierras, deja de vigilar; si la vuelves a abrir, sigue.

> 📍 **¿Dónde está el `.bat`?** Lo tienes en **dos sitios**: en la carpeta del proyecto y **dentro de
> `Ingesta de datos diaria`** (junto a los Excel). Sirve el que te quede más cómodo: los dos hacen
> lo mismo, porque el script ubica el proyecto solo. Si editas un script, edítalo en la carpeta del
> proyecto y avísame para volver a copiarlo.

### Flujo manual (si prefieres controlarlo tú)
1. Guarda los Excel del día (Data, Cronograma, Campos) en la carpeta **`Ingesta de datos diaria`**
   (hermana del proyecto). El archivo **Campos dashboard** debe traer la hoja `Dashboard_KPI` con la
   columna **AP «ESTADO UPS»** y la hoja `Novedades Equipos` (novedades de las oficinas).
2. Doble clic en **`ingesta.bat`** (en la carpeta del proyecto) y elige la opción **2**.
   El script toma el archivo más reciente de cada tipo, ignora los temporales `~$` de Excel,
   ajusta las columnas, valida contra las reglas del panel, escribe `data\` y `deploy_panel\data\`,
   y finalmente hace el **commit + push**.
3. Espera 1-2 minutos y recarga la página pública con **Ctrl+F5**: la fecha del chip `🕒` debe cambiar.

### Las 4 opciones del menú

| Opción | Cuándo usarla |
|---|---|
| **1 · Revisar y generar** | Para ver el resumen y validar antes de publicar (no sube nada) |
| **2 · Generar y PUBLICAR** | El uso normal del día |
| **3 · Solo analizar columnas** | Cuando quieres revisar qué trae un Excel nuevo (`tools\analisis_ingesta.md`) |
| **4 · Solo copia local** | Para probar en tu PC sin tocar la versión publicada |

> 🧩 **Si necesitas analizar o ajustar una columna**, no hay que programar: abre `ingesta.py`,
> busca la sección `1) CONFIGURACION` y edita `ANALISIS_COLUMNAS`
> (por ejemplo `"Oficina": {"tipo": "texto"}` o `"reemplazar": {"SI": "Si"}`).
> El detalle está en el README, sección *Actualización diaria de datos*.

### Actualizar el proyecto en otro equipo
Doble clic en **`actualizar_panel.bat`** (trae los últimos cambios del repositorio).

---

## 11. Problemas comunes

| Síntoma | Causa probable | Solución |
|---|---|---|
| No veo la opción de subir archivos | Estás en modo consulta (o es la URL pública) | Es lo correcto. Para cargar, hazlo en el PC del administrador |
| Subí un archivo y la web no cambia | El uploader **no publica** | Ejecuta `ingesta.bat` → opción 2 |
| La fecha de actualización no avanza | No se ejecutó la ingesta o el push falló | Ejecuta `ingesta.bat` → opción 2 y revisa que diga "Push realizado" |
| Veo seriales como `***1234` | Modo público (enmascarado) | Es normal en la vista pública |
| Los totales no cuadran | Hay filtros activos | Mira los chips de "Filtros activos" o pulsa 🧹 Limpiar filtros |
| "Datos de hace N días" | La data no se ha actualizado | Publicar la data nueva con el `.bat` |

---

## 12. Glosario

| Término | Significado |
|---|---|
| **SBAN** | Código de la oficina (5 dígitos) |
| **MT3 / MT** | Mantenimiento preventivo 3 (el consecutivo asignado cuando se realiza) |
| **Realizado** | Equipo que ya tiene su mantenimiento hecho |
| **Pendiente** | Equipo que aún no tiene mantenimiento |
| **Oficina intervenida** | Oficina donde ya se realizó **al menos 1** mantenimiento (avance **parcial**) |
| **Oficina con … completo** | Oficina donde **todos** los equipos de ese componente ya tienen mantenimiento (avance **completo**). Lo **revisa el panel** con la Data |
| **Reportado** | Lo que la **oficina** escribió en el archivo Campos dashboard (estado de la sede, estado de las UPS). Puede no coincidir con la Data |
| **Campos dashboard** | Archivo que envía la oficina con el estado de cada sede y sus novedades |
| **Data de ejecución** | Archivo con los equipos y los consecutivos de mantenimiento ya realizados |
| **Cronograma** | Programación de las visitas/actividades por sede |
| **Novedad** | Situación particular reportada de un equipo o sede |
| **Facturable** | Equipo cuyo mantenimiento se factura al Banco (Si) o es gestión COLSOF (No) |
| **UPS** | Sistema de alimentación ininterrumpida (baterías/respaldo) |
| **Modo público** | Vista de consulta con datos personales enmascarados |
| **Modo administrador** | Vista con permiso para cargar datos (solo en el PC del administrador) |

---

*Panel MT3 · Banco Agrario · COLSOF — documento de uso interno.*
