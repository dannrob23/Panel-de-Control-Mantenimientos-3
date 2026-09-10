# Panel de Control Mantenimiento Preventivo 3

Dashboard Streamlit con 3 módulos (Componente 1 · Componente 2 Impresoras láser · UPS),
novedades por sede y fecha/hora de última actualización.

> **IMPORTANTE (datos personales)**: este repositorio debe ser **PRIVADO**.
> La app puede ser pública sin login, pero los archivos de datos NO deben quedar
> en un repositorio público. Antes de subir, en GitHub:
> Settings → General → Danger Zone → Change visibility → **Private**.

## Contenido
- `app.py` — aplicación (lee de `data/`).
- `data/` — Data, Campos dashboard y cronograma vigentes + `ultima_actualizacion.json`.
- `ingesta_publicar.py` / `ingesta_publicar.bat` — copia los archivos del día, escribe
  la fecha/hora de actualización y hace `git commit/push`.
- `.streamlit/config.toml`, `requirements.txt`.

## Despliegue en Streamlit Community Cloud
1. Repo privado creado (vacío) en GitHub.
2. Subir este contenido:
   ```
   cd deploy_panel
   git init
   git remote add origin https://github.com/dannrob23/Panel-de-Control-Mantenimientos-3.git
   git branch -M main
   git add -A
   git commit -m "Panel MT3 - despliegue inicial"
   git push -u origin main
   ```
3. En https://share.streamlit.io → **New app** → repositorio, branch `main`,
   main file `app.py` → **Deploy**.
4. Queda una URL pública tipo `https://<nombre>.streamlit.app` (sin login).

## Actualización diaria
1. Copiar los `.xlsx` del día en `..\Ingesta de datos diaria`.
2. Doble clic en `ingesta_publicar.bat` (o programar tarea diaria de Windows que lo ejecute).
3. El script actualiza `data/`, escribe la **fecha/hora** y hace push.
4. Streamlit Cloud redepliega y la app muestra “🕒 Última actualización”.

## Modo público (sin login)
Por defecto `MODO_PUBLICO = True`: seriales y placas se enmascaran (`****1234`) y
las observaciones con datos personales no se muestran. Para uso interno se puede
desactivar creando `.streamlit/secrets.toml` con:
```
modo_publico = false
```
(No subir ese archivo al repositorio.)
