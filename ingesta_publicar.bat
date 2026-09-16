@echo off
chcp 65001>nul
cd /d "%~dp0"
echo ==========================================
echo  INGESTA DIARIA - PANEL MANTENIMIENTO 3
echo ==========================================
echo.
echo NO SUBIR ESTE REPOSITORIO A PUBLICO: contiene datos personales.
echo El repositorio debe ser PRIVADO.
echo.
echo [1/3] Sincronizando app.py y documentacion desde la raiz del proyecto...
python sincronizar_app.py
if errorlevel 1 (
  echo [AVISO] No se pudo sincronizar; se publicara la version actual de deploy_panel.
)
echo.
echo [2/3] Anonimizando datos (data/Data_actual.xlsx y data/Campos_actual.xlsx)...
python sanitizar_publico.py
echo.
echo [3/3] Publicando en GitHub (commit + push)...
python ingesta_publicar.py --no-copy
echo.
echo Proceso terminado.
pause
