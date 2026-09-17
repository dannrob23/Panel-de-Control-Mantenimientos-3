@echo off
chcp 65001>nul
setlocal
cd /d "%~dp0"
title DASHBOARD MANTENIMIENTO PREVENTIVO 3
color 0F
echo ==================================================
echo    DASHBOARD MANTENIMIENTO PREVENTIVO 3 - MT3
echo ==================================================
echo.

python -c "import sys">nul 2>nul
if errorlevel 1 goto no_python

python -c "import streamlit, pandas, openpyxl, plotly">nul 2>nul
if errorlevel 1 goto instalar

goto verificar_puerto

:no_python
echo [ERROR] No se encontro Python en el PATH.
echo         Instale Python 3.12 marcando "Add Python to PATH".
pause
exit /b 1

:instalar
echo Instalando dependencias por primera vez...
python -m pip install -r requirements.txt
echo.
goto verificar_puerto

:verificar_puerto
netstat -ano | findstr /R /C:":8510 .*LISTENING">nul 2>nul
if errorlevel 1 goto arrancar
echo La aplicacion ya esta activa en http://localhost:8510
start "" http://localhost:8510
ping -n 5 127.0.0.1>nul
exit /b 0

:arrancar
echo Abriendo navegador y levantando el servidor...
echo URL: http://localhost:8510
echo Para detener: cierre esta ventana o presione Ctrl+C
echo.
start "abrir-dashboard" cmd /c "ping -n 9 127.0.0.1>nul & start http://localhost:8510"
python -m streamlit run app.py --server.port 8510

echo.
echo El servidor termino.
pause
exit /b 0
