@echo off
chcp 65001>nul
title Actualizacion automatica del panel MT3
REM Ubica la raiz del proyecto (funciona desde la raiz o desde "Ingesta de datos diaria")
set "RAIZ=%~dp0"
if not exist "%RAIZ%vigilar_ingesta.py" set "RAIZ=%~dp0..\"
cd /d "%RAIZ%"
echo ==========================================================
echo   ACTUALIZACION AUTOMATICA DEL PANEL MT3
echo ==========================================================
echo.
echo   Deja esta ventana ABIERTA.
echo.
echo   Cada vez que guardes los Excel en la carpeta
echo   "Ingesta de datos diaria", el panel se actualiza
echo   y se publica solo en GitHub.
echo.
echo   Para detenerlo: cierra la ventana o pulsa Ctrl+C.
echo.
pause
python vigilar_ingesta.py
echo.
echo La vigilancia termino.
pause
