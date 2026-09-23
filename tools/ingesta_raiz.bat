@echo off
chcp 65001>nul
title Ingesta diaria - Panel MT3
REM Ubica la raiz del proyecto (funciona desde la raiz o desde "Ingesta de datos diaria")
set "RAIZ=%~dp0"
if not exist "%RAIZ%ingesta.py" set "RAIZ=%~dp0..\"
cd /d "%RAIZ%"
echo ==========================================================
echo   INGESTA DIARIA - PANEL MANTENIMIENTO PREVENTIVO 3
echo ==========================================================
echo.
echo   [1]  Revisar y generar datos        (no publica)
echo   [2]  Generar y PUBLICAR en GitHub
echo   [3]  Solo analizar columnas         (reporte)
echo   [4]  Solo copia local               (no toca deploy_panel)
echo   [0]  Salir
echo.
set /p opcion="  Elige una opcion [1]: "
if "%opcion%"=="" set opcion=1
echo.

if "%opcion%"=="1" python ingesta.py
if "%opcion%"=="2" python ingesta.py --publicar
if "%opcion%"=="3" python ingesta.py --analisis
if "%opcion%"=="4" python ingesta.py --solo-local

echo.
pause
