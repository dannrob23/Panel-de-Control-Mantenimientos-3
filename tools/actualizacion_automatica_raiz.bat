@echo off
chcp 65001>nul
cd /d "%~dp0"
title Actualizacion automatica del panel MT3
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
