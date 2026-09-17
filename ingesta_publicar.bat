@echo off
chcp 65001>nul
cd /d "%~dp0"
echo ==========================================================
echo   PUBLICAR PANEL MT3  (ingesta + commit + push a GitHub)
echo ==========================================================
echo.
echo   Toma los Excel de "..\Ingesta de datos diaria", los ajusta,
echo   actualiza data\ y deploy_panel\data, y sube los cambios.
echo.
pause
python ingesta.py --publicar
echo.
pause
