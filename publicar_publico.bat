@echo off
chcp 65001>nul
cd /d "%~dp0"
echo ==========================================================
echo   PUBLICAR PANEL MT3  (repositorio privado)
echo ==========================================================
echo.
echo   Este lanzador usa el publicador unico de la raiz del proyecto.
echo   Hace: ajustar columnas + anonimizar + data\ + commit + push.
echo.
pause
python "..\ingesta.py" --publicar
echo.
pause

