@echo off
chcp 65001>nul
cd /d "%~dp0"
echo ==========================================
echo  ACTUALIZAR PANEL MT3
echo  (trae los ultimos cambios de GitHub)
echo ==========================================
echo.
echo NO SUBIR ESTE REPOSITORIO A PUBLICO: contiene datos personales.
echo El repositorio debe ser PRIVADO.
echo.
echo Trayendo cambios...
echo.
git pull origin main
echo.
echo ------------------------------
echo Version actual:
git log --oneline -1
echo ------------------------------
echo.
echo Si no hay errores:
echo   1) Cierra esta ventana
echo   2) Recarga el panel con Ctrl+F5
echo.
echo Si aparece un conflicto, NO borres nada: avisa al responsable.
echo.
pause
