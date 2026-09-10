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
python ingesta_publicar.py
echo.
echo Proceso terminado.
pause
