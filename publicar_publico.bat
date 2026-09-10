@echo off
chcp 65001>nul
cd /d "%~dp0"
echo ==========================================
echo  PUBLICAR PANEL MT3 (repo PUBLICO)
echo ==========================================
echo Genera data anonimizada, actualiza la fecha y sube a GitHub.
echo.
python sanitizar_publico.py
git add -A
git commit -m "Actualizacion anonimizada"
git push
echo.
echo Listo. Streamlit Cloud se actualiza automaticamente.
pause
