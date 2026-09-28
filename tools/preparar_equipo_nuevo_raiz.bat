@echo off
chcp 65001>nul
title Preparar equipo nuevo - Panel MT3
setlocal

REM ==========================================================================
REM  PREPARAR ESTE EQUIPO PARA ACTUALIZAR EL PANEL MT3
REM
REM  Deja cualquier PC Windows listo para publicar la data del dia:
REM    1. Revisa Python y Git y ofrece instalarlos con winget si faltan
REM    2. Clona el repositorio en la carpeta del proyecto
REM    3. Copia los scripts y crea "Ingesta de datos diaria"
REM    4. Instala pandas, numpy y openpyxl
REM    5. Autentica con GitHub, una sola vez por equipo
REM    6. Crea un acceso directo en el Escritorio
REM
REM  Uso normal : doble clic. Se ejecuta UNA VEZ por equipo.
REM  Uso avanzado: Preparar_Equipo_Nuevo.bat "C:\ruta\destino"  sin preguntas
REM
REM  Despues de esto, el uso diario es el mismo de siempre:
REM  copiar los Excel en "Ingesta de datos diaria" y doble clic en
REM  "Actualizar Panel MT3".
REM ==========================================================================

set "REPO=https://github.com/dannrob23/Panel-de-Control-Mantenimientos-3.git"
REM  Backend TLS OpenSSL: en varios equipos el backend schannel falla con
REM  SEC_E_NO_CREDENTIALS y no deja clonar ni subir a GitHub.
set "GIT_CONFIG_COUNT=1"
set "GIT_CONFIG_KEY_0=http.sslBackend"
set "GIT_CONFIG_VALUE_0=openssl"
set "DESTINO=%~1"
if "%DESTINO%"=="" set "DESTINO=%USERPROFILE%\Documents\PROYECTOS\CONTROL INVENTARIO MANTENIMIENTOS 3"

echo ==========================================================
echo   PREPARAR ESTE EQUIPO - PANEL MT3
echo ==========================================================
echo.
echo   Deja este PC listo para actualizar el panel desde aqui.
echo   Se ejecuta UNA VEZ por equipo y tarda unos minutos.
echo.

REM --- 1) Python -----------------------------------------------------------
echo [1/6] Revisando Python...
set "PY="
where python >nul 2>nul
if not errorlevel 1 set "PY=python"
if not defined PY (
    where py >nul 2>nul
    if not errorlevel 1 set "PY=py"
)
if not defined PY (
    echo   No encontre Python en este equipo.
    call :instalar Python Python.Python.3.12
    goto :fin_instalar
)
for /f "delims=" %%V in ('%PY% --version 2^>^&1') do echo   OK %%V

REM --- 2) Git --------------------------------------------------------------
echo.
echo [2/6] Revisando Git...
where git >nul 2>nul
if errorlevel 1 (
    echo   No encontre Git en este equipo.
    call :instalar Git Git.Git
    goto :fin_instalar
)
for /f "delims=" %%V in ('git --version 2^>^&1') do echo   OK %%V

REM --- 3) Carpeta del proyecto y clon --------------------------------------
echo.
echo [3/6] Carpeta del proyecto
echo   Se instalara en:
echo     %DESTINO%
echo.
if not "%~1"=="" goto :carpeta_lista
set "RESP="
set /p "RESP=  Enter para usarla, o escribe otra ruta: "
if not "%RESP%"=="" set "DESTINO=%RESP%"
:carpeta_lista
if not exist "%DESTINO%" mkdir "%DESTINO%"
cd /d "%DESTINO%"
if errorlevel 1 (
    echo   [ERROR] No pude entrar a esa carpeta. Revisa la ruta y vuelve a intentar.
    goto :fin_error
)
if exist "%DESTINO%\deploy_panel\.git" (
    echo   El repositorio ya estaba en este equipo: trayendo los ultimos cambios...
    git -C "deploy_panel" pull --ff-only
    if errorlevel 1 (
        echo   [AVISO] No pude actualizar el repositorio. Se continua con lo que hay.
    )
) else (
    echo   Clonando el repositorio de GitHub. Puede tardar un poco...
    git clone "%REPO%" "deploy_panel"
    if errorlevel 1 (
        echo   [ERROR] No se pudo clonar. Revisa tu conexion a internet y la ruta.
        goto :fin_error
    )
)

REM --- 4) Copiar los archivos de trabajo a la raiz -------------------------
echo.
echo [4/6] Preparando los archivos de trabajo...
if not exist "Ingesta de datos diaria" mkdir "Ingesta de datos diaria"
if not exist "tools" mkdir "tools"
copy /y "deploy_panel\app.py" "app.py" >nul
copy /y "deploy_panel\requirements.txt" "requirements.txt" >nul
copy /y "deploy_panel\README.md" "README.md" >nul
copy /y "deploy_panel\TUTORIAL.md" "TUTORIAL.md" >nul
copy /y "deploy_panel\tools\ingesta_raiz.py" "ingesta.py" >nul
copy /y "deploy_panel\tools\vigilar_ingesta_raiz.py" "vigilar_ingesta.py" >nul
copy /y "deploy_panel\tools\ingesta_raiz.bat" "ingesta.bat" >nul
copy /y "deploy_panel\tools\actualizacion_automatica_raiz.bat" "actualizacion_automatica.bat" >nul
copy /y "deploy_panel\tools\actualizar_panel_raiz.bat" "actualizar_panel.bat" >nul
copy /y "deploy_panel\tools\verificar_datos_raiz.py" "tools\verificar_datos.py" >nul
if not exist "actualizar_panel.bat" (
    echo   [ERROR] No pude copiar los scripts desde deploy_panel\tools\.
    echo   Abre Git Bash o CMD en "%DESTINO%" y ejecuta:
    echo     git clone %REPO% deploy_panel
    goto :fin_error
)
if not exist "ingesta.py" (
    echo   [ERROR] Falta ingesta.py. El repositorio no se clono completo.
    goto :fin_error
)
echo   OK archivos listos.

REM --- 5) Librerias de Python ---------------------------------------------
echo.
echo [5/6] Librerias de Python que necesita la ingesta...
%PY% -c "import pandas, numpy, openpyxl" >nul 2>nul
if errorlevel 1 (
    echo   Faltan: instalando. Puede tardar unos minutos...
    %PY% -m pip install --upgrade pip >nul 2>nul
    %PY% -m pip install pandas numpy openpyxl
    %PY% -c "import pandas, numpy, openpyxl" >nul 2>nul
    if errorlevel 1 (
        echo   [ERROR] No se pudieron instalar las librerias.
        echo   Prueba a mano:  %PY% -m pip install pandas numpy openpyxl
        goto :fin_error
    )
    echo   OK instaladas.
) else (
    echo   OK ya estaban instaladas.
)

REM --- 6) Acceso a GitHub y acceso directo --------------------------------
echo.
echo [6/6] Acceso a GitHub. Solo la primera vez en este equipo.
echo   Si aparece una ventana pidiendo usuario y token, autenticate.
echo.
git -C "deploy_panel" fetch origin
if errorlevel 1 (
    echo   [ATENCION] No pude conectar con GitHub o autenticarme.
    echo   Vuelve a ejecutar este archivo y autenticate cuando lo pida.
    echo   Con Git para Windows basta iniciar sesion con tu cuenta de GitHub.
) else (
    echo   OK acceso concedido: la credencial queda guardada en este equipo.
)

set "ACCESO="
set "LANZADOR=%DESTINO%\Actualizar Panel MT3.bat"
> "%LANZADOR%" echo @echo off
>> "%LANZADOR%" echo cd /d "%DESTINO%"
>> "%LANZADOR%" echo call actualizar_panel.bat
for /f "usebackq delims=" %%D in (`powershell -NoProfile -Command "[Environment]::GetFolderPath('Desktop')"`) do set "ESC=%%D"
if defined ESC copy /y "%LANZADOR%" "%ESC%\Actualizar Panel MT3.bat" >nul 2>nul
if exist "%ESC%\Actualizar Panel MT3.bat" set "ACCESO=1"
if defined ACCESO (
    echo   OK acceso directo creado en el Escritorio.
) else (
    echo   [AVISO] No pude crear el acceso directo en el Escritorio.
    echo   Usa el lanzador que quedo en la carpeta del proyecto:
    echo     %LANZADOR%
)

REM --- Comprobacion final -------------------------------------------------
echo.
echo ----------------------------------------------------------
echo   Comprobando que todo quedo bien en este equipo...
echo ----------------------------------------------------------
%PY% "tools\verificar_datos.py"

echo.
echo ==========================================================
echo   LISTO. Como se usa este equipo de ahora en adelante:
echo ==========================================================
echo.
echo   1. Copia los Excel del dia en:
echo        %DESTINO%\Ingesta de datos diaria
if defined ACCESO (
    echo   2. Doble clic en el Escritorio: "Actualizar Panel MT3"
) else (
    echo   2. Doble clic en: %DESTINO%\actualizar_panel.bat
)
echo   3. Espera 1-3 minutos. Al final te dice si quedo publicado.
echo.
echo   El panel de la nube se actualiza solo 1-2 minutos despues.
echo.
echo   Si te equivocas o algo falla, ejecuta esta misma pantalla asi:
echo        actualizar_panel.bat ayuda
echo.
echo   Para VER el panel en este PC, opcional:
echo        %PY% -m pip install -r requirements.txt
echo        %PY% -m streamlit run app.py
echo.
goto :fin_ok


REM ==========================================================================
REM  SUBRUTINA: instalar Python o Git con winget
REM ==========================================================================
:instalar
where winget >nul 2>nul
if errorlevel 1 (
    echo   [ATENCION] Este Windows no trae winget, no puedo instalarlo solo.
    echo.
    echo   Instala %~1 a mano desde:
    if /i "%~1"=="Python" echo     https://www.python.org/downloads/windows/  - marca Add python.exe to PATH
    if /i "%~1"=="Git" echo     https://git-scm.com/download/win
    echo.
    echo   Despues CIERRA esta ventana y vuelve a ejecutar este archivo.
    exit /b 1
)
echo   Puedo instalarlo ahora con winget.
set "RESP="
set /p "RESP=  Escribe S y Enter para instalarlo: "
if /i not "%RESP%"=="S" (
    echo   Omitido. Instala %~1 y vuelve a ejecutar este archivo.
    exit /b 1
)
winget install --id %~2 -e --accept-source-agreements --accept-package-agreements
echo.
echo   Hecho. CIERRA esta ventana y vuelve a ejecutar este archivo:
echo   esta ventana no ve los programas recien instalados.
exit /b 1


:fin_instalar
echo.
echo   Falta instalar algo antes de continuar. Vuelve a ejecutar este archivo
echo   cuando lo tengas instalado.
endlocal
pause
exit /b 1

:fin_error
echo.
echo   La preparacion no termino bien. Lee los mensajes de arriba.
endlocal
pause
exit /b 1

:fin_ok
endlocal
pause
exit /b 0
