@echo off
chcp 65001>nul
setlocal enabledelayedexpansion
title Actualizar Dashboard MT3
REM ==========================================================================
REM  ACTUALIZAR DASHBOARD MT3   (ciclo completo en un solo doble clic)
REM
REM  Que hace, en orden:
REM    1. Revisa que hay Excel en la carpeta "Ingesta de datos diaria"
REM    2. Revisa la columna AA (sedes-jefatura) y avisa de nombres por corregir
REM    3. Trae los ultimos cambios de GitHub (repo: deploy_panel, rama main)
REM    4. Ejecuta la ingesta: ajusta columnas, valida y anonimiza
REM    5. Regenera data\ y deploy_panel\data\ + commit + push a GitHub
REM       Streamlit Cloud se redespliega solo en 1-2 minutos.
REM
REM  Modos:  (sin argumento) ciclo completo  |  revisar  |  pull  |  local  |  ayuda
REM  Se puede lanzar desde la raiz o desde "Ingesta de datos diaria".
REM  El repositorio debe seguir siendo PRIVADO (contiene datos anonimizados
REM  y la documentacion del proyecto; la raiz NO se publica nunca).
REM ==========================================================================

REM --- 1) Ubicar la raiz del proyecto ---------------------------------------
REM  (la raiz es la carpeta que contiene ingesta.py y deploy_panel)
set "RAIZ=%~dp0"
if not exist "%RAIZ%ingesta.py" set "RAIZ=%~dp0..\"
cd /d "%RAIZ%"

REM --- 2) Configuracion de git para este equipo -----------------------------
REM  En este PC el backend TLS "schannel" falla con SEC_E_NO_CREDENTIALS al
REM  conectar con GitHub. La ingesta ya usa "openssl"; aqui se replica igual
REM  para que el pull de este archivo tambien funcione.
set "GIT_CONFIG_COUNT=1"
set "GIT_CONFIG_KEY_0=http.sslBackend"
set "GIT_CONFIG_VALUE_0=openssl"
set "GIT_TERMINAL_PROMPT=0"
set "GCM_INTERACTIVE=never"

REM --- 3) Modo de trabajo ---------------------------------------------------
set "MODO=completo"
if /i "%~1"=="pull"         set "MODO=pull"
if /i "%~1"=="traer"        set "MODO=pull"
if /i "%~1"=="local"        set "MODO=local"
if /i "%~1"=="sin-publicar" set "MODO=local"
if /i "%~1"=="revisar"      set "MODO=revisar"
if /i "%~1"=="ayuda"        set "MODO=ayuda"
if /i "%~1"=="help"         set "MODO=ayuda"
if /i "%~1"=="-h"           set "MODO=ayuda"
if /i "%~1"=="--help"       set "MODO=ayuda"

echo ==========================================================
echo   ACTUALIZAR DASHBOARD MT3
echo ==========================================================
echo.

if "%MODO%"=="ayuda" (
    echo   Uso:   actualizar_panel.bat [modo]
    echo.
    echo     sin argumento   Ciclo completo: datos + commit + push a GitHub
    echo     revisar         Revisar la columna AA del archivo - jefaturas - sin publicar
    echo     pull            Solo traer los cambios de GitHub a este PC
    echo     local           Regenerar los datos SIN publicar - solo prueba
    echo     ayuda           Esta pantalla
    echo.
    echo   Uso normal: doble clic en el archivo, sin escribir nada.
    echo   Datos personales: el repositorio debe seguir siendo PRIVADO.
    echo.
    goto :fin_ok
)

echo   NO SUBIR ESTE REPOSITORIO A PUBLICO: contiene datos personales.
echo   El repositorio debe seguir siendo PRIVADO.
echo.

REM --- 4) Python disponible? (solo lo necesitan los modos con ingesta) ------
if not "%MODO%"=="pull" (
    where python >nul 2>nul
    if errorlevel 1 (
        echo   [ERROR] No encuentro "python" en el PATH de este equipo.
        echo   Instala Python o abre este archivo desde el PC del administrador.
        echo.
        goto :fin_error
    )
)

REM --- 5) Elegir el camino --------------------------------------------------
if "%MODO%"=="pull"    goto :traer
if "%MODO%"=="local"   goto :local
if "%MODO%"=="revisar" goto :revisar
goto :completo


REM ==========================================================================
REM  MODO COMPLETO  (uso normal: actualizar los datos y publicarlos)
REM ==========================================================================
:completo
echo [1/5] Revisando los Excel de "Ingesta de datos diaria"...
call :revisar_excel
if errorlevel 1 goto :fin_error
echo.
echo [2/5] Revisando la columna AA (sedes-jefatura de la Data)...
call :revisar_aa
if errorlevel 1 goto :fin_error
echo.
echo [3/5] Trayendo los ultimos cambios de GitHub...
git -C "deploy_panel" pull --ff-only origin main
if errorlevel 1 (
    echo   [AVISO] No se pudieron traer los cambios de GitHub.
    echo   Se continua igual: si el push lo rechaza, al final sale el comando exacto.
) else (
    echo   [OK] Repositorio al dia.
)
echo.
echo [4/5] Generando datos y publicando...
echo   - tarda 1-3 minutos; NO cierres esta ventana
echo.
>> "tools\actualizar_panel.log" echo [%date% %time%] modo=completo
python ingesta.py --publicar
set "RC=!ERRORLEVEL!"
>> "tools\actualizar_panel.log" echo [%date% %time%] modo=completo rc=!RC!
echo.
echo [5/5] Resultado de la publicacion
if "!RC!"=="0" echo   [OK] Dashboard actualizado y publicado en GitHub.
if "!RC!"=="0" echo        Streamlit Cloud se actualiza en 1-2 minutos.
if "!RC!"=="1" echo   [ERROR] Faltan archivos o la carpeta de ingesta. Revisa los mensajes de arriba.
if "!RC!"=="2" echo   [ERROR] Faltan columnas obligatorias en los Excel: NO se publico nada.
if "!RC!"=="2" echo        Avisa al responsable del panel antes de reintentar.
if "!RC!"=="3" echo   [AVISO] Los datos y el commit quedaron listos, pero el push fallo.
if "!RC!"=="3" echo        Ejecuta una sola vez y vuelve a intentar:
if "!RC!"=="3" echo            git -C deploy_panel push origin main
goto :verificacion


REM ==========================================================================
REM  MODO REVISAR  (solo revisa la columna AA: jefaturas de la Data)
REM ==========================================================================
:revisar
echo [1/1] Revisando la columna AA (sedes-jefatura de la Data)...
call :revisar_aa
if errorlevel 1 goto :fin_error
echo.
echo   Revision terminada. No se cambio ni se publico nada.
goto :fin_ok


REM ==========================================================================
REM  MODO PULL  (solo traer los ultimos cambios del panel a este PC)
REM ==========================================================================
:traer
echo [1/1] Trayendo los ultimos cambios de GitHub...
git -C "deploy_panel" pull --ff-only origin main
if errorlevel 1 (
    echo.
    echo   [ERROR] No se pudieron traer los cambios.
    echo   Suele pasar si hay un conflicto en data\*.xlsx, por ejemplo cuando
    echo   se publico desde otro PC al mismo tiempo.
    echo   NO borres nada: avisa al responsable del panel.
    echo.
    goto :fin_error
)
echo   [OK] Cambios traidos.
goto :verificacion


REM ==========================================================================
REM  MODO LOCAL  (regenerar los datos sin publicar: sirve para probar)
REM ==========================================================================
:local
echo [1/3] Revisando los Excel de "Ingesta de datos diaria"...
call :revisar_excel
if errorlevel 1 goto :fin_error
echo.
echo [2/3] Revisando la columna AA (sedes-jefatura de la Data)...
call :revisar_aa
if errorlevel 1 goto :fin_error
echo.
echo [3/3] Generando los datos SIN publicar...
echo   - tarda 1-3 minutos; NO cierres esta ventana
echo.
>> "tools\actualizar_panel.log" echo [%date% %time%] modo=local
python ingesta.py --seco
set "RC=!ERRORLEVEL!"
>> "tools\actualizar_panel.log" echo [%date% %time%] modo=local rc=!RC!
echo.
if "!RC!"=="0" echo   [OK] Datos regenerados en tu PC. NO se subio nada a GitHub.
if "!RC!"=="0" echo        Para publicarlos: vuelve a ejecutar este archivo sin argumentos.
if not "!RC!"=="0" echo   [ERROR] La ingesta no termino bien (codigo !RC!). Revisa los mensajes de arriba.
goto :verificacion


REM ==========================================================================
REM  SUBRUTINA: revisar la columna AA del Excel de Data (sedes-jefatura)
REM  No detiene el flujo: si hay avisos, pide Enter para que se alcancen a leer.
REM ==========================================================================
:revisar_aa
if not exist "tools\revisar_jefaturas.py" (
    echo   [AVISO] No encuentro tools\revisar_jefaturas.py: se omite la revision.
    echo   Vuelve a preparar el equipo con Preparar_Equipo_Nuevo.bat.
    exit /b 0
)
%PY% "tools\revisar_jefaturas.py"
if errorlevel 2 (
    echo.
    echo   [ERROR] No se pudo revisar la columna AA del archivo de Data.
    exit /b 1
)
if errorlevel 1 (
    echo.
    echo   [AVISO] Conviene mirar los avisos de arriba antes de publicar.
    echo   El panel corrige lo que puede al mostrar, pero el Excel de origen es el que manda.
    echo.
    set "RESP="
    set /p "RESP=  Enter para continuar, o Ctrl+C para cancelar: "
)
exit /b 0


REM ==========================================================================
REM  SUBRUTINA: revisar los Excel de la carpeta de ingesta
REM ==========================================================================
:revisar_excel
set "N_XLSX=0"
pushd "%RAIZ%Ingesta de datos diaria" 2>nul
if errorlevel 1 (
    echo   [ERROR] No existe la carpeta "Ingesta de datos diaria" en:
    echo           %RAIZ%
    echo   Copia ahi los Excel del dia y vuelve a ejecutar este archivo.
    exit /b 1
)
for /f %%F in ('dir /b /a-d "*.xlsx" 2^>nul ^| find /c /v ""') do set "N_XLSX=%%F"
if "%N_XLSX%"=="0" (
    echo   [ERROR] No hay ningun archivo .xlsx en "Ingesta de datos diaria".
    echo   Copia ahi los Excel del dia y vuelve a ejecutar este archivo.
    popd
    exit /b 1
)
set "PRIMERO="
set "FECHA_X="
for /f "delims=" %%F in ('dir /b /a-d /o-d "*.xlsx" 2^>nul') do (
    if not defined PRIMERO (
        set "PRIMERO=%%~nxF"
        set "FECHA_X=%%~tF"
    )
)
echo   Excel en la carpeta : %N_XLSX%
echo   El mas reciente     : %PRIMERO%
echo   Fecha del mas nuevo : %FECHA_X%
echo   - La ingesta elige sola el mas nuevo de cada tipo: Data, Campos, cronogramas
popd
exit /b 0


REM ==========================================================================
REM  VERIFICACION FINAL
REM ==========================================================================
:verificacion
echo.
echo ----------------------------------------------------------
echo   Estado del repositorio que ve el panel en la nube:
git -C "deploy_panel" log --oneline -1
git -C "deploy_panel" status -sb
echo ----------------------------------------------------------
echo.
echo   Si arriba NO aparece la palabra "ahead", todo quedo subido a GitHub.
echo   Recarga el panel en el navegador con Ctrl+F5 (recarga dura).
echo.
goto :fin_ok


:fin_ok
endlocal
pause
exit /b 0

:fin_error
echo.
endlocal
pause
exit /b 1
