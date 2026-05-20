@echo off
REM ========================================================================
REM  LANZADOR PIPELINE KRUK - DEMANDAS
REM  Dispara el job de Jenkins que procesa la remesa actual:
REM    1. Naming CSV   2. Stamp PDF   3. Sign PDF   4. Convertir rutas
REM ========================================================================

setlocal EnableDelayedExpansion

REM ---- CONFIGURACION (rellenar SIN los signos << >>) ----------------------
set "JENKINS_URL=http://jenkinsaws.pfslegal.es:8080"
set "JOB_NAME=RENOMBRADO_KRUK_SIN_INDICES"
set "JENKINS_USER=jenkins_legal"
set "JENKINS_TOKEN=11caa2afbdef97d4fc9e15ae7a5e274ca2"
REM -------------------------------------------------------------------------

REM ---- Auto-limpieza defensiva: quita << y >> por si quedaron ------------
set "JENKINS_URL=!JENKINS_URL:<<=!"
set "JENKINS_URL=!JENKINS_URL:>>=!"
set "JOB_NAME=!JOB_NAME:<<=!"
set "JOB_NAME=!JOB_NAME:>>=!"
set "JENKINS_USER=!JENKINS_USER:<<=!"
set "JENKINS_USER=!JENKINS_USER:>>=!"
set "JENKINS_TOKEN=!JENKINS_TOKEN:<<=!"
set "JENKINS_TOKEN=!JENKINS_TOKEN:>>=!"

title KRUK - Lanzar pipeline demandas

echo.
echo ========================================================================
echo   KRUK - Pipeline de demandas
echo ========================================================================
echo.
echo   Jenkins : !JENKINS_URL!
echo   Job     : !JOB_NAME!
echo   Usuario : !JENKINS_USER!
echo.

REM ---- Comprobar que curl existe ------------------------------------------
where curl >nul 2>nul
if errorlevel 1 (
    echo  *** ERROR: 'curl' no esta disponible en este Windows.
    echo  *** Windows 10/11 y Server 2019+ lo traen de serie en System32.
    echo.
    pause
    exit /b 1
)

set /p CONFIRMAR="Pulsa ENTER para lanzar la build (CTRL+C para cancelar)... "
echo.

REM ---- Fichero temporal para respuestas de curl ---------------------------
set "TMPFILE=%TEMP%\jenkins_response_%RANDOM%.txt"

REM ========================================================================
REM  PASO 1: Obtener crumb CSRF
REM ========================================================================
echo [1/2] Obteniendo crumb CSRF...
curl -s -u "!JENKINS_USER!:!JENKINS_TOKEN!" ^
     -o "!TMPFILE!" ^
     -w "HTTP %%{http_code}" ^
     "!JENKINS_URL!/crumbIssuer/api/json" > "%TEMP%\http_code.txt"

set /p HTTP_CRUMB=<"%TEMP%\http_code.txt"
del "%TEMP%\http_code.txt" >nul 2>nul

echo   Respuesta: !HTTP_CRUMB!

REM Detectar si Jenkins no usa CSRF (404 en crumbIssuer)
set "CRUMB_VALUE="
if "!HTTP_CRUMB!"=="HTTP 200" (
    REM Extraer el campo "crumb":"..." del JSON
    for /f "usebackq tokens=1,2 delims=:," %%A in ("!TMPFILE!") do (
        echo %%A | findstr /C:"crumb" >nul && (
            set "TMPVAL=%%B"
            set "TMPVAL=!TMPVAL:"=!"
            set "TMPVAL=!TMPVAL: =!"
            if not defined CRUMB_VALUE if not "!TMPVAL!"=="" set "CRUMB_VALUE=!TMPVAL!"
        )
    )
    echo   Crumb obtenido: !CRUMB_VALUE!
) else if "!HTTP_CRUMB!"=="HTTP 404" (
    echo   Jenkins sin proteccion CSRF habilitada, continuando sin crumb.
) else (
    echo.
    echo  *** ERROR obteniendo crumb. Respuesta de Jenkins:
    echo  ----------------------------------------------------------
    type "!TMPFILE!"
    echo.
    echo  ----------------------------------------------------------
    echo  Verifica URL, usuario y token. Si el server requiere VPN,
    echo  comprueba que estas conectado.
    echo.
    del "!TMPFILE!" >nul 2>nul
    pause
    exit /b 1
)

del "!TMPFILE!" >nul 2>nul

REM ========================================================================
REM  PASO 2: Lanzar la build
REM ========================================================================
echo.
echo [2/2] Lanzando build...

set "CURL_ARGS=-s -o "!TMPFILE!" -w "HTTP %%{http_code}" -X POST -u "!JENKINS_USER!:!JENKINS_TOKEN!""
if defined CRUMB_VALUE set "CURL_ARGS=!CURL_ARGS! -H "Jenkins-Crumb:!CRUMB_VALUE!""

curl !CURL_ARGS! "!JENKINS_URL!/job/!JOB_NAME!/build" > "%TEMP%\http_code.txt"

set /p HTTP_BUILD=<"%TEMP%\http_code.txt"
del "%TEMP%\http_code.txt" >nul 2>nul

echo   Respuesta: !HTTP_BUILD!

REM Jenkins devuelve 201 Created cuando encola la build correctamente
echo !HTTP_BUILD! | findstr /R "20[01]" >nul
if errorlevel 1 (
    echo.
    echo  *** ERROR lanzando la build. Respuesta de Jenkins:
    echo  ----------------------------------------------------------
    type "!TMPFILE!" 2>nul
    echo.
    echo  ----------------------------------------------------------
    del "!TMPFILE!" >nul 2>nul
    pause
    exit /b 1
)

del "!TMPFILE!" >nul 2>nul

echo.
echo ========================================================================
echo   BUILD LANZADA CORRECTAMENTE
echo   Sigue el progreso en:
echo   !JENKINS_URL!/job/!JOB_NAME!/
echo ========================================================================
echo.
echo Se abrira el navegador con el job...
timeout /t 2 /nobreak >nul
start "" "!JENKINS_URL!/job/!JOB_NAME!/"

echo.
pause
endlocal
