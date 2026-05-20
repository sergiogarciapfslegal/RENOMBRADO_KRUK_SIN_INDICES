@echo off
REM ========================================================================
REM  LANZADOR PIPELINE KRUK - DEMANDAS
REM  Dispara el job de Jenkins que procesa la remesa actual:
REM    1. Naming CSV   2. Stamp PDF   3. Sign PDF   4. Convertir rutas
REM
REM  Antes de usar, rellena los 4 valores marcados con <<...>> mas abajo.
REM  Para obtener el API TOKEN: Jenkins > tu usuario > Configurar > Add new Token
REM ========================================================================

setlocal EnableDelayedExpansion

REM ---- CONFIGURACION (rellenar) -------------------------------------------
set "JENKINS_URL=<<https://jenkins.tuempresa.es>>"
set "JOB_NAME=<<nombre-del-job>>"
set "JENKINS_USER=<<usuario.jenkins>>"
set "JENKINS_TOKEN=<<11abcd...token...ef99>>"
REM -------------------------------------------------------------------------

title KRUK - Lanzar pipeline demandas

echo.
echo ========================================================================
echo   KRUK - Pipeline de demandas
echo ========================================================================
echo.
echo   Jenkins : %JENKINS_URL%
echo   Job     : %JOB_NAME%
echo   Usuario : %JENKINS_USER%
echo.
set /p CONFIRMAR="Pulsa ENTER para lanzar la build (o CTRL+C para cancelar)... "
echo.

REM ---- Obtener crumb (proteccion CSRF de Jenkins) -------------------------
echo [1/2] Obteniendo crumb de seguridad...
for /f "tokens=2 delims=:" %%A in (
    'curl -s -u "%JENKINS_USER%:%JENKINS_TOKEN%" "%JENKINS_URL%/crumbIssuer/api/json" ^| findstr /C:"crumb"'
) do (
    set "CRUMB_RAW=%%A"
)
set "CRUMB=!CRUMB_RAW:"=!"
set "CRUMB=!CRUMB:,!CRUMB_FIELD=!"
for /f "tokens=1 delims=," %%B in ("!CRUMB!") do set "CRUMB_VALUE=%%B"

REM ---- Lanzar la build ----------------------------------------------------
echo [2/2] Lanzando build en Jenkins...
curl -s -o nul -w "  HTTP %%{http_code}\n" ^
     -X POST ^
     -u "%JENKINS_USER%:%JENKINS_TOKEN%" ^
     -H "Jenkins-Crumb:!CRUMB_VALUE!" ^
     "%JENKINS_URL%/job/%JOB_NAME%/build"

if errorlevel 1 (
    echo.
    echo  *** ERROR al contactar con Jenkins. Revisa la URL, el usuario y el token.
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================================================
echo   Build lanzada. Sigue el progreso en:
echo   %JENKINS_URL%/job/%JOB_NAME%/
echo ========================================================================
echo.
echo Se abrira el navegador con el estado de la build...
timeout /t 2 /nobreak >nul
start "" "%JENKINS_URL%/job/%JOB_NAME%/"

echo.
pause
endlocal
