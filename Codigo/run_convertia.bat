@echo off
setlocal enabledelayedexpansion
REM PENDIENTE: cuando se renombre la carpeta raiz de CLAUDE a "Agente Convertia", cambiar la linea de abajo
cd /d "C:\Users\dy2059\Mis Archivos\CLAUDE\Codigo"

set FECHA_INICIO=%~1
set FECHA_FIN=%~2
set MODO_FORZAR=%~3

set ARGS=
if not "%FECHA_INICIO%"=="" (
    if "%FECHA_FIN%"=="" (
        echo Falta la fecha fin. Uso: run_convertia.bat [fecha-inicio] [fecha-fin] [forzar]
        echo Ejemplo: run_convertia.bat 2026-08-01 2026-08-16
        exit /b 1
    )
    set ARGS=--fecha-inicio %FECHA_INICIO% --fecha-fin %FECHA_FIN%
)
if /i "%MODO_FORZAR%"=="forzar" set ARGS=%ARGS% --forzar

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set STAMP=%%i
set LOG=run_log_!STAMP!.txt

echo INICIO %DATE% %TIME% > "!LOG!"
python -u "descargar_reportes_convertia.py" !ARGS! >> "!LOG!" 2>&1
echo EXIT_CODE=%ERRORLEVEL% >> "!LOG!"
echo FIN %DATE% %TIME% >> "!LOG!"

echo Listo. Log: !LOG!
endlocal
