@echo off
setlocal

REM ============================================================
REM ACTUALIZACION NOCTURNA BOT VENDEDORES
REM Sube SOLO los archivos que forman parte de produccion.
REM ============================================================

cd /d E:\PRUEBASEXTRACTOR\Automatizacion\WhatsApp\botvendedores

echo.
echo ========================================
echo   BOT VENDEDORES - ACTUALIZACION
echo ========================================
echo.

echo [1/4] Estado actual de Git:
git status --short

echo.
echo [2/4] Preparando archivos de produccion...

REM Registra modificaciones y eliminaciones de archivos ya controlados.
git add -u

REM Agrega unicamente los archivos actualmente utilizados.
git add actualizar_db.bat ^
        COACH_WORKER_V10.py ^
        FERIADOS.xlsx ^
        Procfile ^
        requirements.txt ^
        vendedores.xlsx ^
        ventas.db ^
        ventas_backup.db ^
        WEBHOOK_VENDEDORES_V3_CONSOLIDADO_ED.py

if errorlevel 1 (
    echo.
    echo ========================================
    echo ERROR preparando cambios
    echo ========================================
    exit /b 1
)

echo.
echo [3/4] Verificando cambios preparados:
git status --short

echo.
echo [4/4] Creando commit...

git diff --cached --quiet
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo No hay cambios para subir
    echo ========================================
    exit /b 0
)

git commit -m "Auto update - BD y archivos de produccion"

if errorlevel 1 (
    echo.
    echo ========================================
    echo ERROR creando commit
    echo ========================================
    exit /b 1
)

echo.
echo Subiendo a GitHub...
git push origin main

if errorlevel 1 (
    echo.
    echo ========================================
    echo ERROR haciendo push a GitHub
    echo ========================================
    exit /b 1
)

echo.
echo ========================================
echo ACTUALIZACION COMPLETADA
echo ========================================
echo.
echo GitHub recibio la actualizacion.
echo Render podra desplegar la nueva ventas.db.
echo El Worker utilizara la BD actualizada a las 08:00.
echo.

endlocal
exit /b 0
