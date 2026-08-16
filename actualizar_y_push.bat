@echo off
REM ===================================================
REM 🚀 ACTUALIZAR BD Y HACER PUSH A GITHUB
REM Versión: V3 - Con reportes mejorados
REM ===================================================

setlocal enabledelayedexpansion

REM Log file
set LOGFILE=E:\PRUEBASEXTRACTOR\Automatizacion\WhatsApp\botvendedores\logs\actualizar_y_push.log

REM Cambiar a directorio del proyecto
cd /d "E:\PRUEBASEXTRACTOR\Automatizacion\WhatsApp\botvendedores"

if errorlevel 1 (
    echo [%date% %time%] ERROR: No se pudo acceder a la carpeta >> "%LOGFILE%"
    exit /b 1
)

echo [%date% %time%] ===== INICIANDO ACTUALIZACIÓN ===== >> "%LOGFILE%"

REM Ejecutar script de extracción
echo [%date% %time%] Ejecutando ACTUALIZADOR_DB_OPTIMIZADO.py >> "%LOGFILE%"
python ACTUALIZADOR_DB_OPTIMIZADO.py >> "%LOGFILE%" 2>&1

if errorlevel 1 (
    echo [%date% %time%] ERROR: Script de actualización falló >> "%LOGFILE%"
    exit /b 1
)

echo [%date% %time%] BD actualizada exitosamente >> "%LOGFILE%"

REM Git commands
echo [%date% %time%] Git add archivos >> "%LOGFILE%"
git add WEBHOOK_VENDEDORES_MEJORADO.py >> "%LOGFILE%" 2>&1
git add ACTUALIZADOR_DB_OPTIMIZADO.py >> "%LOGFILE%" 2>&1
git add prueba_vendedores.py >> "%LOGFILE%" 2>&1
git add ventas.db >> "%LOGFILE%" 2>&1

if errorlevel 1 (
    echo [%date% %time%] ERROR: git add falló >> "%LOGFILE%"
    exit /b 1
)

REM Obtener fecha y hora para commit
for /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set mydate=%%c-%%a-%%b)
for /f "tokens=1-2 delims=/:" %%a in ('time /t') do (set mytime=%%a-%%b)

echo [%date% %time%] Ejecutando git commit >> "%LOGFILE%"
git commit -m "📊 Auto: Reportes mejorados actualizado %mydate% %mytime%" >> "%LOGFILE%" 2>&1

if errorlevel 1 (
    echo [%date% %time%] INFO: Sin cambios en BD (ya estaba actualizado) >> "%LOGFILE%"
    echo [%date% %time%] ===== COMPLETADO (sin cambios) ===== >> "%LOGFILE%"
    exit /b 0
)

echo [%date% %time%] Ejecutando git push >> "%LOGFILE%"
git push origin main >> "%LOGFILE%" 2>&1

if errorlevel 1 (
    echo [%date% %time%] ERROR: git push falló >> "%LOGFILE%"
    exit /b 1
)

echo [%date% %time%] ✅ COMPLETADO EXITOSAMENTE >> "%LOGFILE%"
echo [%date% %time%] BD actualizada, commiteada y pusheada >> "%LOGFILE%"
echo [%date% %time%] ===== FIN ===== >> "%LOGFILE%"
exit /b 0
