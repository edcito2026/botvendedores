@echo off
REM ===================================================
REM 🚀 ACTUALIZAR BD Y HACER PUSH A GITHUB
REM ===================================================

setlocal enabledelayedexpansion

REM Colores para output
color 0A

echo.
echo ===================================================
echo   📊 ACTUALIZADOR DE BD + GIT PUSH
echo ===================================================
echo.

REM Cambiar a directorio del proyecto
cd /d "E:\PRUEBASEXTRACTOR\Automatizacion\WhatsApp\botvendedores"

if errorlevel 1 (
    echo ❌ Error: No se pudo acceder a la carpeta
    pause
    exit /b 1
)

echo 📁 Directorio: %cd%
echo.

REM Ejecutar script de extracción
echo 🚀 Ejecutando extractor de BD...
echo.

python ACTUALIZADOR_DB.py

if errorlevel 1 (
    echo.
    echo ❌ ERROR: El script de actualización falló
    echo.
    pause
    exit /b 1
)

echo.
echo ✅ BD actualizada exitosamente
echo.

REM Obtener fecha y hora actual
for /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set mydate=%%c-%%a-%%b)
for /f "tokens=1-2 delims=/:" %%a in ('time /t') do (set mytime=%%a-%%b)

echo 📝 Preparando commit...
echo   Timestamp: %mydate% %mytime%
echo.

REM Git commands
git add ventas.db

if errorlevel 1 (
    echo ❌ Error en: git add
    pause
    exit /b 1
)

git commit -m "📊 Update: ventas.db %mydate% %mytime%"

if errorlevel 1 (
    echo ⚠️  Sin cambios en ventas.db (ya estaba actualizado)
    echo.
    pause
    exit /b 0
)

echo 📤 Haciendo push a GitHub...
echo.

git push origin main

if errorlevel 1 (
    echo ❌ Error en: git push
    echo.
    pause
    exit /b 1
)

echo.
echo ===================================================
echo   ✅ COMPLETADO EXITOSAMENTE
echo ===================================================
echo.
echo ✓ BD actualizada
echo ✓ Cambios commiteados
echo ✓ Push a GitHub
echo.
echo ⏰ Ejecutado: %mydate% %mytime%
echo.
pause
