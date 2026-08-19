@echo off
REM Script para detectar y actualizar SOLO archivos modificados en GitHub
REM Git automáticamente detecta qué cambió

cd /d E:\PRUEBASEXTRACTOR\Automatizacion\WhatsApp\botvendedores

echo ========================================
echo Detectando archivos modificados...
echo ========================================

REM Mostrar qué archivos cambió
git status --short

echo.
echo ========================================
echo Subiendo cambios a GitHub...
echo ========================================

REM Agregar SOLO los tipos de archivo que nos interesan (git detecta automáticamente cuáles modificaron)
git add *.py *.db *.xlsx Procfile requirements.txt 2>nul

REM Intentar commit (solo si hay cambios)
git commit -m "Auto update - archivos modificados" 2>nul

if %ERRORLEVEL% EQU 0 (
    REM Si el commit fue exitoso, hacer push
    git push
    echo.
    echo ========================================
    echo ✅ Cambios enviados a GitHub
    echo ========================================
) else (
    echo.
    echo ========================================
    echo ⚠️ No hay cambios para subir
    echo ========================================
)

pause
