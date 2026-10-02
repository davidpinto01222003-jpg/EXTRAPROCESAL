@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Filtrar solo BBVA

echo ============================================================
echo   DEJAR SOLO FNG Y LEASING A FAVOR DE BBVA
echo   (usa el Excel que ya genero EJECUTAR.bat, sin volver a leer el Drive)
echo ============================================================
echo.

set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (
    where python >nul 2>nul && set "PY=python"
)
if not defined PY (
    echo  No encontre Python. Ejecuta primero EJECUTAR.bat.
    pause
    exit /b 1
)

%PY% filtrar_bbva.py %1

echo.
pause
