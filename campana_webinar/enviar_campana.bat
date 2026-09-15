@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo   Campana del webinar - OSCAL Consultores Juridicos
echo   "Recuperacion de cartera con EPS intervenidas"
echo ============================================================
echo.
echo   1. Vista previa (no envia nada)
echo   2. Correo de prueba a una direccion mia
echo   3. ENVIO REAL
echo   4. Ver el estado de la campana
echo   5. Revisar el buzon (rebotes y retiros)
echo.
set /p opcion="  Elija una opcion (1-5): "

if "%opcion%"=="1" python enviar_campana.py vista-previa
if "%opcion%"=="2" (
    set /p destino="  A que correo envio la prueba: "
    call python enviar_campana.py prueba --para %%destino%%
)
if "%opcion%"=="3" python enviar_campana.py enviar
if "%opcion%"=="4" python enviar_campana.py estado
if "%opcion%"=="5" python enviar_campana.py revisar-buzon

echo.
pause
