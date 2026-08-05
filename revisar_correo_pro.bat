@echo off
cd /d "%~dp0"

echo ============================================================
echo  Revisando el correo en MODO PRO y repartiendo cada correo
echo  a la carpeta de su proceso (demandado / radicado / cuenta).
echo  Lo que no coincida va a "_SIN CLASIFICAR - REVISAR A MANO".
echo.
echo  Se puede hacer POR PARTES: si lo cortas o se cae la
echo  conexion, la proxima vez sigue donde se quedo.
echo  (revisar_correo_pro.py)...
echo ============================================================
python revisar_correo_pro.py

pause
