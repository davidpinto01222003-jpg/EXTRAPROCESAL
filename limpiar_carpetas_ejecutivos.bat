@echo off
cd /d "%~dp0"

echo ============================================================
echo  Renombrando carpetas que quedaron solo con el numero, y
echo  borrando carpetas VACIAS de procesos activos/suspendidos/
echo  reorganizacion/remitida/etc (nunca terminados)...
echo  (limpiar_carpetas_procesos_ejecutivos.py)...
echo ============================================================
python limpiar_carpetas_procesos_ejecutivos.py

pause
