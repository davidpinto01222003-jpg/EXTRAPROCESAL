@echo off
cd /d "%~dp0"

echo ============================================================
echo  Clasificando procesos ejecutivos (informacion no procesal +
echo  auto de terminacion) segun el Excel de control...
echo  (clasificar_procesos_ejecutivos.py)...
echo ============================================================
python clasificar_procesos_ejecutivos.py

pause
