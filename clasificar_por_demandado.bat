@echo off
cd /d "%~dp0"

echo ============================================================
echo  Clasificando informacion extraprocesal por el nombre del
echo  demandado (carpeta de descargas + Gmail)...
echo  (clasificar_por_demandado.py)...
echo ============================================================
python clasificar_por_demandado.py

pause
