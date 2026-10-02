@echo off
cd /d "%~dp0"

echo ============================================================
echo  Depurando procesos Supersociedades: garantias FNG y LEASING
echo  (depurar_fng_leasing.py)...
echo ============================================================
python depurar_fng_leasing.py %1

pause
