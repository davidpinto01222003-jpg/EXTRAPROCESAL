@echo off
cd /d "%~dp0"

echo ============================================================
echo  Buscando convocatorias juridicas nuevas en el SECOP II...
echo  (buscar_convocatorias_secop.py)
echo ============================================================
python buscar_convocatorias_secop.py %*

pause
