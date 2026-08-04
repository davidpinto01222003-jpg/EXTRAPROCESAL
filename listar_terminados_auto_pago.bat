@echo off
cd /d "%~dp0"

echo ============================================================
echo  Generando la lista de procesos terminados por auto o por
echo  pago (partes y radicado)...
echo  (listar_terminados_auto_pago.py)...
echo ============================================================
python listar_terminados_auto_pago.py

pause
