@echo off
cd /d "%~dp0"

echo ============================================================
echo  Generando Word para marcar carpetas fisicas...
echo  (generar_carpetas_acuerdos_pago.py)
echo  Excel esperado: ACUERDOS_DE_PAGO.xlsx
echo ============================================================
python generar_carpetas_acuerdos_pago.py

pause
