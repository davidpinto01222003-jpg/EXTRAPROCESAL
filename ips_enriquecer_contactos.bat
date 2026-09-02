@echo off
cd /d "%~dp0"

echo ============================================================
echo  Buscando POR DONDE radicarle la propuesta a cada IPS:
echo  correo de contratacion, area juridica, pagina de
echo  proveedores y convocatorias juridicas en SECOP.
echo.
echo  Requiere haber corrido antes ips_area_metropolitana.bat
echo  Resultado: datos_ips\IPS_Area_Metropolitana_CON_CANAL.xlsx
echo  (ips_enriquecer_contactos.py)
echo ============================================================
python ips_enriquecer_contactos.py %*

pause
