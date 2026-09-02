@echo off
cd /d "%~dp0"

echo ============================================================
echo  Buscando POR DONDE radicarle la propuesta a cada IPS:
echo  pagina oficial, correo de contratacion, area juridica
echo  y convocatorias (SECOP por NIT y por nombre + su web).
echo.
echo  Requiere haber corrido antes ips_area_metropolitana.bat
echo  Resultado: datos_ips\IPS_Area_Metropolitana_CON_CANAL.xlsx
echo  (ips_enriquecer_contactos.py)
echo ============================================================
python ips_enriquecer_contactos.py %*

pause
