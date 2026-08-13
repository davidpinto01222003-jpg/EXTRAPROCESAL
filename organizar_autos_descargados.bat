@echo off
cd /d "%~dp0"

echo ============================================================
echo  Organizando los AUTOS que ya bajaste a tu carpeta de
echo  Descargas: se renombran con el numero de proceso del Excel
echo  y se guardan en "PROCESOS TERMINADOS POR AUTO".
echo  (NO busca nada en Google Drive ni en el correo)
echo  (organizar_autos_descargados.py)...
echo ============================================================
python organizar_autos_descargados.py

pause
