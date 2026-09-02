@echo off
cd /d "%~dp0"

echo ============================================================
echo  Base de datos de IPS (sociedades) del Area Metropolitana
echo  de Bucaramanga -- Bucaramanga, Floridablanca, Giron
echo  y Piedecuesta, desde el REPS del Ministerio de Salud.
echo.
echo  Resultado en la carpeta datos_ips\
echo  (ips_area_metropolitana.py)
echo ============================================================
python ips_area_metropolitana.py %*

pause
