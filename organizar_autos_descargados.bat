@echo off
cd /d "%~dp0"

echo ============================================================
echo  Organizando los AUTOS que ya bajaste a tu carpeta de
echo  Descargas: se renombran con el numero de proceso del Excel
echo  y se guardan en "PROCESOS TERMINADOS POR AUTO".
echo  (NO busca nada en Google Drive ni en el correo)
echo  (organizar_autos_descargados.py)...
echo ============================================================

rem Este script solo necesita estas tres librerias (no watchdog, no
rem playwright, no las de Google Drive). Si falta alguna, se instala
rem sola la primera vez y ya no vuelve a pasar.
python -c "import openpyxl, pypdf, docx" >nul 2>&1
if errorlevel 1 (
    echo.
    echo  Faltan librerias. Instalandolas una sola vez, espera un momento...
    echo.
    python -m pip install pypdf python-docx openpyxl
    echo.
)

python organizar_autos_descargados.py

pause
