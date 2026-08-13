@echo off
cd /d "%~dp0"

echo ============================================================
echo  Organizando los AUTOS que ya bajaste a tu carpeta de
echo  Descargas: se renombran con el numero de proceso del Excel
echo  y se guardan en "PROCESOS TERMINADOS POR AUTO".
echo  (NO busca nada en Google Drive ni en el correo)
echo  (organizar_autos_descargados.py)...
echo ============================================================

rem Este script solo necesita estas librerias (no watchdog, no
rem playwright, no las de Google Drive). Si falta alguna, se instala
rem sola la primera vez y ya no vuelve a pasar. "cryptography" es la
rem que permite abrir los PDF cifrados con AES -- sin ella, esos autos
rem dan "cryptography>=3.1 is required for AES algorithm" y se pierden.
python -c "import openpyxl, pypdf, docx, cryptography" >nul 2>&1
if errorlevel 1 (
    echo.
    echo  Faltan librerias. Instalandolas una sola vez, espera un momento...
    echo.
    python -m pip install pypdf python-docx openpyxl cryptography
    echo.
)

python organizar_autos_descargados.py

pause
