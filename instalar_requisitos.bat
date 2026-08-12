@echo off
cd /d "%~dp0"

echo ============================================================
echo  Instalando lo que necesitan las herramientas de este
echo  proyecto. Esto se hace UNA SOLA VEZ.
echo.
echo  Necesita internet y que Python ya este instalado
echo  (con la casilla "Add Python to PATH" marcada).
echo ============================================================
echo.

python --version
if errorlevel 1 (
    echo.
    echo  *** NO se encontro Python. ***
    echo  Instalalo desde https://www.python.org/downloads/
    echo  y MARCA la casilla "Add Python to PATH" al instalarlo.
    echo  Despues vuelve a ejecutar este archivo.
    echo.
    pause
    exit /b 1
)

echo.
echo  Instalando (puede tardar unos minutos)...
echo.
python -m pip install --upgrade pip
python -m pip install watchdog python-docx openpyxl pypdf

if errorlevel 1 (
    echo.
    echo  *** Algo fallo al instalar. ***
    echo  Revisa que tengas internet y vuelve a intentarlo.
    echo.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Listo. Ya puedes usar descargar_correos_palabras_clave.bat
echo.
echo  ANTES de la primera vez: crea el archivo
echo  credenciales_sgde.txt con tu correo y tu contrasena de
echo  aplicacion de Gmail (ver credenciales_sgde.example.txt).
echo ============================================================
echo.
pause
