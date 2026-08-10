@echo off
cd /d "%~dp0"

echo ============================================================
echo  INSTALACION DEL BUSCADOR DE EMPLEO
echo.
echo  Esto baja de internet lo que el programa necesita para
echo  funcionar. Se hace UNA SOLA VEZ y puede tardar varios
echo  minutos (lo mas pesado es el navegador, unos 150 MB).
echo.
echo  Necesitas Python instalado. Si no lo tienes, bajalo de
echo  https://www.python.org/downloads/ y MUY IMPORTANTE:
echo  marca la casilla "Add Python to PATH" durante la
echo  instalacion.
echo ============================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
  echo  NO ENCUENTRO PYTHON EN ESTE COMPUTADOR.
  echo.
  echo  Instalalo desde https://www.python.org/downloads/
  echo  marcando "Add Python to PATH", cierra esta ventana y
  echo  vuelve a darle doble clic a este archivo.
  echo.
  pause
  exit /b 1
)

echo  [1 de 2] Instalando las librerias de Python...
echo.
python -m pip install --upgrade pip
python -m pip install -r requirements_empleo.txt
if errorlevel 1 goto :fallo

echo.
echo  [2 de 2] Bajando el navegador que lee los portales...
echo.
python -m playwright install chromium
if errorlevel 1 goto :fallo

echo.
echo ============================================================
echo  LISTO. Ya quedo instalado.
echo.
echo  AHORA SIGUE ESTO, EN ESTE ORDEN:
echo.
echo   1. Copia "perfil_laboral.example.json", renombralo a
echo      "perfil_laboral.json" y llenalo con tus datos.
echo.
echo   2. Copia "credenciales_empleo.example.txt", renombralo a
echo      "credenciales_empleo.txt" y pon tu correo con una
echo      contrasena de aplicacion de Gmail.
echo.
echo   3. Abre una terminal aqui y corre:
echo         python buscar_empleo.py --login
echo      para entrar TU MISMO a Computrabajo y elempleo.
echo.
echo   4. Doble clic en "probar_empleo.bat" (modo prueba: no
echo      manda nada) hasta que el filtro quede a tu gusto.
echo.
echo   5. Doble clic en "abrir_movil.bat" para instalar la app
echo      en el celular escaneando el codigo QR.
echo.
echo  Todo esto esta explicado con detalle en LEEME.md.
echo ============================================================
echo.
pause
exit /b 0

:fallo
echo.
echo ============================================================
echo  ALGO FALLO EN LA INSTALACION.
echo.
echo  Lo mas comun es que no haya internet, o que el antivirus
echo  este bloqueando la descarga. Revisa el mensaje rojo de
echo  arriba, corrige y vuelve a darle doble clic a este archivo.
echo ============================================================
echo.
pause
exit /b 1
