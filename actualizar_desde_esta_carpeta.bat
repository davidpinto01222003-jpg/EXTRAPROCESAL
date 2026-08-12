@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo  ACTUALIZAR una instalacion anterior
echo ============================================================
echo.
echo  Copia los archivos NUEVOS (los de esta carpeta) encima de
echo  una instalacion vieja que tengas en otro lado, sin tocar
echo  tus credenciales ni lo que ya descargaste.
echo.
echo  Solo hace falta si ya tenias la herramienta en otra
echo  carpeta. Si esta es tu unica copia, cierra esta ventana y
echo  usa directamente descargar_correos_palabras_clave.bat
echo.
echo ============================================================
echo.

set /p DESTINO="Arrastra aqui la carpeta VIEJA y pulsa Enter: "

rem Quita las comillas que Windows agrega al arrastrar una carpeta.
set DESTINO=%DESTINO:"=%

if not exist "%DESTINO%\descargar_correos_palabras_clave.py" (
    echo.
    echo  *** En "%DESTINO%" no hay ninguna instalacion. ***
    echo  Tiene que ser la carpeta donde esta
    echo  descargar_correos_palabras_clave.py
    echo.
    pause
    exit /b 1
)

echo.
echo  Copiando a: %DESTINO%
echo.

rem Se copian TODOS los .py: el archivo principal necesita la version
rem nueva de revisar_correo_pro.py (le pide un limite de tiempo mayor),
rem y con la vieja daria error al arrancar.
copy /y "%~dp0*.py" "%DESTINO%\" >nul
copy /y "%~dp0*.bat" "%DESTINO%\" >nul
copy /y "%~dp0COMO USAR - descargar correos.txt" "%DESTINO%\" >nul

if errorlevel 1 (
    echo  *** No se pudieron copiar los archivos. ***
    echo  Cierra el programa si lo tienes abierto y reintenta.
    pause
    exit /b 1
)

rem El archivo de progreso viejo trae correos dados por revisados con
rem las reglas anteriores (las que bajaban documentos que no iban). Se
rem borra para que los vuelva a mirar con las reglas nuevas.
if exist "%DESTINO%\descargar_correos_palabras_clave_progreso.json" (
    del /q "%DESTINO%\descargar_correos_palabras_clave_progreso.json"
    echo  Se borro el progreso viejo: volvera a revisar todo con las
    echo  reglas nuevas.
)

echo.
echo ============================================================
echo  Listo. Abre la carpeta vieja y usa alli
echo  descargar_correos_palabras_clave.bat
echo.
echo  Para confirmar que quedo actualizada, la primera linea
echo  del log debe decir:  VERSION 3
echo ============================================================
echo.
pause
