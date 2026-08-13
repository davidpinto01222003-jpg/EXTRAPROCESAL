@echo off
cd /d "%~dp0"

echo ============================================================
echo  Organizando lo que ya bajaste a tu carpeta de Descargas y
echo  que TERMINA un proceso: se renombra con el numero de
echo  proceso del Excel y se guarda en "PROCESOS TERMINADOS POR
echo  AUTO" / "POR PAGO" / etc, segun el estado.
echo  (NO busca nada en Google Drive ni en el correo)
echo  (organizar_autos_descargados.py)...
echo ============================================================
echo.

rem Sin bloques de varias lineas entre parentesis -- if ... ( ) -- ni
rem redirecciones: cmd.exe se cierra en silencio con esos bloques si el
rem .bat tiene saltos de linea de Unix, y el script no arrancaba. Si
rem falta alguna libreria, el propio .py la instala y reintenta.
python organizar_autos_descargados.py

echo.
pause
