@echo off
cd /d "%~dp0"

echo ============================================================
echo  PARAR EL BUSCADOR
echo.
echo  Cierra la vigilancia y la app del celular que esten
echo  corriendo sin ventana. Lo que ya se postulo queda guardado.
echo.
echo  OJO: esto cierra TODOS los programas de Python que esten
echo  corriendo sin ventana en este equipo. Si tienes otro
echo  corriendo asi (por ejemplo el de procesos juridicos), se
echo  cierra tambien.
echo ============================================================
echo.

taskkill /IM pythonw.exe /F

echo.
echo  Para volver a arrancarlo: "empleo_automatico.bat".
echo.
pause
