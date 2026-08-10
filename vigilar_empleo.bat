@echo off
cd /d "%~dp0"

echo ============================================================
echo  Vigilancia de empleo: revisa los portales cada 2 horas y
echo  postula solo a lo que encaja con tu perfil, de forma
echo  indefinida.
echo.
echo  Cierra esta ventana o Ctrl+C cuando quieras detenerlo.
echo  Lo que ya se postulo queda guardado; al volver a arrancar
echo  no se repite.
echo ============================================================
echo.

python buscar_empleo.py --vigilar

echo.
pause
