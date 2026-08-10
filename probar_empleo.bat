@echo off
cd /d "%~dp0"

echo ============================================================
echo  MODO PRUEBA (simulacion): busca y califica las vacantes,
echo  te muestra a cuales postularia y por que, pero NO manda
echo  absolutamente nada.
echo.
echo  Corre esto PRIMERO, y varias veces, hasta que la lista de
echo  vacantes que muestra sea la que tu mismo escogerias. Ahi
echo  ya puedes usar buscar_empleo.bat de verdad.
echo ============================================================
echo.

python buscar_empleo.py --simular

echo.
pause
