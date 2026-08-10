@echo off
cd /d "%~dp0"

echo ============================================================
echo  MODO PRUEBA (simulacion)
echo.
echo  Busca y califica las vacantes de verdad, te muestra a
echo  cuales postularia y por que, pero NO manda absolutamente
echo  nada.
echo.
echo  Vas a VER el navegador abriendo Computrabajo, elempleo y
echo  los demas portales, uno por uno. Dejalo trabajar solo: no
echo  cierres esa ventana ni le muevas el mouse encima.
echo.
echo  Corre esto PRIMERO, y varias veces, hasta que la lista de
echo  vacantes que muestra sea la que tu mismo escogerias. Ahi
echo  ya puedes usar buscar_empleo.bat de verdad.
echo ============================================================
echo.

python buscar_empleo.py --simular --ver

echo.
pause
