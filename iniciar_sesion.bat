@echo off
cd /d "%~dp0"

echo ============================================================
echo  INICIAR SESION EN LOS PORTALES  (una sola vez)
echo.
echo  Esto es lo que le permite al programa POSTULARSE por ti
echo  dentro de Computrabajo y elempleo, con tu propia cuenta.
echo.
echo  COMO FUNCIONA:
echo    - Se abre un navegador normal, uno por portal.
echo    - TU escribes tu usuario y tu clave, como siempre.
echo    - El programa NO ve ni guarda esas claves: quedan
echo      guardadas como sesion del navegador, igual que cuando
echo      entras a mano y el portal te reconoce al volver.
echo    - Cuando termines en un portal, vuelves a ESTA ventana
echo      y presionas Enter para pasar al siguiente.
echo.
echo  Si NO haces esto, el buscador funciona igual, pero solo
echo  puede postularse por CORREO (a las ofertas que publican
echo  una direccion de correo). Las demas te las deja marcadas
echo  como "para ti", con el link, para que las termines tu.
echo.
echo  Presiona una tecla para abrir el navegador.
echo ============================================================
pause >nul
echo.

python buscar_empleo.py --login

echo.
echo ============================================================
echo  Listo. Esas sesiones quedan guardadas y se reusan solas.
echo  Solo hay que repetir esto si algun portal te cierra la
echo  sesion por inactividad (cada varios meses).
echo ============================================================
echo.
pause
