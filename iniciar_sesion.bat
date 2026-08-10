@echo off
cd /d "%~dp0"

echo ============================================================
echo  INICIAR SESION EN LOS PORTALES  (una sola vez)
echo.
echo  Esto es lo que le permite al programa POSTULARSE por ti
echo  dentro de Computrabajo, elempleo, Magneto y el Servicio
echo  Publico de Empleo, con tus propias cuentas.
echo.
echo  COMO FUNCIONA:
echo    - Se abre UN navegador con VARIAS PESTANAS, una por
echo      portal, todas de una vez.
echo    - TU escribes tu usuario y tu clave en cada pestana,
echo      como siempre. El programa NO ve ni guarda esas claves:
echo      quedan como sesion del navegador, igual que cuando
echo      entras a mano y el portal te reconoce al volver.
echo    - NO CIERRES el navegador hasta terminar en todas.
echo    - Cuando ya entraste en todas, vuelves a ESTA ventana y
echo      presionas Enter UNA sola vez.
echo    - Al final te dice en cual quedo la sesion y en cual no.
echo.
echo  Si algun portal no te abre, no importa: entra a ese a mano
echo  en una pestana nueva de ESA MISMA ventana y sigue igual.
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
echo  Esas sesiones quedan guardadas y se reusan solas.
echo  Puedes repetir esto cuando quieras: no se pierde nada.
echo ============================================================
echo.
pause
