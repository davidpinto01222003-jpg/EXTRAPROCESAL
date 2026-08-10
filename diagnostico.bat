@echo off
cd /d "%~dp0"

echo ============================================================
echo  DIAGNOSTICO DE LOS PORTALES
echo.
echo  Esto revisa, portal por portal, por que no estan saliendo
echo  ofertas. Vas a VER el navegador abriendose y probando cada
echo  direccion: dejalo trabajar, no lo cierres.
echo.
echo  De cada portal averigua:
echo    - si la pagina abrio o no,
echo    - si el portal pidio verificacion (captcha),
echo    - si abrio pero el programa no reconocio las ofertas.
echo.
echo  Al terminar deja un informe en:
echo     datos_empleo\diagnostico\informe.txt
echo.
echo  Ese archivo es el que hay que mandar para corregir el
echo  portal que este fallando. Se abre con el Bloc de notas.
echo.
echo  Tarda uno o dos minutos. Presiona una tecla para empezar.
echo ============================================================
pause >nul
echo.

python buscar_empleo.py --diagnostico

echo.
echo ============================================================
echo  Listo. El informe quedo en:
echo     datos_empleo\diagnostico\informe.txt
echo ============================================================
echo.
pause
