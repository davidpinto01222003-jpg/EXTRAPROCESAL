@echo off
cd /d "%~dp0"

echo ============================================================
echo  Buscador de empleo: una pasada por todos los portales.
echo.
echo  Busca en Computrabajo, elempleo, Magneto y el Servicio
echo  Publico de Empleo, califica cada vacante contra tu perfil
echo  (perfil_laboral.json) y postula a las que pasan el umbral.
echo.
echo  Al terminar deja el resultado en:
echo     datos_empleo\postulaciones.xlsx
echo ============================================================
echo.

python buscar_empleo.py %*

echo.
pause
