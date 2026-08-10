@echo off
cd /d "%~dp0"

echo ============================================================
echo  ARRANQUE AUTOMATICO
echo.
echo  Deja el buscador de empleo funcionando solo: cada vez que
echo  prendas el computador y entres a tu usuario, arrancan sin
echo  ventanas la vigilancia (busca cada 2 horas) y la app del
echo  celular.
echo.
echo  No tienes que abrir nada mas. Cuando encuentre algo, te
echo  llega un correo al celular con el resumen.
echo ============================================================
echo.

schtasks /Create /TN "BuscadorEmpleo" /TR "\"%~dp0empleo_automatico.bat\"" /SC ONLOGON /F

if errorlevel 1 (
  echo.
  echo  NO SE PUDO REGISTRAR.
  echo  Cierra esta ventana, busca este mismo archivo en la carpeta,
  echo  dale clic DERECHO y elige "Ejecutar como administrador".
) else (
  echo.
  echo  LISTO. Ya queda programado.
  echo.
  echo  Para que empiece AHORA sin reiniciar, dale doble clic a
  echo  "empleo_automatico.bat".
  echo.
  echo  Para quitarlo despues: "desinstalar_inicio_automatico.bat".
)

echo.
pause
