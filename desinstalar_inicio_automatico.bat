@echo off
cd /d "%~dp0"

echo ============================================================
echo  QUITAR EL ARRANQUE AUTOMATICO
echo.
echo  Deja de arrancar solo al prender el computador. NO borra
echo  nada de lo que ya hiciste: tus postulaciones, el Excel y
echo  tu perfil quedan intactos.
echo ============================================================
echo.

schtasks /Delete /TN "BuscadorEmpleo" /F

if errorlevel 1 (
  echo.
  echo  No estaba registrado, o hay que hacerlo como administrador
  echo  (clic derecho sobre este archivo, "Ejecutar como administrador").
) else (
  echo.
  echo  Listo, ya no arranca solo.
  echo  Si ademas quieres pararlo AHORA: "parar_empleo.bat".
)

echo.
pause
