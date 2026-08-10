@echo off
cd /d "%~dp0"

echo ============================================================
echo  PRUEBA DE POSTULACION  (una sola vacante, con tu permiso)
echo.
echo  Esta es LA prueba que importa: responde si el programa de
echo  verdad puede postularse por ti, o no.
echo.
echo  QUE HACE:
echo    1. Busca en los portales una vacante que pase tu filtro.
echo    2. La abre a la vista y te dice que encontro:
echo         - si la oferta publica un correo de contacto,
echo         - si encontro el boton de "Postularme",
echo         - o si no hay por donde.
echo    3. TE PREGUNTA si quieres que se postule de verdad.
echo       Si dices que no, no se manda absolutamente nada.
echo.
echo  Si no encuentra el boton, te dice que nombres esta
echo  buscando: mira la pagina abierta y fijate como se llama
echo  ese boton en el portal. Con ese dato se corrige.
echo.
echo  Presiona una tecla para empezar.
echo ============================================================
pause >nul
echo.

python buscar_empleo.py --probar-postulacion --ver

echo.
pause
