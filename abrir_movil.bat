@echo off
cd /d "%~dp0"

echo ============================================================
echo  APP DEL CELULAR
echo.
echo  Este PC queda haciendo el trabajo (buscar en los portales y
echo  mandar las hojas de vida) y tu celular queda de control
echo  remoto: ver, decidir y disparar busquedas desde la mano.
echo.
echo  Abajo va a salir la direccion que tienes que abrir en el
echo  celular. El celular debe estar en el MISMO WiFi que este PC.
echo.
echo  DEJA ESTA VENTANA ABIERTA mientras uses la app.
echo ============================================================
echo.

python app_movil.py %*

echo.
pause
