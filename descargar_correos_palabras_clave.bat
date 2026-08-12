@echo off
cd /d "%~dp0"

echo ============================================================
echo  Buscando en el correo los mensajes de PAGO OFICIOSO,
echo  DERECHO DE PETICION y SENTENCIA DE TUTELA, y descargando
echo  sus adjuntos a UNA SOLA carpeta, renombrados con el
echo  radicado / numero de cuenta / asunto que traigan.
echo.
echo  Se puede hacer POR PARTES: si lo cortas o se cae la
echo  conexion, la proxima vez sigue donde se quedo.
echo  (descargar_correos_palabras_clave.py)...
echo ============================================================
python descargar_correos_palabras_clave.py

pause
