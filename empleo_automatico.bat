@echo off
cd /d "%~dp0"

REM Arranca las dos piezas SIN VENTANAS (pythonw), para que el PC trabaje
REM solo y no te estorbe:
REM
REM   1. La vigilancia: revisa los portales cada 2 horas y postula.
REM   2. La app del celular: para que puedas ver y mandar desde el telefono.
REM
REM Este es el archivo que registra `instalar_inicio_automatico.bat` para
REM que corra solo al prender el computador. Tambien puedes darle doble
REM clic tu mismo cuando quieras.
REM
REM Como no abren ventana, para saber que estan corriendo mira el
REM Administrador de tareas (busca "pythonw.exe"), o simplemente abre la
REM app en el celular: si conecta, esta corriendo.
REM
REM Para PARARLAS: `parar_empleo.bat`.

start "" pythonw.exe app_movil.py
start "" pythonw.exe buscar_empleo.py --vigilar
