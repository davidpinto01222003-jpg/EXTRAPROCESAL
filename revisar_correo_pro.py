"""
Revisa tu Gmail en MODO PRO y reparte cada correo a la carpeta del
proceso que le corresponde (por DEMANDADO, RADICADO o NUMERO DE
CUENTA, comparando contra el Excel de control). Lo que no coincide con
ningun proceso NO se pierde: se guarda en una carpeta aparte para
revisarlo a mano.

POR QUE ESTE SCRIPT EXISTE (y en que se diferencia de
clasificar_por_demandado.py)
============================================================
El paso 2 de clasificar_por_demandado.py le PREGUNTABA a Gmail una vez
por cada termino del Excel: "¿tienes correos que digan DAVID CAICEDO?",
"¿y 68001400302120180078701?", "¿y 2018-00787?"... con ~650 procesos
activos eso son MAS DE 4000 busquedas seguidas contra el servidor. Da
igual lo bien escrita que este cada busqueda: ningun servidor de correo
(y menos Gmail, que corta por proteccion anti-scripts) aguanta miles de
busquedas seguidas de un mismo cliente. De ahi venian, todos, los
sintomas que se vieron en la practica: "Could not parse command", "Too
many protocol errors", "socket error: EOF", y el programa quedandose
colgado sin responder.

Este script INVIERTE el planteamiento, que es lo unico que lo arregla
de raiz:

    ANTES:  4242 preguntas a Gmail  ->  Gmail corta  ->  nunca termina
    AHORA:  se BAJAN los correos UNA vez  ->  se comparan en TU PC

Es decir: NO se le pregunta nada a Gmail sobre demandados, radicados ni
cuentas. Se baja el contenido de los correos de una ventana de fechas
(ver DIAS_HACIA_ATRAS) y TODA la comparacion contra el Excel se hace
localmente, en tu computador, sin red de por medio. Comparar 4242
terminos contra un texto es instantaneo cuando se hace en memoria; lo
que era imposible era preguntarlo 4242 veces por internet.

MODO PRO: DONDE BUSCA LOS DATOS DENTRO DE CADA CORREO
============================================================
No se conforma con el asunto. De cada correo revisa:
  1. el ASUNTO,
  2. el CUERPO del mensaje,
  3. el NOMBRE de cada archivo adjunto,
  4. y -- esto es lo "pro" -- el TEXTO DE ADENTRO de cada adjunto
     PDF/DOCX (ver LEER_TEXTO_DE_ADJUNTOS). Un auto casi nunca trae el
     radicado en el asunto del correo: lo trae impreso dentro del PDF.
     Sin leer el PDF, ese correo jamas se podria emparejar.

Ademas, cada adjunto se compara POR SEPARADO (no todo pegado en un solo
texto). Asi el nombre del demandado se busca en el ENCABEZADO de cada
documento -- que es donde de verdad esta ("DEMANDADO: ANA LUCIA GUECHA
ARENAS") -- y no se diluye entre el texto de los otros adjuntos.

A QUE PROCESO VA CADA CORREO
============================================================
Se usa la MISMA regla de emparejamiento que el resto del proyecto (ver
clasificar_procesos_ejecutivos._procesos_que_coinciden_con_correo):
coincide su RADICADO, su CUENTA, O el DEMANDADO completo -- basta
cualquiera de los tres. El radicado se reconoce plano (23 digitos),
con guiones/puntos/espacios, y en sus formas cortas (2018-00787,
2018-787, 2018-00787-00). El demandado exige TODAS sus palabras
significativas, como palabra completa, en el encabezado del documento.

Si un correo coincide con VARIOS procesos (ej. una cuenta que aparece
en dos procesos del mismo cliente), se guarda en la carpeta de TODOS
ellos -- es informacion que aplica a todos. Si no coincide con ninguno,
va a CARPETA_SIN_CLASIFICAR para que la revises a mano; nunca se
descarta ni se adivina.

POR QUE NO SE TRABA (todos los bugs previstos)
============================================================
Cada uno de estos puntos corresponde a un fallo real que se vio en la
practica con el enfoque anterior:

* SE PUEDE HACER POR PARTES, y retoma solo. Lleva un archivo de control
  (ARCHIVO_PROGRESO) con los correos ya revisados. Si lo cortas, se cae
  la luz, o Gmail te corta la conexion, la proxima corrida SIGUE DONDE
  SE QUEDO en vez de empezar de cero.
* SIEMPRE TERMINA. Por corrida revisa como maximo MAX_CORREOS_POR_CORRIDA
  correos y se detiene ordenadamente. Nunca se queda "toda la tarde"
  colgado: si falta mas, lo dice en el log y basta volver a correrlo.
* NADA SE QUEDA ESPERANDO PARA SIEMPRE. Toda operacion de red tiene
  DOBLE limite de tiempo: el del socket, y uno "duro" (un hilo aparte
  que corta a los TIMEOUT_SEGUNDOS pase lo que pase). Esto ultimo hace
  falta porque en equipos con antivirus que inspecciona el correo
  (Avast/Kaspersky/ESET/McAfee) el limite normal del socket NO se
  dispara y el programa se queda muerto sin dar ni un error.
* SI SE CAE LA CONEXION, RECONECTA Y SIGUE. Hasta MAX_RECONEXIONES
  veces, retomando exactamente en el correo donde iba (por eso se usan
  UIDs, ver el punto siguiente).
* USA UID, NO NUMERO DE ORDEN. El "numero" de un correo cambia si
  llegan o se borran mensajes mientras el script corre -- reconectar
  con numeros de orden te hace saltar o repetir correos en silencio.
  El UID es fijo. Ademas se verifica el UIDVALIDITY del buzon: si Gmail
  lo cambia (los UIDs viejos dejan de ser validos), el archivo de
  progreso se descarta solo, avisando, en vez de dar por revisados
  correos equivocados.
* UN CORREO DAÑADO NO TUMBA LA CORRIDA. Cada correo se procesa aislado:
  si uno falla (MIME roto, PDF corrupto, adjunto gigante, nombre de
  archivo imposible), se registra y se sigue con el siguiente.
* NO SE PIERDE LO YA HECHO. El progreso se guarda cada
  GUARDAR_PROGRESO_CADA_N correos, no solo al final.
* NO PISA ARCHIVOS NI REVIENTA CON RUTAS LARGAS DE WINDOWS. Reutiliza
  el guardado seguro del resto del proyecto (_ruta_archivo_libre +
  _ruta_larga_segura).
* SE VE QUE ESTA VIVO. Deja un aviso de avance cada
  AVISO_PROGRESO_CADA_N correos, aunque no encuentre nada -- un log en
  silencio se ve igual que un programa colgado.

Respeta MODO_PRUEBA (por defecto True): primero muestra en el log que
haria, SIN mover ni descargar nada. Revisa el log y, cuando se vea
bien, cambia MODO_PRUEBA = False.

Necesita credenciales_sgde.txt (las mismas de siempre, contraseña de
aplicacion de Gmail -- ver README). Si no existe, avisa y no hace nada.
"""

import datetime
import email
import imaplib
import io
import json
import logging
import os
import re
import threading
import time
from pathlib import Path

import clasificar_procesos_ejecutivos as base
import buscar_faltantes_en_drive as buscador
import validar_renombrar_carpetas as cruce_excel
import procesos_juridicos as organizador

# ============================= CONFIGURACION =============================

# True (por defecto): NO descarga ni guarda nada -- solo muestra en el
# log a que carpeta iria cada correo. Revisa el log y, cuando se vea
# bien, cambia esto a False para que lo haga de verdad.
MODO_PRUEBA = True

# Cuantos dias hacia atras revisar. 730 = los ultimos 2 años. No revisa
# correos mas viejos que esto.
#
# Una ventana grande NO hace que la corrida sea eterna: el script va
# POR PARTES (ver MAX_CORREOS_POR_CORRIDA y MAX_MINUTOS_POR_CORRIDA) y
# recuerda lo ya revisado, asi que la puesta al dia de 2 años se hace
# en varias corridas y cada una termina en un rato acotado. Una vez al
# dia, cuando ya no queden pendientes, cada corrida solo revisa lo
# nuevo (unos pocos correos) y termina en segundos.
DIAS_HACIA_ATRAS = 730

# Maximo de correos a revisar POR CORRIDA. Es lo que garantiza que el
# script SIEMPRE termine en un rato razonable: si quedan mas, lo dice
# en el log y basta volver a correrlo (retoma solo donde se quedo, ver
# ARCHIVO_PROGRESO). Subelo si quieres avanzar mas por corrida.
MAX_CORREOS_POR_CORRIDA = 500

# Ademas del tope por cantidad, un tope por TIEMPO: la corrida se
# cierra ordenadamente al pasar estos minutos, aunque no haya llegado
# a MAX_CORREOS_POR_CORRIDA. Hace falta porque los correos son MUY
# desiguales: 500 correos de texto se revisan en minutos, pero 500 con
# adjuntos escaneados pesados pueden tardar horas. Asi cada corrida
# dura lo que tu decidas, y el resto queda para la siguiente (nunca se
# pierde nada, ver ARCHIVO_PROGRESO). Ponlo en 0 para no limitar por
# tiempo.
MAX_MINUTOS_POR_CORRIDA = 30

# True (por defecto): empieza por los correos MAS NUEVOS y va hacia
# atras. Con una ventana grande (ej. 730 dias = 2 años) esto importa:
# si son miles de correos y necesitas varias corridas, conviene que lo
# primero que quede clasificado sea lo mas reciente -- que casi siempre
# es lo mas urgente. Ponlo en False para ir del mas viejo al mas nuevo.
EMPEZAR_POR_LOS_MAS_NUEVOS = True

# Cuantos correos se bajan por cada peticion a Gmail. Bajarlos de a
# varios es mas rapido que de a uno, pero pedir demasiados de golpe
# hace que la peticion tarde MAS del limite de tiempo y la conexion se
# caiga: en un caso real, pedir 20 correos con adjuntos escaneados
# tardaba mas de 60s y Gmail cortaba. 5 es un punto medio seguro.
TAMANO_LOTE_DESCARGA = 5

# True (por defecto): en vez de bajar TODOS los correos de la ventana
# de fechas, se le pide a Gmail que devuelva solo los que ya mencionan
# alguna de las FRASES_EXTRAPROCESALES (en el asunto o en el texto).
#
# Sin esto, dos años de correo eran 25.134 mensajes en un caso real --
# bajarlos todos completos satura la conexion y tomaria decenas de
# corridas. Filtrando de entrada quedan unos cientos: los que de verdad
# interesan. Son ~8 consultas a Gmail en total, no una por termino del
# Excel (eso era lo que fallaba antes).
#
# Solo aplica si SOLO_INFORMACION_EXTRAPROCESAL esta en True (si
# quieres TODO, no hay por que acotar). Ponlo en False para revisar
# todos los correos de la ventana, sabiendo que sera mucho mas lento.
ACOTAR_POR_TIPO_EN_GMAIL = True

# Frases que se le piden a Gmail para acotar (ver ACOTAR_POR_TIPO_EN_GMAIL).
# A proposito son TROZOS de palabra, no la palabra completa, para que
# una sola sirva para todas sus variantes y no dependa de las tildes
# (la busqueda de IMAP es por subcadena):
#   "PETICI"   -> PETICION, PETICIÓN, PETICIONES, "respuesta a su peticion"...
#   "OFICIOSO" -> PAGO OFICIOSO, PAGOS OFICIOSOS, "pago de manera oficiosa"
#   "TUTELA"   -> TUTELA, ACCION DE TUTELA
#   "PQR"      -> PQR, PQRS, PQRSD (como se rotulan las respuestas)
FRASES_EXTRAPROCESALES = ["PETICI", "TUTELA", "OFICIOSO", "PQR"]

# Limite de tiempo (segundos) para CUALQUIER operacion contra Gmail.
# Se aplica por DUPLICADO: al socket, y con un limite "duro" aparte
# (ver _con_limite_de_tiempo) porque el del socket no siempre se
# dispara en equipos con antivirus que inspecciona el correo.
TIMEOUT_SEGUNDOS = 60

# Cuantas veces reintentar reconectar si Gmail corta la conexion.
MAX_RECONEXIONES = 10

# Cuantas veces reintentar UN MISMO lote antes de saltarlo. Si un lote
# falla siempre (incluso con conexion nueva), el problema es de ESE
# lote: se salta para no gastar ahi todos los reintentos.
MAX_REINTENTOS_POR_LOTE = 3

# True (modo pro): ademas del asunto y el cuerpo, LEE EL TEXTO de los
# adjuntos PDF/DOCX para buscar ahi el radicado/cuenta/demandado. Es lo
# que permite emparejar un auto cuyo radicado solo esta impreso dentro
# del PDF. Ponlo en False si quieres que vaya mas rapido y te conformas
# con el asunto/cuerpo/nombre del adjunto.
LEER_TEXTO_DE_ADJUNTOS = True

# Adjuntos mas pesados que esto no se abren para leerles el texto (un
# escaneo grande puede tardar muchisimo). Igual se guardan completos:
# esto solo evita leerlos para emparejar.
MAX_MB_ADJUNTO_PARA_LEER = 25

# Donde se crean/buscan las carpetas de cada proceso. Por defecto la
# misma que usa el resto del proyecto.
CARPETA_PROCESOS = base.CARPETA_PROCESOS

# Carpeta (dentro de CARPETA_PROCESOS) donde van los correos que NO
# coincidieron con ningun proceso, para revisarlos a mano.
CARPETA_SIN_CLASIFICAR = "_SIN CLASIFICAR - REVISAR A MANO"

# True (por defecto): guarda SOLO informacion EXTRAPROCESAL, que es lo
# que se necesita:
#     derecho de peticion / peticion / PQR-PQRS-PQRSD,
#     LAS RESPUESTAS a esos derechos de peticion,
#     tutela / accion de tutela,
#     pago oficioso.
# Todo lo demas (demandas, memoriales, mandamientos, sentencias,
# embargos... ver PALABRAS_PROCESAL_EXCLUIR) se omite aunque coincida
# con un proceso. Ver es_extraprocesal aqui abajo y
# clasificar_procesos_ejecutivos.es_informacion_no_procesal.
#
# Ponlo en False si alguna vez quieres bajar TODO lo que coincida con
# un proceso, sin filtrar por tipo.
SOLO_INFORMACION_EXTRAPROCESAL = True

# False (por defecto): no exige que el correo mencione a ESSA. Ponlo en
# True si tu Gmail mezcla correos de otros clientes y quieres filtrar.
EXIGIR_MENCION_ESSA = False

# Archivo donde se guarda que correos ya se revisaron (para poder
# hacerlo por partes y retomar donde se quedo).
ARCHIVO_PROGRESO = os.path.join(os.path.dirname(__file__), "revisar_correo_pro_progreso.json")

# Reportes en CSV de lo que se hizo y de lo que quedo sin emparejar.
CARPETA_REPORTES = os.path.dirname(base.RUTA_EXCEL_CONTROL) or os.path.dirname(__file__)
ARCHIVO_REPORTE = os.path.join(CARPETA_REPORTES, "revisar_correo_pro_clasificados.csv")
ARCHIVO_SIN_COINCIDENCIA = os.path.join(CARPETA_REPORTES, "revisar_correo_pro_sin_coincidencia.csv")

ARCHIVO_LOG = os.path.join(os.path.dirname(__file__), "revisar_correo_pro.log")

# Cada cuantos correos se guarda el progreso en disco (para no perder
# lo hecho si el programa se corta a la mitad).
GUARDAR_PROGRESO_CADA_N = 25

# Cada cuantos correos se deja un aviso de avance en el log.
AVISO_PROGRESO_CADA_N = 25

# ========================================================================


def configurar_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(message)s",
        datefmt="%H:%M:%S",
        handlers=[
            logging.FileHandler(ARCHIVO_LOG, encoding="utf-8", mode="w"),
            logging.StreamHandler(),
        ],
    )


# ==================== Progreso (para hacerlo por partes) ====================


def cargar_progreso(uidvalidity_actual, archivo=None):
    """
    Lee que correos ya se revisaron en corridas anteriores. Devuelve un
    set de UIDs (numeros).

    Si el UIDVALIDITY del buzon cambio, TODOS los UIDs guardados dejan
    de referirse a los mismos correos (asi funciona IMAP), asi que el
    progreso se descarta -- avisando -- en vez de dar por revisados
    correos que en realidad son otros. Sin esta comprobacion, un cambio
    de UIDVALIDITY haria que el script se saltara correos nuevos en
    silencio, que es el peor error posible aca: perder correos sin que
    nadie se entere.

    'archivo' permite usar OTRO archivo de progreso (por defecto, el de
    este script). Lo usa descargar_correos_palabras_clave.py, que revisa
    el mismo Gmail pero para otra cosa: si compartieran el archivo, cada
    script daria por revisados los correos que proceso el otro y se
    saltaria correos sin avisar.
    """
    archivo = archivo or ARCHIVO_PROGRESO
    if not os.path.exists(archivo):
        return set()
    try:
        with open(archivo, "r", encoding="utf-8") as f:
            datos = json.load(f)
    except (OSError, ValueError) as error:
        logging.warning("[Progreso] No se pudo leer %s (%s) -- se empieza de cero.", archivo, error)
        return set()

    uidvalidity_guardado = datos.get("uidvalidity")
    if uidvalidity_actual and uidvalidity_guardado and str(uidvalidity_guardado) != str(uidvalidity_actual):
        logging.warning(
            "[Progreso] Gmail cambio el identificador del buzon (UIDVALIDITY %s -> %s): los correos ya "
            "revisados no se pueden reconocer con seguridad, asi que se revisa todo de nuevo (es lo seguro: "
            "asi no se salta ningun correo por error).", uidvalidity_guardado, uidvalidity_actual,
        )
        return set()

    revisados = set()
    for uid in datos.get("uids_revisados", []):
        try:
            revisados.add(int(uid))
        except (TypeError, ValueError):
            continue
    return revisados


def guardar_progreso(uids_revisados, uidvalidity, archivo=None):
    """
    Guarda en disco que correos ya se revisaron. Se escribe primero en
    un archivo temporal y despues se reemplaza el bueno, para que un
    corte de luz a mitad de la escritura no deje el archivo de progreso
    corrupto (y con el, la duda de que se reviso y que no).

    'archivo' permite usar OTRO archivo de progreso -- ver cargar_progreso.
    """
    archivo = archivo or ARCHIVO_PROGRESO
    temporal = archivo + ".tmp"
    try:
        with open(temporal, "w", encoding="utf-8") as f:
            json.dump(
                {
                    "uidvalidity": str(uidvalidity or ""),
                    "actualizado": datetime.datetime.now().isoformat(timespec="seconds"),
                    "uids_revisados": sorted(uids_revisados),
                },
                f,
            )
        os.replace(temporal, archivo)
    except OSError as error:
        logging.warning("[Progreso] No se pudo guardar el avance en %s: %s", archivo, error)


# ==================== Conexion con Gmail ====================


def _con_limite_de_tiempo(mail, funcion, *args, **kwargs):
    """
    Corre 'funcion' con un limite de tiempo DURO de TIMEOUT_SEGUNDOS,
    en un hilo aparte.

    Es una SEGUNDA linea de defensa ademas del limite del socket: en
    equipos con antivirus que inspecciona el correo (o con un proxy de
    por medio), el limite del socket puede NO dispararse nunca y el
    programa se queda esperando una respuesta que no va a llegar --
    colgado, sin ningun error en el log. Esto lo corta igual.

    El hilo es DAEMON a proposito: si cerrar el socket no alcanza a
    desbloquearlo, un hilo daemon nunca impide que el programa termine.
    Al agotarse el tiempo levanta TimeoutError (que es un OSError, asi
    que lo atrapa el mismo manejo de "se cayo la conexion").
    """
    resultado = {}

    def _ejecutar():
        try:
            resultado["valor"] = funcion(*args, **kwargs)
        except BaseException as error:  # noqa: BLE001 -- se re-lanza tal cual abajo
            resultado["error"] = error

    hilo = threading.Thread(target=_ejecutar, daemon=True)
    hilo.start()
    hilo.join(timeout=TIMEOUT_SEGUNDOS)
    if hilo.is_alive():
        try:
            sock = getattr(mail, "sock", None)
            if sock is not None:
                sock.close()
        except Exception:
            pass
        raise TimeoutError(f"Gmail no respondio en {TIMEOUT_SEGUNDOS}s (limite duro)")
    if "error" in resultado:
        raise resultado["error"]
    return resultado.get("valor")


def conectar(usuario, app_password):
    """
    Abre una conexion nueva con Gmail, inicia sesion y selecciona la
    carpeta de TODOS los correos. Devuelve (conexion, uidvalidity), o
    (None, None) si algo fallo.

    Cada paso deja su propia linea en el log: si alguna vez se traba,
    el ultimo mensaje dice exactamente donde -- sin tener que adivinar.
    """
    try:
        # Cada paso va con el limite de tiempo DURO, no solo con el del
        # socket: el login y la apertura de la carpeta tambien se
        # pueden quedar esperando para siempre en un equipo con
        # antivirus que inspecciona el correo.
        logging.info("[Correo] Conectando con Gmail (limite %ds por operacion)...", TIMEOUT_SEGUNDOS)
        mail = imaplib.IMAP4_SSL("imap.gmail.com", timeout=TIMEOUT_SEGUNDOS)
        logging.info("[Correo] Conectado -- iniciando sesion...")
        _con_limite_de_tiempo(mail, mail.login, usuario, app_password)
        logging.info("[Correo] Sesion iniciada -- abriendo la carpeta de todos los correos...")
        if not _con_limite_de_tiempo(mail, buscador.seleccionar_todos_los_correos, mail):
            logging.error("[Correo] No se pudo abrir la carpeta de 'Todos los correos' de Gmail.")
            return None, None
    except Exception as error:
        logging.error("[Correo] No se pudo conectar con Gmail: %s", error)
        return None, None

    uidvalidity = ""
    try:
        respuesta = mail.response("UIDVALIDITY")[1]
        if respuesta and respuesta[0]:
            uidvalidity = respuesta[0].decode("ascii", errors="ignore").strip()
    except Exception:
        uidvalidity = ""

    logging.info("[Correo] Carpeta abierta, listo para revisar.")
    return mail, uidvalidity


def _conectar_o_none(credenciales, uidvalidity_esperado):
    """
    Reconecta y comprueba que el buzon siga siendo "el mismo" (que no
    haya cambiado el UIDVALIDITY). Devuelve (conexion, uidvalidity), o
    (None, None) si no se pudo reconectar o si el buzon cambio -- en
    ese caso el que llama debe detenerse: los UIDs que quedaban
    pendientes ya no se refieren a los mismos correos.
    """
    mail, uidvalidity_nuevo = conectar(*credenciales)
    if mail is None:
        logging.error(
            "[Correo] No se pudo reconectar -- se detiene aqui. Lo revisado queda guardado; vuelve a correr "
            "el script para seguir donde se quedo."
        )
        return None, None
    if uidvalidity_nuevo and uidvalidity_esperado and uidvalidity_nuevo != uidvalidity_esperado:
        logging.error(
            "[Correo] Gmail cambio el identificador del buzon a mitad de la corrida -- se detiene aqui por "
            "seguridad (los UIDs pendientes ya no son de fiar). Vuelve a correr el script."
        )
        cerrar(mail)
        return None, None
    return mail, uidvalidity_nuevo


def cerrar(mail):
    """
    Cierra la sesion sin arriesgarse a colgar el programa: si la
    conexion ya quedo en mal estado, un logout() puede quedarse
    esperando una respuesta que nunca llega. Cualquier error se ignora
    -- a estas alturas ya no importa un cierre prolijo, solo no
    quedarse pegado.
    """
    if mail is None:
        return
    try:
        sock = getattr(mail, "sock", None)
        if sock is not None:
            sock.settimeout(TIMEOUT_SEGUNDOS)
        mail.logout()
    except Exception:
        pass


def listar_uids(mail):
    """
    Pide los UIDs de los correos a revisar, en UN PUÑADO de consultas
    (no una por termino del Excel -- eso es lo que hacia fallar al
    enfoque anterior).

    Con ACOTAR_POR_TIPO_EN_GMAIL (por defecto) NO se piden los correos
    de toda la ventana de fechas, sino solo los que ya mencionan alguna
    de las FRASES_EXTRAPROCESALES en su ASUNTO o en su TEXTO. La razon
    es puramente practica y salio de un caso real: dos años de correo
    eran 25.134 mensajes, y bajarlos TODOS completos (con adjuntos
    escaneados) satura la conexion -- Gmail termina cortandola. Filtrando
    de entrada por tipo quedan unos cientos: los que de verdad
    interesan, y se bajan sin problema.

    Cada consulta es lo mas basico del protocolo: SINCE (fecha) +
    SUBJECT/TEXT (una frase), IMAP4rev1 estandar, en ASCII puro, sin
    extensiones de Gmail (X-GM-RAW), sin acentos y sin combinar nada con
    OR. Son ~8 consultas en total, no 4242.

    Si una consulta falla, se registra y se sigue con las demas: es
    preferible revisar de menos que no revisar nada.
    """
    desde = (datetime.date.today() - datetime.timedelta(days=DIAS_HACIA_ATRAS)).strftime("%d-%b-%Y")

    def _buscar(*criterios):
        typ, datos = _con_limite_de_tiempo(mail, mail.uid, "SEARCH", None, *criterios)
        if typ != "OK" or not datos or not datos[0]:
            return []
        return [int(u) for u in datos[0].split() if u.isdigit()]

    if not (ACOTAR_POR_TIPO_EN_GMAIL and SOLO_INFORMACION_EXTRAPROCESAL):
        logging.info("[Correo] Pidiendo TODOS los correos de los ultimos %d dia(s)...", DIAS_HACIA_ATRAS)
        return _buscar("SINCE", desde)

    encontrados = set()
    # El ASUNTO primero (es lo mas rapido del lado del servidor y lo
    # que trae la mayoria), y despues el TEXTO completo, que ademas
    # atrapa los correos con asunto generico ("Notificacion 12345")
    # cuyo cuerpo si dice de que se trata.
    for campo in ("SUBJECT", "TEXT"):
        for frase in FRASES_EXTRAPROCESALES:
            try:
                nuevos = _buscar("SINCE", desde, campo, frase)
            except Exception as error:
                logging.warning(
                    "   [Correo] Fallo la busqueda %s '%s' (%s) -- se sigue con las demas.", campo, frase, error,
                )
                continue
            antes = len(encontrados)
            encontrados.update(nuevos)
            logging.info(
                "[Correo] %s '%s': %d correo(s) (%d nuevos).", campo, frase, len(nuevos), len(encontrados) - antes,
            )
    return sorted(encontrados)


_PATRON_UID_EN_RESPUESTA = re.compile(rb"UID\s+(\d+)")


def descargar_lote(mail, uids):
    """
    Baja de una sola vez el contenido completo de varios correos (por
    su UID). Devuelve {uid: mensaje_ya_parseado}.

    Bajarlos de a lotes -- y no de a uno -- es lo que hace que revisar
    cientos de correos sea viable: es una peticion a Gmail por lote en
    vez de una por correo.

    La respuesta de un FETCH multiple viene intercalada y el orden NO
    esta garantizado, asi que el UID de cada mensaje se saca de su
    propia cabecera de respuesta (no se asume que venga en el mismo
    orden que se pidio). Un mensaje suelto que no se pueda parsear se
    ignora sin tumbar el lote completo.
    """
    lista = ",".join(str(u) for u in uids)
    typ, datos = _con_limite_de_tiempo(mail, mail.uid, "FETCH", lista, "(UID BODY.PEEK[])")
    if typ != "OK" or not datos:
        return {}

    mensajes = {}
    for parte in datos:
        if not isinstance(parte, tuple) or len(parte) < 2 or not parte[1]:
            continue
        coincidencia = _PATRON_UID_EN_RESPUESTA.search(parte[0] or b"")
        if not coincidencia:
            continue
        uid = int(coincidencia.group(1))
        try:
            mensajes[uid] = email.message_from_bytes(parte[1])
        except Exception as error:
            logging.warning("   [Correo] No se pudo leer el correo UID %s: %s", uid, error)
    return mensajes


# ==================== Leer el contenido de un correo ====================


def _texto_de_pdf_en_memoria(contenido: bytes) -> str:
    """Texto de un PDF que esta en memoria (sin escribirlo al disco primero)."""
    if cruce_excel.PdfReader is None:
        return ""
    try:
        lector = cruce_excel.PdfReader(io.BytesIO(contenido))
        return "\n".join((pagina.extract_text() or "") for pagina in lector.pages)
    except Exception:
        return ""


def _texto_de_docx_en_memoria(contenido: bytes) -> str:
    """Texto de un DOCX que esta en memoria (sin escribirlo al disco primero)."""
    if cruce_excel.docx is None:
        return ""
    try:
        documento = cruce_excel.docx.Document(io.BytesIO(contenido))
        return "\n".join(p.text for p in documento.paragraphs)
    except Exception:
        return ""


def texto_de_adjunto(nombre: str, contenido: bytes) -> str:
    """
    Texto de adentro de un adjunto PDF/DOCX, para poder buscarle el
    radicado/cuenta/demandado. Devuelve "" si no es un tipo que se
    pueda leer, si pesa demasiado, o si esta dañado -- nunca revienta:
    un PDF corrupto no puede tumbar la revision del correo.
    """
    if not LEER_TEXTO_DE_ADJUNTOS:
        return ""
    if len(contenido) > MAX_MB_ADJUNTO_PARA_LEER * 1024 * 1024:
        return ""
    extension = Path(nombre).suffix.lower()
    if extension == ".pdf":
        return _texto_de_pdf_en_memoria(contenido)
    if extension == ".docx":
        return _texto_de_docx_en_memoria(contenido)
    return ""


def leer_correo(mensaje):
    """
    Saca de un mensaje ya descargado todo lo que hace falta: asunto,
    cuerpo (en texto plano), fecha, y los adjuntos con su contenido.
    Mismo formato que usa el resto del proyecto para guardar correos
    (ver clasificar_procesos_ejecutivos._guardar_correo).
    """
    asunto = buscador._decodificar_asunto(mensaje.get("Subject", ""))
    fecha = buscador._fecha_del_correo(mensaje)
    cuerpo = ""
    adjuntos = []
    for parte in mensaje.walk():
        try:
            nombre_adjunto = parte.get_filename()
            if nombre_adjunto:
                contenido = parte.get_payload(decode=True)
                if contenido:
                    adjuntos.append((organizador.sanear_nombre(nombre_adjunto), contenido))
            elif parte.get_content_type() in ("text/html", "text/plain"):
                texto_bruto = parte.get_payload(decode=True)
                if texto_bruto:
                    texto = texto_bruto.decode(parte.get_content_charset() or "utf-8", errors="ignore")
                    cuerpo += base._texto_plano_html(texto) + " "
        except Exception:
            # Una parte MIME rota (codificacion invalida, nombre de
            # archivo imposible) no puede hacer que se pierda el resto
            # del correo -- se salta esa parte y se sigue.
            continue
    return {"asunto": asunto, "cuerpo": cuerpo, "adjuntos": adjuntos, "fecha": fecha}


def analizar_correo(correo):
    """
    Prepara, UNA SOLA VEZ, las piezas del correo que hacen falta tanto
    para saber DE QUE TIPO es como para saber A QUE PROCESO va. Cada
    pieza es un par (titulo, texto):

      1. el asunto  ->  asunto + cuerpo
      2. cada adjunto  ->  su nombre + el texto de adentro del PDF/DOCX

    Se devuelven POR SEPARADO -- y no todo pegado en un texto gigante --
    porque las dos revisiones miran solo el ENCABEZADO del texto:

    * el DEMANDADO se busca en los primeros caracteres (ver
      _procesos_que_coinciden_con_correo), que es donde de verdad esta
      en un documento judicial ("DEMANDADO: ...");
    * el TIPO de documento se decide por el titulo, o por el inicio del
      contenido (ver es_informacion_no_procesal).

    Si se pegara todo junto, el encabezado del segundo adjunto en
    adelante quedaria fuera de esas ventanas y esos documentos no se
    podrian ni emparejar ni reconocer. Ademas leer el texto de un PDF
    es lo mas lento de todo, asi que se hace aqui una vez y se
    reutiliza en las dos revisiones.
    """
    piezas = []
    titulo_correo = buscador._normalizar_para_comparar(correo["asunto"])
    cuerpo = buscador._normalizar_para_comparar(
        correo["asunto"] + " " + correo["cuerpo"] + " " + " ".join(n for n, _ in correo["adjuntos"])
    )
    piezas.append((titulo_correo, cuerpo))
    for nombre, contenido in correo["adjuntos"]:
        nombre_norm = buscador._normalizar_para_comparar(nombre)
        texto = texto_de_adjunto(nombre, contenido)
        piezas.append((nombre_norm, buscador._normalizar_para_comparar(nombre + "\n" + texto)))
    return piezas


def es_extraprocesal(piezas):
    """
    True si el correo ES informacion extraprocesal: un derecho de
    peticion / peticion / PQR, una RESPUESTA a uno de ellos, una
    tutela, o un pago oficioso.

    Se revisa cada pieza POR SEPARADO (el asunto por un lado, y el
    nombre + contenido de CADA adjunto por otro): basta con que UNA lo
    sea. Esto importa de verdad -- un correo con asunto generico
    ("Notificacion 12345") pero con "RESPUESTA DERECHO DE PETICION.pdf"
    adjunto SI es lo que buscamos, y mirando solo el asunto (o todo el
    texto pegado, donde el nombre del adjunto queda enterrado despues
    de un cuerpo largo) se perderia.
    """
    for titulo, contenido in piezas:
        es_del_tipo, motivo = base.es_informacion_no_procesal(titulo, contenido)
        if es_del_tipo:
            return True, motivo
    return False, "no es tutela, ni derecho de peticion (ni su respuesta), ni pago oficioso"


def procesos_del_correo(piezas, indices):
    """
    Todos los procesos con los que este correo coincide, buscando en
    cada pieza por separado (ver analizar_correo) y juntando los
    resultados sin repetir.
    """
    encontrados = {}
    for _titulo, contenido in piezas:
        for proceso in base._procesos_que_coinciden_con_correo(contenido, indices):
            encontrados[proceso["nombre_carpeta"]] = proceso
    return list(encontrados.values())


# ==================== Guardar ====================


def _nombre_para_carpeta_manual(correo) -> str:
    """Nombre de subcarpeta legible para un correo sin coincidencias."""
    fecha = correo["fecha"].isoformat() if correo["fecha"] else "sin_fecha"
    asunto = (correo["asunto"] or "correo").strip()[:60]
    return organizador.sanear_nombre(f"{fecha} - {asunto}")


def guardar_en_proceso(correo, proceso) -> int:
    """Guarda el correo (adjuntos, o su texto si no trae) en la carpeta del proceso."""
    destino = Path(CARPETA_PROCESOS) / proceso["nombre_carpeta"]
    return base._guardar_correo(correo, destino)


def guardar_sin_clasificar(correo) -> int:
    """
    Guarda un correo que no coincidio con ningun proceso, en su propia
    subcarpeta dentro de CARPETA_SIN_CLASIFICAR -- una carpeta por
    correo para que se pueda revisar de a uno sin que se mezclen los
    adjuntos de correos distintos.
    """
    destino = Path(CARPETA_PROCESOS) / CARPETA_SIN_CLASIFICAR / _nombre_para_carpeta_manual(correo)
    return base._guardar_correo(correo, destino)


# ==================== Revisar un correo ====================


def revisar_correo(correo, indices, resumen, filas_reporte, filas_sin_coincidencia):
    """
    Decide a donde va UN correo y lo guarda. Devuelve None -- todo lo
    que hay que reportar queda en 'resumen' y en las listas de filas.
    """
    asunto = correo["asunto"] or "(sin asunto)"
    # Se prepara UNA sola vez (leer el texto de los PDF es lo mas
    # lento) y se reutiliza para las dos revisiones: tipo y proceso.
    piezas = analizar_correo(correo)
    motivo_tipo = ""

    if SOLO_INFORMACION_EXTRAPROCESAL:
        es_del_tipo, motivo_tipo = es_extraprocesal(piezas)
        if not es_del_tipo:
            resumen["descartados_por_tipo"] += 1
            return

    if EXIGIR_MENCION_ESSA:
        todo_el_texto = " ".join(contenido for _titulo, contenido in piezas)
        if not any(buscador._nombre_coincide(todo_el_texto, t) for t in buscador.TERMINOS_DEMANDANTE_VALIDO):
            resumen["descartados_por_essa"] += 1
            return

    procesos = procesos_del_correo(piezas, indices)
    fecha_texto = correo["fecha"].isoformat() if correo["fecha"] else ""
    detalle_tipo = f" ({motivo_tipo})" if motivo_tipo else ""

    if not procesos:
        resumen["sin_coincidencia"] += 1
        filas_sin_coincidencia.append((asunto, fecha_texto, motivo_tipo, len(correo["adjuntos"])))
        if MODO_PRUEBA:
            logging.info(
                "[SIMULACION] Sin coincidencia: '%s'%s -> iria a '%s'.",
                asunto, detalle_tipo, CARPETA_SIN_CLASIFICAR,
            )
            return
        guardados = guardar_sin_clasificar(correo)
        logging.info(
            "[Sin coincidencia] '%s'%s -> %d archivo(s) en '%s'.",
            asunto, detalle_tipo, guardados, CARPETA_SIN_CLASIFICAR,
        )
        return

    for proceso in procesos:
        filas_reporte.append((asunto, fecha_texto, motivo_tipo, proceso["numero"], proceso["nombre_carpeta"]))
        if MODO_PRUEBA:
            logging.info(
                "[SIMULACION] '%s'%s -> proceso %s ('%s').",
                asunto, detalle_tipo, proceso["numero"], proceso["nombre_carpeta"],
            )
            resumen["clasificados"] += 1
            continue
        guardados = guardar_en_proceso(correo, proceso)
        logging.info(
            "[Clasificado] '%s'%s -> proceso %s ('%s'), %d archivo(s).",
            asunto, detalle_tipo, proceso["numero"], proceso["nombre_carpeta"], guardados,
        )
        resumen["clasificados"] += 1


# ==================== Principal ====================


def procesar():
    if not os.path.exists(base.RUTA_EXCEL_CONTROL):
        logging.error("No se encontro el Excel configurado en RUTA_EXCEL_CONTROL: %r", base.RUTA_EXCEL_CONTROL)
        return

    credenciales = organizador.leer_credenciales()
    if not credenciales:
        logging.error(
            "No hay %s (o le faltan datos) -- sin las credenciales de Gmail no se puede revisar el correo. "
            "Ver el README.", organizador.ARCHIVO_CREDENCIALES,
        )
        return

    procesos = base.leer_procesos_control()
    con_radicado, terminados, _sin_estado = base.clasificar_procesos(procesos)
    todos = con_radicado + terminados
    logging.info(
        "Excel: %d proceso(s) activo(s)/suspendido(s)/etc + %d terminado(s) = %d para comparar.",
        len(con_radicado), len(terminados), len(todos),
    )
    # Se compara contra TODOS los procesos (activos Y terminados): a la
    # carpeta de un proceso ya terminado igual pueden llegar correos
    # (el auto que lo termina, una respuesta tardia, etc).
    indices = base._indexar_procesos_para_correo(todos)

    mail, uidvalidity = conectar(*credenciales)
    if mail is None:
        return

    resumen = {"clasificados": 0, "sin_coincidencia": 0, "descartados_por_tipo": 0, "descartados_por_essa": 0, "fallidos": 0}
    filas_reporte, filas_sin_coincidencia = [], []
    revisados = cargar_progreso(uidvalidity)

    try:
        if ACOTAR_POR_TIPO_EN_GMAIL and SOLO_INFORMACION_EXTRAPROCESAL:
            logging.info(
                "[Correo] Pidiendo a Gmail SOLO los correos extraprocesales de los ultimos %d dia(s) "
                "(%d frase(s) x asunto y texto = %d consulta(s), no una por proceso)...",
                DIAS_HACIA_ATRAS, len(FRASES_EXTRAPROCESALES), len(FRASES_EXTRAPROCESALES) * 2,
            )
        try:
            uids = listar_uids(mail)
        except Exception as error:
            logging.error("[Correo] No se pudo obtener la lista de correos: %s", error)
            return

        pendientes = [u for u in uids if u not in revisados]
        # El UID crece con el tiempo, asi que ordenar al reves = del
        # correo mas nuevo al mas viejo (ver EMPEZAR_POR_LOS_MAS_NUEVOS).
        if EMPEZAR_POR_LOS_MAS_NUEVOS:
            pendientes.sort(reverse=True)
        logging.info(
            "[Correo] %d correo(s) en la ventana de %d dia(s); %d ya revisados en corridas anteriores; "
            "quedan %d por revisar (empezando por los mas %s).",
            len(uids), DIAS_HACIA_ATRAS, len(uids) - len(pendientes), len(pendientes),
            "nuevos" if EMPEZAR_POR_LOS_MAS_NUEVOS else "viejos",
        )
        if not pendientes:
            logging.info("[Correo] No hay nada nuevo que revisar. Todo al dia.")
            return

        de_esta_corrida = pendientes[:MAX_CORREOS_POR_CORRIDA]
        if len(pendientes) > len(de_esta_corrida):
            corridas_estimadas = -(-len(pendientes) // MAX_CORREOS_POR_CORRIDA)  # division hacia arriba
            logging.info(
                "[Correo] Esta corrida revisa hasta %d correo(s) (tope por cantidad: MAX_CORREOS_POR_CORRIDA; "
                "tope por tiempo: %s). Quedan %d para las proximas -- son unas %d corrida(s) mas para ponerte "
                "al dia; vuelve a correr el script las veces que haga falta, siempre sigue donde se quedo.",
                len(de_esta_corrida),
                f"{MAX_MINUTOS_POR_CORRIDA} min" if MAX_MINUTOS_POR_CORRIDA else "sin tope",
                len(pendientes) - len(de_esta_corrida), corridas_estimadas,
            )

        reconexiones = 0
        procesados_en_esta_corrida = 0
        comenzo_en = time.monotonic()

        for inicio in range(0, len(de_esta_corrida), TAMANO_LOTE_DESCARGA):
            # Tope por TIEMPO: se revisa ANTES de empezar cada grupo,
            # nunca a mitad de uno, para cerrar siempre en un punto
            # limpio (con el progreso guardado y sin correos a medias).
            #
            # El PRIMER grupo siempre se procesa (inicio > 0), pase lo
            # que pase con el reloj: si no, un tope de tiempo demasiado
            # bajo -- o una maquina muy lenta -- haria que cada corrida
            # terminara sin revisar ni un solo correo, y el usuario
            # correria el script una y otra vez sin avanzar NUNCA.
            if MAX_MINUTOS_POR_CORRIDA and inicio > 0:
                minutos = (time.monotonic() - comenzo_en) / 60
                if minutos >= MAX_MINUTOS_POR_CORRIDA:
                    logging.info(
                        "[Correo] Se cumplieron los %d minuto(s) de esta corrida (MAX_MINUTOS_POR_CORRIDA) -- "
                        "se cierra aqui con %d correo(s) revisados. Lo que falta queda guardado: vuelve a "
                        "correr el script para seguir donde se quedo.",
                        MAX_MINUTOS_POR_CORRIDA, procesados_en_esta_corrida,
                    )
                    break

            lote = de_esta_corrida[inicio:inicio + TAMANO_LOTE_DESCARGA]
            intentos = 0
            mensajes = None
            lote_abandonado = False

            while mensajes is None:
                try:
                    # descargar_lote ya aplica el limite de tiempo duro
                    # internamente sobre la peticion a Gmail.
                    mensajes = descargar_lote(mail, lote)
                except (imaplib.IMAP4.abort, imaplib.IMAP4.error, OSError) as error:
                    intentos += 1
                    if intentos > MAX_REINTENTOS_POR_LOTE:
                        logging.error(
                            "[Correo] Un grupo de %d correo(s) sigue fallando despues de %d intento(s) (%s) -- "
                            "se salta y se sigue con los demas. NO se dan por revisados: quedan pendientes "
                            "para la proxima corrida.", len(lote), intentos, error,
                        )
                        resumen["fallidos"] += len(lote)
                        lote_abandonado = True
                        break
                    if reconexiones >= MAX_RECONEXIONES:
                        logging.error(
                            "[Correo] Se perdio la conexion (%s) y ya se reconecto %d vez/veces -- se detiene "
                            "aqui. Lo revisado hasta ahora queda guardado: vuelve a correr el script para "
                            "seguir donde se quedo.", error, reconexiones,
                        )
                        return
                    reconexiones += 1
                    logging.warning(
                        "[Correo] Se perdio la conexion con Gmail (%s) -- reconectando (intento %d/%d) y "
                        "siguiendo donde se quedo (%d correo(s) revisados en esta corrida)...",
                        error, reconexiones, MAX_RECONEXIONES, procesados_en_esta_corrida,
                    )
                    cerrar(mail)
                    mail, _uidvalidity_nuevo = _conectar_o_none(credenciales, uidvalidity)
                    if mail is None:
                        return

            if lote_abandonado:
                # A proposito NO se marcan como revisados: perder
                # correos en silencio es el peor resultado posible.
                # Quedan pendientes y se reintentan en la proxima
                # corrida (donde quizas la conexion este mejor).
                #
                # Pero SI hay que reconectar antes de seguir: esta
                # conexion acaba de fallar 4 veces seguidas, y si el
                # fallo fue un corte a mitad de una descarga, quedan
                # datos del correo a medio leer en la conexion. Seguir
                # usandola hace que la SIGUIENTE orden lea esa basura
                # como si fuera respuesta del servidor -- en el log
                # real se vio exactamente eso: "unexpected response:
                # b'Delivered-To: ...'". Con una conexion nueva se
                # arranca limpio.
                logging.info("[Correo] Se descarta la conexion (quedo en mal estado) y se abre una nueva...")
                cerrar(mail)
                mail, _uidvalidity_nuevo = _conectar_o_none(credenciales, uidvalidity)
                if mail is None:
                    return
                continue

            for uid in lote:
                mensaje = mensajes.get(uid)
                if mensaje is None:
                    # No llego en la respuesta (borrado, movido, o
                    # ilegible). Se marca como revisado igual: si no,
                    # cada corrida volveria a intentarlo para siempre.
                    revisados.add(uid)
                    continue
                try:
                    correo = leer_correo(mensaje)
                    revisar_correo(correo, indices, resumen, filas_reporte, filas_sin_coincidencia)
                except Exception as error:
                    # Un correo problematico se registra y se sigue --
                    # jamas puede tumbar la corrida completa.
                    resumen["fallidos"] += 1
                    logging.warning("   [Correo] No se pudo procesar el correo UID %s: %s", uid, error)
                revisados.add(uid)
                procesados_en_esta_corrida += 1

                if procesados_en_esta_corrida % GUARDAR_PROGRESO_CADA_N == 0:
                    guardar_progreso(revisados, uidvalidity)
                if procesados_en_esta_corrida % AVISO_PROGRESO_CADA_N == 0:
                    logging.info(
                        "[Correo] ...van %d/%d correo(s) revisados en esta corrida (%d clasificado(s), "
                        "%d sin coincidencia)...",
                        procesados_en_esta_corrida, len(de_esta_corrida),
                        resumen["clasificados"], resumen["sin_coincidencia"],
                    )
    finally:
        guardar_progreso(revisados, uidvalidity)
        cerrar(mail)

    logging.info(
        "Listo: %d correo(s) clasificado(s) a la carpeta de su proceso, %d sin coincidencia (a '%s'), "
        "%d con error.",
        resumen["clasificados"], resumen["sin_coincidencia"], CARPETA_SIN_CLASIFICAR, resumen["fallidos"],
    )
    if resumen["descartados_por_tipo"]:
        logging.info(
            "%d correo(s) se omitieron por no ser tutela/derecho de peticion/pago oficioso "
            "(SOLO_INFORMACION_EXTRAPROCESAL activo).", resumen["descartados_por_tipo"],
        )
    if resumen["descartados_por_essa"]:
        logging.info(
            "%d correo(s) se omitieron por no mencionar a ESSA (EXIGIR_MENCION_ESSA activo).",
            resumen["descartados_por_essa"],
        )

    if filas_reporte and cruce_excel._escribir_csv_tolerante(
        ARCHIVO_REPORTE, ["Asunto", "Fecha", "Por que es extraprocesal", "Proceso", "Carpeta"],
        filas_reporte, "Clasificados (correo)",
    ):
        logging.info("[Reporte] %d fila(s) guardadas en: %s", len(filas_reporte), ARCHIVO_REPORTE)
    if filas_sin_coincidencia and cruce_excel._escribir_csv_tolerante(
        ARCHIVO_SIN_COINCIDENCIA, ["Asunto", "Fecha", "Por que es extraprocesal", "Adjuntos"],
        filas_sin_coincidencia, "Sin coincidencia (correo)",
    ):
        logging.info("[Reporte] %d fila(s) guardadas en: %s", len(filas_sin_coincidencia), ARCHIVO_SIN_COINCIDENCIA)

    if MODO_PRUEBA:
        logging.info(
            "MODO_PRUEBA esta activo: NO se descargo ni se guardo nada todavia, solo se mostro que se haria. "
            "Revisa el log y, si se ve bien, cambia MODO_PRUEBA = False al inicio de este script y vuelve a "
            "correrlo."
        )


def main():
    configurar_logging()
    procesar()


if __name__ == "__main__":
    main()
