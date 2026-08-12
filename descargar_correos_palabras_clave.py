"""
Busca en tu Gmail los correos de PAGO OFICIOSO, DERECHO DE PETICION y
SENTENCIA DE TUTELA, y DESCARGA sus adjuntos a UNA SOLA CARPETA,
renombrandolos con el RADICADO, el NUMERO DE CUENTA o el ASUNTO del
correo -- lo que encuentre, en ese orden.

EN QUE SE DIFERENCIA DE revisar_correo_pro.py
============================================================
Los dos leen el mismo Gmail y buscan cosas parecidas, pero sirven para
cosas distintas:

                      revisar_correo_pro.py     ESTE script
    A donde va todo   a la carpeta de CADA       a UNA SOLA carpeta
                      proceso (una por proceso)  (todo junto)
    Necesita el Excel SI (compara contra el      NO (no lo abre
    de control        control de procesos)       siquiera)
    Que busca         todo lo extraprocesal      SOLO las 3 frases
                      (peticion/PQR/tutela/      pedidas (ver
                      pago oficioso)             FRASES_A_BUSCAR)
    Como nombra       deja el nombre original    RENOMBRA con el
                      del adjunto                radicado/cuenta/asunto

Usa este cuando lo que quieres es TENER LOS DOCUMENTOS A LA MANO, todos
juntos y ya identificados, sin depender de que el proceso este en el
Excel ni de que exista su carpeta. Usa el otro cuando lo que quieres es
que cada correo quede archivado en la carpeta de su proceso.

Los dos se pueden usar a la vez: cada uno lleva su PROPIO archivo de
progreso, asi que no se estorban ni se saltan correos entre ellos.

QUE CORREOS SE DESCARGAN
============================================================
Solo los que SON de uno de estos tres tipos (ver FRASES_A_BUSCAR):

    PAGO OFICIOSO
    DERECHO DE PETICION
    SENTENCIA DE TUTELA

De cada uno se reconocen sus variantes normales de escritura: con o sin
tilde ("PETICIÓN" = "PETICION"), en singular o plural, y las formas con
que se llaman en la practica (a una sentencia de tutela casi siempre le
dicen "fallo de tutela").

El documento tiene que SER de uno de esos tipos, no basta con que los
MENCIONE (ver tipos_de_pieza). Se decide asi:

  1. si el TITULO lo dice (el asunto, o el nombre del archivo:
     "RESPUESTA DERECHO DE PETICION.pdf") -> entra;
  2. si el titulo dice que es OTRA COSA -- un auto, un mandamiento, una
     demanda, un poder... (PALABRAS_QUE_DESCARTAN) -> no entra, aunque
     su contenido mencione la frase;
  3. si el titulo no dice ni una cosa ni la otra, la frase tiene que
     estar en el ENCABEZADO del documento, que es donde un escrito
     judicial se titula a si mismo.

El paso 2 es el que evita el error tipico: un correo "Notifica AUTO
AVOCA accion de tutela" cuyo cuerpo cuenta que el accionante alega la
vulneracion de su "derecho de peticion" NO es un derecho de peticion.

Ademas, de un correo se bajan SOLO los adjuntos que son del tipo
buscado (SOLO_EL_DOCUMENTO_QUE_COINCIDE): de una notificacion que trae
la sentencia de tutela junto con la constancia de envio, el poder y un
pantallazo, se baja la sentencia y nada mas.

Por defecto (EXIGIR_ADJUNTOS) solo se descargan los correos que TRAEN
ADJUNTOS, que es lo que se pidio. Si quieres guardar tambien los que
vienen en puro texto (se guardan como .txt), pon EXIGIR_ADJUNTOS = False.

COMO SE NOMBRA CADA ARCHIVO
============================================================
    <IDENTIFICADOR> - <TIPO> - <nombre original del adjunto>.pdf

El IDENTIFICADOR se busca en este orden, y se usa el PRIMERO que
aparezca (de mas confiable a menos):

  1. el RADICADO (23 digitos), asi venga escrito con guiones, puntos o
     espacios entre sus grupos -- "68001-40-03-001-2024-00050-00" es el
     mismo numero que "68001400300120240005000";
  2. el NUMERO DE CUENTA, cuando el texto lo rotula como tal ("CUENTA
     No. 1234567", "NIC 1234567", "cuenta contrato 1234567"). Se exige
     el rotulo a proposito: un numero suelto de 7 digitos en un correo
     puede ser cualquier cosa (un telefono, un radicado de PQR, una
     factura), y nombrar un documento con el numero equivocado es peor
     que no nombrarlo;
  3. si no hay ninguno de los dos, la FECHA + el ASUNTO del correo.

Ejemplos de como quedan los archivos en la carpeta:

    68001400300120240005000 - SENTENCIA DE TUTELA - fallo primera instancia.pdf
    1234567 - PAGO OFICIOSO - respuesta essa.pdf
    2026-03-14 Notificacion judicial - DERECHO DE PETICION - anexo.pdf

Ademas se deja un CSV (ARCHIVO_REPORTE) dentro de la misma carpeta con
una fila por archivo guardado: de que correo salio, de que tipo es, que
identificador se le puso y DE DONDE se saco -- para poder revisar de un
vistazo los que quedaron nombrados por asunto (que son los que quizas
quieras renombrar a mano).

NUNCA PISA UN ARCHIVO YA GUARDADO. Si dos correos generan el mismo
nombre, el segundo queda como "..._2.pdf" (ver _ruta_archivo_libre).

POR QUE NO SE TRABA (y por que no repite trabajo)
============================================================
Reutiliza la conexion con Gmail de revisar_correo_pro.py, que ya trae
resuelto lo que fallaba en la practica: limite de tiempo DURO por
operacion (para equipos con antivirus que inspecciona el correo),
reconexion automatica, descarga por lotes, UIDs en vez de numeros de
orden, y tope por cantidad y por tiempo para que la corrida SIEMPRE
termine. Encima de eso, tres cosas propias de este script, cada una
sacada de un fallo real:

* EL PROGRESO SE GUARDA DESPUES DE CADA CORREO (no cada 25, ver
  GUARDAR_PROGRESO_CADA_N). Con 25, una caida despues de bajar 24
  correos dejaba esos 24 en el disco pero sin registrar, y la corrida
  siguiente los bajaba otra vez: aparecian duplicados "..._2.pdf".
* SE LE PIDE A GMAIL LA FRASE EXACTA, y solo correos con adjuntos (ver
  BUSQUEDA_EXACTA_EN_GMAIL). Buscando por palabras sueltas, "TUTELA"
  sola aparecia en 926 correos (firmas, hilos citados, "responder a
  todos") y habia que bajarlos todos completos para descartarlos.
* MAS TIEMPO Y LOTES MAS PEQUEÑOS (TIMEOUT_SEGUNDOS, TAMANO_LOTE_DESCARGA):
  aca casi todo correo trae adjuntos escaneados, y con el limite de 60s
  de revisar_correo_pro.py se cortaban descargas que iban bien, solo
  lentas -- la corrida se iba entera en reconectar sin avanzar.

Si lo cortas o se cae la conexion, la proxima corrida SIGUE DONDE SE
QUEDO (ARCHIVO_PROGRESO, que es suyo y no el del otro script).

Necesita credenciales_sgde.txt (las mismas de siempre, contraseña de
aplicacion de Gmail -- ver README). Si no existe, avisa y no hace nada.
"""

import datetime
import imaplib
import io
import logging
import os
import re
import time
import zipfile
from pathlib import Path

import buscar_faltantes_en_drive as buscador
import procesos_juridicos as organizador
import validar_renombrar_carpetas as cruce_excel
import revisar_correo_pro as correo_pro
import clasificar_procesos_ejecutivos as base

# ============================= CONFIGURACION =============================

# Nombre de la carpeta UNICA donde se guarda todo lo descargado.
NOMBRE_CARPETA_DESCARGAS = "CORREOS DESCARGADOS - PAGO OFICIOSO, PETICION Y TUTELA"


def _carpeta_descargas_por_defecto() -> str:
    """
    Donde dejar la carpeta de descargas si no se configura a mano.

    Se prueba, en orden: (1) al lado de la carpeta de procesos del resto
    del proyecto -- que es donde se espera tener todo junto; (2) la
    carpeta Documentos del usuario que este usando el equipo; (3) al
    lado de este mismo script.

    Se comprueba que cada una EXISTA antes de usarla, en vez de dar por
    hecho la primera: la ruta de CARPETA_PROCESOS esta escrita para un
    equipo concreto (C:\\Users\\Francy\\...), y en OTRO equipo crearia
    una carpeta "C:\\Users\\Francy" nueva y vacia -- el usuario dejaria
    corriendo la herramienta y despues no encontraria por ningun lado
    los archivos descargados.
    """
    candidatas = [
        os.path.dirname(base.CARPETA_PROCESOS),
        os.path.join(os.path.expanduser("~"), "Documents"),
        os.path.expanduser("~"),
    ]
    for carpeta in candidatas:
        if carpeta and os.path.isdir(carpeta):
            return os.path.join(carpeta, NOMBRE_CARPETA_DESCARGAS)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), NOMBRE_CARPETA_DESCARGAS)


# Carpeta UNICA donde se guarda todo lo descargado. Si no existe, se
# crea. Cambiala por la que quieras -- por ejemplo:
#     CARPETA_DESCARGAS = r"C:\\Users\\TuUsuario\\Documents\\CORREOS DESCARGADOS"
# Al arrancar, el log dice siempre en que carpeta va a guardar, para que
# no haya que adivinarlo.
CARPETA_DESCARGAS = _carpeta_descargas_por_defecto()

# False: descarga de verdad. True: solo muestra en el log lo que
# descargaria y con que nombre, sin guardar nada.
#
# OJO -- aqui el valor por defecto es AL REVES que en el resto del
# proyecto (donde MODO_PRUEBA arranca en True). La razon: los otros
# scripts MUEVEN, RENOMBRAN o BORRAN carpetas que ya existen, y ahi una
# equivocacion cuesta caro. Este solo CREA archivos nuevos dentro de una
# carpeta propia -- no toca, mueve ni borra nada de lo que ya tienes --
# asi que lo peor que puede pasar es que descargue algo que no querias,
# y para eso basta borrarlo. Ponlo en True si prefieres ver primero el
# log antes de bajar nada.
MODO_PRUEBA = False

# Cuantos dias hacia atras revisar. 730 = los ultimos 2 años.
DIAS_HACIA_ATRAS = 730

# True (por defecto): solo se descargan los correos que TRAEN ADJUNTOS
# -- que es lo que se pidio. Con False, un correo del tipo buscado que
# venga en puro texto tambien se guarda, como un .txt con su asunto y
# su cuerpo (para no perder, por ejemplo, una respuesta de pago
# oficioso escrita en el cuerpo del correo).
EXIGIR_ADJUNTOS = True

# Extensiones de adjunto que SI se guardan. Lo que no este aqui se
# omite (firmas .p7s, calendarios .ics, etc). Los .zip no se guardan
# como zip: se abren y se sacan los PDF/DOCX de adentro.
EXTENSIONES_A_GUARDAR = {
    ".pdf", ".docx", ".doc", ".rtf", ".odt", ".txt",
    ".xlsx", ".xls", ".csv",
    ".jpg", ".jpeg", ".png", ".tif", ".tiff",
}

# Las imagenes mas livianas que esto (en KB) se omiten: casi siempre
# son el logo, la firma o los iconos de redes sociales del pie del
# correo, no un documento escaneado. Un documento escaneado de verdad
# pesa mucho mas que esto.
MIN_KB_IMAGEN = 40

# True: al nombre del archivo se le pega tambien el nombre original del
# adjunto (recortado si hace falta). Ayuda a distinguir los varios
# adjuntos de un mismo correo. Ponlo en False si quieres nombres mas
# cortos.
INCLUIR_NOMBRE_ORIGINAL = True

# Largo maximo del nombre de archivo (sin contar la carpeta). Windows
# corta en 255, pero un nombre asi de largo es incomodo de leer y suma
# al limite de 260 caracteres de la RUTA completa.
MAX_CARACTERES_NOMBRE = 120

# Cuantos caracteres del asunto se usan cuando el correo no trae ni
# radicado ni cuenta y hay que nombrarlo por el asunto.
MAX_CARACTERES_ASUNTO = 60

# True (por defecto): de un correo se bajan SOLO los adjuntos que son
# del tipo buscado, no todos.
#
# Un correo de notificacion judicial trae la sentencia de tutela y, con
# ella, el auto, la constancia de envio, el poder y un pantallazo. Solo
# el primero es lo que se pidio. Con False se bajan todos los documentos
# de cualquier correo que coincida.
#
# Si el que coincidio fue el CORREO (por su asunto) y ningun adjunto
# coincide por si mismo, no hay como saber cual es "el bueno" y se bajan
# todos igual -- es preferible eso a no bajar nada.
SOLO_EL_DOCUMENTO_QUE_COINCIDE = True

# Cuantos caracteres del principio de un texto cuentan como su
# "encabezado" para decidir de que tipo es (ver tipos_de_pieza). Un
# documento judicial se titula a si mismo en las primeras lineas; mas
# alla de eso, la frase que aparezca casi siempre es una mencion de
# pasada. Subirlo trae mas documentos que no van; bajarlo puede dejar
# afuera los que traen un membrete largo antes del titulo.
VENTANA_ENCABEZADO = 900

# True (por defecto): ademas del asunto, el cuerpo y el nombre de los
# adjuntos, se LEE EL TEXTO de adentro de los PDF/DOCX adjuntos, tanto
# para reconocer el tipo de documento como para buscarle el radicado o
# la cuenta. Es lo mas lento de todo, pero es lo que hace que funcione:
# el radicado casi nunca esta en el asunto, esta impreso dentro del PDF.
LEER_TEXTO_DE_ADJUNTOS = True

# De cada PDF se leen solo las primeras paginas: es donde esta el titulo
# del documento y su radicado. Leer un escaneo de 80 paginas entero
# tarda minutos -- con la conexion a Gmail abierta esperando, que es
# justo lo que hace que el servidor la corte.
MAX_PAGINAS_PDF_PARA_LEER = 5

# Adjuntos mas pesados que esto no se abren para leerles el texto (un
# escaneo grande puede tardar muchisimo). Igual se guardan completos:
# esto solo evita leerlos para identificarlos.
MAX_MB_ADJUNTO_PARA_LEER = 25

# Maximo de correos a revisar POR CORRIDA, y tope por tiempo (minutos).
# Es lo que garantiza que la corrida SIEMPRE termine: si queda mas, lo
# dice en el log y basta volver a correrlo (sigue donde se quedo).
# MAX_MINUTOS_POR_CORRIDA = 0 quita el tope por tiempo.
MAX_CORREOS_POR_CORRIDA = 500
MAX_MINUTOS_POR_CORRIDA = 30

# True: empieza por los correos MAS NUEVOS (que suelen ser los mas
# urgentes) y va hacia atras.
EMPEZAR_POR_LOS_MAS_NUEVOS = True

# Cuantos correos se bajan por cada peticion a Gmail.
#
# 2, no 5: aca casi TODOS los correos traen adjuntos (por eso se
# buscan), y muchos son escaneos pesados. Pedir 5 de golpe hacia que una
# sola peticion tuviera que traer decenas de MB, tardara mas del limite
# de tiempo y la conexion se cayera -- y al reintentar el mismo grupo
# volvia a pasar lo mismo. Con 2 cada peticion es mas corta y, si algo
# falla, se pierde menos trabajo.
TAMANO_LOTE_DESCARGA = 2

# Limite de tiempo para cada operacion contra Gmail.
#
# 180s, no los 60s de revisar_correo_pro.py: ese script revisa correos
# de cualquier tipo (la mayoria livianos) y 60s le sobran, pero bajar un
# par de escaneos de varios MB por una conexion domestica se pasa de 60s
# con facilidad. Cuando eso ocurria, el limite cortaba una descarga que
# iba PERFECTAMENTE BIEN, solo lenta, y la corrida se iba en reconectar
# una y otra vez sin avanzar.
TIMEOUT_SEGUNDOS = 180

# Reintentos y reconexiones: se reutilizan de revisar_correo_pro.py.
MAX_RECONEXIONES = correo_pro.MAX_RECONEXIONES
MAX_REINTENTOS_POR_LOTE = correo_pro.MAX_REINTENTOS_POR_LOTE

# Los tres tipos que se buscan, con sus variantes de escritura. La
# comparacion se hace SIN TILDES y en MAYUSCULAS (ver _texto_comparable),
# asi que no hace falta repetir "PETICIÓN" y "PETICION": con una basta.
#
# La clave de cada grupo ("PAGO OFICIOSO", ...) es el nombre del tipo
# que queda escrito en el NOMBRE DEL ARCHIVO; las de la lista son solo
# las formas en que puede venir escrito en el correo.
FRASES_A_BUSCAR = {
    "PAGO OFICIOSO": [
        "PAGO OFICIOSO",
        "PAGOS OFICIOSOS",
        "PAGO DE MANERA OFICIOSA",
        "PAGO EN FORMA OFICIOSA",
    ],
    # "DERECHO DE PETICION" cubre de una vez la peticion Y su respuesta:
    # una respuesta se rotula "RESPUESTA DERECHO DE PETICION ...", asi
    # que la frase aparece igual.
    "DERECHO DE PETICION": [
        "DERECHO DE PETICION",
        "DERECHOS DE PETICION",
    ],
    # En la practica a la sentencia de tutela casi siempre le dicen
    # "FALLO de tutela" -- es el mismo documento, por eso cuenta.
    "SENTENCIA DE TUTELA": [
        "SENTENCIA DE TUTELA",
        "SENTENCIA TUTELA",
        "FALLO DE TUTELA",
        "FALLO TUTELA",
        "SENTENCIA DE PRIMERA INSTANCIA TUTELA",
        "SENTENCIA DE SEGUNDA INSTANCIA TUTELA",
    ],
}

# True (por defecto): se le pide a Gmail que busque las FRASES EXACTAS
# (la misma busqueda que escribirlas entre comillas en la barra de
# Gmail), y ademas que devuelva solo correos CON ADJUNTOS.
#
# Esto es lo que decide si la corrida dura minutos u horas. Buscando por
# palabras sueltas, en un caso real Gmail devolvia 971 candidatos --
# porque "TUTELA" aparece por si sola en 926 correos: en la firma del
# remitente, en el hilo citado abajo, en un "responder a todos" de otra
# cosa. Bajar 971 correos COMPLETOS, con sus adjuntos escaneados, es lo
# que hacia que la descarga tardara horas y que la conexion terminara
# cayendose. Pidiendo la frase exacta quedan los que de verdad lo son.
#
# Ponlo en False para volver a la busqueda amplia (por si alguna vez
# sospechas que Gmail esta dejando algo por fuera): revisa muchos mas
# correos, mucho mas lento, pero el filtro fino en tu PC es el mismo.
BUSQUEDA_EXACTA_EN_GMAIL = True

# Si el TITULO (asunto del correo o nombre del archivo) trae alguna de
# estas marcas, la pieza se descarta aunque su contenido mencione una de
# las frases buscadas: el titulo ya dijo que es otra cosa. Ver el paso 2
# de tipos_de_pieza.
#
# De aqui salio el caso real: "Notifica AUTO AVOCA accion de tutela..."
# cuyo cuerpo menciona el "derecho de peticion" del accionante.
#
# OJO con dos ausencias, las dos a proposito:
#
# * "SENTENCIA" NO esta (aunque el resto del proyecto si la excluye, ver
#   PALABRAS_PROCESAL_EXCLUIR): aca una SENTENCIA DE TUTELA es
#   justamente de lo que se trata. Si un archivo se llama solo
#   "SENTENCIA.pdf" y su encabezado dice que es de tutela, tiene que
#   poder entrar.
# * "NOTIFICACION" tampoco: los juzgados notifican TODO, incluidas las
#   sentencias de tutela, asi que descartaria justo lo que se busca.
#   Lo que si descarta es "AUTO", que es un documento distinto.
PALABRAS_QUE_DESCARTAN = [
    "AUTO", "AUTOS", "AVOCA", "ADMISORIO", "ADMITE",
    "MANDAMIENTO", "DEMANDA", "MEMORIAL", "TRASLADO", "EXCEPCIONES",
    "RECURSO DE REPOSICION", "RECURSO DE APELACION", "REQUERIMIENTO",
    "MEDIDA CAUTELAR", "EMBARGO", "SECUESTRO", "PODER",
    "LIQUIDACION DE CREDITO", "SOLICITUD DE CONCILIACION",
    "CONSTANCIA", "CITACION", "EDICTO", "FACTURA",
]

# Lo que se le pide a Gmail cuando BUSQUEDA_EXACTA_EN_GMAIL esta en
# False: TROZOS de palabra, sin tildes (la busqueda de IMAP es por
# subcadena y en ASCII):
#   "OFICIOSO" -> PAGO OFICIOSO, PAGOS OFICIOSOS, pago de manera oficiosa
#   "PETICI"   -> PETICION, PETICIÓN, PETICIONES
#   "TUTELA"   -> SENTENCIA DE TUTELA, FALLO DE TUTELA
#
# Sea cual sea la busqueda, lo que llegue se vuelve a revisar en tu PC
# con FRASES_A_BUSCAR (que es mas estricto): un correo que solo diga
# "TUTELA" de pasada se descarta igual, sin bajarlo.
FRASES_PARA_PEDIR_A_GMAIL = ["OFICIOSO", "PETICI", "TUTELA"]

# Archivo donde se guarda que correos ya se revisaron (para poder
# hacerlo por partes). Es SUYO, distinto al de revisar_correo_pro.py.
ARCHIVO_PROGRESO = os.path.join(
    os.path.dirname(__file__), "descargar_correos_palabras_clave_progreso.json"
)

ARCHIVO_LOG = os.path.join(os.path.dirname(__file__), "descargar_correos_palabras_clave.log")

# Reporte de lo descargado. Va DENTRO de la carpeta de descargas, para
# que quede junto con los archivos a los que se refiere.
NOMBRE_REPORTE = "_correos descargados.csv"

# Cada cuantos correos se guarda el progreso. 1 = DESPUES DE CADA
# CORREO, a proposito.
#
# Guardarlo cada 25 (como hace revisar_correo_pro.py, que no descarga
# nada pesado) sale carisimo aca: si se cae la conexion despues de bajar
# 24 correos con adjuntos, esos 24 ya estan en el disco pero ninguno
# quedo registrado, asi que la proxima corrida los vuelve a bajar
# enteros y quedan duplicados ("..._2.pdf"). Escribir un JSON de unos
# pocos KB es instantaneo comparado con volver a descargar un solo
# correo con adjuntos escaneados.
GUARDAR_PROGRESO_CADA_N = 1

# Cada cuantos correos se deja un aviso de avance en el log.
AVISO_PROGRESO_CADA_N = 10

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


# ==================== Reconocer el tipo de correo ====================


_PATRON_ESPACIOS = re.compile(r"\s+")


def _texto_comparable(texto: str) -> str:
    """
    Deja el texto listo para buscarle las frases: EN MAYUSCULAS, SIN
    TILDES (buscador._normalizar_para_comparar) y con todos los espacios
    en blanco reducidos a UNO SOLO.

    Lo de los espacios no es cosmetico: en un PDF real la frase casi
    nunca viene en una sola linea limpia -- llega como "SENTENCIA DE\\n
    TUTELA" o con doble espacio entre palabras. Sin unificarlos, buscar
    "SENTENCIA DE TUTELA" no encontraria justo el documento que mas
    importa.
    """
    return _PATRON_ESPACIOS.sub(" ", buscador._normalizar_para_comparar(texto))


def _texto_de_pdf(contenido: bytes) -> str:
    """
    Texto de las PRIMERAS MAX_PAGINAS_PDF_PARA_LEER paginas de un PDF que
    esta en memoria.

    No se leen todas a proposito, y son dos motivos distintos:

    * VELOCIDAD. Sacarle el texto a un escaneo de 80 paginas puede tardar
      minutos, y hay que hacerlo con cientos de correos. Con la conexion
      a Gmail abierta esperando, esa demora es justo la que hace que el
      servidor la corte.
    * PRECISION. Lo que se busca (que TIPO de documento es, y su
      radicado) esta siempre en el encabezado, en la primera pagina. Lo
      que aparece en la pagina 40 suele ser una cita de otro proceso o
      un anexo, y meterlo a la comparacion solo genera falsos positivos.

    Devuelve "" si el PDF esta dañado -- nunca revienta.
    """
    if cruce_excel.PdfReader is None:
        return ""
    try:
        lector = cruce_excel.PdfReader(io.BytesIO(contenido))
        paginas = list(lector.pages)[:MAX_PAGINAS_PDF_PARA_LEER]
        return "\n".join((pagina.extract_text() or "") for pagina in paginas)
    except Exception:
        return ""


def _texto_de_adjunto(nombre: str, contenido: bytes) -> str:
    """
    Texto de adentro de un adjunto PDF/DOCX. Devuelve "" si no es de un
    tipo que se pueda leer, si pesa mas de MAX_MB_ADJUNTO_PARA_LEER, o
    si esta dañado -- nunca revienta: un PDF corrupto no puede tumbar la
    revision del correo.
    """
    if not LEER_TEXTO_DE_ADJUNTOS:
        return ""
    if len(contenido) > MAX_MB_ADJUNTO_PARA_LEER * 1024 * 1024:
        return ""
    extension = Path(nombre).suffix.lower()
    if extension == ".pdf":
        return _texto_de_pdf(contenido)
    if extension == ".docx":
        return correo_pro._texto_de_docx_en_memoria(contenido)
    return ""


def piezas_del_correo(correo):
    """
    Prepara, UNA SOLA VEZ, las piezas del correo donde se va a buscar
    todo (el tipo de documento y el radicado/cuenta). Cada pieza es un
    diccionario:

        {"titulo":     el asunto, o el nombre del adjunto
         "encabezado": el principio del texto (VENTANA_ENCABEZADO)
         "texto":      todo el texto disponible de esa pieza
         "adjunto":    el (nombre, contenido) del adjunto, o None si es
                       la pieza del correo mismo}

    La pieza 0 es SIEMPRE el correo (asunto + cuerpo); despues va una
    pieza por adjunto. Se separan asi -- y no todo pegado en un texto
    gigante -- por dos razones:

    * cada adjunto se decide POR SI MISMO (ver tipos_de_pieza), que es lo
      que permite bajar la sentencia de tutela de un correo y dejar
      afuera los otros cuatro anexos que no tienen nada que ver;
    * leer el texto de un PDF es lo mas lento de todo, asi que se hace
      una sola vez aqui y se reutiliza para el tipo y para el radicado.
    """
    piezas = [{
        "titulo": _texto_comparable(correo["asunto"]),
        "encabezado": _texto_comparable(correo["cuerpo"])[:VENTANA_ENCABEZADO],
        "texto": _texto_comparable(correo["asunto"] + " " + correo["cuerpo"]),
        "adjunto": None,
    }]
    for nombre, contenido in correo["adjuntos"]:
        texto = _texto_comparable(_texto_de_adjunto(nombre, contenido))
        piezas.append({
            "titulo": _texto_comparable(nombre),
            "encabezado": texto[:VENTANA_ENCABEZADO],
            "texto": texto,
            "adjunto": (nombre, contenido),
        })
    return piezas


def _menciona(texto: str, frase: str) -> bool:
    """
    True si 'frase' aparece en 'texto' como palabra(s) completa(s), no
    pegada a otras letras. Hace falta para las marcas cortas: sin esto,
    "AUTO" coincidiria dentro de "AUTOMATICO" o "AUTORIZA", y "PODER"
    dentro de "PODERDANTE".
    """
    return re.search(rf"(?<![A-ZÑ]){re.escape(frase)}(?![A-ZÑ])", texto) is not None


def tipos_de_pieza(pieza):
    """
    De que tipo(s) es UNA pieza (el correo, o un adjunto suelto). Vacio =
    de ninguno.

    Es riguroso a proposito -- el documento debe SER una sentencia de
    tutela, un derecho de peticion o un pago oficioso, no basta con que
    lo MENCIONE. Se decide en tres pasos, igual que hace el resto del
    proyecto (ver clasificar_procesos_ejecutivos.es_informacion_no_procesal):

      1. la frase esta en el TITULO (el asunto del correo, o el nombre
         del archivo: "RESPUESTA DERECHO DE PETICION.pdf" no deja lugar
         a dudas) -> SI, sin mirar nada mas;
      2. el titulo dice que es OTRA COSA (ver PALABRAS_QUE_DESCARTAN) ->
         NO, aunque el contenido mencione la frase;
      3. sin marcas en el titulo: la frase tiene que estar en el
         ENCABEZADO del texto -- los primeros VENTANA_ENCABEZADO
         caracteres, donde un documento judicial se titula a si mismo.

    Los pasos 2 y 3 son los que arreglan el caso que se vio en la
    practica: un correo "Notifica AUTO AVOCA accion de tutela ..." cuyo
    cuerpo explica que el accionante alega la vulneracion de su "derecho
    de peticion". Buscando la frase en todo el texto, ese correo entraba
    entero -- con sus cinco anexos -- como si fuera un derecho de
    peticion. Ahora el paso 2 lo descarta por lo que su propio titulo
    dice que es: un auto.
    """
    encontrados = []
    for tipo, variantes in FRASES_A_BUSCAR.items():
        if any(v in pieza["titulo"] for v in variantes):
            encontrados.append(tipo)
    if encontrados:
        return encontrados

    if any(_menciona(pieza["titulo"], marca) for marca in PALABRAS_QUE_DESCARTAN):
        return []

    for tipo, variantes in FRASES_A_BUSCAR.items():
        if any(v in pieza["encabezado"] for v in variantes):
            encontrados.append(tipo)
    return encontrados


def tipos_del_correo(piezas):
    """
    Los tipos que tiene el correo en conjunto (la union de los de todas
    sus piezas), sin repetir y en el orden de FRASES_A_BUSCAR. Vacio = no
    es de los que buscamos.

    Un mismo correo puede ser de mas de uno (ej. una respuesta que
    contesta un derecho de peticion Y anuncia un pago oficioso); en ese
    caso el nombre del archivo los lleva a los dos, separados por " + ",
    para no guardar el mismo documento dos veces.
    """
    encontrados = []
    for pieza in piezas:
        for tipo in tipos_de_pieza(pieza):
            if tipo not in encontrados:
                encontrados.append(tipo)
    return [t for t in FRASES_A_BUSCAR if t in encontrados]


# ==================== Radicado / cuenta / asunto ====================


# El numero de cuenta SOLO se acepta si el texto lo rotula como tal.
# Sin el rotulo, cualquier numero de 4 a 12 digitos del correo (un
# telefono, una factura, un numero de PQR) se colaria como "cuenta" y el
# documento quedaria nombrado con un numero que no es -- peor que no
# nombrarlo. Se aceptan los rotulos que se usan en la practica:
# "CUENTA No. 1234567", "CUENTA CONTRATO 1234567", "NIC 1234567",
# "NIU 1234567". Los "(?<!\d)"/"(?!\d)" evitan cortar un numero mas
# largo por la mitad.
_PATRON_CUENTA = re.compile(
    r"\b(?:NIC|NIU|CUENTA(?:\s+CONTRATO)?|CONTRATO)\b"
    r"[\s:.\-]*(?:N[O°º]?\.?|NRO\.?|NUM(?:ERO)?\.?|#)?[\s:.\-]*"
    r"(?<!\d)(\d{4,12})(?!\d)"
)


def _radicado_en(texto: str):
    """El primer radicado de 23 digitos del texto (plano o con guiones/puntos/espacios), o None."""
    radicados = cruce_excel._radicados_en_texto(texto)
    return sorted(radicados)[0] if radicados else None


def _cuenta_en(texto: str):
    """El primer numero rotulado como cuenta/NIC/NIU en el texto, o None -- ver _PATRON_CUENTA."""
    coincidencia = _PATRON_CUENTA.search(texto)
    if not coincidencia:
        return None
    cuenta = coincidencia.group(1)
    # Una "cuenta" de puros ceros (o demasiado corta) no identifica nada
    # -- misma regla que usa el resto del proyecto para no arrastrar
    # falsos positivos.
    return cuenta if buscador._cuenta_es_valida_para_buscar(cuenta) else None


def identificador_del_correo(correo, piezas):
    """
    Con que se va a nombrar este correo. Devuelve (identificador, origen),
    donde 'origen' dice de donde salio ("radicado", "cuenta" o "asunto")
    -- queda anotado en el CSV para poder revisar despues los que
    quedaron nombrados por asunto.

    Se prueba primero el RADICADO en TODAS las piezas y solo si no hay
    ninguno se pasa a la CUENTA: el radicado identifica el proceso
    completo y es lo que usa el resto del proyecto para nombrar
    carpetas, asi que si el correo lo trae en cualquier parte es
    preferible a una cuenta que aparezca en el asunto.

    Dentro de cada uno se busca en orden de confianza: primero los
    TITULOS (el asunto y los nombres de los archivos, que es donde el
    numero viene puesto a proposito para identificar el envio) y solo
    despues el texto de adentro, donde el numero puede ser la referencia
    a OTRO proceso ("en relacion con el radicado ...").
    """
    titulos = [p["titulo"] for p in piezas]
    textos = [p["texto"] for p in piezas]

    for texto in titulos + textos:
        radicado = _radicado_en(texto)
        if radicado:
            return radicado, "radicado"

    for texto in titulos + textos:
        cuenta = _cuenta_en(texto)
        if cuenta:
            return cuenta, "cuenta"

    fecha = correo["fecha"].isoformat() if correo["fecha"] else "sin fecha"
    asunto = (correo["asunto"] or "").strip()[:MAX_CARACTERES_ASUNTO].strip()
    return (f"{fecha} {asunto}".strip() if asunto else fecha), "asunto"


# ==================== Que adjuntos se guardan ====================


def _archivos_de_zip(contenido: bytes):
    """
    Los PDF/DOCX de adentro de un .zip adjunto, como [(nombre, bytes), ...].
    Un zip dañado no tumba nada: se devuelve lo que se alcance a sacar.
    """
    sacados = []
    try:
        with zipfile.ZipFile(io.BytesIO(contenido)) as archivo_zip:
            for info in archivo_zip.infolist():
                if info.is_dir() or Path(info.filename).suffix.lower() not in (".pdf", ".docx"):
                    continue
                try:
                    with archivo_zip.open(info) as origen:
                        sacados.append((organizador.sanear_nombre(Path(info.filename).name), origen.read()))
                except Exception:
                    continue
    except (zipfile.BadZipFile, OSError):
        pass
    return sacados


def _adjuntos_que_corresponden(piezas):
    """
    Los adjuntos que SON de los tipos buscados, decidido documento por
    documento (ver tipos_de_pieza). Si ninguno lo es por si mismo,
    devuelve None -- que significa "el que coincidio fue el correo, no un
    adjunto en particular".

    Esta es la diferencia entre bajar un archivo o bajar cinco: un correo
    de notificacion trae la sentencia de tutela Y el auto, la constancia,
    el poder y el pantallazo del envio. Solo el primero es lo que se
    pidio; los demas se dejan en el correo.
    """
    del_tipo = [p["adjunto"] for p in piezas if p["adjunto"] is not None and tipos_de_pieza(p)]
    return del_tipo or None


def adjuntos_a_guardar(correo, adjuntos=None):
    """
    Los adjuntos que de verdad hay que guardar: se descartan los tipos
    que no son documentos (EXTENSIONES_A_GUARDAR) y las imagenes
    demasiado livianas (el logo y la firma del pie del correo, ver
    MIN_KB_IMAGEN), y se abren los .zip para sacar los PDF/DOCX de
    adentro.

    'adjuntos' permite pasar solo algunos (los que corresponden al tipo
    buscado); por defecto se revisan todos los del correo.
    """
    utiles = []
    for nombre, contenido in (correo["adjuntos"] if adjuntos is None else adjuntos):
        extension = Path(nombre).suffix.lower()
        if extension == ".zip":
            utiles.extend(_archivos_de_zip(contenido))
            continue
        if extension not in EXTENSIONES_A_GUARDAR:
            continue
        es_imagen = extension in (".jpg", ".jpeg", ".png", ".tif", ".tiff")
        if es_imagen and len(contenido) < MIN_KB_IMAGEN * 1024:
            continue
        utiles.append((nombre, contenido))
    return utiles


# ==================== Nombrar y guardar ====================


def nombre_de_archivo(identificador: str, tipos, nombre_original: str) -> str:
    """
    Arma el nombre final:  "<identificador> - <tipo> - <nombre original>.pdf"

    Si no cabe en MAX_CARACTERES_NOMBRE se recorta, pero NUNCA a costa
    de lo que sirve para reconocer el documento:

      * primero se recorta el nombre original del adjunto (lo que menos
        importa: es el nombre con el que lo mando el remitente);
      * si aun asi no cabe, se recorta el IDENTIFICADOR -- pero solo
        pasa cuando viene de un asunto larguisimo; un radicado o una
        cuenta siempre caben enteros;
      * el TIPO se conserva completo pase lo que pase (es lo que
        permite ordenar la carpeta por tipo de documento), y la
        EXTENSION tambien: un ".pdf" recortado a ".pd" seria un archivo
        que Windows ya no sabe abrir.

    Al recortar se quitan los espacios y guiones que queden colgando al
    final, para que no salgan nombres como "... ASUNTO -.pdf".
    """
    extension = Path(nombre_original).suffix.lower()
    etiqueta_tipo = " + ".join(tipos)
    espacio_util = MAX_CARACTERES_NOMBRE - len(extension)

    espacio_identificador = espacio_util - len(" - ") - len(etiqueta_tipo)
    if espacio_identificador > 0:
        identificador = identificador[:espacio_identificador].strip(" -.")
    nombre = f"{identificador} - {etiqueta_tipo}" if identificador else etiqueta_tipo

    if INCLUIR_NOMBRE_ORIGINAL:
        original = Path(nombre_original).stem.strip()
        # Si el propio nombre del adjunto ya dice el tipo ("SENTENCIA DE
        # TUTELA primera instancia.pdf"), no se repite: quedaria
        # "... - SENTENCIA DE TUTELA - SENTENCIA DE TUTELA primera
        # instancia.pdf".
        if etiqueta_tipo in _texto_comparable(original):
            nombre = identificador or etiqueta_tipo
        espacio_libre = espacio_util - len(nombre) - len(" - ")
        if original and espacio_libre > 0:
            nombre = f"{nombre} - {original[:espacio_libre].strip()}"

    return organizador.sanear_nombre(nombre[:espacio_util].strip(" -.") + extension)


def _crear_carpeta_descargas():
    Path(CARPETA_DESCARGAS).mkdir(parents=True, exist_ok=True)


def guardar_correo(correo, tipos, identificador, archivos, filas_reporte) -> int:
    """
    Guarda en CARPETA_DESCARGAS los archivos de UN correo, ya
    renombrados. Devuelve cuantos se guardaron.

    Si el correo no trae adjuntos utiles (solo pasa con
    EXIGIR_ADJUNTOS = False) se guarda su asunto y su cuerpo como un
    .txt, para no perder la informacion.

    Un archivo que falle al escribirse (ruta imposible, disco lleno,
    antivirus) se registra y se sigue con los demas -- un adjunto
    problematico no puede hacer que se pierda el resto del correo.
    """
    fecha_texto = correo["fecha"].isoformat() if correo["fecha"] else ""
    asunto = correo["asunto"] or "(sin asunto)"
    carpeta = Path(CARPETA_DESCARGAS)
    guardados = 0

    if not archivos:
        contenido = f"Asunto: {asunto}\nFecha: {fecha_texto}\n\n{correo['cuerpo']}".encode("utf-8")
        archivos = [("correo.txt", contenido)]

    for nombre_original, contenido in archivos:
        nombre_final = nombre_de_archivo(identificador, tipos, nombre_original)
        if MODO_PRUEBA:
            logging.info("[SIMULACION] '%s' -> se guardaria como '%s'.", nombre_original, nombre_final)
            filas_reporte.append(
                (nombre_final, identificador, " + ".join(tipos), fecha_texto, asunto, nombre_original)
            )
            guardados += 1
            continue
        try:
            ruta = buscador._ruta_archivo_libre(carpeta, nombre_final)
            with open(organizador._ruta_larga_segura(str(ruta)), "wb") as f:
                f.write(contenido)
        except OSError as error:
            logging.warning(
                "   No se pudo guardar '%s' del correo '%s' (%s) -- se sigue con los demas adjuntos.",
                nombre_final, asunto, error,
            )
            continue
        filas_reporte.append(
            (ruta.name, identificador, " + ".join(tipos), fecha_texto, asunto, nombre_original)
        )
        guardados += 1
    return guardados


# ==================== Revisar un correo ====================


def revisar_correo(correo, resumen, filas_reporte):
    """Decide si UN correo se descarga y, si si, lo guarda ya renombrado."""
    asunto = correo["asunto"] or "(sin asunto)"

    # El orden importa: si se exigen adjuntos y no los trae, se descarta
    # ANTES de leer el texto de nada (que es lo lento).
    if EXIGIR_ADJUNTOS and not correo["adjuntos"]:
        resumen["sin_adjuntos"] += 1
        return

    piezas = piezas_del_correo(correo)
    tipos = tipos_del_correo(piezas)
    if not tipos:
        resumen["descartados_por_tipo"] += 1
        return

    # Si hay adjuntos que SON del tipo buscado, se bajan solo esos. Si el
    # que coincidio fue el correo mismo (por su asunto), no hay como
    # saber cual de sus adjuntos es "el bueno" y se bajan todos.
    del_tipo = _adjuntos_que_corresponden(piezas) if SOLO_EL_DOCUMENTO_QUE_COINCIDE else None
    archivos = adjuntos_a_guardar(correo, del_tipo)
    if del_tipo is not None and len(del_tipo) < len(correo["adjuntos"]):
        resumen["anexos_omitidos"] += len(correo["adjuntos"]) - len(del_tipo)

    if EXIGIR_ADJUNTOS and not archivos:
        # Traia adjuntos, pero ninguno era un documento (solo el logo
        # del pie de firma, un .p7s, etc).
        resumen["sin_adjuntos"] += 1
        logging.info("[Omitido] '%s' es %s pero no trae ningun documento adjunto.", asunto, " + ".join(tipos))
        return

    identificador, origen = identificador_del_correo(correo, piezas)
    guardados = guardar_correo(correo, tipos, identificador, archivos, filas_reporte)

    resumen["descargados"] += 1
    resumen["archivos"] += guardados
    resumen[f"por_{origen}"] += 1
    logging.info(
        "[%s] '%s' (%s) -> %d archivo(s) como '%s' (%s).",
        "SIMULACION" if MODO_PRUEBA else "Descargado",
        asunto, " + ".join(tipos), guardados, identificador, origen,
    )


# ==================== Gmail ====================


def _buscar_uids(mail, *criterios):
    """Lanza UNA busqueda contra Gmail y devuelve la lista de UIDs (vacia si no hubo resultados)."""
    typ, datos = correo_pro._con_limite_de_tiempo(
        mail, mail.uid, "SEARCH", None, *criterios, timeout=TIMEOUT_SEGUNDOS
    )
    if typ != "OK" or not datos or not datos[0]:
        return []
    return [int(u) for u in datos[0].split() if u.isdigit()]


def _consulta_exacta_de_gmail() -> str:
    """
    La consulta para Gmail con las FRASES EXACTAS, tal como se escribiria
    en su barra de busqueda:

        has:attachment ("pago oficioso" OR "derecho de peticion" OR ...)

    Las comillas son las que convierten la busqueda en una FRASE: sin
    ellas Gmail busca los correos que tengan esas palabras en cualquier
    parte y por separado, que es de donde salian los cientos de
    candidatos que no eran nada.

    Va envuelta en comillas y con las de adentro escapadas, que es como
    IMAP espera recibir un texto con comillas.
    """
    frases = [v.lower() for variantes in FRASES_A_BUSCAR.values() for v in variantes]
    consulta = " OR ".join(f'"{frase}"' for frase in frases)
    if EXIGIR_ADJUNTOS:
        consulta = f"has:attachment ({consulta})"
    return '"' + consulta.replace("\\", "\\\\").replace('"', '\\"') + '"'


def listar_uids(mail):
    """
    Los UIDs de los correos candidatos, dentro de la ventana de
    DIAS_HACIA_ATRAS.

    Con BUSQUEDA_EXACTA_EN_GMAIL (por defecto) es UNA sola consulta, con
    las frases exactas y solo correos con adjuntos -- ver
    _consulta_exacta_de_gmail. Usa X-GM-RAW, que es la busqueda propia de
    Gmail (la misma de su barra de busqueda) y que el proyecto ya usa en
    otros scripts.

    Si esa consulta falla o no devuelve NADA, se cae automaticamente a la
    busqueda amplia (SUBJECT/TEXT por trozos de palabra, lo mas basico
    del protocolo, que funciona en cualquier servidor). Es a proposito:
    entre revisar de mas y no revisar nada, se prefiere revisar de mas.
    """
    desde = (datetime.date.today() - datetime.timedelta(days=DIAS_HACIA_ATRAS)).strftime("%d-%b-%Y")

    if BUSQUEDA_EXACTA_EN_GMAIL:
        try:
            encontrados = _buscar_uids(mail, "SINCE", desde, "X-GM-RAW", _consulta_exacta_de_gmail())
            if encontrados:
                logging.info(
                    "[Correo] Busqueda exacta%s: %d correo(s) candidatos.",
                    " (solo con adjuntos)" if EXIGIR_ADJUNTOS else "", len(encontrados),
                )
                return sorted(encontrados)
            logging.warning(
                "[Correo] La busqueda exacta no devolvio ningun correo -- se prueba con la busqueda amplia "
                "por si acaso (mas lenta)."
            )
        except Exception as error:
            logging.warning(
                "[Correo] Fallo la busqueda exacta (%s) -- se sigue con la busqueda amplia (mas lenta).", error,
            )

    encontrados = set()
    for campo in ("SUBJECT", "TEXT"):
        for frase in FRASES_PARA_PEDIR_A_GMAIL:
            try:
                nuevos = _buscar_uids(mail, "SINCE", desde, campo, frase)
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


# ==================== Principal ====================


def procesar():
    credenciales = organizador.leer_credenciales()
    if not credenciales:
        logging.error(
            "No hay %s (o le faltan datos) -- sin las credenciales de Gmail no se puede revisar el correo. "
            "Haz una copia de credenciales_sgde.example.txt, renombrala a credenciales_sgde.txt y pon ahi tu "
            "correo y tu contraseña de aplicacion de Gmail (paso 3 de 'COMO USAR - descargar correos.txt').",
            organizador.ARCHIVO_CREDENCIALES,
        )
        return

    if not MODO_PRUEBA:
        try:
            _crear_carpeta_descargas()
        except OSError as error:
            logging.error(
                "No se pudo crear la carpeta de descargas %r (%s). Revisa la ruta en CARPETA_DESCARGAS al "
                "inicio de este script.", CARPETA_DESCARGAS, error,
            )
            return
    logging.info("Todo se va a guardar en: %s", CARPETA_DESCARGAS)
    logging.info(
        "Se buscan correos de: %s (ultimos %d dia(s)%s).",
        ", ".join(FRASES_A_BUSCAR), DIAS_HACIA_ATRAS,
        ", solo los que traigan adjuntos" if EXIGIR_ADJUNTOS else "",
    )

    mail, uidvalidity = correo_pro.conectar(*credenciales, timeout=TIMEOUT_SEGUNDOS)
    if mail is None:
        return

    resumen = {
        "descargados": 0, "archivos": 0, "descartados_por_tipo": 0, "sin_adjuntos": 0,
        "anexos_omitidos": 0, "fallidos": 0, "por_radicado": 0, "por_cuenta": 0, "por_asunto": 0,
    }
    filas_reporte = []
    revisados = correo_pro.cargar_progreso(uidvalidity, ARCHIVO_PROGRESO)

    try:
        logging.info(
            "[Correo] Pidiendo a Gmail los correos candidatos (%s)...",
            "busqueda exacta, 1 consulta" if BUSQUEDA_EXACTA_EN_GMAIL
            else f"busqueda amplia, {len(FRASES_PARA_PEDIR_A_GMAIL) * 2} consultas",
        )
        try:
            uids = listar_uids(mail)
        except Exception as error:
            logging.error("[Correo] No se pudo obtener la lista de correos: %s", error)
            return

        pendientes = [u for u in uids if u not in revisados]
        # El UID crece con el tiempo: ordenar al reves = del mas nuevo
        # al mas viejo.
        if EMPEZAR_POR_LOS_MAS_NUEVOS:
            pendientes.sort(reverse=True)
        logging.info(
            "[Correo] %d correo(s) candidatos; %d ya revisados en corridas anteriores; quedan %d por revisar.",
            len(uids), len(uids) - len(pendientes), len(pendientes),
        )
        if not pendientes:
            logging.info("[Correo] No hay nada nuevo que descargar. Todo al dia.")
            return

        de_esta_corrida = pendientes[:MAX_CORREOS_POR_CORRIDA]
        if len(pendientes) > len(de_esta_corrida):
            logging.info(
                "[Correo] Esta corrida revisa hasta %d correo(s); quedan %d para las proximas. Vuelve a correr "
                "el script las veces que haga falta: siempre sigue donde se quedo.",
                len(de_esta_corrida), len(pendientes) - len(de_esta_corrida),
            )

        reconexiones = 0
        procesados = 0
        comenzo_en = time.monotonic()

        for inicio in range(0, len(de_esta_corrida), TAMANO_LOTE_DESCARGA):
            # El tope por tiempo se revisa ANTES de empezar cada grupo,
            # nunca a mitad de uno, para cerrar en un punto limpio. El
            # PRIMER grupo siempre se procesa (inicio > 0): si no, un
            # tope demasiado bajo haria que cada corrida terminara sin
            # revisar ni un correo y nunca se avanzaria.
            if MAX_MINUTOS_POR_CORRIDA and inicio > 0:
                if (time.monotonic() - comenzo_en) / 60 >= MAX_MINUTOS_POR_CORRIDA:
                    logging.info(
                        "[Correo] Se cumplieron los %d minuto(s) de esta corrida -- se cierra aqui con %d "
                        "correo(s) revisados. Vuelve a correr el script para seguir donde se quedo.",
                        MAX_MINUTOS_POR_CORRIDA, procesados,
                    )
                    break

            lote = de_esta_corrida[inicio:inicio + TAMANO_LOTE_DESCARGA]
            intentos = 0
            mensajes = None
            lote_abandonado = False

            while mensajes is None:
                try:
                    mensajes = correo_pro.descargar_lote(mail, lote, timeout=TIMEOUT_SEGUNDOS)
                except (imaplib.IMAP4.abort, imaplib.IMAP4.error, OSError) as error:
                    intentos += 1
                    if intentos > MAX_REINTENTOS_POR_LOTE:
                        logging.error(
                            "[Correo] Un grupo de %d correo(s) sigue fallando despues de %d intento(s) (%s) -- "
                            "se salta y se sigue. NO se dan por revisados: quedan pendientes para la proxima.",
                            len(lote), intentos, error,
                        )
                        resumen["fallidos"] += len(lote)
                        lote_abandonado = True
                        break
                    if reconexiones >= MAX_RECONEXIONES:
                        logging.error(
                            "[Correo] Se perdio la conexion (%s) y ya se reconecto %d vez/veces -- se detiene "
                            "aqui. Lo descargado queda guardado; vuelve a correr el script para seguir.",
                            error, reconexiones,
                        )
                        return
                    reconexiones += 1
                    logging.warning(
                        "[Correo] Se perdio la conexion con Gmail (%s) -- reconectando (intento %d/%d)...",
                        error, reconexiones, MAX_RECONEXIONES,
                    )
                    correo_pro.cerrar(mail)
                    mail, _uidvalidity_nuevo = correo_pro._conectar_o_none(credenciales, uidvalidity, timeout=TIMEOUT_SEGUNDOS)
                    if mail is None:
                        return

            if lote_abandonado:
                # La conexion acaba de fallar varias veces seguidas: si
                # el corte fue a mitad de una descarga quedan datos a
                # medio leer, y la SIGUIENTE orden los leeria como si
                # fueran la respuesta del servidor. Se abre una nueva.
                logging.info("[Correo] Se descarta la conexion (quedo en mal estado) y se abre una nueva...")
                correo_pro.cerrar(mail)
                mail, _uidvalidity_nuevo = correo_pro._conectar_o_none(credenciales, uidvalidity, timeout=TIMEOUT_SEGUNDOS)
                if mail is None:
                    return
                continue

            for uid in lote:
                mensaje = mensajes.get(uid)
                if mensaje is None:
                    # No llego en la respuesta (borrado, movido o
                    # ilegible). Se marca como revisado igual: si no,
                    # cada corrida volveria a intentarlo para siempre.
                    revisados.add(uid)
                    continue
                try:
                    correo = correo_pro.leer_correo(mensaje)
                    revisar_correo(correo, resumen, filas_reporte)
                except Exception as error:
                    # Un correo problematico se registra y se sigue --
                    # jamas puede tumbar la corrida completa.
                    resumen["fallidos"] += 1
                    logging.warning("   [Correo] No se pudo procesar el correo UID %s: %s", uid, error)
                revisados.add(uid)
                procesados += 1

                if procesados % GUARDAR_PROGRESO_CADA_N == 0:
                    correo_pro.guardar_progreso(revisados, uidvalidity, ARCHIVO_PROGRESO)
                if procesados % AVISO_PROGRESO_CADA_N == 0:
                    logging.info(
                        "[Correo] ...van %d/%d correo(s) revisados en esta corrida (%d descargado(s), "
                        "%d archivo(s))...",
                        procesados, len(de_esta_corrida), resumen["descargados"], resumen["archivos"],
                    )
    finally:
        correo_pro.guardar_progreso(revisados, uidvalidity, ARCHIVO_PROGRESO)
        correo_pro.cerrar(mail)

    logging.info(
        "Listo: %d correo(s) descargado(s), %d archivo(s) guardado(s) en '%s'. "
        "%d correo(s) no eran de los tipos buscados, %d sin documentos adjuntos, %d con error.",
        resumen["descargados"], resumen["archivos"], CARPETA_DESCARGAS,
        resumen["descartados_por_tipo"], resumen["sin_adjuntos"], resumen["fallidos"],
    )
    logging.info(
        "Nombrados por radicado: %d, por numero de cuenta: %d, por asunto (no traian ni radicado ni cuenta): %d.",
        resumen["por_radicado"], resumen["por_cuenta"], resumen["por_asunto"],
    )
    if resumen["anexos_omitidos"]:
        logging.info(
            "%d adjunto(s) se dejaron sin bajar por no ser del tipo buscado (eran anexos del correo: "
            "constancias, poderes, pantallazos). Si los quieres todos, pon "
            "SOLO_EL_DOCUMENTO_QUE_COINCIDE = False.", resumen["anexos_omitidos"],
        )

    if filas_reporte and not MODO_PRUEBA:
        ruta_reporte = os.path.join(CARPETA_DESCARGAS, NOMBRE_REPORTE)
        if cruce_excel._escribir_csv_tolerante(
            ruta_reporte,
            ["Archivo guardado", "Identificador", "Tipo", "Fecha del correo", "Asunto", "Adjunto original"],
            filas_reporte, "Correos descargados",
        ):
            logging.info("[Reporte] %d fila(s) guardadas en: %s", len(filas_reporte), ruta_reporte)

    if MODO_PRUEBA:
        logging.info(
            "MODO_PRUEBA esta activo: NO se descargo nada, solo se mostro que se haria y con que nombre. "
            "Revisa el log y, si se ve bien, cambia MODO_PRUEBA = False al inicio de este script."
        )


def main():
    configurar_logging()
    procesar()


if __name__ == "__main__":
    main()
