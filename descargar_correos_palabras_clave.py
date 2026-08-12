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
dicen "fallo de tutela"). La frase se busca en el ASUNTO, en el CUERPO,
en el NOMBRE de cada adjunto y -- lo importante -- en el TEXTO DE
ADENTRO de cada adjunto PDF/DOCX: un correo con asunto generico
("Notificacion 12345") pero con "FALLO DE TUTELA.pdf" adjunto SI es de
los que buscamos, y mirando solo el asunto se perderia.

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

POR QUE NO SE TRABA
============================================================
Reutiliza tal cual la conexion con Gmail de revisar_correo_pro.py, que
ya trae resuelto todo lo que fallaba en la practica: limite de tiempo
DURO por operacion (para equipos con antivirus que inspecciona el
correo), reconexion automatica, descarga por lotes, UIDs en vez de
numeros de orden, y tope por cantidad y por tiempo para que la corrida
SIEMPRE termine. Ademas lleva su propio archivo de progreso
(ARCHIVO_PROGRESO): si lo cortas o se cae la conexion, la proxima
corrida SIGUE DONDE SE QUEDO en vez de volver a bajar todo.

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

# Carpeta UNICA donde se guarda todo lo descargado. Si no existe, se
# crea. Cambiala por la que quieras.
CARPETA_DESCARGAS = os.path.join(
    os.path.dirname(base.CARPETA_PROCESOS) or os.path.dirname(__file__),
    "CORREOS DESCARGADOS - PAGO OFICIOSO, PETICION Y TUTELA",
)

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

# True (por defecto): ademas del asunto, el cuerpo y el nombre de los
# adjuntos, se LEE EL TEXTO de adentro de los PDF/DOCX adjuntos, tanto
# para reconocer el tipo de documento como para buscarle el radicado o
# la cuenta. Es lo mas lento de todo, pero es lo que hace que funcione:
# el radicado casi nunca esta en el asunto, esta impreso dentro del PDF.
LEER_TEXTO_DE_ADJUNTOS = True

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

# Cuantos correos se bajan por cada peticion a Gmail. Pedir demasiados
# de golpe hace que la peticion tarde mas del limite y la conexion se
# caiga; 5 es un punto medio seguro (ver revisar_correo_pro.py).
TAMANO_LOTE_DESCARGA = 5

# Los limites de red (tiempo de espera, reintentos, reconexiones) se
# reutilizan de revisar_correo_pro.py -- es el mismo Gmail y ya estan
# ajustados a lo que aguanta en la practica. Si alguna vez hay que
# tocarlos, se tocan alla y los dos scripts quedan iguales.
TIMEOUT_SEGUNDOS = correo_pro.TIMEOUT_SEGUNDOS
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

# Lo que se le pide a GMAIL para acotar la busqueda del lado del
# servidor, antes de bajar nada. A proposito son TROZOS de palabra y
# sin tildes (la busqueda de IMAP es por subcadena y en ASCII):
#   "OFICIOSO" -> PAGO OFICIOSO, PAGOS OFICIOSOS, pago de manera oficiosa
#   "PETICI"   -> PETICION, PETICIÓN, PETICIONES
#   "TUTELA"   -> SENTENCIA DE TUTELA, FALLO DE TUTELA
# Son 3 frases x (asunto + texto) = 6 consultas a Gmail en total, no una
# por proceso. Bajar los ~25.000 correos de dos años completos, en
# cambio, satura la conexion -- por eso se acota aqui.
#
# Estas frases solo ACOTAN: sobre lo que llegue se vuelve a revisar en
# tu PC con FRASES_A_BUSCAR (que es mas estricto), asi que un correo que
# solo diga "TUTELA" de pasada igual se descarta despues.
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

# Cada cuantos correos se guarda el progreso / se avisa el avance.
GUARDAR_PROGRESO_CADA_N = 25
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
        return correo_pro._texto_de_pdf_en_memoria(contenido)
    if extension == ".docx":
        return correo_pro._texto_de_docx_en_memoria(contenido)
    return ""


def piezas_del_correo(correo):
    """
    Prepara, UNA SOLA VEZ, los textos donde se va a buscar todo (el tipo
    de documento y el radicado/cuenta). Devuelve una lista de textos ya
    normalizados, EN ORDEN DE CONFIANZA para identificar el correo:

      1. el asunto
      2. el nombre de cada adjunto
      3. el texto de adentro de cada adjunto PDF/DOCX
      4. el cuerpo del correo

    El orden importa para el IDENTIFICADOR (ver identificador_del_correo):
    el radicado que viene en el asunto o en el nombre del archivo es mas
    de fiar que uno que aparezca perdido en el cuerpo, donde puede ser
    la referencia a OTRO proceso ("en relacion con el radicado ...").

    Se devuelven por separado -- y no todo pegado -- porque leer el
    texto de un PDF es lo mas lento de todo y asi se hace una sola vez
    para las dos revisiones.
    """
    piezas = [_texto_comparable(correo["asunto"])]
    textos_de_adjuntos = []
    for nombre, contenido in correo["adjuntos"]:
        piezas.append(_texto_comparable(nombre))
        texto = _texto_de_adjunto(nombre, contenido)
        if texto:
            textos_de_adjuntos.append(_texto_comparable(texto))
    piezas.extend(textos_de_adjuntos)
    piezas.append(_texto_comparable(correo["cuerpo"]))
    return piezas


def tipos_del_correo(piezas):
    """
    Los tipos (de FRASES_A_BUSCAR) que este correo tiene, sin repetir y
    en el orden en que estan definidos. Vacio = no es de los que
    buscamos.

    Un mismo correo puede ser de mas de uno (ej. una respuesta que
    contesta un derecho de peticion Y anuncia un pago oficioso); en ese
    caso el nombre del archivo los lleva a los dos, separados por " + ",
    para no tener que guardar el mismo documento dos veces.
    """
    encontrados = []
    for tipo, variantes in FRASES_A_BUSCAR.items():
        if any(any(v in pieza for v in variantes) for pieza in piezas):
            encontrados.append(tipo)
    return encontrados


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

    Se prueba primero el RADICADO en TODAS las piezas (asunto, nombres
    de adjunto, contenido, cuerpo) y solo si no hay ninguno se pasa a la
    CUENTA: el radicado identifica el proceso completo y es lo que usa
    el resto del proyecto para nombrar carpetas, asi que si el correo lo
    trae en cualquier parte es preferible a una cuenta que aparezca en
    el asunto. Dentro de cada uno se respeta el orden de confianza de
    las piezas (ver piezas_del_correo).
    """
    for pieza in piezas:
        radicado = _radicado_en(pieza)
        if radicado:
            return radicado, "radicado"

    for pieza in piezas:
        cuenta = _cuenta_en(pieza)
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


def adjuntos_a_guardar(correo):
    """
    Los adjuntos que de verdad hay que guardar: se descartan los tipos
    que no son documentos (EXTENSIONES_A_GUARDAR) y las imagenes
    demasiado livianas (el logo y la firma del pie del correo, ver
    MIN_KB_IMAGEN), y se abren los .zip para sacar los PDF/DOCX de
    adentro.
    """
    utiles = []
    for nombre, contenido in correo["adjuntos"]:
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

    archivos = adjuntos_a_guardar(correo)
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


def listar_uids(mail):
    """
    Los UIDs de los correos candidatos, en pocas consultas: cada frase de
    FRASES_PARA_PEDIR_A_GMAIL por ASUNTO y por TEXTO, dentro de la
    ventana de DIAS_HACIA_ATRAS. Son 6 consultas en total.

    Cada consulta es lo mas basico del protocolo (SINCE + SUBJECT/TEXT,
    en ASCII puro, sin extensiones de Gmail ni OR): es lo unico que
    aguanta cualquier servidor sin cortar. Si una falla, se registra y
    se sigue con las demas -- es preferible revisar de menos que no
    revisar nada.
    """
    desde = (datetime.date.today() - datetime.timedelta(days=DIAS_HACIA_ATRAS)).strftime("%d-%b-%Y")
    encontrados = set()

    for campo in ("SUBJECT", "TEXT"):
        for frase in FRASES_PARA_PEDIR_A_GMAIL:
            try:
                typ, datos = correo_pro._con_limite_de_tiempo(
                    mail, mail.uid, "SEARCH", None, "SINCE", desde, campo, frase
                )
            except Exception as error:
                logging.warning(
                    "   [Correo] Fallo la busqueda %s '%s' (%s) -- se sigue con las demas.", campo, frase, error,
                )
                continue
            if typ != "OK" or not datos or not datos[0]:
                logging.info("[Correo] %s '%s': 0 correo(s).", campo, frase)
                continue
            nuevos = [int(u) for u in datos[0].split() if u.isdigit()]
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
            "Ver el README.", organizador.ARCHIVO_CREDENCIALES,
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

    mail, uidvalidity = correo_pro.conectar(*credenciales)
    if mail is None:
        return

    resumen = {
        "descargados": 0, "archivos": 0, "descartados_por_tipo": 0, "sin_adjuntos": 0,
        "fallidos": 0, "por_radicado": 0, "por_cuenta": 0, "por_asunto": 0,
    }
    filas_reporte = []
    revisados = correo_pro.cargar_progreso(uidvalidity, ARCHIVO_PROGRESO)

    try:
        logging.info(
            "[Correo] Pidiendo a Gmail los correos candidatos (%d frase(s) x asunto y texto = %d consulta(s))...",
            len(FRASES_PARA_PEDIR_A_GMAIL), len(FRASES_PARA_PEDIR_A_GMAIL) * 2,
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
                    mensajes = correo_pro.descargar_lote(mail, lote)
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
                    mail, _uidvalidity_nuevo = correo_pro._conectar_o_none(credenciales, uidvalidity)
                    if mail is None:
                        return

            if lote_abandonado:
                # La conexion acaba de fallar varias veces seguidas: si
                # el corte fue a mitad de una descarga quedan datos a
                # medio leer, y la SIGUIENTE orden los leeria como si
                # fueran la respuesta del servidor. Se abre una nueva.
                logging.info("[Correo] Se descarta la conexion (quedo en mal estado) y se abre una nueva...")
                correo_pro.cerrar(mail)
                mail, _uidvalidity_nuevo = correo_pro._conectar_o_none(credenciales, uidvalidity)
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
