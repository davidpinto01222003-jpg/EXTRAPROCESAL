"""
Depura un listado de procesos de la Superintendencia de Sociedades
(un Excel con ~130 concursados) contra las carpetas de tu Google
Drive, y genera un Excel nuevo con:

  1. Cuales de esos procesos tienen carpeta en el Drive (la carpeta de
     cada proceso se busca por el NOMBRE DEL CONCURSADO -- y por el NIT,
     si el Excel lo trae).
  2. Cuales de los que estan en el Drive tienen una GARANTIA FNG (Fondo
     Nacional de Garantias): numero de garantia/certificado, obligacion
     o pagare, entidad, valor/cobertura y FECHA DE VENCIMIENTO (y si a
     hoy esta VIGENTE o VENCIDA).
  3. Cuales tienen LEASING (o "arrendamiento financiero"): numero de
     contrato, entidad, bien, FECHA DE INICIO, FECHA DE TERMINACION y
     plazo (si solo hay inicio + plazo, la fecha de terminacion se
     calcula y queda marcada como "estimada").

Para cada proceso revisa TODOS los documentos de su carpeta (y de sus
subcarpetas): PDF, Word (.docx), Excel (.xlsx), texto, correos .eml y
documentos nativos de Google (Docs/Sheets/Slides). Ademas del contenido,
se tiene en cuenta el NOMBRE del archivo/carpeta (ej. "Garantia FNG.pdf").

Las fechas y numeros se sacan automaticamente del texto que rodea cada
mencion de FNG / leasing, asi que pueden quedar datos incompletos o de
mas: cada hallazgo trae el FRAGMENTO de texto donde se encontro y el
ENLACE al documento, para que lo confirmes con un clic.

Los PDF ESCANEADOS (imagen, sin texto) no se pueden leer sin OCR: si su
nombre sugiere que son de FNG/leasing/garantia/contrato quedan listados
en la hoja REVISAR_A_MANO del Excel de salida.

Este script es de SOLO LECTURA: nunca mueve, borra, renombra ni
descarga nada al disco en tu Drive -- los documentos se leen en memoria.
El unico archivo que escribe es el Excel de resultado (y un cache
depurar_fng_leasing_cache.json para que la siguiente corrida no vuelva
a leer los documentos que no han cambiado).

Dos formas de leer el Drive (ver CONFIGURACION):
  - API de Google Drive (por defecto): usa las MISMAS credenciales que
    buscar_faltantes_en_drive.py (credenciales_drive.json /
    token_drive.json). Ver README, "Configurar el acceso a Google Drive".
  - Carpeta local: si tienes "Google Drive para escritorio", puedes
    poner en CARPETA_LOCAL_DRIVE la ruta de la carpeta sincronizada (ej.
    G:\\Mi unidad\\PROCESOS SUPERSOCIEDADES) y no hace falta la API.

Uso:
    python depurar_fng_leasing.py                     (usa RUTA_EXCEL_PROCESOS)
    python depurar_fng_leasing.py "C:\\ruta\\listado.xlsx"
"""

import datetime
import difflib
import email
import io
import json
import logging
import os
import re
import subprocess
import sys
import threading
import time
import unicodedata
import warnings
import webbrowser
from concurrent.futures import ThreadPoolExecutor
from email import policy
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
except ImportError:
    Credentials = None

try:
    from pypdf import PdfReader
except ImportError:
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        PdfReader = None

try:
    import docx
except ImportError:
    docx = None

# ============================= CONFIGURACION =============================

DIRECTORIO = os.path.dirname(os.path.abspath(__file__))

# Excel con el listado de procesos de Supersociedades. Tambien se puede
# pasar como argumento (o arrastrar el archivo sobre depurar_fng_leasing.bat).
# Si queda vacio y no se pasa argumento, usa el unico Excel que haya en
# la carpeta del script; si hay varios, lo pregunta al arrancar.
RUTA_EXCEL_PROCESOS = r""

# Hoja del Excel a leer (None = la primera hoja).
HOJA_EXCEL_PROCESOS = None

# Nombre EXACTO del encabezado de la columna con el nombre del
# concursado. None = se detecta solo (busca encabezados como
# "CONCURSADO", "RAZON SOCIAL", "SOCIEDAD", "DEUDOR", "NOMBRE"...).
COLUMNA_CONCURSADO = None

# Igual, para la columna del NIT (opcional, ayuda a encontrar la carpeta
# si su nombre trae el NIT). None = se detecta sola; si no hay, se omite.
COLUMNA_NIT = None

# Igual, para la columna del numero de EXPEDIENTE de Supersociedades (o
# el RADICADO). Si la carpeta del proceso trae ese numero en el nombre,
# tambien se encuentra por ahi. None = se detecta sola.
COLUMNA_EXPEDIENTE = None

# --- Filtro por TIPO de proceso ---
# Columna que dice el tipo de proceso (None = se detecta sola: "PROCESO",
# "TIPO DE PROCESO", "TIPO"...).
COLUMNA_TIPO_PROCESO = None

# Solo se revisan los procesos cuyo tipo contenga alguna de estas
# palabras (sin importar tildes ni mayusculas: "REORGANIZACION" tambien
# toma "Reorganización abreviada"). Lista vacia [] = revisar TODOS.
TIPOS_PROCESO_A_INCLUIR = [
    "REORGANIZACION",            # incluye REORGANIZACION ABREVIADA
    "INSOLVENCIA",               # incluye INSOLVENCIA IPNNC (persona natural no comerciante)
    "NAR",                       # negociacion de emergencia de acuerdo de reorganizacion
    "NEAR",
    "RECUPERACION EMPRESARIAL",
    "LIQUIDACION",               # liquidacion judicial / simplificada (Ley 1116)
    "VALIDACION",                # validacion judicial de acuerdo extrajudicial
    "ACUERDO DE REESTRUCTURACION",
    "CONCORDATO",
    "LEY 1116",
]

# --- De donde leer el Drive ---
# Si pones aqui una ruta local (carpeta sincronizada con "Google Drive
# para escritorio"), se lee de ahi y NO se usa la API de Google.
CARPETA_LOCAL_DRIVE = r""

# Solo para la API: carpeta de Drive donde estan las carpetas de los
# procesos. Puede ser el NOMBRE de la carpeta, su ID o el enlace
# (https://drive.google.com/drive/folders/XXXX). Vacio = busca en TODO
# el Drive (Mi unidad + compartido conmigo + unidades compartidas).
CARPETA_RAIZ_DRIVE = ""

# Cuantos niveles por debajo de CARPETA_RAIZ_DRIVE se buscan carpetas
# de procesos (1 = solo las carpetas que estan directamente adentro).
PROFUNDIDAD_BUSQUEDA_CARPETAS = 3

CREDENCIALES_DRIVE = os.path.join(DIRECTORIO, "credenciales_drive.json")
TOKEN_DRIVE = os.path.join(DIRECTORIO, "token_drive.json")
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

# "rapido" (recomendado con la API): primero le pregunta al buscador de
# Google Drive que documentos mencionan FNG/leasing, y solo descarga
# esos (mas los que tengan FNG/leasing/garantia/contrato en el nombre).
# Mucho mas rapido. "completo": lee TODOS los documentos de cada carpeta
# (mas lento, pero no depende del indice de Google). Con
# CARPETA_LOCAL_DRIVE siempre se usa "completo".
MODO_LECTURA = "rapido"

# Que tan parecido debe ser el nombre de la carpeta al del concursado
# (0 a 1) cuando no contiene TODAS sus palabras importantes.
UMBRAL_SIMILITUD_NOMBRE = 0.82

# Limites para no quedarse pegado en archivos gigantes.
MAX_MB_ARCHIVO = 40
MAX_PAGINAS_PDF = 120

# Descargas en paralelo (solo API).
NUM_HILOS = 4

ARCHIVO_SALIDA = os.path.join(
    DIRECTORIO, f"depuracion_fng_leasing_{datetime.date.today():%Y-%m-%d}.xlsx"
)
ARCHIVO_CACHE = os.path.join(DIRECTORIO, "depurar_fng_leasing_cache.jsonl")
ARCHIVO_CACHE_VIEJO = os.path.join(DIRECTORIO, "depurar_fng_leasing_cache.json")  # formato anterior

# Avance de la corrida: si se corta (se apaga/suspende el PC, se cierra la
# ventana), al volver a ejecutar se RETOMA desde el proceso donde quedo,
# sin volver a listar el Drive ni revisar los procesos ya terminados. Se
# borra solo cuando la corrida termina completa. Si el avance tiene mas
# de estos dias, se descarta y se empieza de nuevo (para no usar datos
# viejos del Drive).
ARCHIVO_PROGRESO = os.path.join(DIRECTORIO, "depurar_fng_leasing_progreso.json")
DIAS_VALIDEZ_PROGRESO = 3

# Si falla la conexion (ej. el PC se suspendio y al volver no hay red),
# cuantas veces reintentar cada proceso y cuantos segundos esperar entre
# intentos antes de darlo por fallido.
REINTENTOS_POR_PROCESO = 4
ESPERA_REINTENTO_SEG = 30
ARCHIVO_LOG = os.path.join(DIRECTORIO, "depurar_fng_leasing.log")
ARCHIVO_ENLACE_AUTORIZACION = os.path.join(DIRECTORIO, "enlace_autorizacion_google.txt")

# ===========================================================================

HOY = datetime.date.today()

MIME_CARPETA = "application/vnd.google-apps.folder"
MIME_ATAJO = "application/vnd.google-apps.shortcut"
MIME_EXPORTAR_TEXTO = {
    "application/vnd.google-apps.document": "text/plain",
    "application/vnd.google-apps.spreadsheet": "text/csv",
    "application/vnd.google-apps.presentation": "text/plain",
}
EXTENSIONES_LEIBLES = {".pdf", ".docx", ".xlsx", ".xlsm", ".txt", ".csv", ".eml", ".html", ".htm", ".xml", ".json"}

ENCABEZADOS_CONCURSADO = [
    "concursado", "nombre del concursado", "razon social", "nombre o razon social",
    "sociedad", "deudor", "empresa", "nombre de la sociedad", "nombre", "cliente", "demandante",
]
ENCABEZADOS_NIT = ["nit", "nit concursado", "nit del concursado", "identificacion", "documento"]
ENCABEZADOS_EXPEDIENTE = ["expediente", "no expediente", "numero de expediente", "radicado", "rad"]
ENCABEZADOS_TIPO = ["tipo de proceso", "tipo proceso", "proceso", "tipo"]

# Palabras que NO sirven para identificar a un concursado (tipo de
# sociedad, etapa del proceso, conectores) -- se ignoran al comparar el
# nombre del Excel contra el nombre de la carpeta.
PALABRAS_IGNORADAS = {
    "S", "A", "SA", "SAS", "LTDA", "LIMITADA", "EU", "ESP", "CI", "BIC", "SCA", "SC", "EN", "CIA", "Y", "E",
    "DE", "DEL", "LA", "EL", "LOS", "LAS", "COMPANIA", "SOCIEDAD", "POR", "ACCIONES", "SIMPLIFICADA",
    "ANONIMA", "REORGANIZACION", "LIQUIDACION", "JUDICIAL", "VALIDACION", "ACUERDO", "PROCESO", "CONCURSO",
    "INSOLVENCIA", "CONCURSADO", "SUPERSOCIEDADES", "SUPERINTENDENCIA", "SOCIEDADES", "EXPEDIENTE",
    "LEY", "1116", "PERSONA", "NATURAL", "COMERCIANTE",
}

# Para reconocer la entidad acreedora/arrendadora cerca de la mencion.
ENTIDADES = [
    "BANCOLOMBIA", "LEASING BANCOLOMBIA", "BANCO DE BOGOTA", "BANCO DE OCCIDENTE", "BANCO POPULAR",
    "BANCO AV VILLAS", "AV VILLAS", "DAVIVIENDA", "BBVA", "ITAU", "SCOTIABANK", "COLPATRIA",
    "BANCO CAJA SOCIAL", "BANCO AGRARIO", "BANCOOMEVA", "BANCO GNB", "SUDAMERIS", "BANCO PICHINCHA",
    "BANCO FINANDINA", "FINANDINA", "BANCO W", "BANCAMIA", "MIBANCO", "BANCO FALABELLA",
    "BANCO SERFINANZA", "SERFINANZA", "COLTEFINANCIERA", "CREDIFINANCIERA", "BANCO SANTANDER",
    "BANCO MUNDO MUJER", "BANCO UNION", "GIROS Y FINANZAS", "BANCOLDEX", "FINAGRO", "FINDETER",
    "LEASING DE OCCIDENTE", "LEASING BOLIVAR", "LEASING CORFICOLOMBIANA", "CORFICOLOMBIANA",
    "BANCO COOPERATIVO COOPCENTRAL", "COOPCENTRAL", "CONFIAR", "COOFINEP", "LULO BANK", "BANCO CONTACTAR",
    "CAJA SOCIAL", "OCCIDENTE",
]

MESES = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6, "JULIO": 7,
    "AGOSTO": 8, "SEPTIEMBRE": 9, "SETIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12,
    "ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "JUN": 6, "JUL": 7, "AGO": 8, "SEP": 9, "SEPT": 9,
    "OCT": 10, "NOV": 11, "DIC": 12,
}
_MES_RE = "|".join(sorted(MESES, key=len, reverse=True))

RE_FNG = re.compile(r"(?<![A-Z0-9])F\.?\s?N\.?\s?G(?![A-Z0-9])|FONDO\s+NACIONAL\s+DE\s+GARANTIA")
RE_LEASING = re.compile(
    r"(?<![A-Z])LEASING(?![A-Z])|ARRENDAMIENTO\s+FINANCIERO|ARRENDAMIENTO\s+CON\s+OPCION\s+DE\s+COMPRA"
)
# Palabras en el NOMBRE de un archivo que hacen que valga la pena leerlo
# aunque el buscador de Drive no lo haya marcado.
RE_NOMBRE_RELEVANTE = re.compile(r"FNG|FONDO NACIONAL|GARANT|LEASING|ARRENDAMIENTO|CONTRATO|CERTIFICADO")

_NUM = r"(?:N\s?[O°\.]{1,2}|NO\.?|NRO\.?|NUM\.?|NUMERO|#)"
RE_NUM_GARANTIA = re.compile(
    rf"(?:GARANTIA|CERTIFICADO|CERT\.?)\s*(?:FNG\s*)?(?:DE\s+GARANTIA\s*)?(?:FNG\s*)?{_NUM}\s*[:\.]?\s*([0-9][0-9A-Z\-]{{3,}})"
)
RE_NUM_OBLIGACION = re.compile(
    rf"(?:OBLIGACION(?:ES)?|PAGARE|CREDITO)\s*{_NUM}\s*[:\.]?\s*([0-9][0-9A-Z\-]{{3,}})"
)
RE_NUM_CONTRATO = re.compile(
    rf"(?:CONTRATO\s+(?:DE\s+)?(?:LEASING|ARRENDAMIENTO\s+FINANCIERO)?\s*(?:[A-Z]+\s+)?|LEASING\s*){_NUM}\s*[:\.]?\s*([0-9][0-9A-Z\-]{{2,}})"
)
RE_VALOR = re.compile(r"\$\s?([0-9]{1,3}(?:[\.,][0-9]{3})+(?:[\.,][0-9]{1,2})?|[0-9]{5,})")
RE_COBERTURA = re.compile(r"COBERTURA[^%\n]{0,60}?([0-9]{1,3}(?:[\.,][0-9]+)?)\s?%|([0-9]{1,3}(?:[\.,][0-9]+)?)\s?%[^%\n]{0,40}?COBERTURA")
RE_PLAZO = re.compile(r"PLAZO\s+(?:DE\s+|TOTAL\s+(?:DE\s+)?)?(?:[A-Z\s]{0,40}\(\s*)?([0-9]{1,3})\s*\)?\s*MESES")
RE_BIEN = re.compile(
    r"(?<![A-Z])(VEHICULOS?|CAMIONETAS?|CAMIONES|CAMION|TRACTOMULAS?|VOLQUETAS?|MAQUINARIA|MAQUINAS?|"
    r"EQUIPOS?|INMUEBLES?|BODEGAS?|LOCAL(?:ES)? COMERCIAL(?:ES)?|OFICINAS?|LOTES?|MONTACARGAS|"
    r"RETROEXCAVADORAS?|BUS(?:ES)?|AUTOMOVILES|AUTOMOVIL|PLACA\s+[A-Z]{3}\s?-?\s?[0-9]{2,3}[A-Z]?)(?![A-Z])"
)

PALABRAS_FECHA_FNG = {
    "VENCIMIENTO": ["VENCIMIENTO", "VENCE", "VENCERA", "FECHA FINAL", "FECHA DE FINALIZACION", "HASTA",
                    "FECHA LIMITE", "EXPIRA", "FECHA DE TERMINACION", "VIGENTE HASTA"],
    "EXPEDICION": ["EXPEDICION", "EXPEDIDA", "EXPEDIDO", "EMISION", "DESEMBOLSO", "DESDE", "OTORGAMIENTO",
                   "FECHA DE INICIO", "INICIO DE VIGENCIA", "CONSTITUCION"],
}
PALABRAS_FECHA_LEASING = {
    "TERMINACION": ["TERMINACION", "VENCIMIENTO", "VENCE", "FINALIZA", "FINALIZACION", "FECHA FINAL", "HASTA",
                    "ULTIMO CANON", "OPCION DE COMPRA", "EXPIRA", "FECHA DE TERMINO"],
    "INICIO": ["INICIO", "INICIACION", "SUSCRI", "FIRMA", "FIRMADO", "DESDE", "CELEBRA", "FECHA DEL CONTRATO",
               "DESEMBOLSO", "ACTIVACION", "ENTREGA", "PRIMER CANON"],
}

_cache_lock = threading.Lock()


# ============================= Utilidades =============================


def configurar_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
        handlers=[logging.FileHandler(ARCHIVO_LOG, encoding="utf-8"), logging.StreamHandler()],
    )
    # pypdf avisa por cada PDF con fuentes raras ("fontTools is required...",
    # "Ignoring wrong pointing object"...): no afecta la lectura, solo
    # llena la pantalla. Se ocultan esos avisos.
    for nombre in ("pypdf", "PyPDF2", "googleapiclient.discovery_cache"):
        logging.getLogger(nombre).setLevel(logging.ERROR)
    warnings.filterwarnings("ignore", module="pypdf")


def normalizar(texto: str) -> str:
    """Mayusculas y sin tildes, conservando la MISMA longitud (para poder
    recortar el fragmento original en las mismas posiciones)."""
    salida = []
    for c in texto or "":
        base = "".join(x for x in unicodedata.normalize("NFKD", c) if not unicodedata.combining(x))
        salida.append(base[:1] if base else " ")
    return "".join(salida).upper()


def tokens_nombre(texto: str):
    limpio = re.sub(r"[^A-Z0-9]+", " ", normalizar(texto))
    return [t for t in limpio.split() if t not in PALABRAS_IGNORADAS and not (t.isdigit() and len(t) < 5)]


def solo_digitos(texto) -> str:
    return re.sub(r"\D", "", str(texto or ""))


def nit_base(nit) -> str:
    """NIT sin digito de verificacion (ej. 900.123.456-7 -> 900123456)."""
    texto = str(nit or "").strip()
    if not texto:
        return ""
    texto = re.split(r"\s*-\s*\d\s*$", texto)[0]
    digitos = solo_digitos(texto)
    return digitos if len(digitos) >= 6 else ""


def fecha_valida(anio, mes, dia):
    try:
        anio, mes, dia = int(anio), int(mes), int(dia)
        if not 1990 <= anio <= 2070:
            return None
        return datetime.date(anio, mes, dia)
    except (ValueError, TypeError):
        return None


_RE_FECHAS = [
    # 15/03/2025, 15-03-2025, 15.03.2025
    (re.compile(r"(?<!\d)(\d{1,2})\s?[/\-\.]\s?(\d{1,2})\s?[/\-\.]\s?(\d{4})(?!\d)"), lambda m: (m[3], m[2], m[1])),
    # 2025/03/15, 2025-03-15
    (re.compile(r"(?<!\d)(\d{4})\s?[/\-\.]\s?(\d{1,2})\s?[/\-\.]\s?(\d{1,2})(?!\d)"), lambda m: (m[1], m[2], m[3])),
    # 15 de marzo de 2025, 15 MAR 2025, 15-mar-2025
    (re.compile(rf"(?<!\d)(\d{{1,2}})\s*(?:DE\s+|-|/)?\s*({_MES_RE})\.?\s*(?:DE\s+|DEL\s+|-|/|,)?\s*(?:ANO\s+)?(\d{{4}})(?!\d)"),
     lambda m: (m[3], MESES[m[2]], m[1])),
    # marzo 15 de 2025
    (re.compile(rf"(?<![A-Z])({_MES_RE})\.?\s+(\d{{1,2}})\s*(?:DE|DEL|,)?\s*(\d{{4}})(?!\d)"),
     lambda m: (m[3], MESES[m[1]], m[2])),
    # "a los veinte (20) dias del mes de marzo de dos mil veinte (2020)"
    (re.compile(rf"\((\d{{1,2}})\)\s*DIAS?\s+DEL\s+MES\s+DE\s+({_MES_RE})\s+(?:DE|DEL)\s+(?:[A-Z\s]{{0,40}}\(\s*)?(\d{{4}})"),
     lambda m: (m[3], MESES[m[2]], m[1])),
]


def fechas_en_texto(texto_norm: str, desde: int = 0, hasta: int = None):
    """[(posicion, fecha)] de todas las fechas reconocibles en el tramo."""
    hasta = len(texto_norm) if hasta is None else hasta
    tramo = texto_norm[desde:hasta]
    encontradas = {}
    for regex, partes in _RE_FECHAS:
        for m in regex.finditer(tramo):
            f = fecha_valida(*partes(m))
            if f:
                encontradas.setdefault(desde + m.start(), f)
    return sorted(encontradas.items())


def clasificar_fecha(texto_norm: str, posicion: int, palabras_por_tipo: dict, alcance: int = 90):
    """Mira las palabras que estan JUSTO ANTES de la fecha (hasta 'alcance'
    caracteres) y devuelve el tipo cuya palabra este mas cerca, o None."""
    previo = texto_norm[max(0, posicion - alcance):posicion]
    mejor_tipo, mejor_pos = None, -1
    for tipo, palabras in palabras_por_tipo.items():
        for palabra in palabras:
            p = previo.rfind(palabra)
            if p > mejor_pos:
                mejor_tipo, mejor_pos = tipo, p
    return mejor_tipo


def sumar_meses(fecha: datetime.date, meses: int) -> datetime.date:
    mes_total = fecha.month - 1 + meses
    anio, mes = fecha.year + mes_total // 12, mes_total % 12 + 1
    for dia in (fecha.day, 30, 29, 28):
        try:
            return datetime.date(anio, mes, dia)
        except ValueError:
            continue
    return fecha


def unicos(valores):
    vistos, salida = set(), []
    for v in valores:
        if v and v not in vistos:
            vistos.add(v)
            salida.append(v)
    return salida


def fmt_fechas(fechas):
    return ", ".join(f.strftime("%d/%m/%Y") for f in sorted(set(fechas)))


# ============================= Lectura del Excel =============================


def _encabezado_norm(valor) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", normalizar(str(valor or "")).lower())).strip()


def numeros_expediente(valor):
    """Numeros de expediente/radicado de una celda (ej. 69686.0 -> ['69686'];
    'J2026... // 84862' -> [..., '84862']). Solo numeros de 5+ digitos."""
    if valor is None or isinstance(valor, (datetime.date, datetime.datetime)):
        return []
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return unicos(re.findall(r"\d{5,}", str(valor)))


def tipo_incluido(tipo) -> bool:
    if not TIPOS_PROCESO_A_INCLUIR:
        return True
    texto = " " + re.sub(r"[^A-Z0-9]+", " ", normalizar(str(tipo or ""))) + " "
    return any(f" {normalizar(t).strip()} " in texto for t in TIPOS_PROCESO_A_INCLUIR)


def leer_listado(ruta: str):
    """Lee el listado y aplica el filtro por tipo de proceso. Devuelve un
    dict con encabezados, filas (valores originales) e indices de columnas."""
    libro = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    hoja = libro[HOJA_EXCEL_PROCESOS] if HOJA_EXCEL_PROCESOS else libro.worksheets[0]
    filas = [list(f) for f in hoja.iter_rows(values_only=True)]
    libro.close()

    def buscar_columna(encabezados, explicito, candidatos):
        normalizados = [_encabezado_norm(e) for e in encabezados]
        if explicito:
            objetivo = _encabezado_norm(explicito)
            return normalizados.index(objetivo) if objetivo in normalizados else None
        for candidato in candidatos:
            if candidato in normalizados:
                return normalizados.index(candidato)
        for candidato in candidatos:
            for i, n in enumerate(normalizados):
                if candidato in n.split() or n.startswith(candidato):
                    return i
        return None

    for i_enc, fila in enumerate(filas[:25]):
        idx_conc = buscar_columna(fila, COLUMNA_CONCURSADO, ENCABEZADOS_CONCURSADO)
        if idx_conc is None:
            continue
        info = {
            "idx_conc": idx_conc,
            "idx_nit": buscar_columna(fila, COLUMNA_NIT, ENCABEZADOS_NIT),
            "idx_exp": buscar_columna(fila, COLUMNA_EXPEDIENTE, ENCABEZADOS_EXPEDIENTE),
            "idx_tipo": buscar_columna(fila, COLUMNA_TIPO_PROCESO, ENCABEZADOS_TIPO),
        }
        datos = [f for f in filas[i_enc + 1:] if f and idx_conc < len(f) and str(f[idx_conc] or "").strip()]
        info["total_filas"] = len(datos)
        if TIPOS_PROCESO_A_INCLUIR and info["idx_tipo"] is None:
            logging.warning("OJO: no encontre la columna del tipo de proceso; se revisan TODAS las filas. "
                            "Pon su encabezado en COLUMNA_TIPO_PROCESO.")
        elif TIPOS_PROCESO_A_INCLUIR:
            it = info["idx_tipo"]
            datos = [f for f in datos if it < len(f) and tipo_incluido(f[it])]
        # Quita las columnas del final que no tienen encabezado ni datos.
        ancho = len(fila)
        while ancho > 1 and fila[ancho - 1] in (None, "") and all(
            len(f) < ancho or f[ancho - 1] in (None, "") for f in datos
        ):
            ancho -= 1
        info["encabezados"] = [
            str(e).strip() if e not in (None, "") else f"Columna {j + 1}" for j, e in enumerate(fila[:ancho])
        ]
        info["filas"] = datos
        return info
    raise RuntimeError(
        "No encontre en el Excel la columna con el nombre del concursado. Pon su encabezado exacto en "
        "COLUMNA_CONCURSADO (al inicio de depurar_fng_leasing.py)."
    )


# ============================= Fuentes (Drive API / carpeta local) =============================


def _ruta_chrome():
    for base in (os.environ.get("PROGRAMFILES"), os.environ.get("PROGRAMFILES(X86)"), os.environ.get("LOCALAPPDATA")):
        if base:
            ruta = os.path.join(base, "Google", "Chrome", "Application", "chrome.exe")
            if os.path.exists(ruta):
                return ruta
    return None


class _NavegadorAutorizacion(webbrowser.BaseBrowser):
    """Abre el enlace de autorizacion de Google en CHROME (no en el
    navegador predeterminado, que puede ser Edge sin tu cuenta), y ademas
    lo copia al portapapeles y lo guarda en un .txt -- copiarlo a mano de
    la ventana negra lo corta en varias lineas y Google responde 404."""

    def open(self, url, new=0, autoraise=True):
        try:
            with open(ARCHIVO_ENLACE_AUTORIZACION, "w", encoding="utf-8") as f:
                f.write(url + "\n")
        except OSError:
            pass
        if sys.platform.startswith("win"):
            try:
                subprocess.run(["clip"], input=url, text=True, check=False)
            except OSError:
                pass
        chrome = _ruta_chrome()
        if chrome:
            try:
                subprocess.Popen([chrome, url])
                return True
            except OSError:
                pass
        return webbrowser.open(url, new=new, autoraise=autoraise)


class FuenteDriveAPI:
    def __init__(self):
        if Credentials is None:
            raise RuntimeError(
                "Faltan las librerias de Google Drive. Instala con: pip install -r requirements.txt"
            )
        self.creds = self._autenticar()
        self._local = threading.local()
        self.servicio  # crea el del hilo principal

    def _autenticar(self):
        creds = None
        if os.path.exists(TOKEN_DRIVE):
            creds = Credentials.from_authorized_user_file(TOKEN_DRIVE, DRIVE_SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not os.path.exists(CREDENCIALES_DRIVE):
                    raise RuntimeError(
                        f"No existe {CREDENCIALES_DRIVE}. Sigue los pasos del README (\"Configurar el acceso a "
                        "Google Drive\"), o usa CARPETA_LOCAL_DRIVE si tienes Google Drive para escritorio."
                    )
                flow = InstalledAppFlow.from_client_secrets_file(CREDENCIALES_DRIVE, DRIVE_SCOPES)
                webbrowser.register("autorizar_drive", None, _NavegadorAutorizacion("autorizar_drive"))
                creds = flow.run_local_server(
                    port=0, browser="autorizar_drive",
                    authorization_prompt_message=(
                        "\n>>> Autoriza el acceso a tu Google Drive en el navegador (se intento abrir en CHROME).\n"
                        ">>> Si no se abrio, o se abrio en otro navegador: abre Chrome, haz clic en la barra de\n"
                        ">>> direcciones y pega con Ctrl+V -- el enlace COMPLETO ya esta copiado. Tambien quedo en:\n"
                        f">>> {ARCHIVO_ENLACE_AUTORIZACION}\n"
                        ">>> NO cierres esta ventana mientras autorizas.\n"
                    ),
                    success_message="Listo, ya puedes cerrar esta pestana y volver a la ventana negra.",
                )
            with open(TOKEN_DRIVE, "w", encoding="utf-8") as f:
                f.write(creds.to_json())
        return creds

    @property
    def servicio(self):
        # httplib2 no es seguro entre hilos: un cliente por hilo.
        if not hasattr(self._local, "servicio"):
            self._local.servicio = build("drive", "v3", credentials=self.creds, cache_discovery=False)
        return self._local.servicio

    def _listar(self, consulta, campos="id, name, mimeType, parents, modifiedTime, size, webViewLink, shortcutDetails"):
        resultados, token = [], None
        while True:
            r = self.servicio.files().list(
                q=consulta, pageSize=1000, pageToken=token,
                fields=f"nextPageToken, files({campos})",
                supportsAllDrives=True, includeItemsFromAllDrives=True, corpora="allDrives",
            ).execute()
            resultados.extend(r.get("files", []))
            token = r.get("nextPageToken")
            if not token:
                return resultados

    def _id_raiz(self):
        valor = CARPETA_RAIZ_DRIVE.strip()
        if not valor:
            return None
        m = re.search(r"/folders/([A-Za-z0-9_\-]+)", valor) or re.search(r"[?&]id=([A-Za-z0-9_\-]+)", valor)
        if m:
            return m.group(1)
        if re.fullmatch(r"[A-Za-z0-9_\-]{20,}", valor):
            return valor
        nombre = valor.replace("\\", "\\\\").replace("'", "\\'")
        encontradas = self._listar(f"name = '{nombre}' and mimeType = '{MIME_CARPETA}' and trashed = false")
        if not encontradas:
            raise RuntimeError(f"No encontre en Drive ninguna carpeta llamada '{valor}' (CARPETA_RAIZ_DRIVE).")
        if len(encontradas) > 1:
            logging.warning("Hay %d carpetas llamadas '%s'; uso la primera. Mejor pon su enlace o ID.",
                            len(encontradas), valor)
        return encontradas[0]["id"]

    def listar_carpetas_candidatas(self):
        """Carpetas donde puede estar la de cada proceso: [{id, nombre, padres, enlace}]."""
        raiz = self._id_raiz()
        carpetas = []
        if raiz is None:
            logging.info("Listando TODAS las carpetas del Drive (puede tardar un poco)...")
            for c in self._listar(f"mimeType = '{MIME_CARPETA}' and trashed = false"):
                carpetas.append(self._carpeta(c))
            return carpetas
        logging.info("Listando las carpetas dentro de CARPETA_RAIZ_DRIVE (hasta %d niveles)...",
                     PROFUNDIDAD_BUSQUEDA_CARPETAS)
        nivel = [raiz]
        for _ in range(PROFUNDIDAD_BUSQUEDA_CARPETAS):
            siguiente = []
            for padre in nivel:
                for c in self._listar(f"'{padre}' in parents and mimeType = '{MIME_CARPETA}' and trashed = false"):
                    carpetas.append(self._carpeta(c))
                    siguiente.append(c["id"])
            nivel = siguiente
            if not nivel:
                break
        return carpetas

    @staticmethod
    def _carpeta(c):
        return {"id": c["id"], "nombre": c["name"], "padres": c.get("parents", []),
                "enlace": c.get("webViewLink") or f"https://drive.google.com/drive/folders/{c['id']}"}

    def listar_archivos(self, carpeta_id):
        """Todos los archivos (no carpetas) dentro de la carpeta, recursivo."""
        archivos, pendientes, visitadas = [], [(carpeta_id, "")], set()
        while pendientes:
            actual, ruta = pendientes.pop()
            if actual in visitadas:
                continue
            visitadas.add(actual)
            for item in self._listar(f"'{actual}' in parents and trashed = false"):
                mime = item["mimeType"]
                if mime == MIME_ATAJO:
                    destino = item.get("shortcutDetails", {})
                    if destino.get("targetMimeType") == MIME_CARPETA:
                        pendientes.append((destino["targetId"], f"{ruta}{item['name']}/"))
                        continue
                    if not destino.get("targetId"):
                        continue
                    item = dict(item, id=destino["targetId"], mimeType=destino.get("targetMimeType", ""))
                    mime = item["mimeType"]
                if mime == MIME_CARPETA:
                    pendientes.append((item["id"], f"{ruta}{item['name']}/"))
                    continue
                archivos.append({
                    "id": item["id"], "nombre": item["name"], "ruta": f"{ruta}{item['name']}", "mime": mime,
                    "version": f"{item.get('modifiedTime', '')}|{item.get('size', '')}",
                    "tamano": int(item.get("size") or 0),
                    "enlace": item.get("webViewLink") or f"https://drive.google.com/file/d/{item['id']}/view",
                })
        return archivos

    def ids_con_texto(self, terminos):
        """IDs de los archivos cuyo CONTENIDO menciona alguno de los terminos,
        segun el buscador de Google Drive (para el modo 'rapido')."""
        ids = set()
        for termino in terminos:
            t = termino.replace("'", "\\'")
            try:
                for f in self._listar(f"fullText contains '{t}' and trashed = false", campos="id"):
                    ids.add(f["id"])
            except Exception as e:  # noqa: BLE001
                logging.warning("No se pudo usar el buscador de Drive para '%s' (%s).", termino, e)
                return None
        return ids

    def leer(self, archivo):
        """(bytes, extension) del archivo, o (None, motivo)."""
        mime = archivo["mime"]
        if mime in MIME_EXPORTAR_TEXTO:
            peticion = self.servicio.files().export_media(fileId=archivo["id"], mimeType=MIME_EXPORTAR_TEXTO[mime])
            extension = ".csv" if mime.endswith("spreadsheet") else ".txt"
        elif mime.startswith("application/vnd.google-apps"):
            return None, "tipo de Google no exportable"
        else:
            extension = Path(archivo["nombre"]).suffix.lower()
            if mime == "application/pdf":
                extension = ".pdf"
            if extension not in EXTENSIONES_LEIBLES:
                return None, f"formato {extension or mime} no se lee"
            if archivo["tamano"] > MAX_MB_ARCHIVO * 1024 * 1024:
                return None, f"pesa mas de {MAX_MB_ARCHIVO} MB"
            peticion = self.servicio.files().get_media(fileId=archivo["id"], supportsAllDrives=True)
        buffer = io.BytesIO()
        descarga = MediaIoBaseDownload(buffer, peticion)
        listo = False
        while not listo:
            _, listo = descarga.next_chunk()
        return buffer.getvalue(), extension


class FuenteLocal:
    def __init__(self, raiz):
        self.raiz = Path(raiz)
        if not self.raiz.is_dir():
            raise RuntimeError(f"No existe la carpeta CARPETA_LOCAL_DRIVE: {raiz}")

    def listar_carpetas_candidatas(self):
        carpetas = []
        nivel = [self.raiz]
        for _ in range(PROFUNDIDAD_BUSQUEDA_CARPETAS):
            siguiente = []
            for padre in nivel:
                try:
                    hijos = [p for p in padre.iterdir() if p.is_dir()]
                except OSError:
                    continue
                for p in hijos:
                    carpetas.append({"id": str(p), "nombre": p.name, "padres": [str(padre)], "enlace": str(p)})
                    siguiente.append(p)
            nivel = siguiente
        return carpetas

    def listar_archivos(self, carpeta_id):
        base = Path(carpeta_id)
        archivos = []
        for p in base.rglob("*"):
            if p.is_file():
                try:
                    st = p.stat()
                except OSError:
                    continue
                archivos.append({
                    "id": str(p), "nombre": p.name, "ruta": str(p.relative_to(base)), "mime": "",
                    "version": f"{st.st_mtime}|{st.st_size}", "tamano": st.st_size, "enlace": str(p),
                })
        return archivos

    def ids_con_texto(self, terminos):
        return None

    def leer(self, archivo):
        extension = Path(archivo["nombre"]).suffix.lower()
        if extension in (".gdoc", ".gsheet", ".gslides"):
            return None, "documento de Google (usa la API para leerlo)"
        if extension not in EXTENSIONES_LEIBLES:
            return None, f"formato {extension or '(sin extension)'} no se lee"
        if archivo["tamano"] > MAX_MB_ARCHIVO * 1024 * 1024:
            return None, f"pesa mas de {MAX_MB_ARCHIVO} MB"
        return Path(archivo["id"]).read_bytes(), extension


# ============================= Extraccion de texto =============================


def texto_de_bytes(contenido: bytes, extension: str) -> str:
    if extension == ".pdf":
        if PdfReader is None:
            raise RuntimeError("falta pypdf (pip install pypdf)")
        lector = PdfReader(io.BytesIO(contenido))
        paginas = lector.pages[:MAX_PAGINAS_PDF]
        return "\n".join((p.extract_text() or "") for p in paginas)
    if extension == ".docx":
        if docx is None:
            raise RuntimeError("falta python-docx")
        documento = docx.Document(io.BytesIO(contenido))
        partes = [p.text for p in documento.paragraphs]
        for tabla in documento.tables:
            for fila in tabla.rows:
                partes.append(" | ".join(c.text for c in fila.cells))
        return "\n".join(partes)
    if extension in (".xlsx", ".xlsm"):
        libro = openpyxl.load_workbook(io.BytesIO(contenido), read_only=True, data_only=True)
        partes = []
        for hoja in libro.worksheets:
            for fila in hoja.iter_rows(values_only=True):
                valores = [
                    v.strftime("%d/%m/%Y") if isinstance(v, (datetime.date, datetime.datetime)) else str(v)
                    for v in fila if v is not None
                ]
                if valores:
                    partes.append(" | ".join(valores))
        libro.close()
        return "\n".join(partes)
    if extension == ".eml":
        mensaje = email.message_from_bytes(contenido, policy=policy.default)
        partes = [f"ASUNTO: {mensaje.get('subject', '')}", f"FECHA: {mensaje.get('date', '')}"]
        for parte in mensaje.walk():
            if parte.get_content_type() in ("text/plain", "text/html"):
                try:
                    partes.append(parte.get_content())
                except Exception:  # noqa: BLE001
                    pass
        return re.sub(r"<[^>]+>", " ", "\n".join(partes))
    texto = None
    for codificacion in ("utf-8", "latin-1"):
        try:
            texto = contenido.decode(codificacion)
            break
        except UnicodeDecodeError:
            continue
    if extension in (".html", ".htm", ".xml"):
        texto = re.sub(r"<[^>]+>", " ", texto or "")
    return texto or ""


class Cache:
    """Texto ya leido de cada documento (para no volver a descargarlo).
    Se guarda en un archivo de LINEAS (una por documento) que se va
    agregando al instante: si el programa se corta, no se pierde nada."""

    def __init__(self, ruta):
        self.ruta = ruta
        self.datos = {}
        if os.path.exists(ARCHIVO_CACHE_VIEJO):
            # Migra el cache del formato anterior (un solo JSON).
            try:
                with open(ARCHIVO_CACHE_VIEJO, encoding="utf-8") as f:
                    viejo = json.load(f)
                with open(ruta, "a", encoding="utf-8") as f:
                    for clave, e in viejo.items():
                        f.write(json.dumps(dict(e, id=clave), ensure_ascii=False) + "\n")
                os.replace(ARCHIVO_CACHE_VIEJO, ARCHIVO_CACHE_VIEJO + ".migrado")
            except (OSError, ValueError):
                pass
        if os.path.exists(ruta):
            with open(ruta, encoding="utf-8", errors="replace") as f:
                for linea in f:
                    try:
                        e = json.loads(linea)
                        self.datos[e.pop("id")] = e
                    except (ValueError, KeyError):
                        continue  # linea cortada por un apagon: se ignora
        self._archivo = open(ruta, "a", encoding="utf-8")

    def obtener(self, archivo):
        e = self.datos.get(archivo["id"])
        if e and e.get("version") == archivo["version"]:
            return e
        return None

    def guardar_entrada(self, archivo, texto, motivo):
        e = {"version": archivo["version"], "texto": texto, "motivo": motivo}
        with _cache_lock:
            self.datos[archivo["id"]] = e
            self._archivo.write(json.dumps(dict(e, id=archivo["id"]), ensure_ascii=False) + "\n")
            self._archivo.flush()

    def guardar(self):
        with _cache_lock:
            try:
                self._archivo.flush()
                os.fsync(self._archivo.fileno())
            except (OSError, ValueError):
                pass


def _json_a_disco(obj):
    if isinstance(obj, (datetime.date, datetime.datetime)):
        return {"__fecha__": obj.isoformat()[:10]}
    if isinstance(obj, set):
        return sorted(obj)
    raise TypeError(type(obj))


def _json_de_disco(d):
    if "__fecha__" in d and len(d) == 1:
        return datetime.date.fromisoformat(d["__fecha__"])
    return d


class Progreso:
    """Avance de la corrida actual, para retomar si se corta."""

    def __init__(self, ruta, firma):
        self.ruta = ruta
        self.datos = {"firma": firma, "inicio": datetime.datetime.now().isoformat(), "procesos": {}}
        if not os.path.exists(ruta):
            return
        try:
            with open(ruta, encoding="utf-8") as f:
                previo = json.load(f, object_hook=_json_de_disco)
            inicio = datetime.datetime.fromisoformat(previo["inicio"])
        except (OSError, ValueError, KeyError, TypeError):
            return
        if previo.get("firma") != firma:
            logging.info("(Hay un avance guardado de OTRA configuracion/listado: se empieza de nuevo.)")
        elif datetime.datetime.now() - inicio > datetime.timedelta(days=DIAS_VALIDEZ_PROGRESO):
            logging.info("(El avance guardado tiene mas de %d dias: se empieza de nuevo.)", DIAS_VALIDEZ_PROGRESO)
        else:
            self.datos = previo
            logging.info(">>> RETOMANDO la corrida del %s: %d procesos ya estaban revisados.",
                         inicio.strftime("%d/%m/%Y %H:%M"), len(previo.get("procesos", {})))

    def get(self, clave, defecto=None):
        return self.datos.get(clave, defecto)

    def poner(self, clave, valor):
        self.datos[clave] = valor
        self.guardar()

    def proceso(self, clave):
        return self.datos["procesos"].get(clave)

    def marcar_proceso(self, clave, valor):
        self.datos["procesos"][clave] = valor
        self.guardar()

    def guardar(self):
        tmp = self.ruta + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self.datos, f, ensure_ascii=False, default=_json_a_disco)
        os.replace(tmp, self.ruta)

    def borrar(self):
        for ruta in (self.ruta, self.ruta + ".tmp"):
            try:
                os.remove(ruta)
            except OSError:
                pass


def mantener_pc_despierto():
    """En Windows, le pide al sistema que NO se suspenda mientras el
    programa este corriendo (la pantalla si se puede apagar, eso no
    afecta). Se suelta solo al cerrar el programa."""
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
        ES_CONTINUOUS, ES_SYSTEM_REQUIRED = 0x80000000, 0x00000001
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)
        logging.info("(El PC no se va a suspender mientras esto corre.)")
    except Exception:  # noqa: BLE001
        pass


def obtener_texto(fuente, cache: Cache, archivo):
    """(texto, motivo_si_no_se_pudo_leer)."""
    e = cache.obtener(archivo)
    if e is not None:
        return e["texto"], e["motivo"]
    texto, motivo = "", ""
    try:
        contenido, extension = fuente.leer(archivo)
        if contenido is None:
            motivo = extension
        else:
            texto = texto_de_bytes(contenido, extension)
            if extension == ".pdf" and len(texto.strip()) < 40:
                motivo = "PDF escaneado (sin texto, necesita OCR)"
    except Exception as e:  # noqa: BLE001
        motivo = f"no se pudo leer ({type(e).__name__}: {str(e)[:120]})"
    cache.guardar_entrada(archivo, texto, motivo)
    return texto, motivo


# ============================= Cruce concursado -> carpeta =============================


def _token_igual(a, b):
    if a == b:
        return True
    if len(a) >= 5 and len(b) >= 5:
        return difflib.SequenceMatcher(None, a, b).ratio() >= 0.88
    return False


def puntaje(tokens_proceso, nit, expedientes, carpeta):
    nombre = carpeta["nombre"]
    if nit and nit in solo_digitos(nombre):
        return 1.0
    if any(re.search(rf"(?<!\d){e}(?!\d)", nombre) for e in expedientes):
        return 1.0
    if not tokens_proceso:
        return 0.0
    tokens_carpeta = carpeta.setdefault("_tokens", tokens_nombre(nombre))
    if not tokens_carpeta:
        return 0.0
    presentes = sum(1 for t in tokens_proceso if any(_token_igual(t, c) for c in tokens_carpeta))
    cobertura = presentes / len(tokens_proceso)
    if cobertura == 1.0:
        # Penaliza un poco las carpetas con muchas palabras extra (para
        # preferir "ACME" sobre "ACME - ANEXOS ACREEDOR XYZ").
        extra = max(0, len(tokens_carpeta) - len(tokens_proceso))
        return 0.99 - min(extra, 6) * 0.01
    a, b = " ".join(tokens_proceso), " ".join(tokens_carpeta)
    m = difflib.SequenceMatcher(None, a, b)
    if m.real_quick_ratio() < UMBRAL_SIMILITUD_NOMBRE or m.quick_ratio() < UMBRAL_SIMILITUD_NOMBRE:
        return cobertura * 0.8
    return max(cobertura * 0.8, m.ratio())


def buscar_carpetas_del_proceso(nombre, nit, expedientes, carpetas, padres_de):
    tokens = tokens_nombre(nombre)
    puntuadas = sorted(
        ((puntaje(tokens, nit, expedientes, c), c) for c in carpetas), key=lambda x: x[0], reverse=True
    )
    puntuadas = [(p, c) for p, c in puntuadas if p > 0][:15]
    if not puntuadas or puntuadas[0][0] < UMBRAL_SIMILITUD_NOMBRE:
        return [], puntuadas[:3]
    mejor = puntuadas[0][0]
    elegidas = [c for p, c in puntuadas if p >= UMBRAL_SIMILITUD_NOMBRE and p >= mejor - 0.03]
    # Si una elegida esta DENTRO de otra elegida, basta con la de arriba
    # (la lectura es recursiva).
    ids = {c["id"] for c in elegidas}

    def es_descendiente(cid):
        vistos, pendientes = set(), list(padres_de.get(cid, []))
        while pendientes:
            p = pendientes.pop()
            if p in ids:
                return True
            if p not in vistos:
                vistos.add(p)
                pendientes.extend(padres_de.get(p, []))
        return False

    elegidas = [c for c in elegidas if not es_descendiente(c["id"])]
    alternativas = [(p, c) for p, c in puntuadas if c not in elegidas and p >= UMBRAL_SIMILITUD_NOMBRE * 0.85][:3]
    return elegidas, alternativas


# ============================= Analisis FNG / LEASING =============================


def _ventanas(texto_norm, regex, antes=500, despues=700):
    """Tramos [ini, fin) alrededor de cada mencion, fusionando los que se pisan."""
    tramos = []
    for m in regex.finditer(texto_norm):
        ini, fin = max(0, m.start() - antes), min(len(texto_norm), m.end() + despues)
        if tramos and ini <= tramos[-1][1]:
            tramos[-1][1] = max(tramos[-1][1], fin)
        else:
            tramos.append([ini, fin])
    return tramos


def _fragmento(texto, texto_norm, regex, largo=350):
    m = regex.search(texto_norm)
    if not m:
        return ""
    ini = max(0, m.start() - largo // 3)
    return re.sub(r"\s+", " ", texto[ini:ini + largo]).strip()


def _entidades(tramo):
    halladas = [e for e in ENTIDADES if re.search(rf"(?<![A-Z]){re.escape(e)}(?![A-Z])", tramo)]
    # Quita las que estan contenidas en otra mas larga (OCCIDENTE dentro de BANCO DE OCCIDENTE).
    return [e for e in halladas if not any(e != o and e in o for o in halladas)]


def analizar_fng(texto, nombre_archivo):
    texto_norm = normalizar(texto)
    nombre_norm = normalizar(nombre_archivo)
    en_nombre = bool(RE_FNG.search(nombre_norm))
    tramos = _ventanas(texto_norm, RE_FNG)
    if not tramos and not en_nombre:
        return None
    if en_nombre and not tramos and texto_norm.strip():
        # Documento cuyo NOMBRE dice FNG: se analiza completo.
        tramos = [[0, min(len(texto_norm), 6000)]]
    numeros, obligaciones, entidades, valores, coberturas = [], [], [], [], []
    vencimientos, expediciones, otras = [], [], []
    for ini, fin in tramos:
        tramo = texto_norm[ini:fin]
        numeros += RE_NUM_GARANTIA.findall(tramo)
        obligaciones += RE_NUM_OBLIGACION.findall(tramo)
        entidades += _entidades(tramo)
        valores += ["$" + v for v in RE_VALOR.findall(tramo)]
        coberturas += [(a or b) + "%" for a, b in RE_COBERTURA.findall(tramo)]
        for pos, f in fechas_en_texto(texto_norm, ini, fin):
            tipo = clasificar_fecha(texto_norm, pos, PALABRAS_FECHA_FNG)
            {"VENCIMIENTO": vencimientos, "EXPEDICION": expediciones}.get(tipo, otras).append(f)
    return {
        "menciones": len(RE_FNG.findall(texto_norm)) + (1 if en_nombre else 0),
        "numeros": unicos(numeros)[:6], "obligaciones": unicos(obligaciones)[:6],
        "entidades": unicos(e for e in entidades if e != "FNG")[:5],
        "valores": unicos(valores)[:5], "coberturas": unicos(coberturas)[:3],
        "vencimientos": sorted(set(vencimientos)), "expediciones": sorted(set(expediciones)),
        "otras_fechas": sorted(set(otras))[:8],
        "fragmento": _fragmento(texto, texto_norm, RE_FNG) or f"(mencionado en el nombre: {nombre_archivo})",
        "sin_texto": not texto_norm.strip(),
    }


def analizar_leasing(texto, nombre_archivo):
    texto_norm = normalizar(texto)
    nombre_norm = normalizar(nombre_archivo)
    en_nombre = bool(RE_LEASING.search(nombre_norm))
    tramos = _ventanas(texto_norm, RE_LEASING, antes=500, despues=900)
    if not tramos and not en_nombre:
        return None
    if en_nombre and texto_norm.strip():
        # Un contrato de leasing: las fechas pueden estar en cualquier
        # clausula, se analiza el documento entero (hasta un limite).
        tramos = [[0, min(len(texto_norm), 40000)]]
    contratos, entidades, bienes, plazos, inicios, terminaciones, otras = [], [], [], [], [], [], []
    for ini, fin in tramos:
        tramo = texto_norm[ini:fin]
        contratos += RE_NUM_CONTRATO.findall(tramo)
        entidades += _entidades(tramo)
        bienes += [re.sub(r"\s+", " ", b) for b in RE_BIEN.findall(tramo)]
        plazos += [int(p) for p in RE_PLAZO.findall(tramo) if 0 < int(p) <= 360]
        for pos, f in fechas_en_texto(texto_norm, ini, fin):
            tipo = clasificar_fecha(texto_norm, pos, PALABRAS_FECHA_LEASING)
            {"INICIO": inicios, "TERMINACION": terminaciones}.get(tipo, otras).append(f)
    estimadas = []
    if inicios and plazos and not terminaciones:
        estimadas = [sumar_meses(min(inicios), plazos[0])]
    return {
        "menciones": len(RE_LEASING.findall(texto_norm)) + (1 if en_nombre else 0),
        "contratos": unicos(contratos)[:6], "entidades": unicos(entidades)[:5],
        "bienes": unicos(bienes)[:6], "plazos": unicos(plazos)[:3],
        "inicios": sorted(set(inicios)), "terminaciones": sorted(set(terminaciones)),
        "terminaciones_estimadas": estimadas, "otras_fechas": sorted(set(otras))[:8],
        "fragmento": _fragmento(texto, texto_norm, RE_LEASING) or f"(mencionado en el nombre: {nombre_archivo})",
        "sin_texto": not texto_norm.strip(),
    }


# ============================= Proceso principal =============================


def revisar_proceso(fuente, cache, carpetas, ids_indexados):
    """Lee los documentos de las carpetas del proceso y devuelve
    (hallazgos_fng, hallazgos_leasing, sin_leer, total_archivos)."""
    archivos = []
    for c in carpetas:
        for a in fuente.listar_archivos(c["id"]):
            a["carpeta"] = c["nombre"]
            archivos.append(a)

    def debe_leerse(a):
        if ids_indexados is None:
            return True
        return a["id"] in ids_indexados or bool(RE_NOMBRE_RELEVANTE.search(normalizar(a["ruta"])))

    a_leer = [a for a in archivos if debe_leerse(a)]
    hilos = NUM_HILOS if isinstance(fuente, FuenteDriveAPI) else 1
    with ThreadPoolExecutor(max_workers=max(1, hilos)) as ex:
        textos = list(ex.map(lambda a: obtener_texto(fuente, cache, a), a_leer))

    fng, leasing, sin_leer = [], [], []
    for a, (texto, motivo) in zip(a_leer, textos):
        nombre_y_ruta = f"{a['carpeta']}/{a['ruta']}"
        rf = analizar_fng(texto, nombre_y_ruta)
        rl = analizar_leasing(texto, nombre_y_ruta)
        if rf:
            fng.append((a, rf))
        if rl:
            leasing.append((a, rl))
        if motivo and RE_NOMBRE_RELEVANTE.search(normalizar(nombre_y_ruta)):
            sin_leer.append((a, motivo))
    return fng, leasing, sin_leer, len(archivos)


def resumir_proceso(fng, leasing):
    r = {}
    venc = sorted({f for _, h in fng for f in h["vencimientos"]})
    r["fng"] = "SI" if fng else "NO"
    r["fng_docs"] = len(fng)
    r["fng_numeros"] = ", ".join(unicos(n for _, h in fng for n in h["numeros"]))
    r["fng_obligaciones"] = ", ".join(unicos(n for _, h in fng for n in h["obligaciones"]))
    r["fng_entidades"] = ", ".join(unicos(n for _, h in fng for n in h["entidades"]))
    r["fng_valores"] = ", ".join(unicos(n for _, h in fng for n in h["valores"] + h["coberturas"])[:6])
    r["fng_vencimientos"] = fmt_fechas(venc)
    futuros = [f for f in venc if f >= HOY]
    r["fng_proximo"] = min(futuros) if futuros else (max(venc) if venc else None)
    if not fng:
        r["fng_estado"] = ""
    elif not venc:
        r["fng_estado"] = "SIN FECHA - revisar"
    else:
        r["fng_estado"] = "VIGENTE" if futuros else "VENCIDA"

    inicios = sorted({f for _, h in leasing for f in h["inicios"]})
    terms = sorted({f for _, h in leasing for f in h["terminaciones"]})
    estimadas = sorted({f for _, h in leasing for f in h["terminaciones_estimadas"]})
    r["leasing"] = "SI" if leasing else "NO"
    r["leasing_docs"] = len(leasing)
    r["leasing_contratos"] = ", ".join(unicos(n for _, h in leasing for n in h["contratos"]))
    r["leasing_entidades"] = ", ".join(unicos(n for _, h in leasing for n in h["entidades"]))
    r["leasing_bienes"] = ", ".join(unicos(n for _, h in leasing for n in h["bienes"])[:6])
    r["leasing_inicios"] = fmt_fechas(inicios)
    r["leasing_terminaciones"] = fmt_fechas(terms) + (
        (" | estimada: " if terms else "estimada: ") + fmt_fechas(estimadas) if estimadas else ""
    )
    r["leasing_plazos"] = ", ".join(str(p) for p in unicos(p for _, h in leasing for p in h["plazos"]))
    fin = terms or estimadas
    if not leasing:
        r["leasing_estado"] = ""
    elif not fin:
        r["leasing_estado"] = "SIN FECHA - revisar"
    else:
        r["leasing_estado"] = "VIGENTE" if max(fin) >= HOY else "TERMINADO"
    return r


# ============================= Excel de salida =============================

ENCABEZADO_FILL = PatternFill("solid", fgColor="1F4E78")
ENCABEZADO_FONT = Font(bold=True, color="FFFFFF")
FILL_SI = PatternFill("solid", fgColor="E2EFDA")
FILL_VENCIDA = PatternFill("solid", fgColor="F8CBAD")
FILL_REVISAR = PatternFill("solid", fgColor="FFF2CC")


def _hoja(libro, titulo, encabezados, anchos=None):
    hoja = libro.create_sheet(titulo)
    hoja.append(encabezados)
    for i, celda in enumerate(hoja[1], start=1):
        celda.fill, celda.font = ENCABEZADO_FILL, ENCABEZADO_FONT
        celda.alignment = Alignment(wrap_text=True, vertical="center")
        hoja.column_dimensions[get_column_letter(i)].width = (anchos or {}).get(encabezados[i - 1], 18)
    hoja.freeze_panes = "B2"
    return hoja


def _link(celda, url, texto=None):
    if url:
        celda.value = texto or url
        if url.startswith("http"):
            celda.hyperlink = url
            celda.font = Font(color="0563C1", underline="single")


def _cerrar_hoja(hoja):
    if hoja.max_row > 1:
        hoja.auto_filter.ref = hoja.dimensions
    for fila in hoja.iter_rows(min_row=2):
        for c in fila:
            if isinstance(c.value, datetime.date):
                c.number_format = "DD/MM/YYYY"
            c.alignment = Alignment(vertical="top", wrap_text=isinstance(c.value, str) and len(c.value) > 40)


def escribir_excel(ruta, encabezados_originales, resultados, contadores):
    libro = openpyxl.Workbook()
    libro.remove(libro.active)

    # --- RESUMEN ---
    r = libro.create_sheet("RESUMEN")
    r.column_dimensions["A"].width, r.column_dimensions["B"].width = 55, 14
    r.append(["Depuracion de procesos Supersociedades: garantias FNG y LEASING"])
    r["A1"].font = Font(bold=True, size=14)
    r.append([f"Generado el {datetime.datetime.now():%d/%m/%Y %H:%M}"])
    r.append([])
    for etiqueta, valor in contadores:
        r.append([etiqueta, valor])
        r.cell(row=r.max_row, column=1).font = Font(bold=True)
    r.append([])
    r.append(["Nota: las fechas y numeros se extraen automaticamente del texto; confirma cada uno con el "
              "fragmento y el enlace de las hojas FNG_DETALLE y LEASING_DETALLE."])

    # --- PROCESOS ---
    nuevas = [
        "¿EN DRIVE?", "CARPETA(S) EN DRIVE", "ENLACE CARPETA", "ARCHIVOS EN CARPETA",
        "¿TIENE FNG?", "DOCS CON FNG", "N° GARANTIA FNG", "OBLIGACION / PAGARE", "ENTIDAD (FNG)",
        "VALOR / COBERTURA", "FECHA(S) VENCIMIENTO FNG", "PROXIMO VENCIMIENTO FNG", "ESTADO FNG",
        "¿TIENE LEASING?", "DOCS CON LEASING", "N° CONTRATO LEASING", "ENTIDAD (LEASING)", "BIEN",
        "FECHA(S) INICIO LEASING", "FECHA(S) TERMINACION LEASING", "PLAZO (MESES)", "ESTADO LEASING",
        "OBSERVACIONES",
    ]
    anchos = {"CARPETA(S) EN DRIVE": 35, "ENLACE CARPETA": 22, "FECHA(S) VENCIMIENTO FNG": 26,
              "FECHA(S) INICIO LEASING": 24, "FECHA(S) TERMINACION LEASING": 30, "OBSERVACIONES": 45,
              "N° GARANTIA FNG": 22, "N° CONTRATO LEASING": 22}
    anchos.update({e: 30 for e in encabezados_originales[:3]})
    hp = _hoja(libro, "PROCESOS", encabezados_originales + nuevas, anchos)
    hf = _hoja(libro, "FNG_DETALLE", [
        "CONCURSADO", "DOCUMENTO", "RUTA EN LA CARPETA", "ENLACE", "N° GARANTIA", "OBLIGACION / PAGARE",
        "ENTIDAD", "VALOR(ES)", "COBERTURA", "FECHA(S) VENCIMIENTO", "FECHA(S) EXPEDICION / DESEMBOLSO",
        "OTRAS FECHAS EN EL TEXTO", "MENCIONES", "FRAGMENTO",
    ], {"CONCURSADO": 30, "DOCUMENTO": 35, "RUTA EN LA CARPETA": 35, "FRAGMENTO": 80,
        "FECHA(S) VENCIMIENTO": 24, "OTRAS FECHAS EN EL TEXTO": 30})
    hl = _hoja(libro, "LEASING_DETALLE", [
        "CONCURSADO", "DOCUMENTO", "RUTA EN LA CARPETA", "ENLACE", "N° CONTRATO", "ENTIDAD", "BIEN",
        "FECHA(S) INICIO", "FECHA(S) TERMINACION", "TERMINACION ESTIMADA (INICIO + PLAZO)", "PLAZO (MESES)",
        "OTRAS FECHAS EN EL TEXTO", "MENCIONES", "FRAGMENTO",
    ], {"CONCURSADO": 30, "DOCUMENTO": 35, "RUTA EN LA CARPETA": 35, "FRAGMENTO": 80,
        "FECHA(S) INICIO": 22, "FECHA(S) TERMINACION": 24, "OTRAS FECHAS EN EL TEXTO": 30})
    hn = _hoja(libro, "NO_ENCONTRADOS_EN_DRIVE", [
        "CONCURSADO", "TIPO DE PROCESO", "NIT", "EXPEDIENTE", "CARPETA MAS PARECIDA", "SIMILITUD", "ENLACE",
    ], {"CONCURSADO": 40, "TIPO DE PROCESO": 24, "CARPETA MAS PARECIDA": 45, "ENLACE": 30})
    hr = _hoja(libro, "REVISAR_A_MANO", [
        "CONCURSADO", "DOCUMENTO", "RUTA EN LA CARPETA", "ENLACE", "MOTIVO",
    ], {"CONCURSADO": 30, "DOCUMENTO": 40, "RUTA EN LA CARPETA": 40, "MOTIVO": 45})

    columna_estado_fng = len(encabezados_originales) + nuevas.index("ESTADO FNG") + 1
    columna_estado_leasing = len(encabezados_originales) + nuevas.index("ESTADO LEASING") + 1
    columna_tiene_fng = len(encabezados_originales) + nuevas.index("¿TIENE FNG?") + 1
    columna_tiene_leasing = len(encabezados_originales) + nuevas.index("¿TIENE LEASING?") + 1
    columna_enlace = len(encabezados_originales) + nuevas.index("ENLACE CARPETA") + 1

    for res in resultados:
        fila_original = list(res["fila"]) + [None] * (len(encabezados_originales) - len(res["fila"]))
        fila_original = fila_original[:len(encabezados_originales)]
        s = res.get("resumen") or {}
        en_drive = bool(res["carpetas"])
        hp.append(fila_original + [
            "SI" if en_drive else "NO",
            " | ".join(c["nombre"] for c in res["carpetas"]),
            None,
            res.get("total_archivos", ""),
            s.get("fng", ""), s.get("fng_docs", ""), s.get("fng_numeros", ""), s.get("fng_obligaciones", ""),
            s.get("fng_entidades", ""), s.get("fng_valores", ""), s.get("fng_vencimientos", ""),
            s.get("fng_proximo"), s.get("fng_estado", ""),
            s.get("leasing", ""), s.get("leasing_docs", ""), s.get("leasing_contratos", ""),
            s.get("leasing_entidades", ""), s.get("leasing_bienes", ""), s.get("leasing_inicios", ""),
            s.get("leasing_terminaciones", ""), s.get("leasing_plazos", ""), s.get("leasing_estado", ""),
            res.get("observaciones", ""),
        ])
        fila = hp.max_row
        if res["carpetas"]:
            _link(hp.cell(row=fila, column=columna_enlace), res["carpetas"][0]["enlace"], "Abrir carpeta")
        for col, valor in ((columna_tiene_fng, s.get("fng")), (columna_tiene_leasing, s.get("leasing"))):
            if valor == "SI":
                hp.cell(row=fila, column=col).fill = FILL_SI
        for col in (columna_estado_fng, columna_estado_leasing):
            v = hp.cell(row=fila, column=col).value or ""
            if v in ("VENCIDA", "TERMINADO"):
                hp.cell(row=fila, column=col).fill = FILL_VENCIDA
            elif v.startswith("SIN FECHA"):
                hp.cell(row=fila, column=col).fill = FILL_REVISAR
            elif v == "VIGENTE":
                hp.cell(row=fila, column=col).fill = FILL_SI

        for a, h in res.get("fng", []):
            hf.append([res["nombre"], a["nombre"], f"{a['carpeta']}/{a['ruta']}", None,
                       ", ".join(h["numeros"]), ", ".join(h["obligaciones"]), ", ".join(h["entidades"]),
                       ", ".join(h["valores"]), ", ".join(h["coberturas"]), fmt_fechas(h["vencimientos"]),
                       fmt_fechas(h["expediciones"]), fmt_fechas(h["otras_fechas"]), h["menciones"],
                       h["fragmento"] + ("  [SIN TEXTO: solo se detecto por el nombre]" if h["sin_texto"] else "")])
            _link(hf.cell(row=hf.max_row, column=4), a["enlace"], "Abrir")
        for a, h in res.get("leasing", []):
            hl.append([res["nombre"], a["nombre"], f"{a['carpeta']}/{a['ruta']}", None,
                       ", ".join(h["contratos"]), ", ".join(h["entidades"]), ", ".join(h["bienes"]),
                       fmt_fechas(h["inicios"]), fmt_fechas(h["terminaciones"]),
                       fmt_fechas(h["terminaciones_estimadas"]), ", ".join(map(str, h["plazos"])),
                       fmt_fechas(h["otras_fechas"]), h["menciones"],
                       h["fragmento"] + ("  [SIN TEXTO: solo se detecto por el nombre]" if h["sin_texto"] else "")])
            _link(hl.cell(row=hl.max_row, column=4), a["enlace"], "Abrir")
        if not en_drive:
            mejor = res.get("alternativas") or []
            p, c = mejor[0] if mejor else (None, None)
            hn.append([res["nombre"], res.get("tipo", ""), res["nit"], ", ".join(res.get("expedientes", [])),
                       c["nombre"] if c else "", round(p, 2) if p else "", None])
            if c:
                _link(hn.cell(row=hn.max_row, column=7), c["enlace"], "Abrir")
        for a, motivo in res.get("sin_leer", []):
            hr.append([res["nombre"], a["nombre"], f"{a['carpeta']}/{a['ruta']}", None, motivo])
            _link(hr.cell(row=hr.max_row, column=4), a["enlace"], "Abrir")

    for hoja in (hp, hf, hl, hn, hr):
        _cerrar_hoja(hoja)
    libro.save(ruta)


# ============================= main =============================


def pedir_ruta_excel():
    if len(sys.argv) > 1:
        return sys.argv[1].strip().strip('"')
    if RUTA_EXCEL_PROCESOS:
        return RUTA_EXCEL_PROCESOS
    # Si en la carpeta del script hay UN solo Excel (que no sea un
    # resultado anterior), se usa ese sin preguntar.
    candidatos = [
        p for p in Path(DIRECTORIO).glob("*.xlsx")
        if not p.name.lower().startswith(("depuracion_fng_leasing", "~$"))
    ]
    if len(candidatos) == 1:
        return str(candidatos[0])
    return input("Arrastra aqui el Excel con el listado de procesos y presiona Enter: ").strip().strip('"')


def procesar():
    ruta_excel = pedir_ruta_excel()
    if not ruta_excel or not os.path.exists(ruta_excel):
        raise RuntimeError(f"No existe el Excel de procesos: '{ruta_excel}'")
    info = leer_listado(ruta_excel)
    encabezados, filas = info["encabezados"], info["filas"]
    idx_conc, idx_nit, idx_exp, idx_tipo = info["idx_conc"], info["idx_nit"], info["idx_exp"], info["idx_tipo"]
    logging.info("Excel: %s", ruta_excel)
    logging.info("   Columna concursado: '%s'%s%s", encabezados[idx_conc],
                 f" | NIT: '{encabezados[idx_nit]}'" if idx_nit is not None else "",
                 f" | expediente: '{encabezados[idx_exp]}'" if idx_exp is not None else "")
    if idx_tipo is not None and TIPOS_PROCESO_A_INCLUIR:
        logging.info("   Filtro por '%s': %d de %d filas son de insolvencia/reorganizacion y se van a revisar.",
                     encabezados[idx_tipo], len(filas), info["total_filas"])
    else:
        logging.info("   %d procesos a revisar.", len(filas))

    if CARPETA_LOCAL_DRIVE.strip():
        fuente = FuenteLocal(CARPETA_LOCAL_DRIVE.strip())
        logging.info("Leyendo el Drive desde la carpeta local %s", CARPETA_LOCAL_DRIVE)
    else:
        fuente = FuenteDriveAPI()
        logging.info("Conectado a Google Drive (API).")

    firma = "|".join([
        os.path.basename(ruta_excel), str(len(filas)), CARPETA_LOCAL_DRIVE.strip(), CARPETA_RAIZ_DRIVE.strip(),
        MODO_LECTURA, ",".join(TIPOS_PROCESO_A_INCLUIR), str(PROFUNDIDAD_BUSQUEDA_CARPETAS),
    ])
    progreso = Progreso(ARCHIVO_PROGRESO, firma)

    carpetas = progreso.get("carpetas")
    if carpetas is None:
        carpetas = fuente.listar_carpetas_candidatas()
        progreso.poner("carpetas", carpetas)
    else:
        logging.info("(Lista de carpetas del Drive tomada del avance guardado.)")
    padres_de = {c["id"]: c["padres"] for c in carpetas}
    logging.info("%d carpetas candidatas en el Drive.", len(carpetas))

    ids_indexados = None
    if MODO_LECTURA == "rapido" and isinstance(fuente, FuenteDriveAPI):
        if "ids_indexados" in progreso.datos:
            guardados = progreso.get("ids_indexados")
            ids_indexados = set(guardados) if guardados is not None else None
        else:
            logging.info("Preguntando al buscador de Drive que documentos mencionan FNG / leasing...")
            ids_indexados = fuente.ids_con_texto(
                ["FNG", "Fondo Nacional de Garantías", "leasing", "arrendamiento financiero"]
            )
            progreso.poner("ids_indexados", ids_indexados)
        if ids_indexados is not None:
            logging.info("   %d documentos del Drive mencionan FNG / leasing.", len(ids_indexados))

    cache = Cache(ARCHIVO_CACHE)
    resultados = []
    try:
        for n, fila in enumerate(filas, start=1):
            nombre = str(fila[idx_conc]).strip()
            nit = nit_base(fila[idx_nit]) if idx_nit is not None and idx_nit < len(fila) else ""
            expedientes = numeros_expediente(fila[idx_exp]) if idx_exp is not None and idx_exp < len(fila) else []
            tipo = str(fila[idx_tipo] or "").strip() if idx_tipo is not None and idx_tipo < len(fila) else ""
            elegidas, alternativas = buscar_carpetas_del_proceso(nombre, nit, expedientes, carpetas, padres_de)
            res = {"fila": fila, "nombre": nombre, "nit": nit, "expedientes": expedientes, "tipo": tipo,
                   "carpetas": elegidas, "alternativas": alternativas}
            observaciones = []
            if not elegidas:
                logging.info("[%d/%d] %s -> NO esta en el Drive", n, len(filas), nombre)
                resultados.append(res)
                continue
            if len(elegidas) > 1:
                observaciones.append(f"Se revisaron {len(elegidas)} carpetas con nombre parecido")
            if alternativas:
                observaciones.append("Otras carpetas parecidas: " + "; ".join(c["nombre"] for _, c in alternativas))
            clave = "|".join([nombre, tipo, ",".join(expedientes), ",".join(c["id"] for c in elegidas)])
            hecho = progreso.proceso(clave)
            if hecho is not None:
                fng, leasing, sin_leer, total = hecho["fng"], hecho["leasing"], hecho["sin_leer"], hecho["total"]
                ya_hecho = True
            else:
                ya_hecho = False
                for intento in range(1, REINTENTOS_POR_PROCESO + 1):
                    try:
                        fng, leasing, sin_leer, total = revisar_proceso(fuente, cache, elegidas, ids_indexados)
                        progreso.marcar_proceso(clave, {"fng": fng, "leasing": leasing, "sin_leer": sin_leer,
                                                        "total": total})
                        break
                    except Exception as e:  # noqa: BLE001
                        if intento < REINTENTOS_POR_PROCESO:
                            logging.warning("   Fallo revisando %s (%s). Reintento %d/%d en %d s...", nombre,
                                            str(e)[:150], intento, REINTENTOS_POR_PROCESO - 1, ESPERA_REINTENTO_SEG)
                            time.sleep(ESPERA_REINTENTO_SEG)
                            continue
                        logging.exception("   Error revisando %s", nombre)
                        fng, leasing, sin_leer, total = [], [], [], 0
                        observaciones.append(f"ERROR al revisar la carpeta (vuelve a ejecutar para reintentar): {e}")
            if sin_leer:
                observaciones.append(f"{len(sin_leer)} documento(s) relevante(s) sin poder leer (ver REVISAR_A_MANO)")
            res.update(fng=fng, leasing=leasing, sin_leer=sin_leer, total_archivos=total,
                       resumen=resumir_proceso(fng, leasing), observaciones=". ".join(observaciones))
            s = res["resumen"]
            logging.info("[%d/%d] %s%s -> carpeta '%s' (%d archivos) | FNG: %s %s | LEASING: %s %s",
                         n, len(filas), nombre, " (ya revisado)" if ya_hecho else "", elegidas[0]["nombre"], total,
                         s["fng"], s["fng_vencimientos"], s["leasing"], s["leasing_terminaciones"])
            resultados.append(res)
    finally:
        cache.guardar()

    en_drive = [r for r in resultados if r["carpetas"]]
    con_fng = [r for r in en_drive if r.get("resumen", {}).get("fng") == "SI"]
    con_leasing = [r for r in en_drive if r.get("resumen", {}).get("leasing") == "SI"]
    contadores = []
    if idx_tipo is not None and TIPOS_PROCESO_A_INCLUIR:
        contadores.append(("Filas del listado original", info["total_filas"]))
        contadores.append(("Procesos de insolvencia/reorganizacion (filtrados)", len(resultados)))
        for t, cuantos in sorted(
            {normalizar(r["tipo"]): sum(1 for x in resultados if normalizar(x["tipo"]) == normalizar(r["tipo"]))
             for r in resultados}.items()
        ):
            contadores.append((f"   {t}", cuantos))
    else:
        contadores.append(("Procesos en el listado", len(resultados)))
    contadores += [
        ("Procesos con carpeta en el Drive", len(en_drive)),
        ("Procesos NO encontrados en el Drive", len(resultados) - len(en_drive)),
        ("Procesos con garantia FNG", len(con_fng)),
        ("   FNG vigente", sum(1 for r in con_fng if r["resumen"]["fng_estado"] == "VIGENTE")),
        ("   FNG vencida", sum(1 for r in con_fng if r["resumen"]["fng_estado"] == "VENCIDA")),
        ("   FNG sin fecha de vencimiento (revisar)", sum(1 for r in con_fng if r["resumen"]["fng_estado"].startswith("SIN"))),
        ("Procesos con LEASING", len(con_leasing)),
        ("   Leasing vigente", sum(1 for r in con_leasing if r["resumen"]["leasing_estado"] == "VIGENTE")),
        ("   Leasing terminado", sum(1 for r in con_leasing if r["resumen"]["leasing_estado"] == "TERMINADO")),
        ("   Leasing sin fecha (revisar)", sum(1 for r in con_leasing if r["resumen"]["leasing_estado"].startswith("SIN"))),
        ("Procesos con FNG y LEASING", sum(1 for r in con_fng if r in con_leasing)),
        ("Documentos relevantes sin poder leer (REVISAR_A_MANO)", sum(len(r.get("sin_leer", [])) for r in resultados)),
    ]
    escribir_excel(ARCHIVO_SALIDA, encabezados, resultados, contadores)
    if not any("ERROR al revisar" in (r.get("observaciones") or "") for r in resultados):
        progreso.borrar()  # corrida completa: la proxima vez se empieza de cero (con el cache de textos)
    else:
        logging.info("Hubo procesos con ERROR: si vuelves a ejecutar, solo se reintentan esos.")
    logging.info("")
    for etiqueta, valor in contadores:
        logging.info("%-58s %s", etiqueta, valor)
    logging.info("")
    logging.info("Listo. Resultado en: %s", ARCHIVO_SALIDA)
    if sys.platform.startswith("win"):
        try:
            os.startfile(ARCHIVO_SALIDA)  # abre el Excel de resultado
        except OSError:
            pass


def main():
    configurar_logging()
    mantener_pc_despierto()
    try:
        procesar()
    except Exception as e:  # noqa: BLE001
        logging.error("ERROR: %s", e)
        sys.exit(1)


if __name__ == "__main__":
    main()
