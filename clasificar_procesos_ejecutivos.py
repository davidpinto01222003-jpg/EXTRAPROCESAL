"""
Especializa la búsqueda en Google Drive para el informe
"3. CONTROL PROCESOS EJECUTIVOS ESSA...xlsm" (hoja `HOJA_EXCEL_CONTROL`,
por defecto `DatosProcesados1` -- la versión de la hoja `ACTIVOS` ya
aplanada a un solo encabezado por fila).

A diferencia de `buscar_faltantes_en_drive.py` (que descarga TODO el
contenido relacionado con el radicado), este script trabaja FILA POR
FILA del Excel (no agrupa por número de proceso) y aplica dos reglas
de negocio según el `ESTADO PROCESAL` de cada fila:

1. Cualquier fila que NO sea "terminada" (ver regla 2) -- `ACTIVO`,
   `ACTIVOS CON TITULOS`, `SUSPENDIDO`, `REORGANIZACION`, `REMITIDA A
   CASTIGO`/`A PREPAGO`, etc: la carpeta se nombra `"<numero>.
   <radicado>"` si la fila ya tiene un radicado válido de 23 dígitos,
   o `"<numero>. <ESTADO PROCESAL>"` si todavía no lo tiene (mismo
   formato que la regla 2). SOLO se sube información NO procesal:
   tutelas, derechos de petición, y correos de cobro (ver
   `es_informacion_no_procesal`) -- se busca tanto en Google Drive como
   (opcional, ver `BUSCAR_EN_CORREO`) en TODO el Gmail.

2. Procesos terminados por pago, por auto, por contrato/prepago, o que
   nunca se presentaron (`ESTADO PROCESAL` que empieza con `TERMINADO`
   o `NO INICIO` -- igual que `crear_carpetas_terminados_castigo.py`):
   la carpeta se nombra `"<numero>. <ESTADO PROCESAL EXACTO del Excel>"`
   (ej. `"245. TERMINADO POR AUTO"`), y SOLO se sube el documento que
   deja constancia de que el proceso NO sigue su curso -- un auto de
   terminación, de aceptación de retiro de la demanda, de
   desistimiento, etc (ver `es_auto_terminador`, no exige que aparezca
   literalmente la palabra "AUTO"). Si no se encuentra ese documento en
   Drive, el proceso queda listado en `ARCHIVO_PENDIENTES_TERMINADOS`
   para que lo descargues a mano.

IMPORTANTE -- procesos "acumulados" y filas duplicadas: el Excel repite
el mismo número de proceso en más de una fila en dos casos distintos:
  - Cuentas/demandados distintos bajo el MISMO radicado (proceso
    "acumulado": un solo expediente judicial que agrupa varias cuentas).
  - Menos frecuente: el mismo número de proceso con un radicado
    DISTINTO en cada fila (numeración administrativa repetida por
    error o por reuso, no es el mismo expediente).
En AMBOS casos cada fila del Excel se organiza en SU PROPIA carpeta
(nunca se fusionan) -- si dos filas producen el mismo nombre de carpeta
(mismo número + mismo radicado/estado), la segunda (y siguientes) se
numeran "<nombre>_2", "<nombre>_3", etc, igual que el resto del
proyecto nombra duplicados. Ver `_asignar_nombres_de_carpeta`.

La clasificación de "información no procesal" y de "documento que
termina el proceso" es por PALABRAS CLAVE (nombre/asunto y, si es
PDF/DOCX o el cuerpo de un correo, su contenido) -- es una heurística,
no perfecta. Cada decisión queda registrada en el log para que la
revises y ajustes las listas de palabras clave
(`PALABRAS_TIPO_INFORMACION_FUERTES`, `PALABRAS_CORREO_DE_COBRO`,
`PALABRAS_PROCESO_NO_CONTINUA`) si hace falta.

Reutiliza toda la infraestructura de `buscar_faltantes_en_drive.py`
(autenticación de Drive, búsqueda por radicado/radicado corto/cuenta,
validación de que el documento sea de ESSA y del demandado correcto,
descarga de PDF/exportables de Google) -- no repite esa lógica. Para
Gmail, reutiliza el mismo mecanismo de búsqueda que ese script (IMAP
`X-GM-RAW` sobre "Todos los mensajes", igual que buscar en la barra de
búsqueda de Gmail), pero además revisa el CUERPO del correo y CUALQUIER
adjunto (no solo enlaces de Drive o adjuntos .zip).

Requiere las mismas credenciales que `buscar_faltantes_en_drive.py`:
`credenciales_drive.json`/`token_drive.json` para Drive (ver README), y
`credenciales_sgde.txt` para Gmail (opcional -- si no existe, la
búsqueda en correo simplemente se omite, ver `BUSCAR_EN_CORREO`).

Respeta MODO_PRUEBA (por defecto True): en modo prueba solo BUSCA y
CLASIFICA, mostrando qué subiría y a qué carpeta, sin crear carpetas ni
descargar nada todavía.
"""

import csv
import email
import imaplib
import io
import logging
import os
import re
import zipfile
from pathlib import Path

import openpyxl

import buscar_faltantes_en_drive as buscador
import crear_carpetas_terminados_castigo as terminados_folder
import procesos_juridicos as organizador
import validar_renombrar_carpetas as cruce_excel

# ============================= CONFIGURACION =============================

# Ruta al informe de Excel de procesos ejecutivos (.xlsm).
RUTA_EXCEL_CONTROL = r"C:\Users\Francy\OneDrive\INFORME ENTREGA ESSA\3. CONTROL PROCESOS EJECUTIVOS ESSA 29072026 .xlsm"

# Hoja del Excel a leer -- "DatosProcesados1" ya trae, en un solo
# encabezado por fila, todas las columnas que hacen falta (No., ESTADO
# PROCESAL, CUENTA, DEMANDADO, JUZGADO, RADICADO), a diferencia de la
# hoja "ACTIVOS" (que trae los mismos datos pero repartidos en varias
# filas de encabezado y cientos de columnas de honorarios/costas).
HOJA_EXCEL_CONTROL = "DatosProcesados1"
FILA_ENCABEZADO_CONTROL = 1

COLUMNA_NO = "No."
COLUMNA_ESTADO = "ESTADO PROCESAL"
COLUMNA_CUENTA = "CUENTA"
COLUMNA_DEMANDADO = "DEMANDADO"
COLUMNA_JUZGADO = "JUZGADO"
COLUMNA_RADICADO = "RADICADO"

# Prefijos de ESTADO PROCESAL que cuentan como "terminado" para este
# script -- igual que crear_carpetas_terminados_castigo.py: TERMINADO
# (por pago, por auto, por contrato/prepago) y NO INICIO (nunca se
# presentó la demanda, quedó como reclamación administrativa). CUALQUIER
# otro estado (ACTIVO, SUSPENDIDO, REORGANIZACION, REMITIDA A CASTIGO/
# PREPAGO, etc) se organiza con la carpeta "numero. radicado" de siempre.
PREFIJOS_ESTADO_TERMINADO = ("TERMINADO", "NO INICIO")

CARPETA_PROCESOS = cruce_excel.CARPETA_PROCESOS

ARCHIVO_LOG = os.path.join(os.path.dirname(__file__), "clasificar_procesos_ejecutivos.log")

# Procesos terminados (por pago/auto/contrato/no inicio) en los que NO
# se encontró el documento que los termina -- para que los descargues a mano.
ARCHIVO_PENDIENTES_TERMINADOS = os.path.join(os.path.dirname(__file__), "terminados_sin_auto_pendientes.csv")

# True (por defecto): no crea carpetas ni descarga nada, solo busca y
# muestra qué haría. False: aplica los cambios de verdad.
MODO_PRUEBA = True

# True (por defecto): ademas de Drive, busca en TODO tu Gmail (via IMAP,
# usando credenciales_sgde.txt -- ver README) tutelas, derechos de
# peticion y correos de cobro relacionados con cada proceso. Si ese
# archivo de credenciales no existe, esta busqueda se omite sola, sin
# error.
BUSCAR_EN_CORREO = True

# --------------------- Palabras clave de clasificación ---------------------

# Frases que, si aparecen en el contenido (nombre de archivo, texto del
# documento, o asunto/cuerpo de un correo), confirman que es un derecho
# de petición o una tutela. A propósito NO se incluye "SOLICITUD": es
# una palabra demasiado común dentro de cualquier memorial procesal
# ("se solicita...") y generaba muchas imprecisiones.
PALABRAS_TIPO_INFORMACION_FUERTES = [
    "DERECHO DE PETICION",
    "ACCION DE TUTELA",
    "TUTELA",
]

# Frases que identifican un correo/documento de COBRO (a un banco, EPS,
# municipio, etc, reclamando el pago de la obligación) -- cuenta igual
# que una tutela o un derecho de petición como "información no procesal".
PALABRAS_CORREO_DE_COBRO = [
    "COBRO PREJURIDICO", "CUENTA DE COBRO", "ESTADO DE CUENTA",
    "RECORDATORIO DE PAGO", "REQUERIMIENTO DE PAGO", "NOTIFICACION DE COBRO",
    "GESTION DE COBRO", "COBRO DE CARTERA", "COBRO ADMINISTRATIVO",
]

# Frases que dan a entender que el proceso judicial NO sigue su curso
# (terminó por pago, se aceptó el retiro de la demanda, se decretó
# desistimiento o archivo, etc) -- NO hace falta que el documento diga
# literalmente "AUTO", con que aparezca cualquiera de estas frases
# alcanza (ver es_auto_terminador).
PALABRAS_PROCESO_NO_CONTINUA = [
    "TERMINA EL PROCESO", "TERMINACION DEL PROCESO", "SE DA POR TERMINADO",
    "TERMINADO EL PROCESO", "PROCESO TERMINADO", "DECLARA TERMINADO EL PROCESO",
    "TERMINACION POR PAGO", "TERMINACION POR CONTRATO", "TERMINACION DE LA OBLIGACION",
    "DECRETA LA TERMINACION", "DECRETA TERMINACION",
    "ACEPTA EL RETIRO", "ACEPTA RETIRO DE LA DEMANDA", "RETIRO DE LA DEMANDA",
    "SE RETIRA LA DEMANDA", "DESISTIMIENTO DEL PROCESO", "DESISTIMIENTO DE LA DEMANDA",
    "APRUEBA EL DESISTIMIENTO", "ARCHIVA EL PROCESO", "ARCHIVESE EL PROCESO",
    "ARCHIVO DEL PROCESO", "CULMINA EL PROCESO", "FINALIZA EL PROCESO",
    "NO CONTINUA EL PROCESO", "CESE DE LA ACCION EJECUTIVA", "DA POR CONCLUIDO EL PROCESO",
]

# ===========================================================================


def configurar_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
        handlers=[
            logging.FileHandler(ARCHIVO_LOG, encoding="utf-8", mode="w"),
            logging.StreamHandler(),
        ],
    )


def _normalizar_radicado(valor):
    if valor is None:
        return None
    texto = re.sub(r"[\s\-]", "", str(valor).strip())
    return texto if texto.isdigit() and len(texto) == 23 else None


def _encontrar_columna(encabezados, nombre_buscado):
    for idx, valor in enumerate(encabezados):
        if valor and str(valor).strip().upper() == nombre_buscado.strip().upper():
            return idx
    raise ValueError(
        f"No se encontro la columna '{nombre_buscado}' en la fila {FILA_ENCABEZADO_CONTROL} "
        f"de la hoja '{HOJA_EXCEL_CONTROL}'."
    )


def leer_procesos_control():
    """
    Lee HOJA_EXCEL_CONTROL FILA POR FILA (sin agrupar por No. -- ver
    módulo docstring sobre procesos acumulados). Devuelve una lista de
    {fila_excel, numero, estado, radicado, cuenta, demandado, juzgado}.
    """
    wb = openpyxl.load_workbook(RUTA_EXCEL_CONTROL, data_only=True, read_only=True)
    if HOJA_EXCEL_CONTROL not in wb.sheetnames:
        raise ValueError(
            f"La hoja '{HOJA_EXCEL_CONTROL}' no existe en {RUTA_EXCEL_CONTROL}. "
            f"Hojas disponibles: {wb.sheetnames}"
        )
    ws = wb[HOJA_EXCEL_CONTROL]
    filas = ws.iter_rows(min_row=FILA_ENCABEZADO_CONTROL, max_row=FILA_ENCABEZADO_CONTROL, values_only=True)
    encabezados = next(filas)

    idx_no = _encontrar_columna(encabezados, COLUMNA_NO)
    idx_estado = _encontrar_columna(encabezados, COLUMNA_ESTADO)
    idx_cuenta = _encontrar_columna(encabezados, COLUMNA_CUENTA)
    idx_demandado = _encontrar_columna(encabezados, COLUMNA_DEMANDADO)
    idx_juzgado = _encontrar_columna(encabezados, COLUMNA_JUZGADO)
    idx_radicado = _encontrar_columna(encabezados, COLUMNA_RADICADO)

    procesos = []
    for fila_excel, fila in enumerate(
        ws.iter_rows(min_row=FILA_ENCABEZADO_CONTROL + 1, values_only=True),
        start=FILA_ENCABEZADO_CONTROL + 1,
    ):
        numero_crudo = fila[idx_no]
        if numero_crudo is None or not isinstance(numero_crudo, (int, float)):
            continue  # fila vacia o de notas/leyenda al final de la hoja
        numero = int(numero_crudo)

        estado = str(fila[idx_estado]).strip() if fila[idx_estado] is not None else ""
        cuenta = str(fila[idx_cuenta]).strip() if fila[idx_cuenta] is not None else ""
        demandado = str(fila[idx_demandado]).strip() if fila[idx_demandado] is not None else ""
        juzgado = str(fila[idx_juzgado]).strip() if fila[idx_juzgado] is not None else ""
        radicado = _normalizar_radicado(fila[idx_radicado])

        procesos.append({
            "fila_excel": fila_excel,
            "numero": numero,
            "estado": estado,
            "radicado": radicado,
            "cuenta": cuenta if cuenta.strip("0") else "",
            "demandado": demandado,
            "juzgado": juzgado,
        })
    return procesos


def _asignar_nombres_de_carpeta(procesos):
    """
    Asigna a cada proceso (fila) su nombre de carpeta final a partir de
    "nombre_base", agregando "_2", "_3", etc cuando dos filas DISTINTAS
    (ej. un proceso acumulado: mismo número y mismo radicado en dos
    cuentas) producen el mismo nombre -- nunca se fusionan, cada fila
    del Excel es una carpeta propia (ver módulo docstring).
    """
    contador = {}
    for proceso in procesos:
        base = proceso["nombre_base"]
        contador[base] = contador.get(base, 0) + 1
        indice = contador[base]
        proceso["nombre_carpeta"] = base if indice == 1 else f"{base}_{indice}"


def clasificar_procesos(procesos):
    """
    Separa las filas leidas en dos listas de trabajo: con_radicado
    (todo lo que no es "terminado" -- se busca información no procesal)
    y terminados (se busca el documento que termina el proceso), mas
    una lista informativa sin_estado. Toda fila con ESTADO PROCESAL
    diligenciado recibe una carpeta: con su radicado si ya lo tiene, o
    con "numero. ESTADO PROCESAL" si todavia no (igual que las filas
    terminadas) -- ver módulo docstring.
    """
    con_radicado = []
    terminados = []
    sin_estado = []

    for proceso in procesos:
        if not proceso["estado"]:
            sin_estado.append(proceso)
            continue

        numero, estado = proceso["numero"], proceso["estado"]
        proceso = dict(proceso)
        proceso["cuentas"] = [proceso["cuenta"]] if proceso["cuenta"] else []
        proceso["demandados"] = [proceso["demandado"]] if proceso["demandado"] else []

        if estado.upper().startswith(PREFIJOS_ESTADO_TERMINADO):
            proceso["nombre_base"] = terminados_folder._nombre_carpeta_para(numero, estado) or f"{numero}. {estado}"
            terminados.append(proceso)
        else:
            proceso["nombre_base"] = (
                f"{numero}. {proceso['radicado']}" if proceso["radicado"] else f"{numero}. {estado}"
            )
            con_radicado.append(proceso)

    _asignar_nombres_de_carpeta(con_radicado)
    _asignar_nombres_de_carpeta(terminados)
    return con_radicado, terminados, sin_estado


# ==================== Google Drive (usa buscar_faltantes_en_drive.py) ====================


def autenticar_drive_o_none():
    try:
        return buscador.autenticar_drive()
    except Exception as error:
        logging.error("[Drive] No se pudo conectar con Google Drive: %s", error)
        return None


def _terminos_busqueda(radicado, cuentas):
    terminos = []
    if radicado:
        terminos.append(radicado)
        terminos.extend(buscador.radicados_cortos(radicado))
    for cuenta in cuentas:
        if buscador._cuenta_es_valida_para_buscar(cuenta):
            terminos.append(cuenta)
    return terminos


def _candidatos_en_drive(servicio, terminos):
    vistos = {}
    for termino in terminos:
        for item in buscador.buscar_en_drive(servicio, termino):
            vistos[item["id"]] = item
    return list(vistos.values())


def _archivos_de_candidatos(servicio, candidatos, radicado):
    """
    A partir de los candidatos de Drive (carpetas o archivos sueltos que
    coincidieron con el radicado/cuenta), devuelve [(archivo, carpeta), ...]
    de los PDF/exportables relevantes -- reutilizando el mismo criterio
    de "carpeta propia del caso" que buscar_faltantes_en_drive.py.
    """
    vistos = {}
    for item in candidatos:
        carpeta = buscador.carpeta_contenedora(servicio, item)
        if not carpeta:
            continue

        es_propia = item.get("mimeType") == buscador.MIME_CARPETA or buscador._carpeta_es_dedicada_al_caso(
            item, carpeta, radicado or ""
        )
        if es_propia:
            if radicado:
                archivos = buscador._archivos_relacionados(servicio, carpeta, radicado)
            else:
                # Sin radicado real (ej. NO INICIO): no hay como filtrar
                # por nombre, se listan los PDF de primer nivel de la
                # carpeta tal cual -- la validacion de ESSA/demandado
                # (ver _archivo_es_seguro) sigue aplicando igual.
                try:
                    respuesta = servicio.files().list(
                        q=f"'{carpeta['id']}' in parents and trashed = false",
                        fields="files(id, name, mimeType)",
                    ).execute()
                except buscador.HttpError:
                    respuesta = {}
                archivos = [
                    a for a in respuesta.get("files", [])
                    if a.get("mimeType") != buscador.MIME_CARPETA and buscador._es_pdf_o_exportable(a)
                ]
        elif item.get("mimeType") != buscador.MIME_CARPETA and buscador._es_pdf_o_exportable(item):
            archivos = [item]
        else:
            archivos = []

        for archivo in archivos:
            vistos[archivo["id"]] = (archivo, carpeta)
    return list(vistos.values())


def _archivo_es_seguro(servicio, archivo, nombre_carpeta, demandados):
    """
    Aplica la regla obligatoria de buscar_faltantes_en_drive.py: el
    documento (o su carpeta contenedora) tiene que mencionar a ESSA/
    Electrificadora de Santander, y no debe contradecir al DEMANDADO
    esperado del Excel.
    """
    contextos = [nombre_carpeta, archivo.get("name", "")]
    tiene_essa = any(
        buscador._nombre_coincide(texto, termino)
        for texto in contextos for termino in buscador.TERMINOS_DEMANDANTE_VALIDO
    )
    if not tiene_essa:
        tiene_essa = buscador._archivos_tienen_demandante_valido(servicio, [archivo])
    if not tiene_essa:
        return False, "no se confirmo que el proceso sea de ESSA/Electrificadora de Santander"

    for demandado in demandados:
        resultado = buscador._demandado_coincide_en_varios(contextos, demandado)
        if resultado is False:
            return False, f"el nombre en el documento no corresponde al demandado esperado ({demandado})"

    return True, "ok"


def _info_documento(servicio, archivo):
    """Contenido normalizado (nombre + texto de PDF/DOCX si aplica) -- ver funciones de clasificacion mas abajo."""
    nombre = archivo.get("name", "")
    texto = ""
    if Path(nombre).suffix.lower() in buscador.EXTENSIONES_CONTENIDO_DRIVE:
        texto = buscador._texto_de_archivo_drive(servicio, archivo)
    return buscador._normalizar_para_comparar(nombre + " " + texto)


def es_informacion_no_procesal(contenido_normalizado):
    """
    True si el contenido (nombre/texto de un documento, o asunto/cuerpo
    de un correo) parece ser una tutela, un derecho de petición, o un
    correo/documento de cobro.
    """
    if any(frase in contenido_normalizado for frase in PALABRAS_TIPO_INFORMACION_FUERTES):
        return True, "peticion o tutela"
    if any(frase in contenido_normalizado for frase in PALABRAS_CORREO_DE_COBRO):
        return True, "correo de cobro"
    return False, "no parece una tutela, un derecho de peticion ni un correo de cobro"


def es_auto_terminador(contenido_normalizado):
    """
    True si el contenido da a entender que el proceso judicial NO sigue
    su curso (terminación por pago, retiro/desistimiento de la demanda,
    archivo del proceso, etc) -- no exige que aparezca la palabra
    "AUTO", con que aparezca cualquiera de PALABRAS_PROCESO_NO_CONTINUA
    alcanza.
    """
    return any(frase in contenido_normalizado for frase in PALABRAS_PROCESO_NO_CONTINUA)


# ==================== Gmail (busqueda propia, no la de buscar_faltantes_en_drive.py) ====================


def _texto_plano_html(texto: str) -> str:
    """Quita etiquetas HTML de forma simple, solo para poder buscar palabras clave adentro."""
    return re.sub(r"<[^>]+>", " ", texto or "")


def buscar_correo_informacion_no_procesal(usuario: str, app_password: str, termino: str):
    """
    Busca en TODO el Gmail (via X-GM-RAW, igual que buscar_en_correo de
    buscar_faltantes_en_drive.py -- es la misma busqueda que escribir
    'termino' en la barra de busqueda de Gmail, sobre "Todos los
    mensajes") los correos que mencionen 'termino'. A diferencia de esa
    función (que solo mira enlaces de Drive y adjuntos .zip), esta
    devuelve TODO lo que hace falta para clasificar un correo como
    tutela/derecho de petición/correo de cobro: asunto, cuerpo de texto
    (HTML ya limpiado), y CUALQUIER adjunto (no solo .zip). Devuelve
    [{"asunto", "cuerpo", "adjuntos": [(nombre, bytes), ...], "fecha"}, ...].
    """
    resultados = []
    with imaplib.IMAP4_SSL("imap.gmail.com") as mail:
        mail.login(usuario, app_password)
        mail.select('"[Gmail]/All Mail"', readonly=True)
        typ, datos = mail.search(None, "X-GM-RAW", f'"{termino}"')
        if typ != "OK" or not datos or not datos[0]:
            return resultados
        for id_correo in datos[0].split():
            typ, msg_datos = mail.fetch(id_correo, "(RFC822)")
            if typ != "OK" or not msg_datos or not msg_datos[0]:
                continue
            mensaje = email.message_from_bytes(msg_datos[0][1])
            asunto = buscador._decodificar_asunto(mensaje.get("Subject", ""))
            fecha_correo = buscador._fecha_del_correo(mensaje)
            cuerpo = ""
            adjuntos = []
            for parte in mensaje.walk():
                nombre_adjunto = parte.get_filename()
                if nombre_adjunto:
                    contenido = parte.get_payload(decode=True)
                    if contenido:
                        adjuntos.append((organizador.sanear_nombre(nombre_adjunto), contenido))
                elif parte.get_content_type() in ("text/html", "text/plain"):
                    texto_bruto = parte.get_payload(decode=True)
                    if texto_bruto:
                        texto = texto_bruto.decode(parte.get_content_charset() or "utf-8", errors="ignore")
                        cuerpo += _texto_plano_html(texto) + " "
            resultados.append({"asunto": asunto, "cuerpo": cuerpo, "adjuntos": adjuntos, "fecha": fecha_correo})
    return resultados


def _correo_es_seguro(contenido_normalizado, demandados):
    """Version del chequeo de seguridad de _archivo_es_seguro, pero sobre texto de correo (sin llamar a Drive)."""
    tiene_essa = any(buscador._nombre_coincide(contenido_normalizado, t) for t in buscador.TERMINOS_DEMANDANTE_VALIDO)
    if not tiene_essa:
        return False, "no se confirmo que el proceso sea de ESSA/Electrificadora de Santander"
    for demandado in demandados:
        resultado = buscador._demandado_coincide_en_texto(contenido_normalizado, demandado)
        if resultado is False:
            return False, f"el nombre en el correo no corresponde al demandado esperado ({demandado})"
    return True, "ok"


def _extraer_pdfs_de_zip(contenido: bytes, destino: Path) -> int:
    extraidos = 0
    try:
        with zipfile.ZipFile(io.BytesIO(contenido)) as archivo_zip:
            for info in archivo_zip.infolist():
                if info.is_dir() or Path(info.filename).suffix.lower() not in (".pdf", ".docx"):
                    continue
                nombre_seguro = organizador.sanear_nombre(Path(info.filename).name)
                ruta = buscador._ruta_archivo_libre(destino, nombre_seguro)
                with archivo_zip.open(info) as origen, open(organizador._ruta_larga_segura(str(ruta)), "wb") as f:
                    f.write(origen.read())
                extraidos += 1
    except zipfile.BadZipFile:
        pass
    return extraidos


def _guardar_correo(correo: dict, destino: Path) -> int:
    """
    Guarda los adjuntos PDF/DOCX del correo (extrayendo los de dentro
    de un .zip si aplica). Si no trae ningun adjunto util, guarda el
    asunto + cuerpo como un .txt simple, para no perder la informacion
    (ej. un correo de cobro en texto plano, sin adjuntos).
    """
    destino.mkdir(parents=True, exist_ok=True)
    guardados = 0
    for nombre_adjunto, contenido in correo["adjuntos"]:
        extension = Path(nombre_adjunto).suffix.lower()
        if extension == ".zip":
            guardados += _extraer_pdfs_de_zip(contenido, destino)
        elif extension in (".pdf", ".docx"):
            ruta = buscador._ruta_archivo_libre(destino, nombre_adjunto)
            with open(organizador._ruta_larga_segura(str(ruta)), "wb") as f:
                f.write(contenido)
            guardados += 1

    if guardados == 0:
        fecha_texto = correo["fecha"].isoformat() if correo["fecha"] else "sin_fecha"
        nombre_txt = organizador.sanear_nombre(f"{fecha_texto} - {correo['asunto'] or 'correo'}.txt")
        ruta = buscador._ruta_archivo_libre(destino, nombre_txt)
        with open(organizador._ruta_larga_segura(str(ruta)), "w", encoding="utf-8") as f:
            f.write(f"Asunto: {correo['asunto']}\nFecha: {correo['fecha']}\n\n{correo['cuerpo']}")
        guardados = 1
    return guardados


def _procesar_correo_no_procesal(proceso, destino: Path, nombre_carpeta: str, credenciales_correo) -> int:
    numero = proceso["numero"]
    usuario, app_password = credenciales_correo
    terminos = _terminos_busqueda(proceso["radicado"], proceso["cuentas"])
    if not terminos:
        return 0

    subidos = 0
    vistos = set()
    for termino in terminos:
        try:
            correos = buscar_correo_informacion_no_procesal(usuario, app_password, termino)
        except Exception as error:
            logging.error("[Correo] Proceso %s: fallo buscando '%s': %s", numero, termino, error)
            continue

        for correo in correos:
            clave = (correo["asunto"], correo["fecha"])
            if clave in vistos:
                continue

            contenido_norm = buscador._normalizar_para_comparar(
                correo["asunto"] + " " + correo["cuerpo"] + " " + " ".join(n for n, _ in correo["adjuntos"])
            )
            es_no_procesal, motivo = es_informacion_no_procesal(contenido_norm)
            if not es_no_procesal:
                continue

            seguro, motivo_seguridad = _correo_es_seguro(contenido_norm, proceso["demandados"])
            if not seguro:
                logging.info("   (se omite el correo '%s' del proceso %s: %s)", correo["asunto"], numero, motivo_seguridad)
                continue

            vistos.add(clave)
            if MODO_PRUEBA:
                logging.info(
                    "[SIMULACION] Proceso %s: subiria el correo '%s' (%s) a '%s'.",
                    numero, correo["asunto"], motivo, nombre_carpeta,
                )
                subidos += 1
                continue

            guardados = _guardar_correo(correo, destino)
            logging.info(
                "[Descargado] Proceso %s: correo '%s' (%s) -> %d archivo(s) en %s",
                numero, correo["asunto"], motivo, guardados, nombre_carpeta,
            )
            subidos += guardados
    return subidos


# ==================== Carpetas en disco ====================


def _listar_carpetas_existentes():
    carpeta_raiz = Path(CARPETA_PROCESOS)
    carpeta_raiz.mkdir(parents=True, exist_ok=True)
    return {
        d.name for d in carpeta_raiz.iterdir()
        if d.is_dir() and d.name not in cruce_excel.CARPETAS_A_IGNORAR
    }


def _descargar_archivo(servicio, archivo, destino: Path) -> Path:
    destino.mkdir(parents=True, exist_ok=True)
    nombre_seguro = organizador.sanear_nombre(archivo["name"])
    ruta_local = buscador._ruta_archivo_libre(destino, nombre_seguro)
    if archivo["mimeType"] in buscador.MIME_EXPORTAR:
        buscador._exportar_google_doc(servicio, archivo["id"], archivo["mimeType"], ruta_local)
    else:
        buscador._descargar_archivo_binario(servicio, archivo["id"], ruta_local)
    return ruta_local


# ==================== Procesamiento por fila ====================


def procesar_con_radicado(servicio, proceso, carpetas_existentes, credenciales_correo):
    """Filas que NO son terminadas (activo, suspendido, reorganizacion, remitida a castigo/prepago, etc)."""
    numero, radicado = proceso["numero"], proceso["radicado"]
    nombre_carpeta = proceso["nombre_carpeta"]
    destino = Path(CARPETA_PROCESOS) / nombre_carpeta

    if nombre_carpeta not in carpetas_existentes:
        if MODO_PRUEBA:
            logging.info("[SIMULACION] Proceso %s (%s, fila %s): se crearia la carpeta '%s'.", numero, proceso["estado"], proceso["fila_excel"], nombre_carpeta)
        else:
            destino.mkdir(parents=True, exist_ok=True)
            carpetas_existentes.add(nombre_carpeta)
            logging.info("[Creada] Proceso %s (%s, fila %s): carpeta '%s'.", numero, proceso["estado"], proceso["fila_excel"], nombre_carpeta)

    terminos = _terminos_busqueda(radicado, proceso["cuentas"])
    candidatos = _candidatos_en_drive(servicio, terminos)
    archivos = _archivos_de_candidatos(servicio, candidatos, radicado)

    subidos = 0
    for archivo, carpeta in archivos:
        seguro, motivo = _archivo_es_seguro(servicio, archivo, carpeta.get("name", ""), proceso["demandados"])
        if not seguro:
            logging.info("   (se omite '%s' del proceso %s: %s)", archivo["name"], numero, motivo)
            continue

        contenido_norm = _info_documento(servicio, archivo)
        es_no_procesal, motivo = es_informacion_no_procesal(contenido_norm)
        if not es_no_procesal:
            logging.info("   (se omite '%s' del proceso %s: %s)", archivo["name"], numero, motivo)
            continue

        if MODO_PRUEBA:
            logging.info("[SIMULACION] Proceso %s: subiria '%s' (%s) a '%s'.", numero, archivo["name"], motivo, nombre_carpeta)
        else:
            ruta = _descargar_archivo(servicio, archivo, destino)
            logging.info("[Descargado] Proceso %s: '%s' -> %s", numero, archivo["name"], ruta)
        subidos += 1

    if credenciales_correo:
        subidos += _procesar_correo_no_procesal(proceso, destino, nombre_carpeta, credenciales_correo)

    if subidos == 0:
        logging.info(
            "Proceso %s (%s, radicado %s): no se encontro informacion no procesal (tutelas/derechos de "
            "peticion/correos de cobro) para subir todavia.", numero, proceso["estado"], radicado,
        )


def procesar_terminado(servicio, proceso, carpetas_existentes, pendientes):
    numero, estado, radicado = proceso["numero"], proceso["estado"], proceso["radicado"]
    nombre_carpeta = proceso["nombre_carpeta"]
    destino = Path(CARPETA_PROCESOS) / nombre_carpeta

    if nombre_carpeta not in carpetas_existentes:
        if MODO_PRUEBA:
            logging.info("[SIMULACION] Proceso %s (%s, fila %s): se crearia la carpeta '%s'.", numero, estado, proceso["fila_excel"], nombre_carpeta)
        else:
            destino.mkdir(parents=True, exist_ok=True)
            carpetas_existentes.add(nombre_carpeta)
            logging.info("[Creada] Proceso %s (%s, fila %s): carpeta '%s'.", numero, estado, proceso["fila_excel"], nombre_carpeta)

    terminos = _terminos_busqueda(radicado, proceso["cuentas"])
    if not terminos:
        pendientes.append(proceso)
        logging.warning("Proceso %s (%s): no hay radicado ni cuenta valida para buscar en Drive, queda pendiente.", numero, estado)
        return

    candidatos = _candidatos_en_drive(servicio, terminos)
    archivos = _archivos_de_candidatos(servicio, candidatos, radicado)

    encontrado = False
    for archivo, carpeta in archivos:
        seguro, motivo = _archivo_es_seguro(servicio, archivo, carpeta.get("name", ""), proceso["demandados"])
        if not seguro:
            logging.info("   (se omite '%s' del proceso %s: %s)", archivo["name"], numero, motivo)
            continue

        contenido_norm = _info_documento(servicio, archivo)
        if not es_auto_terminador(contenido_norm):
            continue

        if MODO_PRUEBA:
            logging.info("[SIMULACION] Proceso %s (%s): subiria '%s' (termina el proceso) a '%s'.", numero, estado, archivo["name"], nombre_carpeta)
        else:
            ruta = _descargar_archivo(servicio, archivo, destino)
            logging.info("[Descargado] Proceso %s (%s): '%s' -> %s", numero, estado, archivo["name"], ruta)
        encontrado = True
        # No se corta el ciclo: puede haber mas de un documento relevante (ej. primera y segunda instancia).

    if not encontrado:
        pendientes.append(proceso)
        logging.warning("Proceso %s (%s): no se encontro el documento que termina el proceso, queda pendiente.", numero, estado)


# ==================== Orquestacion ====================


def procesar():
    if not RUTA_EXCEL_CONTROL or not os.path.exists(RUTA_EXCEL_CONTROL):
        logging.error("No se encontro el Excel configurado en RUTA_EXCEL_CONTROL: %r", RUTA_EXCEL_CONTROL)
        return

    procesos = leer_procesos_control()
    con_radicado, terminados, sin_estado = clasificar_procesos(procesos)
    logging.info(
        "Excel: %d fila(s) con numero de proceso -- %d activo/suspendido/reorganizacion/remitida/etc "
        "(carpeta con radicado, o con ESTADO PROCESAL si todavia no lo tiene), %d terminadas (pago/auto/"
        "contrato/no inicio), %d sin ESTADO PROCESAL diligenciado.",
        len(procesos), len(con_radicado), len(terminados), len(sin_estado),
    )

    servicio = autenticar_drive_o_none()
    if servicio is None:
        logging.error("No se puede continuar sin conexion a Google Drive.")
        return

    credenciales_correo = None
    if BUSCAR_EN_CORREO:
        credenciales_correo = organizador.leer_credenciales()
        if not credenciales_correo:
            logging.warning(
                "[Correo] No hay %s (o le faltan datos); se omite la busqueda en Gmail.",
                organizador.ARCHIVO_CREDENCIALES,
            )

    carpetas_existentes = _listar_carpetas_existentes()

    for proceso in con_radicado:
        try:
            procesar_con_radicado(servicio, proceso, carpetas_existentes, credenciales_correo)
        except Exception as error:
            logging.error("[Error] Proceso %s (fila %s) fallo y se omite -- se sigue con el resto: %s", proceso["numero"], proceso["fila_excel"], error)

    pendientes = []
    for proceso in terminados:
        try:
            procesar_terminado(servicio, proceso, carpetas_existentes, pendientes)
        except Exception as error:
            logging.error("[Error] Proceso %s (fila %s) fallo y se omite -- se sigue con el resto: %s", proceso["numero"], proceso["fila_excel"], error)

    if pendientes:
        with open(ARCHIVO_PENDIENTES_TERMINADOS, "w", newline="", encoding="utf-8-sig") as f:
            escritor = csv.writer(f, delimiter=";")
            escritor.writerow(["No.", "Fila Excel", "Estado", "Radicado", "Cuentas", "Demandados", "Juzgado"])
            for proceso in pendientes:
                escritor.writerow([
                    proceso["numero"], proceso["fila_excel"], proceso["estado"], proceso["radicado"] or "",
                    ", ".join(proceso["cuentas"]), ", ".join(proceso["demandados"]), proceso["juzgado"],
                ])
        logging.warning(
            "%d proceso(s) terminado(s) se quedaron SIN el documento que los termina -- descargalos a mano, "
            "quedaron listados en %s.", len(pendientes), ARCHIVO_PENDIENTES_TERMINADOS,
        )

    if sin_estado:
        logging.warning("%d fila(s) no tienen ESTADO PROCESAL diligenciado todavia en el Excel:", len(sin_estado))
        for proceso in sin_estado:
            logging.warning("   - Fila %s, proceso %s", proceso["fila_excel"], proceso["numero"])

    if MODO_PRUEBA:
        logging.info(
            "MODO_PRUEBA esta activo: no se creo ninguna carpeta ni se descargo nada de verdad todavia. "
            "Revisa el log y, si se ve bien, cambia MODO_PRUEBA = False al inicio de este script y vuelve a correrlo."
        )


def main():
    configurar_logging()
    procesar()


if __name__ == "__main__":
    main()
