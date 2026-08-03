"""
Especializa la búsqueda en Google Drive para el informe
"3._CONTROL_PROCESOS_EJECUTIVOS_ESSA...xlsm" (hoja `HOJA_EXCEL_CONTROL`,
por defecto `DatosProcesados1` -- la versión de la hoja `ACTIVOS` ya
aplanada a un solo encabezado por fila).

A diferencia de `buscar_faltantes_en_drive.py` (que descarga TODO el
contenido relacionado con el radicado), este script aplica dos reglas
de negocio distintas según el `ESTADO PROCESAL` de cada proceso:

1. Procesos en un estado de `ESTADOS_INFORMACION_NO_PROCESAL` (`ACTIVO`,
   `ACTIVOS CON TITULOS`, `SUSPENDIDO`, `REORGANIZACION` -- los mismos
   de `validar_renombrar_carpetas.ESTADOS_A_CONTAR`): la carpeta se
   nombra `"<numero>. <radicado>"` (igual que el resto del proyecto), y
   SOLO se sube información NO procesal -- derechos de petición,
   tutelas y solicitudes dirigidas a una entidad DISTINTA al juzgado
   del proceso (bancos, EPS, ministerios, municipios, etc). Los
   memoriales, recursos, contestaciones y demás actuaciones dirigidas
   al juzgado del proceso NO se suben. Ver `es_informacion_no_procesal`.

2. Procesos terminados por pago, por auto, por contrato/prepago, o que
   nunca se presentaron (`ESTADO PROCESAL` que empieza con `TERMINADO`
   o `NO INICIO` -- igual que `crear_carpetas_terminados_castigo.py`):
   la carpeta se nombra `"<numero>. <ESTADO PROCESAL EXACTO del Excel>"`
   (ej. `"245. TERMINADO POR AUTO"`), y SOLO se sube el AUTO que
   termina el proceso (el que decreta la terminación por pago, acepta
   el retiro de la demanda, o decreta la terminación). Ver
   `es_auto_terminador`. Si no se encuentra ese auto en Drive, el
   proceso queda listado en `ARCHIVO_PENDIENTES_TERMINADOS` para que lo
   descargues a mano.

Los procesos en cualquier OTRO estado (`REMITIDA*`, `DESISTIMIENTO DE
PRETENSIONES`, etc) quedan FUERA del alcance de este script a
propósito -- no se tocan, y se reportan en el log (usa
`crear_carpetas_terminados_castigo.py` para esos si hace falta).

La clasificación de "información no procesal" y de "auto que termina
el proceso" es por PALABRAS CLAVE (nombre del archivo y, si es PDF/DOCX,
sus primeras páginas de contenido) -- ES una heurística, no perfecta.
Cada decisión (por qué se subió o por qué se omitió cada archivo) queda
registrada en el log para que la revises y ajustes las listas de
palabras clave (`PALABRAS_TIPO_INFORMACION_FUERTES`,
`PALABRAS_TIPO_INFORMACION_SOLO_NOMBRE`, `PALABRAS_PROCESAL_JUZGADO`,
`PALABRAS_AUTO_TERMINADOR`) si hace falta.

Reutiliza toda la infraestructura de `buscar_faltantes_en_drive.py`
(autenticación, búsqueda por radicado/radicado corto/cuenta, validación
de que el documento sea de ESSA y del demandado correcto, descarga de
PDF/exportables de Google) -- no repite esa lógica.

Requiere las mismas credenciales que `buscar_faltantes_en_drive.py`
(`credenciales_drive.json` / `token_drive.json`, ver README).

Respeta MODO_PRUEBA (por defecto True): en modo prueba solo BUSCA y
CLASIFICA, mostrando qué subiría y a qué carpeta, sin crear carpetas ni
descargar nada todavía.
"""

import csv
import logging
import os
import re
from pathlib import Path

import openpyxl

import buscar_faltantes_en_drive as buscador
import crear_carpetas_terminados_castigo as terminados_folder
import procesos_juridicos as organizador
import validar_renombrar_carpetas as cruce_excel

# ============================= CONFIGURACION =============================

# Ruta al informe de Excel de procesos ejecutivos (.xlsm). Cambia esto a
# donde tengas guardado "3._CONTROL_PROCESOS_EJECUTIVOS_ESSA...xlsm".
RUTA_EXCEL_CONTROL = r"C:\Users\User\Documents\3._CONTROL_PROCESOS_EJECUTIVOS_ESSA.xlsm"

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

# Mismos estados que validar_renombrar_carpetas.ESTADOS_A_CONTAR -- se
# organizan con la carpeta "<numero>. <radicado>" de siempre, pero aquí
# SOLO se les sube información no procesal (ver módulo docstring).
ESTADOS_INFORMACION_NO_PROCESAL = {estado.upper() for estado in cruce_excel.ESTADOS_A_CONTAR}

# Prefijos de ESTADO PROCESAL que cuentan como "terminado" para este
# script -- igual que crear_carpetas_terminados_castigo.py: TERMINADO
# (por pago, por auto, por contrato/prepago) y NO INICIO (nunca se
# presentó la demanda, quedó como reclamación administrativa).
PREFIJOS_ESTADO_TERMINADO = ("TERMINADO", "NO INICIO")

CARPETA_PROCESOS = cruce_excel.CARPETA_PROCESOS

ARCHIVO_LOG = os.path.join(os.path.dirname(__file__), "clasificar_procesos_ejecutivos.log")

# Procesos terminados (por pago/auto/contrato/no inicio) en los que NO
# se encontró el auto que los termina -- para que los descargues a mano.
ARCHIVO_PENDIENTES_TERMINADOS = os.path.join(os.path.dirname(__file__), "terminados_sin_auto_pendientes.csv")

# True (por defecto): no crea carpetas ni descarga nada, solo busca y
# muestra qué haría. False: aplica los cambios de verdad.
MODO_PRUEBA = True

# --------------------- Palabras clave de clasificación ---------------------

# Frases que, si aparecen en el CONTENIDO (o el nombre) de un documento,
# confirman por sí solas que es información no procesal (derecho de
# petición o tutela) -- son frases poco ambiguas, casi no aparecen
# "de pasada" dentro de un memorial normal dirigido al juzgado.
PALABRAS_TIPO_INFORMACION_FUERTES = [
    "DERECHO DE PETICION",
    "ACCION DE TUTELA",
    "TUTELA",
]

# "SOLICITUD" es demasiado común DENTRO del contenido de cualquier
# memorial ("se solicita...") como para usarla ahí -- solo cuenta si
# aparece en el NOMBRE del archivo (donde sí suele describir el TIPO de
# documento, ej. "SOLICITUD INFORMACION BANCOLOMBIA.pdf").
PALABRAS_TIPO_INFORMACION_SOLO_NOMBRE = [
    "SOLICITUD",
]

# Si el documento menciona alguna de estas palabras (marcas típicas de
# una actuación procesal) Y además menciona al JUZGADO del proceso, se
# descarta como información no procesal -- lo más probable es que sea
# un memorial/actuación dirigida al juzgado, no una solicitud a un
# tercero. Si el documento no menciona al juzgado (ej. una tutela o
# derecho de petición real dirigida a otra entidad), esta lista no
# bloquea nada.
PALABRAS_PROCESAL_JUZGADO = [
    "MEMORIAL", "DEMANDA", "CONTESTACION", "EXCEPCIONES", "RECURSO",
    "REPOSICION", "APELACION", "ALEGATOS", "TRASLADO", "MANDAMIENTO",
    "AUTO", "SENTENCIA", "NOTIFICACION", "EMBARGO", "SECUESTRO",
    "REMATE", "LIQUIDACION", "PODER", "SUSTITUCION", "CURADOR",
    "EMPLAZAMIENTO", "DILIGENCIA",
]

# Frases que confirman que un AUTO es el que TERMINA el proceso (por
# pago, por retiro/desistimiento de la demanda, o que decreta la
# terminación en general). Debe aparecer ADEMÁS la palabra "AUTO" (ver
# es_auto_terminador) -- sin eso, un memorial que simplemente PIDE la
# terminación no cuenta, solo el AUTO que la decreta.
PALABRAS_AUTO_TERMINADOR = [
    "TERMINA EL PROCESO", "TERMINACION DEL PROCESO", "TERMINACION POR PAGO",
    "TERMINACION POR CONTRATO", "TERMINACION DE LA OBLIGACION",
    "DECRETA LA TERMINACION", "DECRETA TERMINACION",
    "ACEPTA EL RETIRO", "ACEPTA RETIRO DE LA DEMANDA", "RETIRO DE LA DEMANDA",
    "TERMINACION POR DESISTIMIENTO", "APRUEBA EL DESISTIMIENTO",
    "ARCHIVA EL PROCESO", "ARCHIVESE EL PROCESO",
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
    Lee HOJA_EXCEL_CONTROL y agrupa las filas por columna No. (un
    proceso puede tener varias cuentas/demandados en varias filas).
    Devuelve {numero: {"estados": {...}, "radicados": {...},
    "cuentas": {...}, "demandados": {...}, "juzgados": {...}}}.
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

    por_numero = {}
    for fila in ws.iter_rows(min_row=FILA_ENCABEZADO_CONTROL + 1, values_only=True):
        numero_crudo = fila[idx_no]
        if numero_crudo is None or not isinstance(numero_crudo, (int, float)):
            continue  # fila vacia o de notas/leyenda al final de la hoja
        numero = int(numero_crudo)

        estado = str(fila[idx_estado]).strip() if fila[idx_estado] is not None else ""
        cuenta = str(fila[idx_cuenta]).strip() if fila[idx_cuenta] is not None else ""
        demandado = str(fila[idx_demandado]).strip() if fila[idx_demandado] is not None else ""
        juzgado = str(fila[idx_juzgado]).strip() if fila[idx_juzgado] is not None else ""
        radicado = _normalizar_radicado(fila[idx_radicado])

        registro = por_numero.setdefault(numero, {
            "estados": set(), "radicados": set(), "cuentas": set(),
            "demandados": set(), "juzgados": set(),
        })
        if estado:
            registro["estados"].add(estado)
        if radicado:
            registro["radicados"].add(radicado)
        if cuenta and cuenta.strip("0"):
            registro["cuentas"].add(cuenta)
        if demandado:
            registro["demandados"].add(demandado)
        if juzgado:
            registro["juzgados"].add(juzgado)
    return por_numero


def clasificar_procesos(por_numero):
    """
    Separa los procesos leídos en dos listas de trabajo (activos,
    terminados) y tres listas informativas (fuera_de_alcance, sin_estado,
    ambiguos). Ver módulo docstring para las reglas de clasificación.
    """
    activos = []
    terminados = []
    fuera_de_alcance = []
    sin_estado = []
    ambiguos = []

    for numero, datos in sorted(por_numero.items()):
        if len(datos["estados"]) > 1:
            ambiguos.append((numero, f"varios ESTADO PROCESAL distintos: {sorted(datos['estados'])}"))
            continue
        if len(datos["radicados"]) > 1:
            ambiguos.append((numero, f"varios RADICADO distintos: {sorted(datos['radicados'])}"))
            continue
        if not datos["estados"]:
            sin_estado.append(numero)
            continue

        estado = next(iter(datos["estados"]))
        estado_norm = estado.upper()
        radicado = next(iter(datos["radicados"])) if datos["radicados"] else None
        proceso = {
            "numero": numero,
            "estado": estado,
            "radicado": radicado,
            "cuentas": sorted(datos["cuentas"]),
            "demandados": sorted(datos["demandados"]),
            "juzgado": next(iter(sorted(datos["juzgados"])), ""),
        }

        if estado_norm in ESTADOS_INFORMACION_NO_PROCESAL:
            if not radicado:
                fuera_de_alcance.append((numero, estado, "activo/suspendido/reorganizacion sin radicado de 23 digitos todavia"))
                continue
            activos.append(proceso)
        elif estado_norm.startswith(PREFIJOS_ESTADO_TERMINADO):
            terminados.append(proceso)
        else:
            fuera_de_alcance.append((numero, estado, "fuera del alcance de este script"))

    return activos, terminados, fuera_de_alcance, sin_estado, ambiguos


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
    """(nombre_normalizado, contenido_normalizado) -- ver funciones de clasificacion mas abajo."""
    nombre = archivo.get("name", "")
    texto = ""
    if Path(nombre).suffix.lower() in buscador.EXTENSIONES_CONTENIDO_DRIVE:
        texto = buscador._texto_de_archivo_drive(servicio, archivo)
    nombre_norm = buscador._normalizar_para_comparar(nombre)
    contenido_norm = buscador._normalizar_para_comparar(nombre + " " + texto)
    return nombre_norm, contenido_norm


def _menciona_palabras(contenido_normalizado, texto_esperado):
    palabras_esperadas = buscador._palabras_significativas(texto_esperado)
    if not palabras_esperadas:
        return False
    palabras_en_contenido = set(re.findall(r"[A-ZÑ]+", contenido_normalizado))
    return bool(palabras_esperadas & palabras_en_contenido)


def es_informacion_no_procesal(nombre_normalizado, contenido_normalizado, juzgado):
    """
    True si el documento parece ser un derecho de petición, tutela o
    solicitud dirigida a una entidad DISTINTA al juzgado del proceso.
    Ver las listas PALABRAS_TIPO_INFORMACION_* y PALABRAS_PROCESAL_JUZGADO.
    """
    tiene_tipo = (
        any(frase in contenido_normalizado for frase in PALABRAS_TIPO_INFORMACION_FUERTES)
        or any(frase in nombre_normalizado for frase in PALABRAS_TIPO_INFORMACION_SOLO_NOMBRE)
    )
    if not tiene_tipo:
        return False, "no menciona derecho de peticion, tutela ni solicitud"

    tiene_marcador_procesal = any(frase in contenido_normalizado for frase in PALABRAS_PROCESAL_JUZGADO)
    if tiene_marcador_procesal and juzgado and _menciona_palabras(contenido_normalizado, juzgado):
        return False, "parece un documento procesal dirigido al juzgado del proceso, no a otra entidad"

    return True, "ok"


def es_auto_terminador(nombre_normalizado, contenido_normalizado):
    """True si el documento parece ser el AUTO que termina el proceso (por pago, retiro de demanda, o terminacion en general)."""
    contenido_completo = nombre_normalizado + " " + contenido_normalizado
    tiene_auto = re.search(r"\bAUTO\b", contenido_completo) is not None
    tiene_frase_terminacion = any(frase in contenido_completo for frase in PALABRAS_AUTO_TERMINADOR)
    return tiene_auto and tiene_frase_terminacion


# ==================== Carpetas en disco ====================


def _listar_carpetas_existentes():
    carpeta_raiz = Path(CARPETA_PROCESOS)
    carpeta_raiz.mkdir(parents=True, exist_ok=True)
    return {
        d.name for d in carpeta_raiz.iterdir()
        if d.is_dir() and d.name not in cruce_excel.CARPETAS_A_IGNORAR
    }


def _carpeta_existente_para_numero(numero, carpetas_existentes):
    prefijo = f"{numero}. "
    exacto = f"{numero}."
    for nombre in carpetas_existentes:
        if nombre.startswith(prefijo) or nombre == exacto:
            return nombre
    return None


def _descargar_archivo(servicio, archivo, destino: Path) -> Path:
    destino.mkdir(parents=True, exist_ok=True)
    nombre_seguro = organizador.sanear_nombre(archivo["name"])
    ruta_local = buscador._ruta_archivo_libre(destino, nombre_seguro)
    if archivo["mimeType"] in buscador.MIME_EXPORTAR:
        buscador._exportar_google_doc(servicio, archivo["id"], archivo["mimeType"], ruta_local)
    else:
        buscador._descargar_archivo_binario(servicio, archivo["id"], ruta_local)
    return ruta_local


# ==================== Procesamiento por proceso ====================


def procesar_activo(servicio, proceso, carpetas_existentes):
    numero, radicado, juzgado = proceso["numero"], proceso["radicado"], proceso["juzgado"]
    existente = _carpeta_existente_para_numero(numero, carpetas_existentes)
    nombre_carpeta = existente or f"{numero}. {radicado}"
    destino = Path(CARPETA_PROCESOS) / nombre_carpeta

    if not existente:
        if MODO_PRUEBA:
            logging.info("[SIMULACION] Proceso %s (%s): se crearia la carpeta '%s'.", numero, proceso["estado"], nombre_carpeta)
        else:
            destino.mkdir(parents=True, exist_ok=True)
            carpetas_existentes.add(nombre_carpeta)
            logging.info("[Creada] Proceso %s (%s): carpeta '%s'.", numero, proceso["estado"], nombre_carpeta)

    terminos = _terminos_busqueda(radicado, proceso["cuentas"])
    candidatos = _candidatos_en_drive(servicio, terminos)
    archivos = _archivos_de_candidatos(servicio, candidatos, radicado)

    subidos = 0
    for archivo, carpeta in archivos:
        seguro, motivo = _archivo_es_seguro(servicio, archivo, carpeta.get("name", ""), proceso["demandados"])
        if not seguro:
            logging.info("   (se omite '%s' del proceso %s: %s)", archivo["name"], numero, motivo)
            continue

        nombre_norm, contenido_norm = _info_documento(servicio, archivo)
        es_no_procesal, motivo = es_informacion_no_procesal(nombre_norm, contenido_norm, juzgado)
        if not es_no_procesal:
            logging.info("   (se omite '%s' del proceso %s: %s)", archivo["name"], numero, motivo)
            continue

        if MODO_PRUEBA:
            logging.info("[SIMULACION] Proceso %s: subiria '%s' (informacion no procesal) a '%s'.", numero, archivo["name"], nombre_carpeta)
        else:
            ruta = _descargar_archivo(servicio, archivo, destino)
            logging.info("[Descargado] Proceso %s: '%s' -> %s", numero, archivo["name"], ruta)
        subidos += 1

    if subidos == 0:
        logging.info(
            "Proceso %s (%s, radicado %s): no se encontro informacion no procesal (peticiones/tutelas/"
            "solicitudes) para subir todavia.", numero, proceso["estado"], radicado,
        )


def procesar_terminado(servicio, proceso, carpetas_existentes, pendientes):
    numero, estado, radicado = proceso["numero"], proceso["estado"], proceso["radicado"]
    existente = _carpeta_existente_para_numero(numero, carpetas_existentes)
    nombre_carpeta = existente or terminados_folder._nombre_carpeta_para(numero, estado) or f"{numero}. {estado}"
    destino = Path(CARPETA_PROCESOS) / nombre_carpeta

    if not existente:
        if MODO_PRUEBA:
            logging.info("[SIMULACION] Proceso %s (%s): se crearia la carpeta '%s'.", numero, estado, nombre_carpeta)
        else:
            destino.mkdir(parents=True, exist_ok=True)
            carpetas_existentes.add(nombre_carpeta)
            logging.info("[Creada] Proceso %s (%s): carpeta '%s'.", numero, estado, nombre_carpeta)

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

        nombre_norm, contenido_norm = _info_documento(servicio, archivo)
        if not es_auto_terminador(nombre_norm, contenido_norm):
            continue

        if MODO_PRUEBA:
            logging.info("[SIMULACION] Proceso %s (%s): subiria el auto '%s' a '%s'.", numero, estado, archivo["name"], nombre_carpeta)
        else:
            ruta = _descargar_archivo(servicio, archivo, destino)
            logging.info("[Descargado] Proceso %s (%s): auto '%s' -> %s", numero, estado, archivo["name"], ruta)
        encontrado = True
        # No se corta el ciclo: puede haber mas de un auto relevante (ej. primera y segunda instancia).

    if not encontrado:
        pendientes.append(proceso)
        logging.warning("Proceso %s (%s): no se encontro el auto que termina el proceso, queda pendiente.", numero, estado)


# ==================== Orquestacion ====================


def procesar():
    if not RUTA_EXCEL_CONTROL or not os.path.exists(RUTA_EXCEL_CONTROL):
        logging.error("No se encontro el Excel configurado en RUTA_EXCEL_CONTROL: %r", RUTA_EXCEL_CONTROL)
        return

    por_numero = leer_procesos_control()
    activos, terminados, fuera_de_alcance, sin_estado, ambiguos = clasificar_procesos(por_numero)
    logging.info(
        "Excel: %d proceso(s) distintos -- %d activos/suspendidos/reorganizacion, %d terminados "
        "(pago/auto/contrato/no inicio), %d fuera de alcance, %d sin estado, %d con datos ambiguos.",
        len(por_numero), len(activos), len(terminados), len(fuera_de_alcance), len(sin_estado), len(ambiguos),
    )

    servicio = autenticar_drive_o_none()
    if servicio is None:
        logging.error("No se puede continuar sin conexion a Google Drive.")
        return

    carpetas_existentes = _listar_carpetas_existentes()

    for proceso in activos:
        try:
            procesar_activo(servicio, proceso, carpetas_existentes)
        except Exception as error:
            logging.error("[Error] Proceso %s (activo) fallo y se omite -- se sigue con el resto: %s", proceso["numero"], error)

    pendientes = []
    for proceso in terminados:
        try:
            procesar_terminado(servicio, proceso, carpetas_existentes, pendientes)
        except Exception as error:
            logging.error("[Error] Proceso %s (terminado) fallo y se omite -- se sigue con el resto: %s", proceso["numero"], error)

    if pendientes:
        with open(ARCHIVO_PENDIENTES_TERMINADOS, "w", newline="", encoding="utf-8-sig") as f:
            escritor = csv.writer(f, delimiter=";")
            escritor.writerow(["No.", "Estado", "Radicado", "Cuentas", "Demandados", "Juzgado"])
            for proceso in pendientes:
                escritor.writerow([
                    proceso["numero"], proceso["estado"], proceso["radicado"] or "",
                    ", ".join(proceso["cuentas"]), ", ".join(proceso["demandados"]), proceso["juzgado"],
                ])
        logging.warning(
            "%d proceso(s) terminado(s) se quedaron SIN el auto que los termina -- descargalos a mano, "
            "quedaron listados en %s.", len(pendientes), ARCHIVO_PENDIENTES_TERMINADOS,
        )

    if fuera_de_alcance:
        logging.info(
            "%d proceso(s) con un ESTADO PROCESAL fuera del alcance de este script (ni activo/suspendido/"
            "reorganizacion, ni terminado por pago/auto/contrato/no inicio):", len(fuera_de_alcance),
        )
        for numero, estado, motivo in fuera_de_alcance:
            logging.info("   - Proceso %s (%r): %s", numero, estado, motivo)

    if sin_estado:
        logging.warning("%d proceso(s) no tienen ESTADO PROCESAL diligenciado todavia en el Excel: %s", len(sin_estado), sin_estado)

    if ambiguos:
        logging.warning(
            "%d proceso(s) tienen datos ambiguos en el Excel (mas de un ESTADO PROCESAL o RADICADO distinto "
            "para el mismo No.) -- se omitieron, revisalos a mano:", len(ambiguos),
        )
        for numero, motivo in ambiguos:
            logging.warning("   - Proceso %s: %s", numero, motivo)

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
