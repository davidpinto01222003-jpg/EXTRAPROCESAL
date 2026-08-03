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
   CASTIGO`/`A PREPAGO`, etc -- se organiza con la carpeta de siempre
   `"<numero>. <radicado>"`, y SOLO se sube información NO procesal:
   derechos de petición, tutelas y solicitudes (ver
   `es_informacion_no_procesal`). Si la fila no tiene un radicado válido
   de 23 dígitos todavía, no se puede crear su carpeta y queda
   reportada aparte (ver `sin_radicado` en el log).

2. Procesos terminados por pago, por auto, por contrato/prepago, o que
   nunca se presentaron (`ESTADO PROCESAL` que empieza con `TERMINADO`
   o `NO INICIO` -- igual que `crear_carpetas_terminados_castigo.py`):
   la carpeta se nombra `"<numero>. <ESTADO PROCESAL EXACTO del Excel>"`
   (ej. `"245. TERMINADO POR AUTO"`), y SOLO se sube el documento que
   deja constancia de que el proceso NO sigue su curso -- un auto de
   terminación, de aceptación de retiro de la demanda, de
   desistimiento, etc (ver `es_auto_terminador`, ya NO exige que
   aparezca literalmente la palabra "AUTO"). Si no se encuentra ese
   documento en Drive, el proceso queda listado en
   `ARCHIVO_PENDIENTES_TERMINADOS` para que lo descargues a mano.

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
termina el proceso" es por PALABRAS CLAVE (nombre del archivo y, si es
PDF/DOCX, sus primeras páginas de contenido) -- es una heurística, no
perfecta. Cada decisión queda registrada en el log para que la revises
y ajustes las listas de palabras clave
(`PALABRAS_TIPO_INFORMACION_FUERTES`,
`PALABRAS_TIPO_INFORMACION_SOLO_NOMBRE`, `PALABRAS_PROCESO_NO_CONTINUA`)
si hace falta.

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

# --------------------- Palabras clave de clasificación ---------------------

# Frases que, si aparecen en el CONTENIDO (o el nombre) de un documento,
# confirman que es un derecho de petición o una tutela.
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
    """Separa las filas leidas en dos listas de trabajo (con_radicado, terminados) y dos informativas (sin_estado, sin_radicado)."""
    con_radicado = []
    terminados = []
    sin_estado = []
    sin_radicado = []

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
            if not proceso["radicado"]:
                sin_radicado.append(proceso)
                continue
            proceso["nombre_base"] = f"{numero}. {proceso['radicado']}"
            con_radicado.append(proceso)

    _asignar_nombres_de_carpeta(con_radicado)
    _asignar_nombres_de_carpeta(terminados)
    return con_radicado, terminados, sin_estado, sin_radicado


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


def es_informacion_no_procesal(nombre_normalizado, contenido_normalizado):
    """
    True si el documento parece ser un derecho de petición, tutela o
    solicitud -- sin importar si menciona o no al juzgado del proceso
    (lo único que importa es que SEA una petición/tutela/solicitud).
    """
    tiene_tipo = (
        any(frase in contenido_normalizado for frase in PALABRAS_TIPO_INFORMACION_FUERTES)
        or any(frase in nombre_normalizado for frase in PALABRAS_TIPO_INFORMACION_SOLO_NOMBRE)
    )
    if not tiene_tipo:
        return False, "no menciona derecho de peticion, tutela ni solicitud"
    return True, "ok"


def es_auto_terminador(nombre_normalizado, contenido_normalizado):
    """
    True si el documento da a entender que el proceso judicial NO sigue
    su curso (terminación por pago, retiro/desistimiento de la demanda,
    archivo del proceso, etc) -- NO exige que aparezca la palabra
    "AUTO", con que aparezca cualquiera de PALABRAS_PROCESO_NO_CONTINUA
    alcanza.
    """
    contenido_completo = nombre_normalizado + " " + contenido_normalizado
    return any(frase in contenido_completo for frase in PALABRAS_PROCESO_NO_CONTINUA)


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


def procesar_con_radicado(servicio, proceso, carpetas_existentes):
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

        nombre_norm, contenido_norm = _info_documento(servicio, archivo)
        es_no_procesal, motivo = es_informacion_no_procesal(nombre_norm, contenido_norm)
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

        nombre_norm, contenido_norm = _info_documento(servicio, archivo)
        if not es_auto_terminador(nombre_norm, contenido_norm):
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
    con_radicado, terminados, sin_estado, sin_radicado = clasificar_procesos(procesos)
    logging.info(
        "Excel: %d fila(s) con numero de proceso -- %d con radicado (activo/suspendido/reorganizacion/"
        "remitida/etc), %d terminadas (pago/auto/contrato/no inicio), %d sin estado, %d sin radicado valido.",
        len(procesos), len(con_radicado), len(terminados), len(sin_estado), len(sin_radicado),
    )

    servicio = autenticar_drive_o_none()
    if servicio is None:
        logging.error("No se puede continuar sin conexion a Google Drive.")
        return

    carpetas_existentes = _listar_carpetas_existentes()

    for proceso in con_radicado:
        try:
            procesar_con_radicado(servicio, proceso, carpetas_existentes)
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

    if sin_radicado:
        logging.warning(
            "%d fila(s) no tienen un radicado de 23 digitos valido todavia y no son de un estado 'terminado' "
            "-- no se les pudo crear carpeta, revisalas a mano:", len(sin_radicado),
        )
        for proceso in sin_radicado:
            logging.warning("   - Fila %s, proceso %s (%s)", proceso["fila_excel"], proceso["numero"], proceso["estado"])

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
