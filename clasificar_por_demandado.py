"""
Clasifica información extraprocesal por el NOMBRE DEL DEMANDADO, en dos
pasos independientes (ambos respetan MODO_PRUEBA):

1. CARPETA DE DESCARGAS MANUALES (CARPETA_DESCARGAS_MANUAL, por
   defecto tu carpeta "Downloads"): revisa cada PDF/DOCX que haya ahí
   (nombre de archivo y, si hace falta, su contenido) y lo MUEVE
   directo a la carpeta del proceso cuyo DEMANDADO coincide -- ej. si
   el demandado es "ALBERTO SUAREZ" y el archivo menciona "ALBERTO
   SUAREZ", se mueve a "<numero>. <radicado o ESTADO>" en
   CARPETA_PROCESOS. No hay ninguna restricción de TIPO de documento
   aquí (a diferencia de clasificar_procesos_ejecutivos.py) -- se
   asume que ya es información extraprocesal porque tú la descargaste
   a propósito; el único trabajo de este script es encontrar a qué
   proceso corresponde.

   La coincidencia por demandado es ESTRICTA a propósito: tienen que
   aparecer TODAS las palabras significativas del nombre del demandado
   en el archivo (ver _demandado_coincide_fuerte), no basta con que
   coincida una sola -- a diferencia del resto del proyecto (donde el
   radicado/cuenta ya corrobora el proceso y basta con que coincida
   CUALQUIER dato), aquí el nombre del demandado es la ÚNICA señal
   disponible, así que tiene que ser una coincidencia sólida antes de
   mover un documento de verdad. Si el archivo no coincide con ningún
   demandado, o coincide con más de uno (dos procesos activos contra
   la misma persona), se deja donde está y se reporta para que lo
   revises a mano.

2. GMAIL: busca cada DEMANDADO de los procesos activos en TODO tu
   Gmail (no solo la bandeja de entrada) -- al revés de
   clasificar_procesos_ejecutivos.py (que busca por TIPO de documento y
   despues empareja por demandado/radicado/cuenta), aquí se busca
   DIRECTAMENTE por el nombre de cada demandado, y del resultado SOLO
   se descarga lo que además sea un derecho de petición o una tutela
   (NO pagos oficiosos esta vez -- ver PALABRAS_TIPO_A_DESCARGAR) --
   igual de riguroso que el resto del proyecto: si el nombre del
   archivo/asunto ya trae una marca de documento procesal (demanda,
   memorial, etc), se descarta aunque mencione la tutela/petición de
   pasada. También se exige, igual que en el resto del proyecto, que
   el correo mencione a ESSA/Electrificadora de Santander.

   En vez de abrir una conexión de Gmail POR CADA demandado (con
   cientos de procesos activos eso tardaría horas), se abre UNA sola
   conexión y se buscan los demandados en LOTES combinados con OR
   (ver TAMANO_LOTE_CORREO) -- Gmail permite eso mismo en su propia
   barra de búsqueda. Por cada correo encontrado en un lote, se revisa
   cuál(es) demandado(s) del lote coinciden de verdad (misma
   coincidencia ESTRICTA del paso 1).

Reutiliza toda la lectura del Excel y las carpetas de
clasificar_procesos_ejecutivos.py (misma hoja ACTIVOS, mismo
CARPETA_PROCESOS) -- solo considera procesos ACTIVOS (todo lo que no
es terminado/no inicio), igual que la regla de "información no
procesal" del resto del proyecto. Requiere las mismas credenciales
credenciales_sgde.txt para Gmail (ver README) -- si no existe, el
paso 2 se omite solo, sin error.
"""

import imaplib
import logging
import os
from pathlib import Path

import clasificar_procesos_ejecutivos as base
import buscar_faltantes_en_drive as buscador
import validar_renombrar_carpetas as cruce_excel
import procesos_juridicos as organizador

# ============================= CONFIGURACION =============================

# Carpeta donde TU descargas a mano la información extraprocesal antes
# de que este script la reparta -- por defecto tu carpeta de Descargas
# de Windows. Cambiala si usas otra carpeta.
CARPETA_DESCARGAS_MANUAL = os.path.join(os.path.expanduser("~"), "Downloads")

# True (por defecto): además de la carpeta de descargas, busca en TODO
# tu Gmail. Si no existe credenciales_sgde.txt, este paso se omite
# solo, sin error.
BUSCAR_EN_CORREO = True

# Cuántos demandados se combinan en UNA sola búsqueda de Gmail (con
# OR) -- evita abrir una conexión/búsqueda separada por cada uno de
# los cientos de demandados activos, que sería muy lento.
TAMANO_LOTE_CORREO = 15

# Del correo encontrado por demandado, SOLO se descarga si además es
# un derecho de petición o una tutela -- a propósito NO incluye pago
# oficioso esta vez (a diferencia de clasificar_procesos_ejecutivos.py),
# el usuario pidió específicamente "solo si se trata de derecho de
# petición o tutela".
PALABRAS_TIPO_A_DESCARGAR = base.PALABRAS_TIPO_INFORMACION_FUERTES

# True (por defecto): no mueve archivos ni descarga correos de verdad,
# solo revisa y muestra qué haría.
MODO_PRUEBA = True

ARCHIVO_LOG = os.path.join(os.path.dirname(__file__), "clasificar_por_demandado.log")

# ===========================================================================


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


def _indexar_demandados(con_radicado):
    """{demandado_exacto: [proceso, ...]} -- el mismo nombre de demandado puede tener mas de un proceso activo."""
    indice = {}
    for proceso in con_radicado:
        for demandado in proceso["demandados"]:
            demandado = demandado.strip()
            if demandado:
                indice.setdefault(demandado, []).append(proceso)
    return indice


def _demandado_coincide_fuerte(texto_normalizado, demandado):
    """
    Coincidencia FUERTE: TODAS las palabras significativas del
    demandado (ver buscador._palabras_significativas) tienen que
    aparecer en el texto -- no basta con que coincida una sola. A
    diferencia del resto del proyecto (donde el radicado/cuenta ya
    corrobora el proceso y basta con que coincida CUALQUIER dato),
    aqui el nombre del demandado es la UNICA señal disponible.
    """
    palabras = buscador._palabras_significativas(demandado)
    if not palabras:
        return False
    return all(buscador._nombre_coincide(texto_normalizado, palabra) for palabra in palabras)


def _procesos_que_coinciden_fuerte(texto_normalizado, indice_demandados):
    encontrados = {}
    for demandado, procesos in indice_demandados.items():
        if _demandado_coincide_fuerte(texto_normalizado, demandado):
            for proceso in procesos:
                encontrados[proceso["nombre_carpeta"]] = proceso
    return list(encontrados.values())


def _es_tutela_o_peticion(nombre_normalizado, contenido_normalizado):
    """
    Igual de estricto que clasificar_procesos_ejecutivos.es_informacion_no_procesal,
    pero SOLO tutela/derecho de peticion (ver PALABRAS_TIPO_A_DESCARGAR
    -- sin pago oficioso).
    """
    if any(frase in nombre_normalizado for frase in PALABRAS_TIPO_A_DESCARGAR):
        return True, "tutela/derecho de peticion (nombre del archivo/asunto)"
    if any(marca in nombre_normalizado for marca in base.PALABRAS_PROCESAL_EXCLUIR):
        return False, "el nombre indica que es un documento procesal (demanda/memorial/solicitud/etc)"
    if any(frase in contenido_normalizado[:400] for frase in PALABRAS_TIPO_A_DESCARGAR):
        return True, "tutela/derecho de peticion (encabezado del contenido)"
    return False, "no parece una tutela ni un derecho de peticion"


# ==================== Paso 1: carpeta de descargas manuales ====================


def _texto_de_archivo(ruta: Path) -> str:
    if ruta.suffix.lower() == ".pdf":
        return cruce_excel._texto_de_pdf(ruta)
    if ruta.suffix.lower() == ".docx":
        return cruce_excel._texto_de_docx(ruta)
    return ""


def clasificar_carpeta_descargas(indice_demandados, carpeta_raiz):
    carpeta_descargas = Path(CARPETA_DESCARGAS_MANUAL)
    if not carpeta_descargas.exists():
        logging.warning("[Descargas] No existe %s -- se omite el paso 1.", carpeta_descargas)
        return 0, [], []

    movidos = 0
    sin_coincidencia = []
    ambiguos = []

    for ruta in sorted(carpeta_descargas.rglob("*")):
        if not ruta.is_file() or ruta.suffix.lower() not in (".pdf", ".docx"):
            continue

        nombre_norm = buscador._normalizar_para_comparar(ruta.name)
        texto = _texto_de_archivo(ruta)
        contenido_norm = buscador._normalizar_para_comparar(ruta.name + " " + texto)

        coincidencias = _procesos_que_coinciden_fuerte(contenido_norm, indice_demandados)

        if not coincidencias:
            sin_coincidencia.append(ruta.name)
            continue

        if len(coincidencias) > 1:
            numeros = ", ".join(str(p["numero"]) for p in coincidencias)
            ambiguos.append((ruta.name, numeros))
            logging.warning(
                "[Ambiguo] '%s' coincide con mas de un proceso activo (%s) -- se deja donde esta, revisalo a mano.",
                ruta.name, numeros,
            )
            continue

        proceso = coincidencias[0]
        destino_carpeta = carpeta_raiz / proceso["nombre_carpeta"]

        if MODO_PRUEBA:
            logging.info(
                "[SIMULACION] '%s' -> proceso %s ('%s').", ruta.name, proceso["numero"], proceso["nombre_carpeta"],
            )
            movidos += 1
            continue

        base._crear_carpeta(destino_carpeta)
        destino_archivo = buscador._ruta_archivo_libre(destino_carpeta, organizador.sanear_nombre(ruta.name))
        try:
            ruta.rename(organizador._ruta_larga_segura(str(destino_archivo)))
            logging.info(
                "[Movido] '%s' -> proceso %s ('%s').", ruta.name, proceso["numero"], proceso["nombre_carpeta"],
            )
            movidos += 1
        except OSError as error:
            logging.warning("   (no se pudo mover '%s': %s)", ruta.name, error)

    return movidos, sin_coincidencia, ambiguos


# ==================== Paso 2: Gmail ====================


def _lotes(lista, tamano):
    for inicio in range(0, len(lista), tamano):
        yield lista[inicio:inicio + tamano]


def buscar_correo_por_demandados(usuario, app_password, demandados):
    """
    UNA sola conexion IMAP para TODOS los demandados -- los busca en
    lotes de TAMANO_LOTE_CORREO combinados con OR (misma sintaxis que
    la barra de busqueda de Gmail), en vez de abrir una conexion por
    cada uno. Devuelve {demandado: [correo, ...]}.
    """
    resultados_por_demandado = {}
    with imaplib.IMAP4_SSL("imap.gmail.com") as mail:
        mail.login(usuario, app_password)
        mail.select('"[Gmail]/All Mail"', readonly=True)

        for lote in _lotes(demandados, TAMANO_LOTE_CORREO):
            consulta = "(" + " OR ".join(f'"{d.replace(chr(34), "")}"' for d in lote) + ")"
            try:
                typ, datos = mail.search(None, "X-GM-RAW", consulta)
            except imaplib.IMAP4.error as error:
                logging.error("[Correo] Fallo buscando el lote %s: %s", lote, error)
                continue
            if typ != "OK" or not datos or not datos[0]:
                continue

            for id_correo in datos[0].split():
                correo = base._leer_correo(mail, id_correo)
                if correo is None:
                    continue
                contenido_norm = buscador._normalizar_para_comparar(
                    correo["asunto"] + " " + correo["cuerpo"] + " " + " ".join(n for n, _ in correo["adjuntos"])
                )
                for demandado in lote:
                    if _demandado_coincide_fuerte(contenido_norm, demandado):
                        resultados_por_demandado.setdefault(demandado, []).append(correo)
    return resultados_por_demandado


def procesar_correos_por_demandado(resultados_por_demandado, indice_demandados, carpeta_raiz):
    adjuntados = 0
    descartados = 0

    for demandado, correos in resultados_por_demandado.items():
        procesos = indice_demandados.get(demandado, [])
        if not procesos:
            continue

        for correo in correos:
            asunto_norm = buscador._normalizar_para_comparar(correo["asunto"])
            contenido_norm = buscador._normalizar_para_comparar(
                correo["asunto"] + " " + correo["cuerpo"] + " " + " ".join(n for n, _ in correo["adjuntos"])
            )
            es_tutela_o_peticion, motivo = _es_tutela_o_peticion(asunto_norm, contenido_norm)
            if not es_tutela_o_peticion:
                descartados += 1
                continue

            tiene_essa = any(buscador._nombre_coincide(contenido_norm, t) for t in buscador.TERMINOS_DEMANDANTE_VALIDO)
            if not tiene_essa:
                logging.info(
                    "   [Correo] se omite '%s' (demandado %s): no se confirmo que sea de ESSA/Electrificadora "
                    "de Santander.", correo["asunto"], demandado,
                )
                continue

            for proceso in procesos:
                numero, nombre_carpeta = proceso["numero"], proceso["nombre_carpeta"]
                destino = carpeta_raiz / nombre_carpeta

                if MODO_PRUEBA:
                    logging.info(
                        "[SIMULACION] Proceso %s (demandado %s): subiria el correo '%s' (%s) a '%s'.",
                        numero, demandado, correo["asunto"], motivo, nombre_carpeta,
                    )
                    adjuntados += 1
                    continue

                base._crear_carpeta(destino)
                guardados = base._guardar_correo(correo, destino)
                logging.info(
                    "[Descargado] Proceso %s (demandado %s): correo '%s' (%s) -> %d archivo(s) en %s",
                    numero, demandado, correo["asunto"], motivo, guardados, nombre_carpeta,
                )
                adjuntados += guardados

    logging.info(
        "[Correo] %d correo(s)/archivo(s) descargados; %d correo(s) coincidian con el demandado pero no eran "
        "tutela/derecho de peticion, se omitieron.", adjuntados, descartados,
    )


# ==================== Orquestacion ====================


def procesar():
    if not base.RUTA_EXCEL_CONTROL or not os.path.exists(base.RUTA_EXCEL_CONTROL):
        logging.error("No se encontro el Excel configurado en RUTA_EXCEL_CONTROL: %r", base.RUTA_EXCEL_CONTROL)
        return

    procesos = base.leer_procesos_control()
    con_radicado, _terminados, _sin_estado = base.clasificar_procesos(procesos)

    indice_demandados = _indexar_demandados(con_radicado)
    logging.info(
        "Excel: %d proceso(s) activo/suspendido/reorganizacion/remitida/etc, %d demandado(s) distinto(s) "
        "para cruzar.", len(con_radicado), len(indice_demandados),
    )

    carpeta_raiz = Path(base.CARPETA_PROCESOS)
    if not carpeta_raiz.exists():
        logging.error("No existe la carpeta configurada en CARPETA_PROCESOS: %s", carpeta_raiz)
        return

    logging.info("Paso 1: clasificando los archivos de %s por demandado...", CARPETA_DESCARGAS_MANUAL)
    movidos, sin_coincidencia, ambiguos = clasificar_carpeta_descargas(indice_demandados, carpeta_raiz)
    logging.info(
        "Paso 1: %d archivo(s) %s, %d sin ninguna coincidencia, %d ambiguo(s) (mas de un proceso posible).",
        movidos, "simulados para mover (MODO_PRUEBA activo)" if MODO_PRUEBA else "movidos",
        len(sin_coincidencia), len(ambiguos),
    )
    if sin_coincidencia:
        logging.warning("%d archivo(s) no coincidieron con ningun demandado activo:", len(sin_coincidencia))
        for nombre in sin_coincidencia:
            logging.warning("   - %s", nombre)

    if not BUSCAR_EN_CORREO:
        return

    credenciales = organizador.leer_credenciales()
    if not credenciales:
        logging.warning(
            "[Correo] No hay %s (o le faltan datos); se omite el paso 2 (busqueda en Gmail).",
            organizador.ARCHIVO_CREDENCIALES,
        )
        return

    logging.info(
        "Paso 2: buscando %d demandado(s) en todo Gmail (en lotes de %d, solo tutela/derecho de peticion)...",
        len(indice_demandados), TAMANO_LOTE_CORREO,
    )
    try:
        resultados = buscar_correo_por_demandados(*credenciales, list(indice_demandados.keys()))
    except Exception as error:
        logging.error("[Correo] Fallo la busqueda en Gmail, se omite: %s", error)
        return

    logging.info("[Correo] %d demandado(s) tuvieron al menos un correo que los menciona.", len(resultados))
    procesar_correos_por_demandado(resultados, indice_demandados, carpeta_raiz)

    if MODO_PRUEBA:
        logging.info(
            "MODO_PRUEBA esta activo: no se movio ningun archivo ni se descargo ningun correo todavia. "
            "Revisa el log y, si se ve bien, cambia MODO_PRUEBA = False al inicio de este script y vuelve a "
            "correrlo."
        )


def main():
    configurar_logging()
    procesar()


if __name__ == "__main__":
    main()
