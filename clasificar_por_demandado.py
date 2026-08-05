"""
Clasifica información extraprocesal (derecho de petición, tutela, y
pago oficioso -- tanto lo presentado como las RESPUESTAS que da la
entidad a la que se envió) por proceso, en dos pasos independientes
(ambos respetan MODO_PRUEBA):

1. CARPETA DE DESCARGAS MANUALES (CARPETA_DESCARGAS_MANUAL, por
   defecto tu carpeta "Downloads"): revisa cada PDF/DOCX que haya ahí y
   lo MUEVE directo a la carpeta del proceso que le corresponde. No hay
   ninguna restricción de TIPO de documento aquí (a diferencia de
   clasificar_procesos_ejecutivos.py) -- se asume que ya es información
   extraprocesal porque tú la descargaste a propósito; el único
   trabajo de este script es encontrar a qué proceso corresponde. Por
   defecto (SOLO_DESCARGAS_DE_HOY) solo revisa los archivos creados o
   modificados HOY -- una carpeta de Descargas normal acumula años de
   archivos de todo tipo (demandas, autos, etc, no solo lo que bajaste
   hoy), y no tiene sentido volver a revisarlos en cada corrida.

2. GMAIL: busca en TODO tu Gmail (no solo la bandeja de entrada) por
   el nombre de cada demandado, el radicado (y sus formas cortas), y
   la cuenta de cada proceso activo -- no se limita al nombre exacto
   del demandado: si el nombre no aparece en un correo pero sí su
   radicado o su cuenta, también se encuentra. Del resultado SOLO se
   descarga lo que además sea un derecho de petición, una tutela, o un
   pago oficioso (ver clasificar_procesos_ejecutivos.es_informacion_no_procesal
   -- incluye tanto lo presentado como las respuestas que da la
   entidad), igual de riguroso que el resto del proyecto. También se
   exige, igual que en el resto del proyecto, que el correo mencione a
   ESSA/Electrificadora de Santander.

   En vez de abrir una conexión de Gmail por cada término (con cientos
   de procesos activos eso tardaría horas), se abre UNA sola conexión
   y se buscan todos los términos en LOTES combinados con OR (ver
   TAMANO_LOTE_CORREO) -- Gmail permite eso mismo en su propia barra
   de búsqueda.

En AMBOS pasos, a qué proceso corresponde un documento/correo se
decide con la MISMA regla que usa clasificar_procesos_ejecutivos.py
para emparejar el correo global con un proceso: coincide su radicado,
su cuenta, O el demandado COMPLETO (todas sus palabras significativas,
no basta con una sola -- ver MIN_PALABRAS_DEMANDADO_PARA_CRUZAR en ese
mismo script; evita que una palabra suelta/generica de un demandado mal
diligenciado en el Excel "coincida" con archivos que no tienen nada que
ver) -- basta con que coincida CUALQUIERA de los tres, no hace falta
que coincidan todos (ver
clasificar_procesos_ejecutivos._procesos_que_coinciden_con_correo). Si
YA se encuentra el proceso, la carpeta se CREA sola si todavía no
existe (mismo nombre "<numero>. <radicado o ESTADO>" que usa
clasificar_procesos_ejecutivos.py) -- no hace falta haberla creado
antes. Si un archivo/correo no coincide con ningún proceso activo, se
deja intacto y queda registrado en un CSV (ARCHIVO_SIN_COINCIDENCIA /
ARCHIVO_CORREO_SIN_COINCIDENCIA) para que lo revises a mano -- nunca se
adivina. Si coincide con MÁS de uno (ej. dos procesos activos contra el
mismo demandado), primero se intenta desambiguar por el RADICADO: si el
texto menciona el radicado específico de UNO SOLO de los candidatos, se
usa ese (ver _desambiguar_por_radicado); si sigue sin poder decidirse,
recién ahí se deja intacto y queda en ARCHIVO_AMBIGUOS -- junto con el
MOTIVO exacto de cada coincidencia (ver _motivo_coincidencia: si fue
por radicado, por cuenta, y/o cuál demandado exacto con qué palabras),
para poder diagnosticar de verdad la causa (ej. confirmar si el dato de
DEMANDADO de ese proceso en el Excel está bien diligenciado o no) en
vez de tener que adivinarla.

Rendimiento y precisión (paso 1): para cada archivo primero se prueba
la coincidencia SOLO con su NOMBRE (rápido, sin abrir nada) -- solo si
el nombre no basta para encontrar ningún proceso se descarga/lee su
contenido (PDF/DOCX), que es lo más lento. Esto además es más preciso:
el contenido completo de un documento largo puede coincidir con más
procesos de los que el nombre por sí solo sugeriría (nunca con menos),
así que cuando el nombre ya da una única coincidencia clara, usar esa
evita diluirla con texto de más.

Reutiliza toda la lectura del Excel y las carpetas de
clasificar_procesos_ejecutivos.py (misma hoja ACTIVOS, mismo
CARPETA_PROCESOS) -- solo considera procesos ACTIVOS (todo lo que no
es terminado/no inicio), igual que la regla de "información no
procesal" del resto del proyecto. Requiere las mismas credenciales
credenciales_sgde.txt para Gmail (ver README) -- si no existe, el
paso 2 se omite solo, sin error.
"""

import datetime
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

# True (por defecto): en la carpeta de descargas, SOLO revisa los
# archivos modificados/descargados HOY -- una carpeta de Descargas
# normal acumula años de archivos de todo tipo (demandas, autos,
# mandamientos viejos, etc, no solo la información extraprocesal que
# bajaste hoy), y no tiene sentido volver a revisarlos en cada corrida.
# Ponlo en False si quieres que revise TODOS los archivos sin importar
# la fecha.
SOLO_DESCARGAS_DE_HOY = True

# True (por defecto): además de la carpeta de descargas, busca en TODO
# tu Gmail. Si no existe credenciales_sgde.txt, este paso se omite
# solo, sin error.
BUSCAR_EN_CORREO = True

# Cuántos términos (demandados + radicados + cuentas) se combinan en
# UNA sola búsqueda de Gmail (con OR) -- evita abrir una
# conexión/búsqueda separada por cada uno de los cientos de términos
# de los procesos activos, que sería muy lento.
TAMANO_LOTE_CORREO = 15

# True (por defecto): no mueve archivos ni descarga correos de verdad,
# solo revisa y muestra qué haría.
MODO_PRUEBA = True

ARCHIVO_LOG = os.path.join(os.path.dirname(__file__), "clasificar_por_demandado.log")

# Archivos de CARPETA_DESCARGAS_MANUAL que no coincidieron con ningun
# proceso activo, o que coincidieron con mas de uno (no se movieron,
# quedan para que los revises a mano).
ARCHIVO_SIN_COINCIDENCIA = os.path.join(os.path.dirname(__file__), "clasificar_por_demandado_sin_coincidencia.csv")
ARCHIVO_AMBIGUOS = os.path.join(os.path.dirname(__file__), "clasificar_por_demandado_ambiguos.csv")

# Correos de Gmail que SI eran tutela/derecho de peticion/pago
# oficioso, pero no coincidieron con el radicado/cuenta/demandado de
# ningun proceso activo conocido (no se descargaron).
ARCHIVO_CORREO_SIN_COINCIDENCIA = os.path.join(os.path.dirname(__file__), "clasificar_por_demandado_correo_sin_coincidencia.csv")

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


# ==================== Paso 1: carpeta de descargas manuales ====================


def _es_de_hoy(ruta: Path) -> bool:
    """
    True si 'ruta' se creo o se modifico HOY -- se revisan las dos
    fechas (no solo una) porque algunos adjuntos de correo conservan la
    fecha de modificacion ORIGINAL del remitente (no la de cuando tu lo
    descargaste), mientras que la fecha de creacion en el disco si
    refleja cuando el archivo llego a tu carpeta de Descargas.
    """
    hoy = datetime.date.today()
    info = ruta.stat()
    return hoy in (datetime.date.fromtimestamp(info.st_ctime), datetime.date.fromtimestamp(info.st_mtime))


def _texto_de_archivo(ruta: Path) -> str:
    if ruta.suffix.lower() == ".pdf":
        return cruce_excel._texto_de_pdf(ruta)
    if ruta.suffix.lower() == ".docx":
        return cruce_excel._texto_de_docx(ruta)
    return ""


def _desambiguar_por_radicado(texto_normalizado, coincidencias):
    """
    Si 'coincidencias' tiene mas de un proceso posible, intenta
    reducirlo a UNO SOLO revisando si el RADICADO (o su forma corta) de
    alguno de ellos aparece en el texto -- si trae el radicado
    especifico de UN SOLO candidato, se usa ese en vez de quedar
    ambiguo (ej. varios procesos con demandado "AUTO ACEPTA RETIRO..."
    -- si el documento menciona el radicado exacto de uno de ellos, ya
    no hace falta revisarlo a mano). Si el radicado de MAS de un
    candidato aparece, o si no aparece ninguno, se deja igual (sigue
    ambiguo, se reporta para revision manual).
    """
    if len(coincidencias) <= 1:
        return coincidencias

    con_radicado_en_texto = [
        proceso for proceso in coincidencias
        if proceso["radicado"] and any(
            buscador._nombre_coincide(texto_normalizado, termino)
            for termino in [proceso["radicado"]] + buscador.radicados_cortos(proceso["radicado"])
        )
    ]
    if len(con_radicado_en_texto) == 1:
        return con_radicado_en_texto
    return coincidencias


def _procesos_para_archivo(ruta, indices):
    """
    Encuentra los procesos que coinciden con 'ruta', priorizando
    SOLO el nombre del archivo (rapido, sin abrir nada) antes de leer
    su contenido: si el nombre YA da una unica coincidencia clara, no
    hace falta descargar/leer el PDF/DOCX -- eso ahorra la parte mas
    lenta (extraer texto) para la mayoria de los archivos con un
    nombre descriptivo (ej. "peticion alberto suarez.pdf"). Ademas es
    mas PRECISO: el contenido completo de un documento largo puede
    coincidir por casualidad con mas procesos de los que el nombre por
    si solo sugeriria (nunca con MENOS -- el contenido normalizado
    siempre incluye el nombre), asi que si el nombre ya da una unica
    coincidencia clara, usar esa es mas confiable que diluirla con el
    resto del texto.

    Si el nombre NO da ninguna coincidencia, o queda AMBIGUO (varios
    procesos posibles), se lee el contenido: puede aportar una
    coincidencia nueva, o -- si ya habia varias candidatas -- un
    radicado especifico que desambigue cual de ellas es la correcta
    (ver _desambiguar_por_radicado).

    Devuelve (coincidencias, texto_usado) -- 'texto_usado' es el texto
    normalizado que de verdad se uso para decidir (solo el nombre, o
    nombre+contenido), para poder diagnosticar despues POR QUE cada
    proceso coincidio (ver _motivo_coincidencia).
    """
    nombre_norm = buscador._normalizar_para_comparar(ruta.name)
    coincidencias = base._procesos_que_coinciden_con_correo(nombre_norm, indices)
    if len(coincidencias) == 1:
        return coincidencias, nombre_norm

    texto = _texto_de_archivo(ruta)
    contenido_norm = buscador._normalizar_para_comparar(ruta.name + " " + texto)
    coincidencias_contenido = base._procesos_que_coinciden_con_correo(contenido_norm, indices)

    candidatos = coincidencias_contenido if coincidencias_contenido else coincidencias
    return _desambiguar_por_radicado(contenido_norm, candidatos), contenido_norm


def _motivo_coincidencia(texto_normalizado, proceso):
    """
    Diagnostico: POR QUE 'proceso' aparece entre las coincidencias de
    'texto_normalizado' -- coincidio su radicado, su cuenta, y/o cual(es)
    de sus demandados (con las palabras exactas que coincidieron). Se
    usa solo al reportar un caso ambiguo, para poder ver en el
    log/CSV la razon EXACTA de cada coincidencia -- por ejemplo, para
    confirmar si el dato de DEMANDADO en el Excel para ese proceso esta
    bien diligenciado o no, en vez de tener que adivinarlo.
    """
    motivos = []
    radicado = proceso.get("radicado")
    if radicado:
        terminos = [radicado] + buscador.radicados_cortos(radicado)
        if any(buscador._nombre_coincide(texto_normalizado, t) for t in terminos):
            motivos.append(f"radicado={radicado}")
    for cuenta in proceso.get("cuentas", []):
        if buscador._cuenta_es_valida_para_buscar(cuenta) and buscador._nombre_coincide(texto_normalizado, cuenta):
            motivos.append(f"cuenta={cuenta}")
    for demandado in proceso.get("demandados", []):
        palabras = buscador._palabras_significativas(demandado)
        if len(palabras) >= base.MIN_PALABRAS_DEMANDADO_PARA_CRUZAR and all(
            base._palabra_demandado_coincide(texto_normalizado, p) for p in palabras
        ):
            motivos.append(f"demandado='{demandado}'")
    return "; ".join(motivos) if motivos else "razon desconocida (revisa manualmente)"


def clasificar_carpeta_descargas(indices, carpeta_raiz):
    carpeta_descargas = Path(CARPETA_DESCARGAS_MANUAL)
    if not carpeta_descargas.exists():
        logging.warning("[Descargas] No existe %s -- se omite el paso 1.", carpeta_descargas)
        return 0, [], []

    movidos = 0
    sin_coincidencia = []
    ambiguos = []
    omitidos_por_fecha = 0

    for ruta in sorted(carpeta_descargas.rglob("*")):
        if not ruta.is_file() or ruta.suffix.lower() not in (".pdf", ".docx"):
            continue

        if SOLO_DESCARGAS_DE_HOY and not _es_de_hoy(ruta):
            omitidos_por_fecha += 1
            continue

        coincidencias, texto_usado = _procesos_para_archivo(ruta, indices)

        if not coincidencias:
            sin_coincidencia.append(ruta.name)
            continue

        if len(coincidencias) > 1:
            detalle = "; ".join(
                f"{p['numero']} ({_motivo_coincidencia(texto_usado, p)})" for p in coincidencias
            )
            numeros = ", ".join(str(p["numero"]) for p in coincidencias)
            ambiguos.append((ruta.name, numeros, detalle))
            logging.warning(
                "[Ambiguo] '%s' coincide con mas de un proceso activo -- se deja donde esta, revisalo a mano: %s",
                ruta.name, detalle,
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

    if SOLO_DESCARGAS_DE_HOY and omitidos_por_fecha:
        logging.info(
            "[Descargas] %d archivo(s) de otras fechas se omitieron (SOLO_DESCARGAS_DE_HOY activo) --"
            " ponlo en False si quieres revisar TODA la carpeta.", omitidos_por_fecha,
        )

    return movidos, sin_coincidencia, ambiguos


# ==================== Paso 2: Gmail ====================


def _lotes(lista, tamano):
    for inicio in range(0, len(lista), tamano):
        yield lista[inicio:inicio + tamano]


def _terminos_de_busqueda(con_radicado):
    """
    TODOS los terminos de busqueda de los procesos activos, sin
    duplicados: nombres de demandado + radicado (y sus formas cortas)
    + cuentas validas -- ver clasificar_procesos_ejecutivos._terminos_busqueda
    para el radicado/cuenta. No se limita al demandado: si un correo no
    lo menciona pero si el radicado o la cuenta, tambien se encuentra.
    """
    vistos = set()
    terminos = []
    for proceso in con_radicado:
        candidatos = list(proceso["demandados"]) + base._terminos_busqueda(proceso["radicado"], proceso["cuentas"])
        for termino in candidatos:
            termino = (termino or "").strip()
            clave = termino.upper()
            if termino and clave not in vistos:
                vistos.add(clave)
                terminos.append(termino)
    return terminos


def buscar_correo_por_procesos(usuario, app_password, con_radicado):
    """
    UNA sola conexion IMAP para TODOS los procesos activos -- busca,
    en lotes combinados con OR (ver TAMANO_LOTE_CORREO), los demandados
    + radicados + cuentas de _terminos_de_busqueda. Devuelve la lista
    de correos encontrados, sin duplicados (por asunto+fecha) -- a cual
    proceso corresponde cada uno se decide despues, con la misma regla
    de coincidencia que el resto del proyecto (ver
    clasificar_procesos_ejecutivos._procesos_que_coinciden_con_correo).
    """
    terminos = _terminos_de_busqueda(con_radicado)
    logging.info("[Correo] %d termino(s) distinto(s) para buscar (demandados + radicados + cuentas).", len(terminos))

    vistos = {}
    with imaplib.IMAP4_SSL("imap.gmail.com") as mail:
        mail.login(usuario, app_password)
        mail.select('"[Gmail]/All Mail"', readonly=True)

        for lote in _lotes(terminos, TAMANO_LOTE_CORREO):
            consulta = "(" + " OR ".join(f'"{t.replace(chr(34), "")}"' for t in lote) + ")"
            try:
                typ, datos = mail.search(None, "X-GM-RAW", consulta)
            except imaplib.IMAP4.error as error:
                logging.error("[Correo] Fallo buscando el lote %s: %s", lote, error)
                continue
            if typ != "OK" or not datos or not datos[0]:
                continue

            for id_correo in datos[0].split():
                correo = base._leer_correo(mail, id_correo)
                if correo is not None:
                    vistos[(correo["asunto"], correo["fecha"])] = correo

    return list(vistos.values())


def procesar_correos(correos, con_radicado, carpeta_raiz):
    indices = base._indexar_procesos_para_correo(con_radicado)
    adjuntados = 0
    descartados = 0
    sin_proceso = []

    for correo in correos:
        asunto_norm = buscador._normalizar_para_comparar(correo["asunto"])
        contenido_norm = buscador._normalizar_para_comparar(
            correo["asunto"] + " " + correo["cuerpo"] + " " + " ".join(n for n, _ in correo["adjuntos"])
        )
        es_no_procesal, motivo = base.es_informacion_no_procesal(asunto_norm, contenido_norm)
        if not es_no_procesal:
            descartados += 1
            continue

        procesos_coincidentes = base._procesos_que_coinciden_con_correo(contenido_norm, indices)
        procesos_coincidentes = _desambiguar_por_radicado(contenido_norm, procesos_coincidentes)
        if not procesos_coincidentes:
            fecha_texto = correo["fecha"].isoformat() if correo["fecha"] else ""
            sin_proceso.append((correo["asunto"], fecha_texto, motivo))
            logging.info(
                "   [Correo] '%s' parece %s, pero no coincide con el radicado/cuenta/demandado de ningun "
                "proceso activo -- no se pudo emparejar, se omite.", correo["asunto"], motivo,
            )
            continue

        tiene_essa = any(buscador._nombre_coincide(contenido_norm, t) for t in buscador.TERMINOS_DEMANDANTE_VALIDO)
        if not tiene_essa:
            logging.info(
                "   [Correo] se omite '%s': no se confirmo que sea de ESSA/Electrificadora de Santander.",
                correo["asunto"],
            )
            continue

        for proceso in procesos_coincidentes:
            numero, nombre_carpeta = proceso["numero"], proceso["nombre_carpeta"]
            destino = carpeta_raiz / nombre_carpeta

            if MODO_PRUEBA:
                logging.info(
                    "[SIMULACION] Proceso %s: subiria el correo '%s' (%s) a '%s'.",
                    numero, correo["asunto"], motivo, nombre_carpeta,
                )
                adjuntados += 1
                continue

            base._crear_carpeta(destino)
            guardados = base._guardar_correo(correo, destino)
            logging.info(
                "[Descargado] Proceso %s: correo '%s' (%s) -> %d archivo(s) en %s",
                numero, correo["asunto"], motivo, guardados, nombre_carpeta,
            )
            adjuntados += guardados

    logging.info(
        "[Correo] %d correo(s)/archivo(s) descargados; %d correo(s) no eran tutela/derecho de peticion/pago "
        "oficioso; %d coincidian con el tipo pero no con ningun proceso activo conocido.",
        adjuntados, descartados, len(sin_proceso),
    )
    return sin_proceso


# ==================== Orquestacion ====================


def procesar():
    if not base.RUTA_EXCEL_CONTROL or not os.path.exists(base.RUTA_EXCEL_CONTROL):
        logging.error("No se encontro el Excel configurado en RUTA_EXCEL_CONTROL: %r", base.RUTA_EXCEL_CONTROL)
        return

    procesos = base.leer_procesos_control()
    con_radicado, _terminados, _sin_estado = base.clasificar_procesos(procesos)
    logging.info(
        "Excel: %d proceso(s) activo/suspendido/reorganizacion/remitida/etc para cruzar.", len(con_radicado),
    )

    carpeta_raiz = Path(base.CARPETA_PROCESOS)
    if not carpeta_raiz.exists():
        logging.error("No existe la carpeta configurada en CARPETA_PROCESOS: %s", carpeta_raiz)
        return

    indices = base._indexar_procesos_para_correo(con_radicado)

    logging.info("Paso 1: clasificando los archivos de %s...", CARPETA_DESCARGAS_MANUAL)
    movidos, sin_coincidencia, ambiguos = clasificar_carpeta_descargas(indices, carpeta_raiz)
    logging.info(
        "Paso 1: %d archivo(s) %s, %d sin ninguna coincidencia, %d ambiguo(s) (mas de un proceso posible).",
        movidos, "simulados para mover (MODO_PRUEBA activo)" if MODO_PRUEBA else "movidos",
        len(sin_coincidencia), len(ambiguos),
    )
    if sin_coincidencia and cruce_excel._escribir_csv_tolerante(
        ARCHIVO_SIN_COINCIDENCIA, ["Archivo"], [(n,) for n in sin_coincidencia], "Sin coincidencia (descargas)",
    ):
        logging.info("[Reporte] %d archivo(s) sin coincidencia guardados en: %s", len(sin_coincidencia), ARCHIVO_SIN_COINCIDENCIA)
    if ambiguos and cruce_excel._escribir_csv_tolerante(
        ARCHIVO_AMBIGUOS, ["Archivo", "Procesos posibles", "Por que coincidio cada uno"], ambiguos, "Ambiguos (descargas)",
    ):
        logging.info("[Reporte] %d archivo(s) ambiguo(s) guardados en: %s", len(ambiguos), ARCHIVO_AMBIGUOS)

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
        "Paso 2: buscando en todo Gmail (demandados + radicados + cuentas, en lotes de %d, tutela/derecho "
        "de peticion/pago oficioso)...", TAMANO_LOTE_CORREO,
    )
    try:
        correos = buscar_correo_por_procesos(*credenciales, con_radicado)
    except Exception as error:
        logging.error("[Correo] Fallo la busqueda en Gmail, se omite: %s", error)
        return

    logging.info("[Correo] %d correo(s) encontrados en total -- clasificando y emparejando...", len(correos))
    correo_sin_proceso = procesar_correos(correos, con_radicado, carpeta_raiz)
    if correo_sin_proceso and cruce_excel._escribir_csv_tolerante(
        ARCHIVO_CORREO_SIN_COINCIDENCIA, ["Asunto", "Fecha", "Motivo"], correo_sin_proceso, "Sin coincidencia (correo)",
    ):
        logging.info(
            "[Reporte] %d correo(s) sin coincidencia guardados en: %s", len(correo_sin_proceso), ARCHIVO_CORREO_SIN_COINCIDENCIA,
        )

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
