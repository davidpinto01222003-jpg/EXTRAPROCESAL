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
   la misma razón, este paso busca entre TODOS los procesos del Excel,
   ACTIVOS y TERMINADOS (a diferencia del paso 2, que solo busca entre
   los activos) -- también puedes soltar aquí el auto que termina un
   proceso ya terminado (ej. un desistimiento tácito), no solo
   información extraprocesal de un proceso activo. Por defecto
   (SOLO_DESCARGAS_DE_HOY) solo revisa los archivos creados o
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
mismo demandado), SIEMPRE se intenta desambiguar por el RADICADO: si el
texto menciona el radicado específico de UNO SOLO de los candidatos, se
usa ese (ver _desambiguar_por_radicado) -- y el intento (haya
funcionado o no) SIEMPRE queda en el log con el prefijo "[Radicado]",
para que quede claro que sí se revisó aunque no siempre alcance a
desambiguar (un auto corto no siempre repite su propio radicado en el
texto). Si sigue sin poder decidirse, recién ahí se deja intacto y
queda en ARCHIVO_AMBIGUOS -- junto con el MOTIVO exacto de cada
coincidencia (ver _motivo_coincidencia: si fue por radicado, por
cuenta, y/o cuál demandado exacto con qué palabras),
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
CARPETA_PROCESOS). El paso 2 (Gmail) SÍ se restringe solo a procesos
ACTIVOS (todo lo que no es terminado/no inicio), igual que la regla de
"información no procesal" del resto del proyecto -- ahí solo se
descarga tutela/petición/pago oficioso, que no aplica a un proceso ya
terminado. Requiere las mismas credenciales credenciales_sgde.txt para
Gmail (ver README) -- si no existe, el paso 2 se omite solo, sin error.
"""

import datetime
import imaplib
import logging
import os
import re
import threading
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

# Con cientos de procesos activos hay que hacer CIENTOS de búsquedas
# (una por lote) sobre la MISMA conexión IMAP -- Gmail corta la
# conexión si la nota con demasiadas búsquedas/descargas seguidas en
# poco tiempo (proteccion propia de Gmail contra scripts, no depende
# de nada que puedas configurar en el codigo). Cuando eso pasa, en vez
# de rendirse a mitad de camino, se abre una conexión NUEVA (login +
# seleccionar la carpeta otra vez) y se sigue justo donde se quedó --
# esto es cuántas veces se reintenta reconectar en TODA la búsqueda
# antes de darse por vencido de verdad.
MAX_RECONEXIONES_CORREO = 8

# Límite (en segundos) para cualquier operación de red con Gmail
# (conectar, login, buscar, descargar un correo). Sin esto, un socket
# de Python espera una respuesta PARA SIEMPRE si la conexión queda en
# un estado "a medias" -- se ve como el programa colgado, sin ningún
# error ni progreso en el log.
TIMEOUT_CORREO_SEGUNDOS = 30

# Cada cuántos lotes se deja un aviso de "sigo trabajando" en el log
# durante la búsqueda en Gmail. Con cientos de lotes, pasar varios
# minutos sin NINGUNA línea nueva (porque todo va bien, simplemente
# toma tiempo) se ve igual que el programa colgado.
AVISO_PROGRESO_CORREO_LOTES = 10

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


def _desambiguar_por_radicado(texto_normalizado, coincidencias, contexto=""):
    """
    Si 'coincidencias' tiene mas de un proceso posible, intenta
    reducirlo a UNO SOLO revisando el RADICADO. Primero busca el
    radicado COMPLETO (plano O con guiones/puntos/espacios entre sus
    grupos -- ver cruce_excel._radicados_en_texto): si el radicado
    exacto de UN SOLO candidato aparece, se usa ese de inmediato, aunque
    el texto TAMBIEN contenga -- por pura coincidencia -- la forma
    CORTA de otro candidato (dos juzgados distintos pueden repetir el
    mismo año-consecutivo, solo cambia el codigo del juzgado o la
    instancia). Solo si NINGUN candidato tiene su radicado completo en
    el texto se recurre a la forma corta como ultimo recurso. Si el
    radicado (completo o corto) de MAS de un candidato aparece, o si no
    aparece ninguno (un auto puede no repetir su propio radicado en el
    texto), se deja igual (sigue ambiguo, se reporta para revision
    manual). SIEMPRE deja en el log si lo intento y por que no alcanzo,
    para que quede claro que el radicado SI se revisa, aunque no
    siempre pueda desambiguar.
    """
    if len(coincidencias) <= 1:
        return coincidencias

    # El radicado dentro de un documento casi nunca aparece "plano" --
    # normalmente trae guiones/puntos/espacios entre sus grupos (ej.
    # "68001-40-03-001-2024-00050-00"). cruce_excel._radicados_en_texto
    # reconoce esa forma (y la plana) y ya devuelve el numero limpio de
    # separadores, listo para comparar contra el radicado del Excel.
    radicados_del_texto = set(cruce_excel._radicados_en_texto(texto_normalizado))

    # El radicado COMPLETO (23 digitos) es la señal mas fuerte -- se
    # revisa primero y, si señala a UN solo candidato, se usa ese aunque
    # el texto TAMBIEN contenga, por pura coincidencia, la forma CORTA
    # de otro candidato: dos juzgados distintos pueden repetir el mismo
    # año-consecutivo (solo cambia el codigo del juzgado o la
    # instancia), asi que una forma corta compartida no puede pesar
    # igual que el radicado completo y exacto de un candidato.
    con_radicado_exacto = [
        proceso for proceso in coincidencias
        if proceso["radicado"] and proceso["radicado"] in radicados_del_texto
    ]
    con_radicado_en_texto = con_radicado_exacto or [
        proceso for proceso in coincidencias
        if proceso["radicado"] and any(
            buscador._nombre_coincide(texto_normalizado, t) for t in buscador.radicados_cortos(proceso["radicado"])
        )
    ]
    if len(con_radicado_en_texto) == 1:
        logging.info(
            "   [Radicado] '%s': de %d proceso(s) posibles por demandado, el radicado del texto senala al proceso "
            "%s -- ya no queda ambiguo.", contexto, len(coincidencias), con_radicado_en_texto[0]["numero"],
        )
        return con_radicado_en_texto

    if con_radicado_en_texto:
        logging.info(
            "   [Radicado] '%s': el texto menciona el radicado de %d de los %d candidatos -- sigue ambiguo, no se "
            "puede elegir uno solo.", contexto, len(con_radicado_en_texto), len(coincidencias),
        )
    elif radicados_del_texto:
        logging.info(
            "   [Radicado] '%s': el texto SI trae radicado(s) (%s), pero ninguno coincide con los %d candidatos "
            "por demandado -- sigue ambiguo (revisa si alguno de esos radicados deberia estar en el Excel).",
            contexto, ", ".join(sorted(radicados_del_texto)), len(coincidencias),
        )
    else:
        logging.info(
            "   [Radicado] '%s': no se encontro NINGUN radicado en el texto -- sigue ambiguo.%s",
            contexto, _fragmento_radicado_en_texto(texto_normalizado),
        )
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
    return _desambiguar_por_radicado(contenido_norm, candidatos, contexto=ruta.name), contenido_norm


def _fragmento_contexto(texto: str, palabra: str, ventana: int = 50) -> str:
    """Un pedazo de 'texto' alrededor de donde aparece 'palabra' (como palabra completa), para poder ver EN QUE CONTEXTO coincidio."""
    coincidencia = re.search(rf"(?<![A-ZÑ]){re.escape(palabra)}(?![A-ZÑ])", texto)
    if not coincidencia:
        return ""
    inicio = max(0, coincidencia.start() - ventana)
    fin = min(len(texto), coincidencia.end() + ventana)
    return f' ("...{texto[inicio:fin].strip()}...")'


def _fragmento_radicado_en_texto(texto: str, ventana: int = 80) -> str:
    """
    Un pedazo de 'texto' alrededor de la palabra "RADICADO" (o "RAD.",
    abreviatura comun) -- se usa SOLO como diagnostico, cuando ni
    siquiera se reconocio ningun radicado de 23 digitos en el texto,
    para poder ver como esta escrito de verdad ahi (y ajustar el
    reconocimiento si hace falta) en vez de seguir adivinando.
    """
    coincidencia = re.search(r"RADICAD[OA]|\bRAD\.?\s*(?:NO|N)\b", texto)
    if not coincidencia:
        return ""
    inicio = max(0, coincidencia.start() - 10)
    fin = min(len(texto), coincidencia.end() + ventana)
    return f' Texto cerca de "radicado": "...{texto[inicio:fin].strip()}..."'


def _motivo_coincidencia(texto_normalizado, proceso):
    """
    Diagnostico: POR QUE 'proceso' aparece entre las coincidencias de
    'texto_normalizado' -- coincidio su radicado, su cuenta, y/o cual(es)
    de sus demandados (con las palabras exactas que coincidieron, y un
    pedazo del texto donde aparecieron, para ver el contexto real). Se
    usa solo al reportar un caso ambiguo, para poder ver en el
    log/CSV la razon EXACTA de cada coincidencia -- por ejemplo, para
    confirmar si el dato de DEMANDADO en el Excel para ese proceso esta
    bien diligenciado, o si el documento simplemente menciona ese
    nombre en otro contexto (ej. el nombre del JUZGADO), en vez de
    tener que adivinarlo. Usa la misma "ventana" que la busqueda real
    (solo el encabezado del documento para el demandado, ver
    clasificar_procesos_ejecutivos.VENTANA_DEMANDADO_CARACTERES) para
    que el diagnostico sea fiel a lo que de verdad decidio el cruce.
    """
    motivos = []
    radicado = proceso.get("radicado")
    if radicado:
        radicados_del_texto = set(cruce_excel._radicados_en_texto(texto_normalizado))
        coincide = radicado in radicados_del_texto or any(
            buscador._nombre_coincide(texto_normalizado, t) for t in buscador.radicados_cortos(radicado)
        )
        if coincide:
            motivos.append(f"radicado={radicado}")
    for cuenta in proceso.get("cuentas", []):
        if buscador._cuenta_es_valida_para_buscar(cuenta) and buscador._nombre_coincide(texto_normalizado, cuenta):
            motivos.append(f"cuenta={cuenta}")
    encabezado = base._quitar_membrete_juzgado(texto_normalizado)[:base.VENTANA_DEMANDADO_CARACTERES]
    for demandado in proceso.get("demandados", []):
        palabras = buscador._palabras_significativas(demandado)
        if len(palabras) >= base.MIN_PALABRAS_DEMANDADO_PARA_CRUZAR and all(
            base._palabra_demandado_coincide(encabezado, p) for p in palabras
        ):
            palabra_mas_larga = max(palabras, key=len)
            motivos.append(f"demandado='{demandado}'{_fragmento_contexto(encabezado, palabra_mas_larga)}")
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


def _conectar_gmail(usuario, app_password):
    """
    Abre una conexion IMAP nueva, hace login, y selecciona la carpeta
    de "todos los correos" -- usado tanto para la conexion inicial como
    para RECONECTAR si Gmail corta la conexion a mitad de la busqueda
    (ver buscar_correo_por_procesos). Devuelve la conexion lista para
    usar, o None si la conexion, el login, o el select fallaron.

    CON TIMEOUT explicito (TIMEOUT_CORREO_SEGUNDOS): sin esto, un socket
    de Python se queda esperando una respuesta PARA SIEMPRE si la
    conexion queda en un estado "a medias" (ni cerrada del todo ni
    respondiendo) -- que es justo lo que le pasa a veces a la conexion
    vieja despues de que Gmail la corta. Eso se ve como el programa
    "colgado" sin ningun error ni progreso en el log. Con el timeout, si
    no hay respuesta a tiempo salta un error normal (que este mismo
    fallo ya sabe manejar) en vez de quedarse esperando para siempre.
    Ademas TODA la llamada va dentro del try -- antes "imaplib.IMAP4_SSL(...)"
    (que ya intenta conectar/hacer el handshake TLS) quedaba POR FUERA
    del try, asi que un fallo justo ahi tampoco se atrapaba.
    """
    try:
        # Logs paso a paso: si algun dia esto se vuelve a "colgar", el
        # ultimo mensaje que se alcance a ver en el log dice EXACTAMENTE
        # en cual paso se quedo (conectar, iniciar sesion, o seleccionar
        # la carpeta) en vez de tener que adivinar.
        logging.info("[Correo] Conectando con Gmail (timeout %ds)...", TIMEOUT_CORREO_SEGUNDOS)
        mail = imaplib.IMAP4_SSL("imap.gmail.com", timeout=TIMEOUT_CORREO_SEGUNDOS)
        logging.info("[Correo] Conectado -- iniciando sesion...")
        mail.login(usuario, app_password)
        logging.info("[Correo] Sesion iniciada -- seleccionando la carpeta de correos...")
    except Exception as error:
        logging.error("[Correo] No se pudo conectar con Gmail: %s", error)
        return None
    if not buscador.seleccionar_todos_los_correos(mail):
        logging.error("[Correo] No se pudo seleccionar la carpeta de 'Todos los correos' de Gmail.")
        return None
    logging.info("[Correo] Carpeta seleccionada -- listo para seguir buscando.")
    return mail


def _cerrar_sesion_gmail(mail):
    """
    Intenta un LOGOUT limpio, pero sin arriesgarse a colgar el
    programa: si la conexion ya quedo en un estado raro, un logout()
    sin timeout se puede quedar esperando una respuesta que nunca
    llega. Cualquier error (o timeout) se ignora -- ya no importa un
    cierre prolijo, solo que no se cuelgue.
    """
    try:
        sock = getattr(mail, "sock", None)
        if sock is not None:
            sock.settimeout(TIMEOUT_CORREO_SEGUNDOS)
        mail.logout()
    except Exception:
        pass


def _con_limite_de_tiempo_duro(mail, funcion, *args, **kwargs):
    """
    Ejecuta 'funcion' en un hilo DAEMON aparte con un limite duro de
    TIMEOUT_CORREO_SEGUNDOS -- una SEGUNDA linea de defensa ademas del
    timeout que ya tiene el socket (ver TIMEOUT_CORREO_SEGUNDOS en
    _conectar_gmail). El timeout del socket depende de que Python
    detecte "no llego nada en N segundos"; en la practica eso puede NO
    dispararse en ciertos entornos (antivirus/proxy corporativo
    interceptando la conexion, particularidades de Windows, etc), y
    ahi el programa se queda "colgado" de verdad SIN que el timeout del
    socket lo salve -- que es justo lo que se vio en un log real: la
    conexion se establecio bien, pero la primera busqueda se quedo
    pegada sin ningun error ni progreso, mucho mas alla de los 30s.

    Este limite es independiente de eso: corre 'funcion' en un hilo
    aparte y espera COMO MUCHO TIMEOUT_CORREO_SEGUNDOS segundos. Si no
    termino a tiempo, fuerza el cierre del socket (para intentar
    liberar el hilo bloqueado) y sigue de largo sin esperarlo mas --
    se usa un hilo DAEMON a proposito: si el cierre del socket NO
    alcanza a desbloquearlo (puede pasar), un hilo daemon NUNCA
    impide que el programa termine, a diferencia de otros mecanismos
    (ej. concurrent.futures.ThreadPoolExecutor) que sí esperan a sus
    hilos internos al salir.

    Relanza cualquier excepcion de 'funcion' tal cual; si se agota el
    tiempo, levanta TimeoutError (subclase de OSError -- el mismo
    except que ya maneja la reconexion en buscar_correo_por_procesos
    lo atrapa sin necesitar un caso aparte).
    """
    resultado = {}

    def _ejecutar():
        try:
            resultado["valor"] = funcion(*args, **kwargs)
        except BaseException as error:
            resultado["error"] = error

    hilo = threading.Thread(target=_ejecutar, daemon=True)
    hilo.start()
    hilo.join(timeout=TIMEOUT_CORREO_SEGUNDOS)
    if hilo.is_alive():
        try:
            sock = getattr(mail, "sock", None)
            if sock is not None:
                sock.close()
        except Exception:
            pass
        raise TimeoutError(
            f"Gmail no respondio en {TIMEOUT_CORREO_SEGUNDOS}s (limite duro, no fue el timeout normal del socket)"
        )
    if "error" in resultado:
        raise resultado["error"]
    return resultado.get("valor")


def _buscar_y_leer_un_lote(mail, consulta, vistos):
    """
    Busca UN lote (SEARCH X-GM-RAW) y lee (FETCH) cada correo que
    encuentre, agregandolo a 'vistos' -- separado en su propia funcion
    para poder correrlo dentro de _con_limite_de_tiempo_duro (ver
    buscar_correo_por_procesos).
    """
    typ, datos = buscador.buscar_x_gm_raw(mail, consulta)
    if typ == "OK" and datos and datos[0]:
        for id_correo in datos[0].split():
            correo = base._leer_correo(mail, id_correo)
            if correo is not None:
                vistos[(correo["asunto"], correo["fecha"])] = correo


def buscar_correo_por_procesos(usuario, app_password, con_radicado):
    """
    UNA sola conexion IMAP para TODOS los procesos activos -- busca,
    en lotes combinados con OR (ver TAMANO_LOTE_CORREO), los demandados
    + radicados + cuentas de _terminos_de_busqueda. Devuelve la lista
    de correos encontrados, sin duplicados (por asunto+fecha) -- a cual
    proceso corresponde cada uno se decide despues, con la misma regla
    de coincidencia que el resto del proyecto (ver
    clasificar_procesos_ejecutivos._procesos_que_coinciden_con_correo).

    Con cientos de procesos activos son CIENTOS de busquedas seguidas
    sobre la MISMA conexion -- Gmail la corta si la nota con demasiadas
    busquedas/descargas seguidas en poco tiempo (proteccion propia de
    Gmail, no depende de nada configurable aqui). Cuando eso pasa, en
    vez de rendirse a mitad de camino, se RECONECTA (login + seleccionar
    la carpeta otra vez) y se reintenta el MISMO lote que se estaba
    procesando -- asi la busqueda completa los 100% de los terminos en
    vez de quedarse solo con los que alcanzo a revisar antes del primer
    corte (ver MAX_RECONEXIONES_CORREO para el limite de reintentos).
    """
    terminos = _terminos_de_busqueda(con_radicado)
    logging.info("[Correo] %d termino(s) distinto(s) para buscar (demandados + radicados + cuentas).", len(terminos))

    vistos = {}
    # OJO: NO se usa "with imaplib.IMAP4_SSL(...) as mail:" -- si Gmail
    # corta la conexion, el LOGOUT implicito del "with" al salir
    # revienta con un error de socket, y esa excepcion REEMPLAZA
    # cualquier "return" que hubiera adentro del bloque -- se perdian
    # TODOS los correos ya encontrados en los lotes que si funcionaron.
    # Con try/finally, el LOGOUT se intenta igual pero si falla se
    # ignora, y lo que ya se encontro en 'vistos' siempre se devuelve.
    mail = _conectar_gmail(usuario, app_password)
    if mail is None:
        logging.error("[Correo] Se omite la busqueda en correo.")
        return []

    total_lotes = -(-len(terminos) // TAMANO_LOTE_CORREO)  # division hacia arriba, sin importar math
    logging.info("[Correo] Buscando en %d lote(s) de hasta %d termino(s) cada uno...", total_lotes, TAMANO_LOTE_CORREO)

    reconexiones_usadas = 0
    try:
        for indice_lote, lote in enumerate(_lotes(terminos, TAMANO_LOTE_CORREO), start=1):
            # Aviso de progreso cada AVISO_PROGRESO_CORREO_LOTES lotes --
            # sin esto, una busqueda de cientos de lotes no deja NINGUNA
            # señal de vida en el log durante varios minutos seguidos
            # (nada falla, simplemente toma tiempo), y eso se ve
            # exactamente igual que "el programa esta colgado" aunque
            # este avanzando bien.
            if indice_lote > 1 and (indice_lote - 1) % AVISO_PROGRESO_CORREO_LOTES == 0:
                logging.info(
                    "[Correo] ...van %d/%d lote(s) revisados, %d correo(s) encontrados hasta el momento...",
                    indice_lote - 1, total_lotes, len(vistos),
                )
            # Sin tildes (ver buscador.texto_para_busqueda_gmail) --
            # Gmail busca igual sin distinguirlas, y asi alcanza un
            # quoted-string ASCII normal sin necesitar CHARSET ni
            # literales (ver el porque en esa funcion).
            consulta = "(" + " OR ".join(
                f'"{buscador.texto_para_busqueda_gmail(t).replace(chr(34), "")}"' for t in lote
            ) + ")"
            while True:
                try:
                    # 'consulta' ya viene sin tildes (texto_para_busqueda_gmail),
                    # asi que alcanza un SEARCH normal (ver
                    # buscador.buscar_x_gm_raw) sin CHARSET ni
                    # literales -- un termino con tilde/ñ demostro DOS
                    # problemas distintos con esos mecanismos: el
                    # quoted-string normal de IMAP revienta con BAD si
                    # de todas formas le llegan bytes no-ASCII, y el
                    # literal (que si acepta cualquier octeto por
                    # protocolo) se colgaba sin ningun error en ciertas
                    # redes/antivirus -- el intercambio "esperar el '+'
                    # de continuacion" que exige un literal es un
                    # patron de trafico que algunos proxies no manejan
                    # bien.
                    #
                    # Corre DENTRO de _con_limite_de_tiempo_duro -- el
                    # FETCH de cada correo encontrado va incluido (la
                    # conexion se puede caer ahi igual de facil, o mas:
                    # hay un FETCH por cada correo encontrado). Esto es
                    # una SEGUNDA linea de defensa ademas del timeout
                    # del propio socket: si ese timeout no llega a
                    # dispararse por algun motivo del entorno (ver
                    # _con_limite_de_tiempo_duro), este limite duro
                    # igual garantiza que no se quede colgado.
                    _con_limite_de_tiempo_duro(mail, _buscar_y_leer_un_lote, mail, consulta, vistos)
                    break  # lote resuelto (con o sin resultados) -- sigue al siguiente
                except (imaplib.IMAP4.abort, OSError) as error:
                    # Error de CONEXION (no de un comando puntual) --
                    # Gmail cierra la conexion si la nota inactiva o
                    # con demasiadas busquedas/descargas seguidas.
                    # OSError (incluye socket.timeout) tambien cuenta:
                    # con TIMEOUT_CORREO_SEGUNDOS puesto en el socket,
                    # una conexion que quedo "a medias" revienta con
                    # esto en vez de colgarse para siempre esperando
                    # una respuesta que nunca llega.
                    if reconexiones_usadas >= MAX_RECONEXIONES_CORREO:
                        logging.error(
                            "[Correo] Se perdio la conexion con Gmail (%s) y ya se intento reconectar %d vez/veces "
                            "sin exito -- se detiene aqui, ya se guardaron los %d correo(s) encontrados hasta el "
                            "momento.", error, reconexiones_usadas, len(vistos),
                        )
                        return list(vistos.values())
                    reconexiones_usadas += 1
                    logging.warning(
                        "[Correo] Se perdio la conexion con Gmail a mitad de la busqueda (%s) -- reconectando "
                        "(intento %d/%d) y sigue donde se quedo (van %d correo(s) encontrados)...",
                        error, reconexiones_usadas, MAX_RECONEXIONES_CORREO, len(vistos),
                    )
                    _cerrar_sesion_gmail(mail)
                    mail = _conectar_gmail(usuario, app_password)
                    if mail is None:
                        logging.error(
                            "[Correo] No se pudo reconectar -- se detiene aqui, ya se guardaron los %d "
                            "correo(s) encontrados hasta el momento.", len(vistos),
                        )
                        return list(vistos.values())
                    # vuelve a intentar EL MISMO lote con la conexion nueva
                except Exception as error:
                    # Cualquier otro error de UN lote puntual (ej. BAD
                    # de un solo comando) se salta y sigue con el
                    # siguiente -- no tumba la busqueda completa.
                    logging.error("[Correo] Fallo buscando el lote %s: %s", lote, error)
                    break
    finally:
        _cerrar_sesion_gmail(mail)

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
        procesos_coincidentes = _desambiguar_por_radicado(contenido_norm, procesos_coincidentes, contexto=correo["asunto"])
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
    con_radicado, terminados, _sin_estado = base.clasificar_procesos(procesos)
    logging.info(
        "Excel: %d proceso(s) activo/suspendido/reorganizacion/remitida/etc, %d terminado(s) para cruzar.",
        len(con_radicado), len(terminados),
    )

    carpeta_raiz = Path(base.CARPETA_PROCESOS)
    if not carpeta_raiz.exists():
        logging.error("No existe la carpeta configurada en CARPETA_PROCESOS: %s", carpeta_raiz)
        return

    # Paso 1 (carpeta de descargas) NO tiene restriccion de tipo -- vale
    # tanto para informacion extraprocesal de un proceso activo como
    # para el auto que termina un proceso YA terminado (ej. un
    # "desistimiento tacito"), asi que busca entre TODOS los procesos,
    # activos y terminados. El correo (paso 2) SI sigue restringido a
    # activos, porque ahi solo se descarga tutela/peticion/pago
    # oficioso -- eso no aplica a procesos ya terminados.
    indices_descargas = base._indexar_procesos_para_correo(con_radicado + terminados)

    logging.info("Paso 1: clasificando los archivos de %s...", CARPETA_DESCARGAS_MANUAL)
    movidos, sin_coincidencia, ambiguos = clasificar_carpeta_descargas(indices_descargas, carpeta_raiz)
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
