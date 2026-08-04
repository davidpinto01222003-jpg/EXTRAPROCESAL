"""
Limpia y termina de nombrar las carpetas de CARPETA_PROCESOS (la misma
carpeta que usa clasificar_procesos_ejecutivos.py) en dos pasos
independientes, ambos respetando MODO_PRUEBA:

1. CONSOLIDA, por número de proceso, todas las carpetas "de trabajo"
   que haya en el disco para ese número (ej. una vieja "1055. ACTIVO"
   de cuando todavía no tenía radicado, y una nueva "1055. <radicado>"
   ya con el radicado -- o una vieja "1051. ACTIVO" y una nueva "1051.
   TERMINADO POR PAGO" porque el proceso cambió de estado) contra el
   nombre que le corresponde HOY según el Excel -- "<numero>. <radicado>"
   si ya tiene radicado, o "<numero>. <ESTADO PROCESAL>" si todavía no
   (misma regla que usa clasificar_procesos_ejecutivos.clasificar_procesos,
   para activos y para terminados por igual). Esto cubre tanto las
   carpetas que quedaron con SOLO el número (ej. "245." o "245") como
   las que quedaron con un nombre VIEJO que ya no corresponde al estado
   o radicado actual del proceso:
     - Si la carpeta con el nombre correcto YA EXISTE en el disco: la
       carpeta vieja/suelta se BORRA si está completamente vacía (ya
       quedó reemplazada por la correcta), o se reporta para revisión
       manual si todavía tiene contenido adentro (no se combina solo,
       para no arriesgar mezclar archivos de dos carpetas a ciegas).
     - Si la carpeta con el nombre correcto NO existe todavía: la
       carpeta vieja/suelta se RENOMBRA directo al nombre correcto
       (tenga o no contenido, no se pierde nada).
   Si el número es ambiguo (aparece en más de una carpeta de trabajo
   POSIBLE según el Excel -- ej. el mismo número con radicados
   distintos), o si hay más de una carpeta vieja candidata sin ninguna
   forma segura de saber cuál corresponde, se omite y se reporta para
   revisar a mano.

2. BORRA las carpetas de procesos ACTIVOS -- es decir, todo lo que NO
   es "terminado"/"no inicio" (activo, suspendido, en reorganización,
   remitida a castigo/prepago, etc, igual que la regla 1 del módulo
   clasificar_procesos_ejecutivos) -- que estén COMPLETAMENTE VACÍAS
   (sin ningún archivo adentro, ni siquiera en subcarpetas -- no cuenta
   basura que crea Windows solo, como desktop.ini/Thumbs.db). Los
   procesos TERMINADOS/NO INICIO NUNCA se tocan aquí, estén vacíos o
   no -- para esos ya existe terminados_sin_auto_pendientes.csv, que es
   la lista correcta de qué hace falta descargar a mano.

ADVERTENCIA: borrar una carpeta es IRREVERSIBLE. Respeta MODO_PRUEBA
(por defecto True): en modo prueba solo revisa y muestra qué haría, sin
tocar ninguna carpeta todavía.

Reutiliza toda la lectura del Excel de clasificar_procesos_ejecutivos.py
(misma hoja ACTIVOS, mismo CARPETA_PROCESOS, mismas reglas de nombre) --
no hay nada que configurar aparte.
"""

import logging
import os
import re
import shutil
from pathlib import Path

import clasificar_procesos_ejecutivos as base
import validar_renombrar_carpetas as cruce_excel

# ============================= CONFIGURACION =============================

ARCHIVO_LOG = os.path.join(os.path.dirname(__file__), "limpiar_carpetas_procesos_ejecutivos.log")

# True (por defecto): no renombra ni borra nada de verdad, solo revisa y
# muestra qué haría. False: aplica los cambios de verdad (irreversible
# para el borrado).
MODO_PRUEBA = True

# ===========================================================================

# Reconoce CUALQUIER carpeta "de trabajo" de un proceso por su número
# inicial -- "245", "245.", "245. ACTIVO", "245. <radicado>", etc (el
# numero de proceso nunca pasa de unos pocos miles, por eso se limita a
# 1-5 digitos -- asi no se confunde por error una carpeta que sea SOLO
# un radicado de 23 digitos, sin el "numero. " por delante, con un
# numero de proceso).
PATRON_CARPETA_DE_PROCESO = re.compile(r"^(\d{1,5})(?:\.(?:\s+(.*))?)?$")


def configurar_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
        handlers=[
            logging.FileHandler(ARCHIVO_LOG, encoding="utf-8", mode="w"),
            logging.StreamHandler(),
        ],
    )


def _nombre_correcto_por_numero(con_radicado, terminados):
    """
    {numero: nombre_carpeta} -- el nombre de carpeta que le corresponde
    HOY a cada número de proceso según el Excel. Si un número aparece
    con más de una carpeta de trabajo posible (ej. el mismo número con
    radicados distintos), se excluye de este mapa -- queda ambiguo, se
    reporta aparte en vez de adivinar. Devuelve (mapa, ambiguos).
    """
    por_numero = {}
    ambiguos = set()
    for proceso in con_radicado + terminados:
        numero = proceso["numero"]
        if numero in por_numero and por_numero[numero] != proceso["nombre_carpeta"]:
            ambiguos.add(numero)
        else:
            por_numero[numero] = proceso["nombre_carpeta"]
    for numero in ambiguos:
        por_numero.pop(numero, None)
    return por_numero, ambiguos


def _agrupar_carpetas_por_numero(carpeta_raiz):
    """{numero: [carpeta, ...]} -- TODAS las carpetas de primer nivel que reconoce PATRON_CARPETA_DE_PROCESO, agrupadas por su numero de proceso."""
    grupos = {}
    try:
        carpetas = sorted(carpeta_raiz.iterdir(), key=lambda p: p.name)
    except OSError as error:
        logging.error("[Disco] No se pudo leer %s: %s", carpeta_raiz, error)
        return grupos

    for carpeta in carpetas:
        if not carpeta.is_dir() or carpeta.name in cruce_excel.CARPETAS_A_IGNORAR:
            continue
        coincide = PATRON_CARPETA_DE_PROCESO.match(carpeta.name)
        if not coincide:
            continue
        numero = int(coincide.group(1))
        grupos.setdefault(numero, []).append(carpeta)
    return grupos


def consolidar_carpetas_por_numero(carpeta_raiz, nombre_correcto_por_numero, ambiguos):
    """
    Para cada numero de proceso con nombre correcto conocido: borra las
    carpetas VIEJAS/sueltas que ya quedaron reemplazadas por la correcta
    (si estan vacias), renombra la unica carpeta vieja a la correcta
    (si la correcta todavia no existe), y reporta cualquier caso
    ambiguo o con contenido para revision manual -- ver docstring del
    modulo. Devuelve (renombradas, borradas, reportadas).
    """
    grupos = _agrupar_carpetas_por_numero(carpeta_raiz)

    renombradas = 0
    borradas = 0
    reportadas = 0

    for numero, carpetas in sorted(grupos.items()):
        if numero in ambiguos:
            logging.warning(
                "[Ambiguo] Proceso %s: no se toca ninguna de sus carpetas (%s) -- tiene mas de una carpeta "
                "de trabajo posible en el Excel (radicados distintos), revisalo a mano.",
                numero, ", ".join(c.name for c in carpetas),
            )
            reportadas += 1
            continue

        nombre_correcto = nombre_correcto_por_numero.get(numero)
        if not nombre_correcto:
            logging.warning(
                "[Sin dato] Proceso %s: no se toca su carpeta (%s) -- no aparece en el Excel (o todavia no "
                "tiene ESTADO PROCESAL diligenciado).", numero, ", ".join(c.name for c in carpetas),
            )
            reportadas += 1
            continue

        correcta = next((c for c in carpetas if c.name == nombre_correcto), None)
        viejas = [c for c in carpetas if c is not correcta]

        if correcta is not None:
            for vieja in viejas:
                if cruce_excel.contar_archivos(vieja) != 0:
                    logging.warning(
                        "[Conflicto] Proceso %s: '%s' quedo vieja (la carpeta correcta ya es '%s'), pero "
                        "todavia tiene contenido adentro -- se omite, revisala a mano.",
                        numero, vieja.name, nombre_correcto,
                    )
                    reportadas += 1
                    continue
                if MODO_PRUEBA:
                    logging.info(
                        "[SIMULACION] Se borraria '%s' (proceso %s -- vieja y vacia, ya reemplazada por '%s').",
                        vieja.name, numero, nombre_correcto,
                    )
                    borradas += 1
                    continue
                try:
                    shutil.rmtree(base.organizador._ruta_larga_segura(str(vieja)))
                    borradas += 1
                    logging.info(
                        "[Borrada] '%s' (proceso %s -- vieja y vacia, ya reemplazada por '%s').",
                        vieja.name, numero, nombre_correcto,
                    )
                except OSError as error:
                    logging.warning("   (no se pudo borrar '%s': %s)", vieja.name, error)
            continue

        if len(viejas) > 1:
            logging.warning(
                "[Ambiguo] Proceso %s: hay %d carpetas sueltas (%s) y ninguna se llama todavia '%s' -- no se "
                "sabe cual renombrar, revisalo a mano.",
                numero, len(viejas), ", ".join(c.name for c in viejas), nombre_correcto,
            )
            reportadas += 1
            continue

        vieja = viejas[0]
        if MODO_PRUEBA:
            logging.info("[SIMULACION] '%s'  ->  '%s'", vieja.name, nombre_correcto)
        else:
            vieja.rename(carpeta_raiz / nombre_correcto)
            logging.info("[Renombrada] '%s'  ->  '%s'", vieja.name, nombre_correcto)
        renombradas += 1

    return renombradas, borradas, reportadas


def borrar_carpetas_activas_vacias(carpeta_raiz, con_radicado):
    borradas = 0
    no_vacias = 0
    for proceso in con_radicado:
        carpeta = carpeta_raiz / proceso["nombre_carpeta"]
        if not carpeta.exists() or not carpeta.is_dir():
            continue
        if cruce_excel.contar_archivos(carpeta) != 0:
            no_vacias += 1
            continue

        if MODO_PRUEBA:
            logging.info(
                "[SIMULACION] Se borraria '%s' (proceso %s, %s -- vacia, sin ningun documento adentro).",
                carpeta.name, proceso["numero"], proceso["estado"],
            )
            borradas += 1
            continue

        try:
            shutil.rmtree(base.organizador._ruta_larga_segura(str(carpeta)))
            borradas += 1
            logging.info(
                "[Borrada] '%s' (proceso %s, %s -- vacia, sin ningun documento adentro).",
                carpeta.name, proceso["numero"], proceso["estado"],
            )
        except OSError as error:
            logging.warning("   (no se pudo borrar '%s': %s)", carpeta.name, error)
    return borradas, no_vacias


def procesar():
    if not base.RUTA_EXCEL_CONTROL or not os.path.exists(base.RUTA_EXCEL_CONTROL):
        logging.error("No se encontro el Excel configurado en RUTA_EXCEL_CONTROL: %r", base.RUTA_EXCEL_CONTROL)
        return

    procesos = base.leer_procesos_control()
    con_radicado, terminados, _sin_estado = base.clasificar_procesos(procesos)
    logging.info(
        "Excel: %d proceso(s) activo/suspendido/reorganizacion/remitida/etc, %d terminado(s) (pago/auto/"
        "contrato/no inicio).", len(con_radicado), len(terminados),
    )

    carpeta_raiz = Path(base.CARPETA_PROCESOS)
    if not carpeta_raiz.exists():
        logging.error("No existe la carpeta configurada en CARPETA_PROCESOS: %s", carpeta_raiz)
        return

    logging.info("Paso 1: consolidando carpetas sueltas/viejas contra el nombre correcto de cada proceso...")
    nombre_correcto_por_numero, ambiguos = _nombre_correcto_por_numero(con_radicado, terminados)
    renombradas, borradas_viejas, reportadas = consolidar_carpetas_por_numero(carpeta_raiz, nombre_correcto_por_numero, ambiguos)
    logging.info(
        "Paso 1: %d carpeta(s) %s, %d carpeta(s) vieja(s) vacia(s) %s, %d caso(s) reportado(s) para revisar a mano.",
        renombradas, "simuladas para renombrar (MODO_PRUEBA activo)" if MODO_PRUEBA else "renombradas",
        borradas_viejas, "simuladas para borrar (MODO_PRUEBA activo)" if MODO_PRUEBA else "borradas",
        reportadas,
    )

    logging.info(
        "Paso 2: borrando carpetas VACIAS de procesos activos/suspendidos/reorganizacion/remitida/etc "
        "(nunca terminados/no inicio)..."
    )
    borradas, no_vacias = borrar_carpetas_activas_vacias(carpeta_raiz, con_radicado)
    logging.info(
        "Paso 2: %d carpeta(s) %s; %d proceso(s) activo(s) tienen carpeta pero YA tienen contenido adentro "
        "(no se tocaron).", borradas,
        "simuladas para borrar (MODO_PRUEBA activo)" if MODO_PRUEBA else "borradas PERMANENTEMENTE",
        no_vacias,
    )

    if MODO_PRUEBA:
        logging.info(
            "MODO_PRUEBA esta activo: no se renombro ni se borro nada todavia. Revisa el log y, si se ve "
            "bien, cambia MODO_PRUEBA = False al inicio de este script y vuelve a correrlo."
        )


def main():
    configurar_logging()
    procesar()


if __name__ == "__main__":
    main()
