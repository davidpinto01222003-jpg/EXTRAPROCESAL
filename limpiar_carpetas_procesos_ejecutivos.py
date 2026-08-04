"""
Limpia y termina de nombrar las carpetas de CARPETA_PROCESOS (la misma
carpeta que usa clasificar_procesos_ejecutivos.py) en dos pasos
independientes, ambos respetando MODO_PRUEBA:

1. RENOMBRA las carpetas que quedaron con SOLO el número de proceso
   como nombre (ej. "245." o "245", sin radicado ni ESTADO PROCESAL) al
   nombre que les corresponde HOY según el Excel -- "<numero>.
   <radicado>" si ya tiene radicado, o "<numero>. <ESTADO PROCESAL>" si
   todavía no (exactamente la misma regla que usa
   clasificar_procesos_ejecutivos.clasificar_procesos, para activos y
   para terminados por igual). Si el número es ambiguo (aparece en más
   de una carpeta de trabajo distinta -- ej. el mismo número con
   radicados diferentes en el Excel), se omite y se reporta para
   revisar a mano: no hay forma segura de saber a cuál le corresponde.

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

# Una carpeta "solo con el número" -- ej. "245." o "245" (con o sin el
# punto, con o sin espacios al final), sin radicado ni ESTADO PROCESAL.
PATRON_SOLO_NUMERO = re.compile(r"^(\d+)\.?\s*$")


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


def renombrar_carpetas_solo_numero(carpeta_raiz, nombre_correcto_por_numero, ambiguos):
    renombradas = 0
    try:
        carpetas = sorted(carpeta_raiz.iterdir(), key=lambda p: p.name)
    except OSError as error:
        logging.error("[Disco] No se pudo leer %s: %s", carpeta_raiz, error)
        return renombradas

    for carpeta in carpetas:
        if not carpeta.is_dir() or carpeta.name in cruce_excel.CARPETAS_A_IGNORAR:
            continue
        coincide = PATRON_SOLO_NUMERO.match(carpeta.name)
        if not coincide:
            continue
        numero = int(coincide.group(1))

        if numero in ambiguos:
            logging.warning(
                "[Ambiguo] '%s' no se renombra: el proceso %s tiene mas de una carpeta de trabajo posible "
                "en el Excel (radicados distintos) -- revisalo a mano.",
                carpeta.name, numero,
            )
            continue

        nombre_correcto = nombre_correcto_por_numero.get(numero)
        if not nombre_correcto:
            logging.warning(
                "[Sin dato] '%s' no se renombra: el proceso %s no aparece en el Excel (o todavia no tiene "
                "ESTADO PROCESAL diligenciado).", carpeta.name, numero,
            )
            continue

        if nombre_correcto == carpeta.name:
            continue

        destino = carpeta_raiz / nombre_correcto
        if destino.exists():
            logging.warning(
                "[Conflicto] '%s' deberia renombrarse a '%s' pero esa carpeta ya existe -- se omite, revisa "
                "a mano (puede que haya contenido repartido entre las dos).", carpeta.name, nombre_correcto,
            )
            continue

        if MODO_PRUEBA:
            logging.info("[SIMULACION] '%s'  ->  '%s'", carpeta.name, nombre_correcto)
        else:
            carpeta.rename(destino)
            logging.info("[Renombrada] '%s'  ->  '%s'", carpeta.name, nombre_correcto)
        renombradas += 1
    return renombradas


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

    logging.info("Paso 1: renombrando carpetas que quedaron solo con el numero...")
    nombre_correcto_por_numero, ambiguos = _nombre_correcto_por_numero(con_radicado, terminados)
    renombradas = renombrar_carpetas_solo_numero(carpeta_raiz, nombre_correcto_por_numero, ambiguos)
    logging.info(
        "Paso 1: %d carpeta(s) %s.", renombradas,
        "simuladas para renombrar (MODO_PRUEBA activo)" if MODO_PRUEBA else "renombradas",
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
