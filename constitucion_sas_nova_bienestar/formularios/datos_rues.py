# -*- coding: utf-8 -*-
"""Valores del caso NOVA BIENESTAR INTEGRAL S.A.S. para la carátula del RUES."""

TEL = "3156849271"
CORREO = "CONTACTO@NOVABIENESTARINTEGRAL.COM"
DIR = "CARRERA 37 N° 10A-58 OFICINA 804"


def digitos(prefijo, inicio, cantidad, texto):
    """Reparte `texto` carácter por carácter en casillas numeradas."""
    salida = {}
    for i, ch in enumerate(texto):
        n = inicio + i
        salida[prefijo if n == 0 else "%s_%d" % (prefijo, n + 1)] = ch
    return salida


def caratula():
    v = {}
    # --- Hoja 1, sección 2: identificación
    v["Cuadro de texto 2_42"] = "NOVA BIENESTAR INTEGRAL S.A.S."
    # --- Hoja 1, sección 3: ubicación y datos generales
    v["Cuadro de texto 2_76"] = DIR
    v["Cuadro de texto 2_79"] = "X"                     # ubicación: OFICINA
    v["Cuadro de texto 2_84"] = "MEDELLÍN"
    v.update(digitos("Cuadro de texto 2", 87, 3, "001"))   # cód. municipio
    v["Cuadro de texto 2_85"] = "ANTIOQUIA"
    v.update(digitos("Cuadro de texto 2", 90, 2, "05"))    # cód. departamento
    v["Cuadro de texto 2_87"] = "COLOMBIA"
    v.update(digitos("Cuadro de texto 2", 92, 10, TEL))    # teléfono 1
    v["Cuadro de texto 2_206"] = CORREO
    # notificación judicial
    v["Cuadro de texto 2_123"] = DIR
    v["Cuadro de texto 2_125"] = "MEDELLÍN"
    v.update(digitos("Cuadro de texto 2", 128, 3, "001"))
    v["Cuadro de texto 2_126"] = "ANTIOQUIA"
    v.update(digitos("Cuadro de texto 2", 131, 2, "05"))
    v["Cuadro de texto 2_128"] = "COLOMBIA"
    v.update(digitos("Cuadro de texto 2", 133, 10, TEL))
    v["Cuadro de texto 2_207"] = CORREO
    # --- Hoja 1, sección 4: actividades económicas
    v.update(digitos("Cuadro de texto 2", 163, 4, "8692"))
    v.update(digitos("Cuadro de texto 2", 168, 4, "9311"))
    v.update(digitos("Cuadro de texto 2", 173, 4, "9609"))
    v.update(digitos("Cuadro de texto 2", 183, 8, "01102026"))
    v.update(digitos("Cuadro de texto 2", 191, 8, "01102026"))
    # --- Hoja 2, sección 5: información financiera (balance de apertura)
    v["Campo num#C3#A9rico 1"] = "600000000"      # activo corriente
    v["Campo num#C3#A9rico 1_2"] = "0"            # activo no corriente
    v["Campo num#C3#A9rico 2"] = "600000000"      # activo total
    for k in ("1_3", "1_4", "1_5"):               # pasivos
        v["Campo num#C3#A9rico " + k] = "0"
    v["Campo num#C3#A9rico 1_6"] = "600000000"    # patrimonio neto
    v["Campo num#C3#A9rico 1_7"] = "600000000"    # pasivo + patrimonio
    for i in range(9, 17):                        # estado de resultados
        v["Campo num#C3#A9rico 1_%d" % i] = "0"
    v["Cuadro de texto 1_2"] = "2"                # grupo NIIF
    v["Campo num#C3#A9rico 1_17"] = "0"           # nacional público
    v["Campo num#C3#A9rico 1_18"] = "100"         # nacional privado
    v["Campo num#C3#A9rico 1_19"] = "0"           # extranjero público
    v["Campo num#C3#A9rico 1_20"] = "0"           # extranjero privado
    v["Campo num#C3#A9rico 1_21"] = "65"          # % mujeres en el capital
    # --- Hoja 2, sección 8: estado actual
    v["Cuadro de texto 1_9"] = "9"                # número de empleados
    v["Cuadro de texto 1_14"] = "2"               # mujeres en cargos directivos
    v["Cuadro de texto 1_19"] = "5"               # empleadas mujeres
    v["Cuadro de texto 1_20"] = "1"               # cuántos establecimientos
    v["Campo num#C3#A9rico 3"] = "0"              # % empleados temporales
    # --- Hoja 2: firma
    v["Cuadro de texto 3"] = "MARIANA LÓPEZ CASTAÑO"
    v["Cuadro de texto 3_4"] = "1.039.672.481"
    return v


CASILLAS = {
    "Casilla 1_8": "/Yes",    # matrícula / inscripción en registro mercantil
    "Casilla 1_20": "/Yes",   # zona urbana (domicilio)
    "Casilla 1_22": "/Yes",   # zona urbana (notificación)
    "Casilla 1_25": "/Yes",   # sede administrativa en arriendo
    "Casilla 1_28": "/Yes",   # autoriza notificación por correo electrónico
    "Casilla 1_33": "/Yes",   # tiene establecimientos: sí
    "Casilla 1_36": "/Yes",   # innovación: no
    "Casilla 1_38": "/Yes",   # empresa familiar: no
    "Casilla 1_39": "/Yes",   # Ley 1780: declara que cumple
    "Casilla 1_43": "/Yes",   # aportante a seguridad social: sí
    "Casilla 1_46": "/Yes",   # menos de 200 cotizantes
    "Casilla 1_49": "/Yes",   # tipo de documento: C.C.
}


def anexo1():
    """Anexo 1 — matrícula del establecimiento de comercio NOVA BIENESTAR."""
    v = {}
    # Datos del establecimiento
    v["Cuadro de texto 1"] = "NOVA BIENESTAR"
    v["Cuadro de texto 1_2"] = DIR
    v.update(digitos("Cuadro de texto 2", 20, 10, TEL))      # teléfono 1
    v["Cuadro de texto 1_4"] = "MEDELLÍN"
    v.update(digitos("Cuadro de texto 2", 50, 3, "001"))
    v["Cuadro de texto 1_5"] = "ANTIOQUIA"
    v.update(digitos("Cuadro de texto 2", 53, 2, "05"))
    v["Cuadro de texto 1_13"] = CORREO
    v["Campo num#C3#A9rico 1"] = "55000000"                  # activos vinculados
    v["Campo num#C3#A9rico 1_2"] = "9"                       # trabajadores
    # Actividad económica del establecimiento
    v.update(digitos("Cuadro de texto 2", 66, 4, "8692"))
    v.update(digitos("Cuadro de texto 2", 71, 4, "9311"))
    v.update(digitos("Cuadro de texto 2", 76, 4, "9609"))
    v["Cuadro de texto 3"] = ("PRESTACIÓN DE SERVICIOS DE BIENESTAR, "
                              "REHABILITACIÓN FÍSICA Y ACONDICIONAMIENTO INTEGRAL")
    # Propietario
    v["Cuadro de texto 1_28"] = "NOVA BIENESTAR INTEGRAL S.A.S."
    v["Cuadro de texto 1_7"] = DIR
    v["Cuadro de texto 1_12"] = "MEDELLÍN"
    v.update(digitos("Cuadro de texto 2", 98, 3, "001"))
    v["Cuadro de texto 1_14"] = "ANTIOQUIA"
    v.update(digitos("Cuadro de texto 2", 101, 2, "05"))
    v.update(digitos("Cuadro de texto 2", 103, 10, TEL))     # teléfono propietario
    v["Cuadro de texto 1_15"] = DIR
    v["Cuadro de texto 1_16"] = "MEDELLÍN"
    v.update(digitos("Cuadro de texto 2", 135, 3, "001"))
    v["Cuadro de texto 1_17"] = "ANTIOQUIA"
    v.update(digitos("Cuadro de texto 2", 133, 2, "05"))
    v["Cuadro de texto 1_18"] = "MARIANA LÓPEZ CASTAÑO"
    v.update(digitos("Cuadro de texto 2", 138, 10, "1039672481"))
    return v


CASILLAS_ANEXO1 = {
    "Casilla 1": "/Yes",      # establecimiento de comercio
    "Casilla 1_4": "/Yes",    # matrícula
    "Casilla 1_9": "/Yes",    # ubicación: oficina
    "Casilla 1_12": "/Yes",   # propietario único
    "Casilla 1_16": "/Yes",   # el local es ajeno (arrendado)
    "Casilla 1_19": "/Yes",   # tipo de identificación del propietario: NIT
    "Casilla 1_22": "/Yes",   # documento del representante legal: C.C.
}
