"""
Prueba el FILTRO de buscar_empleo.py sin tocar internet.

Le pasa ofertas de mentiras (escritas como se publican de verdad en
Colombia) a la misma funcion que decide a que postular, y verifica que
decida lo correcto. Sirve para dos cosas:

  1. Comprobar que el programa quedo bien instalado, sin gastar una
     sola postulacion real ni abrir el navegador.
  2. Ver COMO piensa el filtro, para que sepas que esperar cuando lo
     dejes corriendo de verdad.

    python probar_filtro_empleo.py
"""

import sys

import buscar_empleo as app

PERFIL_DE_PRUEBA = {
    "nombre_completo": "Persona de Prueba",
    "correo": "prueba@ejemplo.com",
    "telefono": "300 000 0000",
    "ciudad": "Bucaramanga",
    "ruta_hoja_de_vida": "",
    "cargos_objetivo": ["auxiliar juridico", "analista de cartera"],
    "palabras_clave_deseables": ["procesos ejecutivos", "excel", "rama judicial"],
    "palabras_excluyentes": ["solo comision", "multinivel"],
    "ciudades": ["Bucaramanga", "Floridablanca"],
    "acepta_remoto": True,
    "anos_experiencia": 3,
    "nivel_educativo": "profesional",
    "salario_minimo": 1600000,
    "resumen_profesional": "Resumen de prueba.",
    "asunto_correo": "Postulacion: {titulo}",
    "carta_presentacion": "",
}

# (titulo, descripcion, que se espera: "postula" o "descarta")
CASOS = [
    (
        "Auxiliar Juridico",
        "Empresa en Bucaramanga requiere auxiliar juridico con 2 anos de "
        "experiencia en procesos ejecutivos. Profesional en derecho. Manejo "
        "de Excel. Salario $2.200.000. Publicado hace 1 dia.",
        "postula",
    ),
    (
        "Analista de Cartera",
        "Se requiere analista de cartera para Floridablanca, experiencia "
        "minima de 3 anos en cobro juridico. Titulo profesional. "
        "Salario a convenir. Enviar hoja de vida a talento@empresa.com",
        "postula",
    ),
    (
        "Auxiliar Juridico Senior",
        "Bucaramanga. Se requiere profesional con 8 anos de experiencia en "
        "litigio. Salario $5.000.000",
        "descarta",  # pide 8 anos y el perfil tiene 3
    ),
    (
        "Asesor Comercial",
        "Bucaramanga. Vendedor con pago solo comision, sin salario base. "
        "Sin experiencia.",
        "descarta",  # palabra excluyente "solo comision"
    ),
    (
        "Auxiliar Juridico",
        "Vacante en Leticia, Amazonas. Profesional en derecho, 1 ano de "
        "experiencia. Presencial.",
        "descarta",  # ciudad fuera de la lista y no es remota
    ),
    (
        "Auxiliar Juridico",
        "Bucaramanga. Profesional en derecho con 2 anos de experiencia. "
        "Salario $1.100.000 mensuales.",
        "descarta",  # paga menos del minimo del perfil
    ),
    (
        "Auxiliar Juridico Remoto",
        "Trabajo remoto desde cualquier ciudad de Colombia. Profesional en "
        "derecho, 1 ano de experiencia en procesos ejecutivos y rama "
        "judicial. Salario $2.000.000",
        "postula",  # otra ciudad, pero es remota y el perfil lo acepta
    ),
    (
        "Auxiliar Juridico",
        "Bucaramanga. Se requiere candidato con maestria en derecho "
        "procesal y 2 anos de experiencia. Salario $3.000.000",
        "descarta",  # pide maestria y el perfil es profesional
    ),
    (
        "Operario de Produccion",
        "Bucaramanga. Bachiller, sin experiencia. Salario $1.800.000",
        "descarta",  # no es un cargo objetivo: no llega al umbral
    ),
]


def probar_lectura_de_requisitos():
    """Las funciones que leen la letra menuda de cada oferta."""
    pruebas = [
        (app.experiencia_pedida, "se requiere 2 anos de experiencia", 2),
        (app.experiencia_pedida, "experiencia minima de 5 anos en el cargo", 5),
        (app.experiencia_pedida, "de 1 a 3 anos de experiencia", 1),
        (app.experiencia_pedida, "18 meses de experiencia certificada", 1),
        (app.experiencia_pedida, "no requiere experiencia", 0),
        (app.experiencia_pedida, "buen ambiente laboral", None),
        (app.salario_ofrecido, "salario $ 2.500.000 mas prestaciones", 2500000),
        (app.salario_ofrecido, "devengaras 3 millones mensuales", 3000000),
        (app.salario_ofrecido, "salario a convenir", None),
        (app.antiguedad_en_dias, "publicado hace 3 dias", 3),
        (app.antiguedad_en_dias, "publicado ayer", 1),
        (app.antiguedad_en_dias, "hace 2 semanas", 14),
    ]
    fallos = 0
    print("LECTURA DE REQUISITOS")
    print("-" * 62)
    for funcion, texto, esperado in pruebas:
        obtenido = funcion(app.sin_tildes(texto))
        ok = obtenido == esperado
        fallos += not ok
        print(f"  [{'ok ' if ok else 'MAL'}] {funcion.__name__}({texto!r})")
        if not ok:
            print(f"         esperaba {esperado!r}, dio {obtenido!r}")
    return fallos


def probar_correos():
    print()
    print("CORREOS DE CONTACTO (se ignoran los buzones que no reciben hojas de vida)")
    print("-" * 62)
    texto = ("Enviar hoja de vida a seleccion@empresa.com.co o al correo "
             "rrhh@empresa.com. No responder a noreply@computrabajo.com")
    encontrados = app.correos_en_texto(texto)
    esperado = ["seleccion@empresa.com.co", "rrhh@empresa.com"]
    ok = encontrados == esperado
    print(f"  [{'ok ' if ok else 'MAL'}] encontrados: {encontrados}")
    return 0 if ok else 1


def probar_decisiones():
    print()
    print("DECISION SOBRE CADA OFERTA")
    print("-" * 62)
    perfil = dict(app.PERFIL_POR_DEFECTO)
    perfil.update(PERFIL_DE_PRUEBA)
    perfil["_cargos"] = [app.sin_tildes(c) for c in perfil["cargos_objetivo"]]
    perfil["_deseables"] = [app.sin_tildes(c) for c in perfil["palabras_clave_deseables"]]
    perfil["_excluyentes"] = [app.sin_tildes(c) for c in perfil["palabras_excluyentes"]]
    perfil["_ciudades"] = [app.sin_tildes(c) for c in perfil["ciudades"]]
    perfil["_nivel"] = app.nivel_a_numero(perfil["nivel_educativo"])
    perfil["_ruta_hoja_de_vida"] = None

    fallos = 0
    for titulo, descripcion, esperado in CASOS:
        vac = app.Vacante(portal="prueba", titulo=titulo,
                          url="https://ejemplo.com/oferta/" + app.a_slug(titulo))
        vac.descripcion = descripcion
        vac.correos = app.correos_en_texto(descripcion)

        ev = app.evaluar_vacante(vac, perfil)
        postularia = (not ev.descartada) and ev.puntaje >= app.UMBRAL_POSTULACION
        obtenido = "postula" if postularia else "descarta"
        ok = obtenido == esperado
        fallos += not ok

        print(f"  [{'ok ' if ok else 'MAL'}] {titulo[:38]:<38} -> {obtenido.upper()}")
        print(f"         {ev.resumen()}")
        if not ok:
            print(f"         SE ESPERABA: {esperado.upper()}")
    return fallos


def main():
    print()
    print("=" * 62)
    print(f" Probando el filtro de empleo (umbral actual: "
          f"{app.UMBRAL_POSTULACION} puntos)")
    print("=" * 62)

    fallos = probar_lectura_de_requisitos() + probar_correos() + probar_decisiones()

    print()
    print("=" * 62)
    if fallos:
        print(f" {fallos} prueba(s) FALLARON.")
    else:
        print(" Todas las pruebas pasaron. El filtro esta funcionando bien.")
    print("=" * 62)
    return 1 if fallos else 0


if __name__ == "__main__":
    sys.exit(main())
