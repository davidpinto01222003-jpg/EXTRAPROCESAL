# Segunda sesión del webinar — envío de invitaciones

**Miércoles 14 de octubre, 8:00 a. m.**
Inscripciones: https://forms.gle/rheWoEX9D3X2NjYm6

---

## Dónde va la campaña

El **6 de octubre salieron 184 correos**, las instituciones más grandes
de la lista. El envío se cortó solo porque Google cierra las sesiones
SMTP largas; **no se agotó ninguna cuota**. Ese defecto ya está
corregido: ahora el programa reconecta y sigue.

**`contactos.xlsx` ya no trae esos 184.** Quedan **2.531**.

| Tramo | Correos | Día |
|---|---|---|
| A | 650 | miércoles 7 |
| B | 650 | jueves 8 |
| C | 650 | viernes 9 |
| — | — | lunes 12: **festivo** |
| D | 581 | martes 13 |

Ponga `maximo_por_dia` en 650 y corra la opción 3 una vez al día. El
programa lleva el registro y nunca repite un correo.

---

## Cuánto se demora, y cómo acortarlo

Esto es lo que más pesa ahora. Del envío del 6 de octubre salió el dato
real: **24 segundos entre un correo y el siguiente**. La pausa
configurada es de 10 a 20 s, o sea unos 15. **Los 9 segundos de más son
el brochure subiendo por la red.**

Con eso, los 2.531 que faltan son:

| | Horas en total | Por día |
|---|---|---|
| Como está ahora | **16,9 h** | 4,2 h |
| Sin adjunto, misma pausa | 11,3 h | 2,8 h |
| Sin adjunto, pausa 6–12 s | **7,0 h** | 1,8 h |

Son más de **cuatro horas diarias** con el computador prendido en el
primer caso, contra menos de dos en el último.

Si quiere acortarlo, en `config.ini`:

- Borre lo que va después de `adjunto =` — el correo baja de 2,15 MB a
  144 KB y sale el brochure. Quien se inscriba lo recibe después.
- Ponga `pausa_minima = 6` y `pausa_maxima = 12`.

El corte del 6 de octubre fue por antigüedad de la conexión, no por
ritmo, así que bajar las pausas no tiene riesgo demostrado.

Si prefiere mantener el brochure, también funciona. Solo cuesta el
doble de tiempo frente al computador.

---

### El orden importa, y no es alfabético

La base viene **ordenada por tamaño de la institución**: camas,
consultorios, salas de procedimientos, ambulancias y número de sedes.
El programa respeta ese orden exacto.

Eso significa que **el primer lote, el del martes, se lo lleva las 250
IPS más grandes** de los cinco departamentos. Mired Barranquilla,
Metrosalud, Pablo Tobón Uribe, Alma Máter, Comfama. Si la campaña
tuviera que pararse a mitad de camino por lo que sea, lo que ya salió
es lo que más vale.

Los tramos A, B, C, D y E que aparecen en la vista previa son
exactamente los lotes de cada día.

El archivo `contactos.xlsx` trae además el puesto en el ranking, el
departamento, el número de sedes y la capacidad de cada IPS, por si
quiere revisarlo o llamar a alguna directamente.

---

## Antes del primer envío

**1. No borre la carpeta `estado`.** Ahí está el registro de lo que ya
salió. El ZIP no la trae a propósito: si usted descomprime encima de la
carpeta que ya tiene, la suya se conserva intacta.

Si arranca en una carpeta nueva, copie de la campaña anterior:

```
campana_webinar\estado\excluidos.csv  →  invitacion_octubre\estado\excluidos.csv
```

Ese archivo tiene los 10 que rebotaron y los dominios muertos.

**No abra `estado\enviados.csv` con Excel.** Excel reescribe las
comillas y daña las filas cuyo nombre de institución lleva una coma. Si
lo quiere mirar, use el Bloc de notas.

**2. Mire la vista previa.** Opción 1 del menú. Confirme que diga
*"Medidas y peso correctos"* y que el texto esté bien.

**3. Mándese un correo de prueba.** Opción 2. Ábralo y **haga clic en el
botón amarillo**: tiene que abrir el formulario de inscripción.

---

## Cómo se usa

Doble clic en **`enviar_campana.bat`**. Sale un menú:

```
  1. Vista previa (no envia nada)
  2. Correo de prueba a una direccion mia
  3. ENVIO REAL del lote de hoy
  4. Ver el estado de la campana
  5. Revisar el buzon (rebotes y retiros)
  6. Verificar que los dominios existan (tarda, es opcional)
```

Le va a pedir la **clave de aplicación de Google**. Se escribe invisible,
eso es normal. No queda guardada en ningún archivo.

Para cambiar el tope diario: abra `config.ini` con el Bloc de notas,
busque `maximo_por_dia` y cambie el número. Guarde.

---

## Durante la campaña

**Revise el buzón cada dos días** (opción 5). Busca rebotes y gente que
pide que la saquen, y los anota para no volverles a escribir. Con una
base de este tamaño van a salir varios.

Si los rebotes pasan del 5 %, pare y avíseme: significa que la base
trae direcciones viejas y seguir mandando castiga la reputación del
dominio.

**El brochure va adjunto.** Con el PDF, cada correo pesa **2.2 MB** en
vez de 144 KB. Funciona, pero tenga presente dos cosas:

- Un adjunto pesado en correo masivo en frío es una señal que los
  filtros miran. En septiembre se adjuntó y llegó a bandeja de entrada
  sin problema, pero eran 415 correos, no 2.849.
- Si los rebotes se disparan o empieza a caer en SPAM, **lo primero que
  hay que quitar es el adjunto**. En `config.ini`, borre lo que va
  después de `adjunto =` y guarde. El correo sale igual, solo con el
  enlace.

---

## El día antes

El martes 13 hay que mandarle a los inscritos el enlace de conexión del
Meet (`https://meet.google.com/hvv-zdpe-syb`). Eso es un envío aparte,
igual que la vez pasada: se baja la lista de respuestas del formulario y
se manda solo a esa gente.
