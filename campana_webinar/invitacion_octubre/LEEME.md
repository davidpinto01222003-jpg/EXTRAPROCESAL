# Segunda sesión del webinar — envío de invitaciones

**Miércoles 14 de octubre, 8:00 a. m.**
Inscripciones: https://forms.gle/rheWoEX9D3X2NjYm6

---

## Lo primero: esto NO se manda de un solo golpe

La base tiene **2.715 correos**, uno por IPS, ordenados de mayor a menor
capacidad instalada. La campaña de septiembre fueron 415. Es casi siete
veces más.

La cuenta `contactoclientes@oscal.net` nunca ha enviado ese volumen, y
**un salto repentino de volumen es lo que más dispara los filtros de
SPAM** — más que el contenido, más que las imágenes. Gmail y Outlook
miran cuánto manda usted normalmente; si de un día para otro pasa de
cero a tres mil, lo tratan como una cuenta comprometida.

La solución es subir de a poco. El programa ya trae un tope diario
(`maximo_por_dia` en `config.ini`) y lleva el registro de lo que ya
mandó, así que **nunca repite un correo**: usted solo sube el número
cada día y vuelve a correr el envío.

### Calendario sugerido

| Día | `maximo_por_dia` | A quién le llega |
|---|---|---|
| Martes 6 de octubre | 250 | tramo A, las 250 IPS más grandes |
| Miércoles 7 | 450 | tramo B |
| Jueves 8 | 650 | tramo C |
| Viernes 9 | 850 | tramo D |
| Lunes 12 | — | **FESTIVO, no enviar** |
| Martes 13 | 600 | tramo E, las 515 restantes |

El número de la columna del medio es **acumulado del día**: el programa
cuenta cuántos lleva enviados hoy y para cuando llega al tope.

**Cada lote se demora.** Entre un correo y el siguiente hay una pausa de
10 a 20 segundos, a propósito: enviar de corrido es lo que delata un
envío automático. En la práctica:

| Correos | Tiempo aproximado |
|---|---|
| 250 | 1 hora |
| 450 | 1 hora 50 |
| 650 | 2 horas 40 |
| 850 | 3 horas 30 |

El computador tiene que quedar prendido y el programa abierto todo ese
rato. Si se interrumpe no pasa nada grave: el programa anota lo que ya
mandó y al volver a correr sigue donde quedó.

El lunes 12 es festivo en Colombia (Día de la Raza). Un correo
institucional ese día lo lee nadie.

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

**1. Traiga los correos excluidos de la campaña pasada.** Copie:

```
campana_webinar\estado\excluidos.csv  →  invitacion_octubre\estado\excluidos.csv
```

Ese archivo tiene los 10 que rebotaron y los dominios muertos. Si no lo
copia, el programa les vuelve a escribir y los rebotes se repiten.

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
