# Segunda sesión del webinar — envío de invitaciones

**Miércoles 14 de octubre, 8:00 a. m.**
Inscripciones: https://forms.gle/rheWoEX9D3X2NjYm6

---

## Lo primero: esto NO se manda de un solo golpe

La base tiene **3.864 correos**. La campaña de septiembre fueron 415.
Es casi diez veces más.

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

| Día | `maximo_por_dia` | Correos |
|---|---|---|
| Martes 6 de octubre | 250 | 250 |
| Miércoles 7 | 450 | 200 más |
| Jueves 8 | 650 | 200 más |
| Viernes 9 | 850 | 200 más |
| Lunes 12 | — | **FESTIVO, no enviar** |
| Martes 13 | 1500 | el resto |

El número de la columna del medio es **acumulado del día**: el programa
cuenta cuántos lleva enviados hoy y para cuando llega al tope.

El lunes 12 es festivo en Colombia (Día de la Raza). Un correo
institucional ese día lo lee nadie.

### Hasta dónde enviar

La base viene en dos grupos, y el programa envía en ese orden:

- **Segmento A — 2.849 correos.** El correo principal de cada IPS.
- **Segmento B — 1.015 correos.** Correos de sedes individuales de IPS
  que **ya recibieron** la invitación en su correo principal.

**La recomendación es enviar solo el segmento A** y dejar el B para una
próxima campaña. Escribirle dos veces a la misma institución, en la
misma semana, por el mismo evento, es exactamente lo que un filtro de
SPAM interpreta como envío masivo indiscriminado.

Con el calendario de arriba el segmento A queda cubierto el martes 13.

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

**El brochure no va adjunto a propósito.** Con 3.864 destinatarios, un
PDF de 1 MB en cada correo multiplica por seis el peso del envío y sube
el riesgo de SPAM. Quien se inscriba lo recibe después.

---

## El día antes

El martes 13 hay que mandarle a los inscritos el enlace de conexión del
Meet (`https://meet.google.com/hvv-zdpe-syb`). Eso es un envío aparte,
igual que la vez pasada: se baja la lista de respuestas del formulario y
se manda solo a esa gente.
