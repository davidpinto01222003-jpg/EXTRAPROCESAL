# Campaña del webinar — OSCAL Consultores Jurídicos

Envío individual y personalizado de la invitación al webinar
**"Recuperación de cartera con EPS intervenidas"**, miércoles 23 de
septiembre, 8:00 a. m., desde `contactoclientes@oscal.net`.

Cada destinatario recibe **su propio correo**, dirigido a su institución.
Nadie ve la dirección de nadie más: no se usa copia oculta.

No necesita instalar nada: el programa solo usa lo que ya trae Python.

---

## Antes de la primera vez

**1. Contraseña de aplicación de Google.** La contraseña normal de la
cuenta no sirve para esto. Entre a
[myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)
con `contactoclientes@oscal.net`, cree una contraseña nueva y guarde las
16 letras. El programa la pide cada vez que envía y **no la guarda en
ningún archivo**.

Si quiere no escribirla cada vez, déjela en una variable de entorno
antes de ejecutar:

```
set OSCAL_SMTP_PASSWORD=lasdieciseisletras
```

**2. La imagen del webinar.** Guarde el banner horizontal (el de
formato ancho, no el cuadrado) dentro de la carpeta `recursos/` con el
nombre `imagen_webinar.jpg`. Sirve igual en `.png`. Ancho recomendado:
1200 píxeles, menos de 200 KB.

Si no pone ninguna imagen el correo se envía igual, solo que sin ella y
sin dejar un recuadro roto.

**3. La base de contactos.** Cree `contactos.xlsx` (o `.csv`) en esta
misma carpeta. Puede partir de `contactos_ejemplo.csv`. Solo el correo
es obligatorio:

| correo | institucion | ciudad |
|---|---|---|
| gerencia@clinicauno.com | Clínica Uno S.A.S. | Bucaramanga |
| cartera@hospital.gov.co | E.S.E. Hospital San Juan | Floridablanca |
| solo-el-correo@otra.com | | |

Los nombres de las columnas pueden variar: sirve *correo*, *email* o
*e-mail*; *institucion*, *entidad*, *empresa* o *razón social*;
*ciudad* o *municipio*. No importan las tildes ni las mayúsculas del
encabezado.

**Bases con varias columnas de correo.** Si la base trae más de una
columna de correo (como la del área metropolitana, que tiene la de
jurídica, la de contratación, el canal de contacto y la del REPS), el
programa escoge en este orden: *jurídica y notificaciones judiciales*,
después *contratación y proveedores*, después *canal para enviar
propuesta*, y de último *correo REPS*. Así el mensaje le llega a quien
de verdad maneja el tema de cartera.

También entiende celdas con texto alrededor, como
`sergioruiz@fcv.org (General)`, y se queda solo con la dirección.

**Orden de envío.** Si hay una columna de *segmento* o *prioridad*, se
respeta: los "A - contactar primero" salen antes que los "B".

**Clientes actuales.** En la base del área metropolitana las filas sin
razón social no son un error: se les quitó el nombre a propósito porque
ya son clientes de la firma, y no hay que invitarlos en frío. El
programa las deja por fuera. Esto se controla en `config.ini`:

```
[base]
excluir_sin_institucion = si
```

Con `no` entran a la campaña y reciben el saludo genérico
"Señores / Ciudad". Si su base no usa ese criterio, póngalo en `no`.

**Municipios.** Si la celda trae varios (`BUCARAMANGA, FLORIDABLANCA,
GIRON`) se usa el primero, que es lo que va en el encabezado de una
carta.

---

## Cómo se usa

En Windows, doble clic en `enviar_campana.bat` y elija la opción. O
desde PowerShell, parado en esta carpeta:

### 1. Ver cómo queda el correo, sin enviar nada

```
python enviar_campana.py vista-previa
```

Arma tres correos de ejemplo con contactos reales de su base y los deja
en `estado/vista_previa/`. Ábralos con doble clic. **No se conecta a
internet ni envía nada.**

### 2. Enviarse un correo de prueba

```
python enviar_campana.py prueba --para tucorreo@gmail.com
```

Envía uno de verdad. Después, en ese correo, entre a los ⋮ tres puntos
y elija **Mostrar original**: debe decir PASS en SPF, DKIM y DMARC.

Sirve también para mail-tester.com, que califica el correo de 1 a 10:

```
python enviar_campana.py prueba --para la-direccion-que-da@srv1.mail-tester.com
```

### 3. El envío real

```
python enviar_campana.py enviar
```

Muestra el plan, pide que escriba `ENVIAR` para confirmar, y arranca.
Se puede interrumpir con Ctrl+C y retomar después con la misma orden:
**nunca le escribe dos veces a la misma dirección.**

Para enviar menos de lo que permite el tope del día:

```
python enviar_campana.py enviar --maximo 150
```

Para ensayar el recorrido completo sin enviar nada:

```
python enviar_campana.py enviar --simulacro
```

### 4. Ver cómo va

```
python enviar_campana.py estado
```

### 5. Descartar dominios que ya no existen

```
python enviar_campana.py verificar
```

Consulta el DNS de cada dominio de la base y muestra cuáles ya no
reciben correo (empresas que cerraron, dominios vencidos). Pregunta
antes de excluirlos. Es gratis y evita rebotes, que son lo que más daña
la reputación del dominio.

Ante una falla de red **no descarta a nadie**: prefiere enviar de más
que perder un cliente bueno. Por eso conviene correrlo dos veces y
quedarse con lo que aparezca en ambas.

**Córralo antes del primer envío.**

### 6. Revisar rebotes y retiros

```
python enviar_campana.py revisar-buzon
```

Entra al buzón, busca los correos que rebotaron y las respuestas que
piden no recibir más, y los anota en `estado/excluidos.csv` para que el
siguiente envío ya no les escriba. También lista las respuestas que
vale la pena mirar. **Conviene correrlo todos los días.**

---

## Lo que hace solo

- Quita duplicados y direcciones mal escritas.
- Corrige dominios mal tecleados (`gmail.co` → `gmail.com`).
- Arregla los nombres en MAYÚSCULA SOSTENIDA, respetando siglas como
  IPS, E.S.E. o S.A.S.
- Espera entre 10 y 20 segundos al azar entre un correo y el siguiente.
- Respeta el tope diario de `config.ini`.
- Si Google corta por límite de envíos, se detiene solo y avisa.
- Deja todo anotado en `estado/enviados.csv`.

---

## Los topes de Google

Google Workspace permite **2.000 destinatarios al día por cuenta**,
contando los correos normales de la oficina. Además el volumen debe
subir de a poco: una cuenta que nunca ha enviado masivamente y de un
día para otro manda mil correos se gana el filtro de SPAM.

La base actual queda en **419 contactos** (454 filas, menos 7
duplicadas y menos 28 clientes actuales), así que cabe entera en cuatro
días sin forzar nada:

| Día | Cantidad | A quién |
|---|---|---|
| Martes 15 | 50 | Segmento A |
| Miércoles 16 | 92 | Resto del segmento A |
| Jueves 17 | 140 | Segmento B |
| Viernes 18 | 140 | Segmento B |
| Lunes 21 | — | Margen por si algo falla |
| Martes 22 | recordatorio | A los que no se inscribieron |

El programa ya envía en ese orden: primero los "A - contactar primero",
después los "B". Solo hay que ir subiendo `maximo_por_dia` en
`config.ini`, o pasar `--maximo 50` en cada corrida.

El fin de semana no conviene enviar: baja la apertura y sube el riesgo
de que lo marquen como SPAM.

---

## El brochure adjunto

El correo va con el brochure en PDF adjunto. Eso hace que cada mensaje
pese **1,3 MB en vez de 23 KB**.

Funciona, pero un adjunto de ese tamaño enviado a miles de desconocidos
sube el riesgo de SPAM. La alternativa es subir el PDF a la página de la
firma y dejar solo el enlace. Para quitar el adjunto, deje el valor
vacío en `config.ini`:

```
adjunto =
```

Si lo hace, acuérdese de quitar también la frase que menciona el
portafolio adjunto, si la agrega al texto.

---

## Si algo sale mal

**"Google rechazó la contraseña"** — está usando la clave normal de la
cuenta. Necesita una contraseña de aplicación de 16 letras.

**"No encuentro la base de contactos"** — el archivo `contactos.xlsx`
no está en esta carpeta, o tiene otro nombre. Revise `config.ini`,
sección `[base]`.

**Los correos llegan a SPAM** — revise primero que la prueba dé PASS en
los tres (SPF, DKIM, DMARC) y que mail-tester dé 9 o más. Si eso está
bien, casi siempre es que se subió el volumen demasiado rápido: baje el
tope diario y escríbale primero a los contactos conocidos.

**Se cortó a la mitad** — vuelva a ejecutar la misma orden. Retoma donde
iba.

---

## Archivos

| Archivo | Para qué |
|---|---|
| `enviar_campana.py` | El programa |
| `enviar_campana.bat` | Menú para Windows |
| `config.ini` | Remitente, asunto, topes y pausas |
| `plantilla_correo.html` | El diseño del correo (se puede editar) |
| `texto_correo.txt` | La versión en texto plano |
| `contactos_ejemplo.csv` | Modelo de la base |
| `recursos/` | Brochure en PDF y la imagen |
| `estado/enviados.csv` | Qué se envió y cuándo |
| `estado/excluidos.csv` | Rebotes, retiros y dominios muertos |

La base real y el registro de envíos **no se suben al repositorio**:
son datos personales de terceros.
