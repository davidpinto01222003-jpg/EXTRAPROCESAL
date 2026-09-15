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

### 5. Revisar rebotes y retiros

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

Calendario sugerido hasta el webinar:

| Día | Cantidad | A quién |
|---|---|---|
| Martes 15 | 150 | Clientes actuales y contactos con correo previo |
| Miércoles 16 | 300 | Los más cercanos del resto |
| Jueves 17 | 600 | |
| Viernes 18 | 1.000 | |
| Lunes 21 | 1.500 | |
| Martes 22 | recordatorio | A los que no se inscribieron |

Son unos 3.550 contactos. El fin de semana no conviene enviar. Acuérdese
de subir `maximo_por_dia` en `config.ini` cada día.

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
| `estado/excluidos.csv` | Rebotes y quienes pidieron el retiro |

La base real y el registro de envíos **no se suben al repositorio**:
son datos personales de terceros.
