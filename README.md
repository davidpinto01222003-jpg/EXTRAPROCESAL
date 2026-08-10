# Procesos Jurídicos: herramienta única (correo → descarga → organizado)

Un solo programa, `procesos_juridicos.py`, que hace dos cosas a la vez:

1. **Vigila tu Gmail**: cuando un juzgado te comparte un expediente por el
   SGDE (Rama Judicial), entra automáticamente al portal, resuelve el
   token de verificación (leyéndolo de tu correo), descarga todo lo que
   haya en "Elementos Compartidos" —incluidas las carpetas que no tienen
   flecha de descarga directa, entrando a ellas y descargando archivo por
   archivo— y deja el resultado ya descomprimido en tu disco duro, en una
   carpeta nombrada con el número de expediente (23 dígitos).
2. **Vigila tu carpeta de Descargas**: por si alguna vez bajas un zip de
   un proceso a mano (de otro sistema, por ejemplo), lo extrae, busca el
   radicado dentro de los PDF/DOCX y lo organiza igual en el disco duro.

Si configuras `RUTA_EXCEL` en `validar_renombrar_carpetas.py` (ver más
abajo), cada carpeta nueva —de correo o manual— se cruza automáticamente
contra ese informe: si el radicado ya aparece ahí, la carpeta queda
nombrada `"numero. radicado"` desde el momento en que se crea, en vez de
solo el radicado. Si el radicado todavía no está en el informe, la
carpeta se deja solo con el radicado (como antes) y queda registrada en
el log; más tarde puedes correr `validar_renombrar_carpetas.py` para
completar el nombre cuando el informe se actualice.

Todo queda registrado en `procesos_juridicos.log`, dentro de la carpeta
destino.

## Instalación (Windows)

1. Instala [Python 3.10+](https://www.python.org/downloads/) marcando la
   opción "Add Python to PATH" durante la instalación.
2. Abre una terminal (`cmd` o PowerShell) en esta carpeta y ejecuta:

   ```
   pip install -r requirements.txt
   playwright install chromium
   ```

## Configuración

Abre `procesos_juridicos.py` y edita, al inicio del archivo, la sección
`CONFIGURACION`:

- `CARPETA_DESCARGAS`: carpeta donde el navegador guarda los `.zip`
  (por ejemplo `C:\Users\TuUsuario\Downloads`).
- `ETIQUETA_DISCO_EXTERNO`: el NOMBRE de tu disco duro externo (por
  ejemplo `"OSCAL"`, tal como aparece en "Este equipo"). El script busca
  el disco por ese nombre entre todas las unidades conectadas y usa la
  letra que encuentre en ese momento -- así no importa si Windows le
  asigna `D:`, `E:` o cualquier otra letra la próxima vez que lo
  conectes. `CARPETA_DESTINO` se calcula sola a partir de esto; no hace
  falta que edites `CARPETA_DESTINO` directamente. Si el disco no se
  encuentra conectado, se usa `CARPETA_DESTINO_RESPALDO` como respaldo
  (y el log avisa claramente que pasó eso).
- `PATRONES_RADICADO`: usados solo para zips que descargues a mano.
  Reconoce por defecto el radicado judicial colombiano de 23 dígitos.

## Configurar el acceso a tu Gmail (para la vigilancia automática)

El programa necesita leer tu correo (vía IMAP) para conseguir el link de
cada expediente y el token de verificación. **Nunca uses tu contraseña
normal de Gmail** — usa una "Contraseña de aplicación":

1. Activa la verificación en 2 pasos (una sola vez):
   https://myaccount.google.com/security
2. Crea una contraseña de aplicación (una sola vez):
   https://myaccount.google.com/apppasswords
3. En esta carpeta, copia `credenciales_sgde.example.txt` y renómbralo a
   `credenciales_sgde.txt`.
4. Ábrelo y reemplaza los valores por tu correo y la contraseña de
   aplicación de 16 letras.

Una vez configurada, el inicio de sesión con esa contraseña **no pide
código ni confirmación en el teléfono** — el 2FA interactivo solo aplica
a logins manuales desde el navegador, no al uso de contraseñas de
aplicación por IMAP.

`credenciales_sgde.txt` está en `.gitignore`: nunca se sube a git, ni
debes compartirlo con nadie.

Si no creas este archivo, el programa sigue funcionando igual, solo que
sin la vigilancia automática de correo (únicamente organiza lo que
descargues a mano).

### ⚠️ Antes de dejarlo corriendo

Verifica que automatizar el acceso al SGDE con tu cuenta autorizada no
infrinja los términos de uso del portal. El programa solo automatiza
pasos que tú ya estás autorizado a hacer manualmente (abrir el link que
te comparten, escribir tu correo y el token que te llega a ti mismo); no
intenta acceder a expedientes ajenos ni evade ninguna medida de
seguridad.

## Uso

Haz doble clic en `iniciar.bat`. Hace las dos cosas en orden, en la misma
ventana:

1. Corre `validar_renombrar_carpetas.py` una vez (revisa/renombra las
   carpetas que ya existen en el disco contra el informe de Excel) y
   muestra su reporte.
2. Cuando termina, arranca `procesos_juridicos.py` y se queda vigilando
   Descargas (y correo, si configuraste `credenciales_sgde.txt`) de forma
   indefinida.

**Importante**: los dos pasos corren UNO DESPUÉS DEL OTRO, no al mismo
tiempo -- mientras el Paso 1 sigue trabajando (puede tardar varios
minutos si `VALIDAR_CONTENIDO_CONTRA_NOMBRE = True`), la vigilancia de
Descargas todavía no ha arrancado, así que un zip que descargues justo
en ese momento no se procesa solo. Si quieres que la vigilancia esté
corriendo siempre, sin depender de que termine el Paso 1, abre
`vigilar.bat` en una ventana aparte -- corre solo `procesos_juridicos.py`
de una vez, sin esperar nada.

O si prefieres correr cada uno por separado, desde una terminal:

```
python validar_renombrar_carpetas.py
python procesos_juridicos.py
```

La primera vez, deja `NAVEGADOR_VISIBLE = True` (así viene por defecto)
para ver el navegador trabajando y confirmar que todo va bien. Cuando
confíes en que funciona, cámbialo a `False` en `procesos_juridicos.py`
para que corra en segundo plano sin abrir ventanas.

Déjalo abierto; ciérralo con `Ctrl+C` (o cerrando la ventana) cuando
termines.

### Si no quieres dejarlo corriendo de fondo

Por defecto el programa se queda vigilando Descargas en tiempo real (y el
correo, si configuraste `credenciales_sgde.txt`) hasta que lo cierres.

Si prefieres correrlo una vez al día y que termine solo —por ejemplo con
el Programador de tareas de Windows, en vez de dejar una ventana
abierta— abre `procesos_juridicos.py` y cambia:

```
SOLO_PROCESAR_HOY_Y_SALIR = True
```

Con esto, cada vez que lo corras organiza únicamente los `.zip` de los
últimos `DIAS_ATRAS_PROCESAR_EXISTENTES` días (por defecto 14, es decir
las últimas dos semanas) que ya estén en `CARPETA_DESCARGAS` (igual que hace
siempre al arrancar) y termina inmediatamente: no vigila el correo ni se
queda esperando descargas nuevas. Si no tienes `credenciales_sgde.txt`,
el correo ya se salta de por sí; este interruptor es para el otro caso,
cuando tampoco quieres que se quede vigilando Descargas.

Ese mismo `DIAS_ATRAS_PROCESAR_EXISTENTES` controla también, al arrancar
en modo vigilancia normal, cuántos días hacia atrás de zips ya existentes
se procesan (ponlo en `1` si solo quieres los de hoy).

## Ejecutarlo automáticamente al iniciar Windows (opcional)

1. Abre el "Programador de tareas" de Windows.
2. Crea una tarea nueva que se ejecute "Al iniciar sesión".
3. Como acción, apunta a `iniciar.bat` (o a `pythonw.exe
   procesos_juridicos.py` si prefieres que corra sin ventana visible).

## Cómo evita repetir trabajo

- Cada expediente ya descargado por correo se guarda en
  `expedientes_procesados.txt` (no se sube a git). Para volver a
  descargar uno, borra esa línea o el archivo completo.
- Los zips descargados a mano se mueven a una subcarpeta `Procesados`
  dentro de tu carpeta de Descargas (no se borran).
- Si ya existe una carpeta con el mismo nombre en el destino, se agrega
  un sufijo (`_2`, `_3`, ...) en vez de sobrescribir.

## Validar y renombrar carpetas contra el informe de Excel

`validar_renombrar_carpetas.py` es una herramienta aparte (no se ejecuta
junto con `procesos_juridicos.py`). Sirve para cuando ya tienes en el disco
duro carpetas nombradas solo con el radicado de 23 dígitos y quieres
verificar que cada una corresponda a un proceso del informe de Excel,
renombrándola a `"<numero>. <radicado>"` (por ejemplo
`133. 68001400301020180087200`).

El cruce se hace por el **valor exacto del radicado** (columna `RADICADO`
del Excel) contra el radicado que aparece en el nombre de cada carpeta; el
número de proceso sale de la columna `No.` de la misma fila. No importa si
faltan carpetas o filas, cada una se empareja de forma independiente.

1. Instala la dependencia nueva (ya incluida en `requirements.txt`):

   ```
   pip install -r requirements.txt
   ```

2. Abre `validar_renombrar_carpetas.py` y edita, al inicio del archivo, la
   sección `CONFIGURACION`:

   - `RUTA_EXCEL`: ruta al informe (`.xlsx` o `.xlsm`).
   - `HOJA_EXCEL`, `FILA_ENCABEZADO`, `COLUMNA_NO`, `COLUMNA_RADICADO`: en
     qué hoja y fila están los encabezados, y cómo se llaman las columnas
     del número de proceso y del radicado.
   - `ETIQUETA_DISCO_EXTERNO`: el NOMBRE de tu disco duro externo (por
     ejemplo `"OSCAL"`). `CARPETA_PROCESOS` se calcula sola buscando ese
     disco por su nombre entre las unidades conectadas, así no importa
     qué letra (`D:`, `E:`, etc) le asigne Windows esta vez.
   - `CARPETA_ENTRADA_ADICIONAL`: una carpeta aparte (por defecto
     `PROCESOS LAUE/ENTREGA EXPEDIENTE ESSA` dentro del disco) donde a
     veces caen entregas masivas de expedientes ya extraídos que todavía
     no se pasan a la raíz. Si existe, el script mueve a la raíz los que
     tengan radicado válido en el Excel y no estén ya en el disco, deja
     donde están los que ya existan (avisando), y deja aparte (avisando)
     los que no tengan proceso en el Excel. Si la ruta no existe, este
     paso simplemente se omite.
   - `CARPETA_DESCARGAS`: tu carpeta de Descargas (se detecta sola). Se usa
     solo para revisar si una carpeta vacía tiene un `.zip` pendiente de
     extraer ahí.

   Esta misma configuración (`RUTA_EXCEL`, `HOJA_EXCEL`, etc) es la que usa
   `procesos_juridicos.py` para nombrar bien las carpetas nuevas apenas las
   crea — no hay que configurarla dos veces.

3. Ejecuta el script (por defecto corre en `MODO_PRUEBA = True`, así que no
   renombra nada todavía, solo muestra un reporte):

   ```
   python validar_renombrar_carpetas.py
   ```

4. Revisa el reporte en pantalla y en `validar_renombrar_carpetas.log`:
   qué se renombraría, qué carpetas ya tienen el nombre correcto, qué
   procesos del Excel no tienen carpeta en el disco, y qué carpetas del
   disco no aparecen en el Excel (radicados repetidos o con formato
   inválido en el Excel se reportan y se omiten del cruce, para no
   arriesgar un renombrado incorrecto).

   Casos especiales:
   - `[Consecutivo]`: la carpeta y el Excel coinciden en los primeros 22
     dígitos del radicado y solo difieren en el último (el proceso ya
     cambió de instancia, ej. de "...00" a "...01") — esto **sí se
     corrige automático**, dejando el radicado del Excel.
   - `[Duplicado en Excel]`: el mismo radicado aparece en **dos o más
     filas** del Excel con números de proceso **distintos** (ej. proceso
     19 y proceso 451 con exactamente el mismo radicado). En vez de
     excluir esas filas del cruce, el script **duplica la carpeta**:
     conserva/renombra la existente con el primer número, y crea una
     copia completa (mismo contenido) por cada número adicional, para
     que cada número de proceso tenga su propia carpeta
     `"numero. radicado"`. Este caso también se reconoce **combinado**
     con `[Consecutivo]`: si la carpeta suelta tiene el último dígito
     distinto al del Excel, y ese radicado del Excel además está
     repetido con varios números, se corrige el dígito **y** se duplica
     para cada número, todo en la misma pasada.
   - `[POSIBLE COINCIDENCIA]`: la carpeta y el Excel difieren por un
     dígito de más o de menos en **otra** posición — esto **no** se
     corrige solo, se reporta para que lo confirmes a mano (dos procesos
     distintos del mismo juzgado y año pueden compartir casi todos los
     dígitos, así que adivinar mal sería peligroso).
   - `[Duplicado]` / `[Duplicados resueltos]`: cuando el mismo radicado
     aparece en más de una carpeta (típico de descargas repetidas), el
     script **nunca borra nada**. Se queda con la carpeta que tenga más
     archivos adentro (la más completa), la renombra con el número más
     reciente del Excel, y mueve las demás copias — intactas — a una
     carpeta `Duplicados_para_revisar` dentro de tu disco, para que las
     revises y borres a mano si de verdad sobran.
   - `[Duplicado sin resolver]`: hay más de una carpeta con el mismo
     radicado, pero ese radicado no está en el Excel — no se puede saber
     cuál conservar, así que no se toca ninguna; revísalas a mano.
   - `[Anidada]` / `[Anidadas resueltas]` / `[Anidadas de otro caso]`:
     cuando una carpeta de proceso queda metida DENTRO de otra (ej.
     `1014. radicado` adentro de `941. radicado`), también se resuelve
     sola, sin borrar ni fusionar contenido — solo mueve la carpeta
     completa. Si el radicado de la anidada es el mismo que el de la
     carpeta que la contiene, **o solo difiere en el último dígito**
     (el consecutivo de instancia/reparto, ej. termina en 0 o en 1 —
     es el mismo caso, no un proceso distinto), se mueve a
     `Duplicados_para_revisar`. Si es un radicado realmente distinto
     (contenido de otro caso mal ubicado), se saca al nivel principal
     del disco para evaluarla en la próxima corrida.
   - `[Carpeta vacía]`: carpetas sin ningún archivo adentro. El script
     revisa si hay un `.zip` con ese mismo radicado, tanto en
     `CARPETA_DESCARGAS` directamente como en su subcarpeta `Procesados`
     (ahí es donde quedaban los zips que en versiones viejas del programa
     se marcaban como "ya procesados" aunque la extracción hubiera
     fallado por completo) y te lo señala — esto solo se reporta, no se
     extrae ni se toca nada automático. Además de salir en el log, queda
     un reporte aparte en `carpetas_vacias.csv` (carpeta, radicado, el
     zip encontrado, y si estaba en Descargas o en Procesados) para que
     lo revises en Excel sin tener que buscar en el log completo.
   - `[Sin nombre reconocible]`: carpetas que no tienen ningún número que
     se parezca a un radicado en su nombre (antes se ignoraban en
     silencio, ahora se listan para que las revises a mano).
   - `[Contenido no corresponde]`: abre los documentos DENTRO de cada
     carpeta (nombres de archivo, y si hace falta el texto de hasta
     `MAX_ARCHIVOS_CONTENIDO_A_REVISAR` PDF/DOCX) y revisa si el radicado
     que aparece adentro corresponde con el radicado del nombre de la
     carpeta. Si el radicado del nombre nunca aparece en su propio
     contenido pero otro sí aparece claramente, se marca como sospechosa
     de tener contenido de otro caso mal ubicado (esto fue justo lo que
     pasó con algunos zips antes de la corrección del radicado-por-nombre-
     de-zip). Queda además en un reporte aparte:
     `contenido_no_corresponde.csv`. Esta revisión tarda más porque tiene
     que leer documentos de todas las carpetas; puedes desactivarla
     poniendo `VALIDAR_CONTENIDO_CONTRA_NOMBRE = False` si prefieres una
     corrida rápida.

5. Si el reporte se ve bien, cambia `MODO_PRUEBA = False` y vuelve a
   correrlo para aplicar los renombrados de verdad. Puedes correrlo las
   veces que quieras: las carpetas que ya tengan el nombre correcto se
   dejan igual.

## Validar el juzgado contra el portal de Rama Judicial

`validar_juzgados_ramajudicial.py` es otra herramienta aparte. Consulta
cada radicado del informe de Excel en el portal público **Consulta de
Procesos Nacional Unificada** (CPNU,
`consultaprocesos.ramajudicial.gov.co`) y compara el despacho que
reporta el portal contra el juzgado anotado en la columna `JUZGADO` del
Excel, para detectar procesos que ya cambiaron de despacho pero el Excel
todavía no se actualizó.

Cómo hace el cruce: el radicado termina en un "consecutivo" (últimos 2
dígitos) que cambia cuando el proceso pasa a otro despacho. El script
consulta el radicado tal como está en el Excel, y luego prueba el
consecutivo siguiente (+1, +2, ...) mientras el portal siga encontrando
resultado; el despacho del último consecutivo que sí exista es el que se
compara contra la columna `JUZGADO`.

⚠️ **Antes de usarlo, ten en cuenta:**

- Este script no se pudo probar contra el portal real antes de
  entregarlo (el entorno donde se escribió no tiene acceso a ese sitio).
  Corre primero con `SOLO_ESTOS_NUMEROS = [1, 133]` (o los que tú
  elijas) para confirmar que funciona antes de lanzarlo contra todo el
  informe, y avísame qué error sale si algo falla para ajustar el script.
- La comparación de nombres de juzgado es tolerante a diferencias de
  formato (el Excel dice "PRIMERO CIVIL MUNICIPAL...", el portal dice
  "JUZGADO 001 CIVIL MUNICIPAL... (SANTANDER)"), pero no es infalible:
  revisa el CSV completo (no solo la lista de diferencias) si algo se ve
  raro.
- Hace una pausa entre cada consulta para no saturar un portal público
  del Estado; con ~950 procesos, considera correrlo en un par de tandas.
  Si lo interrumpes con `Ctrl+C`, la próxima vez sigue donde iba en vez
  de repetir consultas ya hechas (gracias a
  `validar_juzgados_progreso.txt`).
- No hay que resolver ningún captcha (se confirmó manualmente que la
  consulta por número de radicado no lo pide).

Uso:

1. Edita la sección `CONFIGURACION` al inicio del archivo: `RUTA_EXCEL`,
   `HOJA_EXCEL`, y si quieres, `SOLO_ESTOS_NUMEROS` para una prueba
   chica primero.
2. Corre:

   ```
   python validar_juzgados_ramajudicial.py
   ```

3. Revisa `validar_juzgados_reporte.csv` (el detalle completo de todos
   los procesos consultados) y el resumen final en pantalla / en
   `validar_juzgados_ramajudicial.log` (solo las diferencias y los
   radicados que el portal no encontró).

## Buscar en Google Drive/correo los procesos que faltan en el disco

`buscar_faltantes_en_drive.py` es otra herramienta aparte. Lee
`procesos_faltantes_en_disco.csv` (lo genera `validar_renombrar_carpetas.py`)
y busca cada proceso faltante en tu Google Drive y en tu correo de Gmail,
por este orden:

1. El **radicado completo** de 23 dígitos -- si encuentra una
   coincidencia exacta, la descarga automático (tan específico que no
   hay riesgo real de confundirlo con otro caso).
2. El **radicado corto** (ej. `2025-00456`, `2025-456`,
   `2025-00456-00`, o `2025-456-00` -- con o sin ceros a la izquierda
   del consecutivo, y con o sin el consecutivo de instancia/reparto al
   final, ver `radicados_cortos`).
3. El número de **cuenta** -- solo si es lo bastante específica. Una
   cuenta vacía, `"0"`, o de muy pocos dígitos **no se busca**: como la
   búsqueda de Drive es aproximada, un término tan genérico coincidiría
   con miles de carpetas de todo el Drive sin relación, dejando el
   proceso "pegado" un buen rato revisando falsos positivos uno por
   uno.

Los casos 2 y 3 son menos confiables (un radicado corto o una cuenta
puede coincidir por casualidad con archivos de otro proceso); para
esos, antes de descargar también se valida que la carpeta (o sus
archivos) de verdad mencionen ese radicado.

**Regla obligatoria, sin excepción, para TODO lo que se descargue**
(sea confiable o no, radicado completo o corto, carpeta de Drive o
adjunto de correo): el documento tiene que mencionar a **ESSA** o
**ELECTRIFICADORA DE SANTANDER** (como demandante o como demandado --
por nombre, o abriendo el contenido de sus PDF/DOCX si el nombre no lo
dice). Si no la menciona, no se descarga -- ni siquiera si el radicado
coincidió exacto. Un adjunto de correo que no pase esta regla se
extrae igual, pero a `Duplicados_para_revisar` en vez de a la carpeta
del proceso (nunca se pierde, solo no se mezcla con el caso).

**También se valida el DEMANDADO** (columna `DEMANDADO` del Excel, si
existe): el mismo radicado corto o cuenta se puede repetir entre
procesos **distintos** que van contra demandados diferentes -- por
ejemplo, `"2024-00139 CONTRA RIONEGRO"` no es lo mismo que `"2024-00139
CONTRA BOLIVAR"`. Con que el nombre del demandado esperado aparezca en
**cualquier parte** del nombre/asunto/contenido del candidato alcanza
para confirmarlo -- no hace falta que esté después de la palabra
"CONTRA". Solo se descarta un candidato (una carpeta, un archivo dentro
de una carpeta genérica, o un adjunto de correo) cuando sí trae un
`"CONTRA <algo>"` (la forma habitual de nombrar expedientes) pero ese
"algo" es un demandado **distinto** al esperado -- aunque ya haya
pasado la validación de ESSA. Si el candidato no menciona ni al
demandado esperado ni ningún otro `"CONTRA <algo>"`, esta validación no
bloquea nada (no hay evidencia ni a favor ni en contra).

Los candidatos que sí pasan **también se descargan automático**. Si para el mismo proceso aparece
MÁS de un candidato válido (ej. el expediente está repartido en varias
carpetas de Drive -- una con el "poder", otra con el "expediente"),
todos quedan **fusionados dentro de UNA sola carpeta** en el disco (no
se crean "_2", "_3", etc): el primer candidato crea la carpeta, y cada
candidato siguiente que también pase las validaciones se copia dentro
de esa misma carpeta. Como cada candidato es una carpeta de Drive
DISTINTA, puede traer -- por pura coincidencia -- un archivo o
subcarpeta con el mismo nombre que otro candidato (ej. dos carpetas
"PRINCIPAL" de dos procesos distintos, o dos archivos "01. INFORME 1")
sin ser el mismo documento; en ese caso **nunca se reemplaza uno con el
otro** -- el que llega se guarda con un sufijo (ej. "PRINCIPAL_2") para
quedarse con ambos, y nunca se borra nada que ya estuviera.
La diferencia con el caso 1 es que estos quedan marcados aparte en
`faltantes_descargados_a_validar.csv` (una fila por cada candidato que
se fusionó), para que después confirmes que todo corresponde al mismo
proceso y borres a mano lo que no corresponda (el script nunca borra
nada por su cuenta). Si lo que encuentra es un archivo suelto (no una
carpeta), busca la carpeta que lo contiene. Si esa carpeta es "propia"
del caso (su nombre menciona el radicado, o la búsqueda encontró
directamente la carpeta) descarga la carpeta completa. Pero si es una
carpeta **genérica** -- de "informes" o "actuaciones" que junta
documentos de varios procesos, y el archivo que coincidió es solo uno
más ahí adentro -- **no** descarga la carpeta completa (traería folios
de otros procesos sin relación); solo baja los archivos de esa carpeta
que de verdad mencionen el radicado (y, si aplica, al demandante ESSA).

**Orden cronológico automático:** al terminar de descargar/fusionar
cada proceso, sus documentos quedan ordenados por FECHA y numerados
`1. `, `2. `, etc (el más viejo primero). La fecha se busca, en este
orden:

1. En el NOMBRE del archivo -- reconoce varios formatos: ISO
   (`2023-07-24`), `DD/MM/AAAA`, `"24 de julio de 2023"`, `"24 JULIO
   2023"` (sin la palabra "de" -- el formato más común en los nombres
   reales de autos/providencias) y `"Jul 24, 2023"`.
2. Si el nombre no trae fecha, se abren sus primeras páginas (PDF/DOCX)
   para buscarla ahí -- pero como máximo en 20 archivos por carpeta,
   para no quedarse abriendo PDF tras PDF en una carpeta grande (ej.
   fusionada de varios candidatos, con 70+ archivos).
3. Si tampoco hay fecha en el contenido, como último recurso se usa la
   fecha de modificación del propio archivo en el disco -- pero SOLO si
   es de más de un día atrás (para no confundir la hora en que se
   acaba de copiar/extraer con una fecha real). Un documento que llega
   por **adjunto de correo** y no trae fecha propia queda con la fecha
   del correo puesta como fecha de modificación, así que ese es el
   valor que se usa aquí -- mejor guiarse por cuándo llegó el correo
   que dejarlo sin ningún orden.

Los documentos sin NINGUNA fecha reconocible (ni nombre, ni contenido,
ni fecha de correo) quedan al final, en el orden en que ya estaban. Si
el proceso está repartido en subcarpetas (ej. "principal" y "anexos"),
cada una se ordena por su cuenta, sin mezclar los documentos de una con
los de otra. Si vuelves a correr el script y llega un documento más
viejo que los demás, el orden se recalcula solo. Además de las
carpetas que se acaban de descargar/fusionar en esta corrida, **al
empezar cada corrida también se revisan y ordenan TODAS las carpetas
de proceso que ya existan en el disco** (de esta corrida o de
cualquier corrida anterior) -- no hace falta que el proceso se haya
tocado hoy para que quede numerado.

**Solo se descargan archivos PDF:** cualquier otro tipo de archivo
(Word, Excel, imágenes, etc) que aparezca junto a los PDF -- ya sea
dentro de una carpeta de Drive o de un adjunto `.zip` de correo -- se
omite y **nunca** se baja al disco. La única excepción son los
archivos nativos de Google (Doc, Sheet, Slide), que sí se descargan
pero **exportados como PDF** (Drive los convierte automáticamente al
exportarlos, así que igual terminan como `.pdf` en el disco).

**Si un archivo puntual falla al descargarlo** (ruta demasiado larga
para Windows, permisos, antivirus, un corte de red momentáneo), el
script lo salta con una advertencia en el log y sigue con el resto --
un solo archivo problemático nunca detiene la descarga de ese proceso,
ni mucho menos el resto de la lista. Lo mismo si un proceso entero
falla por algo inesperado: se registra el error y se sigue con el
siguiente, en vez de detener toda la corrida a mitad de camino.

**Carpetas temporales sueltas:** lo primero que hace el script al
empezar es revisar si quedaron carpetas `_tmp_fusion_...`,
`_tmp_extraccion_correo_...` (u otras sueltas dentro de
`_tmp_extraccion`) de una corrida anterior que se cerró a la mitad, o
que no se pudieron borrar solas (por ejemplo, un archivo adentro
bloqueado por el antivirus o por OneDrive justo en ese momento). Como
nunca se sabe con certeza si todo su contenido ya quedó copiado en su
carpeta final, **nunca se borran solas**: se mueven a
`Duplicados_para_revisar` para que las revises tú.

**Limpieza automática de corridas anteriores:** al empezar, el script
también revisa si ya quedaron carpetas "_2", "_3", etc en el disco de
corridas viejas (por ejemplo, de antes de que existiera el filtro de
demandante ESSA). Revalida cada una: si menciona a ESSA (por nombre o
contenido), la fusiona dentro de su carpeta principal (sin perder
archivos con nombres repetidos); si no la menciona, la mueve tal cual
a `Duplicados_para_revisar` -- nunca la borra -- para que la revises a
mano, porque probablemente es ruido de otro proceso que compartía
cuenta o radicado corto. **Importante:** esto solo aplica cuando de
verdad hay una carpeta vieja con sufijo `"_N"` -- si el mismo radicado
aparece con **dos o más números de proceso distintos y legítimos**
(porque el Excel tiene ese radicado repetido con varios números, y
`validar_renombrar_carpetas.py` a propósito dejó una carpeta separada
por cada uno), esas carpetas **no se tocan ni se fusionan entre sí**:
fusionarlas destruiría esa separación intencional.

**Revisión de contaminación entre procesos:** también al empezar,
revisa TODAS las carpetas ya descargadas (de cualquier corrida,
incluidas las de hace tiempo) buscando archivos cuyo nombre mencione
un radicado corto **distinto** al de su propia carpeta -- rastro de un
bug ya corregido donde, por ejemplo, `"2023-24"` se confundía con
`"2023-244"` (otro proceso, solo comparte el prefijo numérico) y
terminaban documentos de un caso mezclados en la carpeta del otro. Un
archivo solo se mueve si menciona OTRO radicado y **nunca** el propio
(si cita ambos, se asume que es un documento legítimo que solo
referencia un caso relacionado, y no se toca). Los archivos
sospechosos se mueven a `Duplicados_para_revisar` -- nunca se borran.
Si al sacarlos una carpeta queda completamente vacía (todo su
contenido era de otro proceso), esa carpeta vacía sí se borra, para
que la próxima corrida la vuelva a buscar desde cero con el filtro ya
corregido.

**Revisión de demandado entre procesos:** de la misma forma, también
al empezar, revisa TODAS las carpetas ya descargadas buscando archivos
con un `"CONTRA <algo>"` que no corresponda al **demandado real** del
proceso según el Excel -- mismo rastro que la revisión de
contaminación, pero para el caso de dos procesos que comparten
radicado corto/cuenta y van contra demandados **distintos**. Esos
archivos también se mueven a `Duplicados_para_revisar` -- nunca se
borran. Si es la **carpeta misma** la que parece tener el demandado
equivocado en su nombre, solo se avisa en el log (no se mueve ni
renombra la carpeta sola) para que la revises a mano -- podría
significar que quedó cruzada con el radicado equivocado desde el
principio. Si el Excel no tiene la columna `DEMANDADO`, esta revisión
simplemente se omite.

### Configurar el acceso a Google Drive (una sola vez)

1. Ve a [console.cloud.google.com](https://console.cloud.google.com/) y
   crea un proyecto nuevo (o usa uno existente) -- es gratis.
2. En el menú, ve a **APIs y servicios → Biblioteca**, busca **Google
   Drive API** y dale **Habilitar**.
3. Ve a **APIs y servicios → Credenciales → Crear credenciales → ID de
   cliente de OAuth**. Si te pide configurar antes la "pantalla de
   consentimiento", elige tipo **Externo**, pon cualquier nombre, y
   agrégate a ti misma como "usuario de prueba" (no hace falta publicarla).
4. Tipo de aplicación: **Aplicación de escritorio**. Créala y descarga el
   JSON.
5. Renombra ese archivo a `credenciales_drive.json` y ponlo en la misma
   carpeta que los demás scripts.
6. La primera vez que corras `buscar_faltantes_en_drive.py`, se abre el
   navegador pidiendo que autorices el acceso con tu cuenta de Google.
   Acepta, y queda guardado `token_drive.json` para las próximas veces
   (no hay que repetir esto).

Para que también busque en tu correo, usa el mismo `credenciales_sgde.txt`
que ya tienes configurado para `procesos_juridicos.py` -- si no existe,
esa búsqueda simplemente se omite.

Uso:

1. Corre primero `iniciar.bat` (o `validar_renombrar_carpetas.py`) para
   que `procesos_faltantes_en_disco.csv` esté al día.
2. Haz doble clic en `buscar_drive.bat` (o corre `python
   buscar_faltantes_en_drive.py` desde una terminal en esa carpeta).
3. Por defecto corre en `MODO_PRUEBA = True` (solo busca y te dice qué
   descargaría). Revisa el log, y cuando confíes en el resultado cambia
   `MODO_PRUEBA = False` para descargar de verdad. Después de correrlo
   revisa `faltantes_descargados_a_validar.csv`: son las carpetas que se
   descargaron por una coincidencia menos segura (radicado corto, cuenta,
   o un enlace de correo sin el radicado completo) -- confírmalas y
   borra a mano las que no correspondan.

Notas sobre las búsquedas por radicado corto/cuenta (las menos
confiables): Google Drive no busca por texto exacto, busca por
*prefijo de palabra* (buscar `"2014-26"` también puede traer `"26
julio"`) -- el script filtra esos falsos positivos solos antes de
descargar nada. Ese mismo filtro también evita que un radicado corto
como `"2023-24"` se confunda con uno DISTINTO que solo comparte el
prefijo (ej. `"2023-244"` o `"2023-248"`) -- si justo al lado de la
coincidencia hay otro dígito, no cuenta como coincidencia real. Además,
antes de descargar cualquier candidato que solo coincidió por radicado
corto o cuenta, se verifican dos cosas:

- Que la carpeta candidata (o alguno de sus archivos) de verdad
  mencione ESE radicado -- así, si el mismo número de cuenta aparece en
  varios procesos distintos del mismo cliente a lo largo de los años,
  solo se descarga la carpeta que en realidad corresponde a este caso.
- Que el demandante sea **ESSA/Electrificadora de Santander** -- se
  revisa primero el nombre de la carpeta y de sus archivos; si ninguno
  lo dice, se abre el contenido de hasta 5 PDF/DOCX como muestra. Así,
  si la cuenta o el radicado corto coincide con un proceso de OTRO
  cliente, esa carpeta no se descarga.

Si un proceso ya tiene una carpeta en el disco (por ejemplo porque una
corrida anterior ya lo descargó), el script lo omite por completo sin
buscar ni descargar nada -- para no crear carpetas "_2" duplicadas si
se vuelve a correr sobre un `procesos_faltantes_en_disco.csv`
desactualizado.

## Validar solo las carpetas de la lista de faltantes

`validar_procesos_faltantes.py` es un script aparte, más liviano, para
cuando solo quieres revisar rápido las carpetas de los procesos que
están en `procesos_faltantes_en_disco.csv` (los que ya bajó
`buscar_faltantes_en_drive.py`, u otros que hayas agregado a mano) --
sin esperar a que se revise el disco completo.

Además de `procesos_faltantes_en_disco.csv`, también cruza contra
`carpetas_vacias.csv` (el otro reporte que genera
`validar_renombrar_carpetas.py`): si una carpeta de la lista ya existe
en el disco pero está **completamente vacía** (sin ningún archivo
adentro), se avisa aparte con claridad -- incluyendo si ese reporte ya
había encontrado un `.zip` pendiente en Descargas para ella -- en vez
de tratarla en silencio como si tuviera contenido para ordenar. Una
carpeta vacía **no se toca** (no hay nada que ordenar ni que borrar en
ella).

Para cada proceso de esa lista que **ya tenga carpeta con contenido**
en el disco:

**Por defecto** (`BORRAR_ARCHIVOS_DE_OTRO_PROCESO = False`), lo
**único** que hace es **ordenar cronológicamente** los documentos que
YA están adentro de cada carpeta, numerándolos `1. `, `2. `, etc -- no
mueve ni borra nada, no fusiona nada, no toca ninguna otra carpeta del
disco.

Si además quieres que **borre** los archivos que parezcan de OTRO
proceso -- porque mencionan un radicado corto distinto al de su propia
carpeta, o porque tienen un `"CONTRA <algo>"` que no corresponde al
demandado real según el Excel -- pon
`BORRAR_ARCHIVOS_DE_OTRO_PROCESO = True` al inicio del script.

**⚠️ A diferencia de TODOS los demás scripts de este proyecto** (que
nunca borran nada, solo mueven a `Duplicados_para_revisar`), con
`BORRAR_ARCHIVOS_DE_OTRO_PROCESO = True` los archivos que no coincidan
se **borran de forma permanente e irreversible** -- no quedan en
`Duplicados_para_revisar`, no se pueden recuperar. Esto se pidió así a
propósito, para no acumular carpetas de revisión manual; revisa con
calma el reporte en `MODO_PRUEBA = True` antes de correrlo con
`MODO_PRUEBA = False`. Lo que sí se mantiene igual que en el resto del
proyecto es que la **carpeta en sí nunca se mueve, renombra, ni se
fusiona con otra** -- solo se borran archivos puntuales adentro.

Los procesos de la lista que **todavía no tengan carpeta** en el disco
se cuentan aparte y se omiten -- este script **no descarga nada** (para
eso está `buscar_faltantes_en_drive.py`). Las carpetas que **no** estén
en la lista de faltantes tampoco se tocan, aunque tengan el mismo tipo
de problemas -- este script es deliberadamente angosto, solo mira la
lista.

Uso:

1. Corre primero `validar_renombrar_carpetas.py` para que
   `procesos_faltantes_en_disco.csv` esté al día.
2. Haz doble clic en `validar_faltantes.bat` (o corre `python
   validar_procesos_faltantes.py` desde una terminal en esa carpeta).
3. Por defecto corre en `MODO_PRUEBA = True` (solo revisa y te dice qué
   ordenaría/borraría). Cambia `MODO_PRUEBA = False` al inicio del
   script para aplicar los cambios de verdad.

Todo queda registrado en `validar_procesos_faltantes.log`.

## Crear carpetas de procesos terminados/remitidos/no iniciados

`validar_renombrar_carpetas.py` solo organiza los procesos en los
estados de `ESTADOS_A_CONTAR` (`ACTIVO`, `ACTIVOS CON TITULOS`,
`SUSPENDIDO`, `REORGANIZACION`). Para los procesos en cualquier OTRO
estado -- terminados (por pago, por auto, por prepago, con contrato,
etc), remitidos (a castigo, a prepago, etc), o que nunca llegaron a
iniciarse -- usa `crear_carpetas_terminados_castigo.py`.

Estos procesos normalmente no tienen (o no importa) un radicado real
para organizar documentos -- son más un registro administrativo que un
expediente judicial activo -- así que la carpeta **no** se nombra
`"numero. radicado"` como el resto del proyecto, sino:

- Si el `ESTADO PROCESAL` empieza con **`TERMINADO`** (por pago, por
  auto, por prepago, con contrato, etc): `"<numero>. <ESTADO PROCESAL
  EXACTO del Excel>"` -- ej. `"123. TERMINADO POR AUTO"`.
- Si el `ESTADO PROCESAL` empieza con **`REMITIDA`** (a castigo, a
  prepago, etc) o con **`NO INICIO`**: **solo el número** -- ej.
  `"145."`.

Los procesos con un `ESTADO PROCESAL` que no encaje en ninguna de esas
dos reglas (ej. `"DESISTIMIENTO DE PRETENSIONES"`, `"DEVUELTA INCURRIO
EN GASTOS"`) se dejan **fuera a propósito** -- no se crea carpeta para
ellos, y quedan listados en el log para que decidas qué hacer.

Si **ya existe** una carpeta en el disco para ese número (cualquier
nombre que empiece por `"<numero>. "` o sea exactamente `"<numero>."`
-- por ejemplo porque el proceso ya se organizó de la forma normal con
su radicado real), **no se crea una nueva**. Este script nunca borra,
renombra ni mueve nada -- solo **crea** carpetas vacías nuevas donde
todavía no exista ninguna para ese número. Correrlo varias veces no
duplica nada.

Uso:

1. Haz doble clic en `crear_carpetas_terminados_castigo.bat` (o corre
   `python crear_carpetas_terminados_castigo.py`).
2. Por defecto corre en `MODO_PRUEBA = True` (solo revisa y te dice qué
   carpetas crearía). Cambia `MODO_PRUEBA = False` al inicio del script
   para crearlas de verdad.

Usa la misma configuración (`RUTA_EXCEL`, `CARPETA_PROCESOS`, etc) que
`validar_renombrar_carpetas.py` -- no hay que configurarla dos veces.
Todo queda registrado en `crear_carpetas_terminados_castigo.log`.

### Por qué el total de carpetas puede no coincidir con el total de filas del Excel

Además de crear, el script audita **todas** las filas del Excel (no
solo las que va a crear) para explicar cualquier diferencia entre el
número de filas del Excel y el número de carpetas en el disco. Cada
número de proceso sin carpeta cae en una de estas categorías, y todas
quedan reportadas en el log con su fila y detalle:

- **Número duplicado en el Excel**: el mismo número de proceso aparece
  en más de una fila (con radicado y/o estado distintos). No se crea
  ni se toca nada -- hay que corregir el Excel a mano, porque no se
  puede adivinar cuál fila es la correcta.
- **Activo/Suspendido/Reorganización con radicado**: ya está
  rastreado en el listado de procesos faltantes
  (`procesos_faltantes_en_disco.csv`) -- a propósito **no** se crea
  una carpeta vacía aquí, porque ese proceso lo organiza
  `buscar_faltantes_en_drive.py` buscando su contenido real. Si
  quieres esa carpeta, corre ese script en vez de este.
- **Activo/Suspendido/Reorganización sin radicado**: el radicado
  todavía no está diligenciado en el Excel -- hay que llenarlo antes
  de poder buscar o crear nada para ese proceso.
- **Sin estado procesal diligenciado**: la fila tiene número de
  proceso pero la columna `ESTADO PROCESAL` está vacía todavía.
- **Estado sin clasificar**: no empieza con `TERMINADO`, `REMITIDA` ni
  `NO INICIO` (ver arriba).

### Deshacer: borrar_carpetas_terminados_castigo.py

Si quieres borrar las carpetas vacías de procesos que no son activos,
usa `borrar_carpetas_terminados_castigo.py` (o su iniciador
`borrar_carpetas_terminados_castigo.bat`):

- Para cada número de proceso que **no** esté en `ESTADOS_A_CONTAR`
  (o sea, cualquiera que no sea `ACTIVO`, `ACTIVOS CON TITULOS`,
  `SUSPENDIDO` o `REORGANIZACION`) **y tenga algún `ESTADO PROCESAL`
  diligenciado**, borra su carpeta si está **completamente vacía** --
  sin importar cómo se llame exactamente (`"numero. ESTADO"`, solo
  `"numero."`, o incluso `"numero. radicado"`) ni si el estado es uno
  de los que clasifica `crear_carpetas_terminados_castigo.py`
  (`TERMINADO*`/`REMITIDA*`/`NO INICIO*`) o uno sin clasificar (ej.
  `DESISTIMIENTO`, `DEVUELTA`).
- **Nunca** toca procesos activos/suspendidos/reorganización, filas
  sin `ESTADO PROCESAL` diligenciado todavía, números duplicados en el
  Excel, ni ninguna carpeta que ya tenga contenido adentro (por
  ejemplo porque le agregaste documentos a mano) -- esas se reportan
  en el log para que decidas a mano.
- ⚠️ Borrar una carpeta es **irreversible**. Respeta `MODO_PRUEBA`
  (por defecto `True`): revisa el log primero, y solo cambia
  `MODO_PRUEBA = False` cuando estés segura de que la lista de
  carpetas a borrar es la correcta.

## Clasificar procesos ejecutivos (información no procesal + documento de terminación)

`clasificar_procesos_ejecutivos.py` (o su iniciador
`clasificar_ejecutivos.bat`) es una versión **especializada** de
`buscar_faltantes_en_drive.py` para el informe `3. CONTROL PROCESOS
EJECUTIVOS ESSA...xlsm`, leyendo directo de la hoja **`ACTIVOS`** (la
que se mantiene al día). **No** usa `DatosProcesados1` ni las demás
hojas `DatosProcesadosN` -- son copias aplanadas que no se actualizan
solas cuando editas `ACTIVOS`, así que pueden traer radicados/cuentas/
demandados desactualizados o de relleno. En `ACTIVOS`, un proceso
"acumulado" (varias cuentas bajo un mismo radicado) no repite el No./
ESTADO PROCESAL/RADICADO/JUZGADO en cada fila -- el script arrastra
esos datos hacia abajo automáticamente a las filas de continuación.
Trabaja fila por fila y, en vez de descargar todo el contenido
relacionado con el radicado, aplica dos reglas según el `ESTADO
PROCESAL` de cada una:

1. **Cualquier fila que no sea "terminada"** (ver punto 2) -- activos,
   suspendidos, en reorganización, remitidos a castigo/prepago, etc: la
   carpeta se nombra `"<numero>. <radicado>"` si la fila ya tiene un
   radicado de 23 dígitos válido, o `"<numero>. <ESTADO PROCESAL>"`
   (igual que el punto 2) si todavía no lo tiene. **Solo** se sube
   información **no procesal**: tutelas, derechos de petición, o pagos
   oficiosos -- **nada más** (ni demandas, ni solicitudes de
   conciliación, ni memoriales pidiendo requerir a alguien, aunque
   mencionen esas palabras de pasada). La clasificación es estricta a
   propósito, en este orden:
   1. Si el **nombre** del archivo/asunto ya lo dice (ej. "DERECHO DE
      PETICION ADRESS.pdf"), se acepta directo -- es la señal más
      confiable.
   2. Si el nombre trae una marca clara de ser un documento procesal
      (demanda, memorial, solicitud, mandamiento, recurso, traslado,
      etc), se **descarta**, aunque el contenido mencione de pasada
      una tutela/petición/pago oficioso (ej. una demanda que narra en
      su historial "el demandado interpuso una tutela"). A propósito
      esta lista NO incluye la palabra "contestación" sola -- una
      **respuesta** de un municipio/departamento a un derecho de
      petición o tutela también suele llamarse así, y esas respuestas
      SÍ son información no procesal que hay que descargar (la
      "contestación de la demanda", que sí debe excluirse, ya queda
      cubierta porque menciona "demanda").
   3. Si no hay ninguna marca clara en el nombre, se acepta solo si la
      frase aparece cerca del **inicio** del contenido (el propio
      encabezado/título del documento), no en cualquier parte de un
      PDF de varias páginas.

2. **Terminados por pago, por auto, por contrato/prepago, o que nunca
   se presentaron** (`ESTADO PROCESAL` que empieza con `TERMINADO` o
   `NO INICIO` -- igual que `crear_carpetas_terminados_castigo.py`): la
   carpeta se nombra `"<numero>. <ESTADO PROCESAL EXACTO del Excel>"`
   (ej. `"245. TERMINADO POR AUTO"`), y **solo** se sube el documento
   que deja constancia de que el proceso **no sigue su curso**
   (termina por pago, acepta el retiro de la demanda, decreta
   desistimiento/archivo, etc) -- no hace falta que diga literalmente
   "auto". Si no se encuentra ese documento, el proceso queda listado
   en `terminados_sin_auto_pendientes.csv` para que lo descargues a
   mano.

Con estas dos reglas, **toda** fila del Excel con `ESTADO PROCESAL`
diligenciado recibe una carpeta (solo quedan sin carpeta las filas sin
ese dato todavía).

**Procesos acumulados**: el Excel repite el mismo número de proceso en
varias filas cuando agrupa varias cuentas bajo un mismo radicado
(proceso "acumulado"). Esas filas se **fusionan en una sola carpeta**
(una carpeta por radicado, no una por cuenta) -- sus cuentas y
demandados se juntan para la búsqueda/validación. Solo cuando el mismo
número de proceso tiene un radicado **distinto** en cada fila (poco
frecuente -- numeración administrativa repetida por error, no es el
mismo expediente) sí salen carpetas separadas, porque ahí el nombre
(que incluye el radicado) ya es distinto para cada una.

**Búsqueda en Gmail** (además de Drive): sí es posible escanear **todo
el correo** (no solo la bandeja de entrada) para esto, y el script ya
lo hace si `BUSCAR_EN_CORREO = True` (por defecto) y existe
`credenciales_sgde.txt` (las mismas credenciales que ya usa
`procesos_juridicos.py`/`buscar_faltantes_en_drive.py`, ver más abajo).

A diferencia de Drive (donde sí se busca por radicado/radicado corto/
cuenta de cada proceso, uno por uno), en Gmail la búsqueda es **al
revés**: se busca **una sola vez para toda la corrida** por tutela/
derecho de petición/pago oficioso directamente (no por radicado), y
**cada correo encontrado se empareja después con el proceso correcto**
si coincide su radicado, su cuenta, **o** el demandado (con que
coincida cualquiera de los tres alcanza -- no hace falta que coincidan
los tres). Esto encuentra correos que una búsqueda por radicado se
perdería (ej. un derecho de petición que en el cuerpo solo menciona el
nombre del demandado o la cuenta, no el radicado exacto). Para el
demandado se exige que coincidan **todas** sus palabras significativas
(nombre y apellido, no solo una) y que sean al menos dos -- un nombre
mal diligenciado en el Excel que se reduzca a una sola palabra genérica
no se usa para cruzar, y además esas palabras deben aparecer como
palabra COMPLETA en el texto (no pegadas dentro de otra palabra más
larga, ej. "AUTO" no cuenta si aparece dentro de
"AUTOTERMINAPROCESO.pdf"). Tampoco cuentan como palabra significativa
los términos administrativos/judiciales genéricos que aparecen por
igual en miles de documentos sin relación entre sí (ej. "auto",
"demanda", "medida", "cautelar", "anexos", "oficio" -- ver
`PALABRAS_GENERICAS_DEMANDADO`): sin estos dos filtros, un demandado
mal diligenciado (una frase administrativa en vez de un nombre real)
generaba coincidencias masivas y completamente falsas. Además, antes
de buscar se **quita el membrete del juzgado** (su propio nombre y
dirección/correo del despacho) **y la fecha de cierre/firma** (ver
`_quitar_membrete_juzgado`) -- ambos son literalmente el principio y el
final de cualquier auto, y casi siempre incluyen el nombre del
**municipio** donde queda el juzgado (ej. "JUZGADO PROMISCUO MUNICIPAL
DE CANTAGALLO, BOLIVAR" al inicio, y "CANTAGALLO, DIEZ (10) DE JULIO DE
2025" al firmar), que por pura coincidencia puede ser el demandado de
OTRO proceso sin ninguna relación con ese documento. El membrete se
reconoce aunque el PDF lo parta a la mitad por su propio salto de línea
visual (ej. "...Municipio de Puerto\nWilches (S)." -- el nombre del
municipio quedaría partido en dos renglones del texto extraído, pero
igual se reconoce completo, sin confundir esa unión con un campo real
como "DEMANDADO:" que venga justo debajo). Ya sin esas dos
cosas, el demandado solo se
busca en los primeros `VENTANA_DEMANDADO_CARACTERES` del texto que
queda (el encabezado/carátula real, donde Colombia identifica las
partes de cualquier proceso) -- **no** en el documento completo. El
radicado y la cuenta sí se buscan en el documento completo, porque esas
coincidencias ya son precisas por sí solas -- y reconocen el radicado
tanto **plano** (23 dígitos seguidos) como escrito **con guiones,
puntos, o espacios** entre sus grupos (ej.
"68001-40-03-001-2024-00050-00"), **o incluso con un separador entre
CADA dígito** (ej. "6 8 0 0 1 4 0 0..."), típico de un sello o tabla
escaneada donde el PDF extrae cada dígito con un espacio de más -- ver
`_radicados_en_texto`.
Un mismo correo puede terminar adjuntado a más de un proceso si aplica
a varios (ej. un proceso "acumulado" con varias cuentas). Igual que en
Drive, siempre se exige además que el correo mencione a ESSA/
Electrificadora de Santander.

Si el correo clasifica y se emparejó con algún proceso: se guardan sus
adjuntos PDF/DOCX (extrayendo los que vengan dentro de un `.zip`), o si
no trae ningún adjunto útil, se guarda el asunto + cuerpo como un
`.txt` simple para no perder la información.

La clasificación de "información no procesal" y de "documento que
termina el proceso" es por **palabras clave** (nombre/asunto y, si es
PDF/DOCX o el cuerpo de un correo, su contenido) -- es una heurística,
no perfecta. Cada decisión queda registrada en
`clasificar_procesos_ejecutivos.log` para que la revises y ajustes las
listas de palabras clave al inicio del script si hace falta
(`PALABRAS_TIPO_INFORMACION_FUERTES`, `PALABRAS_PAGO_OFICIOSO`,
`PALABRAS_PROCESO_NO_CONTINUA`).

Reutiliza toda la infraestructura de `buscar_faltantes_en_drive.py`
(las mismas credenciales `credenciales_drive.json`/`token_drive.json`,
la búsqueda por radicado/radicado corto/cuenta, y la validación
obligatoria de que el documento sea de ESSA y del demandado correcto).

**No repite trabajo ya hecho**: cada proceso que termina de revisarse
por completo (Drive, y si encontró lo que buscaba) queda registrado en
`clasificar_procesos_ejecutivos_revisados.txt`. Si vuelves a correr el
script (por ejemplo porque lo interrumpiste, o simplemente para
capturar lo nuevo), esos procesos se **omiten** -- ni Drive ni Gmail --
en vez de volver a revisarlos desde cero. Un proceso "terminado" que se
quedó pendiente (no se encontró el documento que lo termina) **no**
queda marcado, así que se reintenta en cada corrida hasta que aparezca.
Si quieres forzar que se revise todo de nuevo, borra ese archivo.

**Velocidad**: casi todo el tiempo se va esperando la respuesta de
Google Drive (red), no procesando en tu computador -- por eso el script
busca varios procesos **al mismo tiempo** (`NUM_HILOS`, por defecto 8 a
la vez) en vez de uno por uno, lo que reduce el tiempo total casi en esa
misma proporción. Además, cuando el **nombre** de un archivo ya alcanza
para descartarlo (ej. `"DEMANDA EJECUTIVA.pdf"`), ni siquiera se
descarga ni se lee su contenido -- eso evita el paso más lento
(descargar + extraer texto de PDF/DOCX) para la mayoría de los
documentos de cada carpeta, que normalmente son procesales y no
información no procesal. Si tu internet aguanta y quieres que vaya más
rápido todavía, puedes subir `NUM_HILOS` (ej. `15` o `20`); si prefieres
verlo avanzar de a uno (más fácil de leer en el log, o si notas errores
de conexión), bájalo a `1`.

Antes de usarlo, edita al inicio del script:

- `RUTA_EXCEL_CONTROL`: ruta a `3. CONTROL PROCESOS EJECUTIVOS
  ESSA...xlsm`.
- `HOJA_EXCEL_CONTROL` / columnas (`COLUMNA_NO`, `COLUMNA_ESTADO`,
  etc): solo si tu Excel usa otros nombres de hoja/columna.
- `CARPETA_PROCESOS`: dónde se crean las carpetas de cada proceso. A
  diferencia del resto del proyecto (que usa el disco duro externo
  detectado por `ETIQUETA_DISCO_EXTERNO` en `validar_renombrar_carpetas.py`),
  este script tiene su **propio** destino fijo -- no depende de ningún
  disco externo.
- `BUSCAR_EN_CORREO`: ponlo en `False` si no quieres que también
  busque en Gmail (solo Drive).
- `NUM_HILOS`: cuántos procesos se buscan en Drive al mismo tiempo (por
  defecto `8`). Súbelo si quieres que vaya más rápido y tu internet
  aguanta, o bájalo a `1` para verlo avanzar de a uno.

Las filas sin `ESTADO PROCESAL` diligenciado todavía no se pueden
organizar -- quedan reportadas en el log.

Respeta `MODO_PRUEBA` (por defecto `True`): en modo prueba solo busca y
clasifica, mostrando qué subiría y a qué carpeta, sin crear carpetas ni
descargar nada todavía.

## Limpiar y terminar de nombrar las carpetas de procesos ejecutivos

`limpiar_carpetas_procesos_ejecutivos.py` (o su iniciador
`limpiar_carpetas_ejecutivos.bat`) hace dos limpiezas en la misma
`CARPETA_PROCESOS` que usa `clasificar_procesos_ejecutivos.py` (reutiliza
la misma lectura del Excel, mismas reglas de nombre):

1. **Consolida**, por número de proceso, cualquier carpeta que haya
   quedado suelta o vieja frente al nombre que le corresponde **hoy**
   según el Excel -- `"<numero>. <radicado>"` si ya lo tiene, o
   `"<numero>. <ESTADO PROCESAL>"` si no, tanto para activos como para
   terminados. Cubre dos casos:
   - Carpetas con **solo el número** (ej. `"245."` o `"245"`, sin
     radicado ni ESTADO PROCESAL).
   - Carpetas con un nombre **viejo que ya no corresponde** (ej.
     `"1055. ACTIVO"` de cuando el proceso todavía no tenía radicado, y
     ya existe `"1055. <radicado>"`; o `"1051. ACTIVO"` de cuando el
     proceso seguía activo, y ya existe `"1051. TERMINADO POR PAGO"`
     porque cambió de estado).

   Si la carpeta con el nombre correcto **ya existe**, la vieja se
   **borra** (si está completamente vacía) o se **reporta** para
   revisión manual (si todavía tiene contenido adentro -- nunca se
   mezcla solo). Si la correcta **no existe todavía**, la vieja se
   **renombra** directo a la correcta (con o sin contenido, no se
   pierde nada). Si el número es ambiguo (radicados distintos en el
   Excel) o hay más de una carpeta vieja candidata sin forma segura de
   saber cuál es, se omite y se reporta para que lo revises a mano.
2. **Borra** las carpetas de procesos **activos** (todo lo que no es
   `TERMINADO*`/`NO INICIO` -- activo, suspendido, en reorganización,
   remitida a castigo/prepago, etc) que estén **completamente vacías**
   (sin ningún archivo adentro, ni en subcarpetas). Los procesos
   **terminados nunca se tocan aquí, estén vacíos o no** -- para esos ya
   existe `terminados_sin_auto_pendientes.csv`, que es la lista correcta
   de qué falta descargar a mano.

⚠️ Borrar una carpeta es **irreversible**. Respeta `MODO_PRUEBA` (por
defecto `True`): revisa el log primero, y solo cambia `MODO_PRUEBA =
False` cuando estés segura de que la lista de carpetas a borrar/renombrar
es la correcta.

## Clasificar información extraprocesal por proceso (demandado, radicado o cuenta)

`clasificar_por_demandado.py` (o su iniciador
`clasificar_por_demandado.bat`) reparte información extraprocesal a la
carpeta del proceso que le corresponde, en dos pasos. **El paso 2
(Gmail) está desactivado por defecto** (`BUSCAR_EN_CORREO = False`,
ver más abajo) -- el paso 1 (carpeta de descargas) siempre corre solo,
rápido y sin depender de que Gmail responda bien desde tu red/equipo.

1. **Carpeta de descargas manuales** (`CARPETA_DESCARGAS_MANUAL`, por
   defecto tu carpeta "Downloads"): revisa cada PDF/DOCX que haya ahí y
   lo **mueve** directo a la carpeta del proceso que coincide -- ej. si
   el demandado es "ALBERTO SUAREZ" y el archivo lo menciona, se mueve
   a `"<numero>. <radicado o ESTADO>"`. No hay restricción de tipo de
   documento aquí -- se asume que ya es información extraprocesal
   porque tú la descargaste a propósito, y por eso este paso busca
   entre **todos** los procesos del Excel, activos y **terminados**
   (a diferencia del paso 2, que solo busca entre los activos) --
   también puedes soltar aquí el auto que termina un proceso ya
   terminado (ej. un desistimiento tácito). Por defecto
   (`SOLO_DESCARGAS_DE_HOY = True`) **solo revisa lo creado o
   modificado hoy** -- una carpeta de Descargas normal acumula años de
   archivos de todo tipo (demandas, autos viejos, etc, no solo lo que
   bajaste hoy); ponlo en `False` si alguna vez quieres que revise toda
   la carpeta sin importar la fecha.
2. **Gmail** (`BUSCAR_EN_CORREO`, **`False` por defecto** -- ponlo en
   `True` si quieres activarlo, o si ya confirmaste que tu red/antivirus
   no bloquea las búsquedas IMAP; ver la nota de la búsqueda envuelta
   en un solo argumento y de a un término más abajo): busca en **todo**
   tu correo (una sola conexión para todos los términos -- no una
   conexión nueva por cada uno, para que no tarde horas con miles de
   términos) el nombre de cada demandado,
   el radicado (y sus formas cortas), y la cuenta de cada proceso. Del
   resultado **solo descarga** lo que además sea un derecho de
   petición, una tutela, o un pago oficioso -- **tanto lo presentado
   como las respuestas** que da el municipio/departamento al que se
   envió. Igual de riguroso que el resto del proyecto: si el nombre ya
   trae una marca de documento procesal (demanda, memorial, etc), se
   descarta aunque mencione la tutela/petición de pasada; también
   exige que el correo mencione a ESSA/Electrificadora de Santander.

   Para poder buscar necesita seleccionar la carpeta de Gmail que
   contiene **todos** los correos. Si tu cuenta de Gmail está en
   español (o cualquier idioma que no sea inglés), esa carpeta no se
   llama `"[Gmail]/All Mail"` sino su traducción (ej. `"[Gmail]/Todos"`)
   -- el programa primero intenta el nombre en inglés y, si no existe,
   la encuentra automáticamente buscando la carpeta especial marcada
   como "todos los correos" (sin importar cómo se llame en tu idioma),
   así que no tienes que cambiar nada a mano.

   Busca **de a un término por vez** (`TAMANO_LOTE_CORREO = 1`, el
   valor por defecto) -- sin combinar varios términos en una sola
   consulta con `OR`. Se eligió así, a propósito, después de que las
   consultas combinadas dieran demasiados problemas reales distintos
   (tildes, límites de protocolo, la conexión colgándose) para
   confiar en ellas: buscar de a un término es el método más simple
   posible, exactamente el mismo patrón que ya usa sin problemas
   `buscar_en_correo()` en `buscar_faltantes_en_drive.py`. Es más
   lento que combinar términos (miles de búsquedas en vez de
   cientos), pero la conexión persistente + la reconexión automática
   (ver más abajo) hacen que sea viable igual con miles de términos.
   Si alguna vez hace falta más velocidad, `TAMANO_LOTE_CORREO` se
   puede subir para volver a combinar varios términos por búsqueda --
   la consulta completa (con paréntesis, la palabra `OR`, y comillas
   anidadas para frases exactas) se manda como **UN SOLO argumento
   entre comillas** de IMAP, escapando las comillas que ya trae
   adentro. Esto puede sonar como un detalle interno, pero fue el bug
   real de fondo que costó más encontrar: `mail.search()` de la
   librería de Python no agrega comillas por su cuenta, así que sin
   este envoltorio el servidor leía el paréntesis inicial y la palabra
   `OR` como **tokens sueltos de IMAP** en vez de como parte de un
   único texto -- eso producía "SEARCH command error: BAD Could not
   parse command" en **absolutamente todos los lotes combinados** (no
   solo los que tenían tildes), y tras suficientes errores seguidos
   Gmail terminaba cortando la conexión entera con "Too many protocol
   errors". Además, los términos se mandan **sin tildes/diacríticos**
   (ej. "PÉREZ" -> "PEREZ") -- Gmail busca igual sin distinguirlas, así
   que no se pierde ningún resultado, y así el texto es ASCII puro sin
   depender de CHARSET ni de "literales" de IMAP (un mecanismo que sí
   permite caracteres no-ASCII por protocolo, pero que demostró
   colgar la conexión sin ningún error en ciertas redes/antivirus: el
   intercambio "esperar la confirmación del servidor y mandar el texto
   aparte" que exige un literal es un patrón de tráfico poco común que
   algunos proxies no manejan bien). Si un lote puntual falla, se salta
   y sigue con los demás.

   Con cientos de procesos activos son cientos de búsquedas seguidas
   sobre la misma conexión, y **Gmail la corta** si la nota con
   demasiadas búsquedas/descargas seguidas en poco tiempo (protección
   propia de Gmail contra scripts, no depende de nada configurable acá
   -- no tiene que ver con limitar la búsqueda a los últimos años ni
   nada parecido). Cuando eso pasa (a mitad de una búsqueda o a mitad
   de la descarga de un correo puntual), el programa **reconecta solo**
   (login + volver a seleccionar la carpeta) y sigue justo donde se
   quedó, hasta `MAX_RECONEXIONES_CORREO` veces -- así la búsqueda
   completa los 100% de los términos en vez de quedarse solo con los
   que alcanzó a revisar antes del primer corte. Si después de agotar
   los reintentos la conexión sigue sin poder recuperarse, se rinde de
   forma prolija: nunca se pierden los correos que **ya se habían
   encontrado** en los lotes que sí funcionaron, así la conexión
   termine cortándose de forma definitiva. Si un lote puntual sigue
   fallando SIEMPRE, incluso con conexiones nuevas cada vez (no es la
   conexión, es algo de ESE lote en particular), después de
   `MAX_REINTENTOS_POR_LOTE` intentos se lo **salta** (avisando qué
   términos traía, para revisarlos a mano) en vez de gastar ahí todo el
   presupuesto de reconexiones que le hace falta al resto de los
   cientos de lotes que quedan por revisar.

   Toda operación con Gmail tiene un **límite de tiempo**
   (`TIMEOUT_CORREO_SEGUNDOS`, 30s por defecto), aplicado por
   **DUPLICADO**: el timeout normal del socket, y un límite duro
   independiente corriendo cada búsqueda en un hilo aparte. Esto porque
   el timeout del socket depende de que Python detecte "no llegó nada
   en N segundos" -- en algunos entornos (antivirus/proxy corporativo
   en el medio, ciertas particularidades de Windows) esa señal puede no
   llegarle nunca a Python, y ahí el programa se queda colgado DE
   VERDAD, sin ningún error ni progreso en el log, sin que el timeout
   del socket lo salve. El límite duro no depende de eso: corta la
   espera de todas formas después de `TIMEOUT_CORREO_SEGUNDOS`
   segundos, sin importar la causa, y fuerza el cierre de la conexión
   para intentar liberarla. Al reconectar, cada paso (conectar, iniciar
   sesión, seleccionar la carpeta) deja su propia línea en el log -- si
   alguna vez vuelve a "colgarse", el último mensaje que quede en el
   log dice exactamente en cuál paso se quedó, en vez de tener que
   adivinar.

   Con cientos de lotes por revisar, una búsqueda que va perfectamente
   bien puede pasar varios minutos SIN ninguna línea nueva en el log
   (simplemente toma tiempo, nada falla) -- y eso se ve exactamente
   igual que "el programa está colgado". Por eso ahora deja un aviso de
   progreso cada `AVISO_PROGRESO_CORREO_LOTES` lotes (10 por defecto):
   "...van N/M lotes revisados, X correo(s) encontrados hasta el
   momento...", para que quede claro que sigue avanzando aunque no haya
   pasado nada "que reportar".

En **ambos** pasos, a qué proceso corresponde un archivo/correo se
decide con la misma regla que usa `clasificar_procesos_ejecutivos.py`
para emparejar su búsqueda global de correo: coincide su **radicado**,
su **cuenta**, **o** el **demandado completo** (todas sus palabras
significativas -- no basta con una sola, ver
`MIN_PALABRAS_DEMANDADO_PARA_CRUZAR`) -- basta con que coincida
cualquiera de los tres, no hace falta que coincidan todos. No se limita
al nombre exacto del demandado: si un documento no lo menciona pero sí
su radicado o su cuenta, también se encuentra. Si un archivo/correo
coincide con **más de un** proceso (ej. dos procesos activos contra el
mismo demandado), **siempre** se intenta desambiguar por el
**radicado**: primero se busca el radicado **completo** (23 dígitos) de
cada candidato -- si aparece el de uno solo, se usa ese de inmediato,
aunque el texto también mencione, por pura coincidencia, la **forma
corta** (año-consecutivo) de otro candidato (dos juzgados distintos
pueden repetir el mismo año-consecutivo; solo cambia el código del
juzgado o la instancia, algo que la forma corta no distingue). Solo si
ningún candidato tiene su radicado completo en el texto se recurre a la
forma corta como último recurso. El intento (funcione o no) siempre
queda en el log con el prefijo `[Radicado]`,
para que sea claro que sí se revisó (un auto corto no siempre repite su
propio radicado en el texto, así que no siempre alcanza para
desambiguar). Si un archivo/correo
no coincide con ningún proceso activo, o sigue sin poder decidirse
entre varios, se deja intacto y queda registrado en un CSV para que lo
revises a mano -- nunca se adivina: `clasificar_por_demandado_sin_coincidencia.csv`
(archivos de la carpeta de descargas sin ningún proceso posible),
`clasificar_por_demandado_ambiguos.csv` (archivos con más de un proceso
posible -- incluye una columna "Por que coincidio cada uno" con el
motivo EXACTO de cada candidato: si fue por radicado, por cuenta, y/o
cuál demandado específico, con un pedazo del texto donde apareció su
nombre para ver el contexto real -- para poder confirmar la causa real
en vez de adivinarla), y
`clasificar_por_demandado_correo_sin_coincidencia.csv` (correos que sí
eran tutela/petición/pago oficioso pero no coincidieron con ningún
proceso activo conocido). Si YA se encuentra el proceso, la carpeta se
**crea sola** si todavía no existe -- no hace falta haberla creado
antes.

Para que sea lo más **rápido y preciso** posible, el paso 1 primero
prueba la coincidencia SOLO con el **nombre** del archivo (sin abrir
nada); solo si el nombre no basta se lee su contenido (PDF/DOCX), que
es lo más lento -- y es además más preciso, porque el contenido
completo de un documento largo puede coincidir con más procesos de los
que el nombre por sí solo sugeriría (nunca con menos).

Solo considera procesos **activos** (todo lo que no es terminado/no
inicio), igual que la regla de "información no procesal" del resto del
proyecto. Reutiliza la misma lectura del Excel, el mismo
`CARPETA_PROCESOS`, y el mismo clasificador de tipo de documento de
`clasificar_procesos_ejecutivos.py` (incluye pago oficioso, ver abajo),
y las mismas credenciales `credenciales_sgde.txt` para Gmail (ver más
abajo). Respeta `MODO_PRUEBA` (por defecto `True`).

## Revisar el correo en MODO PRO y repartirlo por proceso ⭐

`revisar_correo_pro.py` (o su iniciador `revisar_correo_pro.bat`) es la
forma **recomendada** de sacarle provecho al correo: revisa tu Gmail y
deja cada correo en la carpeta del proceso que le corresponde --
comparando contra el Excel por **demandado**, **radicado** o **número
de cuenta**. Lo que no coincide con ningún proceso **no se pierde**: se
guarda en `_SIN CLASIFICAR - REVISAR A MANO` (una subcarpeta por
correo) para que lo revises a mano.

### Por qué este sí funciona (y el paso 2 de `clasificar_por_demandado.py` no)

El planteamiento anterior le **preguntaba a Gmail una vez por cada
término del Excel**: "¿tienes correos que digan DAVID CAICEDO?", "¿y
68001400302120180078701?", "¿y 2018-00787?"... con ~650 procesos
activos eso son **más de 4000 búsquedas seguidas**. Ningún servidor de
correo aguanta eso de un mismo cliente -- de ahí venían *todos* los
síntomas: `Could not parse command`, `Too many protocol errors`,
`socket error: EOF`, y el programa colgado.

Este script **invierte** el planteamiento, que es lo único que lo
arregla de raíz:

| | Antes | Ahora |
|---|---|---|
| Preguntas a Gmail | **4242** (una por término) | **1** (una por fecha) |
| Dónde se compara | En el servidor de Gmail | **En tu PC**, en memoria |
| Si Gmail corta | Se pierde todo | Reconecta y sigue donde iba |

Es decir: **no se le pregunta nada a Gmail** sobre demandados,
radicados ni cuentas. Toda la comparación contra el Excel se hace
localmente. Comparar 4242 términos contra un texto es instantáneo en
memoria; lo imposible era preguntarlo 4242 veces por internet.

Lo único que se le pide a Gmail es acotar **por tipo**: unas 8
consultas (`SUBJECT` y `TEXT` de "PETICI", "TUTELA", "OFICIOSO",
"PQR" -- `FRASES_EXTRAPROCESALES`). En un caso real, 2 años de correo
eran **25.134 mensajes**: bajarlos todos completos satura la conexión y
tomaría decenas de corridas. Acotando así quedan unos cientos -- los
que de verdad interesan. Son trozos de palabra a propósito, para que
una sola sirva para todas las variantes sin depender de las tildes
(`PETICI` encuentra PETICION, PETICIÓN, PETICIONES y "respuesta a su
petición"). Si prefieres revisar *todos* los correos de la ventana,
pon `ACOTAR_POR_TIPO_EN_GMAIL = False` (será mucho más lento).

### Qué tiene de "modo pro"

No se conforma con el asunto. De cada correo revisa el **asunto**, el
**cuerpo**, el **nombre** de cada adjunto, y -- esto es lo importante
-- el **texto de adentro de cada PDF/DOCX adjunto**
(`LEER_TEXTO_DE_ADJUNTOS`). Un auto casi nunca trae el radicado en el
asunto del correo: lo trae impreso dentro del PDF. Sin leer el PDF, ese
correo jamás se podría emparejar.

Además compara **cada adjunto por separado** (no todo pegado en un solo
texto), para que el nombre del demandado se busque en el *encabezado*
de cada documento -- que es donde de verdad está (`DEMANDADO: ...`) --
y no se diluya entre el texto de los otros adjuntos.

### Qué se guarda (solo información extraprocesal)

Por defecto (`SOLO_INFORMACION_EXTRAPROCESAL = True`) se queda
**únicamente** con:

- **derecho de petición** / **petición** / **PQR, PQRS, PQRSD**
- **las respuestas** a esos derechos de petición (así vengan rotuladas
  como "Respuesta a su petición", "RTA petición", "Respuesta PQRSD…")
- **tutela** / acción de tutela
- **pago oficioso**

Todo lo demás se omite aunque coincida con un proceso: demandas,
memoriales, mandamientos, sentencias, liquidaciones, embargos, medidas
cautelares, etc. (`PALABRAS_PROCESAL_EXCLUIR` en
`clasificar_procesos_ejecutivos.py`).

El tipo se decide mirando **el asunto y cada adjunto por separado** --
no todo el texto pegado. Esto importa: un correo con asunto genérico
("Notificación 12345") pero con `RESPUESTA DERECHO DE PETICION.pdf`
adjunto **sí** se guarda, cuando mirando solo el asunto (o el texto
completo, donde el nombre del adjunto queda enterrado después de un
cuerpo largo) se habría perdido.

En el log y en el CSV queda **por qué** se consideró extraprocesal cada
correo, para poder revisarlo.

### Se puede hacer por partes

Viene configurado para revisar los **últimos 730 días (2 años)**
(`DIAS_HACIA_ATRAS`). Una ventana así **no** hace que la corrida sea
eterna, porque va por partes: lleva un archivo de control
(`revisar_correo_pro_progreso.json`) con los correos ya revisados.

- Si lo **cortas** a la mitad, se cae la luz, o Gmail te corta la
  conexión, la próxima corrida **sigue donde se quedó**.
- Cada corrida se cierra sola por **cantidad**
  (`MAX_CORREOS_POR_CORRIDA`, 500) **o por tiempo**
  (`MAX_MINUTOS_POR_CORRIDA`, 30 min) -- lo que ocurra primero. El tope
  por tiempo importa porque los correos son muy desiguales: 500 de
  texto se revisan en minutos, pero 500 con escaneos pesados pueden
  tardar horas.
- Siempre avanza **al menos un grupo** por corrida, aunque el tope de
  tiempo ya esté cumplido -- así nunca te quedas corriéndolo sin que
  avance nada.
- Empieza por los **más nuevos** (`EMPEZAR_POR_LOS_MAS_NUEVOS`): con 2
  años de correo, lo primero que queda clasificado es lo más reciente,
  que suele ser lo más urgente.
- El log te dice cuántas corridas faltan aproximadamente para ponerte
  al día. Una vez al día, cuando ya no queden pendientes, cada corrida
  revisa solo lo nuevo y termina en segundos.

### Ajustes principales (al inicio del archivo)

- `MODO_PRUEBA` (por defecto `True`): primero muestra en el log qué
  haría, **sin descargar ni guardar nada**. Cuando se vea bien, ponlo
  en `False`.
- `DIAS_HACIA_ATRAS` (730 = 2 años): cuántos días hacia atrás revisar.
- `MAX_CORREOS_POR_CORRIDA` (500) y `MAX_MINUTOS_POR_CORRIDA` (30):
  cuándo se cierra cada corrida. Súbelos para avanzar más de una vez,
  bájalos si prefieres corridas más cortas.
- `EMPEZAR_POR_LOS_MAS_NUEVOS` (`True`), `TAMANO_LOTE_DESCARGA` (20).
- `LEER_TEXTO_DE_ADJUNTOS` (`True`): ponlo en `False` si quieres que
  vaya más rápido y te conformas con asunto/cuerpo/nombre del adjunto.
- `SOLO_INFORMACION_EXTRAPROCESAL` (**`True`** por defecto): se queda
  solo con información extraprocesal (ver el detalle abajo). Ponlo en
  `False` si alguna vez quieres bajar *todo* lo que coincida con un
  proceso, sin filtrar por tipo.
- `EXIGIR_MENCION_ESSA` (`False`): en `True` descarta lo que no
  mencione a ESSA.

### Por qué no se traba

Cada punto corresponde a un fallo real que se vio en la práctica:

- **Doble límite de tiempo** en toda operación de red: el del socket, y
  uno "duro" (un hilo aparte que corta pase lo que pase). Este último
  hace falta porque en equipos con **antivirus que inspecciona el
  correo** (Avast/Kaspersky/ESET/McAfee) el límite normal del socket
  *no se dispara* y el programa se queda muerto sin dar ni un error.
- **Reconecta** hasta `MAX_RECONEXIONES` veces y retoma exactamente
  donde iba.
- **Usa UID, no número de orden**: el "número" de un correo cambia si
  llegan o se borran mensajes mientras el script corre -- reconectar
  con números de orden te hace saltar o repetir correos en silencio.
  También verifica el `UIDVALIDITY` del buzón: si Gmail lo cambia, el
  progreso se descarta solo (avisando) en vez de dar por revisados
  correos equivocados.
- **Un correo dañado no tumba la corrida**: cada correo se procesa
  aislado (MIME roto, PDF corrupto, nombre de archivo imposible).
- **Nunca pierde correos en silencio**: si un grupo falla incluso
  reconectando, se salta para no frenar la corrida pero **no** se da
  por revisado -- queda pendiente para la próxima.
- **Guarda el avance cada `GUARDAR_PROGRESO_CADA_N`** correos, no solo
  al final, y de forma atómica (un corte de luz no corrompe el
  archivo de progreso).
- **Se ve que está vivo**: deja un aviso de avance cada
  `AVISO_PROGRESO_CADA_N` correos.

Necesita `credenciales_sgde.txt` (las mismas de siempre, contraseña de
aplicación de Gmail -- ver más abajo).

## Listar terminados por auto o por pago (partes y radicado)

`listar_terminados_auto_pago.py` (o su iniciador
`listar_terminados_auto_pago.bat`) es un reporte de **solo lectura**:
lee el mismo Excel que `clasificar_procesos_ejecutivos.py` (hoja
`ACTIVOS`, con el mismo arrastre de los procesos "acumulados") y filtra
los procesos cuyo `ESTADO PROCESAL` es exactamente `TERMINADO POR AUTO`
o `TERMINADO POR PAGO` -- **no** incluye `TERMINADO POR CONTRATO`/
`PREPAGO` ni `NO INICIO`. Genera `terminados_auto_pago.csv` con una
fila por proceso: `No.`, `Estado`, `Radicado`, `Demandante` (siempre
ESSA/Electrificadora de Santander -- no hay columna de demandante en el
Excel porque siempre es la misma parte) y `Demandado(s)`. Un proceso
"acumulado" (varias cuentas bajo el mismo radicado) sale en una sola
fila, con todos sus demandados juntos.

Además genera `terminados_auto_pago_importar.xlsx`, listo para
importar a un sistema propio: una fila por caso, dos columnas --
`radicado` y `correo del responsable` (esta última siempre **vacía**;
si tu sistema no reconoce el correo como de tu firma, simplemente crea
el caso sin asignar). Los procesos que todavía no tienen radicado no
salen en este archivo (no hay como importarlos sin radicado) -- quedan
avisados en el log para que los completes a mano.

No crea, renombra ni borra ninguna carpeta -- solo genera los dos
reportes.

## Comparar Excel vs disco (solo un reporte de lo que falta o sobra)

`comparar_excel_disco.py` (o su iniciador `comparar_excel_disco.bat`)
compara, por **radicado**, los procesos del Excel contra las carpetas
que existen en el disco, en las **dos direcciones**:

1. Procesos del Excel que **no** tienen ninguna carpeta en el disco --
   reporte `comparar_excel_disco_faltantes.csv` (fila del Excel,
   cuenta, radicado, juzgado, demandado y estado, cuando esas columnas
   existen).
2. Carpetas del disco cuyo radicado **no** aparece en el Excel --
   reporte `comparar_excel_disco_sobran_en_disco.csv` (nombre de la
   carpeta y radicado) -- útil para detectar carpetas de procesos que
   ya no están vigentes o con el radicado mal escrito.

- Dos radicados se consideran el **mismo proceso** si son idénticos, o
  si solo difieren en el último dígito (el consecutivo de
  instancia/reparto, ej. termina en 0 o en 1) -- la misma tolerancia
  que usa `validar_renombrar_carpetas.py`.
- Es un script **de solo lectura**: nunca mueve, renombra, crea ni
  borra nada, solo compara y reporta. Es el más rápido de correr
  cuando solo quieres saber qué falta o qué sobra, sin tocar ninguna
  carpeta.
- Solo necesita la columna `RADICADO` del Excel para comparar -- no
  depende de ninguna columna de número de proceso, así que funciona
  incluso con informes que no tengan una columna `"No."`.
- Usa la misma configuración (`RUTA_EXCEL`, `CARPETA_PROCESOS`, etc)
  de `validar_renombrar_carpetas.py` -- no hay que configurarla dos
  veces.

## Borrar una carpeta que Windows no deja borrar

Si el Explorador de Windows se niega a borrar una carpeta (rutas muy
largas, metadatos raros, o simplemente se traba) usa
`borrar_carpeta_forzado.bat`:

- Arrastra la carpeta problemática sobre el `.bat` (o dale doble clic y
  pega la ruta completa cuando te la pida).
- Te pide confirmación escribiendo `SI` antes de borrar nada.
- Primero vacía el contenido con `robocopy` (que sí sabe manejar rutas
  largas, a diferencia del Explorador) y luego borra la carpeta ya
  vacía.
- ⚠️ Es **irreversible** -- solo úsalo cuando estés segura de que
  quieres borrar esa carpeta completa y todo su contenido.
- Si después de correrlo la carpeta sigue sin borrarse del todo, es
  que algún archivo de adentro está abierto en otro programa (Word,
  Excel, un PDF, el antivirus escaneándola) -- ciérralo e inténtalo de
  nuevo.

## Buscar y postularse a empleos automáticamente (`buscar_empleo.py`) ⭐

Esta herramienta **no tiene nada que ver con los procesos jurídicos**:
vive en la misma carpeta por comodidad, pero es un programa aparte. Lo
que hace es vigilar los portales de empleo colombianos, mirar cada
vacante nueva, compararla contra **tu perfil** y **mandar tu hoja de
vida** únicamente a las que de verdad encajan contigo.

Se maneja desde el PC, y además **desde el celular**: trae una app
instalable (ver "La app para el celular", más abajo) para ver las
postulaciones y disparar búsquedas desde la mano.

La idea de fondo es esta: postularse a 200 ofertas no sube tus
opciones, las baja (te descartan en el filtro automático de la empresa,
y los reclutadores del mismo grupo ven las postulaciones repetidas). Lo
que sí sube tus opciones es postularte **rápido** —en las primeras
horas— a las ofertas donde **sí cumples los requisitos**. Eso es lo que
automatiza este programa: la velocidad y el filtro, no la cantidad.

### Qué hace, en orden

1. **Lee tu perfil** de `perfil_laboral.json`: qué cargos buscas, en qué
   ciudades, cuántos años de experiencia tienes, qué estudios, cuál es
   tu salario mínimo, qué palabras no quieres ver, y dónde está tu hoja
   de vida.
2. **Busca** en cada portal habilitado usando tus cargos como término de
   búsqueda, con un navegador real (Playwright), igual que si buscaras tú.
3. **Abre cada vacante nueva** y le lee la descripción completa (no solo
   el título: los requisitos de verdad están en el cuerpo del aviso).
4. **La califica de 0 a 100** contra tu perfil.
5. **Postula** solo a las que pasan el umbral (65 puntos por defecto):
   - Si la oferta publica un **correo** de contacto → le manda tu hoja de
     vida adjunta con una carta de presentación armada para *esa* vacante.
     Esto es 100 % automático y no depende de ningún portal.
   - Si no hay correo → entra al portal con **tu sesión** ya iniciada y usa
     el botón de "Postularme" del propio portal.
6. **Anota todo** en `datos_empleo/postulaciones.json` y en un Excel, para
   no postular dos veces a lo mismo y para que revises qué se mandó, a
   dónde y por qué.

### ⚠️ Antes de usarlo — léelo, son dos párrafos

Los portales de empleo prohíben en sus términos de uso el acceso
automatizado, y **LinkedIn en particular suspende cuentas** que lo hagan.
Por eso el programa está armado así:

- La vía principal es el **correo**: mandarle tu hoja de vida al correo
  que la propia oferta publica no infringe nada de nadie.
- En los portales, se usa **tu navegador con tu sesión** —la inicias tú
  mismo, a mano, una sola vez con `--login`— y solo se repiten los pasos
  que tú ya estás autorizado a hacer manualmente. No entra a APIs
  ocultas, no evade captchas ni bloqueos, y espera entre acción y acción
  como una persona.
- **LinkedIn viene desactivado a propósito** en la sección `FUENTES` de
  `buscar_empleo.py`. Actívalo solo si aceptas el riesgo de que te
  restrinjan la cuenta; es tu perfil profesional el que está en juego.

Y hay un límite honesto que conviene tener claro: el programa **manda la
hoja de vida, no consigue el trabajo**. Sirve para no perderte ofertas y
para no gastar tres horas diarias llenando formularios. La entrevista
sigue siendo tuya.

### Instalación

Las dependencias ya están en `requirements.txt` (son las mismas que usa
el resto del proyecto):

```
pip install -r requirements.txt
playwright install chromium
```

### Configuración (dos archivos)

**1. Tu perfil.** Copia `perfil_laboral.example.json`, renómbralo a
`perfil_laboral.json` y llénalo. El archivo de ejemplo trae una nota
explicando cada campo (las claves que empiezan por `nota_` son solo
comentarios, no las borres ni les hagas caso). Los campos que de verdad
deciden todo son:

| Campo | Para qué sirve |
|---|---|
| `cargos_objetivo` | Con esto se busca en los portales, **y** suma 45 puntos si aparece en el título. Escríbelos como los publican las empresas ("auxiliar jurídico", no "trabajo de abogado"). No pongas más de 6 u 8. |
| `palabras_excluyentes` | Si cualquiera aparece en la oferta, se descarta de una. Aquí va lo que no aceptas ("solo comisión", "multinivel", "inversión inicial"). |
| `anos_experiencia` | Los que **tienes**. Se descarta lo que pida más de 1 año por encima. |
| `nivel_educativo` | `bachiller`, `tecnico`, `tecnologo`, `profesional`, `especializacion`, `maestria` o `doctorado`. |
| `ciudades` | Dónde sí trabajarías. Déjalo en `[]` para no filtrar por ciudad. |
| `salario_minimo` | En pesos. Pon `0` para no filtrar por salario. |
| `ruta_hoja_de_vida` | Ruta a tu hoja de vida en PDF. **Sin esto no se puede postular por correo.** |
| `carta_presentacion` | El cuerpo del correo. Acepta comodines `{titulo}`, `{empresa}`, `{portal}`, `{nombre}`, `{resumen}`... que se reemplazan solos con los datos de cada vacante. |

**2. El correo desde el que salen las postulaciones.** Copia
`credenciales_empleo.example.txt`, renómbralo a `credenciales_empleo.txt`
y pon tu correo y una **Contraseña de aplicación** de Gmail (la misma
idea que en `credenciales_sgde.txt`: nunca tu contraseña normal; se crea
en https://myaccount.google.com/apppasswords).

Las claves de Computrabajo, elempleo o LinkedIn **no van en ningún
archivo**. Esas las escribes tú una sola vez con `--login`, y quedan
guardadas como cookies del navegador. El programa nunca las ve.

Los dos archivos con tus datos reales (`perfil_laboral.json` y
`credenciales_empleo.txt`), más toda la carpeta `datos_empleo/`, están en
`.gitignore`: nunca se suben a GitHub.

### Uso

**Primera vez, en este orden:**

```
python buscar_empleo.py --login        (abre el navegador, inicias sesión tú)
python probar_filtro_empleo.py         (prueba el filtro sin internet)
python buscar_empleo.py --simular      (busca de verdad, pero NO manda nada)
```

`--simular` es el paso importante: te muestra a qué vacantes postularía y
con cuántos puntos, sin mandar nada. **Córrelo varias veces y ajusta tu
perfil hasta que la lista que muestra sea la que tú mismo escogerías.**
Ahí sí, ya en serio:

```
python buscar_empleo.py                (una pasada de verdad)
python buscar_empleo.py --vigilar      (queda revisando cada 2 horas)
python buscar_empleo.py --reporte      (exporta el Excel y resume)
python buscar_empleo.py --diagnostico  (revisa si los portales responden)
```

En Windows puedes hacer doble clic en:

- `probar_empleo.bat` → el modo simulación (empieza por aquí).
- `buscar_empleo.bat` → una pasada real.
- `vigilar_empleo.bat` → queda vigilando de forma indefinida.

### Los tres modos

Se ponen en `MODO`, dentro de `buscar_empleo.py`, o con `--simular` /
`--automatico` al llamarlo:

- **`simulacion`** (el de fábrica): busca, califica y te muestra qué
  haría. No manda nada.
- **`semiautomatico`**: manda las postulaciones **por correo** solas, y en
  los portales hace el clic de "Postularme" —que en la mayoría ya postula
  de una—. Si el portal abre **además** un formulario, no lo llena: deja
  la vacante marcada como "pendiente_revision" con su link, para que la
  termines tú.
- **`automatico`**: manda todo solo. En los portales solo envía cuando
  reconoció el formulario completo; si el formulario pide algo que no
  entiende (preguntas del empleador, pruebas, subir documentos), **no
  inventa respuestas**: deja la vacante pendiente para que la termines
  tú. Una respuesta inventada te quema la postulación; dejarla a medias,
  no.

### Cómo decide a qué postular

Primero mira los requisitos **duros**. Si no los cumples, descarta de una
—mandar la hoja de vida ahí solo gasta tu cupo del día—:

- pide más experiencia de la que tienes (con 1 año de tolerancia);
- pide más estudios de los que tienes;
- dice cuánto paga y paga menos de tu mínimo;
- queda en una ciudad que no trabajas y no es remota;
- trae una de tus palabras excluyentes;
- se publicó hace más de 30 días (ya está cerrada aunque siga visible).

Lo que sobrevive suma puntos: el cargo en el título (+45), tus palabras
clave (+20), tu ciudad (+15), la experiencia que sí cumples (+15), el
nivel educativo (+10), que publique correo (+5), que sea recién publicada
(+5). Con 65 o más, postula.

Todo esto se lee de la oferta escrita en español real: reconoce
"experiencia mínima de 2 años", "de 1 a 3 años", "18 meses de
experiencia", "sin experiencia", "$2.500.000", "3 millones", "publicado
hace 3 días", "ayer". Puedes comprobarlo tú mismo con
`python probar_filtro_empleo.py`, que le pasa ofertas de ejemplo al mismo
filtro y verifica que decida bien, sin tocar internet ni gastar una sola
postulación.

Cuando termina, te manda un correo con el resumen (ver "Que trabaje
solo y te avise al celular", más abajo).

Dos protecciones más, que importan tanto como el filtro:

- **Tope diario** (`MAX_POSTULACIONES_DIA`, 15 de fábrica). Lo que sobra
  queda "en_espera" para el día siguiente.
- **Antiduplicados**: si la misma vacante está publicada en Computrabajo
  y en elempleo, se postula una sola vez.

### Qué archivos genera

Todo dentro de `datos_empleo/`:

- `postulaciones.xlsx` — el reporte para leer tú: fecha, estado, puntaje,
  cargo, empresa, portal, cómo se postuló y el link.
- `postulaciones.json` — la memoria del programa (lo que ya vio y lo que
  ya mandó). Si lo borras, volvería a postular a lo mismo.
- `buscar_empleo.log` — el detalle de todo, incluido el motivo exacto por
  el que descartó cada oferta.
- `perfil_navegador/` — tus sesiones de los portales. **No la subas ni la
  compartas**: son tus cookies de sesión.

Los estados posibles son: `postulada`, `descartada`, `duplicada`,
`en_espera` (tope diario lleno), `pendiente_revision` (te toca a ti
terminarla) y `error`.

### Si un portal deja de dar resultados

Los portales cambian su página cada tanto y ahí es donde se rompen
siempre este tipo de programas. Este está armado para que puedas
arreglarlo tú sin tocar código:

1. Corre `python buscar_empleo.py --diagnostico`. Abre cada portal, te
   dice cuál devolvió **SIN RESULTADOS**, con qué dirección exacta lo
   intentó, y te deja una captura de pantalla en
   `datos_empleo/diagnostico/`.
2. Corrige ese portal en la sección **`FUENTES`**, al inicio de
   `buscar_empleo.py`: ahí están, uno por uno, la dirección de búsqueda y
   los selectores de cada portal, comentados. No hay que tocar nada más.

Además, el programa ya trae una red de seguridad: si los selectores de un
portal dejan de funcionar, recoge las ofertas reconociendo los **links**
(que casi nunca cambian de forma) y sigue trabajando, dejando un aviso en
el log. Y **las direcciones de los portales que vienen de fábrica no se
pudieron probar contra internet** al escribir el programa, así que trata
`--diagnostico` como el primer paso obligatorio, no como un extra.

### La app para el celular (`app_movil.py`)

El buscador **no puede correr dentro del teléfono**, y conviene saber por
qué antes de nada: para leer los portales hay que abrir un navegador
Chromium de verdad, y ni Android ni iPhone dejan hacer eso en segundo
plano (además, las tiendas rechazan las apps que se postulan solas). Así
que el trabajo se reparte:

```
   EL PC  ──── busca en los portales y manda las hojas de vida
     │
     │  (tu WiFi)
     ▼
 EL CELULAR ── es la app: ver, decidir y disparar búsquedas
```

La consecuencia práctica es una sola: **el PC tiene que estar encendido**
y con `app_movil.py` corriendo. Si lo apagas, la app abre igual (queda
instalada en el teléfono) pero avisa "PC apagado" y no trae datos nuevos.

#### Cómo se instala en el teléfono

1. En el PC, doble clic en **`abrir_movil.bat`** (o `python app_movil.py`).
   Deja esa ventana abierta.
2. Esa ventana te muestra **un código QR** (se abre solo en pantalla) y
   una dirección tipo `http://192.168.1.15:8777/?t=A1B2C3D4`. Apunta la
   cámara del celular al QR y entras de una — o escribe la dirección a
   mano si prefieres. El celular tiene que estar **en el mismo WiFi que
   el PC**.
3. Instálala como app:
   - **iPhone (Safari)**: botón compartir → *Agregar a pantalla de inicio*.
   - **Android (Chrome)**: menú de los 3 puntos → *Agregar a pantalla de
     inicio*.

Queda con su ícono propio y abre en pantalla completa, sin barra de
navegador. No hay que pasar por ninguna tienda, no hay que firmar nada y
no caduca.

Un detalle técnico, por si en Android no te aparece *Instalar
aplicación* sino solo *Agregar a pantalla de inicio*: es normal y no es
un error. Los navegadores reservan la instalación completa para
conexiones `https`, y esto se sirve por `http` dentro de tu WiFi (montar
`https` de verdad exigiría un certificado y un dominio, para algo que
nunca sale de tu casa). El acceso directo funciona igual: mismo ícono,
misma pantalla completa. Lo único que no tienes es que la app abra sin
el PC prendido — y sin el PC no habría datos que mostrar de todos modos.

#### Qué puedes hacer desde el celular

- **Inicio**: cuántas llevas postuladas hoy contra el tope diario, cuántas
  quedaron "para ti", y los dos botones para lanzar una búsqueda —
  *Probar sin enviar* o *Buscar y postular*—. Mientras el PC busca, ves el
  avance en vivo, línea por línea.
- **Vacantes**: todas las ofertas con su puntaje, filtrables por estado y
  buscables por cargo o empresa. Cada una muestra **por qué** se descartó
  o cómo se postuló, con un botón para abrir la oferta original. Las que
  quedaron "para ti" traen además *Ya la hice*, para sacarlas de la lista
  cuando las remates a mano.
- **Filtro**: cargos, palabras que suman, palabras que descartan, ciudades,
  remoto sí/no, años de experiencia, salario mínimo, nivel educativo,
  umbral, tope diario y modo. Se guarda directo en el `perfil_laboral.json`
  del PC y se aplica de inmediato, sin reiniciar nada.
- **Registro**: el log completo, tal cual, por si algo salió raro.

Lo que **no** se puede tocar desde el celular es a propósito: la hoja de
vida, las credenciales del correo y la carta de presentación se editan
solo en el PC. Son los campos que dejarían todo roto si se tocan por
accidente desde el bus.

#### Seguridad

El panel muestra tus datos y puede mandar postulaciones a tu nombre, así
que viene cerrado:

- **Solo escucha en tu red local.** No lo publiques en internet ni le
  abras puertos al router.
- **Pide un código** de 8 caracteres que se genera solo la primera vez y
  queda en `datos_empleo/token_movil.txt`. Se escribe una vez por teléfono
  y no se vuelve a pedir. Si crees que alguien más lo tiene, borra ese
  archivo y vuelve a arrancar: sale uno nuevo.
- Las claves de los portales **nunca pasan por aquí**: siguen viviendo
  como cookies del navegador del PC, de cuando corriste `--login`.

Si quieres que ni siquiera salga al WiFi (para probar solo en el PC):

```
python app_movil.py --solo-este-pc
python app_movil.py --puerto 9000     (si el 8777 está ocupado)
```

#### Detalles útiles

- No necesita instalar nada nuevo: usa la librería estándar de Python. Los
  íconos se generan con `python movil/generar_iconos.py` (ya vienen hechos;
  solo hay que correrlo si quieres cambiarles el color o el dibujo).
- La app funciona en modo claro y oscuro según como tengas el teléfono.
- La dirección del PC (`192.168...`) **cambia si cambias de WiFi**. Si un
  día no conecta, vuelve a mirar la que muestre la ventana del PC.
- Si quieres que busque sola cada dos horas sin que tú toques nada, deja
  `vigilar_empleo.bat` corriendo en el PC **además** de `abrir_movil.bat`:
  el celular te sirve entonces para ir mirando lo que va cayendo.


### Que trabaje solo y te avise al celular

Lo anterior sirve para mirar y decidir desde el teléfono. Esto es para
que no tengas ni que mirar: el PC busca solo y te escribe cuando pasa
algo.

#### 1. Que te avise al celular (por correo)

Cada vez que termina una pasada, el PC te manda un correo **a ti** con lo
que hizo: qué postuló, qué quedó esperándote, y el link de cada una para
abrirla de un toque. Al iPhone le llega como cualquier notificación de
correo, sin abrir la app ni estar en el WiFi de la casa.

Ya viene encendido; solo necesita que hayas configurado
`credenciales_empleo.txt` (el mismo correo con el que se postula). El
aviso llega a la dirección que tengas en `"correo"` dentro de
`perfil_laboral.json`.

**Solo escribe cuando hay algo que contar.** Si en esa pasada no postuló
nada ni quedó nada pendiente, no manda nada. Un aviso de "no encontré
nada" cada dos horas se vuelve ruido, y el ruido se termina silenciando —
que es justo lo que no queremos. Si aun así prefieres no recibirlos, pon
`AVISAR_POR_CORREO = False` en `buscar_empleo.py`.

#### 2. Que arranque solo con el computador

Doble clic en **`instalar_inicio_automatico.bat`** (si te dice que no
pudo, clic derecho → *Ejecutar como administrador*). Desde ahí, cada vez
que prendas el computador y entres a tu usuario, arrancan solas y sin
ventanas:

- la **vigilancia**, que revisa los portales cada 2 horas y postula;
- la **app del celular**, para que el teléfono conecte cuando lo abras.

No tienes que abrir nada más. Los otros tres archivos son para manejarlo:

| Archivo | Para qué |
|---|---|
| `empleo_automatico.bat` | Arrancar las dos cosas **ahora**, sin reiniciar. |
| `parar_empleo.bat` | Pararlas. Lo ya postulado queda guardado. |
| `desinstalar_inicio_automatico.bat` | Que deje de arrancar solo. No borra nada. |

Como corren sin ventana, para saber si están vivas: abre la app en el
celular (si conecta, están corriendo), o busca `pythonw.exe` en el
Administrador de tareas. Lo que hicieron queda siempre en
`datos_empleo/buscar_empleo.log` y en el Excel.

Un aviso sobre `parar_empleo.bat`: cierra **todos** los programas de
Python que corran sin ventana en ese equipo. Si tienes también el de
procesos jurídicos corriendo así, se cierra igual.


### Portales incluidos

Computrabajo, elempleo.com, Magneto365 y el Servicio Público de Empleo
(el del Gobierno) vienen **activos**. LinkedIn viene **desactivado** por
lo que se explicó arriba. Para agregar otro portal, copia una entrada de
`FUENTES`, cámbiale la dirección y el patrón de link, y ponle
`"habilitada": True`.


## Si algo falla

- El sitio del SGDE puede cambiar de diseño con el tiempo, lo que puede
  romper los selectores que usa Playwright. Si ves errores en el log,
  dime qué cambió en la pantalla y te ayudo a actualizar el script.
- Si un proceso descargado a mano no tiene ningún documento con el
  radicado en el formato esperado, la carpeta se organiza igual usando
  el nombre del zip original, y queda registrado como advertencia en el
  log.
- Si un zip no deja **ningún** archivo al extraerlo (por ejemplo porque
  todo su contenido está protegido con contraseña, las rutas son
  demasiado largas para Windows, o el antivirus puso los archivos en
  cuarentena justo después de extraerlos), el programa avisa con un
  `ERROR` bien visible en el log, **no** crea una carpeta vacía
  disfrazada de "organizada", y **no** mueve el zip a `Procesados` — se
  queda en Descargas para que lo revises o reintentes a mano. Si algunos
  archivos sí se extrajeron pero otros no, la carpeta se organiza igual
  con lo que se pudo, y queda un `WARNING` explicando cuántos fallaron.
- Al extraer un zip, si una carpeta o archivo dentro trae **espacios o
  puntos al final del nombre** (Windows los maneja mal incluso con el
  truco de rutas largas), se le quitan automáticamente al crearlo en el
  disco -- así se evita el error clásico "el sistema no puede encontrar
  la ruta especificada" al leerlo después.
- Si `procesos_faltantes_en_disco.csv` (o los otros reportes CSV) están
  abiertos en Excel cuando corres `validar_renombrar_carpetas.py`, el
  script ya no crashea al intentar guardarlos -- avisa con un `ERROR`
  claro pidiéndote que cierres el archivo, y el resto del análisis
  (que sí se calculó bien) queda visible igual en el log.
- **Importante sobre carpetas vacías de ANTES de esta corrección**: si ya
  tienes carpetas vacías de cuando el programa sí las creaba aunque la
  extracción fallara, `procesos_juridicos.py` **no las va a arreglar
  solo**, aunque subas `DIAS_ATRAS_PROCESAR_EXISTENTES` — el zip
  correspondiente probablemente ya se movió a `Descargas\Procesados`
  (porque el programa viejo lo marcaba como "hecho" sin comprobar), y la
  vigilancia automática nunca mira dentro de esa subcarpeta a propósito
  (para no reprocesar cosas ya hechas de verdad). Corre
  `validar_renombrar_carpetas.py`, revisa `carpetas_vacias.csv` (columna
  "Donde se encontró"): si dice `Descargas/Procesados`, tienes que sacar
  ese zip de ahí a mano (muévelo de vuelta a Descargas) para que se
  vuelva a intentar.
