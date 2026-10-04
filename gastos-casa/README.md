# Cuentas de la Casa

App para llevar los gastos e ingresos del hogar. Se instala en el teléfono como cualquier otra app y **todas las personas de la casa ven y registran lo mismo desde su propio celular**, en tiempo real. Sin internet se sigue usando y se sincroniza al volver la conexión.

## Dos espacios: Hogar y Personal

Arriba del menú hay un selector **Hogar | Personal**.

- **Hogar**: las cuentas de la casa. Es el espacio que se comparte con la familia y se sincroniza en todos los teléfonos.
- **Personal**: tu propia plata (tu salario, tus gastos, tus cuentas y metas). **Se guarda solo en tu teléfono**, no se sube a la nube y nadie más lo ve. Cada persona tiene el suyo.

Cada espacio tiene su propio color (por defecto verde para Hogar y morado para Personal; se cambia en Ajustes) para saber siempre dónde estás.

Para conectar los dos sin escribir doble:
- En **Personal**, al registrar un gasto marca *"Es un aporte para la casa"*: también aparece en Hogar como ingreso a tu nombre.
- En **Hogar**, al registrar un gasto marca *"Lo pagué con mi plata"*: también aparece en tus gastos personales.

Importante: como el espacio personal no está en la nube, haz copias de seguridad desde *Ajustes* estando en Personal.

## Qué hace

- **Inicio**: balance del mes, ingresos, gastos, tasa de ahorro, proyección de gasto al cierre, gastos por categoría, presupuesto, gasto acumulado día a día, quién aporta y metas.
- **Movimientos**: registrar, editar y borrar ingresos y gastos (valor, fecha, categoría, persona, medio de pago, notas), con filtros y exportación a Excel.
- **Informes** diario, mensual y anual, exportables a Excel o PDF (imprimir).
- **Personas**: porcentaje que le toca a cada una, cuánto pagó, cuánto le correspondía y quién le debe a quién.
- **Presupuesto** por categoría, **pagos fijos** que se registran con un clic y **metas de ahorro**.
- **Ajustes**: moneda, tema claro/oscuro, copia de seguridad, compartir con la familia e instalar.

## Puesta en marcha (una sola vez, unos 15 minutos)

Lo hace una persona de la casa. Las demás solo abren un enlace.

### 1. Crear la base de datos gratuita (Firebase)

1. Entra a <https://console.firebase.google.com> con una cuenta de Google y pulsa **Crear proyecto** (por ejemplo `cuentas-casa`). Google Analytics no hace falta.
2. Menú **Compilación → Authentication → Comenzar → Método de acceso → Anónimo → Habilitar → Guardar**.
3. Menú **Compilación → Firestore Database → Crear base de datos**. Elige una ubicación cercana (por ejemplo `southamerica-east1` o `us-east1`) y modo **producción**.
4. En Firestore, pestaña **Reglas**: borra lo que hay, pega el contenido del archivo [`firestore.rules`](firestore.rules) y pulsa **Publicar**.
5. En **Configuración del proyecto (engranaje) → General → Tus apps**, pulsa el icono **`</>`** (web), ponle un nombre y regístrala. Copia el bloque `const firebaseConfig = { ... }` que aparece; lo necesitas en el paso 3.

El plan gratuito (Spark) alcanza de sobra para una familia.

### 2. Publicar la app en internet

Para instalarla en el teléfono, la app tiene que estar en una dirección `https://`. La forma más fácil:

**Netlify Drop (sin instalar nada)**
1. Entra a <https://app.netlify.com/drop> (crea una cuenta gratis si te la pide).
2. Arrastra la carpeta **`gastos-casa`** completa a la página.
3. Te da una dirección tipo `https://nombre-raro.netlify.app`. En *Site configuration → Change site name* puedes cambiarla, por ejemplo `cuentas-familia-perez.netlify.app`.
4. Para publicar una versión nueva más adelante: *Deploys → arrastra otra vez la carpeta*.

**Alternativa: Firebase Hosting** (si tienes Node.js instalado), dentro de la carpeta `gastos-casa`:
```
npm install -g firebase-tools
firebase login
firebase use --add        (elige tu proyecto)
firebase deploy
```
Esto también publica las reglas de seguridad.

No uses GitHub Pages en este repositorio: publicaría también los demás archivos del repo.

### 3. Crear el hogar e invitar a la familia

1. Abre la dirección de la app en tu teléfono → **Ajustes → Compartir con la familia → Crear hogar compartido**.
2. Pega el bloque `firebaseConfig` del paso 1 y pulsa **Crear hogar**. Si ya tenías datos en ese teléfono, se suben.
3. Pulsa **Enviar por WhatsApp** (o **Copiar enlace de invitación**) y mándalo a cada persona de la casa.
4. Cada persona abre el enlace, toca **Unirme** y elige quién es. Desde ahí, al registrar un gasto aparece como quien pagó.

### 4. Instalarla en cada teléfono

- **Android (Chrome)**: menú ⋮ → **Instalar app** (o el botón *Instalar app* que aparece en la propia app).
- **iPhone (Safari)**: botón **Compartir** → **Agregar a pantalla de inicio**.

Queda con su icono, abre a pantalla completa y funciona sin conexión.

## Seguridad y privacidad

- Los datos se guardan en **tu** proyecto de Firebase, no en un servidor de terceros.
- Solo entra quien tiene el enlace de invitación, que lleva un código aleatorio de 20 caracteres. Las reglas impiden listar o buscar hogares. **Comparte el enlace solo con tu familia**: quien lo tenga puede ver y cambiar las cuentas.
- Si un teléfono se pierde o alguien deja la casa, crea un hogar nuevo (Ajustes → Dejar de sincronizar → Crear hogar compartido, con "subir mis datos" marcado) y envía el enlace nuevo a quienes siguen.
- Descarga de vez en cuando una **copia de seguridad** desde Ajustes.

## Uso sin compartir

Si solo abres `index.html` con doble clic en el computador, funciona igual pero los datos quedan solo en ese navegador y no se puede instalar ni compartir.

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `index.html` | La app completa |
| `manifest.webmanifest`, `icons/` | Nombre, colores e iconos al instalarla |
| `sw.js` | Permite abrirla sin conexión. Si cambias la app, sube el número de `VERSION` |
| `vendor/` | Gráficos (Chart.js), iconos (Lucide) y Firebase, incluidos para que funcione sin internet |
| `firestore.rules` | Reglas de seguridad de la base de datos |
| `firebase.json` | Configuración opcional para publicar con Firebase Hosting |
