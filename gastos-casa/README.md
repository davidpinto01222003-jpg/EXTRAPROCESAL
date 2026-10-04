# Cuentas de la Casa

App para llevar los gastos e ingresos mensuales del hogar. Es un solo archivo (`index.html`): ábrelo con doble clic en Chrome, Edge o Firefox. No hay que instalar nada.

## Qué hace

- **Inicio**: balance del mes, ingresos, gastos, tasa de ahorro, proyección de gasto al cierre, gastos por categoría, avance del presupuesto, gasto acumulado día a día, quién aporta y metas.
- **Movimientos**: registrar, editar y borrar ingresos y gastos (valor, fecha, categoría, persona, medio de pago, notas). Filtros por mes, tipo, categoría, persona, medio de pago y texto. Exporta a Excel (CSV).
- **Informes**
  - *Diario*: movimientos del día, promedio diario y tabla día a día del mes con saldo acumulado.
  - *Mensual*: comparación con el mes anterior, detalle por categoría contra presupuesto, ingresos por fuente, medios de pago, gastos más grandes y reparto entre personas.
  - *Anual*: ingresos, gastos y balance mes a mes, ahorro acumulado, categorías del año y aportes por persona.
  - Cada informe se puede exportar a Excel o imprimir / guardar como PDF.
- **Personas**: quienes aportan, con el % de los gastos que le corresponde a cada una. Calcula cuánto pagó cada quien, cuánto le correspondía y quién le debe a quién para quedar a paz y salvo.
- **Presupuesto**: categorías de gasto e ingreso con icono, color y presupuesto mensual.
- **Pagos fijos**: arriendo, servicios, colegio, salario… Se registran en el mes con un clic.
- **Metas de ahorro**: objetivo, fecha, abonos/retiros y cuánto ahorrar al mes para llegar a tiempo.
- **Ajustes**: nombre del hogar, moneda, tema claro/oscuro, copia de seguridad (descargar/restaurar JSON) y borrado.

Atajo: tecla **N** para registrar un movimiento nuevo.

## Dónde se guardan los datos

En el navegador donde abras la app (almacenamiento local). No se envían a ningún servidor. Por eso:

- Descarga una **copia de seguridad** desde *Ajustes* con frecuencia.
- Para usarla en otro computador o celular, copia `index.html`, ábrelo allá y usa *Restaurar copia*.

La primera vez se cargan **datos de ejemplo**; bórralos con el botón *Empezar con mis datos*.

Los iconos, gráficos y tipografías se descargan de internet (Lucide, Chart.js, Google Fonts); sin conexión la app funciona, pero sin gráficos ni iconos.
