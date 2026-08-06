/* ==========================================================================
   Gestor de Pedidos - aplicacion completa (sin servidor, sin librerias).
   Los datos viven en el propio telefono (localStorage). No sale nada a
   internet: si no hay señal en la tienda, la app funciona igual.
   ========================================================================== */
'use strict';

/* ----------------------------------------------------------- utilidades */

const $  = (sel, raiz = document) => raiz.querySelector(sel);
const $$ = (sel, raiz = document) => Array.from(raiz.querySelectorAll(sel));

const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);

function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/** Texto comparable: sin tildes, sin mayusculas. Para buscar y para no
 *  duplicar clientes por escribir "La Esquina" y "la esquina". */
function norm(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

/* --- fechas: siempre en horario local, nunca UTC (evita el bug del dia -1) */

function isoDe(d) {
  const p = n => String(n).padStart(2, '0');
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
}
function fechaDe(iso) {
  const [y, m, d] = String(iso).split('-').map(Number);
  return new Date(y, m - 1, d);
}
const hoyISO = () => isoDe(new Date());

function sumarDias(iso, n) {
  const d = fechaDe(iso);
  d.setDate(d.getDate() + n);
  return isoDe(d);
}
function diasEntre(isoA, isoB) {
  return Math.round((fechaDe(isoB) - fechaDe(isoA)) / 86400000);
}
/** Lunes de la semana de esa fecha. */
function inicioSemana(iso) {
  const d = fechaDe(iso);
  const dow = (d.getDay() + 6) % 7; // 0 = lunes
  return sumarDias(iso, -dow);
}
function inicioMes(iso) {
  const d = fechaDe(iso);
  return isoDe(new Date(d.getFullYear(), d.getMonth(), 1));
}
function finMes(iso) {
  const d = fechaDe(iso);
  return isoDe(new Date(d.getFullYear(), d.getMonth() + 1, 0));
}

/* --- periodos: dia, semana, mes y año se manejan igual en toda la app.
   Un periodo es un tipo mas una fecha "ancla" que cae dentro de el. */

function rangoPeriodo(tipo, ancla) {
  const d = fechaDe(ancla);
  switch (tipo) {
    case 'semana': { const l = inicioSemana(ancla); return { desde: l, hasta: sumarDias(l, 6) }; }
    case 'mes':    return { desde: inicioMes(ancla), hasta: finMes(ancla) };
    case 'anio':   return { desde: `${d.getFullYear()}-01-01`, hasta: `${d.getFullYear()}-12-31` };
    default:       return { desde: ancla, hasta: ancla };
  }
}

function moverPeriodo(tipo, ancla, dir) {
  const d = fechaDe(ancla);
  switch (tipo) {
    case 'semana': return sumarDias(inicioSemana(ancla), dir * 7);
    case 'mes':    return isoDe(new Date(d.getFullYear(), d.getMonth() + dir, 1));
    case 'anio':   return isoDe(new Date(d.getFullYear() + dir, 0, 1));
    default:       return sumarDias(ancla, dir);
  }
}

/** ¿el periodo que contiene esta ancla es el que estamos viviendo? */
function periodoEsActual(tipo, ancla) {
  return rangoPeriodo(tipo, ancla).desde === rangoPeriodo(tipo, hoyISO()).desde;
}

const cap = s => String(s).charAt(0).toUpperCase() + String(s).slice(1);

/** Etiqueta corta, para el selector de periodo. */
function etiquetaPeriodo(tipo, ancla) {
  const d = fechaDe(ancla);
  const r = rangoPeriodo(tipo, ancla);
  const actual = periodoEsActual(tipo, ancla);
  switch (tipo) {
    case 'semana': return actual ? 'Esta semana' : `${fechaCorta(r.desde)} — ${fechaCorta(r.hasta)}`;
    case 'mes':    return `${cap(MESES[d.getMonth()])} ${d.getFullYear()}`;
    case 'anio':   return String(d.getFullYear());
    default:       return etiquetaDia(ancla) || fechaCorta(ancla);
  }
}

/** Etiqueta larga, para la barra de arriba y los documentos. */
function subPeriodo(tipo, ancla) {
  const d = fechaDe(ancla);
  const r = rangoPeriodo(tipo, ancla);
  switch (tipo) {
    case 'semana': return `Del ${fechaCorta(r.desde)} al ${fechaCorta(r.hasta)}`;
    case 'mes':    return `${cap(MESES[d.getMonth()])} de ${d.getFullYear()}`;
    case 'anio':   return `Año ${d.getFullYear()}`;
    default: {
      const e = etiquetaDia(ancla);
      return e ? `${e}, ${fechaLarga(ancla)}` : cap(fechaLarga(ancla));
    }
  }
}

const NOMBRE_TIPO = { dia: 'Día', semana: 'Semana', mes: 'Mes', anio: 'Año' };

const DIAS = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio',
  'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const MESES_C = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];

function fechaLarga(iso) {
  const d = fechaDe(iso);
  return `${DIAS[d.getDay()]} ${d.getDate()} de ${MESES[d.getMonth()]}`;
}
function fechaCorta(iso) {
  const d = fechaDe(iso);
  return `${d.getDate()} ${MESES_C[d.getMonth()]}`;
}
function fechaFactura(iso) {
  const d = fechaDe(iso);
  return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
}
function etiquetaDia(iso) {
  if (iso === hoyISO()) return 'Hoy';
  if (iso === sumarDias(hoyISO(), -1)) return 'Ayer';
  if (iso === sumarDias(hoyISO(), 1)) return 'Mañana';
  return null;
}

/* --- dinero */

const SIN_DECIMALES = ['COP', 'CLP', 'PYG', 'JPY', 'ISK'];

function decimales() {
  return SIN_DECIMALES.includes(S.config.moneda) ? 0 : 2;
}
function dinero(n) {
  const dec = decimales();
  try {
    return new Intl.NumberFormat(S.config.locale || 'es-CO', {
      style: 'currency', currency: S.config.moneda || 'COP',
      minimumFractionDigits: dec, maximumFractionDigits: dec
    }).format(Number(n) || 0);
  } catch (e) {
    return (Number(n) || 0).toFixed(dec);
  }
}
function num(n) {
  return new Intl.NumberFormat('es-CO').format(Number(n) || 0);
}

/* ------------------------------------------------------------- el estado */

const DB_KEY = 'gestor_pedidos_v1';

function estadoInicial() {
  return {
    version: 1,
    config: {
      negocio: 'Mi distribución',
      contacto: '',
      logo: '',            // foto de perfil, guardada como data URI
      configurado: false,  // ya paso por la pantalla de bienvenida
      moneda: 'COP',
      locale: 'es-CO',
      consecutivo: 1
    },
    productos: [],
    clientes: [],
    pedidos: [],
    cargados: {}   // { 'clave-del-rango': [idProducto, ...] } marcas del cargue
  };
}

let S = estadoInicial();

function cargarEstado() {
  try {
    const crudo = localStorage.getItem(DB_KEY);
    if (!crudo) return estadoInicial();
    const dato = JSON.parse(crudo);
    const base = estadoInicial();
    return {
      ...base, ...dato,
      config: { ...base.config, ...(dato.config || {}) },
      productos: dato.productos || [],
      clientes: dato.clientes || [],
      pedidos: dato.pedidos || [],
      cargados: dato.cargados || {}
    };
  } catch (e) {
    console.error('No se pudo leer el almacenamiento', e);
    return estadoInicial();
  }
}

let guardarPendiente = null;
function guardar() {
  clearTimeout(guardarPendiente);
  guardarPendiente = setTimeout(() => {
    try {
      localStorage.setItem(DB_KEY, JSON.stringify(S));
    } catch (e) {
      toast('No se pudo guardar: memoria llena');
      console.error(e);
    }
  }, 60);
}

/* --------------------------------------------------------- consultas */

const productoPorId = id => S.productos.find(p => p.id === id);

function pedidosDe(desde, hasta) {
  return S.pedidos
    .filter(p => p.estado !== 'anulado' && p.fecha >= desde && p.fecha <= hasta)
    .sort((a, b) => (a.fecha === b.fecha ? b.creado - a.creado : a.fecha.localeCompare(b.fecha)));
}
const totalPedido = p => p.items.reduce((s, i) => s + i.cant * (i.precio || 0), 0);
const unidadesPedido = p => p.items.reduce((s, i) => s + i.cant, 0);

/** Suma por producto en un rango: la base del cargue y de los reportes. */
function totalesPorProducto(desde, hasta) {
  const mapa = new Map();
  for (const ped of pedidosDe(desde, hasta)) {
    for (const it of ped.items) {
      if (!it.cant) continue;
      const prod = productoPorId(it.prodId);
      const clave = it.prodId || ('n:' + norm(it.nombre));
      const acc = mapa.get(clave) || {
        id: it.prodId, nombre: it.nombre,
        categoria: prod ? prod.categoria : (it.categoria || 'Sin categoría'),
        unidad: prod ? prod.unidad : (it.unidad || 'paq'),
        cant: 0, valor: 0, clientes: new Set()
      };
      acc.cant += it.cant;
      acc.valor += it.cant * (it.precio || 0);
      acc.clientes.add(ped.cliente);
      mapa.set(clave, acc);
    }
  }
  return Array.from(mapa.values()).sort((a, b) => b.cant - a.cant);
}

function totalesPorCliente(desde, hasta) {
  const mapa = new Map();
  for (const ped of pedidosDe(desde, hasta)) {
    const clave = norm(ped.cliente);
    const acc = mapa.get(clave) || { nombre: ped.cliente, pedidos: 0, cant: 0, valor: 0 };
    acc.pedidos += 1;
    acc.cant += unidadesPedido(ped);
    acc.valor += totalPedido(ped);
    mapa.set(clave, acc);
  }
  return Array.from(mapa.values()).sort((a, b) => b.valor - a.valor);
}

/** Serie para graficar y tabular: por dia, salvo en el año, que va por mes.
 *  Devuelve siempre {clave, etiqueta, etiquetaLarga, pedidos, cant, valor}. */
function serieDelPeriodo(tipo, desde, hasta) {
  if (tipo === 'anio') return serieMensual(desde, hasta);
  return serieDiaria(desde, hasta).map(d => ({
    clave: d.fecha,
    etiqueta: fechaCorta(d.fecha),
    etiquetaLarga: cap(fechaLarga(d.fecha)),
    pedidos: d.pedidos, cant: d.cant, valor: d.valor
  }));
}

function serieMensual(desde, hasta) {
  const acc = new Map();
  for (const p of pedidosDe(desde, hasta)) {
    const k = p.fecha.slice(0, 7);
    const a = acc.get(k) || { pedidos: 0, cant: 0, valor: 0 };
    a.pedidos++; a.cant += unidadesPedido(p); a.valor += totalPedido(p);
    acc.set(k, a);
  }
  const salida = [];
  const ini = fechaDe(desde), fin = fechaDe(hasta);
  const cursor = new Date(ini.getFullYear(), ini.getMonth(), 1);
  while (cursor <= fin) {
    const k = `${cursor.getFullYear()}-${String(cursor.getMonth() + 1).padStart(2, '0')}`;
    const a = acc.get(k) || { pedidos: 0, cant: 0, valor: 0 };
    salida.push({
      clave: k,
      etiqueta: MESES_C[cursor.getMonth()],
      etiquetaLarga: `${cap(MESES[cursor.getMonth()])} ${cursor.getFullYear()}`,
      ...a
    });
    cursor.setMonth(cursor.getMonth() + 1);
  }
  return salida;
}

function serieDiaria(desde, hasta) {
  const acc = new Map();
  for (const p of pedidosDe(desde, hasta)) {
    const a = acc.get(p.fecha) || { fecha: p.fecha, cant: 0, valor: 0, pedidos: 0 };
    a.cant += unidadesPedido(p);
    a.valor += totalPedido(p);
    a.pedidos += 1;
    acc.set(p.fecha, a);
  }
  const salida = [];
  for (let f = desde; f <= hasta; f = sumarDias(f, 1)) {
    salida.push(acc.get(f) || { fecha: f, cant: 0, valor: 0, pedidos: 0 });
  }
  return salida;
}

/* ------------------------------------------------------------- avisos */

let tToast;
function toast(msg) {
  const t = $('#toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(tToast);
  tToast = setTimeout(() => t.classList.remove('show'), 2200);
}

/* ------------------------------------------------------ despachador */

const ACC = {};   // acciones de click, por nombre en data-act

document.addEventListener('click', ev => {
  const el = ev.target.closest('[data-act]');
  if (!el) return;
  const fn = ACC[el.dataset.act];
  if (!fn) return;
  ev.preventDefault();
  fn(el, ev);
});
document.addEventListener('input', ev => {
  const el = ev.target.closest('[data-in]');
  if (el && ACC[el.dataset.in]) ACC[el.dataset.in](el, ev);
});
document.addEventListener('change', ev => {
  const el = ev.target.closest('[data-ch]');
  if (el && ACC[el.dataset.ch]) ACC[el.dataset.ch](el, ev);
});

/* --------------------------------------------------------- navegacion */

let vistaActual = 'pedidos';
const VISTAS = {};   // nombre -> función que dibuja

function irA(tab) {
  vistaActual = tab;
  $$('#tabbar .tab').forEach(b => b.classList.toggle('is-active', b.dataset.tab === tab));
  $$('.view').forEach(v => { v.hidden = v.id !== 'view-' + tab; });
  window.scrollTo(0, 0);
  render();
}

function render() {
  const fn = VISTAS[vistaActual];
  if (fn) fn();
}

$('#tabbar').addEventListener('click', ev => {
  const b = ev.target.closest('.tab');
  if (b) irA(b.dataset.tab);
});

function barra(titulo, sub, acciones) {
  $('#appbar-title').textContent = titulo;
  $('#appbar-sub').textContent = sub || '';
  $('#appbar-actions').innerHTML = (acciones || [])
    .map(a => `<button type="button" data-act="${a.act}">${esc(a.txt)}</button>`).join('');
  $('#appbar-logo').innerHTML = marcaHTML();
  $('#appbar-logo').title = S.config.negocio || '';
}
ACC.irAjustes = () => irA('ajustes');

/* --------------------------------------------- identidad de la empresa */

/** Iniciales para cuando todavia no hay foto: "Distribuidora David" -> DD */
function iniciales(nombre) {
  const p = String(nombre || '').trim().split(/\s+/).filter(Boolean);
  if (!p.length) return '·';
  return ((p[0][0] || '') + (p[1] ? p[1][0] : '')).toUpperCase();
}
function marcaHTML() {
  return S.config.logo
    ? `<img src="${S.config.logo}" alt="">`
    : esc(iniciales(S.config.negocio));
}

/** Achica la foto antes de guardarla: el almacenamiento del navegador es
 *  pequeño y una foto de camara moderna no cabe. */
function procesarFoto(archivo, listo) {
  if (!archivo) return;
  if (!/^image\//.test(archivo.type)) { toast('Elige una imagen'); return; }
  const lector = new FileReader();
  lector.onload = () => {
    const img = new Image();
    img.onload = () => {
      const MAX = 420;
      const escala = Math.min(1, MAX / Math.max(img.width, img.height));
      const w = Math.max(1, Math.round(img.width * escala));
      const h = Math.max(1, Math.round(img.height * escala));
      const lienzo = document.createElement('canvas');
      lienzo.width = w; lienzo.height = h;
      const ctx = lienzo.getContext('2d');
      ctx.fillStyle = '#ffffff';          // fondo para logos con transparencia
      ctx.fillRect(0, 0, w, h);
      ctx.drawImage(img, 0, 0, w, h);
      listo(lienzo.toDataURL('image/jpeg', 0.85));
    };
    img.onerror = () => toast('No se pudo leer esa imagen');
    img.src = lector.result;
  };
  lector.onerror = () => toast('No se pudo leer el archivo');
  lector.readAsDataURL(archivo);
}

/* ------------------------------------------------ pantalla de bienvenida */

function abrirBienvenida() {
  abrirModal(`<div class="modal" id="modal-bienvenida">
    <div class="modal-head"><h2>Bienvenido</h2></div>
    <div class="modal-body">
      <div class="card"><div class="card-body">
        <p class="small muted" style="margin-top:0">Ponle el nombre y la cara de tu negocio.
          Aparecerán al abrir la app y en cada factura que entregues.</p>

        <div class="avatar-fila">
          <div class="avatar" id="bv-avatar">${marcaHTML()}</div>
          <div>
            <button class="btn btn-sm" data-act="elegirFoto">Elegir foto</button>
            <div class="small muted" style="margin-top:4px">Tu logo o una foto. Opcional.</div>
          </div>
        </div>

        <div class="field"><label for="bv-negocio">Nombre de la empresa</label>
          <input id="bv-negocio" type="text" placeholder="Ej: Distribuidora David"
                 value="${esc(S.config.negocio === 'Mi distribución' ? '' : S.config.negocio)}"></div>
        <div class="field"><label for="bv-contacto">Teléfono o NIT (opcional)</label>
          <input id="bv-contacto" type="text" value="${esc(S.config.contacto)}"></div>

        <button class="btn btn-primary btn-block" data-act="terminarBienvenida">Empezar</button>
      </div></div>
      <p class="small muted" style="text-align:center">Todo esto se puede cambiar después en Ajustes.</p>
    </div>
  </div>`);
}

ACC.elegirFoto = () => {
  let inp = $('#selector-foto');
  if (!inp) {
    inp = document.createElement('input');
    inp.type = 'file';
    inp.accept = 'image/*';
    inp.id = 'selector-foto';
    inp.hidden = true;
    inp.addEventListener('change', () => {
      procesarFoto(inp.files && inp.files[0], dataUri => {
        S.config.logo = dataUri;
        guardar();
        $$('#bv-avatar, #aj-avatar').forEach(a => { a.innerHTML = marcaHTML(); });
        $('#appbar-logo').innerHTML = marcaHTML();
        toast('Foto guardada');
      });
      inp.value = '';
    });
    document.body.appendChild(inp);
  }
  inp.click();
};

ACC.quitarFoto = () => {
  S.config.logo = '';
  guardar();
  render();
  $('#appbar-logo').innerHTML = marcaHTML();
  toast('Foto quitada');
};

ACC.terminarBienvenida = () => {
  const nombre = $('#bv-negocio').value.trim();
  if (!nombre) { toast('Escribe el nombre de tu empresa'); $('#bv-negocio').focus(); return; }
  S.config.negocio = nombre;
  S.config.contacto = $('#bv-contacto').value.trim();
  S.config.configurado = true;
  guardar();
  cerrarModal();
  render();
  toast('Listo, ' + nombre);
};

/* Pantalla de entrada: se ve un momento al abrir la app. */
function mostrarEntrada() {
  const sp = $('#splash');
  $('#splash-logo').innerHTML = marcaHTML();
  $('#splash-nombre').textContent = S.config.negocio || '';
  sp.hidden = false;
  const irse = () => {
    sp.classList.add('se-va');
    setTimeout(() => { sp.hidden = true; sp.classList.remove('se-va'); }, 380);
  };
  sp.addEventListener('click', irse, { once: true });
  setTimeout(irse, 1100);
}

/* =======================================================================
   VISTA 1 - PEDIDOS DEL DIA
   ======================================================================= */

let uiFecha = hoyISO();   // fecha ancla del periodo que se esta viendo
let uiTipo = 'dia';       // dia | semana | mes | anio

/** Fecha que se propone al crear un pedido desde el periodo que se ve. */
function fechaParaPedidoNuevo() {
  if (uiTipo === 'dia') return uiFecha;
  const { desde, hasta } = rangoPeriodo(uiTipo, uiFecha);
  const h = hoyISO();
  return (h >= desde && h <= hasta) ? h : desde;   // hoy si cae dentro; si no, el inicio
}

/** ‹ Etiqueta › con el subtitulo solo cuando aporta algo: "Agosto 2026" no
 *  necesita explicacion, pero "Hoy" y "Esta semana" si dicen poco solos. */
function navPeriodo(tipo, ancla, accMover, accHoy) {
  const sub = (tipo === 'dia' || (tipo === 'semana' && periodoEsActual(tipo, ancla)))
    ? subPeriodo(tipo, ancla) : '';
  return `<div class="card"><div class="card-body periodo-nav">
      <button class="btn btn-sm" data-act="${accMover}" data-dir="-1" aria-label="Anterior">‹</button>
      <div class="periodo-titulo">
        <b>${esc(etiquetaPeriodo(tipo, ancla))}</b>
        ${sub ? `<span>${esc(sub)}</span>` : ''}
      </div>
      <button class="btn btn-sm" data-act="${accMover}" data-dir="1" aria-label="Siguiente">›</button>
      ${!periodoEsActual(tipo, ancla) ? `<button class="btn btn-sm" data-act="${accHoy}">Hoy</button>` : ''}
    </div></div>`;
}

function selectorPeriodo(tipo, ancla, accTipo, accMover, accHoy) {
  return `<div class="chips">
      ${Object.entries(NOMBRE_TIPO).map(([k, t]) => `<button class="chip ${tipo === k ? 'is-active' : ''}"
        data-act="${accTipo}" data-k="${k}">${t}</button>`).join('')}
    </div>
    ${navPeriodo(tipo, ancla, accMover, accHoy)}`;
}

function filaPedido(p, conFecha) {
  const entregado = p.estado === 'entregado';
  return `<li><button class="row" data-act="verPedido" data-id="${p.id}">
    <div class="row-main">
      <div class="row-title">${esc(p.cliente)}</div>
      <div class="row-sub">${conFecha ? esc(fechaCorta(p.fecha)) + ' · ' : ''}${num(unidadesPedido(p))} und ·
        <span class="badge ${entregado ? 'badge-ok' : 'badge-pend'}">${entregado ? 'Entregado' : 'Pendiente'}</span>
      </div>
    </div>
    <div class="row-end"><div class="row-amount">${dinero(totalPedido(p))}</div>
      <div class="row-sub">#${p.num}</div></div>
  </button></li>`;
}

/** La relación de pedidos del periodo. En día y semana se ven uno por uno;
 *  en mes se agrupan por día; en año, por mes (tocar un mes lo abre). */
function relacionPedidos(tipo, desde, hasta) {
  const lista = pedidosDe(desde, hasta);
  if (!lista.length) {
    return `<div class="empty"><strong>No hay pedidos en este periodo</strong>
      ${tipo === 'dia' ? 'Toca <b>+ Nuevo pedido</b> para registrar el primero.' : 'Prueba con otra fecha.'}</div>`;
  }

  if (tipo === 'dia') return `<ul class="list">${lista.map(p => filaPedido(p, false)).join('')}</ul>`;

  if (tipo === 'anio') {
    const meses = new Map();
    for (const p of lista) {
      const k = p.fecha.slice(0, 7);
      const a = meses.get(k) || { clave: k, pedidos: 0, cant: 0, valor: 0 };
      a.pedidos++; a.cant += unidadesPedido(p); a.valor += totalPedido(p);
      meses.set(k, a);
    }
    const filas = Array.from(meses.values()).sort((a, b) => b.clave.localeCompare(a.clave));
    return `<ul class="list">${filas.map(m => {
      const d = fechaDe(m.clave + '-01');
      return `<li><button class="row" data-act="abrirMes" data-f="${m.clave}-01">
        <div class="row-main">
          <div class="row-title">${cap(MESES[d.getMonth()])}</div>
          <div class="row-sub">${num(m.pedidos)} pedido${m.pedidos === 1 ? '' : 's'} · ${num(m.cant)} und</div>
        </div>
        <div class="row-end"><div class="row-amount">${dinero(m.valor)}</div>
          <div class="row-sub">ver el mes ›</div></div>
      </button></li>`;
    }).join('')}</ul>`;
  }

  // semana y mes: agrupados por dia, el mas reciente arriba
  const dias = new Map();
  for (const p of lista) {
    if (!dias.has(p.fecha)) dias.set(p.fecha, []);
    dias.get(p.fecha).push(p);
  }
  const fechas = Array.from(dias.keys()).sort((a, b) => b.localeCompare(a));
  return fechas.map(f => {
    const delDia = dias.get(f);
    const val = delDia.reduce((s, p) => s + totalPedido(p), 0);
    const und = delDia.reduce((s, p) => s + unidadesPedido(p), 0);
    return `<div class="grupo-dia">
        <button class="grupo-dia-cab" data-act="abrirDia" data-f="${f}">
          <span>${cap(fechaLarga(f))}</span>
          <span class="num">${num(delDia.length)} ped · ${num(und)} und · ${dinero(val)}</span>
        </button>
        <ul class="list">${delDia.map(p => filaPedido(p, false)).join('')}</ul>
      </div>`;
  }).join('');
}

VISTAS.pedidos = function () {
  const { desde, hasta } = rangoPeriodo(uiTipo, uiFecha);
  barra('Pedidos', subPeriodo(uiTipo, uiFecha));
  $('#fab').hidden = false;

  const lista = pedidosDe(desde, hasta);
  const unidades = lista.reduce((s, p) => s + unidadesPedido(p), 0);
  const valor = lista.reduce((s, p) => s + totalPedido(p), 0);
  const porCobrar = lista.filter(p => p.estado !== 'entregado').reduce((s, p) => s + totalPedido(p), 0);
  const tiendas = new Set(lista.map(p => norm(p.cliente))).size;

  $('#view-pedidos').innerHTML = `
    ${selectorPeriodo(uiTipo, uiFecha, 'uiTipoPeriodo', 'uiMover', 'irHoy')}

    <div class="kpis">
      <div class="kpi"><div class="kpi-label">Pedidos</div><div class="kpi-value num">${num(lista.length)}</div>
        <div class="kpi-note">${num(tiendas)} tienda${tiendas === 1 ? '' : 's'}</div></div>
      <div class="kpi"><div class="kpi-label">Unidades</div><div class="kpi-value num">${num(unidades)}</div></div>
      <div class="kpi"><div class="kpi-label">Valor</div><div class="kpi-value num">${dinero(valor)}</div></div>
      <div class="kpi"><div class="kpi-label">Por cobrar</div><div class="kpi-value num">${dinero(porCobrar)}</div>
        <div class="kpi-note">${porCobrar ? 'sin entregar' : 'todo entregado'}</div></div>
    </div>

    <div class="card">
      <div class="card-head"><h2>Relación de pedidos</h2>
        ${lista.length && uiTipo === 'dia'
          ? `<button class="btn btn-sm" data-act="imprimirFacturasDia">Imprimir todas</button>` : ''}
      </div>
      ${relacionPedidos(uiTipo, desde, hasta)}
    </div>

    ${lista.length ? `<div class="btn-row">
      <button class="btn btn-primary" data-act="verRelacion">${uiTipo === 'dia' ? 'Cuenta del día' : 'Relación en PDF'}</button>
      <button class="btn" data-act="verCargueDelPeriodo">Qué subir al camión</button>
    </div>` : ''}
  `;
};

ACC.uiTipoPeriodo = el => { uiTipo = el.dataset.k; render(); };
ACC.uiMover = el => { uiFecha = moverPeriodo(uiTipo, uiFecha, Number(el.dataset.dir)); render(); };
ACC.irHoy = () => { uiFecha = hoyISO(); render(); };
ACC.abrirMes = el => { uiTipo = 'mes'; uiFecha = el.dataset.f; render(); };
ACC.abrirDia = el => { uiTipo = 'dia'; uiFecha = el.dataset.f; render(); };
ACC.verCargueDelPeriodo = () => {
  const { desde, hasta } = rangoPeriodo(uiTipo, uiFecha);
  cargueRango = { desde, hasta, preset: desde === hasta ? 'dia' : 'rango' };
  irA('cargue');
};

$('#fab').addEventListener('click', () => abrirEditorPedido(null));

/* =======================================================================
   EDITOR DE PEDIDO
   ======================================================================= */

let borrador = null;   // { id, num, fecha, cliente, nota, items: Map(prodId -> cant), estado }
let editorFiltro = { texto: '', categoria: 'todas' };

function abrirEditorPedido(id, reemplazar, clientePrefijado) {
  const existente = id ? S.pedidos.find(p => p.id === id) : null;
  borrador = existente
    ? {
        id: existente.id, num: existente.num, fecha: existente.fecha,
        cliente: existente.cliente, nota: existente.nota || '', estado: existente.estado,
        items: new Map(existente.items.map(i => [i.prodId, i.cant]))
      }
    : {
        id: null, num: null, fecha: fechaParaPedidoNuevo(), cliente: clientePrefijado || '', nota: '',
        estado: 'pendiente', items: new Map()
      };
  editorFiltro = { texto: '', categoria: 'todas' };

  if (!S.productos.length) {
    cerrarModal();
    irA('catalogo');
    toast('Primero agrega productos al catálogo');
    return;
  }
  abrirModal(pintarEditor(), reemplazar ? 'reemplazo' : undefined);
}

function categorias() {
  const set = new Set(S.productos.map(p => p.categoria || 'Sin categoría'));
  return Array.from(set).sort((a, b) => a.localeCompare(b, 'es'));
}

function productosFiltrados() {
  const t = norm(editorFiltro.texto);
  return S.productos
    .filter(p => p.activo !== false)
    .filter(p => editorFiltro.categoria === 'todas' || (p.categoria || 'Sin categoría') === editorFiltro.categoria)
    .filter(p => !t || norm(p.nombre).includes(t) || norm(p.categoria).includes(t))
    .sort((a, b) => (a.categoria || '').localeCompare(b.categoria || '', 'es') || a.nombre.localeCompare(b.nombre, 'es'));
}

function pintarEditor() {
  const cats = categorias();
  return `
  <div class="modal" id="modal-pedido">
    <div class="modal-head">
      <button type="button" data-act="cerrarSubmodal">Cancelar</button>
      <h2>${borrador.id ? 'Editar pedido' : 'Nuevo pedido'}</h2>
    </div>
    <div class="modal-body">
      <div class="card"><div class="card-body">
        <div class="field">
          <label for="ed-cliente">Tienda o persona</label>
          <input id="ed-cliente" type="text" list="dl-clientes" autocomplete="off"
                 placeholder="Ej: Tienda La Esquina" value="${esc(borrador.cliente)}"
                 data-in="edCliente">
          <datalist id="dl-clientes">
            ${S.clientes.map(c => `<option value="${esc(c.nombre)}"></option>`).join('')}
          </datalist>
        </div>
        <div class="field-row">
          <div class="field">
            <label for="ed-fecha">Fecha de entrega</label>
            <input id="ed-fecha" type="date" value="${borrador.fecha}" data-ch="edFecha">
          </div>
        </div>
        <div class="field" style="margin-bottom:0">
          <label for="ed-nota">Nota (opcional)</label>
          <input id="ed-nota" type="text" placeholder="Ej: cobrar el viernes" value="${esc(borrador.nota)}" data-in="edNota">
        </div>
      </div></div>

      <div class="search-bar">
        <input type="search" placeholder="Buscar producto…" value="${esc(editorFiltro.texto)}" data-in="edBuscar">
      </div>
      <div class="chips">
        <button class="chip ${editorFiltro.categoria === 'todas' ? 'is-active' : ''}" data-act="edCat" data-cat="todas">Todas</button>
        ${cats.map(c => `<button class="chip ${editorFiltro.categoria === c ? 'is-active' : ''}"
            data-act="edCat" data-cat="${esc(c)}">${esc(c)}</button>`).join('')}
      </div>

      <div class="card"><div class="card-body tight" id="ed-lista">${pintarListaProductos()}</div></div>
    </div>
    <div class="modal-foot">
      <button class="btn btn-sm" data-act="verResumen" id="ed-resumen-btn">Resumen</button>
      <div class="total" id="ed-total">${resumenTexto()}</div>
      <button class="btn btn-primary" data-act="guardarPedido">Guardar</button>
    </div>
  </div>`;
}

function pintarListaProductos() {
  const lista = productosFiltrados();
  if (!lista.length) return `<div class="empty">Ningún producto coincide con la búsqueda.</div>`;
  let html = '';
  let catActual = null;
  for (const p of lista) {
    const cat = p.categoria || 'Sin categoría';
    if (cat !== catActual && editorFiltro.categoria === 'todas' && !editorFiltro.texto) {
      catActual = cat;
      html += `<div class="section-title" style="margin:12px 14px 4px">${esc(cat)}</div>`;
    }
    html += filaProducto(p);
  }
  return html;
}

function filaProducto(p) {
  const cant = borrador.items.get(p.id) || 0;
  return `<div class="prod-row ${cant ? 'has-qty' : ''}" id="fila-${p.id}">
    <div class="row-main">
      <div class="row-title">${esc(p.nombre)}</div>
      <div class="row-sub">${dinero(p.precio)} · ${esc(p.unidad || 'paq')}</div>
    </div>
    <div class="stepper">
      <button type="button" data-act="edMenos" data-id="${p.id}" aria-label="Quitar uno">−</button>
      <input type="number" inputmode="numeric" min="0" step="1" value="${cant}"
             data-ch="edCant" data-id="${p.id}" aria-label="Cantidad de ${esc(p.nombre)}">
      <button type="button" class="plus" data-act="edMas" data-id="${p.id}" aria-label="Agregar uno">+</button>
    </div>
  </div>`;
}

function itemsDelBorrador() {
  const out = [];
  for (const [prodId, cant] of borrador.items) {
    if (!cant) continue;
    const p = productoPorId(prodId);
    if (!p) continue;
    out.push({ prodId, nombre: p.nombre, precio: p.precio || 0, unidad: p.unidad || 'paq', categoria: p.categoria || '', cant });
  }
  return out.sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));
}

function resumenTexto() {
  const items = itemsDelBorrador();
  const und = items.reduce((s, i) => s + i.cant, 0);
  const val = items.reduce((s, i) => s + i.cant * i.precio, 0);
  return `<b>${dinero(val)}</b><span>${items.length} producto${items.length === 1 ? '' : 's'} · ` +
         `${num(und)} unidad${und === 1 ? '' : 'es'}</span>`;
}

function refrescarPie() {
  const pie = $('#ed-total');
  if (pie) pie.innerHTML = resumenTexto();
}

function ajustarCantidad(prodId, valor) {
  const v = Math.max(0, Math.round(Number(valor) || 0));
  if (v) borrador.items.set(prodId, v); else borrador.items.delete(prodId);
  const fila = document.getElementById('fila-' + prodId);
  if (fila) {
    fila.classList.toggle('has-qty', v > 0);
    const inp = $('input', fila);
    if (inp && Number(inp.value) !== v) inp.value = v;
  }
  refrescarPie();
}

ACC.edCliente = el => { borrador.cliente = el.value; };
ACC.edNota    = el => { borrador.nota = el.value; };
ACC.edFecha   = el => { if (el.value) borrador.fecha = el.value; };
ACC.edBuscar  = el => { editorFiltro.texto = el.value; $('#ed-lista').innerHTML = pintarListaProductos(); };
ACC.edCat     = el => {
  editorFiltro.categoria = el.dataset.cat;
  $$('#modal-pedido .chip').forEach(c => c.classList.toggle('is-active', c.dataset.cat === el.dataset.cat));
  $('#ed-lista').innerHTML = pintarListaProductos();
};
ACC.edMas   = el => ajustarCantidad(el.dataset.id, (borrador.items.get(el.dataset.id) || 0) + 1);
ACC.edMenos = el => ajustarCantidad(el.dataset.id, (borrador.items.get(el.dataset.id) || 0) - 1);
ACC.edCant  = el => ajustarCantidad(el.dataset.id, el.value);

function pintarResumenItems() {
  const items = itemsDelBorrador();
  if (!items.length) return `<div class="empty">El pedido quedó vacío. Vuelve y agrega productos.</div>`;
  return items.map(i => `<div class="prod-row has-qty">
      <div class="row-main"><div class="row-title">${esc(i.nombre)}</div>
        <div class="row-sub">${num(i.cant)} × ${dinero(i.precio)} = ${dinero(i.cant * i.precio)}</div></div>
      <div class="stepper">
        <button type="button" data-act="resMenos" data-id="${i.prodId}" aria-label="Quitar uno">−</button>
        <input type="number" inputmode="numeric" min="0" value="${i.cant}" data-ch="resCant" data-id="${i.prodId}"
               aria-label="Cantidad de ${esc(i.nombre)}">
        <button type="button" class="plus" data-act="resMas" data-id="${i.prodId}" aria-label="Agregar uno">+</button>
      </div>
    </div>`).join('');
}

ACC.verResumen = () => {
  if (!itemsDelBorrador().length) { toast('Todavía no has agregado productos'); return; }
  abrirModal(`<div class="modal" id="modal-resumen">
      <div class="modal-head"><button type="button" data-act="cerrarSubmodal">Volver</button><h2>Resumen del pedido</h2></div>
      <div class="modal-body"><div class="card"><div class="card-body tight" id="res-lista">${pintarResumenItems()}</div></div></div>
      <div class="modal-foot"><div class="total" id="res-total">${resumenTexto()}</div>
        <button class="btn btn-primary" data-act="cerrarSubmodal">Listo</button></div>
    </div>`, 'sub');
};

function ajustarDesdeResumen(prodId, v) {
  ajustarCantidad(prodId, v);   // mantiene sincronizada la lista de atras
  const lista = $('#res-lista');
  if (lista) lista.innerHTML = pintarResumenItems();
  const total = $('#res-total');
  if (total) total.innerHTML = resumenTexto();
}
ACC.resMas   = el => ajustarDesdeResumen(el.dataset.id, (borrador.items.get(el.dataset.id) || 0) + 1);
ACC.resMenos = el => ajustarDesdeResumen(el.dataset.id, (borrador.items.get(el.dataset.id) || 0) - 1);
ACC.resCant  = el => ajustarDesdeResumen(el.dataset.id, el.value);

ACC.guardarPedido = () => {
  const cliente = (borrador.cliente || '').trim();
  const items = itemsDelBorrador();
  if (!cliente) { toast('Escribe el nombre de la tienda'); $('#ed-cliente')?.focus(); return; }
  if (!items.length) { toast('Agrega al menos un producto'); return; }

  // cliente nuevo -> queda en la libreta para la proxima vez
  if (!S.clientes.some(c => norm(c.nombre) === norm(cliente))) {
    S.clientes.push({ id: uid(), nombre: cliente, zona: '', telefono: '' });
  }

  if (borrador.id) {
    const p = S.pedidos.find(x => x.id === borrador.id);
    Object.assign(p, { fecha: borrador.fecha, cliente, nota: borrador.nota, items });
  } else {
    S.pedidos.push({
      id: uid(), num: S.config.consecutivo++, fecha: borrador.fecha, cliente,
      nota: borrador.nota, items, estado: 'pendiente', creado: Date.now()
    });
  }
  guardar();
  uiFecha = borrador.fecha;
  cerrarModal();
  irA('pedidos');
  toast(borrador.id ? 'Pedido actualizado' : 'Pedido guardado');
};

/* =======================================================================
   FACTURA / DETALLE DEL PEDIDO
   ======================================================================= */

function encabezadoMarca() {
  return `<div class="factura-marca">
      <div class="factura-logo">${marcaHTML()}</div>
      <div><h3>${esc(S.config.negocio || 'Mi distribución')}</h3>
        ${S.config.contacto ? `<div class="small muted">${esc(S.config.contacto)}</div>` : ''}</div>
    </div>`;
}

function facturaHTML(p) {
  const total = totalPedido(p);
  return `<div class="factura">
    <div class="factura-head">
      ${encabezadoMarca()}
      <div class="meta"><b>Pedido N° ${p.num}</b><br>${fechaFactura(p.fecha)}<br>
        ${p.estado === 'entregado' ? 'ENTREGADO' : 'PENDIENTE'}</div>
    </div>
    <div class="factura-cliente">Cliente:<br><b>${esc(p.cliente)}</b></div>
    <table>
      <thead><tr><th class="n">Cant</th><th>Producto</th><th class="n">V. unit</th><th class="n">Total</th></tr></thead>
      <tbody>
        ${p.items.map(i => `<tr>
          <td class="n">${num(i.cant)}</td>
          <td>${esc(i.nombre)}${i.unidad ? ` <span class="muted small">(${esc(i.unidad)})</span>` : ''}</td>
          <td class="n">${dinero(i.precio)}</td>
          <td class="n">${dinero(i.cant * i.precio)}</td></tr>`).join('')}
      </tbody>
      <tfoot><tr><td class="n">${num(unidadesPedido(p))}</td><td colspan="2">TOTAL</td>
        <td class="n">${dinero(total)}</td></tr></tfoot>
    </table>
    ${p.nota ? `<div class="factura-nota">Nota: ${esc(p.nota)}</div>` : ''}
    <div class="factura-firma"><div>Entregado por</div><div>Recibido por</div></div>
  </div>`;
}

function textoPedido(p) {
  const l = [];
  l.push(`*${S.config.negocio || 'Pedido'}* — Pedido N° ${p.num}`);
  l.push(`${fechaFactura(p.fecha)}`);
  l.push(`Cliente: ${p.cliente}`);
  l.push('');
  for (const i of p.items) l.push(`• ${i.cant} × ${i.nombre} — ${dinero(i.cant * i.precio)}`);
  l.push('');
  l.push(`TOTAL: ${dinero(totalPedido(p))}`);
  if (p.nota) l.push(`Nota: ${p.nota}`);
  return l.join('\n');
}

ACC.verPedido = el => {
  const p = S.pedidos.find(x => x.id === el.dataset.id);
  if (!p) return;
  const entregado = p.estado === 'entregado';
  abrirModal(`<div class="modal">
    <div class="modal-head"><button type="button" data-act="cerrarSubmodal">Cerrar</button><h2>Pedido N° ${p.num}</h2></div>
    <div class="modal-body">
      ${facturaHTML(p)}
      <div class="btn-row" style="margin-top:14px">
        <button class="btn" data-act="compartirPedido" data-id="${p.id}">Compartir</button>
        <button class="btn" data-act="imprimirPedido" data-id="${p.id}">Imprimir / PDF</button>
      </div>
      <div class="btn-row" style="margin-top:8px">
        <button class="btn" data-act="editarPedido" data-id="${p.id}">Editar</button>
        <button class="btn ${entregado ? '' : 'btn-primary'}" data-act="alternarEntrega" data-id="${p.id}">
          ${entregado ? 'Marcar pendiente' : 'Marcar entregado'}</button>
      </div>
      <button class="btn btn-danger btn-block" style="margin-top:8px" data-act="borrarPedido" data-id="${p.id}">Eliminar pedido</button>
    </div>
  </div>`);
};

// reemplaza la factura por el editor sin tocar el historial
ACC.editarPedido = el => abrirEditorPedido(el.dataset.id, true);

ACC.alternarEntrega = el => {
  const p = S.pedidos.find(x => x.id === el.dataset.id);
  p.estado = p.estado === 'entregado' ? 'pendiente' : 'entregado';
  guardar(); cerrarModal(); render();
  toast(p.estado === 'entregado' ? 'Marcado como entregado' : 'Marcado como pendiente');
};

ACC.borrarPedido = el => {
  const p = S.pedidos.find(x => x.id === el.dataset.id);
  if (!confirm(`¿Eliminar el pedido N° ${p.num} de ${p.cliente}? No se puede deshacer.`)) return;
  S.pedidos = S.pedidos.filter(x => x.id !== p.id);
  guardar(); cerrarModal(); render();
  toast('Pedido eliminado');
};

ACC.compartirPedido = el => {
  const p = S.pedidos.find(x => x.id === el.dataset.id);
  compartir(`Pedido ${p.cliente}`, textoPedido(p));
};

ACC.imprimirPedido = el => {
  const p = S.pedidos.find(x => x.id === el.dataset.id);
  imprimir(facturaHTML(p));
};

ACC.imprimirFacturasDia = () => {
  const lista = pedidosDe(uiFecha, uiFecha);
  if (!lista.length) return;
  imprimir(lista.map(facturaHTML).join(''));
};

/* =======================================================================
   CUENTA DEL DIA - el cierre, tienda por tienda y producto por producto
   ======================================================================= */

const TITULO_DOC = { dia: 'CUENTA DEL DÍA', semana: 'RELACIÓN SEMANAL', mes: 'RELACIÓN MENSUAL', anio: 'RELACIÓN ANUAL' };

/** Resumen por tienda: en un solo dia interesa cada pedido; en periodos
 *  largos interesa cuanto compro cada tienda en total. */
function resumenTiendas(desde, hasta) {
  const mapa = new Map();
  for (const p of pedidosDe(desde, hasta)) {
    const k = norm(p.cliente);
    const a = mapa.get(k) || { nombre: p.cliente, pedidos: 0, cant: 0, valor: 0, pendiente: 0 };
    a.pedidos++;
    a.cant += unidadesPedido(p);
    a.valor += totalPedido(p);
    if (p.estado !== 'entregado') a.pendiente += totalPedido(p);
    mapa.set(k, a);
  }
  return Array.from(mapa.values()).sort((a, b) => b.valor - a.valor);
}

function relacionHTML(tipo, ancla) {
  const { desde, hasta } = rangoPeriodo(tipo, ancla);
  const unDia = desde === hasta;
  const lista = pedidosDe(desde, hasta);
  const productos = totalesPorProducto(desde, hasta);
  const unidades = productos.reduce((s, t) => s + t.cant, 0);
  const valor = lista.reduce((s, p) => s + totalPedido(p), 0);
  const entregados = lista.filter(p => p.estado === 'entregado');
  const cobrado = entregados.reduce((s, p) => s + totalPedido(p), 0);
  const tiendas = resumenTiendas(desde, hasta);

  // en periodos largos, la evolucion: por dia en semana/mes, por mes en año
  const serie = unDia ? [] : serieDelPeriodo(tipo, desde, hasta).filter(x => x.pedidos);

  return `<div class="factura">
    <div class="factura-head">
      ${encabezadoMarca()}
      <div class="meta"><b>${TITULO_DOC[tipo]}</b><br>
        ${unDia ? fechaFactura(desde) : `${fechaFactura(desde)}<br>al ${fechaFactura(hasta)}`}</div>
    </div>

    <div class="doc-totales">
      <div class="doc-tot"><span>Pedidos</span><b>${num(lista.length)}</b></div>
      <div class="doc-tot"><span>Unidades</span><b>${num(unidades)}</b></div>
      <div class="doc-tot"><span>Tiendas</span><b>${num(tiendas.length)}</b></div>
    </div>

    ${unDia ? `
    <div class="doc-seccion">Pedidos del día</div>
    <table>
      <thead><tr><th>N°</th><th>Tienda</th><th class="n">Und</th><th class="n">Valor</th><th>Estado</th></tr></thead>
      <tbody>${lista.map(p => `<tr>
        <td class="n">${p.num}</td>
        <td>${esc(p.cliente)}</td>
        <td class="n">${num(unidadesPedido(p))}</td>
        <td class="n">${dinero(totalPedido(p))}</td>
        <td>${p.estado === 'entregado' ? 'Entregado' : 'Pendiente'}</td></tr>`).join('')}</tbody>
      <tfoot><tr><td colspan="2">TOTAL</td><td class="n">${num(unidades)}</td>
        <td class="n">${dinero(valor)}</td><td></td></tr></tfoot>
    </table>` : `
    <div class="doc-seccion">Por tienda</div>
    <table>
      <thead><tr><th>Tienda</th><th class="n">Ped</th><th class="n">Und</th><th class="n">Valor</th><th class="n">Debe</th></tr></thead>
      <tbody>${tiendas.map(t => `<tr>
        <td>${esc(t.nombre)}</td>
        <td class="n">${num(t.pedidos)}</td>
        <td class="n">${num(t.cant)}</td>
        <td class="n">${dinero(t.valor)}</td>
        <td class="n">${t.pendiente ? dinero(t.pendiente) : '—'}</td></tr>`).join('')}</tbody>
      <tfoot><tr><td>TOTAL</td><td class="n">${num(lista.length)}</td><td class="n">${num(unidades)}</td>
        <td class="n">${dinero(valor)}</td><td class="n">${dinero(valor - cobrado)}</td></tr></tfoot>
    </table>`}

    ${serie.length ? `
    <div class="doc-seccion">${tipo === 'anio' ? 'Mes a mes' : 'Día a día'}</div>
    <table>
      <thead><tr><th>${tipo === 'anio' ? 'Mes' : 'Día'}</th><th class="n">Ped</th>
        <th class="n">Und</th><th class="n">Valor</th></tr></thead>
      <tbody>${serie.map(d => `<tr><td>${esc(d.etiquetaLarga)}</td>
        <td class="n">${num(d.pedidos)}</td><td class="n">${num(d.cant)}</td>
        <td class="n">${dinero(d.valor)}</td></tr>`).join('')}</tbody>
    </table>` : ''}

    <div class="doc-seccion">Por producto</div>
    <table>
      <thead><tr><th class="n">Cant</th><th>Producto</th><th class="n">Valor</th></tr></thead>
      <tbody>${productos.map(t => `<tr>
        <td class="n">${num(t.cant)}</td><td>${esc(t.nombre)}</td>
        <td class="n">${dinero(t.valor)}</td></tr>`).join('')}</tbody>
    </table>

    <div class="doc-seccion">Caja</div>
    <table>
      <tbody>
        <tr><td>Cobrado (pedidos entregados)</td><td class="n">${dinero(cobrado)}</td></tr>
        <tr><td>Por cobrar (pendientes)</td><td class="n">${dinero(valor - cobrado)}</td></tr>
      </tbody>
    </table>

    <div class="doc-gran-total"><span>TOTAL ${unDia ? 'DEL DÍA' : esc(subPeriodo(tipo, ancla).toUpperCase())}</span>
      <b>${dinero(valor)}</b></div>
    <div class="factura-firma"><div>Elaborado por</div><div>Revisado por</div></div>
  </div>`;
}

function textoRelacion(tipo, ancla) {
  const { desde, hasta } = rangoPeriodo(tipo, ancla);
  const unDia = desde === hasta;
  const lista = pedidosDe(desde, hasta);
  const productos = totalesPorProducto(desde, hasta);
  const valor = lista.reduce((s, p) => s + totalPedido(p), 0);
  const l = [`*${S.config.negocio || ''}*`.trim(),
    `${cap(TITULO_DOC[tipo].toLowerCase())} — ${subPeriodo(tipo, ancla)}`, ''];
  l.push('TIENDAS:');
  if (unDia) {
    for (const p of lista) l.push(`• ${p.cliente}: ${dinero(totalPedido(p))}${p.estado === 'entregado' ? '' : ' (pendiente)'}`);
  } else {
    for (const t of resumenTiendas(desde, hasta)) {
      l.push(`• ${t.nombre}: ${num(t.pedidos)} ped · ${dinero(t.valor)}${t.pendiente ? ` (debe ${dinero(t.pendiente)})` : ''}`);
    }
  }
  l.push('', 'PRODUCTOS:');
  for (const t of productos) l.push(`• ${t.cant} × ${t.nombre}`);
  l.push('', `TOTAL: ${dinero(valor)}`);
  l.push(`${lista.length} pedido${lista.length === 1 ? '' : 's'} · ${num(productos.reduce((s, t) => s + t.cant, 0))} unidades`);
  return l.join('\n');
}

ACC.verRelacion = () => {
  const { desde, hasta } = rangoPeriodo(uiTipo, uiFecha);
  if (!pedidosDe(desde, hasta).length) { toast('Este periodo no tiene pedidos'); return; }
  abrirModal(`<div class="modal">
    <div class="modal-head"><button type="button" data-act="cerrarSubmodal">Cerrar</button>
      <h2>${uiTipo === 'dia' ? 'Cuenta del día' : 'Relación de pedidos'}</h2></div>
    <div class="modal-body">
      ${relacionHTML(uiTipo, uiFecha)}
      <div class="btn-row" style="margin-top:14px">
        <button class="btn" data-act="compartirRelacion">Compartir</button>
        <button class="btn" data-act="imprimirRelacion">Imprimir / PDF</button>
      </div>
    </div>
  </div>`);
};
ACC.compartirRelacion = () => compartir('Relación de pedidos', textoRelacion(uiTipo, uiFecha));
ACC.imprimirRelacion = () => imprimir(relacionHTML(uiTipo, uiFecha));

/* =======================================================================
   VISTA 2 - CARGUE DEL CAMION
   ======================================================================= */

let cargueRango = { desde: hoyISO(), hasta: hoyISO(), preset: 'dia' };

function clavesCargue() { return cargueRango.desde + '_' + cargueRango.hasta; }

VISTAS.cargue = function () {
  const { desde, hasta } = cargueRango;
  barra('Cargue del camión', desde === hasta ? fechaLarga(desde) : `${fechaCorta(desde)} — ${fechaCorta(hasta)}`);
  $('#fab').hidden = true;

  const totales = totalesPorProducto(desde, hasta);
  const marcados = new Set(S.cargados[clavesCargue()] || []);
  const unidades = totales.reduce((s, t) => s + t.cant, 0);
  const valor = totales.reduce((s, t) => s + t.valor, 0);
  const nPedidos = pedidosDe(desde, hasta).length;

  // agrupado por categoria, para armar el camion por estante
  const grupos = new Map();
  for (const t of totales) {
    const g = grupos.get(t.categoria) || [];
    g.push(t);
    grupos.set(t.categoria, g);
  }

  const cuerpo = Array.from(grupos.entries()).map(([cat, items]) => `
    <div class="section-title" style="margin:14px 14px 2px">${esc(cat)}</div>
    ${items.map(t => {
      const id = t.id || ('n:' + norm(t.nombre));
      const done = marcados.has(id);
      return `<label class="check-row ${done ? 'done' : ''}">
        <input type="checkbox" ${done ? 'checked' : ''} data-ch="marcarCargado" data-id="${esc(id)}">
        <span class="qty-big">${num(t.cant)}</span>
        <span class="row-main"><span class="row-title">${esc(t.nombre)}</span>
          <span class="row-sub">${esc(t.unidad)} · ${t.clientes.size} cliente${t.clientes.size === 1 ? '' : 's'}</span></span>
        <span class="row-end row-sub">${dinero(t.valor)}</span>
      </label>`;
    }).join('')}`).join('');

  $('#view-cargue').innerHTML = `
    <div class="chips">
      <button class="chip ${cargueRango.preset === 'dia' && cargueRango.desde === hoyISO() ? 'is-active' : ''}" data-act="cargHoy">Hoy</button>
      <button class="chip ${cargueRango.preset === 'dia' && cargueRango.desde === sumarDias(hoyISO(), 1) ? 'is-active' : ''}" data-act="cargManana">Mañana</button>
      <button class="chip ${cargueRango.preset === 'pend' ? 'is-active' : ''}" data-act="cargPendientes">Todo lo pendiente</button>
      <button class="chip ${cargueRango.preset === 'rango' ? 'is-active' : ''}" data-act="cargRango">Elegir fechas</button>
    </div>

    ${cargueRango.preset === 'rango' ? `<div class="card"><div class="card-body field-row">
      <div class="field"><label>Desde</label><input type="date" value="${desde}" data-ch="cargDesde"></div>
      <div class="field"><label>Hasta</label><input type="date" value="${hasta}" data-ch="cargHasta"></div>
    </div></div>` : ''}

    <div class="kpis">
      <div class="kpi"><div class="kpi-label">Productos</div><div class="kpi-value num">${num(totales.length)}</div></div>
      <div class="kpi"><div class="kpi-label">Unidades</div><div class="kpi-value num">${num(unidades)}</div>
        <div class="kpi-note">de ${num(nPedidos)} pedido${nPedidos === 1 ? '' : 's'}</div></div>
    </div>

    <div class="card">
      <div class="card-head"><h2>Subir al camión</h2>
        <span class="small muted num" id="cargue-contador">${num(marcados.size)}/${num(totales.length)} listos</span></div>
      ${totales.length ? `<div class="card-body tight">${cuerpo}</div>
        <div class="card-head" style="border-top:1px solid var(--line);border-bottom:0">
          <b>Total del cargue</b><b class="num">${dinero(valor)}</b></div>`
        : `<div class="empty"><strong>Nada que cargar</strong>No hay pedidos en estas fechas.</div>`}
    </div>

    ${totales.length ? `
    <div class="btn-row">
      <button class="btn" data-act="compartirCargue">Compartir lista</button>
      <button class="btn" data-act="imprimirCargue">Imprimir</button>
    </div>
    <button class="btn btn-block" style="margin-top:8px" data-act="limpiarMarcas">Quitar las marcas</button>
    ` : ''}
  `;
};

ACC.cargHoy = () => { cargueRango = { desde: hoyISO(), hasta: hoyISO(), preset: 'dia' }; render(); };
ACC.cargManana = () => { const m = sumarDias(hoyISO(), 1); cargueRango = { desde: m, hasta: m, preset: 'dia' }; render(); };
ACC.cargPendientes = () => {
  const pend = S.pedidos.filter(p => p.estado === 'pendiente');
  if (!pend.length) { toast('No hay pedidos pendientes'); return; }
  const fechas = pend.map(p => p.fecha).sort();
  cargueRango = { desde: fechas[0], hasta: fechas[fechas.length - 1], preset: 'pend' };
  render();
};
ACC.cargRango = () => { cargueRango = { ...cargueRango, preset: 'rango' }; render(); };
ACC.cargDesde = el => { if (el.value) { cargueRango.desde = el.value; if (cargueRango.hasta < el.value) cargueRango.hasta = el.value; render(); } };
ACC.cargHasta = el => { if (el.value) { cargueRango.hasta = el.value; if (cargueRango.desde > el.value) cargueRango.desde = el.value; render(); } };

/* Se actualiza el DOM a mano en vez de redibujar: si la lista se redibuja,
   el scroll salta al inicio y marcar 30 productos se vuelve imposible. */
ACC.marcarCargado = el => {
  const clave = clavesCargue();
  const set = new Set(S.cargados[clave] || []);
  if (el.checked) set.add(el.dataset.id); else set.delete(el.dataset.id);
  S.cargados[clave] = Array.from(set);
  guardar();
  el.closest('.check-row').classList.toggle('done', el.checked);
  const cont = $('#cargue-contador');
  if (cont) cont.textContent = `${num(set.size)}/${num($$('#view-cargue .check-row').length)} listos`;
};
ACC.limpiarMarcas = () => { delete S.cargados[clavesCargue()]; guardar(); render(); };

function textoCargue() {
  const { desde, hasta } = cargueRango;
  const totales = totalesPorProducto(desde, hasta);
  const l = [`*CARGUE — ${S.config.negocio || ''}*`.trim(),
    desde === hasta ? fechaLarga(desde) : `${fechaCorta(desde)} al ${fechaCorta(hasta)}`, ''];
  let cat = null;
  for (const t of totales) {
    if (t.categoria !== cat) { cat = t.categoria; l.push(`— ${cat} —`); }
    l.push(`${t.cant} × ${t.nombre}`);
  }
  l.push('');
  l.push(`Total: ${num(totales.reduce((s, t) => s + t.cant, 0))} unidades · ${dinero(totales.reduce((s, t) => s + t.valor, 0))}`);
  return l.join('\n');
}
ACC.compartirCargue = () => compartir('Cargue', textoCargue());

ACC.imprimirCargue = () => {
  const { desde, hasta } = cargueRango;
  const totales = totalesPorProducto(desde, hasta);
  imprimir(`<div class="factura">
    <div class="factura-head">
      <div><h3>Cargue del camión</h3><div class="small muted">${esc(S.config.negocio || '')}</div></div>
      <div class="meta">${desde === hasta ? fechaFactura(desde) : fechaFactura(desde) + ' al ' + fechaFactura(hasta)}</div>
    </div>
    <table>
      <thead><tr><th class="n">Cant</th><th>Producto</th><th>Categoría</th><th class="n">Valor</th></tr></thead>
      <tbody>${totales.map(t => `<tr><td class="n">${num(t.cant)}</td><td>${esc(t.nombre)}</td>
        <td>${esc(t.categoria)}</td><td class="n">${dinero(t.valor)}</td></tr>`).join('')}</tbody>
      <tfoot><tr><td class="n">${num(totales.reduce((s, t) => s + t.cant, 0))}</td><td colspan="2">TOTAL</td>
        <td class="n">${dinero(totales.reduce((s, t) => s + t.valor, 0))}</td></tr></tfoot>
    </table>
  </div>`);
};

/* =======================================================================
   VISTA 3 - REPORTES Y PROYECCION
   ======================================================================= */

let repTipo = 'semana';      // dia | semana | mes | anio | rango
let repAncla = hoyISO();
let repRango = { desde: sumarDias(hoyISO(), -13), hasta: hoyISO() };
let repVentana = 28;         // dias de historial para proyectar
let repMetrica = 'unidades'; // unidades | valor

function rangoDelPeriodo() {
  return repTipo === 'rango' ? repRango : rangoPeriodo(repTipo, repAncla);
}
/** Los graficos del año van por mes; los demas por dia, salvo que el rango
 *  elegido a mano sea tan largo que no quepan las barras. */
function granularidadReporte() {
  if (repTipo === 'anio') return 'anio';
  const { desde, hasta } = rangoDelPeriodo();
  return diasEntre(desde, hasta) > 92 ? 'anio' : 'dia';
}

VISTAS.reportes = function () {
  const { desde, hasta } = rangoDelPeriodo();
  barra('Reportes', repTipo === 'rango'
    ? `${fechaCorta(desde)} — ${fechaCorta(hasta)}`
    : subPeriodo(repTipo, repAncla));
  $('#fab').hidden = true;

  const pedidos = pedidosDe(desde, hasta);
  const productos = totalesPorProducto(desde, hasta);
  const clientes = totalesPorCliente(desde, hasta);
  const serie = serieDelPeriodo(granularidadReporte(), desde, hasta);
  const unidades = productos.reduce((s, t) => s + t.cant, 0);
  const valor = productos.reduce((s, t) => s + t.valor, 0);
  const ticket = pedidos.length ? valor / pedidos.length : 0;

  $('#view-reportes').innerHTML = `
    <div class="chips">
      ${Object.entries(NOMBRE_TIPO).map(([k, t]) => `<button class="chip ${repTipo === k ? 'is-active' : ''}"
        data-act="repTipo" data-k="${k}">${t}</button>`).join('')}
      <button class="chip ${repTipo === 'rango' ? 'is-active' : ''}" data-act="repTipo" data-k="rango">Fechas</button>
    </div>

    ${repTipo === 'rango' ? `<div class="card"><div class="card-body field-row">
      <div class="field"><label>Desde</label><input type="date" value="${desde}" data-ch="repDesde"></div>
      <div class="field"><label>Hasta</label><input type="date" value="${hasta}" data-ch="repHasta"></div>
    </div></div>` : navPeriodo(repTipo, repAncla, 'repMover', 'repHoy')}

    <div class="kpis">
      <div class="kpi"><div class="kpi-label">Pedidos</div><div class="kpi-value num">${num(pedidos.length)}</div></div>
      <div class="kpi"><div class="kpi-label">Unidades</div><div class="kpi-value num">${num(unidades)}</div></div>
      <div class="kpi"><div class="kpi-label">Ventas</div><div class="kpi-value num">${dinero(valor)}</div></div>
      <div class="kpi"><div class="kpi-label">Pedido promedio</div><div class="kpi-value num">${dinero(ticket)}</div></div>
    </div>

    ${pedidos.length ? `
      ${desde === hasta ? '' : tarjetaTendencia(serie)}
      ${tarjetaTopProductos(productos, unidades)}
      ${tarjetaProyeccion()}
      ${tarjetaClientes(clientes)}
      <div class="btn-row">
        <button class="btn" data-act="csvDetalle">Exportar detalle CSV</button>
        <button class="btn" data-act="csvResumen">Exportar resumen CSV</button>
      </div>`
    : `<div class="card"><div class="empty"><strong>Sin datos en este periodo</strong>
        Registra pedidos y aquí verás cuánto vendiste y de qué.</div></div>`}
  `;
};

ACC.repTipo = el => { repTipo = el.dataset.k; render(); };
ACC.repMover = el => { repAncla = moverPeriodo(repTipo, repAncla, Number(el.dataset.dir)); render(); };
ACC.repHoy = () => { repAncla = hoyISO(); render(); };
ACC.repDesde = el => { if (el.value) { repRango.desde = el.value; if (repRango.hasta < el.value) repRango.hasta = el.value; render(); } };
ACC.repHasta = el => { if (el.value) { repRango.hasta = el.value; if (repRango.desde > el.value) repRango.desde = el.value; render(); } };
ACC.repMetrica = el => { repMetrica = el.dataset.k; render(); };
ACC.repVentana = el => { repVentana = Number(el.dataset.k); render(); };

/* --- grafico 1: tendencia por dia (barras verticales, una sola serie) --- */

function tarjetaTendencia(serie) {
  const campo = repMetrica === 'valor' ? 'valor' : 'cant';
  const porMes = granularidadReporte() === 'anio';
  const max = Math.max(1, ...serie.map(d => d[campo]));
  const pico = serie.reduce((a, b) => (b[campo] > a[campo] ? b : a), serie[0]);
  // con muchas barras no cabe una etiqueta en cada una: se rotula solo el pico
  const paso = serie.length > 14 ? Math.ceil(serie.length / 7) : 1;

  const barras = serie.map(d => {
    const v = d[campo];
    const alto = (v / max) * 100;
    const esPico = v > 0 && d.clave === pico.clave;
    return `<div class="bar-v ${v ? '' : 'is-zero'}" tabindex="0" data-act="tipDia"
              data-f="${esc(d.etiquetaLarga)}" data-c="${d.cant}" data-v="${d.valor}" data-p="${d.pedidos}">
        ${esPico ? `<span class="bar-v-top">${campo === 'valor' ? dinero(v) : num(v)}</span>` : ''}
        <span class="bar-v-fill" style="height:${Math.max(alto, v ? 3 : 0.8)}%"></span>
      </div>`;
  }).join('');

  const ejes = serie.map((d, i) => `<span>${i % paso === 0 ? esc(d.etiqueta) : ''}</span>`).join('');

  return `<div class="card">
    <div class="card-head"><h2>${porMes ? 'Mes a mes' : 'Día a día'}</h2>
      <div style="display:flex;gap:6px">
        <button class="chip ${repMetrica === 'unidades' ? 'is-active' : ''}" data-act="repMetrica" data-k="unidades">Unidades</button>
        <button class="chip ${repMetrica === 'valor' ? 'is-active' : ''}" data-act="repMetrica" data-k="valor">Ventas</button>
      </div>
    </div>
    <div class="chart">
      <p class="chart-note">${repMetrica === 'valor' ? 'Ventas' : 'Unidades'} por ${porMes ? 'mes' : 'día'} ·
        máximo ${repMetrica === 'valor' ? dinero(max) : num(max)}</p>
      <div class="bars-v">${barras}</div>
      <div class="bars-axis">${ejes}</div>
      <p class="chart-tip" id="tip-dia">Toca una barra para ver el detalle.</p>
    </div>
  </div>`;
}

ACC.tipDia = el => {
  const t = $('#tip-dia');
  if (!t) return;
  t.textContent = `${el.dataset.f}: ${num(el.dataset.c)} unidades · ` +
    `${dinero(Number(el.dataset.v))} · ${num(el.dataset.p)} pedido${el.dataset.p === '1' ? '' : 's'}`;
};

/* --- grafico 2: productos mas movidos (barras horizontales) --- */

function tarjetaTopProductos(productos, unidades) {
  const top = productos.slice(0, 8);
  const max = Math.max(1, ...top.map(t => t.cant));
  const barras = top.map(t => `<div>
      <div class="bar-h-label"><span>${esc(t.nombre)}</span>
        <b>${num(t.cant)} · ${dinero(t.valor)}</b></div>
      <div class="bar-h-track"><div class="bar-h-fill" style="width:${(t.cant / max) * 100}%"></div></div>
    </div>`).join('');

  const filas = productos.map(t => `<tr>
      <td>${esc(t.nombre)}</td>
      <td class="n">${num(t.cant)}</td>
      <td class="n">${dinero(t.valor)}</td>
      <td class="n">${unidades ? Math.round(t.cant / unidades * 100) : 0}%</td>
    </tr>`).join('');

  return `<div class="card">
    <div class="card-head"><h2>Productos más vendidos</h2></div>
    <div class="chart"><div class="bars-h">${barras}</div></div>
    <details><summary style="padding:10px 14px;cursor:pointer;font-size:14px;font-weight:650;border-top:1px solid var(--line)">
      Ver la tabla completa (${productos.length})</summary>
      <div class="table-wrap"><table class="data">
        <thead><tr><th>Producto</th><th class="n">Und</th><th class="n">Valor</th><th class="n">%</th></tr></thead>
        <tbody>${filas}</tbody>
        <tfoot><tr><td>Total</td><td class="n">${num(unidades)}</td>
          <td class="n">${dinero(productos.reduce((s, t) => s + t.valor, 0))}</td><td class="n">100%</td></tr></tfoot>
      </table></div>
    </details>
  </div>`;
}

/* --- proyeccion semanal y mensual ------------------------------------- */

function calcularProyeccion(ventana) {
  const hasta = hoyISO();
  const desde = sumarDias(hasta, -(ventana - 1));
  const pedidos = pedidosDe(desde, hasta);
  if (!pedidos.length) return null;

  const fechas = Array.from(new Set(pedidos.map(p => p.fecha))).sort();
  const primera = fechas[0];
  // solo se promedia sobre dias que realmente pudieron tener venta
  const diasCalendario = Math.max(1, diasEntre(primera, hasta) + 1);
  const diasActivos = fechas.length;

  const productos = totalesPorProducto(desde, hasta).map(t => ({
    nombre: t.nombre,
    cant: t.cant,
    valor: t.valor,
    porDia: t.cant / diasCalendario,
    semana: t.cant / diasCalendario * 7,
    mes: t.cant / diasCalendario * 30
  }));

  const unidades = productos.reduce((s, t) => s + t.cant, 0);
  const valor = productos.reduce((s, t) => s + t.valor, 0);

  return {
    desde: primera, hasta, diasCalendario, diasActivos,
    unidades, valor, productos,
    undDia: unidades / diasCalendario,
    valDia: valor / diasCalendario,
    undSemana: unidades / diasCalendario * 7,
    valSemana: valor / diasCalendario * 7,
    undMes: unidades / diasCalendario * 30,
    valMes: valor / diasCalendario * 30
  };
}

function tarjetaProyeccion() {
  const p = calcularProyeccion(repVentana);
  if (!p) return '';
  const filas = p.productos.slice(0, 30).map(t => `<tr>
      <td>${esc(t.nombre)}</td>
      <td class="n">${num(Math.round(t.semana))}</td>
      <td class="n">${num(Math.round(t.mes))}</td>
    </tr>`).join('');

  return `<div class="card">
    <div class="card-head"><h2>Proyección</h2>
      <div style="display:flex;gap:6px">
        ${[14, 28, 56].map(v => `<button class="chip ${repVentana === v ? 'is-active' : ''}"
            data-act="repVentana" data-k="${v}">${v}d</button>`).join('')}
      </div>
    </div>
    <div class="card-body">
      <div class="kpis" style="margin-bottom:6px">
        <div class="kpi"><div class="kpi-label">Próxima semana</div>
          <div class="kpi-value num">${num(Math.round(p.undSemana))} und</div>
          <div class="kpi-note">≈ ${dinero(p.valSemana)}</div></div>
        <div class="kpi"><div class="kpi-label">Próximo mes</div>
          <div class="kpi-value num">${num(Math.round(p.undMes))} und</div>
          <div class="kpi-note">≈ ${dinero(p.valMes)}</div></div>
      </div>
      ${p.diasCalendario < 7 ? `<p class="aviso">Ojo: apenas hay
        ${p.diasCalendario} día${p.diasCalendario === 1 ? '' : 's'} de historial. Estos números
        son una idea gruesa; se vuelven confiables después de dos o tres semanas registrando
        pedidos.</p>` : ''}
      <p class="small muted" style="margin:8px 2px 0">
        Calculado con ${p.diasCalendario} día${p.diasCalendario === 1 ? '' : 's'} de historial
        (${fechaCorta(p.desde)} a ${fechaCorta(p.hasta)}), de los cuales ${p.diasActivos}
        tuvieron pedidos. Promedio: ${num(Math.round(p.undDia * 10) / 10)} unidades y
        ${dinero(p.valDia)} por día. La proyección es ese promedio × 7 y × 30; no adivina
        temporadas ni festivos.
      </p>
    </div>
    <details><summary style="padding:10px 14px;cursor:pointer;font-size:14px;font-weight:650;border-top:1px solid var(--line)">
      Ver proyección por producto</summary>
      <div class="table-wrap"><table class="data">
        <thead><tr><th>Producto</th><th class="n">Semana</th><th class="n">Mes</th></tr></thead>
        <tbody>${filas}</tbody>
      </table></div>
    </details>
  </div>`;
}

function tarjetaClientes(clientes) {
  const filas = clientes.slice(0, 20).map(c => `<tr>
      <td>${esc(c.nombre)}</td><td class="n">${num(c.pedidos)}</td>
      <td class="n">${num(c.cant)}</td><td class="n">${dinero(c.valor)}</td></tr>`).join('');
  return `<div class="card">
    <div class="card-head"><h2>Clientes del periodo</h2><span class="small muted">${num(clientes.length)}</span></div>
    <div class="table-wrap"><table class="data">
      <thead><tr><th>Cliente</th><th class="n">Pedidos</th><th class="n">Und</th><th class="n">Valor</th></tr></thead>
      <tbody>${filas}</tbody></table></div>
  </div>`;
}

/* --- exportar a Excel (CSV con ; para configuraciones en español) ------ */

function csv(filas) {
  const linea = f => f.map(c => {
    const s = String(c == null ? '' : c);
    return /[;"\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }).join(';');
  return '\ufeff' + filas.map(linea).join('\r\n');
}
function descargar(nombre, texto, tipo) {
  const blob = new Blob([texto], { type: (tipo || 'text/csv') + ';charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = nombre;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

ACC.csvDetalle = () => {
  const { desde, hasta } = rangoDelPeriodo();
  const filas = [['Fecha', 'Pedido', 'Cliente', 'Producto', 'Categoria', 'Cantidad', 'Valor unitario', 'Total', 'Estado']];
  for (const p of pedidosDe(desde, hasta))
    for (const i of p.items)
      filas.push([p.fecha, p.num, p.cliente, i.nombre, i.categoria || '', i.cant, i.precio, i.cant * i.precio, p.estado]);
  descargar(`pedidos_${desde}_a_${hasta}.csv`, csv(filas));
  toast('CSV descargado');
};

ACC.csvResumen = () => {
  const { desde, hasta } = rangoDelPeriodo();
  const filas = [['Producto', 'Categoria', 'Unidades', 'Valor']];
  for (const t of totalesPorProducto(desde, hasta)) filas.push([t.nombre, t.categoria, t.cant, t.valor]);
  descargar(`resumen_${desde}_a_${hasta}.csv`, csv(filas));
  toast('CSV descargado');
};

/* =======================================================================
   VISTA 4 - CATALOGO (productos y clientes)
   ======================================================================= */

let catPestana = 'productos';
let catBusqueda = '';

function contenidoCatalogo() {
  const t = norm(catBusqueda);
  let contenido;

  if (catPestana === 'productos') {
    const lista = S.productos
      .filter(p => !t || norm(p.nombre).includes(t) || norm(p.categoria).includes(t))
      .sort((a, b) => (a.categoria || '').localeCompare(b.categoria || '', 'es') || a.nombre.localeCompare(b.nombre, 'es'));
    contenido = lista.length
      ? `<ul class="list">${lista.map(p => `<li><button class="row" data-act="editarProducto" data-id="${p.id}">
          <div class="row-main"><div class="row-title">${esc(p.nombre)}</div>
            <div class="row-sub">${esc(p.categoria || 'Sin categoría')} · ${esc(p.unidad || 'paq')}</div></div>
          <div class="row-end"><div class="row-amount">${dinero(p.precio)}</div></div></button></li>`).join('')}</ul>`
      : `<div class="empty"><strong>${S.productos.length ? 'Sin coincidencias' : 'Catálogo vacío'}</strong>
          ${S.productos.length ? 'Prueba con otra búsqueda.' : 'Agrega tus productos para empezar a tomar pedidos.'}</div>`;
  } else {
    const lista = S.clientes
      .filter(c => !t || norm(c.nombre).includes(t) || norm(c.zona).includes(t))
      .sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));
    contenido = lista.length
      ? `<ul class="list">${lista.map(c => {
          const hist = pedidosDeCliente(c.nombre);
          const gastado = hist.reduce((s, p) => s + totalPedido(p), 0);
          return `<li><button class="row" data-act="verCliente" data-id="${c.id}">
            <div class="row-main"><div class="row-title">${esc(c.nombre)}</div>
              <div class="row-sub">${esc(c.zona || 'Sin zona')} · ${num(hist.length)} factura${hist.length === 1 ? '' : 's'}</div></div>
            <div class="row-end"><div class="row-amount">${dinero(gastado)}</div></div>
          </button></li>`;
        }).join('')}</ul>`
      : `<div class="empty"><strong>Sin clientes</strong>Se agregan solos al guardar un pedido nuevo.</div>`;
  }
  return contenido;
}

VISTAS.catalogo = function () {
  barra('Catálogo', catPestana === 'productos'
    ? `${S.productos.length} producto${S.productos.length === 1 ? '' : 's'}`
    : `${S.clientes.length} cliente${S.clientes.length === 1 ? '' : 's'}`);
  $('#fab').hidden = true;

  $('#view-catalogo').innerHTML = `
    <div class="chips">
      <button class="chip ${catPestana === 'productos' ? 'is-active' : ''}" data-act="catTab" data-k="productos">Productos</button>
      <button class="chip ${catPestana === 'clientes' ? 'is-active' : ''}" data-act="catTab" data-k="clientes">Clientes</button>
    </div>
    <div class="search-bar"><input type="search" placeholder="Buscar…" value="${esc(catBusqueda)}" data-in="catBuscar"></div>
    <div class="btn-row" style="margin-bottom:12px">
      <button class="btn btn-primary" data-act="${catPestana === 'productos' ? 'nuevoProducto' : 'nuevoCliente'}">
        + ${catPestana === 'productos' ? 'Nuevo producto' : 'Nuevo cliente'}</button>
      ${catPestana === 'productos' && !S.productos.length
        ? `<button class="btn" data-act="cargarEjemplo">Cargar catálogo de ejemplo</button>` : ''}
    </div>
    <div class="card" id="cat-lista">${contenidoCatalogo()}</div>
  `;
};

ACC.catTab = el => { catPestana = el.dataset.k; catBusqueda = ''; render(); };
ACC.catBuscar = el => {
  // solo se redibuja la lista: el cursor se queda donde el usuario escribe
  catBusqueda = el.value;
  $('#cat-lista').innerHTML = contenidoCatalogo();
};

function formProducto(p) {
  return `<div class="modal">
    <div class="modal-head"><button type="button" data-act="cerrarSubmodal">Cancelar</button>
      <h2>${p ? 'Editar producto' : 'Nuevo producto'}</h2></div>
    <div class="modal-body"><div class="card"><div class="card-body">
      <div class="field"><label for="p-nombre">Nombre</label>
        <input id="p-nombre" type="text" placeholder="Ej: Papas de mayonesa" value="${esc(p ? p.nombre : '')}"></div>
      <div class="field"><label for="p-cat">Categoría</label>
        <input id="p-cat" type="text" list="dl-cats" placeholder="Ej: Papas" value="${esc(p ? p.categoria : '')}">
        <datalist id="dl-cats">${categorias().map(c => `<option value="${esc(c)}"></option>`).join('')}</datalist></div>
      <div class="field-row">
        <div class="field"><label for="p-precio">Precio</label>
          <input id="p-precio" type="number" inputmode="decimal" min="0" step="any" value="${p ? p.precio : ''}"></div>
        <div class="field"><label for="p-unidad">Unidad</label>
          <input id="p-unidad" type="text" list="dl-unid" placeholder="paquete" value="${esc(p ? p.unidad : 'paquete')}">
          <datalist id="dl-unid"><option value="paquete"></option><option value="unidad"></option>
            <option value="caja"></option><option value="bulto"></option><option value="display"></option></datalist></div>
      </div>
      <button class="btn btn-primary btn-block" data-act="guardarProducto" data-id="${p ? p.id : ''}">Guardar</button>
      ${p ? `<button class="btn btn-danger btn-block" style="margin-top:8px" data-act="borrarProducto" data-id="${p.id}">Eliminar</button>` : ''}
    </div></div></div>
  </div>`;
}

ACC.nuevoProducto = () => abrirModal(formProducto(null));
ACC.editarProducto = el => abrirModal(formProducto(productoPorId(el.dataset.id)));

ACC.guardarProducto = el => {
  const nombre = $('#p-nombre').value.trim();
  if (!nombre) { toast('Escribe el nombre'); return; }
  const datos = {
    nombre,
    categoria: $('#p-cat').value.trim() || 'Sin categoría',
    precio: Math.max(0, Number($('#p-precio').value) || 0),
    unidad: $('#p-unidad').value.trim() || 'paquete'
  };
  if (el.dataset.id) Object.assign(productoPorId(el.dataset.id), datos);
  else S.productos.push({ id: uid(), activo: true, ...datos });
  guardar(); cerrarModal(); render();
  toast('Producto guardado');
};

ACC.borrarProducto = el => {
  const p = productoPorId(el.dataset.id);
  const usos = S.pedidos.filter(x => x.items.some(i => i.prodId === p.id)).length;
  const aviso = usos
    ? `"${p.nombre}" aparece en ${usos} pedido(s). Esos pedidos conservan su información, pero el producto ya no se podrá pedir. ¿Eliminarlo?`
    : `¿Eliminar "${p.nombre}"?`;
  if (!confirm(aviso)) return;
  S.productos = S.productos.filter(x => x.id !== p.id);
  guardar(); cerrarModal(); render();
  toast('Producto eliminado');
};

function formCliente(c, desdeDetalle) {
  const cancelar = desdeDetalle
    ? `data-act="volverCliente" data-id="${c.id}"`
    : `data-act="cerrarSubmodal"`;
  return `<div class="modal">
    <div class="modal-head"><button type="button" ${cancelar}>Cancelar</button>
      <h2>${c ? 'Editar cliente' : 'Nuevo cliente'}</h2></div>
    <div class="modal-body"><div class="card"><div class="card-body">
      <div class="field"><label for="c-nombre">Tienda o persona</label>
        <input id="c-nombre" type="text" value="${esc(c ? c.nombre : '')}" placeholder="Ej: Tienda La Esquina"></div>
      <div class="field"><label for="c-zona">Zona o barrio</label>
        <input id="c-zona" type="text" value="${esc(c ? c.zona : '')}" placeholder="Ej: Centro"></div>
      <div class="field"><label for="c-tel">Teléfono</label>
        <input id="c-tel" type="tel" value="${esc(c ? c.telefono : '')}"></div>
      <button class="btn btn-primary btn-block" data-act="guardarCliente"
              data-id="${c ? c.id : ''}" data-volver="${desdeDetalle ? '1' : ''}">Guardar</button>
      ${c ? `<button class="btn btn-danger btn-block" style="margin-top:8px" data-act="borrarCliente" data-id="${c.id}">Eliminar</button>` : ''}
    </div></div></div>
  </div>`;
}

ACC.nuevoCliente = () => abrirModal(formCliente(null));
ACC.editarCliente = el => abrirModal(
  formCliente(S.clientes.find(c => c.id === el.dataset.id), el.dataset.desde === 'detalle'),
  el.dataset.desde === 'detalle' ? 'reemplazo' : undefined
);

/* -------------------------------------------------- historial por tienda */

const pedidosDeCliente = nombre => S.pedidos
  .filter(p => norm(p.cliente) === norm(nombre) && p.estado !== 'anulado')
  .sort((a, b) => b.fecha.localeCompare(a.fecha) || b.creado - a.creado);

function abrirDetalleCliente(id, reemplazar) {
  const c = S.clientes.find(x => x.id === id);
  if (!c) return;
  const hist = pedidosDeCliente(c.nombre);
  const total = hist.reduce((s, p) => s + totalPedido(p), 0);
  const pendiente = hist.filter(p => p.estado !== 'entregado').reduce((s, p) => s + totalPedido(p), 0);
  const unidades = hist.reduce((s, p) => s + unidadesPedido(p), 0);

  // que le gusta comprar a esta tienda
  const favoritos = new Map();
  for (const p of hist) for (const i of p.items) {
    favoritos.set(i.nombre, (favoritos.get(i.nombre) || 0) + i.cant);
  }
  const top = Array.from(favoritos.entries()).sort((a, b) => b[1] - a[1]).slice(0, 5);

  const filas = hist.map(p => `<li><button class="row" data-act="verPedido" data-id="${p.id}">
      <div class="row-main">
        <div class="row-title">${esc(fechaLarga(p.fecha))}</div>
        <div class="row-sub">Factura N° ${p.num} · ${num(unidadesPedido(p))} und ·
          <span class="badge ${p.estado === 'entregado' ? 'badge-ok' : 'badge-pend'}">
            ${p.estado === 'entregado' ? 'Entregado' : 'Pendiente'}</span></div>
      </div>
      <div class="row-end"><div class="row-amount">${dinero(totalPedido(p))}</div></div>
    </button></li>`).join('');

  abrirModal(`<div class="modal">
    <div class="modal-head"><button type="button" data-act="cerrarSubmodal">Cerrar</button>
      <h2>${esc(c.nombre)}</h2></div>
    <div class="modal-body">
      <div class="card"><div class="card-body">
        <div class="row-sub" style="margin-bottom:10px">
          ${esc(c.zona || 'Sin zona')}${c.telefono ? ' · ' + esc(c.telefono) : ''}</div>
        <div class="kpis" style="margin-bottom:0">
          <div class="kpi"><div class="kpi-label">Facturas</div><div class="kpi-value num">${num(hist.length)}</div></div>
          <div class="kpi"><div class="kpi-label">Unidades</div><div class="kpi-value num">${num(unidades)}</div></div>
          <div class="kpi"><div class="kpi-label">Total comprado</div><div class="kpi-value num">${dinero(total)}</div></div>
          <div class="kpi"><div class="kpi-label">Por cobrar</div>
            <div class="kpi-value num">${dinero(pendiente)}</div>
            <div class="kpi-note">${pendiente ? 'pedidos sin entregar' : 'al día'}</div></div>
        </div>
      </div></div>

      ${top.length ? `<div class="card">
        <div class="card-head"><h2>Lo que más lleva</h2></div>
        <div class="card-body">${top.map(([nom, cant]) =>
          `<div class="bar-h-label"><span>${esc(nom)}</span><b>${num(cant)}</b></div>`).join('')}</div>
      </div>` : ''}

      <div class="card">
        <div class="card-head"><h2>Facturas de esta tienda</h2><span class="small muted">${num(hist.length)}</span></div>
        ${hist.length ? `<ul class="list">${filas}</ul>`
          : `<div class="empty"><strong>Todavía no le has vendido</strong>Toma el primer pedido desde aquí.</div>`}
      </div>

      <div class="btn-row">
        <button class="btn btn-primary" data-act="pedidoParaCliente" data-nombre="${esc(c.nombre)}">Nuevo pedido</button>
        <button class="btn" data-act="editarCliente" data-id="${c.id}" data-desde="detalle">Editar datos</button>
      </div>
      ${hist.length ? `<div class="btn-row" style="margin-top:8px">
        <button class="btn" data-act="compartirCuentaCliente" data-id="${c.id}">Compartir estado de cuenta</button>
        <button class="btn" data-act="imprimirFacturasCliente" data-id="${c.id}">Imprimir sus facturas</button>
      </div>` : ''}
    </div>
  </div>`, reemplazar ? 'reemplazo' : undefined);
}

ACC.verCliente = el => abrirDetalleCliente(el.dataset.id);
ACC.volverCliente = el => abrirDetalleCliente(el.dataset.id, true);

ACC.pedidoParaCliente = el => abrirEditorPedido(null, true, el.dataset.nombre);

ACC.compartirCuentaCliente = el => {
  const c = S.clientes.find(x => x.id === el.dataset.id);
  const hist = pedidosDeCliente(c.nombre);
  const total = hist.reduce((s, p) => s + totalPedido(p), 0);
  const pend = hist.filter(p => p.estado !== 'entregado');
  const l = [`*${S.config.negocio || ''}*`.trim(), `Estado de cuenta — ${c.nombre}`, ''];
  for (const p of hist.slice(0, 20)) {
    l.push(`${fechaFactura(p.fecha)} · N° ${p.num} · ${dinero(totalPedido(p))}` +
      (p.estado === 'entregado' ? '' : ' (pendiente)'));
  }
  l.push('', `Total comprado: ${dinero(total)}`);
  if (pend.length) l.push(`Por cobrar: ${dinero(pend.reduce((s, p) => s + totalPedido(p), 0))}`);
  compartir('Estado de cuenta', l.join('\n'));
};

ACC.imprimirFacturasCliente = el => {
  const c = S.clientes.find(x => x.id === el.dataset.id);
  const hist = pedidosDeCliente(c.nombre);
  if (!hist.length) return;
  imprimir(hist.map(facturaHTML).join(''));
};

ACC.guardarCliente = el => {
  const nombre = $('#c-nombre').value.trim();
  if (!nombre) { toast('Escribe el nombre'); return; }
  const datos = { nombre, zona: $('#c-zona').value.trim(), telefono: $('#c-tel').value.trim() };
  if (el.dataset.id) {
    const c = S.clientes.find(x => x.id === el.dataset.id);
    // si le cambio el nombre, sus facturas viejas lo siguen
    if (norm(c.nombre) !== norm(nombre)) {
      for (const p of S.pedidos) if (norm(p.cliente) === norm(c.nombre)) p.cliente = nombre;
    }
    Object.assign(c, datos);
  } else {
    S.clientes.push({ id: uid(), ...datos });
  }
  guardar();
  render();
  toast('Cliente guardado');
  if (el.dataset.volver) abrirDetalleCliente(el.dataset.id, true);
  else cerrarModal();
};

ACC.borrarCliente = el => {
  const c = S.clientes.find(x => x.id === el.dataset.id);
  if (!confirm(`¿Eliminar a "${c.nombre}" de la libreta? Sus pedidos se conservan.`)) return;
  S.clientes = S.clientes.filter(x => x.id !== c.id);
  guardar(); cerrarModal(); render();
  toast('Cliente eliminado');
};

const EJEMPLO = [
  ['Papas de mayonesa', 'Papas', 1800, 'paquete'],
  ['Papas limón', 'Papas', 1800, 'paquete'],
  ['Papas pollo', 'Papas', 1800, 'paquete'],
  ['Papas naturales', 'Papas', 1800, 'paquete'],
  ['Cheetos', 'Snacks', 1500, 'paquete'],
  ['Doritos', 'Snacks', 2000, 'paquete'],
  ['Tostitos', 'Snacks', 2000, 'paquete'],
  ['Palomitas', 'Snacks', 1200, 'paquete'],
  ['Chicharrón', 'Snacks', 1500, 'paquete'],
  ['Platanitos', 'Snacks', 1500, 'paquete'],
  ['Galleta wafer', 'Galletas', 1000, 'paquete'],
  ['Galleta de sal', 'Galletas', 1200, 'paquete'],
  ['Ponqué', 'Panadería', 1500, 'unidad'],
  ['Chocolatina', 'Dulces', 800, 'unidad'],
  ['Bombones', 'Dulces', 500, 'paquete'],
  ['Gaseosa personal', 'Bebidas', 2500, 'unidad'],
  ['Agua 600 ml', 'Bebidas', 1500, 'unidad'],
  ['Jugo caja', 'Bebidas', 2000, 'unidad']
];

ACC.cargarEjemplo = () => {
  for (const [nombre, categoria, precio, unidad] of EJEMPLO)
    S.productos.push({ id: uid(), nombre, categoria, precio, unidad, activo: true });
  guardar(); render();
  toast('Catálogo de ejemplo cargado — edítalo a tu gusto');
};

/* =======================================================================
   VISTA 5 - AJUSTES
   ======================================================================= */

VISTAS.ajustes = function () {
  barra('Ajustes', '');
  $('#fab').hidden = true;
  const monedas = [['COP', 'Peso colombiano'], ['MXN', 'Peso mexicano'], ['PEN', 'Sol peruano'],
    ['CLP', 'Peso chileno'], ['ARS', 'Peso argentino'], ['GTQ', 'Quetzal'], ['DOP', 'Peso dominicano'],
    ['USD', 'Dólar'], ['EUR', 'Euro']];

  $('#view-ajustes').innerHTML = `
    <div class="card">
      <div class="card-head"><h2>Mi negocio</h2></div>
      <div class="card-body">
        <div class="avatar-fila">
          <div class="avatar" id="aj-avatar">${marcaHTML()}</div>
          <div>
            <div class="btn-row">
              <button class="btn btn-sm" data-act="elegirFoto">${S.config.logo ? 'Cambiar foto' : 'Poner foto'}</button>
              ${S.config.logo ? `<button class="btn btn-sm" data-act="quitarFoto">Quitar</button>` : ''}
            </div>
            <div class="small muted" style="margin-top:4px">Sale al abrir la app y en las facturas.</div>
          </div>
        </div>
        <div class="field"><label for="a-negocio">Nombre (sale en la factura)</label>
          <input id="a-negocio" type="text" value="${esc(S.config.negocio)}" data-in="cfgNegocio"></div>
        <div class="field"><label for="a-contacto">Teléfono o NIT</label>
          <input id="a-contacto" type="text" value="${esc(S.config.contacto)}" data-in="cfgContacto"></div>
        <div class="field" style="margin-bottom:0"><label for="a-moneda">Moneda</label>
          <select id="a-moneda" data-ch="cfgMoneda">
            ${monedas.map(([c, n]) => `<option value="${c}" ${S.config.moneda === c ? 'selected' : ''}>${c} — ${n}</option>`).join('')}
          </select></div>
      </div>
    </div>

    <div class="card">
      <div class="card-head"><h2>Respaldo</h2></div>
      <div class="card-body">
        <p class="small muted" style="margin-top:0">Los datos viven solo en este teléfono. Si cambias de
        equipo o borras el navegador, se pierden: exporta un respaldo de vez en cuando y guárdalo
        en tu correo o en Drive.</p>
        <div class="btn-row">
          <button class="btn" data-act="exportarTodo">Exportar respaldo</button>
          <button class="btn" data-act="importarTodo">Importar respaldo</button>
        </div>
        <input type="file" id="archivo-import" accept="application/json,.json" hidden data-ch="archivoElegido">
      </div>
    </div>

    <div class="card">
      <div class="card-head"><h2>Resumen</h2></div>
      <div class="table-wrap"><table class="data"><tbody>
        <tr><td>Productos</td><td class="n">${num(S.productos.length)}</td></tr>
        <tr><td>Clientes</td><td class="n">${num(S.clientes.length)}</td></tr>
        <tr><td>Pedidos registrados</td><td class="n">${num(S.pedidos.length)}</td></tr>
        <tr><td>Primer pedido</td><td class="n">${S.pedidos.length ? fechaFactura(S.pedidos.map(p => p.fecha).sort()[0]) : '—'}</td></tr>
      </tbody></table></div>
    </div>

    <div class="card">
      <div class="card-head"><h2>Instalar en el teléfono</h2></div>
      <div class="card-body small">
        <p style="margin-top:0"><b>Android (Chrome):</b> menú ⋮ → <i>Agregar a pantalla principal</i>.</p>
        <p><b>iPhone (Safari):</b> botón Compartir → <i>Agregar a inicio</i>.</p>
        <p style="margin-bottom:0">Queda como una app más: abre sin internet y sin navegador a la vista.</p>
        <div id="zona-instalar"></div>
      </div>
    </div>

    <button class="btn btn-danger btn-block" data-act="borrarTodo">Borrar todos los datos</button>
    <p class="small muted" style="text-align:center;margin-top:14px">Gestor de Pedidos · versión 1.0</p>
  `;
  pintarBotonInstalar();
};

ACC.cfgNegocio  = el => {
  S.config.negocio = el.value;
  guardar();
  if (!S.config.logo) $('#appbar-logo').innerHTML = marcaHTML();   // iniciales al vuelo
};
ACC.cfgContacto = el => { S.config.contacto = el.value; guardar(); };
ACC.cfgMoneda   = el => { S.config.moneda = el.value; guardar(); render(); };

ACC.exportarTodo = () => {
  descargar(`respaldo_pedidos_${hoyISO()}.json`, JSON.stringify(S, null, 2), 'application/json');
  toast('Respaldo descargado');
};
ACC.importarTodo = () => $('#archivo-import').click();
ACC.archivoElegido = el => {
  const f = el.files && el.files[0];
  if (!f) return;
  const lector = new FileReader();
  lector.onload = () => {
    try {
      const dato = JSON.parse(lector.result);
      if (!dato || !Array.isArray(dato.pedidos)) throw new Error('formato');
      if (!confirm('Esto reemplaza todos los datos actuales por los del respaldo. ¿Continuar?')) return;
      S = { ...estadoInicial(), ...dato, config: { ...estadoInicial().config, ...(dato.config || {}) } };
      guardar(); render();
      toast('Respaldo importado');
    } catch (e) {
      toast('El archivo no es un respaldo válido');
    }
  };
  lector.readAsText(f);
  el.value = '';
};

ACC.borrarTodo = () => {
  if (!confirm('Se borrarán productos, clientes y TODOS los pedidos. ¿Seguro?')) return;
  if (!confirm('Última confirmación: esto no se puede deshacer.')) return;
  S = estadoInicial();
  localStorage.removeItem(DB_KEY);
  guardar(); render();
  toast('Todo borrado');
  abrirBienvenida();
};

/* --- instalacion de la PWA */

let promptInstalar = null;
window.addEventListener('beforeinstallprompt', ev => {
  ev.preventDefault();
  promptInstalar = ev;
  pintarBotonInstalar();
});
function pintarBotonInstalar() {
  const zona = $('#zona-instalar');
  if (!zona || !promptInstalar) return;
  zona.innerHTML = `<button class="btn btn-primary btn-block" style="margin-top:10px" data-act="instalar">Instalar ahora</button>`;
}
ACC.instalar = async () => {
  if (!promptInstalar) return;
  promptInstalar.prompt();
  await promptInstalar.userChoice;
  promptInstalar = null;
  render();
};

/* =======================================================================
   MODALES, COMPARTIR E IMPRESION
   ======================================================================= */

/* Cada modal abierto empuja una entrada al historial: asi el boton "atras"
   del telefono cierra la pantalla en vez de salirse de la aplicacion.
   Cuando la app cierra un modal por su cuenta ya actualizo el DOM, asi que
   el popstate que llega despues hay que ignorarlo: si no, se lleva por
   delante la pantalla siguiente. */
let ignorarPop = false;

function retroceder(n) {
  if (!n) return;
  ignorarPop = true;
  history.go(-n);
  setTimeout(() => { ignorarPop = false; }, 500);   // red de seguridad
}

/** modo: undefined = pantalla nueva · 'sub' = encima de otra · 'reemplazo' = cambia la de arriba.
 *  Si ya hay una pantalla abierta, la nueva siempre se apila encima: al cerrarla
 *  se vuelve a la anterior (de la factura al historial de la tienda, por ejemplo). */
function abrirModal(html, modo) {
  const raiz = $('#modal-root');
  document.body.style.overflow = 'hidden';
  if (modo === 'reemplazo' && raiz.lastElementChild) {
    raiz.lastElementChild.outerHTML = html;   // no toca el historial
    return;
  }
  if (raiz.children.length) raiz.insertAdjacentHTML('beforeend', html);
  else raiz.innerHTML = html;
  history.pushState({ modal: raiz.children.length }, '');
}
function cerrarModal() {
  const abiertos = $('#modal-root').children.length;
  $('#modal-root').innerHTML = '';
  document.body.style.overflow = '';
  retroceder(abiertos);
}
function cerrarSubmodal() {
  const raiz = $('#modal-root');
  if (raiz.children.length > 1) { raiz.lastElementChild.remove(); retroceder(1); }
  else cerrarModal();
}
ACC.cerrarModal = cerrarModal;
ACC.cerrarSubmodal = cerrarSubmodal;

window.addEventListener('popstate', () => {
  if (ignorarPop) { ignorarPop = false; return; }
  const raiz = $('#modal-root');
  if (!raiz.children.length) return;
  raiz.lastElementChild.remove();
  if (!raiz.children.length) document.body.style.overflow = '';
});

async function compartir(titulo, texto) {
  try {
    if (navigator.share) { await navigator.share({ title: titulo, text: texto }); return; }
  } catch (e) {
    if (e && e.name === 'AbortError') return;
  }
  try {
    await navigator.clipboard.writeText(texto);
    toast('Copiado — pégalo en WhatsApp');
    return;
  } catch (e) { /* sigue al respaldo */ }
  window.open('https://wa.me/?text=' + encodeURIComponent(texto), '_blank');
}

function imprimir(html) {
  $('#print-root').innerHTML = html;
  setTimeout(() => window.print(), 60);
}
window.addEventListener('afterprint', () => { $('#print-root').innerHTML = ''; });

/* =======================================================================
   ARRANQUE
   ======================================================================= */

S = cargarEstado();
irA('pedidos');

if (S.config.configurado) mostrarEntrada();   // logo y nombre al abrir
else abrirBienvenida();                       // primera vez: quien eres

if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('sw.js').catch(e => console.warn('SW no registrado', e));
  });
}
