/* Logica de la app de empleo en el celular.
   Sin librerias: todo lo que hace es pedirle datos al PC y pintarlos. */

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

let vistaActual = "inicio";
let filtroEstado = "";
let buscando = false;
let temporizador = null;

/* ---------- Hablar con el PC ---------- */

async function pedir(ruta, opciones = {}) {
  const respuesta = await fetch(ruta, {
    ...opciones,
    headers: { "Content-Type": "application/json", ...(opciones.headers || {}) },
  });
  const datos = await respuesta.json().catch(() => ({}));
  if (!respuesta.ok) throw new Error(datos.error || `Error ${respuesta.status}`);
  return datos;
}

function conectado(si, texto) {
  const chip = $("#estado-conexion");
  chip.textContent = texto || (si ? "conectado" : "PC apagado");
  chip.classList.toggle("malo", !si);
}

function brindis(mensaje) {
  const caja = $("#brindis");
  caja.textContent = mensaje;
  caja.classList.remove("oculto");
  clearTimeout(caja._t);
  caja._t = setTimeout(() => caja.classList.add("oculto"), 3200);
}

function escapar(texto) {
  const d = document.createElement("div");
  d.textContent = texto == null ? "" : String(texto);
  return d.innerHTML;
}

/* ---------- Inicio ---------- */

async function cargarEstado() {
  let datos;
  try {
    datos = await pedir("/api/estado");
    conectado(true);
  } catch (e) {
    conectado(false);
    return;
  }

  const aviso = $("#aviso-perfil");
  if (!datos.perfil_ok) {
    aviso.textContent = "Falta configurar el perfil en el PC:\n" + datos.problema_perfil;
    aviso.classList.remove("oculto");
  } else {
    aviso.classList.add("oculto");
  }

  $("#tarjetas").innerHTML = datos.tarjetas.map((t) => `
    <button class="tarjeta ${t.estado === "pendiente_revision" && t.total > 0 ? "destacada" : ""}"
            data-estado="${t.estado}">
      <strong>${t.total}</strong>
      <span>${escapar(t.etiqueta)}</span>
    </button>`).join("");

  $$("#tarjetas .tarjeta").forEach((b) => {
    b.onclick = () => { filtroEstado = b.dataset.estado; irA("vacantes"); };
  });

  const hoy = datos.postuladas_hoy, tope = datos.tope_diario;
  $("#hoy").textContent = `${hoy} de ${tope}`;
  $("#barra-hoy").style.width = Math.min(100, tope ? (hoy / tope) * 100 : 0) + "%";
  $("#nota-tope").textContent =
    `Umbral: ${datos.umbral} puntos · Modo: ${datos.modo}` +
    (hoy >= tope ? " · Ya llegaste al tope de hoy." : "");

  pintarProgreso(datos.tarea);
}

function pintarProgreso(tarea) {
  const caja = $("#progreso");
  const activa = tarea.corriendo;

  if (buscando && !activa) {
    // Acaba de terminar: se refresca todo y se avisa.
    buscando = false;
    brindis(tarea.error ? "La búsqueda falló" : "Búsqueda terminada");
    if (vistaActual === "vacantes") cargarLista();
  }
  buscando = activa;

  $("#btn-simular").disabled = activa;
  $("#btn-real").disabled = activa;

  if (!activa && !tarea.lineas.length) { caja.classList.add("oculto"); return; }
  caja.classList.remove("oculto");

  $("#progreso-titulo").textContent = activa
    ? `Buscando… (modo ${tarea.modo})`
    : (tarea.error ? "Terminó con error" : `Terminó a las ${tarea.fin}`);

  const reloj = $("#progreso-reloj");
  reloj.textContent = activa ? "en curso" : (tarea.error ? "error" : "listo");
  reloj.classList.toggle("vivo", activa);
  reloj.classList.toggle("malo", Boolean(tarea.error));

  const texto = tarea.lineas.join("\n") + (tarea.error ? "\n\n" + tarea.error : "");
  const pre = $("#progreso-lineas");
  const abajo = pre.scrollHeight - pre.scrollTop - pre.clientHeight < 40;
  pre.textContent = texto;
  if (abajo) pre.scrollTop = pre.scrollHeight;
}

async function buscar(modo) {
  const mensaje = modo === "simulacion"
    ? "Va a buscar y calificar SIN enviar nada. ¿Seguimos?"
    : "Va a buscar y ENVIAR tu hoja de vida a las vacantes que pasen el filtro. ¿Seguimos?";
  if (!confirm(mensaje)) return;

  try {
    await pedir("/api/buscar", { method: "POST", body: JSON.stringify({ modo }) });
    buscando = true;
    brindis("El PC empezó a buscar");
    cargarEstado();
  } catch (e) {
    brindis(e.message);
  }
}

/* ---------- Vacantes ---------- */

const ESTADOS = [
  ["", "Todas"],
  ["pendiente_revision", "Para ti"],
  ["postulada", "Postuladas"],
  ["en_espera", "En espera"],
  ["descartada", "Descartadas"],
  ["simulada", "De prueba"],
  ["error", "Con error"],
];

function pintarFiltros() {
  $("#filtros-estado").innerHTML = ESTADOS.map(([valor, nombre]) =>
    `<button data-estado="${valor}" class="${valor === filtroEstado ? "activo" : ""}">${nombre}</button>`
  ).join("");
  $$("#filtros-estado button").forEach((b) => {
    b.onclick = () => { filtroEstado = b.dataset.estado; pintarFiltros(); cargarLista(); };
  });
}

async function cargarLista() {
  const q = $("#buscador").value.trim();
  const parametros = new URLSearchParams({ estado: filtroEstado, q });
  let datos;
  try {
    datos = await pedir("/api/lista?" + parametros);
    conectado(true);
  } catch (e) {
    conectado(false);
    return;
  }

  if (!datos.filas.length) {
    $("#lista").innerHTML = `<p class="vacio">No hay vacantes que mostrar aquí todavía.</p>`;
    return;
  }

  $("#lista").innerHTML = datos.filas.map((v) => `
    <article class="vacante">
      ${v.puntaje !== "" ? `<span class="puntaje">${escapar(v.puntaje)}</span>` : ""}
      <span class="insignia ${escapar(v.estado)}">${escapar(v.etiqueta)}</span>
      <h3>${escapar(v.titulo) || "(sin título)"}</h3>
      <div class="meta">
        ${escapar(v.empresa) || "empresa no publicada"}
        ${v.ciudad ? " · " + escapar(v.ciudad) : ""}
        ${v.portal ? " · " + escapar(v.portal) : ""}
      </div>
      ${v.detalle ? `<div class="detalle">${escapar(v.detalle)}</div>` : ""}
      <div class="pie">
        ${v.url ? `<a href="${escapar(v.url)}" target="_blank" rel="noopener">Abrir oferta</a>` : ""}
        ${v.estado === "pendiente_revision"
          ? `<button data-clave="${escapar(v.clave)}" class="marcar">Ya la hice</button>` : ""}
      </div>
    </article>`).join("");

  $$("#lista .marcar").forEach((b) => {
    b.onclick = async () => {
      try {
        await pedir("/api/marcar", {
          method: "POST",
          body: JSON.stringify({ clave: b.dataset.clave, estado: "revisada" }),
        });
        brindis("Marcada como hecha");
        cargarLista();
      } catch (e) { brindis(e.message); }
    };
  });
}

/* ---------- Filtro (perfil) ---------- */

const lineas = (t) => t.split("\n").map((s) => s.trim()).filter(Boolean);

async function cargarPerfil() {
  let p;
  try { p = await pedir("/api/perfil"); conectado(true); }
  catch (e) { conectado(false); return; }

  $("#f-cargos").value = (p.cargos_objetivo || []).join("\n");
  $("#f-deseables").value = (p.palabras_clave_deseables || []).join("\n");
  $("#f-excluyentes").value = (p.palabras_excluyentes || []).join("\n");
  $("#f-ciudades").value = (p.ciudades || []).join("\n");
  $("#f-remoto").checked = Boolean(p.acepta_remoto);
  $("#f-experiencia").value = p.anos_experiencia ?? 0;
  $("#f-salario").value = p.salario_minimo ?? 0;
  $("#f-nivel").value = p.nivel_educativo || "bachiller";
  $("#f-umbral").value = p.umbral_postulacion ?? 65;
  $("#f-tope").value = p.max_postulaciones_dia ?? 15;
  $("#f-modo").value = p.modo || "simulacion";
}

async function guardarPerfil() {
  const cambios = {
    cargos_objetivo: lineas($("#f-cargos").value),
    palabras_clave_deseables: lineas($("#f-deseables").value),
    palabras_excluyentes: lineas($("#f-excluyentes").value),
    ciudades: lineas($("#f-ciudades").value),
    acepta_remoto: $("#f-remoto").checked,
    anos_experiencia: Number($("#f-experiencia").value || 0),
    salario_minimo: Number($("#f-salario").value || 0),
    nivel_educativo: $("#f-nivel").value,
    umbral_postulacion: Number($("#f-umbral").value || 65),
    max_postulaciones_dia: Number($("#f-tope").value || 15),
    modo: $("#f-modo").value,
  };

  if (!cambios.cargos_objetivo.length) {
    brindis("Necesitas al menos un cargo para buscar");
    return;
  }

  $("#btn-guardar").disabled = true;
  try {
    const r = await pedir("/api/perfil", { method: "POST", body: JSON.stringify(cambios) });
    $("#aviso-guardado").textContent =
      `Guardado en el PC (${r.guardados.length} campos) · ${new Date().toLocaleTimeString()}`;
    brindis("Filtro guardado");
    cargarEstado();
  } catch (e) {
    brindis(e.message);
  } finally {
    $("#btn-guardar").disabled = false;
  }
}

/* ---------- Registro ---------- */

async function cargarLog() {
  try {
    const datos = await pedir("/api/log");
    conectado(true);
    const pre = $("#log");
    pre.textContent = datos.lineas.join("\n") || "(todavía no hay nada)";
    pre.scrollTop = pre.scrollHeight;
  } catch (e) { conectado(false); }
}

/* ---------- Navegacion ---------- */

function irA(vista) {
  vistaActual = vista;
  $$(".vista").forEach((s) => s.classList.toggle("activa", s.id === "vista-" + vista));
  $$(".pestana").forEach((b) => b.classList.toggle("activa", b.dataset.vista === vista));
  window.scrollTo(0, 0);

  if (vista === "vacantes") { pintarFiltros(); cargarLista(); }
  if (vista === "filtro") cargarPerfil();
  if (vista === "registro") cargarLog();
  if (vista === "inicio") cargarEstado();
  reprogramar();
}

function reprogramar() {
  clearInterval(temporizador);
  // Mientras busca conviene refrescar seguido; el resto del tiempo no
  // tiene sentido gastarle bateria al telefono.
  const cada = buscando ? 3000 : 20000;
  temporizador = setInterval(() => {
    if (document.hidden) return;
    if (vistaActual === "inicio" || buscando) cargarEstado();
    if (vistaActual === "registro") cargarLog();
  }, cada);
}

/* ---------- Arranque ---------- */

$$(".pestana").forEach((b) => { b.onclick = () => irA(b.dataset.vista); });
$("#btn-simular").onclick = () => buscar("simulacion");
$("#btn-real").onclick = () => buscar($("#f-modo").value === "automatico" ? "automatico" : "semiautomatico");
$("#btn-guardar").onclick = guardarPerfil;

let tecleo = null;
$("#buscador").oninput = () => { clearTimeout(tecleo); tecleo = setTimeout(cargarLista, 350); };

document.addEventListener("visibilitychange", () => { if (!document.hidden) irA(vistaActual); });

if ("serviceWorker" in navigator) {
  navigator.serviceWorker.register("/sw.js").catch(() => {});
}

cargarEstado();
reprogramar();
