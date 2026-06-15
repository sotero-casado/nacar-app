/* ============================================================
   NÁCAR Abogados — app móvil de consulta (solo lectura)
   Datos: OneDrive (documentos) + Outlook (agenda) en tiempo real,
   y exportaciones Excel de MN Program (clientes y expedientes).
   ============================================================ */
"use strict";

const CFG = window.NACAR_CONFIG || {};
const MODO_DEMO = !CFG.clientId;
const SCOPES = ["User.Read", "Files.Read.All", "Calendars.Read"];
const GRAPH = "https://graph.microsoft.com/v1.0";

/* ---------- estado global ---------- */
let CLIENTES = [];      // [{num,n,movil,tel,mail,exps:[...], _carpeta,_sub,_rootDocs}]
let EXPS = [];          // [{ci, cliente, e:{...}, docs:null|[{n,url}]}]
let CONTRA = [];        // [{nombre, exps:[idx]}]
let CITAS = [];         // [{inicio,fin,tipo,titulo,cliente,numexp,ciudad,esMN}]
let FUENTE = { fClientes: null, fExpedientes: null, usuario: "" };
let conIdx = {};
let historial = [];
let tabActual = "hoy";
let msalApp = null, cuenta = null;

/* ---------- utilidades ---------- */
function norm(s) {
  s = (s || "").normalize("NFD").replace(/[̀-ͯ]/g, "");
  s = s.replace(/^\d+_/, "").toLowerCase().replace(/[.,]/g, "");
  s = s.replace(/\b(sa|sl|slu|sau|s a|s l)\b/g, "");
  return s.replace(/\s+/g, " ").trim();
}
function esc(s) {
  return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function iniciales(n) {
  const p = (n || "?").replace(/,.*$/, "").split(/\s+/).filter(Boolean);
  return ((p[0] || "?")[0] + (p[1] ? p[1][0] : "")).toUpperCase();
}
function iconoDoc(f) {
  const e = (f || "").toLowerCase();
  if (e.endsWith(".zip")) return "ti-file-zip";
  if (e.endsWith(".mnmsg") || e.endsWith(".eml")) return "ti-mail";
  if (e.endsWith(".doc") || e.endsWith(".docx")) return "ti-file-text";
  if (/\.(jpe?g|png|gif|heic)$/.test(e)) return "ti-photo";
  return "ti-file";
}
function fmtFecha(s) {
  if (!s) return "";
  const d = new Date(s);
  return isNaN(d) ? "" : d.toLocaleDateString("es-ES", { day: "2-digit", month: "2-digit", year: "numeric" });
}
function fmtDia(d) {
  return d.toLocaleDateString("es-ES", { weekday: "long", day: "numeric", month: "long" })
    .replace(/^./, c => c.toUpperCase());
}
function fmtHora(d) {
  return d.toLocaleTimeString("es-ES", { hour: "2-digit", minute: "2-digit", timeZone: "Europe/Madrid" });
}
function mismoDia(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}
function claveDia(d) {
  return d.toLocaleDateString("sv-SE", { timeZone: "Europe/Madrid" });
}
function badgeEstado(estado) {
  if (estado === "Abierto") return '<span class="badge suave">Abierto</span>';
  if (estado === "Cerrado") return '<span class="badge gris">Cerrado</span>';
  return "";
}

/* ---------- clasificación de citas de Outlook ---------- */
function clasificarCita(asunto, cuerpo, inicio, fin) {
  const esMN = (cuerpo || "").indexOf("Datos MNprogram") >= 0;
  let cliente = "", resto = asunto || "";
  const m = /^(.*?)\.\s+(.*)$/.exec(asunto || "");
  if (m && esMN) { cliente = m[1].trim(); resto = m[2].trim(); }
  let tipo = "otra";
  if (esMN) {
    if (/^juicio/i.test(resto)) tipo = "juicio";
    else if (/^aportar/i.test(resto)) tipo = "plazo";
    else if (/^confesi/i.test(resto)) tipo = "confesion";
    else tipo = "vencimiento";
  }
  let numexp = "";
  const me = /Expediente:\s*([\d]+\/[\d]+)/.exec(cuerpo || "");
  if (me) numexp = me[1];
  let ciudad = "";
  const mc = /^([A-Za-zÀ-ÿ.\s]{3,25}?)[\r\n]/.exec(cuerpo || "");
  if (mc && mc[1].indexOf("<<<") < 0) ciudad = mc[1].trim();
  return { inicio, fin, tipo, titulo: resto, cliente, numexp, ciudad, esMN };
}
function etiquetaTipo(t) {
  return { juicio: "Juicio", plazo: "Plazo", confesion: "Confesión", vencimiento: "Vencimiento", otra: "Otra cita" }[t] || t;
}
function estiloTipo(t) {
  if (t === "juicio") return { bg: "var(--teal-bg)", color: "var(--teal)", icono: "ti-gavel" };
  if (t === "plazo" || t === "vencimiento") return { bg: "var(--ambar-bg)", color: "var(--ambar)", icono: "ti-file-upload" };
  if (t === "confesion") return { bg: "var(--teal-bg)", color: "var(--teal)", icono: "ti-message-question" };
  return { bg: "var(--fondo-2)", color: "var(--texto-2)", icono: "ti-users" };
}

/* ---------- ciudad de un expediente o cita ---------- */
function ciudadDeCita(c) {
  if (c.ciudad) return c.ciudad;
  const i = expedienteDeCita(c);
  if (i >= 0) {
    const mj = /\(([^)]+)\)\s*$/.exec(EXPS[i].e.juzgado || "");
    if (mj) return mj[1];
  }
  return "";
}
function expedienteDeCita(c) {
  if (!c.numexp) return -1;
  const candidatos = [];
  for (let i = 0; i < EXPS.length; i++) {
    if (EXPS[i].e.numexp === c.numexp) {
      if (c.cliente && norm(EXPS[i].cliente) === norm(c.cliente)) return i;
      candidatos.push(i);
    }
  }
  // sin cliente que coincida: solo es fiable si la referencia no está repetida
  return (!c.cliente && candidatos.length === 1) ? candidatos[0] : -1;
}

/* ---------- avisos ---------- */
function calcularAvisos() {
  const avisos = [];
  const ahora = new Date();
  const porDia = {};
  CITAS.forEach(c => {
    if (c.inicio < ahora) return;
    (porDia[claveDia(c.inicio)] = porDia[claveDia(c.inicio)] || []).push(c);
  });
  Object.keys(porDia).sort().slice(0, 14).forEach(k => {
    const juicios = porDia[k].filter(c => c.tipo === "juicio").sort((a, b) => a.inicio - b.inicio);
    for (let i = 1; i < juicios.length; i++) {
      if (juicios[i].inicio < juicios[i - 1].fin) {
        avisos.push(fmtDia(juicios[i].inicio) + ": dos juicios se solapan — " + fmtHora(juicios[i - 1].inicio) + " y " + fmtHora(juicios[i].inicio));
        break;
      }
    }
    const ciudades = [...new Set(juicios.map(ciudadDeCita).filter(Boolean))];
    if (ciudades.length > 1) {
      avisos.push(fmtDia(juicios[0].inicio) + ": juicios en " + ciudades.join(" y ") + " el mismo día");
    }
    porDia[k].filter(c => c.tipo === "plazo").forEach(c => {
      const dia = c.inicio.getDay();
      if (dia === 0 || dia === 6) avisos.push(fmtDia(c.inicio) + ": un plazo cae en fin de semana — revisar vencimiento real");
    });
  });
  if (FUENTE.fExpedientes) {
    const dias = Math.floor((ahora - FUENTE.fExpedientes) / 86400000);
    if (dias > (CFG.diasAvisoDatosViejos || 7)) {
      avisos.push("Los datos de MN Program tienen " + dias + " días — conviene repetir las exportaciones");
    }
  }
  return avisos;
}

/* ============================================================
   CAPA DE DATOS — Microsoft Graph
   ============================================================ */
async function token() {
  try {
    const r = await msalApp.acquireTokenSilent({ scopes: SCOPES, account: cuenta });
    return r.accessToken;
  } catch (e) {
    await msalApp.acquireTokenRedirect({ scopes: SCOPES });
  }
}
async function graph(ruta) {
  const t = await token();
  const r = await fetch(ruta.startsWith("http") ? ruta : GRAPH + ruta, { headers: { Authorization: "Bearer " + t, Prefer: 'outlook.timezone="Europe/Madrid"' } });
  if (!r.ok) throw new Error("Graph " + r.status + " en " + ruta);
  return r.json();
}
async function graphTodos(ruta) {
  let url = ruta, out = [];
  while (url) {
    const j = await graph(url);
    out = out.concat(j.value || []);
    url = j["@odata.nextLink"] || null;
  }
  return out;
}

/* ---------- exportaciones Excel ---------- */
function tipoDeExcel(rows) {
  if (!rows || !rows.length) return null;
  const hdr = rows[0].map(c => norm(String(c || "")));
  if (hdr.some(h => h.indexOf("numero cliente") >= 0)) return "clientes";
  if (hdr.some(h => h.indexOf("juzgado principal") >= 0 || h === "num exp")) return "expedientes";
  return null;
}
function col(hdr, nombre) {
  const n = norm(nombre);
  for (let i = 0; i < hdr.length; i++) if (norm(String(hdr[i] || "")).indexOf(n) >= 0) return i;
  return -1;
}
function parseClientes(rows) {
  const h = rows[0];
  const iNum = col(h, "Número Cliente"), iNom = col(h, "Nombre"), iMov = col(h, "Móvil"),
    iTel = col(h, "Teléfono"), iMail = col(h, "Correo"), iEst = col(h, "Estado");
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || !r[iNom]) continue;
    out.push({
      num: String(r[iNum] || "").trim(), n: String(r[iNom]).trim(),
      movil: String(r[iMov] || "").trim(), tel: String(r[iTel] || "").trim(),
      mail: String(r[iMail] || "").trim(), estado: String(r[iEst] || "").trim(), exps: []
    });
  }
  return out;
}
function parseExpedientes(rows) {
  const h = rows[0];
  const iAno = col(h, "Año"), iNum = col(h, "Núm. Exp"), iDesc = col(h, "Descripción Expediente"),
    iTpro = col(h, "Tipo Procedimiento"), iCli = col(h, "Cliente"), iCon = col(h, "Contrario"),
    iJuz = col(h, "Juzgado Principal"), iAut = col(h, "Núm. Autos"), iNig = col(h, "NIG"), iEst = col(h, "Estado");
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || !r[iNum]) continue;
    out.push({
      desc: String(r[iDesc] || "").trim() || "Expediente",
      tproc: String(r[iTpro] || "").trim(),
      cliente: String(r[iCli] || "").trim(),
      contrario: String(r[iCon] || "").trim(),
      juzgado: String(r[iJuz] || "").trim(),
      autos: String(r[iAut] || "").trim(),
      nig: String(r[iNig] || "").trim(),
      estado: String(r[iEst] || "").trim(),
      numexp: String(r[iNum]).trim() + "/" + String(r[iAno] || "").trim()
    });
  }
  return out;
}
async function cargarExportaciones() {
  const carpeta = encodeURIComponent(CFG.carpetaExportaciones || "Descargas");
  const items = await graphTodos("/me/drive/root:/" + carpeta + ":/children?$top=200");
  const excels = items.filter(i => i.file && /\.xlsx?$/i.test(i.name))
    .sort((a, b) => new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime))
    .slice(0, 15);
  let clientesRaw = null, expsRaw = null;
  for (const it of excels) {
    if (clientesRaw && expsRaw) break;
    try {
      const r = await fetch(it["@microsoft.graph.downloadUrl"]);
      const buf = await r.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
      const tipo = tipoDeExcel(rows);
      if (tipo === "clientes" && !clientesRaw) {
        clientesRaw = rows; FUENTE.fClientes = new Date(it.lastModifiedDateTime);
      } else if (tipo === "expedientes" && !expsRaw) {
        expsRaw = rows; FUENTE.fExpedientes = new Date(it.lastModifiedDateTime);
      }
    } catch (e) { /* fichero ilegible: probar el siguiente */ }
  }
  if (!clientesRaw && !expsRaw) throw new Error("No se han encontrado las exportaciones de MN Program en " + (CFG.carpetaExportaciones || "Descargas"));
  return { clientes: clientesRaw ? parseClientes(clientesRaw) : [], exps: expsRaw ? parseExpedientes(expsRaw) : [] };
}

/* ---------- calendario ---------- */
async function cargarCalendario() {
  const ini = new Date();
  const fin = new Date(Date.now() + (CFG.diasCalendario || 90) * 86400000);
  const url = "/me/calendarView?startDateTime=" + ini.toISOString() + "&endDateTime=" + fin.toISOString() +
    "&$top=100&$select=subject,bodyPreview,start,end&$orderby=start/dateTime";
  const eventos = await graphTodos(url);
  return eventos.map(ev => clasificarCita(
    ev.subject || "", ev.bodyPreview || "",
    new Date(ev.start.dateTime + (ev.start.timeZone === "UTC" ? "Z" : "")),
    new Date(ev.end.dateTime + (ev.end.timeZone === "UTC" ? "Z" : ""))
  ));
}

/* ---------- documentos de OneDrive ---------- */
const BASE_MN = () => (CFG.carpetaMN || "DOCS-MNPROGRAM_81002");
let mapaCarpetas = null;
// número de autos -> token que MN/LexNET usa en los nombres de fichero
// "376/2025" -> "2025_0000376" ; "0000363/2024" -> "2024_0000363"
function lexnetToken(autos) {
  const m = /(\d+)\s*\/\s*(\d{4})/.exec(autos || "");
  if (!m) return null;
  return m[2] + "_" + m[1].replace(/^0+/, "").padStart(7, "0");
}
function descCoincide(nombreSub, desc) {
  const resto = nombreSub.replace(/^Exp\d*/, "").trim().toLowerCase();
  const d = (desc || "").toLowerCase();
  return !!resto && !!d && (resto.indexOf(d) >= 0 || d.indexOf(resto) >= 0);
}
// nº de autos distintos (formato LexNET AAAA_NNNNNNN) presentes en una lista de ficheros
function autosDistintos(files) {
  const set = new Set();
  files.forEach(f => { const m = /(\d{4})_(\d{7})_/.exec(f.n || ""); if (m) set.add(m[1] + m[2]); });
  return set.size;
}
// Algunas carpetas de MN son "cajones" con documentos de MUCHOS trabajadores
// (p. ej. ExpCantidad de MARKTEL, 161 ficheros). Si la carpeta mezcla varios
// autos, devolvemos solo los documentos de ESTE expediente (por nº de autos o
// por el nombre del contrario); si es una carpeta dedicada, devolvemos todo.
function filtrarDocsExpediente(files, e) {
  if (autosDistintos(files) < 3) return files;
  const tok = lexnetToken(e.autos);
  const apellidos = norm(e.contrario || "").split(/\s+/).filter(w => w.length >= 4);
  return files.filter(f => {
    if (tok && f.n.indexOf(tok) >= 0) return true;
    if (apellidos.length >= 2) {
      const n = norm(f.n);
      if (apellidos.filter(w => n.indexOf(w) >= 0).length >= 2) return true;
    }
    return false;
  });
}
async function asegurarCarpetas() {
  if (mapaCarpetas) return;
  const cache = localStorage.getItem("nacar_carpetas2");
  if (cache) {
    const j = JSON.parse(cache);
    if (Date.now() - j.t < 86400000) { mapaCarpetas = j.m; return; }
  }
  const ruta = encodeURIComponent(BASE_MN());
  const items = await graphTodos("/me/drive/root:/" + ruta + "/Usu2:/children?$top=999&$select=name,folder");
  mapaCarpetas = {};
  // un mismo cliente puede tener VARIAS carpetas con el mismo nombre normalizado
  items.forEach(i => { if (i.folder) { const k = norm(i.name); (mapaCarpetas[k] = mapaCarpetas[k] || []).push(i.name); } });
  localStorage.setItem("nacar_carpetas2", JSON.stringify({ t: Date.now(), m: mapaCarpetas }));
}
async function cargarDocsCliente(c) {
  if (c._docsCargados || MODO_DEMO) return;
  await asegurarCarpetas();
  c._docsCargados = true;
  const carpetas = mapaCarpetas[norm(c.n)] || [];
  if (!carpetas.length) { c._sinCarpeta = true; return; }
  c._rutas = [];
  c._rootDocs = [];
  const subs = [];
  for (const carpeta of carpetas) {
    const ruta = encodeURIComponent(BASE_MN() + "/Usu2/" + carpeta);
    c._rutas.push(ruta);
    let items = [];
    try { items = await graphTodos("/me/drive/root:/" + ruta + ":/children?$top=500&$select=name,file,folder,webUrl,createdDateTime,lastModifiedDateTime"); }
    catch (e) { continue; }
    items.forEach(i => {
      if (i.file) c._rootDocs.push({ n: i.name, url: i.webUrl, f: i.createdDateTime || i.lastModifiedDateTime });
      else if (i.folder) subs.push({ name: i.name, ruta, usada: false });
    });
  }
  // Emparejar por NOMBRE solo cuando es fiable: 1) código exacto Exp{año}{num};
  // 2) descripción con un único candidato (evita asignar mal en empresas con
  // decenas de carpetas "Despido"). El resto se resuelve por nº de autos al abrir.
  c.exps.forEach(e => {
    if (e.desc === "Documentación general") return;
    let elegida = null;
    const me = /^(\d+)\/(\d{4})$/.exec(e.numexp || "");
    if (me) elegida = subs.find(s => !s.usada && new RegExp("^Exp" + me[2] + me[1] + "(\\D|$)").test(s.name));
    if (!elegida) {
      const cands = subs.filter(s => !s.usada && descCoincide(s.name, e.desc));
      if (cands.length === 1) elegida = cands[0];
    }
    if (elegida) { elegida.usada = true; e._subcarpeta = elegida.name; e._subRuta = elegida.ruta; }
  });
  c._subSueltas = subs.filter(s => !s.usada).map(s => s.name);
}
async function cargarDocsExpediente(c, e) {
  if (e._docs || MODO_DEMO) return;
  // 1) subcarpeta ya emparejada por nombre
  if (e._subcarpeta && e._subRuta) {
    try {
      const items = await graphTodos("/me/drive/root:/" + e._subRuta + "/" + encodeURIComponent(e._subcarpeta) + ":/children?$top=500&$select=name,file,webUrl,createdDateTime,lastModifiedDateTime");
      const todos = items.filter(i => i.file).map(i => ({ n: i.name, url: i.webUrl, f: i.createdDateTime || i.lastModifiedDateTime }));
      e._docs = filtrarDocsExpediente(todos, e);
      if (e._docs.length) return;
    } catch (err) { /* sigue al fallback por autos */ }
  }
  // 2) fallback: localizar la carpeta por el nº de autos dentro de los ficheros
  //    de LexNET (p. ej. autos 376/2025 -> ficheros "2025_0000376_...")
  const tok = lexnetToken(e.autos);
  if (tok && c._rutas && c._rutas.length) {
    for (const ruta of c._rutas) {
      let hits = [];
      try { hits = await graphTodos("/me/drive/root:/" + ruta + ":/search(q='" + tok + "')?$top=25&$select=name,file,parentReference,webUrl"); }
      catch (err) { continue; }
      const hit = hits.find(h => h.file && h.name.indexOf(tok) >= 0 && h.parentReference && h.parentReference.id);
      if (hit) {
        const items = await graphTodos("/me/drive/items/" + hit.parentReference.id + "/children?$top=500&$select=name,file,webUrl,createdDateTime,lastModifiedDateTime");
        const todos = items.filter(i => i.file).map(i => ({ n: i.name, url: i.webUrl, f: i.createdDateTime || i.lastModifiedDateTime }));
        e._docs = filtrarDocsExpediente(todos, e);
        if (hit.parentReference.name) e._subcarpeta = hit.parentReference.name;
        if (e._docs.length) return;
      }
    }
  }
  e._docs = e._docs || [];
}

/* ---------- indexado tras la carga ---------- */
function indexar(listaClientes, listaExps, listaCitas) {
  CLIENTES = listaClientes;
  const porNombre = {};
  CLIENTES.forEach((c, i) => { if (!(norm(c.n) in porNombre)) porNombre[norm(c.n)] = i; });
  listaExps.forEach(x => {
    let ci = porNombre[norm(x.cliente)];
    if (ci === undefined) {
      CLIENTES.push({ num: "", n: x.cliente || "(sin cliente)", movil: "", tel: "", mail: "", estado: "", exps: [] });
      ci = CLIENTES.length - 1;
      porNombre[norm(x.cliente)] = ci;
    }
    CLIENTES[ci].exps.push(x);
  });
  EXPS = []; CONTRA = []; conIdx = {};
  CLIENTES.forEach((c, ci) => {
    c.exps.forEach(e => {
      if (e.desc === "Documentación general") return;
      const i = EXPS.length;
      EXPS.push({ ci, cliente: c.n, e, demo: !!e.d });
      if (e.contrario) {
        const k = norm(e.contrario);
        if (!(k in conIdx)) { conIdx[k] = CONTRA.length; CONTRA.push({ nombre: e.contrario, exps: [] }); }
        CONTRA[conIdx[k]].exps.push(i);
      }
    });
  });
  CITAS = listaCitas.sort((a, b) => a.inicio - b.inicio);
}

/* ============================================================
   ARRANQUE
   ============================================================ */
async function init() {
  document.querySelectorAll("nav .tab").forEach(t => t.addEventListener("click", () => irTab(t.dataset.tab)));
  document.getElementById("btn-atras").addEventListener("click", atras);
  document.getElementById("btn-ajustes").addEventListener("click", () => irA(pintarAjustes));
  if (MODO_DEMO) {
    FUENTE.usuario = "Sotero";
    FUENTE.fClientes = new Date("2026-06-11");
    FUENTE.fExpedientes = new Date("2026-06-12");
    const demoCitas = window.NACAR_DEMO.citas.map(c =>
      clasificarCita(c.asunto, c.cuerpo, new Date(c.inicioISO), new Date(c.finISO)));
    const demoClientes = window.NACAR_DEMO.clientes.map(c => ({
      num: c.num, n: c.n, movil: c.movil, tel: c.tel, mail: c.mail, estado: "Alta",
      exps: c.exps.map(e => Object.assign({}, e))
    }));
    const sueltos = [];
    demoClientes.forEach(c => { c.exps = c.exps.filter(e => { if (e.desc !== "Documentación general") return true; c._rootDocsDemo = e.d; return false; }); });
    indexar(demoClientes, [], demoCitas);
    // en demo los expedientes ya vienen dentro de cada cliente
    EXPS = []; CONTRA = []; conIdx = {};
    CLIENTES.forEach((c, ci) => c.exps.forEach(e => {
      const i = EXPS.length;
      EXPS.push({ ci, cliente: c.n, e });
      if (e.contrario) {
        const k = norm(e.contrario);
        if (!(k in conIdx)) { conIdx[k] = CONTRA.length; CONTRA.push({ nombre: e.contrario, exps: [] }); }
        CONTRA[conIdx[k]].exps.push(i);
      }
    }));
    ponerEstado("Modo demo");
    irTab("hoy");
    return;
  }
  // modo conectado
  msalApp = new msal.PublicClientApplication({
    auth: {
      clientId: CFG.clientId,
      authority: "https://login.microsoftonline.com/" + (CFG.tenantId || "common"),
      redirectUri: location.origin + location.pathname.replace(/index\.html$/, "")
    },
    cache: { cacheLocation: "localStorage" }
  });
  await msalApp.initialize();
  const resp = await msalApp.handleRedirectPromise();
  cuenta = (resp && resp.account) || msalApp.getAllAccounts()[0] || null;
  if (!cuenta) { pintarLogin(); return; }
  await cargarTodo();
}
async function cargarTodo() {
  ponerEstado("Cargando datos...");
  pintar('<div class="cargando">Cargando clientes, expedientes y agenda<br>desde tu Microsoft 365...</div>');
  try {
    const yo = await graph("/me");
    FUENTE.usuario = (yo.givenName || yo.displayName || "").split(" ")[0];
    const [exp, citas] = await Promise.all([cargarExportaciones(), cargarCalendario()]);
    indexar(exp.clientes, exp.exps, citas);
    ponerEstado("Conectado");
    irTab("hoy");
  } catch (e) {
    pintar('<div class="vacio"><i class="ti ti-plug-x" style="font-size:34px;color:var(--coral);"></i>' +
      '<p>No se han podido cargar los datos.</p><p class="mini">' + esc(e.message) + '</p>' +
      '<button class="boton-secundario" onclick="cargarTodo()">Reintentar</button></div>');
  }
}
function pintarLogin() {
  ponerEstado("Sin sesión");
  pintar('<div class="pantalla-login">' +
    '<img src="logo-completo.png?v=3" alt="Nácar Abogados" style="width: 150px; height: auto;" />' +
    '<button class="boton-principal" id="btn-login"><i class="ti ti-brand-windows" style="vertical-align:-2px;"></i> Entrar con la cuenta del despacho</button>' +
    '<p class="nota">Acceso de solo lectura.<br>Los datos nunca salen de tu Microsoft 365.</p></div>');
  document.getElementById("btn-login").addEventListener("click", () =>
    msalApp.loginRedirect({ scopes: SCOPES }));
}

/* ============================================================
   NAVEGACIÓN Y PANTALLAS
   ============================================================ */
function pintar(html) { document.getElementById("contenido").innerHTML = html; document.getElementById("contenido").scrollTop = 0; }
function ponerEstado(t) { document.getElementById("hdr-estado").textContent = t; }
function actualizarTabs(tab) {
  tabActual = tab;
  document.querySelectorAll("nav .tab").forEach(t => t.classList.toggle("activa", t.dataset.tab === tab));
}
function actualizarAtras() {
  document.getElementById("btn-atras").style.display = historial.length > 1 ? "flex" : "none";
}
function irTab(tab) {
  const fn = tab === "hoy" ? pintarHoy : tab === "buscar" ? pintarBuscar : tab === "calc" ? pintarCalc : pintarAgenda;
  historial = [fn];
  actualizarTabs(tab);
  fn();
  actualizarAtras();
}
function irA(fn) { historial.push(fn); fn(); actualizarAtras(); }
function atras() {
  if (historial.length > 1) { historial.pop(); historial[historial.length - 1](); actualizarAtras(); }
}

/* ---------- componentes ---------- */
function htmlExp(i) {
  const x = EXPS[i], e = x.e;
  return '<div class="tarjeta-exp" onclick="abrirExpedienteNav(' + i + ')">' +
    '<div class="linea1"><i class="ti ti-briefcase"></i><p>' + esc(e.desc) + (e.contrario ? " · " + esc(e.contrario) : "") + '</p><i class="ti ti-chevron-right" style="color:var(--teal);font-size:15px;"></i></div>' +
    '<div class="badges">' +
    (e.autos ? '<span class="badge lleno">Autos ' + esc(e.autos) + '</span>' : '<span class="mini">Sin autos (fase previa)</span>') +
    badgeEstado(e.estado) +
    '<span class="mini">' + esc(x.cliente) + '</span></div>' +
    (e.juzgado ? '<p class="juz"><i class="ti ti-building-bank" style="font-size:12px;"></i> ' + esc(e.juzgado) + '</p>' : "") +
    '</div>';
}
function htmlCita(c, idx) {
  const es = estiloTipo(c.tipo);
  const expI = expedienteDeCita(c);
  const finde = c.tipo === "plazo" && (c.inicio.getDay() === 0 || c.inicio.getDay() === 6);
  return '<div class="tarjeta-cita" onclick="abrirCitaNav(' + idx + ')">' +
    '<div class="icono-tipo" style="background:' + es.bg + ';"><i class="ti ' + es.icono + '" style="color:' + es.color + ';"></i></div>' +
    '<div class="cuerpo"><div class="linea1"><p>' + esc(c.titulo) + '</p>' +
    '<span class="hora" style="background:' + es.bg + ';color:' + es.color + ';">' + fmtHora(c.inicio) + '</span></div>' +
    '<p class="sub">' + [c.cliente ? esc(c.cliente) : "", expI >= 0 && EXPS[expI].e.autos ? "Autos " + esc(EXPS[expI].e.autos) : ""].filter(Boolean).join(" · ") + '</p>' +
    (expI >= 0 && EXPS[expI].e.juzgado ? '<p class="sub"><i class="ti ti-building-bank" style="font-size:12px;"></i> ' + esc(EXPS[expI].e.juzgado) + '</p>' : "") +
    (finde ? '<p class="aviso-linea"><i class="ti ti-alert-triangle" style="font-size:12px;"></i> Cae en fin de semana</p>' : "") +
    '</div></div>';
}
function abrirCitaNav(idx) {
  const c = CITAS[idx];
  const i = expedienteDeCita(c);
  if (i >= 0) irA(() => abrirExpediente(i));
  else if (c.cliente) {
    const ci = CLIENTES.findIndex(cl => norm(cl.n) === norm(c.cliente));
    if (ci >= 0) { irA(() => abrirCliente(ci)); return; }
    alert("Esta cita no tiene expediente vinculado en las exportaciones actuales.");
  } else alert("Cita del calendario sin datos de MN Program (reunión u otra cita personal).");
}
function abrirExpedienteNav(i) { irA(() => abrirExpediente(i)); }
function abrirClienteNav(i) { irA(() => abrirCliente(i)); }
function abrirContrarioNav(i) { irA(() => abrirContrario(i)); }

/* ---------- pantalla HOY ---------- */
function pintarHoy() {
  actualizarTabs("hoy");
  ponerEstado(MODO_DEMO ? "Modo demo" : "Conectado");
  const ahora = new Date();
  const hoyJ = CITAS.filter(c => c.tipo === "juicio" && mismoDia(c.inicio, ahora));
  const hoyP = CITAS.filter(c => c.tipo !== "juicio" && c.esMN && mismoDia(c.inicio, ahora));
  const finSemana = new Date(ahora); finSemana.setDate(finSemana.getDate() + (7 - finSemana.getDay()) % 7 + 1);
  const semanaJ = CITAS.filter(c => c.tipo === "juicio" && c.inicio >= ahora && c.inicio < finSemana);
  const totalJ = CITAS.filter(c => c.tipo === "juicio" && c.inicio >= ahora);
  let html = '<div class="saludo"><h1>' + saludoHora() + (FUENTE.usuario ? ", " + esc(FUENTE.usuario) : "") + '</h1>' +
    '<p>' + fmtDia(ahora) + " de " + ahora.getFullYear() + '</p></div>' +
    '<div class="metricas">' +
    '<div class="metrica"><p class="num">' + hoyJ.length + '</p><p class="lbl">juicios hoy</p></div>' +
    '<div class="metrica' + (hoyP.length ? " ambar" : "") + '"><p class="num">' + hoyP.length + '</p><p class="lbl">plazos hoy</p></div>' +
    '<div class="metrica"><p class="num">' + semanaJ.length + '</p><p class="lbl">juicios esta semana</p></div>' +
    '<div class="metrica"><p class="num">' + totalJ.length + '</p><p class="lbl">en ' + (CFG.diasCalendario || 90) + ' días</p></div></div>';
  const proximas = CITAS.filter(c => c.inicio >= ahora && c.esMN).slice(0, 8);
  if (proximas.length) {
    html += '<p class="seccion">PRÓXIMAS CITAS</p>';
    let diaAct = "";
    proximas.forEach(c => {
      const d = fmtDia(c.inicio);
      if (d !== diaAct) { diaAct = d; html += '<p class="dia-agenda">' + d + '</p>'; }
      html += htmlCita(c, CITAS.indexOf(c));
    });
  } else {
    html += '<div class="vacio">No hay citas próximas en el calendario</div>';
  }
  html += '<button class="boton-linea" onclick="irTab(\'agenda\')"><p>Ver toda la agenda <i class="ti ti-arrow-right" style="font-size:12px;"></i></p></button>';
  pintar(html);
}
function saludoHora() {
  const h = new Date().getHours();
  return h < 14 ? "Buenos días" : h < 21 ? "Buenas tardes" : "Buenas noches";
}

/* ---------- pantalla BUSCAR ---------- */
let ultimaBusqueda = "";
function pintarBuscar() {
  actualizarTabs("buscar");
  ponerEstado("Búsqueda global");
  pintar('<div class="busqueda"><i class="ti ti-search"></i>' +
    '<input type="search" id="buscador" placeholder="Cliente, autos, NIG, contrario, juzgado..." value="' + esc(ultimaBusqueda) + '" autocomplete="off"></div>' +
    '<div id="resultados"></div>');
  const inp = document.getElementById("buscador");
  inp.addEventListener("input", () => { ultimaBusqueda = inp.value; pintarResultados(inp.value); });
  pintarResultados(ultimaBusqueda);
}
function pintarResultados(f) {
  f = (f || "").toLowerCase().trim();
  const cont = document.getElementById("resultados");
  if (!f) {
    cont.innerHTML = '<p class="contador" style="margin-top:14px;">Un solo buscador para todo: clientes, expedientes, números de autos, NIG, juzgados, contrarios y documentos.</p>' +
      '<div class="atajos">' +
      '<button class="atajo" onclick="pintarResultadosAtajo(\'clientes\')"><i class="ti ti-users"></i><p class="t">Clientes</p><p class="s">' + CLIENTES.length + '</p></button>' +
      '<button class="atajo" onclick="pintarResultadosAtajo(\'expedientes\')"><i class="ti ti-briefcase"></i><p class="t">Expedientes</p><p class="s">' + EXPS.length + '</p></button>' +
      '<button class="atajo" onclick="pintarResultadosAtajo(\'contrarios\')"><i class="ti ti-user"></i><p class="t">Contrarios</p><p class="s">' + CONTRA.length + '</p></button></div>';
    return;
  }
  let html = "";
  const cli = [];
  CLIENTES.forEach((c, i) => { if ((c.n + " " + c.num + " " + c.movil + " " + c.mail).toLowerCase().indexOf(f) >= 0) cli.push(i); });
  if (cli.length) {
    html += '<p class="seccion">CLIENTES (' + cli.length + ')</p>';
    cli.slice(0, 6).forEach(i => {
      const c = CLIENTES[i];
      html += '<div class="fila" onclick="abrirClienteNav(' + i + ')"><div class="avatar">' + esc(iniciales(c.n)) + '</div>' +
        '<div class="cuerpo"><p class="titulo">' + esc(c.n) + '</p><p class="detalle">' + (c.num ? "Nº " + esc(c.num) + " · " : "") + c.exps.length + ' expedientes</p></div>' +
        '<i class="ti ti-chevron-right flecha"></i></div>';
    });
    if (cli.length > 6) html += '<p class="contador">+ ' + (cli.length - 6) + ' clientes más — afina la búsqueda</p>';
  }
  const cons = [];
  CONTRA.forEach((c, i) => { if (c.nombre.toLowerCase().indexOf(f) >= 0) cons.push(i); });
  if (cons.length) {
    html += '<p class="seccion">CONTRARIOS (' + cons.length + ')</p>';
    cons.slice(0, 6).forEach(i => {
      const c = CONTRA[i];
      html += '<div class="fila" onclick="abrirContrarioNav(' + i + ')"><div class="avatar contrario"><i class="ti ti-user"></i></div>' +
        '<div class="cuerpo"><p class="titulo">' + esc(c.nombre) + '</p><p class="detalle">' + c.exps.length + ' procedimiento' + (c.exps.length > 1 ? "s" : "") + '</p></div>' +
        '<i class="ti ti-chevron-right flecha" style="color:var(--coral);"></i></div>';
    });
  }
  const exq = [];
  EXPS.forEach((x, i) => {
    const e = x.e;
    if ((e.desc + " " + e.autos + " " + e.nig + " " + e.juzgado + " " + e.tproc + " " + e.numexp + " " + x.cliente + " " + e.contrario).toLowerCase().indexOf(f) >= 0) exq.push(i);
  });
  if (exq.length) {
    html += '<p class="seccion">EXPEDIENTES (' + exq.length + ')</p>';
    exq.slice(0, 8).forEach(i => html += htmlExp(i));
    if (exq.length > 8) html += '<p class="contador">+ ' + (exq.length - 8) + ' expedientes más — afina la búsqueda</p>';
  }
  const docs = [];
  EXPS.forEach((x, i) => {
    const lista = x.e._docs || x.e.d || [];
    lista.forEach(d => {
      const nombre = d.n || d;
      if (nombre.toLowerCase().indexOf(f) >= 0) docs.push({ nombre, i, f: d.f });
    });
  });
  if (docs.length) {
    html += '<p class="seccion">DOCUMENTOS (EN EXPEDIENTES YA ABIERTOS)</p>';
    docs.slice(0, 8).forEach(o => {
      html += '<div class="fila-doc" style="cursor:pointer;" onclick="abrirExpedienteNav(' + o.i + ')"><i class="ti ' + iconoDoc(o.nombre) + '"></i>' +
        '<div class="nombre-doc"><p>' + esc(o.nombre) + '</p><p class="origen">' + (o.f ? fmtFecha(o.f) + " · " : "") + esc(EXPS[o.i].e.desc) + " · " + esc(EXPS[o.i].cliente) + '</p></div></div>';
    });
  }
  // búsqueda de documentos en todo OneDrive (por nombre y por texto interno)
  if (!MODO_DEMO && f.length >= 2) html += '<div id="docs-od"><p class="seccion">DOCUMENTOS EN ONEDRIVE</p><p class="contador">Buscando…</p></div>';
  cont.innerHTML = html || '<div class="vacio">Nada coincide con esa búsqueda</div>';
  if (!MODO_DEMO && f.length >= 2) buscarDocsOneDriveDebounced(f);
}
let _odTimer = null, _odSeq = 0;
function buscarDocsOneDriveDebounced(f) {
  clearTimeout(_odTimer);
  _odTimer = setTimeout(() => buscarDocsOneDrive(f), 400);
}
async function buscarDocsOneDrive(f) {
  const seq = ++_odSeq;
  const ruta = encodeURIComponent(BASE_MN());
  let hits = [];
  try {
    hits = await graphTodos("/me/drive/root:/" + ruta + ":/search(q='" + f.replace(/'/g, "''") + "')?$top=40&$select=name,webUrl,file,parentReference,createdDateTime,lastModifiedDateTime");
  } catch (e) { /* sin resultados */ }
  if (seq !== _odSeq) return;            // ya hay una búsqueda más reciente
  const cont = document.getElementById("docs-od");
  if (!cont) return;
  const docs = hits.filter(h => h.file);
  if (!docs.length) {
    cont.innerHTML = '<p class="seccion">DOCUMENTOS EN ONEDRIVE</p><p class="contador">Sin documentos que coincidan con «' + esc(f) + '»</p>';
    return;
  }
  let html = '<p class="seccion">DOCUMENTOS EN ONEDRIVE (' + docs.length + (docs.length >= 40 ? "+" : "") + ')</p>';
  docs.forEach(d => {
    const carpeta = (d.parentReference && d.parentReference.name) || "";
    const fec = fmtFecha(d.createdDateTime || d.lastModifiedDateTime);
    html += '<div class="fila-doc" style="cursor:pointer;" onclick="abrirDocUrl(\'' + encodeURIComponent(d.webUrl || "") + '\')"><i class="ti ' + iconoDoc(d.name) + '"></i>' +
      '<div class="nombre-doc"><p>' + esc(d.name) + '</p><p class="origen">' + (fec ? fec + " · " : "") + esc(carpeta) + '</p></div>' +
      '<i class="ti ti-external-link" style="color:var(--texto-3);font-size:15px;flex-shrink:0;"></i></div>';
  });
  html += '<p class="contador">Busca en el nombre y en el texto interno de los documentos. Los PDF escaneados sin OCR solo se encuentran por el nombre.</p>';
  cont.innerHTML = html;
}
function abrirDocUrl(u) { const url = decodeURIComponent(u || ""); if (url) window.open(url, "_blank"); }
function pintarResultadosAtajo(t) {
  const inp = document.getElementById("buscador");
  const cont = document.getElementById("resultados");
  let html = "";
  if (t === "clientes") {
    html = '<p class="seccion">TODOS LOS CLIENTES (' + CLIENTES.length + ')</p>';
    CLIENTES.map((c, i) => ({ c, i })).sort((a, b) => a.c.n.localeCompare(b.c.n)).slice(0, 100).forEach(o => {
      html += '<div class="fila" onclick="abrirClienteNav(' + o.i + ')"><div class="avatar">' + esc(iniciales(o.c.n)) + '</div>' +
        '<div class="cuerpo"><p class="titulo">' + esc(o.c.n) + '</p><p class="detalle">' + o.c.exps.length + ' expedientes</p></div><i class="ti ti-chevron-right flecha"></i></div>';
    });
    if (CLIENTES.length > 100) html += '<p class="contador">Mostrando 100 — usa el buscador para el resto</p>';
  } else if (t === "expedientes") {
    html = '<p class="seccion">EXPEDIENTES ABIERTOS</p>';
    EXPS.forEach((x, i) => { if (x.e.estado !== "Cerrado" && html.length < 120000) html += htmlExp(i); });
  } else {
    html = '<p class="seccion">CONTRARIOS CON VARIOS PLEITOS</p>';
    CONTRA.map((c, i) => ({ c, i })).filter(o => o.c.exps.length > 1).sort((a, b) => b.c.exps.length - a.c.exps.length).slice(0, 60).forEach(o => {
      html += '<div class="fila" onclick="abrirContrarioNav(' + o.i + ')"><div class="avatar contrario"><i class="ti ti-user"></i></div>' +
        '<div class="cuerpo"><p class="titulo">' + esc(o.c.nombre) + '</p><p class="detalle">' + o.c.exps.length + ' procedimientos</p></div><i class="ti ti-chevron-right flecha" style="color:var(--coral);"></i></div>';
    });
  }
  cont.innerHTML = html;
}

/* ---------- pantalla AGENDA ---------- */
let filtroAgenda = "todos";
function pintarAgenda() {
  actualizarTabs("agenda");
  ponerEstado("Agenda · Outlook");
  const ahora = new Date();
  const futuras = CITAS.filter(c => c.fin >= ahora);
  const nJ = futuras.filter(c => c.tipo === "juicio").length;
  const nP = futuras.filter(c => c.tipo === "plazo" || c.tipo === "vencimiento").length;
  const nO = futuras.length - nJ - nP;
  let html = '<div class="caja-info"><i class="ti ti-refresh"></i><p>Sincronizado con tu calendario de Outlook · se actualiza solo</p></div>' +
    '<div class="chips">' +
    chip("todos", "Todas (" + futuras.length + ")") + chip("juicio", "Juicios (" + nJ + ")") +
    chip("plazo", "Plazos (" + nP + ")") + chip("otra", "Otras (" + nO + ")") + '</div><div id="lista-agenda"></div>';
  pintar(html);
  pintarListaAgenda();
}
function chip(id, label) {
  return '<button class="chip' + (filtroAgenda === id ? " activo" : "") + '" onclick="cambiarFiltro(\'' + id + '\')">' + label + '</button>';
}
function cambiarFiltro(f) { filtroAgenda = f; pintarAgenda(); }
function pintarListaAgenda() {
  const ahora = new Date();
  let html = "", diaAct = "";
  CITAS.forEach((c, idx) => {
    if (c.fin < ahora) return;
    const grupo = c.tipo === "juicio" ? "juicio" : (c.tipo === "plazo" || c.tipo === "vencimiento") ? "plazo" : c.tipo === "confesion" ? "juicio" : "otra";
    if (filtroAgenda !== "todos" && grupo !== filtroAgenda) return;
    const d = fmtDia(c.inicio);
    if (d !== diaAct) { diaAct = d; html += '<p class="dia-agenda">' + d + '</p>'; }
    html += htmlCita(c, idx);
  });
  document.getElementById("lista-agenda").innerHTML = html || '<div class="vacio">No hay citas de este tipo</div>';
}

/* ---------- ficha CLIENTE ---------- */
function abrirCliente(i) {
  const c = CLIENTES[i];
  ponerEstado("Ficha de cliente");
  const tel = c.movil || c.tel;
  let html = '<div class="cabecera-ficha"><div class="fila-id">' +
    '<div class="avatar-g">' + esc(iniciales(c.n)) + '</div>' +
    '<div><h2>' + esc(c.n) + '</h2><p class="sub">' + (c.num ? "Cliente nº " + esc(c.num) : "Cliente") + (c.estado ? " · " + esc(c.estado) : "") + '</p></div></div>';
  if (tel || c.mail) {
    html += '<div class="acciones">' +
      (tel ? '<a class="lleno" href="tel:' + esc(tel.replace(/\s/g, "")) + '"><i class="ti ti-phone" style="vertical-align:-2px;"></i> Llamar</a>' : "") +
      (c.mail ? '<a class="hueco" href="mailto:' + esc(c.mail) + '"><i class="ti ti-mail" style="vertical-align:-2px;"></i> Email</a>' : "") + '</div>' +
      '<p class="sub" style="margin-top:10px;">' + (tel ? '<i class="ti ti-device-mobile" style="font-size:13px;"></i> ' + esc(tel) + '  ' : "") +
      (c.mail ? '<i class="ti ti-at" style="font-size:13px;"></i> ' + esc(c.mail) : "") + '</p>';
  } else {
    html += '<p class="sub" style="margin-top:10px;"><i class="ti ti-alert-triangle" style="font-size:13px;color:var(--ambar);"></i> Sin datos de contacto en MN Program</p>';
  }
  html += '</div><p class="seccion">EXPEDIENTES (' + c.exps.filter(e => e.desc !== "Documentación general").length + ')</p><div id="exps-cliente"></div><div id="docs-cliente"></div>';
  pintar(html);
  let lista = "";
  EXPS.forEach((x, xi) => { if (x.ci === i) lista += htmlExp(xi); });
  document.getElementById("exps-cliente").innerHTML = lista || '<div class="vacio">Sin expedientes en la exportación actual</div>';
  pintarDocsGenerales(c);
  if (!MODO_DEMO && !c._docsCargados) {
    cargarDocsCliente(c).then(() => pintarDocsGenerales(c)).catch(() => {});
  }
}
function pintarDocsGenerales(c) {
  const cont = document.getElementById("docs-cliente");
  if (!cont) return;
  const docs = MODO_DEMO ? (c._rootDocsDemo || []).map(n => ({ n, url: null })) : (c._rootDocs || []);
  if (!docs.length) { cont.innerHTML = c._sinCarpeta ? '<p class="contador">Este cliente no tiene carpeta de documentos localizada en OneDrive</p>' : ""; return; }
  let html = '<p class="seccion">DOCUMENTACIÓN GENERAL</p>';
  docs.forEach(d => { html += htmlDoc(d, c); });
  cont.innerHTML = html;
}
function htmlDoc(d, cliente) {
  const id = "d" + Math.random().toString(36).slice(2, 9);
  window["__" + id] = { d, cliente };
  return '<div class="fila-doc"><i class="ti ' + iconoDoc(d.n) + '"></i>' +
    '<div class="nombre-doc" onclick="abrirDoc(\'' + id + '\')"><p>' + esc(d.n) + '</p>' +
    (d.f ? '<p class="origen"><i class="ti ti-calendar-event" style="font-size:11px;vertical-align:-1px;"></i> ' + fmtFecha(d.f) + '</p>' : '') + '</div>' +
    '<button class="btn-doc" aria-label="Compartir" onclick="compartirDoc(\'' + id + '\')"><i class="ti ti-share-2"></i></button></div>';
}
function abrirDoc(id) {
  const o = window["__" + id];
  if (o && o.d.url) window.open(o.d.url, "_blank");
  else alert(MODO_DEMO ? "Modo demo: en la app conectada este documento se abriría desde OneDrive." : "Este documento no tiene enlace disponible.");
}
function compartirDoc(id) {
  const o = window["__" + id];
  if (!o) return;
  const url = o.d.url || "";
  const mail = (o.cliente && o.cliente.mail) || "";
  if (navigator.share && url) {
    navigator.share({ title: o.d.n, url }).catch(() => {});
  } else if (url || mail) {
    location.href = "mailto:" + encodeURIComponent(mail) + "?subject=" + encodeURIComponent("Documentación - NACAR ABOGADOS") +
      "&body=" + encodeURIComponent("Estimado/a cliente:\n\nLe adjunto el enlace al documento " + o.d.n + ":\n" + url + "\n\nUn saludo,\nNACAR ABOGADOS");
  } else {
    alert("Modo demo: en la app conectada se abriría la hoja de compartir de iPhone (email al cliente, WhatsApp, guardar...).");
  }
}

/* ---------- ficha EXPEDIENTE ---------- */
function abrirExpediente(i) {
  const x = EXPS[i], e = x.e, c = CLIENTES[x.ci];
  ponerEstado("Ficha de expediente");
  const ki = e.contrario ? conIdx[norm(e.contrario)] : undefined;
  const citasExp = CITAS.map((ct, idx) => ({ ct, idx })).filter(o => o.ct.numexp && o.ct.numexp === e.numexp && o.ct.fin >= new Date());
  let html = '<div class="cabecera-ficha"><div class="fila-id">' +
    '<div class="avatar-g cuadrado"><i class="ti ti-briefcase"></i></div>' +
    '<div><h2>' + esc(e.desc) + '</h2><p class="sub">' + esc(e.tproc || "Sin tipo de procedimiento") + '</p></div></div>' +
    '<div class="badges" style="margin:12px 0 0;">' +
    (e.autos ? '<span class="badge lleno">Autos ' + esc(e.autos) + '</span>' : '<span class="mini">Sin nº de autos (fase previa / SMAC)</span>') +
    badgeEstado(e.estado) + '</div>' +
    '<table class="tabla-datos">' +
    (e.juzgado ? '<tr><td>Juzgado</td><td>' + esc(e.juzgado) + '</td></tr>' : "") +
    (e.contrario ? '<tr><td>Contrario</td><td><span class="enlace coral" onclick="abrirContrarioNav(' + (ki !== undefined ? ki : -1) + ')">' + esc(e.contrario) + ' <i class="ti ti-chevron-right" style="font-size:11px;"></i></span></td></tr>' : "") +
    (e.nig ? '<tr><td>NIG</td><td style="font-family:ui-monospace,monospace;font-size:11px;">' + esc(e.nig) + '</td></tr>' : "") +
    (e.numexp ? '<tr><td>Ref. MN</td><td>' + esc(e.numexp) + '</td></tr>' : "") +
    '<tr><td>Cliente</td><td><span class="enlace" onclick="abrirClienteNav(' + x.ci + ')">' + esc(x.cliente) + ' <i class="ti ti-chevron-right" style="font-size:11px;"></i></span></td></tr>' +
    '</table></div>';
  if (citasExp.length) {
    html += '<p class="seccion">PRÓXIMAS CITAS DE ESTE EXPEDIENTE</p>';
    citasExp.slice(0, 3).forEach(o => {
      html += '<div class="caja-info" style="margin-top:4px;"><i class="ti ' + estiloTipo(o.ct.tipo).icono + '"></i><p>' +
        etiquetaTipo(o.ct.tipo) + " · " + fmtDia(o.ct.inicio) + " a las " + fmtHora(o.ct.inicio) + '</p></div>';
    });
  }
  html += '<div class="busqueda"><i class="ti ti-search"></i><input type="search" id="buscador-docs" placeholder="Buscar en la documentación..." autocomplete="off"></div>' +
    '<div id="docs-exp"><div class="cargando">Cargando documentos...</div></div>';
  pintar(html);
  const render = () => {
    const docs = MODO_DEMO ? (e.d || []).map(n => ({ n, url: null })) : (e._docs || []);
    const inp = document.getElementById("buscador-docs");
    const f = (inp.value || "").toLowerCase();
    const vis = f ? docs.filter(d => d.n.toLowerCase().indexOf(f) >= 0) : docs;
    let h = "";
    vis.forEach(d => { h += htmlDoc(d, c); });
    if (!docs.length) h = '<p class="contador">' + (e._subcarpeta === undefined && !MODO_DEMO && !c._docsCargados ? "Cargando..." : "La documentación de este expediente está en la carpeta general del cliente.") + '</p>';
    else if (!vis.length) h = '<div class="vacio">Ningún documento coincide</div>';
    const cont = document.getElementById("docs-exp");
    if (cont) cont.innerHTML = h;
  };
  document.getElementById("buscador-docs").addEventListener("input", render);
  if (MODO_DEMO) { render(); return; }
  cargarDocsCliente(c).then(() => cargarDocsExpediente(c, e)).then(render).catch(render);
}

/* ---------- ficha CONTRARIO ---------- */
function abrirContrario(i) {
  if (i < 0) return;
  const c = CONTRA[i];
  ponerEstado("Ficha de contrario");
  const abiertos = c.exps.filter(j => EXPS[j].e.estado === "Abierto").length;
  let html = '<div class="cabecera-ficha contrario"><div class="fila-id">' +
    '<div class="avatar-g coral"><i class="ti ti-user"></i></div>' +
    '<div><h2>' + esc(c.nombre) + '</h2><p class="sub">Contrario · ' + c.exps.length + ' procedimiento' + (c.exps.length > 1 ? "s" : "") +
    ' (' + abiertos + ' abierto' + (abiertos === 1 ? "" : "s") + ')</p></div></div>' +
    (c.exps.length > 1 ? '<p class="sub" style="margin-top:12px;"><i class="ti ti-info-circle" style="font-size:13px;"></i> Este contrario aparece en varios pleitos — conviene revisar el histórico antes del juicio</p>' : "") +
    '</div><p class="seccion">SUS PROCEDIMIENTOS</p>';
  c.exps.forEach(j => { html += htmlExp(j); });
  pintar(html);
}

/* ---------- AJUSTES ---------- */
function pintarAjustes() {
  ponerEstado("Ajustes");
  const fc = FUENTE.fClientes ? FUENTE.fClientes.toLocaleDateString("es-ES") : "—";
  const fe = FUENTE.fExpedientes ? FUENTE.fExpedientes.toLocaleDateString("es-ES") : "—";
  const sinTel = CLIENTES.filter(c => c.estado !== "Baja" && !c.movil && !c.tel).length;
  const sinMail = CLIENTES.filter(c => c.estado !== "Baja" && !c.mail).length;
  let html = '<p class="seccion" style="margin-top:18px;">DATOS</p>' +
    '<div class="fila-ajuste"><i class="ti ti-users"></i><div class="cuerpo"><p class="t">Clientes</p><p class="s">' + CLIENTES.length + ' cargados · exportación del ' + fc + '</p></div></div>' +
    '<div class="fila-ajuste"><i class="ti ti-briefcase"></i><div class="cuerpo"><p class="t">Expedientes</p><p class="s">' + EXPS.length + ' cargados · exportación del ' + fe + '</p></div></div>' +
    '<div class="fila-ajuste"><i class="ti ti-calendar"></i><div class="cuerpo"><p class="t">Agenda</p><p class="s">' + CITAS.length + ' citas · en tiempo real desde Outlook</p></div></div>' +
    (MODO_DEMO ? "" : '<div class="fila-ajuste" style="cursor:pointer;" onclick="refrescar()"><i class="ti ti-refresh"></i><div class="cuerpo"><p class="t">Actualizar ahora</p><p class="s">Vuelve a leer las exportaciones y el calendario</p></div><i class="ti ti-chevron-right flecha"></i></div>') +
    '<p class="seccion">CALIDAD DE DATOS</p>' +
    '<div class="metricas" style="margin-top:6px;">' +
    '<div class="metrica ambar"><p class="num">' + sinTel + '</p><p class="lbl">clientes sin teléfono</p></div>' +
    '<div class="metrica ambar"><p class="num">' + sinMail + '</p><p class="lbl">clientes sin email</p></div></div>' +
    '<div class="fila-ajuste" style="cursor:pointer;" onclick="irA(pintarCalidad)"><i class="ti ti-list-check"></i><div class="cuerpo"><p class="t">Ver clientes incompletos</p><p class="s">Para ir completándolos en MN Program</p></div><i class="ti ti-chevron-right flecha"></i></div>' +
    '<p class="seccion">SESIÓN</p>' +
    '<div class="fila-ajuste"><i class="ti ti-lock"></i><div class="cuerpo"><p class="t">' + (MODO_DEMO ? "Modo demo" : "Conectado como " + esc((cuenta && cuenta.username) || "")) + '</p><p class="s">Solo lectura · los datos no salen de Microsoft 365</p></div></div>' +
    (MODO_DEMO ? '<p class="contador" style="margin:14px 18px;">Para conectar con tus datos reales, sigue la GUIA-INSTALACION y rellena config.js</p>'
      : '<div class="fila-ajuste" style="cursor:pointer;" onclick="cerrarSesion()"><i class="ti ti-logout"></i><div class="cuerpo"><p class="t">Cerrar sesión</p></div></div>');
  pintar(html);
}
function pintarCalidad() {
  ponerEstado("Calidad de datos");
  const incompletos = CLIENTES.map((c, i) => ({ c, i }))
    .filter(o => o.c.estado !== "Baja" && !o.c.movil && !o.c.tel && !o.c.mail && o.c.n.indexOf("VARIOS") < 0);
  let html = '<p class="contador" style="margin-top:16px;">Clientes en alta sin ningún dato de contacto (' + incompletos.length + '). Al completarlos en MN Program, entrarán con la siguiente exportación.</p>';
  incompletos.slice(0, 80).forEach(o => {
    html += '<div class="fila" onclick="abrirClienteNav(' + o.i + ')"><div class="avatar" style="background:var(--ambar-bg);color:var(--ambar);">' + esc(iniciales(o.c.n)) + '</div>' +
      '<div class="cuerpo"><p class="titulo">' + esc(o.c.n) + '</p><p class="detalle">' + o.c.exps.length + ' expedientes</p></div><i class="ti ti-chevron-right flecha"></i></div>';
  });
  pintar(html);
}
function refrescar() {
  localStorage.removeItem("nacar_carpetas");
  localStorage.removeItem("nacar_carpetas2");
  mapaCarpetas = null;
  cargarTodo();
}
function cerrarSesion() {
  if (msalApp) msalApp.logoutRedirect({ account: cuenta });
}

/* ============================================================
   CALCULADORA DE INDEMNIZACIONES
   Reglas: Guía CGPJ oct-2024 (skill calculo-indemnizaciones del despacho)
   - Salario diario = bruto anual / 365 (STS 22-2-2020)
   - Meses redondeados al alza, fecha de efectos inclusive (SSTS 20-7-2009 y otras)
   - Improcedente: 33 d/año (2,75/mes), tope 720 días; doble tramo DT 11ª si
     el contrato es anterior al 12-02-2012 (45 d/año el tramo 1; topes STS 18-2-2016)
   - Objetivo/colectivo: 20 d/año (20/12 por mes), tope 360 días; MSCT tope 270
   - Temporal: días naturales x coeficiente DT 8ª ET (8 a 12 d/año), sin prorrateo
   ============================================================ */
const FMT_EUR = new Intl.NumberFormat("es-ES", { style: "currency", currency: "EUR" });
// Meses indemnizatorios (Guía CGPJ pág. 5): fecha de efectos INCLUSIVE y los días
// que excedan de un mes completo cuentan como un mes entero. Cálculo con enteros
// año/mes/día para que sea inmune a la zona horaria y a los cambios de hora.
function mesesIndem(inicio, finInclusive) {
  let y = finInclusive.getFullYear(), m = finInclusive.getMonth(), d = finInclusive.getDate() + 1;
  const diasMes = new Date(y, m + 1, 0).getDate();
  if (d > diasMes) { d = 1; m++; if (m > 11) { m = 0; y++; } }
  let meses = (y - inicio.getFullYear()) * 12 + (m - inicio.getMonth());
  const resto = d - inicio.getDate();
  if (resto < 0) meses--;
  if (meses < 0) return 0;
  return meses + (resto !== 0 ? 1 : 0);
}
// lee un <input type="date"> como fecha local exacta (evita líos UTC de valueAsDate)
function fechaCampo(id) {
  const v = document.getElementById(id).value;
  if (!v) return null;
  const p = v.split("-").map(Number);
  return new Date(p[0], p[1] - 1, p[2]);
}
function pintarCalc() {
  ponerEstado("Calculadora");
  pintar('<div style="padding:16px;max-width:560px;">' +
    '<p style="margin:0 0 2px;font-size:17px;font-weight:600;">Calculadora de indemnización</p>' +
    '<p style="margin:0 0 14px;font-size:12px;color:var(--texto-2);">Criterios de la Guía CGPJ (oct. 2024) · cálculo orientativo</p>' +
    '<label class="lbl-calc">Tipo de extinción</label>' +
    '<select id="c-tipo" class="campo-calc">' +
    '<option value="improcedente">Despido improcedente / art. 50 ET</option>' +
    '<option value="objetivo">Despido objetivo o colectivo (20 días)</option>' +
    '<option value="msct">Modificación sustancial – MSCT (20 días, tope 9 mens.)</option>' +
    '<option value="temporal">Fin de contrato temporal (DT 8ª ET)</option></select>' +
    '<label class="lbl-calc">Fecha de inicio (antigüedad)</label>' +
    '<input type="date" id="c-inicio" class="campo-calc" />' +
    '<label class="lbl-calc">Fecha de efectos del despido / fin</label>' +
    '<input type="date" id="c-fin" class="campo-calc" />' +
    '<label class="lbl-calc">Salario bruto</label>' +
    '<div style="display:flex;gap:8px;">' +
    '<input type="number" id="c-salario" class="campo-calc" style="flex:2;" placeholder="Importe en euros" inputmode="decimal" />' +
    '<select id="c-unidad" class="campo-calc" style="flex:1.4;"><option value="anual">anual</option><option value="mensual">mensual</option></select>' +
    '<select id="c-pagas" class="campo-calc" style="flex:1;display:none;"><option value="12">12 pagas</option><option value="14" selected>14 pagas</option><option value="15">15 pagas</option></select></div>' +
    '<button class="boton-principal" style="width:100%;margin-top:14px;" onclick="calcularIndem()"><i class="ti ti-calculator" style="vertical-align:-2px;"></i> Calcular</button>' +
    '<div id="c-resultado"></div>' +
    '<p style="margin:16px 0 0;font-size:11px;color:var(--texto-3);">No incluye salarios de tramitación, FOGASA ni mejoras de convenio. El salario en especie (seguro médico, vehículo...) computa como salario: si existe, súmalo al bruto.</p></div>');
  document.getElementById("c-unidad").addEventListener("change", function () {
    document.getElementById("c-pagas").style.display = this.value === "mensual" ? "block" : "none";
  });
}
function calcularIndem() {
  const tipo = document.getElementById("c-tipo").value;
  const inicio = fechaCampo("c-inicio");
  const fin = fechaCampo("c-fin");
  const importe = parseFloat(document.getElementById("c-salario").value);
  const unidad = document.getElementById("c-unidad").value;
  const pagas = parseInt(document.getElementById("c-pagas").value, 10);
  const out = document.getElementById("c-resultado");
  if (!inicio || !fin || !importe || fin <= inicio) {
    out.innerHTML = '<p style="margin:12px 0 0;font-size:13px;color:var(--rojo,#A32D2D);">Revisa los datos: faltan fechas o salario, o el fin es anterior al inicio.</p>';
    return;
  }
  const anual = unidad === "anual" ? importe : importe * pagas;
  const sd = anual / 365;
  const filas = [];
  filas.push(["Salario bruto anual", FMT_EUR.format(anual)]);
  filas.push(["Salario diario (÷365)", FMT_EUR.format(sd)]);
  let dias = 0, topeTxt = "";
  if (tipo === "improcedente") {
    const corte = new Date(2012, 1, 12);
    if (inicio < corte) {
      const finT1 = fin < corte ? fin : new Date(2012, 1, 11);
      const m1 = mesesIndem(inicio, finT1);
      const m2 = fin >= corte ? mesesIndem(corte, fin) : 0;
      const d1 = m1 * 3.75, d2 = m2 * 2.75;
      filas.push(["Antigüedad total", mesesIndem(inicio, fin) + " meses"]);
      filas.push(["Tramo 1 (hasta 11-02-2012, 45 d/año)", m1 + " meses × 3,75 = " + d1.toFixed(2) + " días"]);
      filas.push(["Tramo 2 (desde 12-02-2012, 33 d/año)", m2 + " meses × 2,75 = " + d2.toFixed(2) + " días"]);
      if (d1 >= 720) {
        dias = Math.min(d1, 1260);
        topeTxt = d1 > 1260 ? "Aplicado tope absoluto de 42 mensualidades (1.260 días)" : "El tramo 1 ya supera 720 días: se respeta lo devengado a 11-02-2012 (máx. 1.260)";
      } else {
        dias = Math.min(d1 + d2, 720);
        if (d1 + d2 > 720) topeTxt = "Aplicado tope global de 24 mensualidades (720 días)";
      }
    } else {
      const m = mesesIndem(inicio, fin);
      filas.push(["Antigüedad computada", m + " meses (redondeo al alza)"]);
      filas.push(["Devengo (33 d/año)", m + " × 2,75 = " + (m * 2.75).toFixed(2) + " días"]);
      dias = Math.min(m * 2.75, 720);
      if (m * 2.75 > 720) topeTxt = "Aplicado tope de 24 mensualidades (720 días)";
    }
  } else if (tipo === "objetivo" || tipo === "msct") {
    const m = mesesIndem(inicio, fin);
    const tope = tipo === "msct" ? 270 : 360;
    filas.push(["Antigüedad computada", m + " meses (redondeo al alza)"]);
    filas.push(["Devengo (20 d/año)", m + " × 20/12 = " + (m * 20 / 12).toFixed(2) + " días"]);
    dias = Math.min(m * 20 / 12, tope);
    if (m * 20 / 12 > tope) topeTxt = "Aplicado tope de " + (tope / 30) + " mensualidades (" + tope + " días)";
  } else {
    const diasTrab = Math.round((fin - inicio) / 86400000) + 1;
    const y = inicio.getFullYear();
    const coef = y <= 2011 ? 8 : y === 2012 ? 9 : y === 2013 ? 10 : y === 2014 ? 11 : 12;
    filas.push(["Días naturales trabajados", diasTrab + " días"]);
    filas.push(["Coeficiente DT 8ª (contrato de " + y + ")", coef + " días/año"]);
    dias = diasTrab * coef / 365;
  }
  const total = sd * dias;
  let html = '<div style="margin-top:14px;border:0.5px solid var(--borde);border-radius:12px;overflow:hidden;">';
  filas.forEach(f => { html += '<div style="display:flex;justify-content:space-between;gap:10px;padding:8px 12px;border-bottom:0.5px solid var(--borde);font-size:13px;"><span style="color:var(--texto-2);">' + f[0] + '</span><span style="text-align:right;">' + f[1] + '</span></div>'; });
  html += '<div style="display:flex;justify-content:space-between;padding:11px 12px;background:var(--teal-bg);font-size:14px;"><span style="font-weight:600;">Indemnización</span><span style="font-weight:600;color:var(--teal);">' + FMT_EUR.format(total) + '</span></div></div>' +
    '<p style="margin:8px 0 0;font-size:12px;color:var(--texto-2);">' + dias.toFixed(2) + ' días de salario' + (topeTxt ? ' · ' + topeTxt : '') + '</p>';
  out.innerHTML = html;
}

/* expone funciones usadas desde HTML inline */
Object.assign(window, {
  irTab, abrirClienteNav, abrirExpedienteNav, abrirContrarioNav, abrirCitaNav,
  abrirDoc, compartirDoc, cambiarFiltro, pintarResultadosAtajo, refrescar,
  cerrarSesion, pintarCalidad, irA, cargarTodo, pintarCalc, calcularIndem, abrirDocUrl
});

init().catch(e => {
  pintar('<div class="vacio">Error al iniciar: ' + esc(e.message) + '</div>');
});
