/**
 * Alfajores Vicencio · bandeja de pedidos en la nube (Google Apps Script).
 * Recibe los pedidos de la página pública aunque tu PC esté apagado y los guarda en una hoja de Google.
 * Tu ERP los retira cada vez que está abierto.
 * (Este archivo se carga solo desde el cargador de Apps Script; la clave TOKEN la aporta el cargador.)
 */
const HOJA = "Pedidos";
const COLUMNAS = ["ref", "creado", "nombre", "telefono", "nota", "lineas", "total", "importado", "entrega", "comuna", "direccion", "envio", "correo", "mp_link", "mp_pago", "mp_avisado", "fecha_entrega"];
const HOJA_OP = "Opiniones";
const COL_OP = ["ref", "creado", "estrellas", "comentario", "importado"];
const HOJA_COMP = "Comprobantes";
const COL_COMP = ["ref", "creado", "mime", "importado"];   // desde la columna 5: la imagen del comprobante, en trozos de texto (una celda aguanta 50.000 caracteres)
const TROZO = 45000, MAX_TROZOS = 10;

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function leerCfg() {
  const t = PropertiesService.getScriptProperties().getProperty("cfg");
  try { return t ? JSON.parse(t) : {}; } catch (e) { return {}; }
}

function hoja() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let h = libro.getSheetByName(HOJA);
  if (!h) { h = libro.insertSheet(HOJA); h.appendRow(COLUMNAS); h.setFrozenRows(1); }
  if (h.getLastColumn() < COLUMNAS.length) h.getRange(1, 1, 1, COLUMNAS.length).setValues([COLUMNAS]);  // hoja antigua: agrega los títulos nuevos
  return h;
}

// Despacho: gratis desde el mínimo en las comunas de reparto; bajo eso se cobra el costo. Otras comunas no se aceptan (se rechazan en crearPedido).
function calcularEnvio(c, subtotal, entrega, comuna) {
  const r = c.despacho || {};
  if (entrega !== "despacho" || !r.activo) return 0;
  if ((r.comunas || []).indexOf(comuna) < 0) return Number(r.costo || 0);
  return subtotal >= Number(r.minimo || 0) ? 0 : Number(r.costo || 0);
}

// La página pública pide aquí el catálogo (sin datos de pago).
function doGet(e) {
  const c = leerCfg();
  const ag = c.agenda && c.agenda.desde ? { activo: true, dias: diasDisponibles(c, null).map(x => ({ fecha: x.fecha, libre: x.libre })) } : { activo: false };
  return json({ ok: true, activo: c.activo !== false && (c.productos || []).length > 0,
                productos: (c.productos || []).map(p => publico(p, false)),
                entrega: c.entrega || "", whatsapp: c.whatsapp || "", despacho: c.despacho || { activo: false }, mp: !!(c.mp && mpToken()),
                amigo: !!amigoCodigo(), agenda: ag });
}

function doPost(e) {
  const q = (e && e.parameter) || {};
  if (q.accion === "mp") return avisoMp(e);   // aviso automático de Mercado Pago (se verifica siempre con la API)
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: "Solicitud inválida." }); }
  if (d.accion === "pedido") return crearPedido(d);
  if (d.accion === "mp_link") return mpLink(d);
  if (d.accion === "estado_mp") return estadoMp(d);
  if (d.accion === "amigo") return validarAmigo(d);
  if (d.accion === "opinion") return guardarOpinion(d);
  if (d.accion === "comprobante") return guardarComprobante(d);
  if (d.token !== TOKEN) return json({ ok: false, error: "No autorizado." });
  if (d.accion === "publicar") {
    const props = PropertiesService.getScriptProperties();
    props.setProperty("cfg", JSON.stringify(d.cfg || {}));
    if (typeof d.amigo_codigo === "string") { if (d.amigo_codigo) props.setProperty("amigo_codigo", d.amigo_codigo); else props.deleteProperty("amigo_codigo"); }
    if (typeof d.mp_token === "string") { if (d.mp_token) props.setProperty("mp_token", d.mp_token); else props.deleteProperty("mp_token"); }
    return json({ ok: true });
  }
  if (d.accion === "pendientes") return pendientes();
  if (d.accion === "marcar") return marcar(d.refs || []);
  if (d.accion === "opiniones") return opinionesPendientes();
  if (d.accion === "marcar_opiniones") return marcarOpiniones(d.refs || []);
  if (d.accion === "comprobantes") return comprobantesPendientes();
  if (d.accion === "pagos_mp") return pagosMpPendientes();
  if (d.accion === "marcar_pagos_mp") return marcarPagosMp(d.refs || []);
  if (d.accion === "marcar_comprobantes") return marcarComprobantes(d.refs || []);
  return json({ ok: false, error: "Acción desconocida." });
}

function crearPedido(d) {
  if (d.trampa) return json({ ok: true, pedido: { ref: "x", nombre: "", lineas: [], total: 0 }, pago: {} }); // robot: se simula éxito
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const c = leerCfg();
    if (c.activo === false) return json({ ok: false, error: "Por ahora no estamos tomando pedidos." });
    const cache = CacheService.getScriptCache();
    const clave = "n" + Utilities.formatDate(new Date(), "UTC", "yyyyMMddHH");
    const n = Number(cache.get(clave) || 0);
    if (n >= 40) return json({ ok: false, error: "Hay muchos pedidos seguidos. Escríbenos por WhatsApp o Instagram." });
    const nombre = String(d.nombre || "").replace(/\s+/g, " ").trim().substring(0, 60);
    const nota = String(d.nota || "").replace(/\s+/g, " ").trim().substring(0, 300);
    let tel = String(d.telefono || "").replace(/\D/g, "");
    if (tel.length === 8) tel = "569" + tel;
    if (nombre.length < 2) return json({ ok: false, error: "Escribe tu nombre." });
    if (tel.length < 8 || tel.length > 13) return json({ ok: false, error: "Escribe un teléfono válido." });
    const correo = String(d.correo || "").trim().toLowerCase().substring(0, 80);
    if (correo && !/^[^@\s]+@[^@\s]+\.[^@\s]{2,}$/.test(correo)) return json({ ok: false, error: "Escribe un correo válido (por ejemplo, nombre@gmail.com)." });
    const ref = String(d.ref || "").replace(/[^a-zA-Z0-9-]/g, "").substring(0, 40);
    if (ref.length < 8) return json({ ok: false, error: "Solicitud inválida." });
    const h = hoja();
    const filas = h.getDataRange().getValues();
    const idx = {}; COLUMNAS.forEach((k, i) => idx[k] = i);
    for (let i = 1; i < filas.length; i++) {   // mismo envío repetido (doble toque): devuelve el mismo pedido
      if (String(filas[i][idx.ref]) === ref) {
        return json({ ok: true, pedido: { ref: ref, nombre: filas[i][idx.nombre], lineas: JSON.parse(filas[i][idx.lineas]), total: filas[i][idx.total],
                                          entrega: filas[i][idx.entrega] || "retiro", comuna: filas[i][idx.comuna] || "", direccion: filas[i][idx.direccion] || "",
                                          envio: Number(filas[i][idx.envio] || 0), fecha_entrega: String(filas[i][idx.fecha_entrega] || "") }, pago: c.pago || {} });
      }
    }
    const catalogo = {}; (c.productos || []).forEach(p => catalogo[p.id] = p);
    const amigo = esAmigo(d.codigo);
    if (String(d.codigo || "").trim() && !amigo) return json({ ok: false, error: "Ese código no es válido." });
    const lineas = []; let total = 0, unidades = 0;
    Object.keys(d.cantidades || {}).forEach(k => {
      const p = catalogo[k]; const cant = parseInt(d.cantidades[k], 10);
      if (!p || !(cant > 0)) return;
      if (cant > 100) throw new Error("Para pedidos de más de 100 alfajores, escríbenos por WhatsApp.");
      const precio = precioDe(p, amigo);
      lineas.push({ producto_id: p.id, nombre: p.nombre, cantidad: cant, precio: precio });
      total += cant * precio; unidades += cant * (Number(p.navidad) || 1);
    });
    if (unidades <= 0) return json({ ok: false, error: "Elige al menos 1 alfajor." });
    let fechaEntrega = "";
    if (c.agenda && c.agenda.desde) {
      fechaEntrega = String(d.fecha_entrega || "").substring(0, 10);
      if (!fechaEntrega) return json({ ok: false, error: "Elige el día de entrega." });
      const dia = diasDisponibles(c, filas).filter(x => x.fecha === fechaEntrega)[0];
      if (!dia) return json({ ok: false, error: "Ese día de entrega ya no está disponible. Elige otro, por favor." });
      if (dia.libre < unidades) return json({ ok: false, error: "Para ese día ya no alcanzamos con tantos alfajores. Elige otro día o escríbenos por WhatsApp." });
    }
    const notaFinal = (amigo ? (nota ? nota + " · " : "") + "Con código" : nota).substring(0, 300);
    if (unidades > 200) return json({ ok: false, error: "Para pedidos tan grandes, escríbenos por WhatsApp." });
    const despacho = d.entrega === "despacho" && (c.despacho || {}).activo;
    const entrega = despacho ? "despacho" : "retiro";
    const comuna = despacho ? String(d.comuna || "").replace(/\s+/g, " ").trim().substring(0, 40) : "";
    const direccion = despacho ? String(d.direccion || "").replace(/\s+/g, " ").trim().substring(0, 120) : "";
    if (despacho && (!comuna || direccion.length < 5)) return json({ ok: false, error: "Escribe tu comuna y dirección para el despacho." });
    if (despacho && (c.despacho.comunas || []).indexOf(comuna) < 0) return json({ ok: false, error: "Por ahora solo entregamos a domicilio en " + (c.despacho.comunas || []).join(", ") + ". Puedes retirar gratis en el metro Escuela Militar." });
    const envio = calcularEnvio(c, total, entrega, comuna);
    const creado = Utilities.formatDate(new Date(), "America/Santiago", "yyyy-MM-dd HH:mm");
    h.appendRow([ref, "'" + creado, nombre, "'" + tel, notaFinal, JSON.stringify(lineas), total + envio, "", entrega, comuna, direccion, envio, correo, "", "", "", fechaEntrega ? "'" + fechaEntrega : ""]);
    cache.put(clave, String(n + 1), 3600);
    return json({ ok: true, pedido: { ref: ref, nombre: nombre, lineas: lineas, total: total + envio, creado: creado,
                                      entrega: entrega, comuna: comuna, direccion: direccion, envio: envio, fecha_entrega: fechaEntrega }, pago: c.pago || {} });
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function pendientes() {
  const filas = hoja().getDataRange().getValues();
  const salida = [];
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][7]).trim() !== "") continue;
    salida.push({ ref: String(filas[i][0]), creado: String(filas[i][1]), nombre: filas[i][2], telefono: String(filas[i][3]).replace(/\D/g, ""),
                  nota: filas[i][4], lineas: JSON.parse(filas[i][5]), total: filas[i][6],
                  entrega: String(filas[i][8] || "retiro"), comuna: String(filas[i][9] || ""), direccion: String(filas[i][10] || ""),
                  envio: Number(filas[i][11] || 0), correo: String(filas[i][12] || ""), fecha_entrega: String(filas[i][16] || "") });
  }
  return json({ ok: true, pedidos: salida });
}

function marcar(refs) {
  const h = hoja();
  const filas = h.getDataRange().getValues();
  const set = {}; refs.forEach(r => set[r] = true);
  const ahora = Utilities.formatDate(new Date(), "America/Santiago", "yyyy-MM-dd HH:mm");
  for (let i = 1; i < filas.length; i++) {
    if (set[String(filas[i][0])]) h.getRange(i + 1, 8).setValue(ahora);
  }
  return json({ ok: true });
}

// ---------- Opiniones de clientes (se piden cuando marcas un pedido como entregado) ----------
function hojaOp() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let h = libro.getSheetByName(HOJA_OP);
  if (!h) { h = libro.insertSheet(HOJA_OP); h.appendRow(COL_OP); h.setFrozenRows(1); }
  return h;
}

function guardarOpinion(d) {
  if (d.trampa) return json({ ok: true });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ref = String(d.ref || "").replace(/[^a-zA-Z0-9-]/g, "").substring(0, 40);
    const estrellas = parseInt(d.estrellas, 10);
    if (ref.length < 8 || !(estrellas >= 1 && estrellas <= 5)) return json({ ok: false, error: "Elige de 1 a 5 estrellas." });
    const comentario = String(d.comentario || "").replace(/\s+/g, " ").trim().substring(0, 500);
    const pedidos = hoja().getDataRange().getValues();
    let existe = false;
    for (let i = 1; i < pedidos.length; i++) if (String(pedidos[i][0]) === ref) { existe = true; break; }
    if (!existe) return json({ ok: false, error: "No encontramos tu pedido." });   // solo opinan clientes reales
    const h = hojaOp();
    const filas = h.getDataRange().getValues();
    const creado = Utilities.formatDate(new Date(), "America/Santiago", "yyyy-MM-dd HH:mm");
    for (let i = 1; i < filas.length; i++) {
      if (String(filas[i][0]) === ref) {   // si vuelve a opinar, se actualiza y se vuelve a avisar
        h.getRange(i + 1, 2, 1, 4).setValues([[creado, estrellas, comentario, ""]]);
        return json({ ok: true });
      }
    }
    h.appendRow([ref, "'" + creado, estrellas, comentario, ""]);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function opinionesPendientes() {
  const filas = hojaOp().getDataRange().getValues();
  const salida = [];
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][4]).trim() !== "") continue;
    salida.push({ ref: String(filas[i][0]), creado: String(filas[i][1]), estrellas: Number(filas[i][2]), comentario: String(filas[i][3]) });
  }
  return json({ ok: true, opiniones: salida });
}

function marcarOpiniones(refs) {
  const h = hojaOp();
  const filas = h.getDataRange().getValues();
  const set = {}; refs.forEach(r => set[r] = true);
  const ahora = Utilities.formatDate(new Date(), "America/Santiago", "yyyy-MM-dd HH:mm");
  for (let i = 1; i < filas.length; i++) if (set[String(filas[i][0])]) h.getRange(i + 1, 5).setValue(ahora);
  return json({ ok: true });
}

// ---------- Comprobantes de transferencia (el cliente los adjunta en la página después de pagar) ----------
// La foto llega reducida desde el navegador. Se guarda en una hoja aparte y el ERP la baja a tu PC y la borra de la nube.
function hojaComp() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let h = libro.getSheetByName(HOJA_COMP);
  if (!h) { h = libro.insertSheet(HOJA_COMP); h.appendRow(COL_COMP); h.setFrozenRows(1); }
  return h;
}

function guardarComprobante(d) {
  if (d.trampa) return json({ ok: true });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ref = String(d.ref || "").replace(/[^a-zA-Z0-9-]/g, "").substring(0, 40);
    if (ref.length < 8) return json({ ok: false, error: "Solicitud inválida." });
    const mime = String(d.mime || "").toLowerCase();
    if (["image/jpeg", "image/png", "image/webp"].indexOf(mime) < 0) return json({ ok: false, error: "Sube una foto o captura del comprobante (JPG o PNG)." });
    const datos = String(d.datos || "").replace(/\s+/g, "");
    if (datos.length < 200 || !/^[A-Za-z0-9+\/]+={0,2}$/.test(datos)) return json({ ok: false, error: "No pude leer esa imagen. Prueba con otra o envíala por WhatsApp." });
    if (datos.length > TROZO * MAX_TROZOS) return json({ ok: false, error: "La imagen es muy pesada. Prueba con una captura de pantalla o envíala por WhatsApp." });
    const pedidos = hoja().getDataRange().getValues();
    let existe = false;
    for (let i = 1; i < pedidos.length; i++) if (String(pedidos[i][0]) === ref) { existe = true; break; }
    if (!existe) return json({ ok: false, error: "No encontramos tu pedido." });   // solo adjuntan clientes reales
    const cache = CacheService.getScriptCache();
    const kRef = "c" + ref, kHora = "cg" + Utilities.formatDate(new Date(), "UTC", "yyyyMMddHH");
    const nRef = Number(cache.get(kRef) || 0), nHora = Number(cache.get(kHora) || 0);
    if (nRef >= 6 || nHora >= 80) return json({ ok: false, error: "Ya recibimos varios envíos. Escríbenos por WhatsApp con tu comprobante." });
    const trozos = [];
    for (let i = 0; i < MAX_TROZOS; i++) trozos.push(i * TROZO < datos.length ? "'" + datos.substring(i * TROZO, (i + 1) * TROZO) : "");
    const creado = Utilities.formatDate(new Date(), "America/Santiago", "yyyy-MM-dd HH:mm");
    const fila = [ref, "'" + creado, mime, ""].concat(trozos);
    const h = hojaComp();
    const filas = h.getDataRange().getValues();
    let hecho = false;
    for (let i = 1; i < filas.length; i++) {
      if (String(filas[i][0]) === ref) {   // si lo vuelve a enviar, reemplaza al anterior y se vuelve a avisar
        h.getRange(i + 1, 1, 1, fila.length).setValues([fila]);
        hecho = true; break;
      }
    }
    if (!hecho) h.appendRow(fila);
    cache.put(kRef, String(nRef + 1), 3600); cache.put(kHora, String(nHora + 1), 3600);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function comprobantesPendientes() {
  const filas = hojaComp().getDataRange().getValues();
  const salida = [];
  for (let i = 1; i < filas.length && salida.length < 3; i++) {   // de a 3, para que la respuesta no sea gigante
    if (String(filas[i][3]).trim() !== "") continue;
    const datos = filas[i].slice(4).map(String).join("");
    if (!datos) continue;
    salida.push({ ref: String(filas[i][0]), creado: String(filas[i][1]), mime: String(filas[i][2]), datos: datos });
  }
  return json({ ok: true, comprobantes: salida });
}

function marcarComprobantes(refs) {
  const h = hojaComp();
  const filas = h.getDataRange().getValues();
  const set = {}; refs.forEach(r => set[r] = true);
  const ahora = Utilities.formatDate(new Date(), "America/Santiago", "yyyy-MM-dd HH:mm");
  const vacio = []; for (let i = 0; i < MAX_TROZOS; i++) vacio.push("");
  for (let i = 1; i < filas.length; i++) {
    if (set[String(filas[i][0])]) {
      h.getRange(i + 1, 4).setValue(ahora);                         // recibido por el ERP
      h.getRange(i + 1, 5, 1, MAX_TROZOS).setValues([vacio]);      // y se borra la imagen de la nube
    }
  }
  return json({ ok: true });
}

// ---------- Pago con tarjeta (Mercado Pago Checkout Pro) ----------
// El ERP publica la clave (Access Token) de tu cuenta de Mercado Pago; queda guardada solo en este script, nunca en la página.
// Flujo: la página pide un link de pago para su pedido → el cliente paga en Mercado Pago → Mercado Pago avisa aquí (y la página
// vuelve a preguntar al regresar) → se verifica con la API que el pago esté aprobado y sea por el total → el ERP lo registra.
function mpToken() { return PropertiesService.getScriptProperties().getProperty("mp_token") || ""; }

function mpApi(metodo, ruta, cuerpo) {
  const op = { method: metodo, headers: { Authorization: "Bearer " + mpToken() }, muteHttpExceptions: true };
  if (cuerpo) { op.contentType = "application/json"; op.payload = JSON.stringify(cuerpo); }
  const r = UrlFetchApp.fetch("https://api.mercadopago.com" + ruta, op);
  let j = {}; try { j = JSON.parse(r.getContentText()); } catch (e) {}
  if (r.getResponseCode() >= 300) throw new Error("Mercado Pago: " + ((j && (j.message || j.error)) || r.getResponseCode()));
  return j;
}

function limpiarRef(r) { return String(r || "").replace(/[^a-zA-Z0-9-]/g, "").substring(0, 40); }

function buscarFila(h, ref) {
  const filas = h.getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) if (String(filas[i][0]) === ref) return { n: i + 1, f: filas[i] };
  return null;
}

function mpLink(d) {
  if (!mpToken()) return json({ ok: false, error: "El pago con tarjeta no está disponible. Transfiere, por favor." });
  const ref = limpiarRef(d.ref);
  if (ref.length < 8) return json({ ok: false, error: "Solicitud inválida." });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const h = hoja(), x = buscarFila(h, ref), I = {}; COLUMNAS.forEach((k, i) => I[k] = i);
    if (!x) return json({ ok: false, error: "No encontramos tu pedido." });
    if (String(x.f[I.mp_pago] || "")) return json({ ok: true, pagado: true });
    if (String(x.f[I.mp_link] || "")) return json({ ok: true, url: String(x.f[I.mp_link]) });
    const total = Math.round(Number(x.f[I.total] || 0));
    if (!(total > 0)) return json({ ok: false, error: "Tu pedido no tiene total." });
    const lineas = JSON.parse(x.f[I.lineas] || "[]");
    const detalle = lineas.map(l => l.cantidad + " × " + l.nombre).join(", ").substring(0, 250);
    let volver = String(d.volver || "");
    if (!/^https:\/\/[a-z0-9-]+\.github\.io\/[A-Za-z0-9_\-\/.]*$/.test(volver)) volver = "";
    const pref = { items: [{ id: ref, title: "Alfajores Vicencio", description: detalle, quantity: 1, unit_price: total, currency_id: "CLP" }],
                   external_reference: ref, statement_descriptor: "VICENCIO",
                   payer: { name: String(x.f[I.nombre] || ""), email: String(x.f[I.correo] || "") || undefined } };
    if (volver) {
      const v = volver + "?ref=" + ref + "&mp=1";
      pref.back_urls = { success: v, pending: v, failure: v }; pref.auto_return = "approved";
    }
    try { const u = ScriptApp.getService().getUrl(); if (u) pref.notification_url = u + "?accion=mp"; } catch (e) {}
    const r = mpApi("post", "/checkout/preferences", pref);
    if (!r.init_point) return json({ ok: false, error: "No pudimos abrir el pago con tarjeta. Transfiere, por favor." });
    h.getRange(x.n, I.mp_link + 1).setValue(r.init_point);
    return json({ ok: true, url: r.init_point });
  } catch (err) {
    return json({ ok: false, error: "No pudimos abrir el pago con tarjeta. Transfiere, por favor." });
  } finally {
    lock.releaseLock();
  }
}

// Revisa en Mercado Pago si el pedido tiene un pago aprobado por el total. Si lo tiene, lo anota en la hoja para el ERP.
function mpRevisar(ref, pagoId) {
  let pagos = [];
  if (pagoId) pagos = [mpApi("get", "/v1/payments/" + encodeURIComponent(pagoId))];
  else pagos = (mpApi("get", "/v1/payments/search?sort=date_created&criteria=desc&external_reference=" + encodeURIComponent(ref)).results || []);
  const ap = pagos.filter(p => p && p.status === "approved")[0];
  if (!ap) return { pagado: false, estado: String((pagos[0] || {}).status || "") };
  ref = limpiarRef(ap.external_reference);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const h = hoja(), x = buscarFila(h, ref), I = {}; COLUMNAS.forEach((k, i) => I[k] = i);
    if (!x) return { pagado: false, estado: "sin pedido" };
    if (Number(ap.transaction_amount || 0) + 0.5 < Number(x.f[I.total] || 0)) return { pagado: false, estado: "monto" };
    if (!String(x.f[I.mp_pago] || "")) {
      const neto = ((ap.transaction_details || {}).net_received_amount);
      h.getRange(x.n, I.mp_pago + 1).setValue(JSON.stringify({ id: String(ap.id), monto: Number(ap.transaction_amount), neto: neto == null ? null : Number(neto),
                                                                 fecha: String(ap.date_approved || ""), tipo: String(ap.payment_type_id || "") }));
    }
    return { pagado: true };
  } finally {
    lock.releaseLock();
  }
}

function avisoMp(e) {
  try {
    const q = e.parameter || {};
    let d = {}; try { d = JSON.parse((e.postData || {}).contents || "{}"); } catch (err) {}
    const tipo = String(d.type || q.type || q.topic || "");
    const id = String((d.data || {}).id || q["data.id"] || q.id || "");
    if (mpToken() && tipo.indexOf("payment") >= 0 && /^\d+$/.test(id)) mpRevisar("", id);
  } catch (err) {}
  return json({ ok: true });
}

function estadoMp(d) {
  const ref = limpiarRef(d.ref);
  if (ref.length < 8 || !mpToken()) return json({ ok: false, error: "Solicitud inválida." });
  const cache = CacheService.getScriptCache(), k = "mp" + ref, n = Number(cache.get(k) || 0);
  if (n >= 30) return json({ ok: false, error: "Espera un momento y vuelve a intentar." });
  cache.put(k, String(n + 1), 3600);
  try {
    const h = hoja(), x = buscarFila(h, ref), I = {}; COLUMNAS.forEach((k2, i) => I[k2] = i);
    if (!x) return json({ ok: false, error: "No encontramos tu pedido." });
    if (String(x.f[I.mp_pago] || "")) return json({ ok: true, pagado: true });
    const r = mpRevisar(ref, "");
    return json({ ok: true, pagado: r.pagado, estado: r.estado });
  } catch (err) {
    return json({ ok: false, error: "No pudimos revisar el pago. Intenta en un momento." });
  }
}

function pagosMpPendientes() {
  const filas = hoja().getDataRange().getValues(), I = {}; COLUMNAS.forEach((k, i) => I[k] = i);
  const salida = [];
  for (let i = 1; i < filas.length; i++) {
    if (!String(filas[i][I.mp_pago] || "") || String(filas[i][I.mp_avisado] || "")) continue;
    let p = {}; try { p = JSON.parse(filas[i][I.mp_pago]); } catch (e) { continue; }
    salida.push({ ref: String(filas[i][0]), pago: p });
  }
  return json({ ok: true, pagos: salida });
}

function marcarPagosMp(refs) {
  const h = hoja(), filas = h.getDataRange().getValues(), I = {}; COLUMNAS.forEach((k, i) => I[k] = i);
  const set = {}; refs.forEach(r => set[r] = true);
  const ahora = Utilities.formatDate(new Date(), "America/Santiago", "yyyy-MM-dd HH:mm");
  for (let i = 1; i < filas.length; i++) if (set[String(filas[i][0])]) h.getRange(i + 1, I.mp_avisado + 1).setValue(ahora);
  return json({ ok: true });
}

// ---------- Precios por fecha, código amigo y días de entrega con cupo ----------
function hoySantiago() { return Utilities.formatDate(new Date(), "America/Santiago", "yyyy-MM-dd"); }

// Precio del producto hoy: el último tramo cuya fecha 'desde' ya llegó (el ERP publica los cambios de precio futuros).
// Con código amigo, el precio amigo.
function precioDe(p, amigo) {
  if (amigo && Number(p.precio_amigo) > 0) return Number(p.precio_amigo);
  const hoy = hoySantiago();
  let precio = Number(p.precio);
  (p.tramos || []).forEach(t => { if (String(t.desde) <= hoy) precio = Number(t.precio); });
  return precio;
}

function publico(p, amigo) {
  const x = { id: p.id, nombre: p.nombre, precio: precioDe(p, amigo) };
  if (p.navidad) x.navidad = p.navidad;
  return x;
}

function amigoCodigo() { return PropertiesService.getScriptProperties().getProperty("amigo_codigo") || ""; }
function normal(t) { return String(t || "").trim().toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, ""); }
function esAmigo(codigo) { const k = amigoCodigo(); return !!k && !!String(codigo || "").trim() && normal(codigo) === normal(k); }

function validarAmigo(d) {
  const cache = CacheService.getScriptCache(), clave = "am" + Utilities.formatDate(new Date(), "UTC", "yyyyMMddHH"), n = Number(cache.get(clave) || 0);
  if (n >= 60) return json({ ok: false, error: "Demasiados intentos. Prueba en un rato." });
  cache.put(clave, String(n + 1), 3600);
  if (!esAmigo(d.codigo)) return json({ ok: false, error: "Ese código no es válido." });
  return json({ ok: true, productos: (leerCfg().productos || []).map(p => publico(p, true)) });
}

function sumarDias(fecha, n) {
  const p = String(fecha).split("-").map(Number), t = new Date(Date.UTC(p[0], p[1] - 1, p[2] + n));
  return t.getUTCFullYear() + "-" + ("0" + (t.getUTCMonth() + 1)).slice(-2) + "-" + ("0" + t.getUTCDate()).slice(-2);
}

// Alfajores de un pedido (una caja de N cuenta N).
function unidadesDe(lineas, catalogo) {
  let u = 0;
  (lineas || []).forEach(l => {
    const p = catalogo[l.producto_id] || {};
    const m = /caja de navidad\s+(\d+)/i.exec(String(l.nombre || ""));
    u += Number(l.cantidad || 0) * (Number(p.navidad) || (m ? Number(m[1]) : 1));
  });
  return u;
}

// Días que todavía se pueden elegir, con lo que queda libre (tope interno; la página no lo muestra).
// Ocupado = lo que el ERP informa (pedidos vigentes) + pedidos de la hoja que el ERP aún no conoce.
function diasDisponibles(c, filas) {
  const a = c.agenda || {};
  if (!a.desde || !a.hasta) return [];
  const cupo = Number(a.cupo) || 0, antes = Number(a.dias_antes) || 0, hora = ("0" + (Number(a.hora) || 0)).slice(-2) + ":00";
  const conocidos = {}; (a.refs || []).forEach(r => conocidos[String(r)] = true);
  const ocupado = {}; Object.keys(a.ocupado || {}).forEach(k => ocupado[k] = Number(a.ocupado[k]) || 0);
  const catalogo = {}; (c.productos || []).forEach(p => catalogo[p.id] = p);
  filas = filas || hoja().getDataRange().getValues();
  const I = {}; COLUMNAS.forEach((k, i) => I[k] = i);
  for (let i = 1; i < filas.length; i++) {
    const f = String(filas[i][I.fecha_entrega] || "").replace(/^'/, "").substring(0, 10);
    if (!f || conocidos[String(filas[i][0])]) continue;
    let ls = []; try { ls = JSON.parse(filas[i][I.lineas]); } catch (e) {}
    ocupado[f] = (ocupado[f] || 0) + unidadesDe(ls, catalogo);
  }
  const ahora = Utilities.formatDate(new Date(), "America/Santiago", "yyyy-MM-dd HH:mm");
  const salida = [];
  for (let d = String(a.desde), n = 0; d <= String(a.hasta) && n < 120; d = sumarDias(d, 1), n++) {
    if (ahora >= sumarDias(d, -antes) + " " + hora) continue;   // ya cerró
    const libre = cupo - (ocupado[d] || 0);
    if (libre > 0) salida.push({ fecha: d, libre: libre });
  }
  return salida;
}
