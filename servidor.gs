/**
 * Alfajores Vicencio · bandeja de pedidos en la nube (Google Apps Script).
 * Recibe los pedidos de la página pública aunque tu PC esté apagado y los guarda en una hoja de Google.
 * Tu ERP los retira cada vez que está abierto.
 * (Este archivo se carga solo desde el cargador de Apps Script; la clave TOKEN la aporta el cargador.)
 */
const HOJA = "Pedidos";
const COLUMNAS = ["ref", "creado", "nombre", "telefono", "nota", "lineas", "total", "importado", "entrega", "comuna", "direccion", "envio"];
const HOJA_OP = "Opiniones";
const COL_OP = ["ref", "creado", "estrellas", "comentario", "importado"];

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

// Despacho: gratis desde el mínimo en las comunas de reparto; bajo eso se cobra el costo. Otras comunas: siempre el costo.
function calcularEnvio(c, subtotal, entrega, comuna) {
  const r = c.despacho || {};
  if (entrega !== "despacho" || !r.activo) return 0;
  if ((r.comunas || []).indexOf(comuna) < 0) return Number(r.costo || 0);
  return subtotal >= Number(r.minimo || 0) ? 0 : Number(r.costo || 0);
}

// La página pública pide aquí el catálogo (sin datos de pago).
function doGet(e) {
  const c = leerCfg();
  return json({ ok: true, activo: c.activo !== false && (c.productos || []).length > 0, productos: c.productos || [],
                entrega: c.entrega || "", whatsapp: c.whatsapp || "", despacho: c.despacho || { activo: false } });
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: "Solicitud inválida." }); }
  if (d.accion === "pedido") return crearPedido(d);
  if (d.accion === "opinion") return guardarOpinion(d);
  if (d.token !== TOKEN) return json({ ok: false, error: "No autorizado." });
  if (d.accion === "publicar") {
    PropertiesService.getScriptProperties().setProperty("cfg", JSON.stringify(d.cfg || {}));
    return json({ ok: true });
  }
  if (d.accion === "pendientes") return pendientes();
  if (d.accion === "marcar") return marcar(d.refs || []);
  if (d.accion === "opiniones") return opinionesPendientes();
  if (d.accion === "marcar_opiniones") return marcarOpiniones(d.refs || []);
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
    const ref = String(d.ref || "").replace(/[^a-zA-Z0-9-]/g, "").substring(0, 40);
    if (ref.length < 8) return json({ ok: false, error: "Solicitud inválida." });
    const h = hoja();
    const filas = h.getDataRange().getValues();
    const idx = {}; COLUMNAS.forEach((k, i) => idx[k] = i);
    for (let i = 1; i < filas.length; i++) {   // mismo envío repetido (doble toque): devuelve el mismo pedido
      if (String(filas[i][idx.ref]) === ref) {
        return json({ ok: true, pedido: { ref: ref, nombre: filas[i][idx.nombre], lineas: JSON.parse(filas[i][idx.lineas]), total: filas[i][idx.total],
                                          entrega: filas[i][idx.entrega] || "retiro", comuna: filas[i][idx.comuna] || "", direccion: filas[i][idx.direccion] || "",
                                          envio: Number(filas[i][idx.envio] || 0) }, pago: c.pago || {} });
      }
    }
    const catalogo = {}; (c.productos || []).forEach(p => catalogo[p.id] = p);
    const lineas = []; let total = 0, unidades = 0;
    Object.keys(d.cantidades || {}).forEach(k => {
      const p = catalogo[k]; const cant = parseInt(d.cantidades[k], 10);
      if (!p || !(cant > 0)) return;
      if (cant > 100) throw new Error("Para pedidos de más de 100 alfajores, escríbenos por WhatsApp.");
      lineas.push({ producto_id: p.id, nombre: p.nombre, cantidad: cant, precio: p.precio });
      total += cant * p.precio; unidades += cant;
    });
    if (unidades <= 0) return json({ ok: false, error: "Elige al menos 1 alfajor." });
    if (unidades > 200) return json({ ok: false, error: "Para pedidos tan grandes, escríbenos por WhatsApp." });
    const despacho = d.entrega === "despacho" && (c.despacho || {}).activo;
    const entrega = despacho ? "despacho" : "retiro";
    const comuna = despacho ? String(d.comuna || "").replace(/\s+/g, " ").trim().substring(0, 40) : "";
    const direccion = despacho ? String(d.direccion || "").replace(/\s+/g, " ").trim().substring(0, 120) : "";
    if (despacho && (!comuna || direccion.length < 5)) return json({ ok: false, error: "Escribe tu comuna y dirección para el despacho." });
    const envio = calcularEnvio(c, total, entrega, comuna);
    const creado = Utilities.formatDate(new Date(), "America/Santiago", "yyyy-MM-dd HH:mm");
    h.appendRow([ref, "'" + creado, nombre, "'" + tel, nota, JSON.stringify(lineas), total + envio, "", entrega, comuna, direccion, envio]);
    cache.put(clave, String(n + 1), 3600);
    return json({ ok: true, pedido: { ref: ref, nombre: nombre, lineas: lineas, total: total + envio, creado: creado,
                                      entrega: entrega, comuna: comuna, direccion: direccion, envio: envio }, pago: c.pago || {} });
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
                  envio: Number(filas[i][11] || 0) });
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
