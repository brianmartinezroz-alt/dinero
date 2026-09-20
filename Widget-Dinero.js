// Variables used by Scriptable.
// icon-color: teal; icon-glyph: wallet;
//
// ============================================================
//  D I N E R O  —  widget v3 (con lo que dice el cerebro)
//  Scriptable. El nombre del script puede ser el que quieras.
//  Las pantallas viven en la app web; aquí solo está el número
//  de la semana y la captura en dos toques.
// ============================================================

// ---------- PEGA AQUÍ TUS DOS DIRECCIONES ----------
const ENDPOINT = "https://script.google.com/macros/s/AKfycbzO06EdQhEE_Osf_DNBfPhWVaLRU9PpQUJXY-nyNIZTQJQ5BxJwR3OvRmL37ckwE2E_eg/exec"
const APP_URL  = "https://brianmartinezroz-alt.github.io/dinero/"
// ---------------------------------------------------

// El monto de la semana y los avisos los manda la app (Más → Ajustes).
function semanal() { return (leerJson(CACHE, {}).cerebro || {}).semanal || 2300 }
function cerebro() { return leerJson(CACHE, {}).cerebro || null }

const CATEGORIAS = [
  { id: "gasolina", nombre: "Gasolina", emoji: "⛽️" },
  { id: "comida",   nombre: "Comida",   emoji: "🍽" },
  { id: "antojo",   nombre: "Antojo",   emoji: "🍩" },
  { id: "otro",     nombre: "Otro",     emoji: "•" }
]

const COLOR = {
  fondo1: new Color("#132640"),
  fondo2: new Color("#0C1A2C"),
  mint:   new Color("#8CC2FF"),
  brass:  new Color("#E9C98E"),
  coral:  new Color("#F2877B"),
  texto:  new Color("#F5ECDC"),
  tenue:  new Color("#A9B6C8"),
  track:  new Color("#1B3252"),
  chip:   new Color("#1B3252")
}

const MESES = ["enero","febrero","marzo","abril","mayo","junio",
               "julio","agosto","septiembre","octubre","noviembre","diciembre"]

// ------------------------------------------------------------
// Almacén local: copia de los movimientos y cola de pendientes
// ------------------------------------------------------------
const fm = FileManager.local()
const DIR = fm.joinPath(fm.documentsDirectory(), "dinero")
if (!fm.fileExists(DIR)) fm.createDirectory(DIR, true)
const CACHE = fm.joinPath(DIR, "cache.json")
const COLA  = fm.joinPath(DIR, "cola.json")

function leerJson(ruta, porDefecto) {
  try {
    if (!fm.fileExists(ruta)) return porDefecto
    return JSON.parse(fm.readString(ruta))
  } catch (e) { return porDefecto }
}
function escribirJson(ruta, valor) {
  try { fm.writeString(ruta, JSON.stringify(valor)) } catch (e) {}
}

// ------------------------------------------------------------
// Fechas y formato
// ------------------------------------------------------------
function ymd(d) {
  const m = String(d.getMonth() + 1).padStart(2, "0")
  const x = String(d.getDate()).padStart(2, "0")
  return d.getFullYear() + "-" + m + "-" + x
}
function lunesDe(d) {
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate())
  const w = x.getDay()
  x.setDate(x.getDate() - (w === 0 ? 6 : w - 1))
  return x
}
function proximoLunes(d) { const l = lunesDe(d); l.setDate(l.getDate() + 7); return l }
function sep(n) { return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ",") }
function mx(n) {
  const g = n < 0, v = Math.abs(Math.round(n))
  return (g ? "−$" : "$") + sep(v)
}

// ------------------------------------------------------------
// Red
// ------------------------------------------------------------
async function llamar(cuerpo) {
  if (!ENDPOINT) throw new Error("sin endpoint")
  const r = new Request(ENDPOINT)
  r.method = "POST"
  r.headers = { "Content-Type": "text/plain;charset=utf-8" }
  r.body = JSON.stringify(cuerpo)
  r.timeoutInterval = 12
  const j = await r.loadJSON()
  if (!j || !j.ok) throw new Error((j && j.error) || "respuesta rara")
  return j
}

// Sube lo pendiente y trae lo más reciente. Nunca truena: si no hay
// internet se queda con lo que ya tenía guardado.
async function sincronizar() {
  const cola = leerJson(COLA, [])
  try {
    const j = await llamar(cola.length ? { a: "lote", ops: cola } : { a: "get" })
    escribirJson(COLA, [])
    const est = j.estado || {}
    const datos = { movimientos: j.movimientos || [], ts: Date.now(), cerebro: est.widget || null }
    escribirJson(CACHE, datos)
    return datos
  } catch (e) {
    return leerJson(CACHE, { movimientos: [], ts: 0 })
  }
}

function movimientosLocales() {
  const cache = leerJson(CACHE, { movimientos: [] })
  const cola = leerJson(COLA, [])
  const ms = cache.movimientos.slice()
  // los que todavía no suben también cuentan para el número
  for (const op of cola) if (op.a === "add" && op.mov) {
    if (!ms.some(m => m.id === op.mov.id)) ms.push(op.mov)
  }
  return ms
}

function saldoSemana() {
  const ini = lunesDe(new Date())
  const fin = new Date(ini); fin.setDate(fin.getDate() + 7)
  const a = ymd(ini), b = ymd(fin)
  const g = movimientosLocales()
    .filter(m => m.fecha >= a && m.fecha < b)
    .reduce((x, y) => x + (Number(y.monto) || 0), 0)
  return { gastado: g, saldo: semanal() - g }
}

function colorPorSaldo(s) {
  const p = s / semanal()
  if (s < 0 || p < 0.15) return COLOR.coral
  if (p < 0.40) return COLOR.brass
  return COLOR.mint
}

// ------------------------------------------------------------
// Widget
// ------------------------------------------------------------
function barraImg(pct, w, h, relleno) {
  const ctx = new DrawContext()
  ctx.size = new Size(w, h); ctx.opaque = false; ctx.respectScreenScale = true
  const r = h / 2
  ctx.setFillColor(COLOR.track)
  const p1 = new Path(); p1.addRoundedRect(new Rect(0, 0, w, h), r, r)
  ctx.addPath(p1); ctx.fillPath()
  const p = Math.max(0, Math.min(1, pct))
  if (p > 0) {
    ctx.setFillColor(relleno)
    const p2 = new Path(); p2.addRoundedRect(new Rect(0, 0, Math.max(h, p * w), h), r, r)
    ctx.addPath(p2); ctx.fillPath()
  }
  return ctx.getImage()
}

function urlCategoria(cat) {
  let nombre = "Dinero"
  try { if (Script.name()) nombre = Script.name() } catch (e) {}
  return "scriptable:///run/" + encodeURIComponent(nombre) + "?cat=" + encodeURIComponent(cat)
}

function chip(padre, c) {
  const s = padre.addStack()
  s.backgroundColor = COLOR.chip
  s.cornerRadius = 9
  s.setPadding(6, 8, 6, 8)
  s.centerAlignContent()
  s.url = urlCategoria(c.id)
  const t = s.addText(c.emoji + " " + c.nombre)
  t.font = Font.mediumSystemFont(10.5)
  t.textColor = COLOR.texto
  t.lineLimit = 1
  t.minimumScaleFactor = 0.7
}

function diasRestantes() {
  const hoy = new Date(); const w = hoy.getDay(); return w === 0 ? 1 : 8 - w
}
function textoChico(padre, t, tam, color, lineas) {
  const x = padre.addText(t)
  x.font = Font.systemFont(tam); x.textColor = color || COLOR.tenue
  x.lineLimit = lineas || 1; x.minimumScaleFactor = 0.75
  return x
}
function corto(ymdStr) {
  if (!ymdStr) return ""
  const p = ymdStr.split("-"); return (+p[2]) + " " + MESES[+p[1] - 1].slice(0, 3)
}

function crearWidget(familia) {
  const s = saldoSemana()
  const col = colorPorSaldo(s.saldo)
  const pct = Math.max(0, s.saldo) / semanal()
  const lunes = proximoLunes(new Date())
  const pendientes = leerJson(COLA, []).length
  const cb = cerebro()
  const porDia = Math.max(0, s.saldo) / diasRestantes()

  const w = new ListWidget()
  const g = new LinearGradient()
  g.colors = [COLOR.fondo1, COLOR.fondo2]; g.locations = [0, 1]
  w.backgroundGradient = g
  w.setPadding(14, 16, 14, 16)
  w.url = APP_URL || urlCategoria("")

  const chico = familia === "small"
  const ancho = chico ? 125 : 290

  const top = w.addStack(); top.layoutHorizontally(); top.bottomAlignContent()
  const n = top.addText(mx(Math.abs(s.saldo)))
  n.font = Font.boldSystemFont(chico ? 30 : 36)
  n.textColor = col; n.minimumScaleFactor = 0.5; n.lineLimit = 1
  if (!chico) {
    top.addSpacer()
    const der = top.addStack(); der.layoutVertically()
    const a1 = der.addText(mx(porDia) + "/día")
    a1.font = Font.semiboldSystemFont(13); a1.textColor = COLOR.texto; a1.rightAlignText()
    const a2 = der.addText("hasta el domingo")
    a2.font = Font.systemFont(9.5); a2.textColor = COLOR.tenue; a2.rightAlignText()
  }

  textoChico(w, s.saldo < 0 ? "te pasaste por esto" : "te quedan esta semana", 11)
  w.addSpacer(6)
  w.addImage(barraImg(pct, ancho, 8, col)).imageSize = new Size(ancho, 8)
  w.addSpacer(5)

  if (chico) {
    textoChico(w, s.saldo < 0 ? "lunes " + lunes.getDate() + " reinicia" : mx(porDia) + "/día hasta el dom", 10, COLOR.texto)
  } else {
    // lo que dice el cerebro
    let linea = ""
    if (cb && cb.aviso) linea = "✦ " + cb.aviso
    else if (cb && cb.seguro != null) linea = "✦ Hasta el " + corto(cb.hasta) + " aguantas " + mx(cb.seguro) + "/día"
    else linea = "se reinicia el lunes " + lunes.getDate() + " de " + MESES[lunes.getMonth()]
    if (pendientes) linea = pendientes + " por subir · " + linea
    const t = textoChico(w, linea, 10.5, cb && cb.aviso ? COLOR.brass : COLOR.tenue, familia === "large" ? 3 : 2)
  }

  if (familia === "large" && cb) {
    w.addSpacer(12)
    const fila = w.addStack(); fila.layoutHorizontally(); fila.spacing = 8
    const caja = (tit, val, color) => {
      const c = fila.addStack(); c.layoutVertically(); c.backgroundColor = COLOR.chip; c.cornerRadius = 10
      c.setPadding(8, 10, 8, 10); c.size = new Size(0, 0)
      const v = c.addText(val); v.font = Font.boldSystemFont(15); v.textColor = color || COLOR.texto; v.minimumScaleFactor = 0.6; v.lineLimit = 1
      const tt = c.addText(tit); tt.font = Font.systemFont(9); tt.textColor = COLOR.tenue; tt.lineLimit = 1; tt.minimumScaleFactor = 0.7
    }
    if (cb.saldo != null) caja("saldo BBVA", mx(cb.saldo))
    if (cb.seguro != null) caja("al día hasta " + corto(cb.hasta), mx(cb.seguro), cb.seguro < semanal() / 7 ? COLOR.brass : COLOR.mint)
    if (cb.bajo) caja("punto bajo " + corto(cb.bajo.fecha), mx(cb.bajo.saldo), cb.bajo.saldo < 0 ? COLOR.coral : COLOR.texto)
    const av = (cb.avisos || []).slice(1, 3)
    for (const x of av) { w.addSpacer(8); textoChico(w, "• " + x, 10.5, COLOR.texto, 2) }
    w.addSpacer()
  }

  if (!chico) {
    w.addSpacer(8)
    const fila = w.addStack(); fila.layoutHorizontally(); fila.spacing = 5
    for (const c of CATEGORIAS) chip(fila, c)
    const ia = fila.addStack()
    ia.backgroundColor = new Color("#274A78"); ia.cornerRadius = 9; ia.setPadding(6, 8, 6, 8)
    ia.url = (APP_URL || "") + "?ia=1"
    const it = ia.addText("✦"); it.font = Font.boldSystemFont(11); it.textColor = COLOR.mint
  }

  w.refreshAfterDate = new Date(Date.now() + 5 * 60 * 1000)
  return w
}

// ------------------------------------------------------------
// Registrar gasto: teclado + categorías = dos toques
// ------------------------------------------------------------
async function registrarGasto(catPre) {
  const s = saldoSemana()
  const a = new Alert()
  const c0 = CATEGORIAS.find(c => c.id === catPre)
  a.title = c0 ? c0.emoji + "  " + c0.nombre : "¿Cuánto?"
  a.message = s.saldo >= 0
    ? "Te quedan " + mx(s.saldo) + " esta semana"
    : "Vas " + mx(Math.abs(s.saldo)) + " pasado esta semana"
  const tf = a.addTextField("0", "")
  tf.setDecimalPadKeyboard()

  let cats
  if (catPre) { a.addAction("Guardar"); cats = [catPre] }
  else {
    cats = CATEGORIAS.map(c => c.id)
    for (const c of CATEGORIAS) a.addAction(c.emoji + "  " + c.nombre)
  }
  a.addCancelAction("Cancelar")

  const i = await a.presentAlert()
  if (i < 0) return false

  const monto = parseFloat((a.textFieldValue(0) || "").replace(/[^0-9.\-]/g, ""))
  if (!monto || isNaN(monto) || monto <= 0) return false

  const mov = {
    id: Date.now().toString(36) + "-" + Math.random().toString(36).slice(2, 7),
    fecha: ymd(new Date()),
    monto: Math.round(monto * 100) / 100,
    categoria: cats[i] || "otro",
    es_efectivo: false,
    nota: ""
  }

  // se guarda en la cola primero: si no hay internet, no se pierde
  const cola = leerJson(COLA, [])
  cola.push({ a: "add", mov: mov })
  escribirJson(COLA, cola)
  await sincronizar()
  return true
}

// ------------------------------------------------------------
// Pantalla corta dentro de Scriptable
// ------------------------------------------------------------
async function pantalla() {
  const datos = await sincronizar()
  const s = saldoSemana()
  const pendientes = leerJson(COLA, []).length
  const t = new UITable()
  t.showSeparators = false

  const esp = (h) => { const r = new UITableRow(); r.height = h; r.addText(" "); t.addRow(r) }
  esp(24)

  const r1 = new UITableRow(); r1.height = 86
  const c1 = r1.addText(mx(Math.abs(s.saldo)))
  c1.titleFont = Font.boldSystemFont(54); c1.centerAligned()
  c1.titleColor = colorPorSaldo(s.saldo)
  t.addRow(r1)

  const r2 = new UITableRow(); r2.height = 30
  const c2 = r2.addText(s.saldo < 0 ? "te pasaste por esto" : "te quedan esta semana")
  c2.titleFont = Font.systemFont(15); c2.titleColor = COLOR.tenue; c2.centerAligned()
  t.addRow(r2)

  esp(20)

  const rb = new UITableRow(); rb.height = 62; rb.dismissOnSelect = false
  rb.backgroundColor = new Color("#1B3252")
  const cb = rb.addText("＋   Registrar gasto")
  cb.titleFont = Font.boldSystemFont(20); cb.centerAligned()
  rb.onSelect = async () => {
    await registrarGasto(null)
    await pantalla()
  }
  t.addRow(rb)

  esp(18)

  const cbp = cerebro()
  if (cbp && (cbp.aviso || cbp.seguro != null)) {
    const rc = new UITableRow(); rc.height = 74; rc.dismissOnSelect = false
    const cc = rc.addText("✦ " + (cbp.aviso || ("Hasta el " + corto(cbp.hasta) + " aguantas " + mx(cbp.seguro) + " al día")))
    cc.titleFont = Font.systemFont(14); cc.titleColor = COLOR.brass
    t.addRow(rc)
  }

  if (APP_URL) {
    const rq = new UITableRow(); rq.height = 56
    const cq = rq.addText("✦ Pregúntale al cerebro", "dile qué gastaste o mándale tu captura de BBVA")
    cq.titleFont = Font.semiboldSystemFont(17); cq.titleColor = COLOR.mint; cq.subtitleColor = COLOR.tenue
    rq.onSelect = () => { Safari.open(APP_URL + "?ia=1") }
    t.addRow(rq)

    const ra = new UITableRow(); ra.height = 56
    const ca = ra.addText("Abrir Dinero", "el mes, los sobres y el colchón")
    ca.titleFont = Font.semiboldSystemFont(17)
    ca.subtitleColor = COLOR.tenue
    ra.onSelect = () => { Safari.open(APP_URL) }
    t.addRow(ra)
  }

  const re = new UITableRow(); re.height = 52; re.dismissOnSelect = false
  const ce = re.addText(
    pendientes ? pendientes + " movimiento(s) sin subir" : "Todo sincronizado",
    datos.ts ? "última bajada: " + new Date(datos.ts).toLocaleString() : "sin conexión todavía")
  ce.titleFont = Font.systemFont(14)
  ce.titleColor = pendientes ? COLOR.brass : COLOR.tenue
  ce.subtitleColor = COLOR.tenue
  re.onSelect = async () => { await sincronizar(); await pantalla() }
  t.addRow(re)

  await t.present()
}

// ------------------------------------------------------------
// Arranque
// ------------------------------------------------------------
if (config.runsInWidget) {
  await sincronizar()
  Script.setWidget(crearWidget(config.widgetFamily || "medium"))
  Script.complete()
} else {
  const cat = (args.queryParameters && args.queryParameters.cat) ? args.queryParameters.cat : null
  if (cat) await registrarGasto(cat)
  await pantalla()
  Script.complete()
}
