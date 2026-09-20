/**
 * DINERO — endpoint de datos
 * Apps Script ligado a una hoja de cálculo de Google.
 *
 * Crea dos pestañas automáticamente la primera vez:
 *   movimientos  → id | fecha | monto | categoria | efectivo | nota
 *   estado       → A1 guarda un JSON con checklist, recibidos, cierres y ajustes
 *
 * Se publica como aplicación web (Implementar → Nueva implementación →
 * Aplicación web → Ejecutar como: yo → Quién tiene acceso: cualquier usuario).
 */

var HOJA_MOV = 'movimientos';
var HOJA_EST = 'estado';
var ENCABEZADOS = ['id', 'fecha', 'monto', 'categoria', 'efectivo', 'nota'];

function doGet(e) { return manejar(e); }
function doPost(e) { return manejar(e); }

function manejar(e) {
  // La app manda JSON; los atajos del iPhone mandan un formulario.
  // Se aceptan los dos: si el cuerpo no es JSON, se usan los parámetros.
  var datos = {};
  if (e && e.postData && e.postData.contents) {
    try { datos = JSON.parse(e.postData.contents); } catch (err) { datos = null; }
  }
  if (!datos || typeof datos !== 'object') datos = (e && e.parameter) ? e.parameter : {};
  if (datos.cuerpo) {
    try { datos = JSON.parse(datos.cuerpo); } catch (err) {}
  }

  var accion = datos.a || 'get';

  // La IA va fuera del candado: tarda unos segundos y no debe frenar la hoja.
  if (accion === 'ia') {
    try { return responder(preguntarIA(datos)); }
    catch (err) { return responder({ ok: false, error: String(err) }); }
  }

  var candado = LockService.getScriptLock();
  try {
    candado.waitLock(20000);
  } catch (err) {
    return responder({ ok: false, error: 'La hoja está ocupada, intenta de nuevo' });
  }

  try {
    var r;
    if (accion === 'get')          r = leerTodo();
    else if (accion === 'add')     r = agregarMovimiento(datos.mov);
    else if (accion === 'edit')    r = editarMovimiento(datos.mov);
    else if (accion === 'del')     r = borrarMovimiento(datos.id);
    else if (accion === 'estado')  r = guardarEstado(datos.estado);
    else if (accion === 'lote')    r = aplicarLote(datos.ops);
    else if (accion === 'pago')    r = desdeApplePay(datos);
    else if (accion === 'voz')     r = desdeVoz(datos);
    else                           r = { ok: false, error: 'Acción desconocida: ' + accion };
    return responder(r);
  } catch (err) {
    return responder({ ok: false, error: String(err) });
  } finally {
    candado.releaseLock();
  }
}

function responder(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- hojas ---------- */

function libro() { return SpreadsheetApp.getActiveSpreadsheet(); }

function hojaMovimientos() {
  var h = libro().getSheetByName(HOJA_MOV);
  if (!h) {
    h = libro().insertSheet(HOJA_MOV);
    h.appendRow(ENCABEZADOS);
    h.setFrozenRows(1);
  }
  return h;
}

function hojaEstado() {
  var h = libro().getSheetByName(HOJA_EST);
  if (!h) {
    h = libro().insertSheet(HOJA_EST);
    h.getRange('A1').setValue('{}');
  }
  return h;
}

/* ---------- lectura ---------- */

function leerTodo() {
  var h = hojaMovimientos();
  var filas = h.getDataRange().getValues();
  var movimientos = [];
  for (var i = 1; i < filas.length; i++) {
    var f = filas[i];
    if (!f[0]) continue;
    movimientos.push({
      id: String(f[0]),
      fecha: aFecha(f[1]),
      monto: Number(f[2]) || 0,
      categoria: String(f[3] || 'otro'),
      es_efectivo: f[4] === true || String(f[4]).toLowerCase() === 'true',
      nota: String(f[5] || '')
    });
  }

  var crudo = String(hojaEstado().getRange('A1').getValue() || '{}');
  var estado;
  try { estado = JSON.parse(crudo); } catch (err) { estado = {}; }
  if (!estado.checklist) estado.checklist = {};
  if (!estado.recibidos) estado.recibidos = {};
  if (!estado.cierres)   estado.cierres = {};
  if (!estado.ajustes)   estado.ajustes = [];

  return { ok: true, movimientos: movimientos, estado: estado, ts: Date.now() };
}

// La hoja puede devolver un objeto Date o un texto; siempre sale YYYY-MM-DD.
function aFecha(v) {
  if (v instanceof Date) {
    var y = v.getFullYear();
    var m = ('0' + (v.getMonth() + 1)).slice(-2);
    var d = ('0' + v.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
  return String(v || '').slice(0, 10);
}

/* ---------- escritura ---------- */

function agregarMovimiento(mov) {
  if (!mov || !mov.id) return { ok: false, error: 'Movimiento sin id' };
  var h = hojaMovimientos();
  if (filaDe(h, mov.id) > 0) return { ok: true, repetido: true };   // ya estaba, no duplicar
  h.appendRow([
    mov.id,
    "'" + String(mov.fecha),      // apóstrofo: la hoja lo deja como texto, no lo reinterpreta
    Number(mov.monto) || 0,
    String(mov.categoria || 'otro'),
    mov.es_efectivo === true,
    String(mov.nota || '')
  ]);
  return { ok: true };
}

function editarMovimiento(mov) {
  if (!mov || !mov.id) return { ok: false, error: 'Movimiento sin id' };
  var h = hojaMovimientos();
  var fila = filaDe(h, mov.id);
  if (fila < 1) return agregarMovimiento(mov);
  h.getRange(fila, 1, 1, ENCABEZADOS.length).setValues([[
    mov.id,
    "'" + String(mov.fecha),
    Number(mov.monto) || 0,
    String(mov.categoria || 'otro'),
    mov.es_efectivo === true,
    String(mov.nota || '')
  ]]);
  return { ok: true };
}

function borrarMovimiento(id) {
  if (!id) return { ok: false, error: 'Falta el id' };
  var h = hojaMovimientos();
  var fila = filaDe(h, String(id));
  if (fila > 0) h.deleteRow(fila);
  return { ok: true };
}

function guardarEstado(estado) {
  if (!estado) return { ok: false, error: 'Estado vacío' };
  hojaEstado().getRange('A1').setValue(JSON.stringify(estado));
  return { ok: true };
}

// Varias operaciones en una sola llamada: es lo que manda la cola de offline.
function aplicarLote(ops) {
  if (!ops || !ops.length) return leerTodo();
  for (var i = 0; i < ops.length; i++) {
    var o = ops[i];
    if (o.a === 'add')          agregarMovimiento(o.mov);
    else if (o.a === 'edit')    editarMovimiento(o.mov);
    else if (o.a === 'del')     borrarMovimiento(o.id);
    else if (o.a === 'estado')  guardarEstado(o.estado);
  }
  return leerTodo();
}

/* =====================================================================
   APPLE PAY Y VOZ
   Los atajos del iPhone mandan aquí. Toda la inteligencia vive en este
   archivo, así el atajo es de dos o tres acciones y los cambios de
   palabras clave no obligan a rehacerlo.
   ===================================================================== */

// Cada lista se revisa EN ORDEN: la primera que pegue, gana.
// Agrega aquí los comercios que se te repitan.
var PALABRAS = [
  { cat: 'gasolina', claves: ['gasolina','gasolinera','combustible','pemex','mobil',
      'shell','bp ','oxxo gas','repsol','g500','arco','cargar gas','tanque'] },
  { cat: 'comida',   claves: ['comida','comi','comer','restaurant','restaurante','fonda',
      'taco','tacos','torta','almuerzo','desayuno','cena','uber eats','ubereats',
      'rappi','didi food','sushi','pizza','pollo','super','soriana','chedraui',
      'walmart','bodega aurrera','mercado','despensa'] },
  { cat: 'antojo',   claves: ['antojo','dulce','postre','helado','cafe','café','starbucks',
      'oxxo','7 eleven','seven','chatarra','botana','snack','cerveza','chela',
      'refresco','galleta','pan','panaderia','cine'] }
];

function sinAcentos(s) {
  return String(s || '').toLowerCase()
    .replace(/á/g,'a').replace(/é/g,'e').replace(/í/g,'i')
    .replace(/ó/g,'o').replace(/ú/g,'u').replace(/ü/g,'u').replace(/ñ/g,'n');
}

// Devuelve gasolina / comida / antojo / otro
function clasificar(texto) {
  var t = sinAcentos(texto);
  if (!t) return 'otro';
  for (var i = 0; i < PALABRAS.length; i++) {
    for (var j = 0; j < PALABRAS[i].claves.length; j++) {
      if (t.indexOf(sinAcentos(PALABRAS[i].claves[j])) > -1) return PALABRAS[i].cat;
    }
  }
  return 'otro';
}

function hoyIso() {
  var z = Session.getScriptTimeZone() || 'America/Mexico_City';
  return Utilities.formatDate(new Date(), z, 'yyyy-MM-dd');
}

function aNumero(v) {
  var n = parseFloat(String(v === undefined || v === null ? '' : v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

// Saca el primer número del texto dictado: "gasté 450 de gasolina" → 450
function montoDelTexto(t) {
  var m = String(t || '').replace(/,/g, '').match(/\d+(\.\d+)?/);
  return m ? parseFloat(m[0]) : 0;
}

// Busca un movimiento de los últimos días con el mismo monto (y comercio, si
// viene). Sirve para que la clasificación por voz corrija el gasto que Apple
// Pay ya había guardado, en vez de duplicarlo.
function buscarReciente(monto, comercio, dias) {
  var h = hojaMovimientos();
  var filas = h.getDataRange().getValues();
  var limite = new Date(); limite.setDate(limite.getDate() - (dias || 3));
  var com = sinAcentos(comercio);
  for (var i = filas.length - 1; i > 0; i--) {
    var f = filas[i];
    if (!f[0]) continue;
    var fecha = aFecha(f[1]);
    if (fecha < Utilities.formatDate(limite, Session.getScriptTimeZone() || 'America/Mexico_City', 'yyyy-MM-dd')) break;
    if (Math.abs(Number(f[2]) - monto) > 0.009) continue;
    if (com && sinAcentos(f[5]).indexOf(com) < 0) continue;
    return { fila: i + 1, id: String(f[0]), fecha: fecha, nota: String(f[5] || '') };
  }
  return null;
}

function nuevoId() {
  return 'ap-' + Date.now().toString(36) + '-' + Math.floor(Math.random() * 46656).toString(36);
}

// a=pago  ·  lo dispara la automatización de Transacción al pagar con Wallet.
// Guarda el gasto de inmediato. Si el comercio se reconoce, ya cae categorizado.
function desdeApplePay(d) {
  var monto = aNumero(d.monto);
  if (monto <= 0) return { ok: false, error: 'monto vacío' };
  var comercio = String(d.comercio || '').trim();
  var mov = {
    id: nuevoId(),
    fecha: hoyIso(),
    monto: monto,
    categoria: clasificar(comercio),
    es_efectivo: false,
    nota: comercio
  };
  agregarMovimiento(mov);
  return { ok: true, categoria: mov.categoria, id: mov.id, monto: monto };
}

// a=voz  ·  lo dispara el atajo de dictado.
// Si encuentra el gasto que Apple Pay ya guardó, lo corrige. Si no, lo crea:
// así también sirve para "Oye Siri, gasté 450 en gasolina" con efectivo.
function desdeVoz(d) {
  var texto = String(d.texto || '').trim();
  var comercio = String(d.comercio || '').trim();
  var monto = aNumero(d.monto) || montoDelTexto(texto);
  if (monto <= 0) return { ok: false, error: 'no entendí el monto' };

  var cat = clasificar(texto);
  if (cat === 'otro' && comercio) cat = clasificar(comercio);
  var efectivo = sinAcentos(texto).indexOf('efectivo') > -1;

  var previo = buscarReciente(monto, comercio, 3);
  var nota = texto || comercio;
  if (previo && comercio && nota.indexOf(comercio) < 0) nota = comercio + ' · ' + texto;

  if (previo) {
    editarMovimiento({ id: previo.id, fecha: previo.fecha, monto: monto,
                       categoria: cat, es_efectivo: efectivo, nota: nota });
    return { ok: true, categoria: cat, corregido: true, monto: monto };
  }
  var mov = { id: nuevoId(), fecha: hoyIso(), monto: monto, categoria: cat,
              es_efectivo: efectivo, nota: nota };
  agregarMovimiento(mov);
  return { ok: true, categoria: cat, corregido: false, monto: monto };
}

function filaDe(hoja, id) {
  var col = hoja.getRange(1, 1, Math.max(hoja.getLastRow(), 1), 1).getValues();
  for (var i = 1; i < col.length; i++) {
    if (String(col[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

/* =====================================================================
   ✦ IA — el cerebro de la app le pregunta a Claude por aquí.
   Tu llave vive en: Configuración del proyecto → Propiedades de secuencia
   de comandos → ANTHROPIC_API_KEY. Nunca sale de Google.
   Opcional: MODELO_IA (por defecto claude-sonnet-5; más barato:
   claude-haiku-4-5-20251001).
   ===================================================================== */
var MODELO_POR_DEFECTO = 'claude-sonnet-5';

function preguntarIA(d) {
  var props = PropertiesService.getScriptProperties();
  var llave = props.getProperty('ANTHROPIC_API_KEY');
  if (!llave) return { ok: false, error: 'Falta la llave ANTHROPIC_API_KEY en las propiedades del Apps Script' };
  var modelo = props.getProperty('MODELO_IA') || MODELO_POR_DEFECTO;
  if (!d.mensajes || !d.mensajes.length) return { ok: false, error: 'No hay mensaje' };

  var cuerpo = {
    model: modelo,
    max_tokens: 2000,
    system: String(d.sistema || '').slice(0, 20000),
    messages: d.mensajes.slice(-12)
  };
  if (d.herramienta && d.herramienta.name) {
    cuerpo.tools = [d.herramienta];
    cuerpo.tool_choice = { type: 'tool', name: d.herramienta.name };
  }

  var r = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': llave, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(cuerpo),
    muteHttpExceptions: true
  });
  var codigo = r.getResponseCode();
  var j;
  try { j = JSON.parse(r.getContentText()); } catch (err) { j = null; }
  if (codigo !== 200 || !j) {
    var msg = (j && j.error && j.error.message) ? j.error.message : ('la IA contestó ' + codigo);
    if (codigo === 401) msg = 'La llave de la IA no es válida (revisa ANTHROPIC_API_KEY)';
    if (codigo === 400 && /credit/i.test(msg)) msg = 'Tu cuenta de la API no tiene saldo. Recarga en console.anthropic.com → Billing';
    return { ok: false, error: msg };
  }

  var salida = null, texto = '';
  (j.content || []).forEach(function (b) {
    if (b.type === 'tool_use') salida = b.input;
    if (b.type === 'text') texto += b.text;
  });
  if (!salida) salida = { respuesta: texto || '(sin respuesta)', acciones: [] };
  return { ok: true, r: salida, uso: j.usage, modelo: modelo };
}

// Córrela UNA vez desde el editor para dar permiso de conectarse a la IA.
function probarIA() {
  var r = preguntarIA({
    sistema: 'Contesta en una línea.',
    mensajes: [{ role: 'user', content: 'Di: «El cerebro de Dinero está conectado».' }]
  });
  Logger.log(JSON.stringify(r));
}

/* ---------- prueba manual desde el editor ---------- */
function probar() {
  Logger.log(JSON.stringify(leerTodo()));
}
