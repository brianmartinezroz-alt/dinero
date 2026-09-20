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
   ✦ IA — el cerebro de la app le pregunta a una IA por aquí.
   Gratis: Google Gemini. Pon tu llave en
   Configuración del proyecto → Propiedades de secuencia de comandos:
       GEMINI_API_KEY = (tu llave de aistudio.google.com)
   Opcional: MODELO_IA (por ejemplo gemini-2.5-flash).
   Si algún día pones ANTHROPIC_API_KEY, usa Claude en su lugar (de paga).
   ===================================================================== */
var MODELOS_GEMINI = ['gemini-3.5-flash', 'gemini-flash-latest', 'gemini-2.5-flash', 'gemini-2.5-flash-lite'];
var MODELO_CLAUDE = 'claude-sonnet-5';

function preguntarIA(d) {
  var props = PropertiesService.getScriptProperties();
  if (!d.mensajes || !d.mensajes.length) return { ok: false, error: 'No hay mensaje' };
  if (props.getProperty('ANTHROPIC_API_KEY')) return preguntarClaude(d, props);
  if (props.getProperty('GEMINI_API_KEY')) return preguntarGemini(d, props);
  return { ok: false, error: 'Falta la llave GEMINI_API_KEY en las propiedades del Apps Script' };
}

/* ---------- Gemini (gratis) ---------- */
// Convierte el esquema de la herramienta al formato que pide Gemini.
function esquemaGemini(x) {
  if (!x || typeof x !== 'object') return x;
  var o = {};
  if (x.type) o.type = String(x.type).toUpperCase();
  if (x.description) o.description = x.description;
  if (x.enum) o.enum = x.enum;
  if (x.required) o.required = x.required;
  if (x.items) o.items = esquemaGemini(x.items);
  if (x.properties) { o.properties = {}; for (var k in x.properties) o.properties[k] = esquemaGemini(x.properties[k]); }
  return o;
}

function contenidosGemini(mensajes) {
  return mensajes.slice(-12).map(function (m) {
    var partes = [];
    if (typeof m.content === 'string') partes.push({ text: m.content });
    else (m.content || []).forEach(function (b) {
      if (b.type === 'text') partes.push({ text: b.text });
      if (b.type === 'image' && b.source) partes.push({ inlineData: { mimeType: b.source.media_type, data: b.source.data } });
    });
    return { role: m.role === 'assistant' ? 'model' : 'user', parts: partes };
  });
}

function llamarGemini(modelo, llave, cuerpo) {
  var r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + modelo + ':generateContent', {
    method: 'post', contentType: 'application/json',
    headers: { 'x-goog-api-key': llave },
    payload: JSON.stringify(cuerpo), muteHttpExceptions: true
  });
  var j = null; try { j = JSON.parse(r.getContentText()); } catch (e) {}
  return { codigo: r.getResponseCode(), j: j };
}

function preguntarGemini(d, props) {
  var llave = props.getProperty('GEMINI_API_KEY');
  var elegido = props.getProperty('MODELO_IA');
  var modelos = elegido ? [elegido].concat(MODELOS_GEMINI) : MODELOS_GEMINI.slice();
  var sistema = String(d.sistema || '').slice(0, 20000);
  if (d.herramienta) sistema += '\n\nFORMATO: contesta SOLO con un objeto JSON {"respuesta": "...", "acciones": [...]} según el esquema.';
  var base = {
    systemInstruction: { parts: [{ text: sistema }] },
    contents: contenidosGemini(d.mensajes),
    generationConfig: { maxOutputTokens: 2500, temperature: 0.4, responseMimeType: 'application/json' }
  };
  var conEsquema = JSON.parse(JSON.stringify(base));
  if (d.herramienta && d.herramienta.input_schema) conEsquema.generationConfig.responseSchema = esquemaGemini(d.herramienta.input_schema);

  var ultimo = '';
  for (var i = 0; i < modelos.length; i++) {
    var m = modelos[i];
    var r = llamarGemini(m, llave, conEsquema);
    if (r.codigo === 400) r = llamarGemini(m, llave, base);          // por si no le gustó el esquema
    if (r.codigo === 404) { ultimo = 'modelo ' + m + ' no existe'; continue; }  // prueba el siguiente
    if (r.codigo === 429) return { ok: false, error: 'Llegaste al límite gratis de la IA por ahora. Intenta en un rato.' };
    if (r.codigo === 400 || r.codigo === 403) {
      var em = (r.j && r.j.error && r.j.error.message) || ('la IA contestó ' + r.codigo);
      if (/API key/i.test(em)) em = 'La llave de Gemini no es válida (revisa GEMINI_API_KEY)';
      return { ok: false, error: em };
    }
    if (r.codigo !== 200 || !r.j) { ultimo = 'la IA contestó ' + r.codigo; continue; }

    var c = (r.j.candidates || [])[0];
    var texto = '';
    ((c && c.content && c.content.parts) || []).forEach(function (p) { if (p.text && !p.thought) texto += p.text; });
    var salida = null;
    try { salida = JSON.parse(texto.replace(/^```(json)?/i, '').replace(/```\s*$/, '')); } catch (e) {}
    if (!salida || typeof salida !== 'object') salida = { respuesta: texto || '(sin respuesta)', acciones: [] };
    if (!Array.isArray(salida.acciones)) salida.acciones = [];
    var u = r.j.usageMetadata || {};
    return { ok: true, r: salida, modelo: m,
             uso: { input_tokens: u.promptTokenCount || 0, output_tokens: u.candidatesTokenCount || 0 } };
  }
  return { ok: false, error: 'No pude usar ningún modelo de Gemini (' + ultimo + ')' };
}

/* ---------- Claude (de paga, opcional) ---------- */
function preguntarClaude(d, props) {
  var llave = props.getProperty('ANTHROPIC_API_KEY');
  var modelo = props.getProperty('MODELO_IA') || MODELO_CLAUDE;
  var cuerpo = { model: modelo, max_tokens: 2000, system: String(d.sistema || '').slice(0, 20000), messages: d.mensajes.slice(-12) };
  if (d.herramienta && d.herramienta.name) {
    cuerpo.tools = [d.herramienta];
    cuerpo.tool_choice = { type: 'tool', name: d.herramienta.name };
  }
  var r = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json',
    headers: { 'x-api-key': llave, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(cuerpo), muteHttpExceptions: true
  });
  var codigo = r.getResponseCode(), j = null;
  try { j = JSON.parse(r.getContentText()); } catch (err) {}
  if (codigo !== 200 || !j) {
    var msg = (j && j.error && j.error.message) ? j.error.message : ('la IA contestó ' + codigo);
    if (codigo === 401) msg = 'La llave de Claude no es válida (revisa ANTHROPIC_API_KEY)';
    return { ok: false, error: msg };
  }
  var salida = null, texto = '';
  (j.content || []).forEach(function (b) { if (b.type === 'tool_use') salida = b.input; if (b.type === 'text') texto += b.text; });
  if (!salida) salida = { respuesta: texto || '(sin respuesta)', acciones: [] };
  return { ok: true, r: salida, uso: j.usage, modelo: modelo };
}

// Córrela UNA vez desde el editor para dar permiso de conectarse a la IA.
function probarIA() {
  var r = preguntarIA({
    sistema: 'Contesta con JSON {"respuesta": "...", "acciones": []}.',
    mensajes: [{ role: 'user', content: 'Pon en respuesta: «El cerebro de Dinero está conectado».' }]
  });
  Logger.log(JSON.stringify(r));
}

/* ---------- prueba manual desde el editor ---------- */
function probar() {
  Logger.log(JSON.stringify(leerTodo()));
}
