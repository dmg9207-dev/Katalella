/**
 * Accesorios Katalella — conexión entre la tienda y esta hoja de Google.
 * - La página lee aquí los productos (sin el costo, que es privado).
 * - Cada pedido llega a la pestaña "Pedidos" y se descuenta la cantidad.
 * - Avisos de cada pedido por WhatsApp (CallMeBot) y correo: se configuran en la pestaña "Ajustes".
 * - Contador de visitas por día en la pestaña "Visitas".
 */
var ENVIO_GRATIS_DESDE = 60000;
var VALOR_ENVIO = 18000;
var PESTANA_PEDIDOS = 'Pedidos';
var PESTANA_AJUSTES = 'Ajustes';
var PESTANA_VISITAS = 'Visitas';
var ENCABEZADOS_PEDIDOS = ['Fecha', 'Pedido', 'Estado', 'Cliente', 'Cédula', 'Celular', 'Departamento',
  'Ciudad', 'Dirección', 'Barrio', 'Indicaciones', 'Productos', 'Subtotal', 'Envío', 'Total',
  'Forma de pago', 'Transferir', 'Cobrar al recibir', 'Transportadora', 'Ganancia'];

function doGet(e) {
  if (e && e.parameter && e.parameter.visita) {
    try { contarVisita_(); } catch (err) {}
  }
  var productos = leerProductos_().filter(function (p) { return p.mostrar; }).map(function (p) {
    return { id: p.id, seccion: p.seccion, nombre: p.nombre, precio: p.precio, cantidad: p.cantidad,
      frase: p.frase, descripcion: p.descripcion, fotos: p.fotos };
  });
  return json_({ ok: true, productos: productos, envioGratisDesde: ENVIO_GRATIS_DESDE, valorEnvio: VALOR_ENVIO });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    var c = d.cliente || {};
    ['nombre', 'cedula', 'celular', 'depto', 'ciudad', 'direccion'].forEach(function (k) {
      if (!String(c[k] || '').trim()) throw new Error('Falta completar: ' + k);
    });
    var pago = ['Nequi', 'Bre-B', 'Contraentrega'].indexOf(d.pago) >= 0 ? d.pago : 'Nequi';
    var transportadora = ['Servientrega', 'Interrapidísimo'].indexOf(d.transportadora) >= 0 ? d.transportadora : 'Servientrega';
    if (!d.items || !d.items.length) throw new Error('El carrito está vacío.');

    var hoja = SpreadsheetApp.getActive().getSheets()[0];
    var productos = leerProductos_();
    var lineas = [], subtotal = 0, ganancia = 0, cambios = [];
    d.items.forEach(function (it) {
      var p = productos.filter(function (x) { return x.id === String(it.id); })[0];
      var qty = Math.max(1, Math.floor(Number(it.qty) || 0));
      if (!p || !p.mostrar || !p.precio) throw new Error('Un producto ya no está disponible.');
      if (p.cantidad < qty) throw new Error('Solo quedan ' + p.cantidad + ' de "' + p.nombre + '".');
      subtotal += p.precio * qty;
      ganancia += (p.precio - p.costo) * qty;
      lineas.push(qty + ' x ' + p.nombre + ' ($' + (p.precio * qty).toLocaleString('es-CO') + ')');
      cambios.push({ fila: p.fila, nueva: p.cantidad - qty });
    });

    var envio = subtotal >= ENVIO_GRATIS_DESDE ? 0 : VALOR_ENVIO;
    var total = subtotal + envio;
    var envioAparte = envio > 0 && pago !== 'Contraentrega' && !!d.envioAparte;
    var transferir = pago === 'Contraentrega' ? 0 : subtotal + (envioAparte ? 0 : envio);
    var alRecibir = pago === 'Contraentrega' ? total : (envioAparte ? envio : 0);

    var colCantidad = columna_(hoja, 'cantidad');
    cambios.forEach(function (ch) { hoja.getRange(ch.fila, colCantidad).setValue(ch.nueva); });

    var ped = hojaPedidos_();
    var numero = 'KAT-' + ('0000' + ped.getLastRow()).slice(-4);
    ped.appendRow([new Date(), numero, pago === 'Contraentrega' ? 'Contraentrega · por enviar' : 'Pendiente de pago',
      c.nombre, "'" + c.cedula, "'" + c.celular, c.depto, c.ciudad, c.direccion, c.barrio || '', c.notas || '',
      lineas.join('\n'), subtotal, envio, total, pago + (envioAparte ? ' (envío al recibir)' : ''),
      transferir, alRecibir, transportadora, ganancia]);

    avisar_(numero, c, lineas, subtotal, envio, total, pago, envioAparte, transferir, alRecibir, transportadora, ganancia);

    return json_({ ok: true, pedido: numero, subtotal: subtotal, envio: envio, total: total,
      transferir: transferir, alRecibir: alRecibir });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function leerProductos_() {
  var hoja = SpreadsheetApp.getActive().getSheets()[0];
  var filas = hoja.getDataRange().getValues();
  var cab = filas.shift().map(function (h) { return String(h).trim().toLowerCase(); });
  var i = function (n) { return cab.indexOf(n); };
  var num = function (v) { return Number(String(v).replace(/[^\d]/g, '')) || 0; };
  return filas.map(function (r, k) {
    return {
      fila: k + 2, id: String(r[i('id')]).trim(), seccion: String(r[i('seccion')]).trim(),
      nombre: String(r[i('nombre')]).trim(), precio: num(r[i('precio')]), cantidad: num(r[i('cantidad')]),
      costo: num(r[i('costo')]), mostrar: String(r[i('mostrar')]).trim().toUpperCase() !== 'NO',
      frase: String(r[i('frase')] || ''), descripcion: String(r[i('descripcion')] || ''),
      fotos: String(r[i('fotos')] || '').split(',').map(function (s) { return s.trim(); }).filter(String)
    };
  }).filter(function (p) { return p.id; });
}

function columna_(hoja, nombre) {
  var cab = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim().toLowerCase(); });
  return cab.indexOf(nombre) + 1;
}

function hojaPedidos_() {
  var ss = SpreadsheetApp.getActive();
  var ped = ss.getSheetByName(PESTANA_PEDIDOS);
  if (!ped) {
    ped = ss.insertSheet(PESTANA_PEDIDOS);
    ped.appendRow(ENCABEZADOS_PEDIDOS);
    ped.setFrozenRows(1);
    ped.getRange(1, 1, 1, ENCABEZADOS_PEDIDOS.length).setFontWeight('bold');
  }
  return ped;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function hojaAjustes_() {
  var ss = SpreadsheetApp.getActive();
  var aj = ss.getSheetByName(PESTANA_AJUSTES);
  if (!aj) {
    aj = ss.insertSheet(PESTANA_AJUSTES);
    aj.appendRow(['ajuste', 'valor', 'para qué sirve']);
    aj.appendRow(['whatsapp_avisos', "'573212972635", 'Número que recibe el aviso de cada pedido (57 + celular, sin espacios)']);
    aj.appendRow(['callmebot_apikey', '', 'La APIKey que te envía CallMeBot por WhatsApp']);
    aj.appendRow(['correo_avisos', Session.getEffectiveUser().getEmail(), 'Correo que recibe el aviso de cada pedido (déjalo vacío para no recibir correo)']);
    aj.setFrozenRows(1);
    aj.getRange(1, 1, 1, 3).setFontWeight('bold');
  }
  return aj;
}

function ajustes_() {
  var out = {};
  hojaAjustes_().getDataRange().getValues().forEach(function (r) { out[String(r[0]).trim()] = String(r[1]).trim(); });
  return out;
}

function pesos_(n) { return '$' + Number(n).toLocaleString('es-CO'); }

function avisar_(numero, c, lineas, subtotal, envio, total, pago, envioAparte, transferir, alRecibir, transportadora, ganancia) {
  var texto = '🌸 NUEVO PEDIDO ' + numero + '\n\n' +
    lineas.join('\n') + '\n\n' +
    'Productos: ' + pesos_(subtotal) + '\n' +
    'Envío: ' + (envio ? pesos_(envio) : 'Gratis') + '\n' +
    'Total: ' + pesos_(total) + '\n' +
    'Pago: ' + pago + (envioAparte ? ' (envío al recibir)' : '') + '\n' +
    (transferir ? 'Debe transferir: ' + pesos_(transferir) + '\n' : '') +
    (alRecibir ? '⚠️ Cobrar al recibir: ' + pesos_(alRecibir) + '\n' : '') +
    'Transportadora: ' + transportadora + '\n\n' +
    '📦 ENVIAR A:\n' + c.nombre + '\nC.C. ' + c.cedula + '\nCel: ' + c.celular + '\n' +
    c.direccion + (c.barrio ? ', ' + c.barrio : '') + '\n' + c.ciudad + ', ' + c.depto +
    (c.notas ? '\nIndicaciones: ' + c.notas : '') + '\n\n' +
    'Ganancia: ' + pesos_(ganancia);
  var aj = {};
  try { aj = ajustes_(); } catch (e) {}
  var tel = String(aj.whatsapp_avisos || '').replace(/[^\d]/g, '');
  if (tel && aj.callmebot_apikey) {
    try {
      UrlFetchApp.fetch('https://api.callmebot.com/whatsapp.php?phone=%2B' + tel +
        '&text=' + encodeURIComponent(texto) + '&apikey=' + encodeURIComponent(aj.callmebot_apikey),
        { muteHttpExceptions: true });
    } catch (e) {}
  }
  if (aj.correo_avisos) {
    try { MailApp.sendEmail(aj.correo_avisos, 'Nuevo pedido ' + numero + ' · Katalella', texto); } catch (e) {}
  }
}

function hojaVisitas_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(PESTANA_VISITAS);
  if (!sh) {
    sh = ss.insertSheet(PESTANA_VISITAS);
    sh.appendRow(['Fecha', 'Visitas', '', 'Total visitas']);
    sh.getRange('E1').setFormula('=SUM(B2:B)');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold');
  }
  return sh;
}

function contarVisita_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = hojaVisitas_();
    var hoy = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd');
    var ultima = sh.getLastRow();
    if (ultima > 1 && String(sh.getRange(ultima, 1).getDisplayValue()) === hoy) {
      var celda = sh.getRange(ultima, 2);
      celda.setValue((Number(celda.getValue()) || 0) + 1);
    } else {
      sh.appendRow(["'" + hoy, 1]);
    }
  } finally {
    lock.releaseLock();
  }
}

/** Ejecuta esta función para crear las pestañas "Pedidos", "Ajustes" y "Visitas" y dar los permisos. */
function prepararTienda() {
  hojaPedidos_();
  hojaAjustes_();
  hojaVisitas_();
  Logger.log('Listo. Productos encontrados: ' + leerProductos_().length);
}

/** Envía un aviso de prueba a tu WhatsApp y correo (no crea ningún pedido). */
function probarAviso() {
  avisar_('PRUEBA', { nombre: 'Clienta de prueba', cedula: '0', celular: '3000000000', depto: 'Santander',
    ciudad: 'El Carmen de Chucurí', direccion: 'Calle de prueba', barrio: '', notas: '' },
    ['1 x Producto de prueba ($10.000)'], 10000, 18000, 28000, 'Nequi', false, 28000, 0, 'Servientrega', 5000);
}
