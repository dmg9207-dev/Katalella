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
      frase: p.frase, descripcion: p.descripcion, fotos: p.fotos, tipo: p.tipo, material: p.material };
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

    var local = esLocal_(c.depto, c.ciudad);
    if (local) transportadora = 'Entrega local';
    var envio = (local || subtotal >= ENVIO_GRATIS_DESDE) ? 0 : VALOR_ENVIO;
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

/** Envío gratis para El Carmen de Chucurí, Santander. */
function esLocal_(depto, ciudad) {
  var n = function (t) { return String(t || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase(); };
  return n(depto).indexOf('santander') === 0 && n(depto).indexOf('norte') < 0 && n(ciudad).indexOf('carmen') >= 0;
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
      fotos: String(r[i('fotos')] || '').split(',').map(function (s) { return s.trim(); }).filter(String),
      tipo: i('tipo') >= 0 ? String(r[i('tipo')] || '').trim() : '',
      material: i('material') >= 0 ? String(r[i('material')] || '').trim() : ''
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

/* ===================== Actualización de catálogo (noviembre) =====================
 * Ejecuta "actualizarCatalogo" UNA vez. Es segura: si la ejecutas de nuevo no repite el aumento
 * ni duplica productos.
 *  - NO cambia los precios de los productos que ya estaban.
 *  - Agrega las columnas "material" y "tipo".
 *  - Agrega los productos de oro laminado (pestaña "Oro laminado" en la tienda).
 */
var RODIO = ['n095','n096','n097','n098','n099','n100','n124','n125','n126'];
var ORO = [
  ["o001", "Oro laminado", "Topos", "Topos tortuga con circón", 65000, 55000, "Paso a paso también se llega lejos.", "Tortuguitas con un circón en forma de gota como caparazón: tiernas y brillantes.", "o001-topos-tortuga-con-circon.jpg"],
  ["o002", "Oro laminado", "Topos", "Topos sol diamantado", 35000, 25000, "Sé el sol de tu propio día.", "Discos con corte diamantado que reflejan la luz como un pequeño sol.", "o002-topos-sol-diamantado.jpg"],
  ["o003", "Oro laminado", "Topos", "Topos punto de luz", 38000, 28000, "La sencillez también deslumbra.", "Topo minimalista con un circón engastado: el básico que combina con todo.", "o003-topos-punto-de-luz.jpg"],
  ["o004", "Oro laminado", "Topos", "Topos gota con halo de circones", 55000, 45000, "Tu luz no necesita permiso para brillar.", "Gota central rodeada de pequeños circones, ideal para ocasiones especiales.", "o004-topos-gota-con-halo-de-circones.jpg"],
  ["o005", "Oro laminado", "Topos", "Topos concha", 35000, 25000, "Lleva contigo la calma del mar.", "Conchita dorada con relieve, un guiño al mar para todos los días.", "o005-topos-concha.jpg"],
  ["o006", "Oro laminado", "Topos", "Topos mariposa con circones", 40000, 30000, "Las mariposas no miran atrás: vuelan.", "Mariposas con alas de circones que capturan cada destello.", "o006-topos-mariposa-con-circones.jpg"],
  ["o007", "Oro laminado", "Topos", "Topos flor esmeralda", 42000, 32000, "Crece donde te siembren.", "Florecita dorada con centro verde esmeralda, fresca y delicada.", "o007-topos-flor-esmeralda.jpg"],
  ["o008", "Oro laminado", "Topos", "Topos grano de café", 55000, 45000, "Despierta cada día con ganas de lograrlo.", "Grano de café dorado con textura, perfecto para las amantes del café colombiano.", "o008-topos-grano-de-cafe.jpg"],
  ["o009", "Oro laminado", "Topos", "Topos Mickey con circones", 52000, 42000, "Nunca pierdas tu niña interior.", "Silueta de Mickey cubierta de circones, divertida y brillante.", "o009-topos-mickey-con-circones.jpg"],
  ["o010", "Oro laminado", "Topos", "Topos esfera diamantada", 49000, 39000, "Gira el mundo, pero tú sigues firme.", "Esferas doradas con corte diamantado que brillan desde todos los ángulos.", "o010-topos-esfera-diamantada.jpg"],
  ["o011", "Oro laminado", "Topos", "Topos rosa dorada", 45000, 35000, "Florece aunque nadie te riegue.", "Rosas doradas en relieve, románticas y elegantes.", "o011-topos-rosa-dorada.jpg"],
  ["o012", "Oro laminado", "Topos", "Topos esfera grande texturizada", 70000, 60000, "Hazte espacio, tu brillo lo merece.", "Esfera grande con textura martillada, protagonista en cualquier look.", "o012-topos-esfera-grande-texturizada.jpg"],
  ["o013", "Oro laminado", "Topos", "Topos Virgen Milagrosa", 40000, 30000, "Bajo su manto, siempre segura.", "Medalla ovalada de la Virgen Milagrosa en relieve, símbolo de protección.", "o013-topos-virgen-milagrosa.jpg"],
  ["o014", "Oro laminado", "Topos", "Topos flor, estrella y aro", 40000, 30000, "Combina tus sueños y créalos.", "Diseños combinables: flor de circón rosado, estrella y aro con brillo. Precio por unidad.", "o014-topos-flor-estrella-y-aro.jpg"],
  ["o015", "Oro laminado", "Topos", "Topos circón verde", 30000, 20000, "La esperanza es el color de tu mirada.", "Circón verde redondo engastado, un toque de color elegante.", "o015-topos-circon-verde.jpg"],
  ["o016", "Oro laminado", "Topos", "Aretes gota colgante con circón", 70000, 60000, "Camina como si el mundo fuera tuyo.", "Aretes largos de doble gota calada con circones, ligeros y sofisticados.", "o016-aretes-gota-colgante-con-circon.jpg"],
  ["o017", "Oro laminado", "Topos", "Topos huellita de colores", 40000, 30000, "El amor de cuatro patas también deja huella.", "Huellitas con circones de colores para las amantes de las mascotas. Precio por unidad.", "o017-topos-huellita-de-colores.jpg"],
  ["o018", "Oro laminado", "Topos", "Topos osito", 35000, 25000, "Abrázate fuerte, te lo mereces.", "Ositos dorados de líneas suaves, tiernos y fáciles de combinar.", "o018-topos-osito.jpg"],
  ["o019", "Oro laminado", "Manillas tejidas", "Manilla tejida con balines", 52000, 42000, "Cada nudo guarda un deseo.", "Manilla en hilo tejido con balines dorados. Disponible en rojo, negro o dorado. Precio por unidad.", "o019-manilla-tejida-con-balines.jpg"],
  ["o020", "Oro laminado", "Manillas tejidas", "Manilla tejida perlas y Virgen", 90000, 80000, "La fe y la elegancia van de la mano.", "Tejido dorado con perlas y medalla de la Virgen con circones.", "o020-manilla-tejida-perlas-y-virgen.jpg"],
  ["o021", "Oro laminado", "Manillas tejidas", "Manilla tejida con medalla de santo", 60000, 50000, "Protegida en cada paso que das.", "Manilla tejida con balines y medalla de San Benito, Milagrosa o Guadalupe. Precio por unidad.", "o021-manilla-tejida-con-medalla-de-santo.jpg"],
  ["o022", "Oro laminado", "Manillas tejidas", "Manilla tejida roja San Benito", 85000, 75000, "Que nada ni nadie apague tu paz.", "Hilo rojo tejido con balines y medalla de San Benito para la protección.", "o022-manilla-tejida-roja-san-benito.jpg"],
  ["o023", "Oro laminado", "Manillas tejidas", "Manilla tejida roja infinito", 60000, 50000, "Lo que es para ti, es para siempre.", "Hilo rojo con dije de infinito dorado: amor y protección sin fin.", "o023-manilla-tejida-roja-infinito.jpg"],
  ["o024", "Oro laminado", "Manillas tejidas", "Manilla tejida cruz", 100000, 90000, "Con fe, cada camino se ilumina.", "Tejido beige con balines y cruz calada en el centro.", "o024-manilla-tejida-cruz.jpg"],
  ["o025", "Oro laminado", "Manillas tejidas", "Manilla tejida negra balines", 67000, 57000, "Lo clásico nunca falla.", "Hilo negro tejido con una hilera de balines dorados, elegante y moderna.", "o025-manilla-tejida-negra-balines.jpg"],
  ["o026", "Oro laminado", "Manillas tejidas", "Manilla tejida negra fina", 70000, 60000, "Menos ruido, más brillo.", "Manilla negra delgada con balines dorados, ideal para combinar.", "o026-manilla-tejida-negra-fina.jpg"],
  ["o027", "Oro laminado", "Manillas tejidas", "Manilla tejida roja Virgen", 82000, 72000, "Llevas un ángel contigo.", "Hilo rojo con balines y medalla de la Virgen al centro.", "o027-manilla-tejida-roja-virgen.jpg"],
  ["o028", "Oro laminado", "Manillas tejidas", "Manilla tejida azul", 70000, 60000, "Que tu calma sea tu fuerza.", "Bolitas tejidas en azul con balines dorados, fresca y juvenil.", "o028-manilla-tejida-azul.jpg"],
  ["o029", "Oro laminado", "Manillas tejidas", "Manilla tejida roja de bolas", 86000, 76000, "Tu energía es contagiosa.", "Bolas tejidas en rojo con esferas doradas, llamativa y alegre.", "o029-manilla-tejida-roja-de-bolas.jpg"],
  ["o030", "Oro laminado", "Manillas tejidas", "Manilla tejida vinotinto", 70000, 60000, "Elegancia es seguridad en ti misma.", "Bolas tejidas vinotinto con esferas doradas, elegante para toda ocasión.", "o030-manilla-tejida-vinotinto.jpg"],
  ["o031", "Oro laminado", "Candongas", "Candongas con circones frontales", 66000, 56000, "Pequeños brillos, grandes días.", "Candongas pequeñas con circones al frente, brillo delicado para el día a día.", "o031-candongas-con-circones-frontales.jpg"],
  ["o032", "Oro laminado", "Candongas", "Candongas torcidas pequeñas", 60000, 50000, "Las curvas de la vida también son hermosas.", "Candongas con diseño torcido que juega con la luz.", "o032-candongas-torcidas-pequenas.jpg"],
  ["o033", "Oro laminado", "Candongas", "Candongas pavé de circones", 86000, 76000, "Mereces brillar todos los días.", "Candonga cubierta de circones en pavé, lujo en pequeño formato.", "o033-candongas-pave-de-circones.jpg"],
  ["o034", "Oro laminado", "Candongas", "Candongas aro grande liso", 130000, 120000, "Grandes sueños, grandes aros.", "Aro grande y liso, un clásico atemporal que alarga el rostro.", "o034-candongas-aro-grande-liso.jpg"],
  ["o035", "Oro laminado", "Candongas", "Candongas gruesas medianas", 96000, 86000, "Fuerte, segura y radiante.", "Candongas gruesas de tamaño mediano, cómodas y con mucho brillo.", "o035-candongas-gruesas-medianas.jpg"],
  ["o036", "Oro laminado", "Candongas", "Candongas grandes diamantadas", 93000, 83000, "Muévete y deja brillo a tu paso.", "Aros grandes y finos con corte diamantado que destellan al moverte.", "o036-candongas-grandes-diamantadas.jpg"],
  ["o037", "Oro laminado", "Candongas", "Candongas torcidas medianas", 85000, 75000, "Gira hacia lo que te hace feliz.", "Candongas medianas con giro elegante, ideales para el día y la noche.", "o037-candongas-torcidas-medianas.jpg"],
  ["o038", "Oro laminado", "Candongas", "Candongas tubo gruesas", 105000, 95000, "Atrévete a ser tú.", "Candonga gruesa tipo tubo, la favorita para un look atrevido.", "o038-candongas-tubo-gruesas.jpg"],
  ["o039", "Oro laminado", "Candongas", "Candongas gruesas pequeñas", 65000, 55000, "Lo pequeño también tiene poder.", "Candongas pequeñas y gruesas que se ven en todo momento.", "o039-candongas-gruesas-pequenas.jpg"],
  ["o040", "Oro laminado", "Candongas", "Candongas finas lisas", 52000, 42000, "La delicadeza es tu fortaleza.", "Aro fino y liso, delicado y muy cómodo para usar a diario.", "o040-candongas-finas-lisas.jpg"],
  ["o041", "Oro laminado", "Candongas", "Candongas lisas medianas", 80000, 70000, "Un básico para una mujer extraordinaria.", "Candongas lisas de tamaño medio, el básico que no puede faltar.", "o041-candongas-lisas-medianas.jpg"],
  ["o042", "Oro laminado", "Candongas", "Candongas dobles", 65000, 55000, "Doble brillo, doble actitud.", "Candongas de doble línea que dan volumen sin peso.", "o042-candongas-dobles.jpg"],
  ["o043", "Oro laminado", "Candongas", "Candongas gruesas grandes", 120000, 110000, "Tu presencia se siente.", "Candongas gruesas de tamaño grande, para brillar con fuerza.", "o043-candongas-gruesas-grandes.jpg"],
  ["o044", "Oro laminado", "Candongas", "Candongas lisas grandes o medianas", 55000, 45000, "Elige tu tamaño, elige tu estilo.", "Candongas lisas en tamaño grande o mediano. Indica el tamaño en \"Indicaciones\".", "o044-candongas-lisas-grandes-o-medianas.jpg"],
  ["o045", "Oro laminado", "Candongas", "Candongas lisas pequeñas", 40000, 30000, "Discreta, pero inolvidable.", "Candongas lisas en tamaño pequeño, perfectas para segunda perforación.", "o045-candongas-lisas-pequenas.jpg"],
  ["o046", "Oro laminado", "Candongas", "Candongas con circones grandes", 60000, 50000, "Brilla en grande.", "Candongas con una línea de circones, tamaño grande.", "o046-candongas-con-circones-grandes.jpg"],
  ["o047", "Oro laminado", "Candongas", "Candongas con circones medianas", 55000, 45000, "El equilibrio perfecto entre brillo y estilo.", "Candongas con una línea de circones, tamaño mediano.", "o047-candongas-con-circones-medianas.jpg"],
  ["o048", "Oro laminado", "Candongas", "Candongas con circones pequeñas", 50000, 40000, "Un destello para cada día.", "Candongas con una línea de circones, tamaño pequeño.", "o048-candongas-con-circones-pequenas.jpg"],
  ["o049", "Oro laminado", "Pulseras", "Esclava placa con cadena", 107000, 87000, "Escribe tu propia historia.", "Pulsera esclava con placa lisa, ideal para grabar un nombre o fecha. Precio por unidad.", "o049-esclava-placa-con-cadena.jpg"],
  ["o050", "Oro laminado", "Pulseras", "Esclava placa lisa", 90000, 70000, "Tu nombre, tu sello, tu brillo.", "Placa rectangular lisa en cadena fina, elegante y personalizable.", "o050-esclava-placa-lisa.jpg"],
  ["o051", "Oro laminado", "Pulseras", "Esclava placa gruesa", 120000, 100000, "Deja tu marca en el mundo.", "Placa amplia con cadena de eslabones, presencia y estilo.", "o051-esclava-placa-gruesa.jpg"],
  ["o052", "Oro laminado", "Pulseras", "Pulsera tejido palomita", 90000, 70000, "Brillar es tu estado natural.", "Cadena tejida tipo palomita con mucho volumen y brillo. Precio por unidad.", "o052-pulsera-tejido-palomita.jpg"],
  ["o053", "Oro laminado", "Pulseras", "Pulsera eslabón grano de café", 153000, 133000, "Cada eslabón, una meta cumplida.", "Eslabones tipo grano de café, un diseño clásico y elegante.", "o053-pulsera-eslabon-grano-de-cafe.jpg"],
  ["o054", "Oro laminado", "Pulseras", "Pulsera barbada plana", 160000, 140000, "Fuerte como el oro, suave como tú.", "Cadena barbada plana y ancha, un clásico de lujo.", "o054-pulsera-barbada-plana.jpg"],
  ["o055", "Oro laminado", "Pulseras", "Pulsera trébol negro", 135000, 115000, "Cuatro hojas de buena suerte para ti.", "Cadena fina con tréboles negros de borde dorado, elegancia pura.", "o055-pulsera-trebol-negro.jpg"],
  ["o056", "Oro laminado", "Pulseras", "Pulsera denario Virgen y cruz", 90000, 70000, "Reza, confía y avanza.", "Pulsera tipo denario con medalla de la Virgen y cruz.", "o056-pulsera-denario-virgen-y-cruz.jpg"],
  ["o057", "Oro laminado", "Pulseras", "Pulsera eslabón café fina", 110000, 90000, "Lo delicado también es firme.", "Eslabones tipo grano de café en versión delicada.", "o057-pulsera-eslabon-cafe-fina.jpg"],
  ["o058", "Oro laminado", "Tobilleras", "Tobillera trébol con circones", 94000, 79000, "Camina segura, la suerte te sigue.", "Cadena fina con trébol de circones para lucir en el tobillo.", "o058-tobillera-trebol-con-circones.jpg"],
  ["o059", "Oro laminado", "Tobilleras", "Tobillera de bolitas", 75000, 60000, "Pisa fuerte y con estilo.", "Cadena con bolitas doradas, perfecta para el sol y la playa.", "o059-tobillera-de-bolitas.jpg"],
  ["o060", "Oro laminado", "Tobilleras", "Tobillera de barritas", 80000, 65000, "Cada paso cuenta tu historia.", "Cadena con barritas brillantes, discreta y elegante.", "o060-tobillera-de-barritas.jpg"],
  ["o061", "Oro laminado", "Cadenas", "Cadena veneciana fina 45 cm", 95000, 75000, "Lleva cerca lo que más amas.", "Cadena veneciana fina de 45 cm, ideal para dijes.", "o061-cadena-veneciana-fina-45-cm.jpg"],
  ["o062", "Oro laminado", "Cadenas", "Cadena gruesa 45 cm", 170000, 150000, "Tu fuerza se nota.", "Cadena de eslabón grueso de 45 cm con mucha presencia.", "o062-cadena-gruesa-45-cm.jpg"],
  ["o063", "Oro laminado", "Cadenas", "Cadena brillante 50 cm", 105000, 85000, "Brilla en cada detalle.", "Cadena de eslabones brillantes de 50 cm.", "o063-cadena-brillante-50-cm.jpg"],
  ["o064", "Oro laminado", "Cadenas", "Cadena de eslabones 60 cm", 134000, 114000, "Más largo el camino, más bonito el logro.", "Cadena de eslabón plano de 60 cm, ideal para dijes grandes.", "o064-cadena-de-eslabones-60-cm.jpg"],
  ["o065", "Oro laminado", "Cadenas", "Cadena destellos fina", 95000, 75000, "La luz siempre encuentra su camino.", "Cadena fina de eslabones con destellos que brillan con la luz.", "o065-cadena-destellos-fina.jpg"],
  ["o066", "Oro laminado", "Cadenas", "Cadena 45 cm tejido fino", 140000, 120000, "La elegancia está en los detalles.", "Cadena de tejido fino de 45 cm, elegante y resistente.", "o066-cadena-45-cm-tejido-fino.jpg"],
  ["o067", "Oro laminado", "Cadenas", "Cadena de bolitas 45 cm", 120000, 100000, "Cada bolita, un motivo para sonreír.", "Cadena con bolitas doradas de 45 cm, delicada y moderna.", "o067-cadena-de-bolitas-45-cm.jpg"],
  ["o068", "Oro laminado", "Cadenas", "Cadena torzal 50 cm", 120000, 100000, "Entrelaza tus sueños con tus actos.", "Cadena tipo torzal de 50 cm con brillo intenso.", "o068-cadena-torzal-50-cm.jpg"],
  ["o069", "Oro laminado", "Cadenas", "Cadena caja 50 cm", 115000, 95000, "Lo clásico es para siempre.", "Cadena tipo caja de 50 cm, firme y clásica.", "o069-cadena-caja-50-cm.jpg"],
  ["o070", "Oro laminado", "Cadenas", "Cadena cubana gruesa", 243000, 223000, "Lujo es sentirte bien contigo.", "Cadena cubana gruesa de alto brillo, presencia y estilo.", "o070-cadena-cubana-gruesa.jpg"],
  ["o071", "Oro laminado", "Cadenas", "Cadena destellos 60 cm", 302000, 282000, "Tu luz llega lejos.", "Cadena de eslabones con destellos de 60 cm, larga y elegante.", "o071-cadena-destellos-60-cm.jpg"],
  ["o072", "Oro laminado", "Cadenas", "Cadena con barras", 150000, 130000, "Construye tu camino paso a paso.", "Cadena fina con barritas intercaladas, moderna y delicada.", "o072-cadena-con-barras.jpg"],
  ["o073", "Oro laminado", "Cadenas", "Cadena trébol negro", 156000, 136000, "La suerte se lleva cerca del corazón.", "Cadena fina con tréboles negros de borde dorado.", "o073-cadena-trebol-negro.jpg"],
  ["o074", "Oro laminado", "Cadenas", "Cadena barbada", 171000, 151000, "Lo valioso se queda contigo.", "Cadena barbada plana, un clásico que nunca pasa de moda.", "o074-cadena-barbada.jpg"],
  ["o075", "Oro laminado", "Cadenas", "Cadena plana fina", 126000, 106000, "La belleza está en la sencillez.", "Cadena plana fina con brillo, ideal para usar sola o con dije.", "o075-cadena-plana-fina.jpg"],
  ["o076", "Oro laminado", "Dijes", "Dije medalla Virgen Milagrosa", 88000, 78000, "Bendecida y en paz.", "Medalla ovalada de la Virgen Milagrosa con detalle en relieve.", "o076-dije-medalla-virgen-milagrosa.jpg"],
  ["o077", "Oro laminado", "Dijes", "Dije cruz texturizada", 43000, 33000, "La fe es tu ancla.", "Cruz dorada con textura y borde liso.", "o077-dije-cruz-texturizada.jpg"],
  ["o078", "Oro laminado", "Dijes", "Dije crucifijo", 78000, 68000, "Llévalo cerca del corazón.", "Crucifijo dorado con detalle en relieve.", "o078-dije-crucifijo.jpg"],
  ["o079", "Oro laminado", "Dijes", "Dije llavecita con circones", 35000, 25000, "Tú tienes la llave de tus sueños.", "Llavecita con circones, símbolo de nuevas puertas. Precio por unidad.", "o079-dije-llavecita-con-circones.jpg"],
  ["o080", "Oro laminado", "Dijes", "Dije bolsita de la abundancia", 46000, 36000, "La abundancia llega a quien la espera con fe.", "Bolsita de la abundancia dorada para atraer prosperidad.", "o080-dije-bolsita-de-la-abundancia.jpg"],
  ["o081", "Oro laminado", "Dijes", "Dije árbol de la vida", 48000, 38000, "Echa raíces y crece hacia el cielo.", "Árbol de la vida dentro de un aro de circones.", "o081-dije-arbol-de-la-vida.jpg"]
];

function actualizarCatalogo() {
  var hoja = SpreadsheetApp.getActive().getSheets()[0];
  ['material', 'tipo'].forEach(function (col) {
    if (columna_(hoja, col) === 0) hoja.getRange(1, hoja.getLastColumn() + 1).setValue(col);
  });
  var cId = columna_(hoja, 'id'),
      cMat = columna_(hoja, 'material'), cTipo = columna_(hoja, 'tipo');
  var datos = hoja.getDataRange().getValues();
  var ids = {};
  for (var r = 1; r < datos.length; r++) {
    var id = String(datos[r][cId - 1]).trim(); if (!id) continue; ids[id] = true;
    if (id.charAt(0) === 'o') continue;
    if (!String(datos[r][cMat - 1] || '').trim()) hoja.getRange(r + 1, cMat).setValue(RODIO.indexOf(id) >= 0 ? 'Rodio' : 'Covergold');
  }
  var cab = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim().toLowerCase(); });
  var nuevos = ORO.filter(function (o) { return !ids[o[0]]; }).map(function (o) {
    var v = { id: o[0], seccion: o[1], tipo: o[2], nombre: o[3], precio: o[4], costo: o[5], frase: o[6], descripcion: o[7], fotos: o[8], mostrar: 'SI', cantidad: '', material: 'Oro laminado 18k' };
    return cab.map(function (h) { return v[h] !== undefined ? v[h] : ''; });
  });
  if (nuevos.length) hoja.getRange(hoja.getLastRow() + 1, 1, nuevos.length, cab.length).setValues(nuevos);
  Logger.log('Listo. Productos de oro agregados: ' + nuevos.length);
}
