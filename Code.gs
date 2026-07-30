// ╔══════════════════════════════════════════════════════════╗
// ║  BOTIQUÍN DE ESCRITORIO — Backend v7                     ║
// ║  Cambios sobre v5:                                       ║
// ║   • Rangos DINÁMICOS (ya no se rompe al producto 53+)    ║
// ║   • Folio de ticket (agrupa items de una venta)          ║
// ║   • Costo snapshot al momento de venta (fiscal correcto) ║
// ║   • Hoja Configuración (PropertiesService) + UI          ║
// ║   • Hoja Promos_Items (componentes de promos)            ║
// ║   • Migración con backup automático                      ║
// ║   • Reporte fiscal por rango de fechas                   ║
// ║   • Enviar recibo por correo desde el backend            ║
// ║                                                          ║
// ║  Instalación:                                            ║
// ║  1. Apps Script → reemplaza Code.gs                      ║
// ║  2. Crea archivo HTML "Index" → pega Index.html          ║
// ║  3. Guarda → Ejecutar onOpen                             ║
// ║  4. En Sheets: menú "POS Botiquín"                       ║
// ║     a) "ACTUALIZAR ESTRUCTURA v6" (una vez)              ║
// ║     b) "Preparar datos"                                  ║
// ║  5. Implementar como Web App (yo / cualquiera con Google)║
// ╚══════════════════════════════════════════════════════════╝

var CACHE_KEY = 'bq_cat_v6';
var CFG_PREFIX = 'cfg_';
var STRUCT_VERSION = 'bq_struct_v';
var TARGET_STRUCT = 'v6';

// ══════════════════════════════════════════════
//  MENÚ
// ══════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('POS Botiquin')
    .addItem('🚀 ACTUALIZAR ESTRUCTURA v6 (primera vez)', 'actualizarEstructura')
    .addSeparator()
    .addItem('Preparar datos (recarga caché)', 'prepararDatos')
    .addItem('URL del sistema', 'mostrarUrl')
    .addSeparator()
    .addItem('Venta rápida emergencia', 'ventaRapida')
    .addItem('Reconstruir Fiados', 'reconstruirFiados')
    .addItem('Archivar mes anterior', 'archivarMes')
    .addItem('Limpiar caché', 'limpiarCache')
    .addToUi();
}

function mostrarUrl() {
  SpreadsheetApp.getUi().alert('URL del sistema:\n\n' + ScriptApp.getService().getUrl());
}

// ══════════════════════════════════════════════
//  MIGRACIÓN v5 → v6
// ══════════════════════════════════════════════
function actualizarEstructura() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var currentVer = PropertiesService.getScriptProperties().getProperty(STRUCT_VERSION) || 'v5';
  if (currentVer === TARGET_STRUCT) {
    var resp = ui.alert('Estructura ya está en ' + TARGET_STRUCT + '. ¿Re-ejecutar de todos modos? (idempotente)', ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
  }

  var confirm = ui.alert(
    'Actualización de estructura v5 → v6',
    'Esto hará:\n\n' +
    '1. Backup de la hoja Ventas → _Backup_Ventas_[fecha]\n' +
    '2. Agregar columnas a Ventas: P=Costo_Snapshot, Q=Folio, R=Costo_Snapshot_Source\n' +
    '3. Agregar columna J=archivado a Catálogo\n' +
    '4. Crear hoja "Promos_Items" (vacía)\n' +
    '5. Crear hoja "Configuracion" (vacía)\n' +
    '6. Recalcular costos snapshot de ventas históricas (usa costo actual de Catálogo como aproximación)\n' +
    '7. Asignar folios retroactivos agrupando por fecha+hora+cliente+método\n' +
    '8. Reemplazar fórmulas con rangos $F$500 por rangos de columna entera\n\n' +
    'Los datos existentes NO se borran. Las funciones del backend siguen igual.\n\n' +
    '¿Continuar?',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  try {
    var t0 = Date.now();

    // 1. Backup
    var ventas = ss.getSheetByName('Ventas');
    if (!ventas) throw new Error('No existe hoja Ventas');
    var ts = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd-HHmm');
    var backupName = '_Backup_Ventas_' + ts;
    if (!ss.getSheetByName(backupName)) {
      ventas.copyTo(ss).setName(backupName).hideSheet();
    }

    // 2. Columnas nuevas en Ventas (P, Q, R)
    if (ventas.getLastColumn() < 18) {
      ventas.getRange(3, 16, 1, 3).setValues([['Costo_Snapshot', 'Folio', 'Snapshot_Source']]);
    }

    // 3. Columna J en Catálogo (archivado)
    var cat = ss.getSheetByName('Catalogo');
    if (cat && cat.getRange(5, 10).getValue() !== 'Archivado') {
      cat.getRange(5, 10).setValue('Archivado');
    }

    // 4. Hoja Promos_Items
    var pi = ss.getSheetByName('Promos_Items');
    if (!pi) {
      pi = ss.insertSheet('Promos_Items');
      pi.getRange(1, 1, 1, 4).setValues([['promo_id', 'producto_id', 'cantidad', 'notas']]);
      pi.setFrozenRows(1);
      pi.getRange(1, 1, 1, 4).setBackground('#1E2A1E').setFontColor('#6DFF1A').setFontWeight('bold');
    }

    // 5. Hoja Configuracion (k/v editable a mano si quieren)
    var cfg = ss.getSheetByName('Configuracion');
    if (!cfg) {
      cfg = ss.insertSheet('Configuracion');
      cfg.getRange(1, 1, 1, 3).setValues([['Clave', 'Valor', 'Descripción']]);
      cfg.setFrozenRows(1);
      cfg.getRange(1, 1, 1, 3).setBackground('#1E2A1E').setFontColor('#6DFF1A').setFontWeight('bold');
      var defaults = [
        ['negocio_nombre', 'Botiquín de Escritorio', 'Nombre del negocio (recibo header)'],
        ['negocio_subtitulo', 'by Jairo Lopez', 'Subtítulo del recibo'],
        ['negocio_correo', '', 'Correo del negocio'],
        ['negocio_telefono', '', 'Teléfono'],
        ['negocio_slack', '', 'Canal/usuario de Slack para pedidos'],
        ['negocio_whatsapp', '', 'Número de WhatsApp para pedidos'],
        ['banks_json', '[]', 'Array JSON de cuentas bancarias'],
        ['qrs_json', '[]', 'Array JSON de QRs de pago'],
        ['pago_tarjeta_proximamente', 'true', '"Tarjeta: Próximamente"'],
        ['msg_pago_default', '🎉 ¡Gracias por tu compra! 🛍️ Esperamos verte pronto. Que tengas un excelente día. ✨', 'Mensaje default modo Pago'],
        ['msg_adeudo_default', '📋 Aquí está el detalle de tu cuenta pendiente. 🙏 Cuando puedas te agradezco el pago. ¡Gracias por tu preferencia! 💚', 'Mensaje default modo Adeudo']
      ];
      cfg.getRange(2, 1, defaults.length, 3).setValues(defaults);
    }

    // 6. Costo snapshot retroactivo (lee Catalogo y rellena col P)
    var scan = _scanCatalogoSheet(cat);
    var prodMap = {};
    var pData = cat.getRange(scan.prodStart, 1, scan.prodEnd - scan.prodStart + 1, 9).getValues();
    pData.forEach(function(r) {
      var nom = String(r[1] || '').trim();
      if (nom) prodMap[nom] = Number(r[4]) || 0; // costo
    });
    var prMap = {};
    var prData = cat.getRange(scan.promoStart, 1, scan.promoEnd - scan.promoStart + 1, 6).getValues();
    prData.forEach(function(r) {
      var nom = String(r[1] || '').trim();
      if (nom) prMap[nom] = Number(r[5]) || 0;
    });

    var vData = ventas.getDataRange().getValues();
    var snapWrites = [];
    var folioWrites = [];
    var sourceWrites = [];

    // Agrupar para folios retroactivos
    var grupos = {};
    for (var i = 3; i < vData.length; i++) {
      var row = vData[i];
      var art = String(row[4] || '').trim();
      if (!art) continue;
      // Si ya hay snapshot escrito, respetarlo
      var existingSnap = ventas.getLastColumn() >= 16 ? row[15] : null;
      var existingFolio = ventas.getLastColumn() >= 17 ? row[16] : null;
      var cant = Number(row[5]) || 0;
      var costoUnit = 0;
      if (existingSnap && existingSnap !== '') {
        costoUnit = Number(existingSnap);
      } else {
        var tipo = String(row[2] || '').trim();
        costoUnit = (tipo === 'Promo' ? prMap[art] : prodMap[art]) || 0;
      }

      var fecha = row[0] instanceof Date ? row[0] : new Date(row[0]);
      var fStr = Utilities.formatDate(fecha, ss.getSpreadsheetTimeZone(), 'yyyyMMdd');
      var clave = fStr + '|' + String(row[1] || '') + '|' + String(row[13] || '') + '|' + String(row[11] || '');

      if (!grupos[clave]) {
        grupos[clave] = 'T-OLD-' + fStr + '-' + (Object.keys(grupos).length + 1).toString().padStart(4, '0');
      }
      var folio = existingFolio || grupos[clave];

      snapWrites.push([costoUnit]);
      folioWrites.push([folio]);
      sourceWrites.push([existingSnap ? 'preserved' : 'migration']);
    }

    if (snapWrites.length) {
      var startRow = 4;
      ventas.getRange(startRow, 16, snapWrites.length, 1).setValues(snapWrites);
      ventas.getRange(startRow, 17, folioWrites.length, 1).setValues(folioWrites);
      ventas.getRange(startRow, 18, sourceWrites.length, 1).setValues(sourceWrites);

      // Reescribir J (Costo Total) y K (Ganancia) como fórmulas que usan P en vez de VLOOKUP
      var formulas = [];
      for (var rr = 0; rr < snapWrites.length; rr++) {
        var r = startRow + rr;
        formulas.push([
          '=IF(F' + r + '="","",F' + r + '*P' + r + ')',  // J = cant * costo snapshot
          '=IF(OR(F' + r + '="",J' + r + '=""),"",I' + r + '-J' + r + ')'  // K = total - costo
        ]);
      }
      ventas.getRange(startRow, 10, formulas.length, 2).setFormulas(formulas);
    }

    // 7. Reemplazar fórmulas con $F$500 por columna entera en Inventario
    var inv = ss.getSheetByName('Inventario');
    if (inv) {
      var invLast = inv.getLastRow();
      for (var rr = 5; rr <= invLast; rr++) {
        var gF = inv.getRange(rr, 7).getFormula();
        if (gF && gF.indexOf('$F$500') > -1) {
          inv.getRange(rr, 7).setFormula(gF.replace(/\$F\$4:\$F\$500/g, '$F:$F').replace(/\$E\$4:\$E\$500/g, '$E:$E').replace(/\$C\$4:\$C\$500/g, '$C:$C'));
        }
        var hF = inv.getRange(rr, 8).getFormula();
        if (hF && hF.indexOf('$F$500') > -1) {
          inv.getRange(rr, 8).setFormula(hF.replace(/\$F\$4:\$F\$500/g, '$F:$F').replace(/\$E\$4:\$E\$500/g, '$E:$E').replace(/\$C\$4:\$C\$500/g, '$C:$C'));
        }
      }
    }

    PropertiesService.getScriptProperties().setProperty(STRUCT_VERSION, TARGET_STRUCT);

    var ms = Date.now() - t0;
    ui.alert('✅ Estructura actualizada a ' + TARGET_STRUCT + ' en ' + ms + 'ms.\n\n' +
      'Backup creado: ' + backupName + '\n' +
      'Filas de Ventas procesadas: ' + snapWrites.length + '\n\n' +
      'Ahora ejecuta "Preparar datos" para recargar el caché.');
  } catch (e) {
    ui.alert('Error en migración:\n\n' + String(e) + '\n\nStack:\n' + (e.stack || ''));
  }
}

// ══════════════════════════════════════════════
//  DISCOVERY DINÁMICO DE RANGOS EN CATALOGO
// ══════════════════════════════════════════════
function _scanCatalogoSheet(sh) {
  // Busca marcadores en columna A. Devuelve rangos.
  var last = sh.getLastRow();
  var colA = sh.getRange(1, 1, last, 1).getValues();
  var prodHdr = -1, promoHdr = -1, catHdr = -1;
  for (var i = 0; i < colA.length; i++) {
    var v = String(colA[i][0] || '').trim().toUpperCase();
    if (prodHdr < 0 && v.indexOf('PRODUCTO') === 0) prodHdr = i + 1;
    else if (promoHdr < 0 && (v.indexOf('PROMO') === 0 || v.indexOf('COMBO') === 0)) promoHdr = i + 1;
    else if (catHdr < 0 && v.indexOf('CATEGORIA') === 0 && i > 50) catHdr = i + 1;
  }
  // Fallback a los valores antiguos si no se encuentran marcadores
  if (prodHdr < 0) prodHdr = 4;
  if (promoHdr < 0) promoHdr = 59;
  if (catHdr < 0) catHdr = 82;

  // El header de columnas viene 1-2 filas después
  var prodStart = prodHdr + 2; // suele estar fila 6
  var promoStart = promoHdr + 3; // suele estar fila 63 (header en 62)
  var catStart = catHdr + 3; // suele estar fila 85

  // Encontrar el final de productos = una fila antes del marcador promo
  var prodEnd = Math.max(prodStart, promoHdr - 1);
  var promoEnd = Math.max(promoStart, catHdr - 1);
  var catEnd = last;

  return { prodStart: prodStart, prodEnd: prodEnd, promoStart: promoStart, promoEnd: promoEnd, catStart: catStart, catEnd: catEnd };
}

// ══════════════════════════════════════════════
//  PREPARAR DATOS
// ══════════════════════════════════════════════
function prepararDatos() {
  var ui = SpreadsheetApp.getUi();
  try {
    var t0 = Date.now();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cat = ss.getSheetByName('Catalogo');
    if (!cat) { ui.alert('No existe hoja Catalogo'); return; }

    var scan = _scanCatalogoSheet(cat);
    var prods = [], promos = [], cats = [];

    var pData = cat.getRange(scan.prodStart, 1, scan.prodEnd - scan.prodStart + 1, 10).getValues();
    pData.forEach(function(r) {
      var nom = String(r[1] || '').trim();
      if (!nom || nom.startsWith('\u2190') || nom === 'Nombre Articulo') return;
      var precio = Number(r[3]); if (!precio) return;
      var archivado = r[9] === true || String(r[9] || '').toLowerCase() === 'true' || String(r[9] || '').toLowerCase() === 'si';
      prods.push({
        id: String(r[0] || '').trim(),
        nombre: nom,
        cat: String(r[2] || 'Varios').trim(),
        precio: precio,
        costo: Number(r[4]) || 0,
        notas: String(r[7] || '').trim(),
        prov: String(r[8] || '').trim(),
        archivado: archivado
      });
    });

    var prData = cat.getRange(scan.promoStart, 1, scan.promoEnd - scan.promoStart + 1, 6).getValues();
    prData.forEach(function(r) {
      var nom = String(r[1] || '').trim();
      if (!nom || nom.startsWith('\u2190') || ['ID', 'Nombre Promo', 'Nueva Promo'].indexOf(nom) >= 0) return;
      var precio = Number(r[2]); if (!precio) return;
      promos.push({
        id: String(r[0] || '').trim(),
        nombre: nom,
        precio: precio,
        ahorro: Number(r[3]) || 0,
        contenido: String(r[4] || '').trim(),
        costo: Number(r[5]) || 0
      });
    });

    var cData = cat.getRange(scan.catStart, 2, Math.max(1, scan.catEnd - scan.catStart + 1), 1).getValues();
    cData.forEach(function(r) {
      var c = String(r[0] || '').trim();
      if (c && c !== 'Categoria' && cats.indexOf(c) < 0) cats.push(c);
    });

    var payload = JSON.stringify({ prods: prods, promos: promos, cats: cats });
    PropertiesService.getScriptProperties().setProperty(CACHE_KEY, payload);

    var ms = Date.now() - t0;
    ui.alert('✅ Datos listos en ' + ms + 'ms\n\n' +
      prods.length + ' productos (' + prods.filter(function(p){return !p.archivado;}).length + ' activos)\n' +
      promos.length + ' promos\n' +
      cats.length + ' categorías\n\n' +
      'Abre la URL del sistema para usar.');
  } catch (e) {
    ui.alert('Error: ' + String(e) + '\n\n' + (e.stack || ''));
  }
}

function limpiarCache() {
  PropertiesService.getScriptProperties().deleteProperty(CACHE_KEY);
  SpreadsheetApp.getUi().alert('Caché limpiado. Ejecuta "Preparar datos" antes de usar.');
}

// ══════════════════════════════════════════════
//  ENTRY POINT WEB APP
// ══════════════════════════════════════════════
function doGet() {
  var t = HtmlService.createTemplateFromFile('Index');
  return t.evaluate()
    .setTitle('Botiquín de Escritorio')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width,initial-scale=1');
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// ══════════════════════════════════════════════
//  HELPERS GENERALES
// ══════════════════════════════════════════════
function _ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function _tz() { return _ss().getSpreadsheetTimeZone(); }
function _fmt(v) { if (v instanceof Date) return Utilities.formatDate(v, _tz(), 'dd/MM/yyyy'); return String(v || '').substring(0, 10); }
function _fmtDT(v) { if (v instanceof Date) return Utilities.formatDate(v, _tz(), 'dd/MM/yyyy HH:mm'); return String(v || ''); }
function _lastRow(sh) {
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 3; i--) {
    if (data[i][0] instanceof Date || (data[i][0] && data[i][0] !== '')) return i + 1;
  }
  return 4;
}

function _nuevoFolio() {
  var now = new Date();
  var key = 'folio_seq_' + Utilities.formatDate(now, _tz(), 'yyyyMMdd');
  var props = PropertiesService.getScriptProperties();
  var seq = parseInt(props.getProperty(key) || '0') + 1;
  props.setProperty(key, String(seq));
  return 'T-' + Utilities.formatDate(now, _tz(), 'yyyyMMdd') + '-' + String(seq).padStart(4, '0');
}

// ══════════════════════════════════════════════
//  API: CATALOGO / CLIENTES / CONFIG
// ══════════════════════════════════════════════
function getCatalogo() {
  var raw = PropertiesService.getScriptProperties().getProperty(CACHE_KEY);
  if (!raw) return { ok: false, err: 'CACHE_VACIO', prods: [], promos: [], cats: [] };
  var d = JSON.parse(raw);
  d.ok = true;
  d.struct = PropertiesService.getScriptProperties().getProperty(STRUCT_VERSION) || 'v5';
  return d;
}

function getClientes() {
  try {
    var sh = _ss().getSheetByName('Clientes');
    if (!sh || sh.getLastRow() < 2) return { ok: true, clientes: [] };
    // Asegurar columnas extendidas A=Nombre B=Telefono C=Correo D=Slack E=WhatsApp F=Notas
    if (sh.getRange(1, 3).getValue() !== 'Correo') {
      sh.getRange(1, 1, 1, 6).setValues([['Nombre','Telefono','Correo','Slack','WhatsApp','Notas']]);
    }
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues()
      .filter(function(r) { return r[0] && String(r[0]).trim(); })
      .map(function(r) { return {
        nombre: String(r[0]).trim(),
        tel: r[1] || '',
        correo: r[2] || '',
        slack: r[3] || '',
        whatsapp: r[4] || '',
        notas: r[5] || ''
      }; });
    return { ok: true, clientes: rows };
  } catch (e) { return { ok: false, err: String(e) }; }
}

function getPerfilCliente(nombre) {
  try {
    var sh = _ss().getSheetByName('Clientes');
    var ventas = _ss().getSheetByName('Ventas');
    var datos = { nombre: nombre, tel: '', correo: '', slack: '', whatsapp: '', notas: '' };
    if (sh && sh.getLastRow() >= 2) {
      // asegurar columnas
      if (sh.getRange(1, 3).getValue() !== 'Correo') {
        sh.getRange(1, 1, 1, 6).setValues([['Nombre','Telefono','Correo','Slack','WhatsApp','Notas']]);
      }
      var data = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
      for (var i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
          datos = { nombre: String(data[i][0]).trim(), tel: data[i][1]||'', correo: data[i][2]||'', slack: data[i][3]||'', whatsapp: data[i][4]||'', notas: data[i][5]||'', row: i + 2 };
          break;
        }
      }
    }
    // Historial completo
    var vData = ventas.getDataRange().getValues();
    var historial = [];
    var totalGastado = 0, totalPendiente = 0, ultimaCompra = '', numTickets = {};
    for (var i = 3; i < vData.length; i++) {
      var row = vData[i];
      if (String(row[13] || '').trim().toLowerCase() !== nombre.toLowerCase()) continue;
      var fStr = _fmt(row[0]);
      var tot = Number(row[8]) || 0;
      var est = String(row[12] || '').trim();
      historial.push({
        row: i + 1, fecha: fStr, hora: String(row[1]||''),
        art: String(row[4]||''), cant: Number(row[5])||0,
        precio: Number(row[6])||0, total: tot,
        metodo: String(row[11]||''), estado: est,
        folio: String(row[16]||''), notas: String(row[14]||'')
      });
      totalGastado += tot;
      if (est === 'Fiado' || est === 'Al rato te pago') totalPendiente += tot;
      if (row[16]) numTickets[String(row[16])] = true;
      if (!ultimaCompra || fStr > ultimaCompra) ultimaCompra = fStr;
    }
    historial.reverse();
    return {
      ok: true, datos: datos, historial: historial,
      stats: { totalGastado: totalGastado, totalPendiente: totalPendiente, numTickets: Object.keys(numTickets).length, ultimaCompra: ultimaCompra, totalCompras: historial.length }
    };
  } catch (e) { return { ok: false, err: String(e) }; }
}

function guardarCliente(d) {
  try {
    var sh = _ss().getSheetByName('Clientes');
    if (!sh) { sh = _ss().insertSheet('Clientes'); sh.getRange(1, 1, 1, 6).setValues([['Nombre','Telefono','Correo','Slack','WhatsApp','Notas']]); }
    if (sh.getRange(1, 3).getValue() !== 'Correo') {
      sh.getRange(1, 1, 1, 6).setValues([['Nombre','Telefono','Correo','Slack','WhatsApp','Notas']]);
    }
    var nom = String(d.nombre || '').trim();
    if (!nom) return { ok: false, msg: 'Falta nombre' };
    // Si trae row, actualizar; si no, buscar por nombre o crear
    var row = d.row;
    if (!row) {
      var last = sh.getLastRow();
      if (last >= 2) {
        var names = sh.getRange(2, 1, last - 1, 1).getValues().flat().map(function(v) { return String(v).trim().toLowerCase(); });
        var idx = names.indexOf(nom.toLowerCase());
        if (idx >= 0) row = idx + 2;
      }
      if (!row) row = sh.getLastRow() + 1;
    }
    sh.getRange(row, 1, 1, 6).setValues([[nom, d.tel||'', d.correo||'', d.slack||'', d.whatsapp||'', d.notas||'']]);
    return { ok: true, msg: nom + ' guardado.' };
  } catch (e) { return { ok: false, msg: String(e) }; }
}

function obtenerConfiguracion() {
  try {
    var props = PropertiesService.getScriptProperties().getProperties();
    var cfg = {};
    Object.keys(props).forEach(function(k) {
      if (k.indexOf(CFG_PREFIX) === 0) cfg[k.substring(CFG_PREFIX.length)] = props[k];
    });
    // Hoja Configuracion: para valores grandes (logos, imágenes QR base64, etc.)
    // que no caben en PropertiesService (max ~9KB por valor).
    // Tiene prioridad sobre Properties cuando ambos existen (canónica para blobs).
    var sh = _ss().getSheetByName('Configuracion');
    if (sh && sh.getLastRow() >= 2) {
      var data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
      data.forEach(function(r) {
        var k = String(r[0] || '').trim();
        if (!k) return;
        var v = String(r[1] || '');
        if (v) cfg[k] = v; // sheet gana si tiene contenido
        else if (cfg[k] === undefined) cfg[k] = '';
      });
    }
    return { ok: true, cfg: cfg };
  } catch (e) { return { ok: false, err: String(e) }; }
}

/**
 * v7.2 — Guardar configuración con ruteo automático por tamaño.
 *   - Valores <= 8000 chars  → PropertiesService (rápido).
 *   - Valores > 8000 chars   → Hoja Configuracion (cabe hasta 50k chars/celda).
 * Esto permite guardar imágenes de QR (data URLs ~50-250 KB) y logos sin romperse.
 * Se mantiene una sola fuente de verdad por clave (sin duplicados).
 */
function guardarConfiguracion(obj) {
  try {
    var props = PropertiesService.getScriptProperties();
    var sh = _ss().getSheetByName('Configuracion');
    if (!sh) {
      sh = _ss().insertSheet('Configuracion');
      sh.getRange(1, 1, 1, 3).setValues([['Clave', 'Valor', 'Descripción']]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, 3).setBackground('#1E2A1E').setFontColor('#6DFF1A').setFontWeight('bold');
    }

    // Mapa actual de claves en hoja → fila
    var sheetMap = {};
    if (sh.getLastRow() >= 2) {
      var data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
      data.forEach(function(r, i) {
        var k = String(r[0] || '').trim();
        if (k) sheetMap[k] = { row: i + 2 };
      });
    }

    Object.keys(obj).forEach(function(k) {
      var v = String(obj[k] == null ? '' : obj[k]);
      var large = v.length > 8000;
      var propKey = CFG_PREFIX + k;
      if (large) {
        // Va a la hoja
        if (sheetMap[k]) {
          sh.getRange(sheetMap[k].row, 2).setValue(v);
        } else {
          sh.appendRow([k, v, '(blob automático)']);
        }
        // Limpiar de Properties si quedó algo (evitar stale)
        try { props.deleteProperty(propKey); } catch (e) {}
      } else {
        // Va a Properties
        try {
          props.setProperty(propKey, v);
        } catch (e) {
          // Si por algún motivo Properties rebota (>9KB borderline), cae a hoja
          if (sheetMap[k]) sh.getRange(sheetMap[k].row, 2).setValue(v);
          else sh.appendRow([k, v, '(fallback)']);
        }
        // Limpiar de hoja si la clave estaba ahí (evitar stale)
        if (sheetMap[k]) sh.getRange(sheetMap[k].row, 2).clearContent();
      }
    });
    return { ok: true };
  } catch (e) { return { ok: false, err: String(e) + ' ' + (e.stack || '') }; }
}

// ══════════════════════════════════════════════
//  API: VENTAS
// ══════════════════════════════════════════════
function registrarVenta(d) {
  try {
    var ss = _ss();
    var sh = ss.getSheetByName('Ventas');
    var cat = ss.getSheetByName('Catalogo');
    var tz = ss.getSpreadsheetTimeZone();
    var now = new Date();
    var hora = Utilities.formatDate(now, tz, 'HH:mm');
    var start = _lastRow(sh) + 1;
    var folio = _nuevoFolio();

    var scan = _scanCatalogoSheet(cat);
    // Cachear costos para snapshot
    var prodCost = {}, promoCost = {};
    cat.getRange(scan.prodStart, 1, scan.prodEnd - scan.prodStart + 1, 5).getValues().forEach(function(r) {
      var n = String(r[1] || '').trim(); if (n) prodCost[n] = Number(r[4]) || 0;
    });
    cat.getRange(scan.promoStart, 1, scan.promoEnd - scan.promoStart + 1, 6).getValues().forEach(function(r) {
      var n = String(r[1] || '').trim(); if (n) promoCost[n] = Number(r[5]) || 0;
    });

    var rows = [];
    d.items.forEach(function(it, idx) {
      var r = start + idx;
      var snap = it.tipo === 'Promo' ? (promoCost[it.nombre] || 0) : (prodCost[it.nombre] || 0);
      rows.push([
        now,                          // A Fecha
        hora,                          // B Hora
        it.tipo,                       // C Tipo
        '=IF(E' + r + '="","",IF(C' + r + '="Promo","Promo",IFERROR(VLOOKUP(E' + r + ',Catalogo!$B:$C,2,0),"Varios")))', // D Categoría
        it.nombre,                     // E Artículo
        it.cantidad,                   // F Cant
        it.precio,                     // G Precio
        '',                            // H Precio Compra (legacy)
        it.precio * it.cantidad,       // I Total
        '=IF(F' + r + '="","",F' + r + '*P' + r + ')',                  // J Costo Total
        '=IF(OR(F' + r + '="",J' + r + '=""),"",I' + r + '-J' + r + ')', // K Ganancia
        d.metodo,                       // L Método
        d.estado,                       // M Estado
        d.cliente || '',                // N Cliente
        d.notas || '',                  // O Notas
        snap,                           // P Costo_Snapshot
        folio,                          // Q Folio
        'sale'                          // R Snapshot_Source
      ]);
    });

    sh.getRange(start, 1, rows.length, 18).setValues(rows);
    sh.getRange(start, 1, rows.length, 1).setNumberFormat('DD/MM/YYYY');

    if (d.cliente && d.cliente.trim()) _autoCliente(d.cliente.trim());

    // Si hay promos con componentes, descontar stock automáticamente vía actualización del cache
    // (las fórmulas de Inventario lo recogen al recalcular)

    var total = d.items.reduce(function(s, it) { return s + it.precio * it.cantidad; }, 0);
    return { ok: true, total: total, rows: rows.length, folio: folio };
  } catch (e) { return { ok: false, err: String(e) }; }
}

function getVentas(params) {
  try {
    var sh = _ss().getSheetByName('Ventas');
    var tz = _tz();
    var data = sh.getDataRange().getValues();
    var p = params || {};
    var now = new Date();
    var hoy = Utilities.formatDate(now, tz, 'dd/MM/yyyy');
    var mes = Utilities.formatDate(now, tz, 'MM/yyyy');

    var rows = [];
    for (var i = 3; i < data.length; i++) {
      var row = data[i];
      var art = String(row[4] || '').trim(); if (!art) continue;
      var fStr = _fmt(row[0]); if (!fStr || fStr.length < 8) continue;
      if (p.filtro === 'hoy' && fStr !== hoy) continue;
      if (p.filtro === 'mes' && fStr.substring(3) !== mes) continue;

      rows.push({
        row: i + 1,
        fecha: fStr,
        hora: String(row[1] || ''),
        tipo: String(row[2] || ''),
        cat: String(row[3] || ''),
        art: art,
        cant: Number(row[5]) || 0,
        precio: Number(row[6]) || 0,
        total: Number(row[8]) || 0,
        costo: Number(row[9]) || 0,
        gan: Number(row[10]) || 0,
        metodo: String(row[11] || ''),
        estado: String(row[12] || ''),
        cliente: String(row[13] || ''),
        notas: String(row[14] || ''),
        folio: String(row[16] || '')
      });
    }
    rows.reverse();

    // Agrupar por folio si se solicita
    if (p.modo === 'tickets') {
      var tickets = {};
      rows.forEach(function(r) {
        var f = r.folio || ('NO-FOLIO-' + r.fecha + '-' + r.hora + '-' + r.cliente);
        if (!tickets[f]) {
          tickets[f] = {
            folio: r.folio || '',
            fecha: r.fecha, hora: r.hora,
            cliente: r.cliente, metodo: r.metodo, estado: r.estado,
            notas: r.notas, items: [], total: 0, gan: 0, costo: 0
          };
        }
        tickets[f].items.push(r);
        tickets[f].total += r.total;
        tickets[f].gan += r.gan;
        tickets[f].costo += r.costo;
      });
      var tList = Object.keys(tickets).map(function(k) { return tickets[k]; });
      return { ok: true, tickets: tList, count: tList.length };
    }

    return { ok: true, rows: rows };
  } catch (e) { return { ok: false, err: String(e), rows: [] }; }
}

// ══════════════════════════════════════════════
//  API: FIADOS
// ══════════════════════════════════════════════
function getFiados() {
  try {
    var sh = _ss().getSheetByName('Ventas');
    var data = sh.getDataRange().getValues();
    var map = {}, total = 0;
    for (var i = 3; i < data.length; i++) {
      var art = String(data[i][4] || '').trim(); if (!art) continue;
      var est = String(data[i][12] || '').trim();
      if (est !== 'Fiado' && est !== 'Al rato te pago') continue;
      var cli = String(data[i][13] || 'Desconocido').trim() || 'Desconocido';
      var tot = Number(data[i][8]) || 0;
      if (!map[cli]) map[cli] = { cli: cli, items: [], tot: 0, folios: {} };
      map[cli].items.push({
        row: i + 1, fecha: _fmt(data[i][0]), art: art,
        cant: data[i][5], tot: tot, metodo: String(data[i][11] || ''),
        folio: String(data[i][16] || ''), precio: Number(data[i][6]) || 0
      });
      if (data[i][16]) map[cli].folios[String(data[i][16])] = true;
      map[cli].tot += tot;
      total += tot;
    }
    var lista = Object.keys(map).map(function(k) {
      map[k].numFolios = Object.keys(map[k].folios).length;
      delete map[k].folios;
      return map[k];
    }).sort(function(a, b) { return b.tot - a.tot; });
    return { ok: true, lista: lista, total: total };
  } catch (e) { return { ok: false, err: String(e), lista: [], total: 0 }; }
}

function marcarPagado(rows) {
  try {
    var sh = _ss().getSheetByName('Ventas');
    rows.forEach(function(r) { sh.getRange(r, 13).setValue('Pagado'); });
    return { ok: true, msg: rows.length + ' ventas marcadas como Pagado.' };
  } catch (e) { return { ok: false, msg: String(e) }; }
}

// ══════════════════════════════════════════════
//  API: REPORTES
// ══════════════════════════════════════════════
function getReportes(filtro) {
  try {
    var sh = _ss().getSheetByName('Ventas');
    var tz = _tz();
    var data = sh.getDataRange().getValues();
    var now = new Date();
    var hoy = Utilities.formatDate(now, tz, 'dd/MM/yyyy');
    var mes = Utilities.formatDate(now, tz, 'MM/yyyy');
    var totalV = 0, totalG = 0, fiados = 0;
    var porProd = {}, porMetodo = {}, porCat = {}, porDia = {}, porHora = {};

    for (var i = 3; i < data.length; i++) {
      var row = data[i];
      var art = String(row[4] || '').trim(); if (!art) continue;
      var fecha = row[0] instanceof Date ? row[0] : new Date(row[0]);
      var fStr = _fmt(fecha); if (!fStr || fStr.length < 8) continue;
      if (filtro === 'hoy' && fStr !== hoy) continue;
      if (filtro === 'mes' && fStr.substring(3) !== mes) continue;
      var v = Number(row[8]) || 0, g = Number(row[10]) || 0;
      var m = String(row[11] || 'Efectivo').trim();
      var cat = String(row[3] || '').trim();
      var e = String(row[12] || '').trim();
      totalV += v; totalG += g;
      if (e === 'Fiado' || e === 'Al rato te pago') fiados += v;
      porMetodo[m] = (porMetodo[m] || 0) + v;
      if (cat) porCat[cat] = (porCat[cat] || 0) + v;
      if (!porProd[art]) porProd[art] = { cant: 0, total: 0, gan: 0 };
      porProd[art].cant += Number(row[5]) || 1;
      porProd[art].total += v;
      porProd[art].gan += g;
      if (!porDia[fStr]) porDia[fStr] = { v: 0, g: 0 };
      porDia[fStr].v += v; porDia[fStr].g += g;
      // Heatmap día-de-semana × hora
      var dow = fecha.getDay(); // 0=Domingo
      var horaStr = String(row[1] || '').substring(0, 2);
      var hh = parseInt(horaStr) || 0;
      var hKey = dow + '_' + hh;
      porHora[hKey] = (porHora[hKey] || 0) + v;
    }

    var top = Object.keys(porProd)
      .map(function(k) { return { nombre: k, cant: porProd[k].cant, total: porProd[k].total, gan: porProd[k].gan }; })
      .sort(function(a, b) { return b.total - a.total; }).slice(0, 10);
    var dias = Object.keys(porDia).sort()
      .map(function(k) { return { fecha: k, v: porDia[k].v, g: porDia[k].g }; }).slice(-30);

    return {
      ok: true, totalV: totalV, totalG: totalG, fiados: fiados,
      margen: totalV > 0 ? (totalG / totalV * 100) : 0,
      top: top, porMetodo: porMetodo, porCat: porCat, dias: dias, porHora: porHora
    };
  } catch (e) { return { ok: false, err: String(e) }; }
}

function getReporteFiscal(inicio, fin) {
  try {
    var sh = _ss().getSheetByName('Ventas');
    var data = sh.getDataRange().getValues();
    var tz = _tz();
    var dIni = new Date(inicio + 'T00:00:00');
    var dFin = new Date(fin + 'T23:59:59');
    var ventas = 0, costos = 0, ganancia = 0;
    var porMetodo = {}, porEstado = {}, porDia = {};
    var foliosSet = {};
    var items = [];
    for (var i = 3; i < data.length; i++) {
      var row = data[i];
      var art = String(row[4] || '').trim(); if (!art) continue;
      var f = row[0] instanceof Date ? row[0] : new Date(row[0]);
      if (f < dIni || f > dFin) continue;
      var v = Number(row[8]) || 0, c = Number(row[9]) || 0, g = Number(row[10]) || 0;
      var m = String(row[11] || 'Efectivo').trim();
      var est = String(row[12] || 'Pagado').trim();
      var fStr = _fmt(f);
      ventas += v; costos += c; ganancia += g;
      porMetodo[m] = (porMetodo[m] || 0) + v;
      porEstado[est] = (porEstado[est] || 0) + v;
      if (!porDia[fStr]) porDia[fStr] = { ventas: 0, ganancia: 0, tickets: {} };
      porDia[fStr].ventas += v;
      porDia[fStr].ganancia += g;
      var folio = String(row[16] || ('row' + (i + 1)));
      porDia[fStr].tickets[folio] = true;
      foliosSet[folio] = true;
      items.push({
        folio: folio, fecha: fStr, hora: String(row[1] || ''),
        producto: art, cant: row[5], precio: row[6], total: v, costo: c, ganancia: g,
        metodo: m, estado: est, cliente: String(row[13] || '')
      });
    }
    var diasArr = Object.keys(porDia).sort().map(function(k) {
      return { fecha: k, ventas: porDia[k].ventas, ganancia: porDia[k].ganancia, tickets: Object.keys(porDia[k].tickets).length };
    });
    var totalTickets = Object.keys(foliosSet).length;
    return {
      ok: true,
      periodo: { inicio: inicio, fin: fin },
      totales: { ventas: ventas, costos: costos, ganancia: ganancia, ivaEstimado: 0 },
      porMetodo: porMetodo, porEstado: porEstado, porDia: diasArr,
      totalTickets: totalTickets,
      ticketPromedio: totalTickets > 0 ? ventas / totalTickets : 0,
      items: items
    };
  } catch (e) { return { ok: false, err: String(e) }; }
}

// ══════════════════════════════════════════════
//  API: INVENTARIO / RESTOCK / PRODUCTOS (sin cambios mayores)
// ══════════════════════════════════════════════
function getInventario() {
  try {
    var inv = _ss().getSheetByName('Inventario');
    var data = inv.getRange(5, 1, 200, 15).getValues();
    var all = [], bajos = [];
    data.forEach(function(row) {
      var nom = String(row[0] || '').trim();
      if (!nom || nom === 'TOTAL INVENTARIO') return;
      var stock = Number(row[9]) || 0, alerta = Number(row[10]) || 3;
      var estado = String(row[12] || '').trim(), tam = Number(row[2]) || 1, cp = Number(row[3]) || 0;
      var item = { nombre: nom, stock: Math.round(stock), alerta: alerta, estado: estado || 'OK', tamPaq: tam, costoPaq: cp, urgente: ['AGOTADO', 'URGENTE'].indexOf(estado) >= 0 };
      all.push(item);
      if (['AGOTADO', 'URGENTE', 'BAJO'].indexOf(estado) >= 0 || stock <= alerta) {
        var sug = Math.max(tam, Math.ceil((alerta * 3 - stock) / tam) * tam);
        bajos.push(Object.assign({}, item, { sugerido: sug }));
      }
    });
    bajos.sort(function(a, b) { return (b.urgente ? 1 : 0) - (a.urgente ? 1 : 0); });
    return { ok: true, all: all, bajos: bajos };
  } catch (e) { return { ok: false, err: String(e), all: [], bajos: [] }; }
}

function getStockActual(nombre) {
  try {
    var cat = _ss().getSheetByName('Catalogo');
    var inv = _ss().getSheetByName('Inventario');
    var scan = _scanCatalogoSheet(cat);
    var pData = cat.getRange(scan.prodStart, 1, scan.prodEnd - scan.prodStart + 1, 9).getValues();
    var prod = null;
    pData.forEach(function(r) { if (String(r[1] || '').trim() === nombre.trim()) prod = r; });
    if (!prod) return { ok: false, err: 'No encontrado' };
    var iData = inv.getRange(5, 1, 200, 10).getValues();
    var sa = 0, tam = 1, cp = 0;
    iData.forEach(function(r) { if (String(r[0] || '').trim() === nombre.trim()) { tam = r[2] || 1; cp = r[3] || 0; sa = r[9] || 0; } });
    return {
      ok: true, nombre: prod[1], precio: Number(prod[3]) || 0, costoUnit: Number(prod[4]) || 0,
      cat: prod[2], prov: prod[8] || '', stockAct: sa, tamPaq: tam, costoPaq: cp
    };
  } catch (e) { return { ok: false, err: String(e) }; }
}

function procesarRestock(d) {
  try {
    var cat = _ss().getSheetByName('Catalogo');
    var inv = _ss().getSheetByName('Inventario');
    var scan = _scanCatalogoSheet(cat);
    var nom = d.nombre.trim(), nC = Number(d.nuevaCant), nU = Number(d.nuevoCosto);
    var sA = Number(d.stockAct), cA = Number(d.costoAct);
    if (!nC || !nU || nC <= 0 || nU <= 0) return { ok: false, err: 'Cantidad y costo deben ser > 0' };
    // Si el stock actual es <= 0 (descuadre), no usar promedio ponderado:
    // el costo promedio nuevo es simplemente el costo del lote nuevo.
    var tU, cP;
    if (sA <= 0) {
      tU = nC;
      cP = nU;
    } else {
      tU = sA + nC;
      cP = ((sA * cA) + (nC * nU)) / tU;
    }
    var pData = cat.getRange(scan.prodStart, 1, scan.prodEnd - scan.prodStart + 1, 9).getValues();
    pData.forEach(function(r, i) {
      if (String(r[1] || '').trim() === nom) {
        cat.getRange(scan.prodStart + i, 5).setValue(Math.round(cP * 10000) / 10000);
        if (d.nuevoPrecio) cat.getRange(scan.prodStart + i, 4).setValue(Number(d.nuevoPrecio));
        if (d.prov) cat.getRange(scan.prodStart + i, 9).setValue(d.prov);
      }
    });
    var iData = inv.getRange(5, 1, 200, 6).getValues();
    iData.forEach(function(r, i) {
      if (String(r[0] || '').trim() === nom) {
        inv.getRange(i + 5, 6).setValue(Number(r[5]) + nC);
        inv.getRange(i + 5, 4).setValue(Math.round(cP * (Number(r[2]) || 1) * 100) / 100);
      }
    });
    var raw = PropertiesService.getScriptProperties().getProperty(CACHE_KEY);
    if (raw) {
      var c = JSON.parse(raw);
      c.prods.forEach(function(p) { if (p.nombre === nom) { p.costo = Math.round(cP * 10000) / 10000; if (d.nuevoPrecio) p.precio = Number(d.nuevoPrecio); } });
      PropertiesService.getScriptProperties().setProperty(CACHE_KEY, JSON.stringify(c));
    }
    return { ok: true, costoP: cP.toFixed(4), nuevoStock: tU, msg: nom + ': +' + nC + 'u. Costo prom $' + cP.toFixed(2) + '/u' };
  } catch (e) { return { ok: false, err: String(e) }; }
}

function getProdData(nombre) {
  try {
    var cat = _ss().getSheetByName('Catalogo');
    var inv = _ss().getSheetByName('Inventario');
    var scan = _scanCatalogoSheet(cat);
    var pData = cat.getRange(scan.prodStart, 1, scan.prodEnd - scan.prodStart + 1, 10).getValues();
    var found = null, catRow = -1;
    pData.forEach(function(r, i) { if (String(r[1] || '').trim() === nombre.trim()) { found = r; catRow = scan.prodStart + i; } });
    if (!found) return { ok: false, err: 'No encontrado' };
    var iData = inv.getRange(5, 1, 200, 11).getValues();
    var st = 0, al = 3, tp = 1;
    iData.forEach(function(r) { if (String(r[0] || '').trim() === nombre.trim()) { st = r[5] || 0; tp = r[2] || 1; al = r[10] || 3; } });
    return {
      ok: true, catRow: catRow, nombre: found[1], id: found[0], cat: found[2],
      precio: found[3], costo: found[4], notas: found[7] || '', prov: found[8] || '',
      archivado: String(found[9] || '').toLowerCase() === 'true',
      stock: st, tamPaq: tp, alertaMin: al
    };
  } catch (e) { return { ok: false, err: String(e) }; }
}

function guardarEdicion(d) {
  try {
    var cat = _ss().getSheetByName('Catalogo'), inv = _ss().getSheetByName('Inventario');
    var r = Number(d.catRow), nN = d.nombre.trim(), oN = d.nombreOld.trim();
    cat.getRange(r, 2).setValue(nN); cat.getRange(r, 3).setValue(d.cat);
    cat.getRange(r, 4).setValue(Number(d.precio)); cat.getRange(r, 5).setValue(Number(d.costo));
    cat.getRange(r, 8).setValue(d.notas || ''); cat.getRange(r, 9).setValue(d.prov || '');
    if (d.archivado !== undefined) cat.getRange(r, 10).setValue(d.archivado ? 'true' : 'false');
    var iData = inv.getRange(5, 1, 200, 11).getValues();
    iData.forEach(function(row, i) {
      if (String(row[0] || '').trim() === oN) {
        if (nN !== oN) inv.getRange(i + 5, 1).setValue(nN);
        inv.getRange(i + 5, 11).setValue(Number(d.alertaMin) || 3);
      }
    });
    var raw = PropertiesService.getScriptProperties().getProperty(CACHE_KEY);
    if (raw) {
      var c = JSON.parse(raw);
      c.prods.forEach(function(p) {
        if (p.nombre === oN) {
          p.nombre = nN; p.cat = d.cat;
          p.precio = Number(d.precio); p.costo = Number(d.costo);
          p.notas = d.notas || ''; p.prov = d.prov || '';
          if (d.archivado !== undefined) p.archivado = d.archivado;
        }
      });
      PropertiesService.getScriptProperties().setProperty(CACHE_KEY, JSON.stringify(c));
    }
    return { ok: true, msg: nN + ' actualizado.' };
  } catch (e) { return { ok: false, err: String(e) }; }
}

function agregarProducto(d) {
  try {
    var cat = _ss().getSheetByName('Catalogo'), inv = _ss().getSheetByName('Inventario');
    var scan = _scanCatalogoSheet(cat);
    var id = d.id.trim().toUpperCase(), nombre = d.nombre.trim();
    var existing = cat.getRange(scan.prodStart, 1, scan.prodEnd - scan.prodStart + 1, 1).getValues().flat();
    if (existing.indexOf(id) >= 0) return { ok: false, err: 'ID "' + id + '" ya existe.' };

    // Encontrar fila libre dentro del rango, o insertar antes del header de Promos
    var catRow = -1;
    for (var r = scan.prodStart; r <= scan.prodEnd; r++) {
      var v = cat.getRange(r, 2).getValue();
      if (!v || String(v).startsWith('\u2190')) { catRow = r; break; }
    }
    if (catRow < 0) {
      // Insertar una fila antes del marcador de Promos para expandir dinámicamente
      var promoHdr = scan.promoStart - 3;
      cat.insertRowBefore(promoHdr);
      catRow = promoHdr;
    }

    cat.getRange(catRow, 1).setValue(id);
    cat.getRange(catRow, 2).setValue(nombre);
    cat.getRange(catRow, 3).setValue(d.cat || 'Varios');
    cat.getRange(catRow, 4).setValue(Number(d.precio));
    cat.getRange(catRow, 5).setValue(Number(d.costo));
    cat.getRange(catRow, 6).setFormula('=D' + catRow + '-E' + catRow);
    cat.getRange(catRow, 7).setFormula('=IF(D' + catRow + '>0,(D' + catRow + '-E' + catRow + ')/D' + catRow + ',0)');
    cat.getRange(catRow, 8).setValue(d.notas || '');
    cat.getRange(catRow, 9).setValue(d.prov || '');
    cat.getRange(catRow, 10).setValue('false');

    // Inventario
    var totalRow = null;
    for (var ir = 5; ir <= inv.getLastRow(); ir++) {
      if (inv.getRange(ir, 1).getValue() === 'TOTAL INVENTARIO') { totalRow = ir; break; }
    }
    if (!totalRow) totalRow = inv.getLastRow() + 1;
    inv.insertRowBefore(totalRow);
    var nr = totalRow, tam = Number(d.tamPaquete) || 1, cp = Number(d.costoPaquete) || (Number(d.costo) * tam);
    inv.getRange(nr, 1).setValue(nombre); inv.getRange(nr, 2).setValue(d.cat || 'Varios');
    inv.getRange(nr, 3).setValue(tam); inv.getRange(nr, 4).setValue(Math.round(cp * 100) / 100);
    inv.getRange(nr, 5).setFormula('=IF(C' + nr + '>0,D' + nr + '/C' + nr + ',0)');
    inv.getRange(nr, 6).setValue(Number(d.stockInicial) || 0);
    inv.getRange(nr, 7).setFormula('=SUMIFS(Ventas!$F:$F,Ventas!$E:$E,A' + nr + ',Ventas!$C:$C,"Producto")');
    inv.getRange(nr, 9).setFormula('=G' + nr + '+H' + nr);
    inv.getRange(nr, 10).setFormula('=F' + nr + '-I' + nr);
    inv.getRange(nr, 11).setValue(Number(d.alertaMin) || 3);
    inv.getRange(nr, 12).setFormula('=IF(F' + nr + '>0,J' + nr + '/F' + nr + ',0)');
    inv.getRange(nr, 13).setFormula('=IF(J' + nr + '<=0,"AGOTADO",IF(L' + nr + '<=0.15,"URGENTE",IF(L' + nr + '<=0.4,"BAJO",IF(L' + nr + '<=0.7,"OK","BIEN"))))');
    inv.getRange(nr, 14).setFormula('=J' + nr + '*E' + nr);

    var raw = PropertiesService.getScriptProperties().getProperty(CACHE_KEY);
    if (raw) {
      var c = JSON.parse(raw);
      c.prods.push({ id: id, nombre: nombre, cat: d.cat || 'Varios', precio: Number(d.precio), costo: Number(d.costo), notas: d.notas || '', prov: d.prov || '', archivado: false });
      if (c.cats.indexOf(d.cat || 'Varios') < 0) c.cats.push(d.cat || 'Varios');
      PropertiesService.getScriptProperties().setProperty(CACHE_KEY, JSON.stringify(c));
    }
    return { ok: true, msg: nombre + ' agregado (fila ' + catRow + ').' };
  } catch (e) { return { ok: false, err: String(e) }; }
}

function guardarCliente_legacy_removido(d) {
  // Reemplazada por la versión arriba con columnas extendidas. Esta es solo un placeholder
  return { ok: false, msg: 'legacy' };
}

function _autoCliente(nom) {
  try {
    var sh = _ss().getSheetByName('Clientes');
    if (!sh) { guardarCliente({ nombre: nom, tel: '', notas: '' }); return; }
    var last = sh.getLastRow();
    if (last < 2) { guardarCliente({ nombre: nom, tel: '', notas: '' }); return; }
    var names = sh.getRange(2, 1, last - 1, 1).getValues().flat().map(function(v) { return String(v).trim().toLowerCase(); });
    if (names.indexOf(nom.toLowerCase()) < 0) guardarCliente({ nombre: nom, tel: '', notas: '' });
  } catch (e) {}
}

// ══════════════════════════════════════════════
//  API: PROMOS_ITEMS
// ══════════════════════════════════════════════
function obtenerPromoComponentes(promoId) {
  try {
    var sh = _ss().getSheetByName('Promos_Items');
    if (!sh || sh.getLastRow() < 2) return { ok: true, items: [] };
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    var items = data.filter(function(r) { return String(r[0]).trim() === promoId; })
      .map(function(r) { return { producto_id: String(r[1]).trim(), cantidad: Number(r[2]) || 1, notas: r[3] || '' }; });
    return { ok: true, items: items };
  } catch (e) { return { ok: false, err: String(e), items: [] }; }
}

function guardarPromoComponentes(promoId, items) {
  try {
    var sh = _ss().getSheetByName('Promos_Items');
    if (!sh) {
      sh = _ss().insertSheet('Promos_Items');
      sh.getRange(1, 1, 1, 4).setValues([['promo_id', 'producto_id', 'cantidad', 'notas']]);
    }
    // Borrar items existentes de esa promo
    if (sh.getLastRow() >= 2) {
      var data = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
      var keep = data.filter(function(r) { return String(r[0]).trim() !== promoId; });
      sh.getRange(2, 1, sh.getLastRow() - 1, 4).clearContent();
      if (keep.length) sh.getRange(2, 1, keep.length, 4).setValues(keep);
    }
    // Insertar nuevos
    if (items && items.length) {
      var rows = items.map(function(it) { return [promoId, it.producto_id, Number(it.cantidad) || 1, it.notas || '']; });
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
    }
    return { ok: true, msg: promoId + ': ' + (items || []).length + ' componentes guardados.' };
  } catch (e) { return { ok: false, err: String(e) }; }
}

function obtenerTodasPromosComponentes() {
  try {
    var sh = _ss().getSheetByName('Promos_Items');
    if (!sh || sh.getLastRow() < 2) return { ok: true, map: {} };
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    var map = {};
    data.forEach(function(r) {
      var pid = String(r[0]).trim();
      if (!pid) return;
      if (!map[pid]) map[pid] = [];
      map[pid].push({ producto_id: String(r[1]).trim(), cantidad: Number(r[2]) || 1, notas: r[3] || '' });
    });
    return { ok: true, map: map };
  } catch (e) { return { ok: false, err: String(e), map: {} }; }
}

// ══════════════════════════════════════════════
//  API: ENVÍO DE RECIBO POR CORREO
// ══════════════════════════════════════════════
function enviarRecibo(args) {
  try {
    var to = args.to;
    if (!to) return { ok: false, err: 'Falta destinatario' };
    var subject = args.subject || 'Recibo · Botiquín de Escritorio';
    var htmlBody = args.htmlBody || '';
    var plain = args.plainText || htmlBody.replace(/<[^>]+>/g, '');
    var attachments = [];
    if (args.pngDataUrl) {
      var b64 = args.pngDataUrl.split(',')[1];
      attachments.push(Utilities.newBlob(Utilities.base64Decode(b64), 'image/png', 'recibo-' + (args.folio || 'x') + '.png'));
    }
    if (args.pdfHtmlSource) {
      var pdfBlob = HtmlService.createHtmlOutput(args.pdfHtmlSource).getAs('application/pdf');
      pdfBlob.setName('recibo-' + (args.folio || 'x') + '.pdf');
      attachments.push(pdfBlob);
    }
    var nombre = obtenerConfiguracion().cfg.negocio_nombre || 'Botiquín de Escritorio';
    GmailApp.sendEmail(to, subject, plain, {
      htmlBody: htmlBody,
      attachments: attachments,
      name: nombre
    });
    return { ok: true };
  } catch (e) { return { ok: false, err: String(e) }; }
}

// ══════════════════════════════════════════════
//  MANTENIMIENTO (preservado del v5)
// ══════════════════════════════════════════════
function reconstruirFiados() {
  var ss = _ss(), sh = ss.getSheetByName('Ventas');
  var fs = ss.getSheetByName('Fiados') || ss.insertSheet('Fiados');
  fs.getDataRange().clearContent();
  fs.getRange(1, 1, 1, 4).setValues([['Cliente', 'Ventas Rows', 'Estado', 'Ultima']]);
  var data = sh.getDataRange().getValues(), idx = {};
  for (var i = 3; i < data.length; i++) {
    var art = String(data[i][4] || '').trim(); if (!art) continue;
    var est = String(data[i][12] || '').trim();
    if (est !== 'Fiado' && est !== 'Al rato te pago') continue;
    var cli = String(data[i][13] || 'Desconocido').trim() || 'Desconocido';
    if (!idx[cli]) idx[cli] = { rows: [], ultima: _fmt(data[i][0]) };
    idx[cli].rows.push(i + 1); idx[cli].ultima = _fmt(data[i][0]);
  }
  var entries = Object.keys(idx);
  if (entries.length) fs.getRange(2, 1, entries.length, 4).setValues(entries.map(function(k) { return [k, idx[k].rows.join(','), 'Pendiente', idx[k].ultima]; }));
  SpreadsheetApp.getUi().alert('Fiados reconstruidos: ' + entries.length + ' clientes.');
}

function archivarMes() {
  var ss = _ss(), ui = SpreadsheetApp.getUi(), tz = _tz();
  var sh = ss.getSheetByName('Ventas');
  var now = new Date(), prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var label = Utilities.formatDate(prev, tz, 'MMM-yyyy');
  if (ss.getSheetByName('Ventas_' + label)) { ui.alert('Ya existe Ventas_' + label); return; }
  var data = sh.getDataRange().getValues(), arch = [], keep = [];
  for (var i = 3; i < data.length; i++) {
    var f = data[i][0]; if (!f) continue;
    var d = f instanceof Date ? f : new Date(f);
    if (d.getMonth() === prev.getMonth() && d.getFullYear() === prev.getFullYear()) arch.push(data[i]); else keep.push(data[i]);
  }
  if (!arch.length) { ui.alert('Sin ventas de ' + label); return; }
  var ns = ss.insertSheet('Ventas_' + label);
  var cols = sh.getLastColumn();
  sh.getRange(3, 1, 1, cols).copyTo(ns.getRange(1, 1, 1, cols));
  ns.getRange(2, 1, arch.length, cols).setValues(arch);
  sh.getRange(4, 1, data.length - 3, cols).clearContent();
  if (keep.length) sh.getRange(4, 1, keep.length, cols).setValues(keep);
  ui.alert('Archivadas ' + arch.length + ' ventas de ' + label + '.');
}

// ══════════════════════════════════════════════
//  API v7: REPORTES AVANZADOS, COMPRAS, GRANEL
// ══════════════════════════════════════════════

/**
 * Reporte avanzado con filtros granulares.
 * opts: { inicio:'YYYY-MM-DD', fin:'YYYY-MM-DD', weekdays:[0..6], diaMes:[1..31] }
 * weekdays: array opcional (0=Domingo, 1=Lunes, ..., 6=Sábado)
 * diaMes: array opcional (1..31) — solo cuentan ventas con .getDate() en este array
 */
function getReporteAvanzado(opts) {
  try {
    opts = opts || {};
    var sh = _ss().getSheetByName('Ventas');
    var data = sh.getDataRange().getValues();
    var tz = _tz();
    var dIni = opts.inicio ? new Date(opts.inicio + 'T00:00:00') : null;
    var dFin = opts.fin ? new Date(opts.fin + 'T23:59:59') : null;
    var wd = (opts.weekdays && opts.weekdays.length) ? opts.weekdays : null;
    var dm = (opts.diaMes && opts.diaMes.length) ? opts.diaMes : null;

    var totalV = 0, totalG = 0, totalC = 0, fiados = 0;
    var porProd = {}, porMetodo = {}, porCat = {}, porDia = {}, porHora = {}, porWD = {};
    var foliosSet = {};

    for (var i = 3; i < data.length; i++) {
      var row = data[i];
      var art = String(row[4] || '').trim(); if (!art) continue;
      var fecha = row[0] instanceof Date ? row[0] : new Date(row[0]);
      if (isNaN(fecha)) continue;
      if (dIni && fecha < dIni) continue;
      if (dFin && fecha > dFin) continue;
      var dayOfWeek = fecha.getDay();
      if (wd && wd.indexOf(dayOfWeek) < 0) continue;
      if (dm && dm.indexOf(fecha.getDate()) < 0) continue;

      var v = Number(row[8]) || 0, c = Number(row[9]) || 0, g = Number(row[10]) || 0;
      var m = String(row[11] || 'Efectivo').trim();
      var cat = String(row[3] || '').trim();
      var e = String(row[12] || '').trim();
      var fStr = _fmt(fecha);
      var horaStr = String(row[1] || '').substring(0, 2);
      var hh = parseInt(horaStr) || 0;
      var folio = String(row[16] || ('row' + i));

      totalV += v; totalG += g; totalC += c;
      if (e === 'Fiado' || e === 'Al rato te pago') fiados += v;
      porMetodo[m] = (porMetodo[m] || 0) + v;
      if (cat) porCat[cat] = (porCat[cat] || 0) + v;
      if (!porProd[art]) porProd[art] = { cant: 0, total: 0, gan: 0 };
      porProd[art].cant += Number(row[5]) || 1;
      porProd[art].total += v;
      porProd[art].gan += g;
      if (!porDia[fStr]) porDia[fStr] = { v: 0, g: 0, tickets: {} };
      porDia[fStr].v += v; porDia[fStr].g += g;
      porDia[fStr].tickets[folio] = true;
      foliosSet[folio] = true;
      var hKey = dayOfWeek + '_' + hh;
      porHora[hKey] = (porHora[hKey] || 0) + v;
      porWD[dayOfWeek] = (porWD[dayOfWeek] || 0) + v;
    }

    var top = Object.keys(porProd)
      .map(function(k) { return { nombre: k, cant: porProd[k].cant, total: porProd[k].total, gan: porProd[k].gan }; })
      .sort(function(a, b) { return b.total - a.total; }).slice(0, 15);
    var diasArr = Object.keys(porDia).sort()
      .map(function(k) { return { fecha: k, v: porDia[k].v, g: porDia[k].g, tickets: Object.keys(porDia[k].tickets).length }; });
    var totalTickets = Object.keys(foliosSet).length;

    return {
      ok: true, totalV: totalV, totalG: totalG, totalC: totalC, fiados: fiados,
      margen: totalV > 0 ? (totalG / totalV * 100) : 0,
      top: top, porMetodo: porMetodo, porCat: porCat,
      dias: diasArr, porHora: porHora, porWD: porWD,
      totalTickets: totalTickets,
      ticketPromedio: totalTickets > 0 ? totalV / totalTickets : 0
    };
  } catch (e) { return { ok: false, err: String(e) }; }
}

/**
 * Lista de compras inteligente: inventario bajo + datos de ventas del mes actual,
 * agrupado por proveedor. Para tomar decisiones de qué priorizar.
 */
function getListaCompras() {
  try {
    var ss = _ss();
    var inv = ss.getSheetByName('Inventario');
    var cat = ss.getSheetByName('Catalogo');
    var ventas = ss.getSheetByName('Ventas');
    var tz = ss.getSpreadsheetTimeZone();
    var now = new Date();
    var mesActual = Utilities.formatDate(now, tz, 'MM/yyyy');

    // 1. Mapa proveedor por producto desde Catalogo
    var scan = _scanCatalogoSheet(cat);
    var provMap = {};
    var pData = cat.getRange(scan.prodStart, 1, scan.prodEnd - scan.prodStart + 1, 9).getValues();
    pData.forEach(function(r) {
      var nom = String(r[1] || '').trim();
      if (nom) provMap[nom] = String(r[8] || 'Sin proveedor').trim() || 'Sin proveedor';
    });

    // 2. Stats de ventas del mes por producto
    var ventasData = ventas.getDataRange().getValues();
    var statsProd = {};
    for (var i = 3; i < ventasData.length; i++) {
      var row = ventasData[i];
      var art = String(row[4] || '').trim(); if (!art) continue;
      var fStr = _fmt(row[0]);
      if (!fStr || fStr.length < 8 || fStr.substring(3) !== mesActual) continue;
      if (!statsProd[art]) statsProd[art] = { cant: 0, vend: 0, gan: 0 };
      statsProd[art].cant += Number(row[5]) || 1;
      statsProd[art].vend += Number(row[8]) || 0;
      statsProd[art].gan += Number(row[10]) || 0;
    }

    // 3. Inventario
    var invData = inv.getRange(5, 1, 200, 15).getValues();
    var porProv = {};
    var totalEstimado = 0;

    invData.forEach(function(row) {
      var nom = String(row[0] || '').trim();
      if (!nom || nom === 'TOTAL INVENTARIO') return;
      var stock = Number(row[9]) || 0;
      var alerta = Number(row[10]) || 3;
      var estado = String(row[12] || '').trim();
      var tam = Number(row[2]) || 1;
      var cp = Number(row[3]) || 0;
      var urgente = ['AGOTADO', 'URGENTE'].indexOf(estado) >= 0;
      var bajo = urgente || ['BAJO'].indexOf(estado) >= 0 || stock <= alerta;
      if (!bajo) return;

      var sug = Math.max(tam, Math.ceil((alerta * 3 - stock) / tam) * tam);
      var costoEst = cp * (sug / tam);
      totalEstimado += costoEst;

      var prov = provMap[nom] || 'Sin proveedor';
      if (!porProv[prov]) porProv[prov] = { items: [], totalCost: 0 };
      porProv[prov].items.push({
        nombre: nom,
        cat: String(row[1] || ''),
        stock: stock,
        alerta: alerta,
        sugerido: sug,
        tamPaq: tam,
        costoPaq: cp,
        costoEst: costoEst,
        estado: estado || 'BAJO',
        urgente: urgente,
        ventasMes: statsProd[nom] ? statsProd[nom].cant : 0,
        ingresoMes: statsProd[nom] ? statsProd[nom].vend : 0,
        gananciaMes: statsProd[nom] ? statsProd[nom].gan : 0
      });
      porProv[prov].totalCost += costoEst;
    });

    // Ordenar items dentro de cada proveedor: urgentes primero, luego por ventas del mes desc
    Object.keys(porProv).forEach(function(p) {
      porProv[p].items.sort(function(a, b) {
        if (a.urgente !== b.urgente) return b.urgente ? 1 : -1;
        return b.ingresoMes - a.ingresoMes;
      });
    });

    return {
      ok: true,
      mes: mesActual,
      proveedores: porProv,
      totalEstimado: totalEstimado,
      countItems: Object.keys(porProv).reduce(function(s, p) { return s + porProv[p].items.length; }, 0)
    };
  } catch (e) { return { ok: false, err: String(e) }; }
}

/**
 * Devuelve el catálogo completo para el módulo de tarjetas imprimibles.
 * Incluye productos activos + promos activas con sus precios actuales.
 */
function getCatalogoParaTarjetas() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(CACHE_KEY);
    if (!raw) return { ok: false, err: 'Ejecuta "Preparar datos" primero.' };
    var d = JSON.parse(raw);
    // Agrupar productos por categoría
    var porCat = {};
    d.prods.filter(function(p){return !p.archivado;}).forEach(function(p) {
      if (!porCat[p.cat]) porCat[p.cat] = [];
      porCat[p.cat].push({ nombre: p.nombre, precio: p.precio, id: p.id });
    });
    return {
      ok: true,
      categorias: Object.keys(porCat).sort().map(function(c) { return { nombre: c, items: porCat[c] }; }),
      promos: d.promos.map(function(p) { return { nombre: p.nombre, precio: p.precio, ahorro: p.ahorro, contenido: p.contenido, id: p.id }; })
    };
  } catch (e) { return { ok: false, err: String(e) }; }
}

function agregarPromo(d) {
  try {
    var cat = _ss().getSheetByName('Catalogo');
    var scan = _scanCatalogoSheet(cat);
    var id = String(d.id || '').trim().toUpperCase();
    var nombre = String(d.nombre || '').trim();
    var precio = Number(d.precio) || 0;
    if (!id || !nombre || !precio) return { ok: false, err: 'ID, Nombre y Precio son requeridos' };
    var existing = cat.getRange(scan.promoStart, 1, scan.promoEnd - scan.promoStart + 1, 1).getValues().flat();
    if (existing.indexOf(id) >= 0) return { ok: false, err: 'ID de promo "' + id + '" ya existe.' };
    // Buscar fila libre dentro del rango de promos
    var promoRow = -1;
    for (var r = scan.promoStart; r <= scan.promoEnd; r++) {
      var v = cat.getRange(r, 2).getValue();
      if (!v || String(v).startsWith('\u2190')) { promoRow = r; break; }
    }
    if (promoRow < 0) {
      // Insertar una fila antes del marcador de Categorías
      var catHdr = scan.catStart - 3;
      cat.insertRowBefore(catHdr);
      promoRow = catHdr;
    }
    cat.getRange(promoRow, 1).setValue(id);
    cat.getRange(promoRow, 2).setValue(nombre);
    cat.getRange(promoRow, 3).setValue(precio);
    cat.getRange(promoRow, 4).setValue(Number(d.ahorro) || 0);
    cat.getRange(promoRow, 5).setValue(String(d.contenido || ''));
    cat.getRange(promoRow, 6).setValue(Number(d.costo) || 0);

    // Actualizar cache
    var raw = PropertiesService.getScriptProperties().getProperty(CACHE_KEY);
    if (raw) {
      var c = JSON.parse(raw);
      c.promos = c.promos || [];
      c.promos.push({ id: id, nombre: nombre, precio: precio, ahorro: Number(d.ahorro) || 0, contenido: String(d.contenido || ''), costo: Number(d.costo) || 0 });
      PropertiesService.getScriptProperties().setProperty(CACHE_KEY, JSON.stringify(c));
    }
    return { ok: true, msg: nombre + ' agregada (fila ' + promoRow + ').' };
  } catch (e) { return { ok: false, err: String(e) }; }
}

/**
 * v7.2 — Actualiza una promo existente en Catálogo.
 * Encuentra la fila por ID y reescribe nombre/precio/ahorro/contenido/costo.
 * Refresca el cache para que el frontend vea los cambios sin "Preparar datos".
 */
function actualizarPromo(d) {
  try {
    var cat = _ss().getSheetByName('Catalogo');
    var scan = _scanCatalogoSheet(cat);
    var id = String(d.id || '').trim().toUpperCase();
    var nombre = String(d.nombre || '').trim();
    var precio = Number(d.precio) || 0;
    if (!id || !nombre || !precio) return { ok: false, err: 'ID, Nombre y Precio son requeridos' };

    var rng = cat.getRange(scan.promoStart, 1, scan.promoEnd - scan.promoStart + 1, 6).getValues();
    var promoRow = -1;
    for (var i = 0; i < rng.length; i++) {
      if (String(rng[i][0] || '').trim().toUpperCase() === id) { promoRow = scan.promoStart + i; break; }
    }
    if (promoRow < 0) return { ok: false, err: 'Promo "' + id + '" no encontrada' };

    cat.getRange(promoRow, 2).setValue(nombre);
    cat.getRange(promoRow, 3).setValue(precio);
    if (d.ahorro !== undefined && d.ahorro !== null) cat.getRange(promoRow, 4).setValue(Number(d.ahorro) || 0);
    if (d.contenido !== undefined && d.contenido !== null) cat.getRange(promoRow, 5).setValue(String(d.contenido));
    if (d.costo !== undefined && d.costo !== null) cat.getRange(promoRow, 6).setValue(Number(d.costo) || 0);

    // Refrescar cache (mismo patrón que guardarEdicion/procesarRestock)
    var raw = PropertiesService.getScriptProperties().getProperty(CACHE_KEY);
    if (raw) {
      var c = JSON.parse(raw);
      c.promos = c.promos || [];
      var found = false;
      c.promos.forEach(function(p) {
        if (String(p.id).toUpperCase() === id) {
          found = true;
          p.nombre = nombre;
          p.precio = precio;
          if (d.ahorro !== undefined && d.ahorro !== null) p.ahorro = Number(d.ahorro) || 0;
          if (d.contenido !== undefined && d.contenido !== null) p.contenido = String(d.contenido);
          if (d.costo !== undefined && d.costo !== null) p.costo = Number(d.costo) || 0;
        }
      });
      if (!found) {
        c.promos.push({
          id: id, nombre: nombre, precio: precio,
          ahorro: Number(d.ahorro) || 0,
          contenido: String(d.contenido || ''),
          costo: Number(d.costo) || 0
        });
      }
      PropertiesService.getScriptProperties().setProperty(CACHE_KEY, JSON.stringify(c));
    }
    return { ok: true, msg: nombre + ' actualizada (fila ' + promoRow + ').' };
  } catch (e) { return { ok: false, err: String(e) }; }
}

/**
 * v7.2 — Eliminar promo: borra fila de Catalogo + componentes en Promos_Items.
 */
function eliminarPromo(promoId) {
  try {
    var id = String(promoId || '').trim().toUpperCase();
    if (!id) return { ok: false, err: 'ID vacío' };
    var cat = _ss().getSheetByName('Catalogo');
    var scan = _scanCatalogoSheet(cat);
    var rng = cat.getRange(scan.promoStart, 1, scan.promoEnd - scan.promoStart + 1, 6).getValues();
    var promoRow = -1;
    for (var i = 0; i < rng.length; i++) {
      if (String(rng[i][0] || '').trim().toUpperCase() === id) { promoRow = scan.promoStart + i; break; }
    }
    if (promoRow < 0) return { ok: false, err: 'Promo no encontrada' };
    cat.getRange(promoRow, 1, 1, 6).clearContent();

    // Limpiar componentes
    var pi = _ss().getSheetByName('Promos_Items');
    if (pi && pi.getLastRow() >= 2) {
      var data = pi.getRange(2, 1, pi.getLastRow() - 1, 4).getValues();
      var keep = data.filter(function(r) { return String(r[0]).trim().toUpperCase() !== id; });
      pi.getRange(2, 1, pi.getLastRow() - 1, 4).clearContent();
      if (keep.length) pi.getRange(2, 1, keep.length, 4).setValues(keep);
    }

    // Cache
    var raw = PropertiesService.getScriptProperties().getProperty(CACHE_KEY);
    if (raw) {
      var c = JSON.parse(raw);
      c.promos = (c.promos || []).filter(function(p) { return String(p.id).toUpperCase() !== id; });
      PropertiesService.getScriptProperties().setProperty(CACHE_KEY, JSON.stringify(c));
    }
    return { ok: true, msg: 'Promo ' + id + ' eliminada.' };
  } catch (e) { return { ok: false, err: String(e) }; }
}

// ══════════════════════════════════════════════
//  v7.2 — AJUSTE DE INVENTARIO (stock negativo / ventas olvidadas)
// ══════════════════════════════════════════════

/**
 * Detecta productos con stock anómalo: negativo o muy bajo vs alerta.
 * Útil para abrir la página Ajustes y atacarlos directo.
 */
function getStockAnomalo() {
  try {
    var inv = _ss().getSheetByName('Inventario');
    var data = inv.getRange(5, 1, 200, 15).getValues();
    var negativos = [], otros = [];
    data.forEach(function(row, idx) {
      var nom = String(row[0] || '').trim();
      if (!nom || nom === 'TOTAL INVENTARIO') return;
      var stock = Number(row[9]) || 0;
      var inicial = Number(row[5]) || 0;
      var vendido = Number(row[6]) || 0;
      var consPromos = Number(row[7]) || 0;
      var item = {
        nombre: nom, fila: idx + 5,
        stock: Math.round(stock * 100) / 100,
        inicial: Math.round(inicial * 100) / 100,
        vendido: Math.round(vendido * 100) / 100,
        consPromos: Math.round(consPromos * 100) / 100
      };
      if (stock < 0) negativos.push(item);
    });
    negativos.sort(function(a, b) { return a.stock - b.stock; }); // más negativo primero
    return { ok: true, negativos: negativos };
  } catch (e) { return { ok: false, err: String(e), negativos: [] }; }
}

/**
 * Ajusta el inventario de un producto. Dos modos:
 *
 *   modo='venta_olvidada': registra una venta retroactiva de la diferencia que falta.
 *     Útil cuando vendiste sin registrar. Crea una fila en Ventas (con folio AJ-...)
 *     marcada como tal en columna O (Notas) y R (Snapshot_Source='ajuste').
 *     La fórmula de Inventario lo recoge automáticamente.
 *
 *   modo='ajuste_directo': cambia el Stock Inicial (col F de Inventario) para que
 *     Stock Final = conteoReal. Loguea en hoja "Ajustes_Inventario".
 *     Útil para mermas, regalos, productos caducados, robos, errores históricos.
 *
 * args: { nombre, conteoReal, modo, motivo, precio?, metodo? }
 */
function ajustarInventario(args) {
  try {
    var nombre = String(args.nombre || '').trim();
    var modo = String(args.modo || 'venta_olvidada');
    var conteoReal = Number(args.conteoReal);
    var motivo = String(args.motivo || '').trim();
    if (!nombre) return { ok: false, err: 'Falta nombre del producto' };
    if (isNaN(conteoReal) || conteoReal < 0) return { ok: false, err: 'Conteo real inválido' };

    var ss = _ss();
    var inv = ss.getSheetByName('Inventario');
    var cat = ss.getSheetByName('Catalogo');
    if (!inv || !cat) return { ok: false, err: 'Hojas faltantes' };

    // Localizar fila de inventario y datos actuales
    var iData = inv.getRange(5, 1, 200, 15).getValues();
    var invRow = -1, stockActual = 0, inicial = 0, vendido = 0, consPromos = 0;
    for (var i = 0; i < iData.length; i++) {
      if (String(iData[i][0] || '').trim() === nombre) {
        invRow = i + 5;
        stockActual = Number(iData[i][9]) || 0;
        inicial = Number(iData[i][5]) || 0;
        vendido = Number(iData[i][6]) || 0;
        consPromos = Number(iData[i][7]) || 0;
        break;
      }
    }
    if (invRow < 0) return { ok: false, err: 'Producto "' + nombre + '" no está en Inventario' };

    var diff = Math.round((conteoReal - stockActual) * 100) / 100;
    if (diff === 0) return { ok: false, err: 'El stock ya coincide con el conteo. No hay nada que ajustar.' };

    if (modo === 'venta_olvidada') {
      // diff debe ser negativo (vendiste sin registrar → stock real menor que el calculado)
      // Si el calculado es negativo (-3) y el conteo real es 0, diff = +3 ; eso significa
      // que ya vendiste esas 3 unidades pero el sistema no las tiene, así que falta
      // registrarlas. Reinterpretamos: cantidadAVender = stockActual - conteoReal (si > 0).
      var cantOlvidada = Math.round((stockActual - conteoReal) * 100) / 100;
      if (cantOlvidada <= 0) {
        return { ok: false, err: 'Para registrar ventas olvidadas, el conteo real debe ser MENOR al stock calculado (' + stockActual + ').' };
      }
      // Tomar precio del catálogo si no viene en args
      var precio = Number(args.precio);
      if (!precio || isNaN(precio)) {
        var scan = _scanCatalogoSheet(cat);
        var pData = cat.getRange(scan.prodStart, 1, scan.prodEnd - scan.prodStart + 1, 5).getValues();
        for (var j = 0; j < pData.length; j++) {
          if (String(pData[j][1] || '').trim() === nombre) { precio = Number(pData[j][3]) || 0; break; }
        }
      }
      if (!precio) return { ok: false, err: 'No se pudo determinar el precio del producto' };

      // Registrar usando la misma vía que registrarVenta, pero con notas explícitas.
      var notas = '[AJUSTE] Venta olvidada · conteo físico: ' + conteoReal + ' · ' + (motivo || 'sin motivo');
      var result = registrarVenta({
        items: [{ nombre: nombre, precio: precio, tipo: 'Producto', cantidad: cantOlvidada }],
        metodo: String(args.metodo || 'Efectivo'),
        estado: 'Pagado',
        cliente: '',
        notas: notas
      });
      if (!result.ok) return { ok: false, err: 'Error registrando venta: ' + result.err };

      _logAjusteInventario({
        fecha: new Date(), nombre: nombre, modo: 'venta_olvidada',
        stockAntes: stockActual, conteoReal: conteoReal, diff: -cantOlvidada,
        motivo: motivo, refFolio: result.folio, valorAjuste: cantOlvidada * precio
      });

      return {
        ok: true, modo: 'venta_olvidada',
        cantidad: cantOlvidada, precio: precio, total: cantOlvidada * precio,
        folio: result.folio,
        msg: 'Registradas ' + cantOlvidada + ' venta(s) olvidada(s) de ' + nombre + ' por $' + (cantOlvidada * precio).toFixed(2) + ' (folio ' + result.folio + ')'
      };
    }

    if (modo === 'ajuste_directo') {
      // Stock Final = Inicial - (Vendido + ConsPromos) ⇒ Inicial_nuevo = conteoReal + Vendido + ConsPromos
      var nuevoInicial = Math.round((conteoReal + vendido + consPromos) * 100) / 100;
      inv.getRange(invRow, 6).setValue(nuevoInicial);
      // Nota de auditoría en la celda
      var notaCelda = '[AJUSTE ' + Utilities.formatDate(new Date(), _tz(), 'dd/MM/yyyy HH:mm') + '] '
        + 'Stock antes=' + stockActual + ' → ' + conteoReal + '. Inicial: ' + inicial + ' → ' + nuevoInicial
        + (motivo ? '. Motivo: ' + motivo : '');
      try { inv.getRange(invRow, 6).setNote(notaCelda); } catch (e) {}

      _logAjusteInventario({
        fecha: new Date(), nombre: nombre, modo: 'ajuste_directo',
        stockAntes: stockActual, conteoReal: conteoReal, diff: diff,
        motivo: motivo, refFolio: '',
        inicialAntes: inicial, inicialDespues: nuevoInicial
      });

      return {
        ok: true, modo: 'ajuste_directo',
        diff: diff, inicialAntes: inicial, inicialDespues: nuevoInicial,
        msg: nombre + ': stock ' + stockActual + ' → ' + conteoReal + ' (Inicial ajustado: ' + inicial + ' → ' + nuevoInicial + ')'
      };
    }

    // v8: restock_olvidado — surtí pero olvidé registrar el lote, y seguí vendiendo.
    // Suma `diff` (>0) al stock inicial. Si llega un costoLote, recalcula costo promedio.
    if (modo === 'restock_olvidado') {
      if (diff <= 0) {
        return { ok: false, err: 'Para "surtí olvidado" el conteo debe ser MAYOR al stock calculado (' + stockActual + ').' };
      }
      var nuevoInicialR = Math.round((conteoReal + vendido + consPromos) * 100) / 100;
      inv.getRange(invRow, 6).setValue(nuevoInicialR);

      var costoLote = Number(args.costoLote) || 0;
      var costoCatNuevo = null, costoCatViejo = null;
      if (costoLote > 0) {
        // Buscar costo actual en Catálogo
        var scanR = _scanCatalogoSheet(cat);
        var pDataR = cat.getRange(scanR.prodStart, 1, scanR.prodEnd - scanR.prodStart + 1, 5).getValues();
        var catRowR = -1;
        for (var k = 0; k < pDataR.length; k++) {
          if (String(pDataR[k][1] || '').trim() === nombre) {
            catRowR = scanR.prodStart + k;
            costoCatViejo = Number(pDataR[k][4]) || 0;
            break;
          }
        }
        if (catRowR > 0) {
          var costoLoteUnit = costoLote / diff;
          var saR = Math.max(0, stockActual);
          var tuR = saR + diff;
          var cpR = saR > 0 ? (((saR * costoCatViejo) + (diff * costoLoteUnit)) / tuR) : costoLoteUnit;
          costoCatNuevo = Math.round(cpR * 100) / 100;
          cat.getRange(catRowR, 5).setValue(costoCatNuevo);
          // Invalidar cache
          var rawR = PropertiesService.getScriptProperties().getProperty(CACHE_KEY);
          if (rawR) {
            try {
              var cR = JSON.parse(rawR);
              cR.prods.forEach(function(p) { if (p.nombre === nombre) p.costo = costoCatNuevo; });
              PropertiesService.getScriptProperties().setProperty(CACHE_KEY, JSON.stringify(cR));
            } catch (e) {}
          }
        }
      }

      var notaR = '[RESTOCK OLVIDADO ' + Utilities.formatDate(new Date(), _tz(), 'dd/MM/yyyy HH:mm') + '] '
        + 'Stock antes=' + stockActual + ' → ' + conteoReal + ' (+' + diff + '). Inicial: ' + inicial + ' → ' + nuevoInicialR
        + (costoLote > 0 ? '. Costo lote: $' + costoLote.toFixed(2) + ' (unit: $' + (costoLote / diff).toFixed(2) + ' · prom nuevo: $' + costoCatNuevo + ')' : '')
        + (motivo ? '. Motivo: ' + motivo : '');
      try { inv.getRange(invRow, 6).setNote(notaR); } catch (e) {}

      _logAjusteInventario({
        fecha: new Date(), nombre: nombre, modo: 'restock_olvidado',
        stockAntes: stockActual, conteoReal: conteoReal, diff: diff,
        motivo: motivo, refFolio: '',
        inicialAntes: inicial, inicialDespues: nuevoInicialR
      });

      var msgR = nombre + ': +' + diff + ' uds al stock (Inicial ' + inicial + ' → ' + nuevoInicialR + ')';
      if (costoCatNuevo !== null) msgR += '. Costo: $' + costoCatViejo.toFixed(2) + ' → $' + costoCatNuevo.toFixed(2);
      return {
        ok: true, modo: 'restock_olvidado',
        diff: diff, inicialAntes: inicial, inicialDespues: nuevoInicialR,
        costoAntes: costoCatViejo, costoDespues: costoCatNuevo,
        msg: msgR
      };
    }

    return { ok: false, err: 'Modo inválido: ' + modo };
  } catch (e) { return { ok: false, err: String(e) + ' ' + (e.stack || '') }; }
}

function _logAjusteInventario(entry) {
  try {
    var ss = _ss();
    var sh = ss.getSheetByName('Ajustes_Inventario');
    if (!sh) {
      sh = ss.insertSheet('Ajustes_Inventario');
      sh.getRange(1, 1, 1, 10).setValues([['Fecha', 'Producto', 'Modo', 'Stock Antes', 'Conteo Real', 'Diferencia', 'Motivo', 'Folio Ref', 'Inicial Antes', 'Inicial Después']]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, 10).setBackground('#1E2A1E').setFontColor('#6DFF1A').setFontWeight('bold');
      try { sh.hideSheet(); } catch (e) {}
    }
    var row = [
      entry.fecha,
      entry.nombre,
      entry.modo,
      entry.stockAntes,
      entry.conteoReal,
      entry.diff,
      entry.motivo || '',
      entry.refFolio || '',
      entry.inicialAntes !== undefined ? entry.inicialAntes : '',
      entry.inicialDespues !== undefined ? entry.inicialDespues : ''
    ];
    sh.appendRow(row);
  } catch (e) { /* no-op: el log no debe bloquear el ajuste */ }
}

function getAjustesInventario(limit) {
  try {
    var sh = _ss().getSheetByName('Ajustes_Inventario');
    if (!sh || sh.getLastRow() < 2) return { ok: true, ajustes: [] };
    var n = sh.getLastRow() - 1;
    var data = sh.getRange(2, 1, n, 10).getValues();
    var rows = data.map(function(r) {
      return {
        fecha: r[0] instanceof Date ? _fmtDT(r[0]) : String(r[0] || ''),
        nombre: r[1], modo: r[2],
        stockAntes: r[3], conteoReal: r[4], diff: r[5],
        motivo: r[6], folio: r[7],
        inicialAntes: r[8], inicialDespues: r[9]
      };
    }).reverse();
    if (limit) rows = rows.slice(0, Number(limit));
    return { ok: true, ajustes: rows };
  } catch (e) { return { ok: false, err: String(e), ajustes: [] }; }
}

function ventaRapida() {
  var raw = PropertiesService.getScriptProperties().getProperty(CACHE_KEY);
  if (!raw) { SpreadsheetApp.getUi().alert('Ejecuta "Preparar datos" primero.'); return; }
  var data = JSON.parse(raw);
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Venta rápida', 'Nombre del artículo:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var q = resp.getResponseText().toLowerCase().trim();
  var all = data.prods.concat(data.promos.map(function(p) { return { nombre: p.nombre, precio: p.precio, cat: 'Promo' }; }));
  var found = all.filter(function(p) { return p.nombre.toLowerCase().indexOf(q) >= 0; });
  if (!found.length) { ui.alert('No encontrado: ' + q); return; }
  var prod = found[0];
  if (found.length > 1) {
    var lista = found.map(function(p, i) { return (i + 1) + '. ' + p.nombre + ' $' + p.precio; }).join('\n');
    var r2 = ui.prompt('Varios resultados', lista + '\n\nNúmero:', ui.ButtonSet.OK_CANCEL);
    if (r2.getSelectedButton() !== ui.Button.OK) return;
    prod = found[parseInt(r2.getResponseText()) - 1];
  }
  var r3 = ui.prompt('Cantidad', prod.nombre + ' ($' + prod.precio + ')\nCantidad:', ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() !== ui.Button.OK) return;
  var cant = parseInt(r3.getResponseText()) || 1;
  var result = registrarVenta({
    items: [{ nombre: prod.nombre, precio: prod.precio, tipo: prod.cat === 'Promo' ? 'Promo' : 'Producto', cantidad: cant }],
    metodo: 'Efectivo', estado: 'Pagado', cliente: '', notas: ''
  });
  ui.alert(result.ok ? 'Registrado: ' + prod.nombre + ' x' + cant + ' = $' + result.total.toFixed(2) : 'Error: ' + result.err);
}
