/**
 * MALEU — Apps Script v5
 * ─────────────────────────────────────────────────────────────
 * DEPLOY:
 *   Implementar → Nueva implementación → Aplicación web
 *   Ejecutar como: Yo · Acceso: Cualquier usuario
 *
 * SETUP (ejecutar UNA sola vez después de deployar):
 *   Correr la función setupSheets() desde el editor
 *
 * TRIGGER onEdit (instalar una sola vez):
 *   Activadores → Agregar → onEdit → Al editar
 * ─────────────────────────────────────────────────────────────
 */

const SS      = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById('1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY');
const BROWN   = '#3D1C0A';
const ORANGE  = '#F97035';
const CREAM   = '#E8DFC4';

// ════════════════════════════════════════════════════════════
//  CONFIG — Contadores globales persistentes (hoja oculta)
//  Fila 1=headers, Fila 2=H-, Fila 3=C-, Fila 4=OC-
//  Col A=Parámetro, Col B=Valor
// ════════════════════════════════════════════════════════════

/** Lee el siguiente número para un contador y lo incrementa atómicamente.
 *  @param {string} prefix - 'H-', 'C-' o 'OC-'
 *  @returns {string} El nuevo ID formateado (ej: 'H-128', 'OC-090')
 */
function _nextId(prefix) {
  const sh = SS.getSheetByName('Config');
  if (!sh) throw new Error('Hoja Config no encontrada. Ejecutá setupConfig().');

  // Mapeo prefix → fila en Config
  const rowMap = { 'H-': 2, 'C-': 3, 'OC-': 4, 'P-': 5, 'CF-': 6 };
  const row = rowMap[prefix];
  if (!row) throw new Error('Prefix desconocido: ' + prefix);

  const celda = sh.getRange(row, 2); // Col B = valor actual
  const actual = Number(celda.getValue()) || 0;
  const nuevo = actual + 1;
  celda.setValue(nuevo);

  const pad = prefix === 'OC-' ? 3 : 3;
  return prefix + String(nuevo).padStart(pad, '0');
}

/** Setup de la hoja Config (ejecutar UNA vez) */
function setupConfig() {
  let sh = SS.getSheetByName('Config');
  if (!sh) sh = SS.insertSheet('Config');

  // Estructura
  sh.getRange(1, 1, 1, 2).setValues([['Parámetro', 'Valor']])
    .setBackground(BROWN).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');

  // Contadores — inicializar con los máximos actuales de cada hoja
  const contadores = [
    ['Último H-', _scanMax('Home', 'B', /^H-(\d+)$/)],
    ['Último C-', _scanMax('Clubes', 'B', /^C-(\d+)$/)],
    ['Último OC-', _scanMax('Orden de Compra', 'A', /^OC-(\d+)$/)],
    ['Último P-', _scanMax('Pilar', 'B', /^P-(\d+)$/)],
    ['Último CF-', _scanMax('Capital Federal', 'B', /^CF-(\d+)$/)],
  ];
  sh.getRange(2, 1, contadores.length, 2).setValues(contadores);

  // Parámetros del negocio
  sh.getRange(6, 1, 1, 2).setValues([['Parámetro Negocio', 'Valor']])
    .setBackground('#E8DFC4').setFontColor(BROWN)
    .setFontWeight('bold').setFontSize(10);
  const params = [
    ['Envío fuera de Home ($)', 3000],
    ['Stock Crítico (umbral)', 3],
  ];
  sh.getRange(7, 1, params.length, 2).setValues(params);

  // Formato
  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 120);
  sh.getRange('B2:B4').setHorizontalAlignment('center').setFontWeight('bold').setFontSize(12);
  sh.setFrozenRows(1);

  // Ocultar la hoja
  sh.hideSheet();
  sh.setTabColor('#666666');

  SS.toast('Config creada. Contadores: H-' + contadores[0][1] + ', C-' + contadores[1][1] + ', OC-' + contadores[2][1], 'Setup Config', 6);
}

/** Escanea una columna buscando el máximo de un patrón regex */
function _scanMax(sheetName, col, regex) {
  const sh = SS.getSheetByName(sheetName);
  if (!sh) return 0;
  const data = sh.getRange(col + '2:' + col + sh.getLastRow()).getValues();
  let max = 0;
  data.forEach(function(row) {
    const match = String(row[0]).match(regex);
    if (match) { const n = parseInt(match[1], 10); if (n > max) max = n; }
  });
  return max;
}

// ════════════════════════════════════════════════════════════
//  KARDEX — Log inmutable de movimientos de stock
//  Cada cambio de stock deja una huella digital
// ════════════════════════════════════════════════════════════

/**
 * Registra un movimiento de stock en la hoja Kardex.
 * @param {string} abbr - Abreviatura del producto (ej: 'PMu')
 * @param {string} tipo - '+REC', '-SAL', '+DEV', '-AJU'
 * @param {number} qty - Cantidad movida (siempre positiva)
 * @param {number} stockAnterior - Stock antes del movimiento
 * @param {number} stockNuevo - Stock después del movimiento
 * @param {string} canal - 'Home', 'Clubes', 'OC', 'Depósito', 'Manual'
 * @param {string} referencia - ID del pedido/OC (ej: 'H-026', 'OC-037')
 */
function _logKardex(abbr, tipo, qty, stockAnterior, stockNuevo, canal, referencia) {
  const sh = SS.getSheetByName('Kardex');
  if (!sh) return; // Si no existe, silencio

  // Obtener nombre del producto desde Productos
  var nombreProd = abbr;
  var hProd = SS.getSheetByName('Productos');
  if (hProd) {
    var prodData = hProd.getDataRange().getValues();
    for (var r = 1; r < prodData.length; r++) {
      if (String(prodData[r][2]).trim() === abbr) {
        nombreProd = String(prodData[r][1]).trim();
        break;
      }
    }
  }

  // Timestamp Argentina
  var ahora   = new Date();
  var argDate = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var dd = String(argDate.getDate()).padStart(2, '0');
  var mm = String(argDate.getMonth() + 1).padStart(2, '0');
  var yyyy = argDate.getFullYear();
  var hh = String(argDate.getHours()).padStart(2, '0');
  var mi = String(argDate.getMinutes()).padStart(2, '0');

  sh.appendRow([
    dd + '/' + mm + '/' + yyyy + ' ' + hh + ':' + mi,  // A  Fecha/Hora
    nombreProd,                                           // B  Producto
    abbr,                                                 // C  Abreviatura
    tipo,                                                 // D  Tipo
    qty,                                                  // E  Cantidad
    stockAnterior,                                        // F  Stock Anterior
    stockNuevo,                                           // G  Stock Nuevo
    canal,                                                // H  Canal
    referencia || '',                                     // I  Referencia
  ]);
}

/** Setup de la hoja Kardex (ejecutar UNA vez) */
function setupKardex() {
  let sh = SS.getSheetByName('Kardex');
  if (!sh) sh = SS.insertSheet('Kardex');

  const headers = [
    'Fecha/Hora',    // A
    'Producto',      // B
    'Abreviatura',   // C
    'Tipo',          // D
    'Cantidad',      // E
    'Stock Anterior',// F
    'Stock Nuevo',   // G
    'Canal',         // H
    'Referencia',    // I
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(BROWN).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.setFrozenRows(1);
  sh.setRowHeight(1, 40);

  // Anchos
  [140, 200, 80, 65, 75, 100, 100, 90, 100].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // Centrar
  sh.getRange('C2:I5000').setHorizontalAlignment('center');

  // Conditional formatting — Tipo
  const dRange = sh.getRange('D2:D5000');
  const rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith('+')
    .setBackground('#C8E6C9').setFontColor('#1B5E20').setBold(true)
    .setRanges([dRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith('-')
    .setBackground('#FFCDD2').setFontColor('#B71C1C').setBold(true)
    .setRanges([dRange]).build());
  sh.setConditionalFormatRules(rules);

  // Banding
  try { sh.getRange('A2:I5000').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false); }
  catch(ex) {}

  sh.setTabColor('#7B1FA2'); // violeta — distinguir de operativas
}

// ════════════════════════════════════════════════════════════
//  ARCHIVO HISTÓRICO — reemplazo del limpiado manual
//  Mueve pedidos cerrados a hojas de archivo
// ════════════════════════════════════════════════════════════

/** Archiva pedidos Entregados/Cancelados de Home y Clubes.
 *  Solo mueve filas cerradas, las pendientes/reservadas permanecen. */
function archivarSemana() {
  var ui = SpreadsheetApp.getUi();

  // ── Contar filas archivables ──
  var shHome   = SS.getSheetByName('Home');
  var shPilar  = SS.getSheetByName('Pilar');
  var shCaba   = SS.getSheetByName('Capital Federal');
  var shClubes = SS.getSheetByName('Clubes');
  var archHome = 0, archPilar = 0, archCaba = 0, archClubes = 0;

  function _countArchivable(sh, colEstado) {
    if (!sh || sh.getLastRow() <= 1) return 0;
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    var count = 0;
    data.forEach(function(row) {
      var estado = String(row[colEstado]).trim();
      if (estado === 'Entregado' || estado === 'Cancelado') count++;
    });
    return count;
  }

  archHome   = _countArchivable(shHome, 10);
  archPilar  = _countArchivable(shPilar, 10);
  archCaba   = _countArchivable(shCaba, 10);
  archClubes = _countArchivable(shClubes, 13);

  if (archHome === 0 && archPilar === 0 && archCaba === 0 && archClubes === 0) {
    ui.alert('No hay pedidos para archivar', 'No se encontraron pedidos en estado "Entregado" o "Cancelado".', ui.ButtonSet.OK);
    return;
  }

  // ── Confirmar ──
  var resp = ui.alert(
    'Archivar semana',
    'Se van a mover:\n\n' +
    '  Home: ' + archHome + ' pedido(s)\n' +
    '  Pilar: ' + archPilar + ' pedido(s)\n' +
    '  Capital Federal: ' + archCaba + ' pedido(s)\n' +
    '  Clubes: ' + archClubes + ' pedido(s)\n\n' +
    'Las filas se copian a las hojas de archivo y se eliminan de las hojas operativas.\n' +
    'Los pedidos Pendientes y Reservados NO se tocan.\n\n' +
    '¿Confirmar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var totalMovidas = 0;

  // ── Archivar Home ──
  if (shHome && archHome > 0) {
    totalMovidas += _archivarHoja(shHome, 'Archivo Home', 10);
  }

  // ── Archivar Pilar ──
  if (shPilar && archPilar > 0) {
    totalMovidas += _archivarHoja(shPilar, 'Archivo Pilar', 10);
  }

  // ── Archivar Capital Federal ──
  if (shCaba && archCaba > 0) {
    totalMovidas += _archivarHoja(shCaba, 'Archivo Capital Federal', 10);
  }

  // ── Archivar Clubes ──
  if (shClubes && archClubes > 0) {
    totalMovidas += _archivarHoja(shClubes, 'Archivo Clubes', 13);
  }

  SS.toast(totalMovidas + ' pedidos archivados correctamente.', 'Archivo completado', 6);
}

/**
 * Mueve filas Entregado/Cancelado de una hoja operativa a su archivo.
 * @param {Sheet} shOrigen - Hoja operativa (Home o Clubes)
 * @param {string} archivoName - Nombre de la hoja de archivo
 * @param {number} colEstado - Índice 0-based de la columna Estado de Entrega
 * @returns {number} Cantidad de filas movidas
 */
function _archivarHoja(shOrigen, archivoName, colEstado) {
  // Crear hoja archivo si no existe (mismos headers que la operativa)
  var shArchivo = SS.getSheetByName(archivoName);
  if (!shArchivo) {
    shArchivo = SS.insertSheet(archivoName);
    // Copiar headers
    var headers = shOrigen.getRange(1, 1, 1, shOrigen.getLastColumn()).getValues();
    shArchivo.getRange(1, 1, 1, headers[0].length).setValues(headers)
      .setBackground(BROWN).setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center');
    shArchivo.setFrozenRows(1);
    shArchivo.setTabColor('#666666');
  }

  var lastRow = shOrigen.getLastRow();
  if (lastRow <= 1) return 0;

  var data = shOrigen.getRange(2, 1, lastRow - 1, shOrigen.getLastColumn()).getValues();
  var filasArchivar = [];
  var filasEliminar = []; // índices de fila 1-based para borrar

  for (var i = 0; i < data.length; i++) {
    var estado = String(data[i][colEstado]).trim();
    if (estado === 'Entregado' || estado === 'Cancelado') {
      filasArchivar.push(data[i]);
      filasEliminar.push(i + 2); // +2 porque fila 1 = header, i es 0-based
    }
  }

  if (filasArchivar.length === 0) return 0;

  // Copiar al archivo
  var destRow = shArchivo.getLastRow() + 1;
  shArchivo.getRange(destRow, 1, filasArchivar.length, filasArchivar[0].length).setValues(filasArchivar);

  // Eliminar de la hoja operativa (de abajo hacia arriba para no romper índices)
  filasEliminar.reverse().forEach(function(rowNum) {
    shOrigen.deleteRow(rowNum);
  });

  return filasArchivar.length;
}

// ════════════════════════════════════════════════════════════
//  doGet — lectura de datos (compras)
// ════════════════════════════════════════════════════════════
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'compras') return _doGetCompras();
  if (action === 'egresos') return _doGetEgresos();
  return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
}

function _doGetCompras() {
  const sh = SS.getSheetByName('Pedidos_Proveedores');
  if (!sh) {
    return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON);
  }
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) {
    return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON);
  }
  const headers = data[0];
  const rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] instanceof Date
      ? Utilities.formatDate(row[i], 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy')
      : (row[i] !== undefined && row[i] !== null ? String(row[i]) : ''); });
    return obj;
  });
  return ContentService
    .createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

function _doGetEgresos() {
  const sh = SS.getSheetByName('Egresos');
  if (!sh) return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON);
  const headers = data[0];
  const rows = data.slice(1).filter(row => row[0]).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i] instanceof Date
        ? Utilities.formatDate(row[i], 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy')
        : (row[i] !== undefined && row[i] !== null ? String(row[i]) : '');
    });
    return obj;
  });
  return ContentService.createTextOutput(JSON.stringify(rows)).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  doPost — recibe pedidos desde la página + acciones internas
// ════════════════════════════════════════════════════════════
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'compra')       return _doPostCompra(data);
    if (data.action === 'compraLote')  return _doPostCompraLote(data);
    if (data.action === 'updateCompra') return _doUpdateCompra(data);
    return _doPostPedido(data);
  } catch(err) {
    // LOG DE ERROR: guardar pedido fallido para no perderlo jamás
    _logPedidoFallido(e, err);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Guarda pedidos fallidos en hoja "Log Errores" para recuperación.
 * NINGÚN pedido se pierde — si falla el procesamiento, queda registrado acá.
 */
function _logPedidoFallido(e, err) {
  try {
    var sh = SS.getSheetByName('Log Errores');
    if (!sh) {
      sh = SS.insertSheet('Log Errores');
      sh.getRange(1, 1, 1, 6).setValues([['Fecha/Hora', 'Error', 'Canal', 'Cliente', 'Teléfono', 'Data Completa']])
        .setBackground(BROWN).setFontColor('#FFFFFF').setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.setTabColor('#FF0000');
    }
    var rawData = '';
    var canal = '', cliente = '', telefono = '';
    try {
      rawData = e.postData.contents;
      var parsed = JSON.parse(rawData);
      canal = parsed.canal || '';
      cliente = parsed.nombre || '';
      telefono = parsed.telefono || '';
    } catch(ex) { rawData = String(e.postData && e.postData.contents || 'sin datos'); }

    var ahora = new Date();
    var argDate = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
    var timestamp = Utilities.formatDate(argDate, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm:ss');

    sh.appendRow([timestamp, err.message, canal, cliente, telefono, rawData]);
  } catch(logErr) {
    Logger.log('ERROR CRÍTICO en _logPedidoFallido: ' + logErr.message);
  }
}

function _doPostPedido(data) {
    const canal = String(data.canal || 'Home');
    if (canal === 'Clubes')          _doPostClubes(data);
    else if (canal === 'Pilar')      _doPostHome(data, 'Pilar', 'P');
    else if (canal === 'Capital Federal') _doPostHome(data, 'Capital Federal', 'CF');
    else                             _doPostHome(data, 'Home', 'H');

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  _doPostHome — escribe en la hoja "Home" (canal Estancias)
//  Columnas A–AP según el esquema de análisis de Tadeo
// ════════════════════════════════════════════════════════════

// Mapeo id de producto (página web) → columna 1-based de la hoja Home
// Con col A=Hora, Envío en O (15), productos empiezan en col U (21) hasta AM (39)
const HOME_PRODUCT_COLS = {
  5:  21,  // PPM   — Pack Muzarella x2
  6:  22,  // PPJyQ — Pack Jamón y Queso x2
  7:  23,  // PPCyQ — Pack Cebolla y Queso x2
  8:  24,  // SCo   — Sorrentinos Cordero al Malbec
  9:  25,  // SJyQ  — Sorrentinos Jamón y Queso
  10: 26,  // SCa   — Sorrentinos Calabaza y Queso
  11: 27,  // ECaC  — Empanadas Carne a Cuchillo x8
  12: 28,  // EJyQ  — Empanadas Jamón y Queso x8
  17: 29,  // ECyQ  — Empanadas Cebolla y Queso x8
  18: 30,  // EV    — Empanadas Verdura x8
  14: 31,  // TG    — Torta Golosa
  15: 32,  // TLC   — Torta Lemon Crumble
  16: 33,  // TC    — Torta Coco
  13: 34,  // F     — Franui Leche
  19: 35,  // PMu   — Pizza Muzzarella
  1:  36,  // PMa   — Pizza Margarita
  2:  37,  // PJyQ  — Pizza Jamón y Queso
  3:  38,  // PCC   — Pizza Cebolla Caramelizada
  4:  39,  // PJyM  — Pizza Jamón y Morrón
};

// Mapeo id de producto (página web) → abreviatura en hoja Productos (col C)
const PAGE_ID_TO_ABBR = {
  5:  'PPM',   6:  'PPJyQ', 7:  'PPCyQ',
  8:  'SCo',   9:  'SJyQ',  10: 'SCa',
  11: 'ECaC',  12: 'EJyQ',  17: 'ECyQ', 18: 'EV',
  14: 'TG',    15: 'TLC',   16: 'TC',   13: 'F',
  19: 'PMu',  1:  'PMa',  2:  'PJyQ',  3:  'PCC',  4:  'PJyM',
};

function _doPostHome(data, sheetName, prefix) {
  sheetName = sheetName || 'Home';
  prefix = prefix || 'H';
  const sh = SS.getSheetByName(sheetName);
  if (!sh) return;

  // ── N° de pedido desde contador global (Config) ──
  const orderNum = _nextId(prefix + '-');

  // ── Fecha y hora en zona horaria Argentina ────────────────
  const ahora   = new Date();
  const argDate = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));

  const DIAS      = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
  const diaNombre = DIAS[argDate.getDay()];
  const dd        = String(argDate.getDate()).padStart(2, '0');
  const mm        = String(argDate.getMonth() + 1).padStart(2, '0');
  const yyyy      = argDate.getFullYear();
  const fechaStr  = dd + '/' + mm + '/' + yyyy;
  const horaStr   = String(argDate.getHours()).padStart(2, '0') + ':' + String(argDate.getMinutes()).padStart(2, '0');
  const mes       = argDate.getMonth() + 1;
  const semana    = _isoWeek(argDate);

  // ── Pago ─────────────────────────────────────────────────
  const envio           = Number(data.envio) || 0;
  const total           = Number(data.total) || 0;
  const pago            = String(data.pago || '');
  const totalSinEnvio   = total - envio;
  const efectivo        = pago === 'Efectivo'      ? totalSinEnvio : 0;
  const transferencia   = pago === 'Transferencia' ? totalSinEnvio : 0;

  // ── Cantidades de productos por id ────────────────────────
  const qtys = {};
  (data.items || []).forEach(function(item) {
    qtys[Number(item.id)] = Number(item.qty) || 0;
  });

  // ── Costo desde hoja Productos (col C=Abreviatura, col H=Costo) ──
  let costoTotal = 0;
  const hProductos = SS.getSheetByName('Productos');
  if (hProductos) {
    const prodData = hProductos.getDataRange().getValues();
    (data.items || []).forEach(function(item) {
      const abbr = PAGE_ID_TO_ABBR[Number(item.id)];
      if (!abbr) return;
      for (let r = 1; r < prodData.length; r++) {
        if (String(prodData[r][2]).trim() === abbr) { // col C = Abreviatura
          const costoUnit = Number(prodData[r][9]) || 0; // col J = Costo
          costoTotal += costoUnit * (Number(item.qty) || 0);
          break;
        }
      }
    });
  }

  // ── Barrio y Sub Barrio ───────────────────────────────────
  const barrioPrivado = String(data.barrioPrivado || data.barrio || '');
  const subBarrio     = String(data.subBarrio     || '');

  // ── Construir fila de 48 columnas (A a AV) ────────────────
  const row = new Array(50).fill('');
  const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

  row[0]  = horaStr;                            // A  Hora Pedido
  row[1]  = orderNum;                           // B  N° Pedido
  row[2]  = diaNombre;                          // C  Día Pedido
  row[3]  = fechaStr;                           // D  Fecha Pedido
  row[4]  = MESES[mes - 1];                      // E  Mes Pedido (nombre)
  row[5]  = semana;                             // F  Semana Pedido
  row[6]  = yyyy;                               // G  Año Pedido
  row[7]  = String(data.nombre || '');          // H  Cliente
  row[8]  = 'Pendiente';                        // I  Origen (default)
  row[9]  = String(data.dia || '');             // J  Día de entrega elegido
  row[10] = 'Pendiente';                        // K  Estado de Entrega (default)
  row[11] = pago;                               // L  Forma de Pago
  row[12] = 'No Cobrado';                       // M  Estado de Pago (default)
  row[13] = total;                              // N  Total ($)
  row[14] = envio;                              // O  Envío ($)
  row[15] = efectivo;                           // P  Efectivo ($)
  row[16] = transferencia;                      // Q  Transferencia ($)
  row[17] = 0;                                  // R  Propina Efectivo (default $0)
  row[18] = 0;                                  // S  Propina Transferencia (default $0)

  // Productos: cols U–AM (índices 20–38 en base-0)
  Object.keys(HOME_PRODUCT_COLS).forEach(function(id) {
    row[HOME_PRODUCT_COLS[id] - 1] = qtys[Number(id)] || 0;
  });

  // T (col 20) = Facturado → se pone fórmula DESPUÉS del appendRow
  row[39] = costoTotal;                         // Costo
  row[40] = total - costoTotal;                 // Margen Bruto

  if (sheetName === 'Capital Federal') {
    // Capital Federal: AP=Barrio, AQ=Calle, AR=Número, AS=Piso, AT=Teléfono (52 cols)
    row.push('', '');
    row[41] = String(data.barrioCaba || '');   // AP  Barrio
    row[42] = String(data.calle || '');        // AQ  Calle
    row[43] = String(data.numero || '');       // AR  Número
    row[44] = String(data.piso || '');         // AS  Piso
    row[45] = String(data.telefono || '');     // AT  Teléfono
  } else if (sheetName === 'Pilar') {
    // Pilar: AP=Dirección, AQ=Lote/Piso, AR=Teléfono (50 cols)
    row[41] = String(data.barrio || data.direccion || '');  // AP  Dirección
    row[42] = String(data.lote || '');                       // AQ  Lote / Piso
    row[43] = String(data.telefono || '');                   // AR  Teléfono
  } else {
    // Home: AP=Barrio, AQ=Sub Barrio, AR=Domicilio-Lote, AS=Teléfono (51 cols)
    row.push('');
    row[41] = barrioPrivado;                      // AP  Barrio
    row[42] = subBarrio;                          // AQ  Sub Barrio
    row[43] = String(data.lote || '');            // AR  Domicilio - Lote
    row[44] = String(data.telefono || '');        // AS  Teléfono
  }

  sh.appendRow(row);
  // Fórmula Facturado en T (col 20) = Total + Propinas
  var newRow = sh.getLastRow();
  sh.getRange(newRow, 20).setFormula('=N' + newRow + '+R' + newRow + '+S' + newRow);

}

// ════════════════════════════════════════════════════════════
//  _doPostClubes — escribe en la hoja "Clubes"
//  Columnas: Hora, N°Pedido, Día, Fecha, Mes, Semana, Año,
//  Cliente, Club, Deporte, Grupo, Origen, Día Entrega,
//  Estado Entrega, Forma Pago, Estado Pago, Total, Efectivo,
//  Transferencia, Propina Ef, Propina Tr,
//  PMu, PMa, PJyQ, PCC, PJyM, PPM, PPJyQ, PPCyQ,
//  Costo, Margen Bruto, Teléfono
// ════════════════════════════════════════════════════════════

// Mapeo ID producto (web clubes) → columna 1-based hoja Clubes
// Columna V(22)=Facturado, productos empiezan en W(23)
const CLUBES_PRODUCT_COLS = {
  'pmu': 23,  // X  — PMu  — Pizza Muzzarella
  'pma': 24,  // Y  — PMa  — Pizza Margarita
  'pjq': 25,  // Z  — PJyQ — Pizza Jamón y Queso
  'pcc': 26,  // AA — PCC  — Pizza Cebolla Caramelizada
  'pjm': 27,  // AB — PJyM — Pizza Jamón y Morrón
  'pp1': 28,  // AC — PPM  — Pack Muzarella x2
  'pp2': 29,  // AD — PPJyQ — Pack Jamón y Queso x2
  'pp3': 30,  // AE — PPCyQ — Pack Cebolla y Queso x2
};

// Mapeo ID producto (web clubes) → abreviatura en hoja Productos
const CLUBES_ID_TO_ABBR = {
  'pmu':'PMu', 'pma':'PMa', 'pjq':'PJyQ', 'pcc':'PCC', 'pjm':'PJyM',
  'pp1':'PPM', 'pp2':'PPJyQ', 'pp3':'PPCyQ',
};

function _doPostClubes(data) {
  const sh = SS.getSheetByName('Clubes');
  if (!sh) return;

  // ── N° de pedido desde contador global (Config) ──
  const orderNum = _nextId('C-');

  // Fecha y hora Argentina
  const ahora   = new Date();
  const argDate = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  const DIAS    = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
  const MESES   = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  const diaNombre = DIAS[argDate.getDay()];
  const dd   = String(argDate.getDate()).padStart(2, '0');
  const mm   = String(argDate.getMonth() + 1).padStart(2, '0');
  const yyyy = argDate.getFullYear();
  const fechaStr = dd + '/' + mm + '/' + yyyy;
  const horaStr  = String(argDate.getHours()).padStart(2, '0') + ':' + String(argDate.getMinutes()).padStart(2, '0');
  const mes    = MESES[argDate.getMonth()];
  const semana = _isoWeek(argDate);

  // Pago
  const envio           = Number(data.envio) || 0;
  const total           = Number(data.total) || 0;
  const pago            = String(data.pago || '');
  const totalSinEnvio   = total - envio;
  const efectivo        = pago === 'Efectivo'      ? totalSinEnvio : 0;
  const transferencia   = pago === 'Transferencia' ? totalSinEnvio : 0;

  // Cantidades
  const qtys = {};
  (data.items || []).forEach(function(item) {
    qtys[String(item.id)] = Number(item.qty) || 0;
  });

  // Costo desde hoja Productos
  let costoTotal = 0;
  const hProductos = SS.getSheetByName('Productos');
  if (hProductos) {
    const prodData = hProductos.getDataRange().getValues();
    (data.items || []).forEach(function(item) {
      const abbr = CLUBES_ID_TO_ABBR[String(item.id)];
      if (!abbr) return;
      for (let r = 1; r < prodData.length; r++) {
        if (String(prodData[r][2]).trim() === abbr) {
          const costoUnit = Number(String(prodData[r][9]).replace(/[$.]/g,'').replace(/,/g,'')) || 0;
          costoTotal += costoUnit * (Number(item.qty) || 0);
          break;
        }
      }
    });
  }

  // Construir fila de 34 columnas (A a AH) — +1 por col Envío
  const row = new Array(34).fill('');
  row[0]  = horaStr;                           // A  Hora
  row[1]  = orderNum;                          // B  N° Pedido
  row[2]  = diaNombre;                         // C  Día
  row[3]  = fechaStr;                          // D  Fecha
  row[4]  = mes;                               // E  Mes
  row[5]  = semana;                            // F  Semana
  row[6]  = yyyy;                              // G  Año
  row[7]  = String(data.nombre || '');         // H  Cliente
  row[8]  = String(data.club || '');           // I  Club
  row[9]  = String(data.deporte || '');        // J  Deporte
  row[10] = String(data.grupo || '');          // K  Grupo
  row[11] = 'Pendiente';                       // L  Origen
  row[12] = String(data.dia || '');            // M  Día de Entrega
  row[13] = 'Pendiente';                       // N  Estado de Entrega
  row[14] = pago;                              // O  Forma de Pago
  row[15] = 'No Cobrado';                      // P  Estado de Pago
  row[16] = total;                             // Q  Total ($)
  row[17] = envio;                             // R  Envío ($)
  row[18] = efectivo;                          // S  Efectivo
  row[19] = transferencia;                     // T  Transferencia
  row[20] = 0;                                 // U  Propina Efectivo
  row[21] = 0;                                 // V  Propina Transferencia
  // W  Facturado → fórmula se pone después del appendRow

  // Productos: cols X–AE (índices 23–30 en base-0)
  Object.keys(CLUBES_PRODUCT_COLS).forEach(function(id) {
    row[CLUBES_PRODUCT_COLS[id]] = qtys[id] || 0;
  });

  row[31] = costoTotal;                        // AF  Costo
  row[32] = total - costoTotal;                // AG  Margen Bruto
  row[33] = String(data.telefono || '');       // AH  Teléfono

  sh.appendRow(row);
  // Fórmula Facturado en W (col 23) = Total + Propinas
  var newRow = sh.getLastRow();
  sh.getRange(newRow, 23).setFormula('=Q' + newRow + '+U' + newRow + '+V' + newRow);
}

// Número de semana ISO (lunes = primer día de la semana)
function _isoWeek(date) {
  const d      = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

function _doPostCompra(data) {
  const sh = SS.getSheetByName('Pedidos_Proveedores');
  if (!sh) throw new Error('Hoja Pedidos_Proveedores no encontrada. Ejecutá setupSheets().');

  const ahora    = new Date();
  const id       = 'C-' + ahora.getTime();
  const fecha    = Utilities.formatDate(ahora, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');
  const cantidad = parseFloat(data.cantidad) || 0;
  const precio   = parseFloat(data.precio)   || 0;
  const total    = precio ? cantidad * precio : 0;

  sh.appendRow([
    id,
    fecha,
    data.proveedor    || '',
    data.producto     || '',
    cantidad,
    data.unidad       || '',
    precio            || '',
    total             || '',
    'Pendiente',
    data.notas        || '',
    data.fecha        || ''   // fecha entrega estimada (yyyy-MM-dd o vacío)
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, id }))
    .setMimeType(ContentService.MimeType.JSON);
}

function _doPostCompraLote(data) {
  const sh = SS.getSheetByName('Pedidos_Proveedores');
  if (!sh) throw new Error('Hoja Pedidos_Proveedores no encontrada. Ejecutá setupSheets().');

  const ahora     = new Date();
  const loteId    = 'C-' + ahora.getTime();
  const fecha     = Utilities.formatDate(ahora, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');
  const items     = data.items || [];

  const costoTotal = parseFloat(data.costoTotal) || 0;
  items.forEach((item, i) => {
    const id    = loteId + (items.length > 1 ? '-' + (i + 1) : '');
    const qty   = parseFloat(item.cantidad) || 0;
    const total = i === 0 && costoTotal > 0 ? costoTotal : ''; // costo del lote en primera fila
    sh.appendRow([
      id,
      fecha,
      data.proveedor || '',
      item.producto  || '',
      qty,
      item.unidad    || '',
      '',            // Precio unit.
      total,         // Total (costo lote, solo fila 1)
      'Pendiente',
      data.notas     || '',
      data.fecha     || ''
    ]);
  });

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, id: loteId, count: items.length }))
    .setMimeType(ContentService.MimeType.JSON);
}

function _doUpdateCompra(data) {
  const sh = SS.getSheetByName('Pedidos_Proveedores');
  if (!sh) throw new Error('Hoja Pedidos_Proveedores no encontrada.');

  const shData = sh.getDataRange().getValues();
  for (let r = 1; r < shData.length; r++) {
    if (String(shData[r][0]) === String(data.id)) {
      const estadoAnterior = String(shData[r][8]); // col 9 = Estado (índice 8)
      sh.getRange(r + 1, 9).setValue(data.estado);

      // Si transiciona a Entregado por primera vez → actualizar stock físico
      if (data.estado === 'Entregado' && estadoAnterior !== 'Entregado') {
        const producto = String(shData[r][3]); // col 4 = Producto
        const cantidad = parseFloat(shData[r][4]) || 0; // col 5 = Cantidad
        if (producto && cantidad > 0) _actualizarStockFisico(producto, cantidad);
      }

      return ContentService
        .createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: 'ID no encontrado' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Suma cantidad al Stock Físico del producto (col 3) cuando llega una compra
function _actualizarStockFisico(nombreProducto, cantidad) {
  const hProd = SS.getSheetByName('Productos');
  if (!hProd) return;
  const data = hProd.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][1]).trim().toLowerCase() === nombreProducto.trim().toLowerCase()) {
      const celda = hProd.getRange(r + 1, 6); // col F(6) = Stock Físico
      celda.setValue((celda.getValue() || 0) + cantidad);
      break;
    }
  }
}

// ════════════════════════════════════════════════════════════
//  onEdit — actualiza stock al cambiar estado de un pedido
// ════════════════════════════════════════════════════════════

// Mapeo inverso: columna Home (1-based) → Abreviatura en Productos (col C)
// Con Envío en O(15), productos van de U(21) a AM(39)
const HOME_COL_TO_ABBR = {
  21: 'PPM',   // U
  22: 'PPJyQ', // V
  23: 'PPCyQ', // W
  24: 'SCo',   // X
  25: 'SJyQ',  // Y
  26: 'SCa',   // Z
  27: 'ECaC',  // AA
  28: 'EJyQ',  // AB
  29: 'ECyQ',  // AC
  30: 'EV',    // AD
  31: 'TG',    // AE
  32: 'TLC',   // AF
  33: 'TC',    // AG
  34: 'F',     // AH
  35: 'PMu',   // AI
  36: 'PMa',   // AJ
  37: 'PJyQ',  // AK
  38: 'PCC',   // AL
  39: 'PJyM',  // AM
};

// IMPORTANTE: esta función debe configurarse SOLO como trigger instalable.
// NO usar el nombre "onEdit" para evitar doble ejecución (simple + instalable).
// En Activadores: función = onEditHandler, evento = Al editar
function onEditHandler(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  if (sheetName === 'Home' || sheetName === 'Pilar' || sheetName === 'Capital Federal') return _onEditHome(e);
  if (sheetName === 'Clubes')          return _onEditClubes(e);
  if (sheetName === 'Orden de Compra') return _onEditOC(e);
  if (sheetName === 'Pedidos')         return _onEditPedidos(e);
}

// ── Orden de Compra: auto-fill fechas + stock al cambiar Estado (col U=21) ──
// Columnas v2: E=Canal, L=Abreviatura, M=Cantidad, T=Origen, U=Estado, V=FechaPedido, W=FechaRecibido
function _onEditOC(e) {
  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row <= 1 || col !== 21) return; // solo col U (21) = Estado OC

  const sh       = e.range.getSheet();
  const nuevo    = String(e.value || '');
  const anterior = String(e.oldValue || '');
  const origen   = String(sh.getRange(row, 20).getValue()); // T = Origen

  // Timestamp Argentina
  var ahora   = new Date();
  var argDate = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var dd      = String(argDate.getDate()).padStart(2, '0');
  var mm      = String(argDate.getMonth() + 1).padStart(2, '0');
  var yyyy    = argDate.getFullYear();
  var fechaHoy = dd + '/' + mm + '/' + yyyy;

  // → Pendiente: limpiar Fecha Pedido (V=22) y Fecha Recibido (W=23)
  if (nuevo === 'Pendiente') {
    sh.getRange(row, 22).clearContent();
    sh.getRange(row, 23).clearContent();
    return;
  }

  // → Pedido: llenar Fecha Pedido Proveedor (V=22)
  if (nuevo === 'Pedido' && anterior !== 'Pedido') {
    sh.getRange(row, 22).setValue(fechaHoy);
  }

  // → Recibido: llenar Fecha Recibido (W=23) + sumar stock SOLO si Canal=Depósito
  if (nuevo === 'Recibido' && anterior !== 'Recibido') {
    sh.getRange(row, 23).setValue(fechaHoy);

    var canal = String(sh.getRange(row, 5).getValue()).trim(); // E = Canal
    if (canal === 'Depósito' && origen === 'Orden de Compra') {
      var abbr = String(sh.getRange(row, 12).getValue()).trim(); // L = Abreviatura
      var qty  = Number(sh.getRange(row, 13).getValue()) || 0;  // M = Cantidad
      var refOC = String(sh.getRange(row, 1).getValue() || ''); // A = N° Orden
      if (abbr && qty > 0) {
        var hProd = SS.getSheetByName('Productos');
        if (hProd) {
          var prodData = hProd.getDataRange().getValues();
          for (var r = 1; r < prodData.length; r++) {
            if (String(prodData[r][2]).trim() === abbr) {
              var celdaFis = hProd.getRange(r + 1, 6);
              var fisico   = Number(celdaFis.getValue()) || 0;
              celdaFis.setValue(fisico + qty);
              _logKardex(abbr, '+REC', qty, fisico, fisico + qty, 'OC', refOC);
              SS.toast('Stock +' + qty + ' ' + abbr + ' → ' + (fisico + qty), 'Stock actualizado', 4);
              break;
            }
          }
        }
      }
    }
  }

  // ← Sale de Recibido (corrección): restar stock SOLO si Canal=Depósito
  if (anterior === 'Recibido' && nuevo !== 'Recibido') {
    var canal2 = String(sh.getRange(row, 5).getValue()).trim(); // E = Canal
    if (canal2 === 'Depósito' && origen === 'Orden de Compra') {
      var abbr2 = String(sh.getRange(row, 12).getValue()).trim(); // L
      var qty2  = Number(sh.getRange(row, 13).getValue()) || 0;   // M
      var refOC2 = String(sh.getRange(row, 1).getValue() || '');   // A = N° Orden
      if (abbr2 && qty2 > 0) {
        var hProd2 = SS.getSheetByName('Productos');
        if (hProd2) {
          var prodData2 = hProd2.getDataRange().getValues();
          for (var r2 = 1; r2 < prodData2.length; r2++) {
            if (String(prodData2[r2][2]).trim() === abbr2) {
              var celdaFis2 = hProd2.getRange(r2 + 1, 6);
              var fisico2   = Number(celdaFis2.getValue()) || 0;
              var nuevoStock2 = Math.max(0, fisico2 - qty2);
              celdaFis2.setValue(nuevoStock2);
              _logKardex(abbr2, '-AJU', qty2, fisico2, nuevoStock2, 'OC', refOC2);
              break;
            }
          }
        }
      }
    }
    sh.getRange(row, 23).clearContent(); // W = Fecha Recibido
  }
}

// ── Clubes: OC automática + stock cuando cambia Origen o Estado ──
// Clubes: col L(12)=Origen, col N(14)=Estado de Entrega
// R(18)=Envío, W(23)=Facturado, Productos: cols X(24)–AE(31) → PMu,PMa,PJyQ,PCC,PJyM,PPM,PPJyQ,PPCyQ

const CLUBES_COL_TO_ABBR = {
  24: 'PMu',   // X
  25: 'PMa',   // Y
  26: 'PJyQ',  // Z
  27: 'PCC',   // AA
  28: 'PJyM',  // AB
  29: 'PPM',   // AC
  30: 'PPJyQ', // AD
  31: 'PPCyQ', // AE
};

function _onEditClubes(e) {
  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row <= 1) return;

  const sh = e.range.getSheet();

  // Col L (12) = Origen → generar OC con lock
  if (col === 12) {
    const nuevoOrigen = String(e.value || '');
    if (nuevoOrigen === 'Orden de Compra') {
      var lock = LockService.getScriptLock();
      if (lock.tryLock(100)) {
        try {
          var pedidoNum = String(sh.getRange(row, 2).getValue());
          var shOC = SS.getSheetByName('Orden de Compra');
          if (shOC && shOC.getLastRow() > 1) {
            var existentes = shOC.getRange(2, 6, shOC.getLastRow() - 1, 1).getValues();
            for (var i = 0; i < existentes.length; i++) {
              if (String(existentes[i][0]).trim() === pedidoNum) {
                SS.toast('OC ya existe para ' + pedidoNum, 'Duplicado evitado', 4);
                return;
              }
            }
          }
          generarOrdenDeCompra('Clubes', row);
          SS.toast('Orden de Compra generada para ' + pedidoNum, 'OC', 5);
        } finally {
          lock.releaseLock();
        }
      }
    }
    return;
  }

  // Col N (14) = Estado de Entrega → stock solo si Depósito
  if (col === 14) {
    const origen = String(sh.getRange(row, 12).getValue()); // L = Origen
    const nuevo    = String(e.value || '');
    const anterior = String(e.oldValue || '');

    // → Entregado: descontar Stock Físico solo si Depósito
    if (nuevo === 'Entregado' && anterior !== 'Entregado') {
      if (origen === 'Depósito') {
        const hProductos = SS.getSheetByName('Productos');
        if (hProductos) _clubesStockFisico(sh, row, hProductos, -1);
      }
    }
    // ← Sale de Entregado: devolver Stock Físico solo si Depósito
    if (anterior === 'Entregado' && nuevo !== 'Entregado') {
      if (origen === 'Depósito') {
        const hProductos = SS.getSheetByName('Productos');
        if (hProductos) _clubesStockFisico(sh, row, hProductos, +1);
      }
    }
  }
}

// Ajusta Stock Físico (col F=6) de Productos desde Clubes + log Kardex. signo: -1=restar, +1=sumar
function _clubesStockFisico(shClubes, row, hProductos, signo) {
  const cantidades = shClubes.getRange(row, 24, 1, 8).getValues()[0]; // cols X–AE (8 productos)
  const prodData   = hProductos.getDataRange().getValues();
  const refPedido  = String(shClubes.getRange(row, 2).getValue() || ''); // B = N° Pedido

  Object.keys(CLUBES_COL_TO_ABBR).forEach(function(colStr) {
    const colIdx = Number(colStr);
    const abbr   = CLUBES_COL_TO_ABBR[colIdx];
    const qty    = Number(cantidades[colIdx - 24]) || 0;
    if (qty === 0) return;

    for (var r = 1; r < prodData.length; r++) {
      if (String(prodData[r][2]).trim() === abbr) {
        var celdaFis = hProductos.getRange(r + 1, 6); // F = Stock Físico
        var fisico   = Number(celdaFis.getValue()) || 0;
        var nuevoStock = Math.max(0, fisico + (qty * signo));
        celdaFis.setValue(nuevoStock);
        _logKardex(abbr, signo < 0 ? '-SAL' : '+DEV', qty, fisico, nuevoStock, 'Clubes', refPedido);
        break;
      }
    }
  });
}

// ════════════════════════════════════════════════════════════
//  Orden de Compra — auto-generación desde Home/Delivery/Clubes
// ════════════════════════════════════════════════════════════

// Mapeo abreviatura → proveedor (desde hoja Proveedores)
function _getAbbrToProveedor() {
  const hProv = SS.getSheetByName('Proveedores');
  if (!hProv) return {};
  const data = hProv.getDataRange().getValues();
  const map = {};
  let lastProv = '';
  for (let r = 1; r < data.length; r++) {
    if (data[r][2] && String(data[r][2]).trim()) lastProv = String(data[r][2]).trim(); // col C = Proveedor
    const abbr = String(data[r][4]).trim(); // col E = Abreviatura
    if (abbr) map[abbr] = lastProv;
  }
  return map;
}

// Mapeo abreviatura → nombre producto (desde hoja Proveedores)
function _getAbbrToProductName() {
  const hProv = SS.getSheetByName('Proveedores');
  if (!hProv) return {};
  const data = hProv.getDataRange().getValues();
  const map = {};
  let lastProd = '';
  for (let r = 1; r < data.length; r++) {
    if (data[r][1] && String(data[r][1]).trim()) lastProd = String(data[r][1]).trim(); // col B = Producto
    const abbr  = String(data[r][4]).trim(); // col E = Abreviatura
    const gusto = String(data[r][3]).trim(); // col D = Gusto
    if (abbr) map[abbr] = lastProd + (gusto ? ' — ' + gusto : '');
  }
  return map;
}

// Mapeo abreviatura → costo unitario (desde hoja Productos)
function _getAbbrToCosto() {
  const hProd = SS.getSheetByName('Productos');
  if (!hProd) return {};
  const data = hProd.getDataRange().getValues();
  const map = {};
  for (let r = 1; r < data.length; r++) {
    const abbr = String(data[r][2]).trim(); // col C = Abreviatura
    const costo = String(data[r][9]).replace(/[$.]/g, '').replace(/,/g, '').trim(); // col J = Costo
    if (abbr) map[abbr] = parseFloat(costo) || 0;
  }
  return map;
}

// Mapeo abreviatura → precio de venta unitario (desde hoja Productos)
function _getAbbrToPrecio() {
  const hProd = SS.getSheetByName('Productos');
  if (!hProd) return {};
  const data = hProd.getDataRange().getValues();
  const map = {};
  for (let r = 1; r < data.length; r++) {
    const abbr = String(data[r][2]).trim(); // col C = Abreviatura
    const precio = String(data[r][8]).replace(/[$.]/g, '').replace(/,/g, '').trim(); // col I = Precio
    if (abbr) map[abbr] = parseFloat(precio) || 0;
  }
  return map;
}

/**
 * Genera filas en Orden de Compra para un pedido dado.
 * @param {string} canal - 'Home', 'Pilar', 'Capital Federal' o 'Clubes'
 * @param {number} row - fila del pedido en la hoja de origen
 */
function generarOrdenDeCompra(canal, row) {
  const shOrigen = SS.getSheetByName(canal);
  const shOC     = SS.getSheetByName('Orden de Compra');
  if (!shOrigen || !shOC) return;

  const rowData = shOrigen.getRange(row, 1, 1, shOrigen.getLastColumn()).getValues()[0];

  // Determinar columnas de productos según canal
  let colCliente, colPedido, colTelefono;
  let direccion = '';

  if (canal === 'Home') {
    colCliente = 7; colPedido = 1; colTelefono = 44; // AS(45) = Teléfono
    direccion = [rowData[41], rowData[42], 'Lote ' + rowData[43]].filter(Boolean).join(' · '); // AP=Barrio, AQ=SubBarrio, AR=Lote
  } else if (canal === 'Pilar') {
    colCliente = 7; colPedido = 1; colTelefono = 43; // AR(44) = Teléfono
    direccion = [rowData[41], rowData[42]].filter(Boolean).join(' · '); // AP=Dirección, AQ=Lote
  } else if (canal === 'Capital Federal') {
    colCliente = 7; colPedido = 1; colTelefono = 45; // AT(46) = Teléfono
    direccion = [rowData[41], rowData[42] + ' ' + rowData[43], rowData[44]].filter(Boolean).join(' · '); // AP=Barrio, AQ=Calle, AR=Número, AS=Piso
  } else if (canal === 'Clubes') {
    colCliente = 7; colPedido = 1; colTelefono = 33; // AH(34) = Teléfono
    direccion = [rowData[8], rowData[9], rowData[10]].filter(Boolean).join(' · ');
  }

  const abbrToProvMap   = _getAbbrToProveedor();
  const abbrToNameMap   = _getAbbrToProductName();
  const abbrToCostoMap  = _getAbbrToCosto();
  const abbrToPrecioMap = _getAbbrToPrecio();

  // Precios especiales para Clubes (menores al retail)
  const CLUBES_PRECIOS = {
    'PMu':7000, 'PMa':7000, 'PJyQ':7000, 'PCC':7000, 'PJyM':7800,
    'PPM':11000, 'PPJyQ':11000, 'PPCyQ':11000
  };

  const colToAbbrMap = (canal === 'Clubes')
    ? {24:'PMu', 25:'PMa', 26:'PJyQ', 27:'PCC', 28:'PJyM', 29:'PPM', 30:'PPJyQ', 31:'PPCyQ'}
    : HOME_COL_TO_ABBR;

  const cliente    = String(rowData[colCliente] || '');
  const numPedido  = String(rowData[colPedido] || '');
  const telefono   = String(rowData[colTelefono] || '');

  // Fecha, hora, semana y mes en zona Argentina
  const ahora   = new Date();
  const argDate = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  const MESES   = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  const dd      = String(argDate.getDate()).padStart(2, '0');
  const mm      = String(argDate.getMonth() + 1).padStart(2, '0');
  const yyyy    = argDate.getFullYear();
  const hh      = String(argDate.getHours()).padStart(2, '0');
  const mi      = String(argDate.getMinutes()).padStart(2, '0');
  const fechaStr  = dd + '/' + mm + '/' + yyyy + ' ' + hh + ':' + mi;
  const semana    = _isoWeek(argDate);
  const mesNombre = MESES[argDate.getMonth()];

  // ── Construir filas (25 columnas A-Y) ──
  const newRows = [];
  Object.keys(colToAbbrMap).forEach(function(colStr) {
    const colIdx = Number(colStr);
    const abbr   = colToAbbrMap[colIdx];
    const qty    = Number(rowData[colIdx - 1]) || 0;
    if (qty === 0) return;

    const ocId         = _nextId('OC-');
    const costoUnit    = abbrToCostoMap[abbr] || 0;
    const costoTotal   = costoUnit * qty;
    const precioVenta  = (canal === 'Clubes') ? (CLUBES_PRECIOS[abbr] || 0) : (abbrToPrecioMap[abbr] || 0);

    newRows.push([
      ocId,                                      // A  N° Orden
      fechaStr,                                  // B  Fecha Creación
      semana,                                    // C  Semana
      mesNombre,                                 // D  Mes
      canal,                                     // E  Canal
      numPedido,                                 // F  N° Pedido Origen
      cliente,                                   // G  Cliente
      telefono,                                  // H  Teléfono
      direccion,                                 // I  Dirección
      abbrToProvMap[abbr] || '',                 // J  Proveedor
      abbrToNameMap[abbr] || abbr,               // K  Producto
      abbr,                                      // L  Abreviatura
      qty,                                       // M  Cantidad
      costoUnit,                                 // N  Costo Unitario
      costoTotal,                                // O  Costo Total
      precioVenta,                               // P  Precio Venta Unit.
      0,                                         // Q  Ingreso Total (fórmula)
      0,                                         // R  Margen Bruto $ (fórmula)
      0,                                         // S  Margen % (fórmula)
      'Orden de Compra',                         // T  Origen
      'Pendiente',                               // U  Estado OC
      '',                                        // V  Fecha Pedido Prov
      '',                                        // W  Fecha Recibido
      'No',                                      // X  Pagado Proveedor
      'No',                                      // Y  Cobrado Cliente
    ]);
  });

  if (newRows.length > 0) {
    // Debug: log para diagnosticar el problema de escritura
    var lastRow = shOC.getLastRow();
    var startRow = lastRow + 1;
    Logger.log('OC DEBUG: lastRow=' + lastRow + ', startRow=' + startRow + ', newRows=' + newRows.length + ', cols=' + newRows[0].length);
    SS.toast('Escribiendo ' + newRows.length + ' filas en fila ' + startRow + ' (25 cols)', 'OC Debug', 8);

    // Forzar escritura fila por fila con appendRow como fallback robusto
    newRows.forEach(function(row) {
      shOC.appendRow(row);
    });

    // Fórmulas financieras después de appendRow (locale español: punto y coma)
    var actualLastRow = shOC.getLastRow();
    var formulaStart = actualLastRow - newRows.length + 1;
    for (var i = 0; i < newRows.length; i++) {
      var r = formulaStart + i;
      shOC.getRange(r, 17).setFormula('=P' + r + '*M' + r);          // Q = Ingreso
      shOC.getRange(r, 18).setFormula('=Q' + r + '-O' + r);          // R = Margen $
      shOC.getRange(r, 19).setFormula('=R' + r + '/Q' + r);          // S = Margen %
    }

    // Formato moneda y porcentaje
    shOC.getRange(formulaStart, 13, newRows.length, 1).setNumberFormat('0');       // M Cantidad (sin $)
    shOC.getRange(formulaStart, 14, newRows.length, 2).setNumberFormat('$#,##0');  // N-O Costo
    shOC.getRange(formulaStart, 16, newRows.length, 3).setNumberFormat('$#,##0');  // P-R Precio/Ingreso/Margen$
    shOC.getRange(formulaStart, 19, newRows.length, 1).setNumberFormat('0.0%');    // S Margen%
  }
}

// ── Home: sync stock + auto-generar Orden de Compra ──
// Col I (9) = Origen → si cambia a "Orden de Compra", genera filas en OC
// Col K (11) = Estado de Entrega → si cambia a "Entregado", descuenta stock
function _onEditHome(e) {
  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row <= 1) return;

  const sh = e.range.getSheet();

  // Si cambió Origen (col I=9) a "Orden de Compra" → generar OC con lock
  if (col === 9) {
    const nuevoOrigen = String(e.value || '');
    if (nuevoOrigen === 'Orden de Compra') {
      var lock = LockService.getScriptLock();
      if (lock.tryLock(100)) {
        try {
          // Verificar si ya existen OC para este pedido DENTRO del lock
          var pedidoNum = String(sh.getRange(row, 2).getValue());
          var shOC = SS.getSheetByName('Orden de Compra');
          if (shOC && shOC.getLastRow() > 1) {
            var existentes = shOC.getRange(2, 6, shOC.getLastRow() - 1, 1).getValues();
            for (var i = 0; i < existentes.length; i++) {
              if (String(existentes[i][0]).trim() === pedidoNum) {
                SS.toast('OC ya existe para ' + pedidoNum, 'Duplicado evitado', 4);
                return;
              }
            }
          }
          generarOrdenDeCompra(sh.getName(), row);
          SS.toast('Orden de Compra generada para ' + pedidoNum, 'OC', 5);
        } finally {
          lock.releaseLock();
        }
      } else {
        SS.toast('Otra ejecución ya está generando esta OC', 'Lock', 3);
      }
    }
    return;
  }

  if (col !== 11) return; // solo col K (11) = Estado de Entrega

  const origen = String(sh.getRange(row, 9).getValue()); // col I (9) = Origen
  const nuevo    = String(e.value || '');
  const anterior = String(e.oldValue || '');

  // Columna donde empieza "Hora Entrega" varía por hoja (todas +1 por col Envío):
  // Home: col 46, Pilar: col 45, CF: col 47
  const sheetName = sh.getName();
  const colEntrega = sheetName === 'Pilar' ? 45 : sheetName === 'Capital Federal' ? 47 : 46;

  // → Entregado: registrar fecha SIEMPRE + descontar stock solo si Depósito
  if (nuevo === 'Entregado' && anterior !== 'Entregado') {
    _registrarFechaEntrega(sh, row, colEntrega);
    if (origen === 'Depósito') {
      const hProductos = SS.getSheetByName('Productos');
      if (hProductos) _homeStockFisico(sh, row, hProductos, -1);
    }
  }

  // ← Sale de Entregado: limpiar fecha SIEMPRE + devolver stock solo si Depósito
  if (anterior === 'Entregado' && nuevo !== 'Entregado') {
    sh.getRange(row, colEntrega, 1, 6).clearContent();
    if (origen === 'Depósito') {
      const hProductos = SS.getSheetByName('Productos');
      if (hProductos) _homeStockFisico(sh, row, hProductos, +1);
    }
  }
}

// Llena 6 columnas de entrega desde colStart en zona Argentina
function _registrarFechaEntrega(sh, row, colStart) {
  colStart = colStart || 45;
  var ahora   = new Date();
  var argDate = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var DIAS    = ['Domingo','Lunes','Martes','Mi\u00E9rcoles','Jueves','Viernes','S\u00E1bado'];
  var MESES   = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var dd      = String(argDate.getDate()).padStart(2, '0');
  var mm      = String(argDate.getMonth() + 1).padStart(2, '0');
  var yyyy    = argDate.getFullYear();
  var hh      = String(argDate.getHours()).padStart(2, '0');
  var mi      = String(argDate.getMinutes()).padStart(2, '0');

  sh.getRange(row, colStart, 1, 6).setValues([[
    hh + ':' + mi,                     // Hora Entrega
    DIAS[argDate.getDay()],            // Día Entrega
    dd + '/' + mm + '/' + yyyy,        // Fecha Entrega
    MESES[argDate.getMonth()],         // Mes Entrega
    _isoWeek(argDate),                 // Semana Entrega
    yyyy                               // Año Entrega
  ]]);
}

// Ajusta Stock Físico (col F=6) de Productos + log Kardex. signo: -1 = restar, +1 = sumar
function _homeStockFisico(shHome, row, hProductos, signo) {
  const cantidades = shHome.getRange(row, 21, 1, 19).getValues()[0]; // cols U–AM
  const prodData   = hProductos.getDataRange().getValues();
  const refPedido  = String(shHome.getRange(row, 2).getValue() || ''); // B = N° Pedido

  Object.keys(HOME_COL_TO_ABBR).forEach(function(colStr) {
    const colIdx = Number(colStr);
    const abbr   = HOME_COL_TO_ABBR[colIdx];
    const qty    = Number(cantidades[colIdx - 21]) || 0;
    if (qty === 0) return;

    for (let r = 1; r < prodData.length; r++) {
      if (String(prodData[r][2]).trim() === abbr) {
        const celdaFis = hProductos.getRange(r + 1, 6); // F = Stock Físico
        const fisico   = Number(celdaFis.getValue()) || 0;
        const nuevoStock = Math.max(0, fisico + (qty * signo));
        celdaFis.setValue(nuevoStock);
        _logKardex(abbr, signo < 0 ? '-SAL' : '+DEV', qty, fisico, nuevoStock, 'Home', refPedido);
        break;
      }
    }
  });
}

// ── Fórmulas en Productos: Reservado (E) y Disponible (F) ───
// Ejecutar UNA vez — pone fórmulas SUMPRODUCT que se auto-actualizan.
// Abreviatura (col C de Productos) → letra de columna en Home
// FIX: Columnas corregidas (+1) para coincidir con HOME_PRODUCT_COLS
// Productos van de U(21) a AM(39), NO de T a AL
const ABBR_TO_HOME_COL = {
  'PPM':'U', 'PPJyQ':'V', 'PPCyQ':'W',
  'SCo':'X', 'SJyQ':'Y', 'SCa':'Z',
  'ECaC':'AA', 'EJyQ':'AB', 'ECyQ':'AC', 'EV':'AD',
  'TG':'AE', 'TLC':'AF', 'TC':'AG', 'F':'AH',
  'PMu':'AI', 'PMa':'AJ', 'PJyQ':'AK', 'PCC':'AL', 'PJyM':'AM',
};

// Mapeo para Clubes: abreviatura → letra de columna en Clubes
// Productos van de X(24) a AE(31)
const ABBR_TO_CLUBES_COL = {
  'PMu':'X', 'PMa':'Y', 'PJyQ':'Z', 'PCC':'AA', 'PJyM':'AB',
  'PPM':'AC', 'PPJyQ':'AD', 'PPCyQ':'AE',
};

function setupProductosFormulas() {
  const hProd = SS.getSheetByName('Productos');
  if (!hProd) return;
  const data = hProd.getDataRange().getValues();

  // Calcular semana y año actuales en el script (evita funciones de locale)
  var ahora = new Date();
  var argDate = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var semanaActual = _isoWeek(argDate);
  var anioActual   = argDate.getFullYear();

  for (let r = 1; r < data.length; r++) {
    const abbr    = String(data[r][2]).trim(); // col C = Abreviatura
    const homeCol = ABBR_TO_HOME_COL[abbr];
    if (!homeCol) continue;

    const rowNum = r + 1;
    var dep = 'Dep' + '\u00F3sito';

    // Col E (Vendidos Semana) = SUMPRODUCT: Entregados por semana de ENTREGA
    // Incluye Home + Pilar + Capital Federal (misma estructura de columnas)
    // Home: Semana Entrega=col AX(50), Año Entrega=col AY(51)
    // Pilar: Semana Entrega=col AW(49), Año Entrega=col AX(50)
    // CF: Semana Entrega=col AY(51), Año Entrega=col AZ(52)
    var vendidosFormula =
      'SUMPRODUCT((Home!$I$2:$I$10000="' + dep + '")*(Home!$K$2:$K$10000="Entregado")' +
      '*(Home!$AX$2:$AX$10000=' + semanaActual + ')*(Home!$AY$2:$AY$10000=' + anioActual + ')' +
      '*(Home!' + homeCol + '$2:' + homeCol + '$10000))' +
      '+SUMPRODUCT((Pilar!$I$2:$I$10000="' + dep + '")*(Pilar!$K$2:$K$10000="Entregado")' +
      '*(Pilar!$AW$2:$AW$10000=' + semanaActual + ')*(Pilar!$AX$2:$AX$10000=' + anioActual + ')' +
      '*(Pilar!' + homeCol + '$2:' + homeCol + '$10000))' +
      '+SUMPRODUCT((\'Capital Federal\'!$I$2:$I$10000="' + dep + '")*(\'Capital Federal\'!$K$2:$K$10000="Entregado")' +
      '*(\'Capital Federal\'!$AY$2:$AY$10000=' + semanaActual + ')*(\'Capital Federal\'!$AZ$2:$AZ$10000=' + anioActual + ')' +
      '*(\'Capital Federal\'!' + homeCol + '$2:' + homeCol + '$10000))';

    // Clubes: solo para los 8 productos que vende Clubes
    var clubesCol = ABBR_TO_CLUBES_COL[abbr];
    if (clubesCol) {
      vendidosFormula +=
        '+SUMPRODUCT((Clubes!$L$2:$L$10000="' + dep + '")*(Clubes!$N$2:$N$10000="Entregado")' +
        '*(Clubes!' + clubesCol + '$2:' + clubesCol + '$10000))';
    }
    hProd.getRange(rowNum, 5).setFormula('=' + vendidosFormula);

    // Col G (Reservado) = SUMPRODUCT: Reservados activos en TODAS las hojas
    var reservadoFormula =
      'SUMPRODUCT((Home!$I$2:$I$10000="' + dep + '")*(Home!$K$2:$K$10000="Reservado")' +
      '*(Home!' + homeCol + '$2:' + homeCol + '$10000))' +
      '+SUMPRODUCT((Pilar!$I$2:$I$10000="' + dep + '")*(Pilar!$K$2:$K$10000="Reservado")' +
      '*(Pilar!' + homeCol + '$2:' + homeCol + '$10000))' +
      '+SUMPRODUCT((\'Capital Federal\'!$I$2:$I$10000="' + dep + '")*(\'Capital Federal\'!$K$2:$K$10000="Reservado")' +
      '*(\'Capital Federal\'!' + homeCol + '$2:' + homeCol + '$10000))';

    if (clubesCol) {
      reservadoFormula +=
        '+SUMPRODUCT((Clubes!$L$2:$L$10000="' + dep + '")*(Clubes!$N$2:$N$10000="Reservado")' +
        '*(Clubes!' + clubesCol + '$2:' + clubesCol + '$10000))';
    }
    hProd.getRange(rowNum, 7).setFormula('=' + reservadoFormula);

    // Col H (Stock Disponible) = Stock Físico - Reservado
    hProd.getRange(rowNum, 8).setFormula('=F' + rowNum + '-G' + rowNum);

    // Col K (Margen Unitario) = Precio - Costo
    hProd.getRange(rowNum, 11).setFormula('=I' + rowNum + '-J' + rowNum);

    // Col L (Check D-E) = Stock Inicial - Vendidos (referencia para comparar con F)
    hProd.getRange(rowNum, 12).setFormula('=D' + rowNum + '-E' + rowNum);
  }

  SS.toast('Formulas actualizadas (semana ' + semanaActual + '/' + anioActual + ')', 'Productos', 5);
}

// Ejecutar al inicio de cada semana: copia Stock Físico → Stock Inicial + actualiza fórmulas
function resetStockSemanal() {
  const hProd = SS.getSheetByName('Productos');
  if (!hProd) return;
  const lastRow = hProd.getLastRow();
  if (lastRow < 2) return;

  // Copiar col F (Stock Físico) → col D (Stock Inicial Semana)
  const fisico = hProd.getRange(2, 6, lastRow - 1, 1).getValues();
  hProd.getRange(2, 4, lastRow - 1, 1).setValues(fisico);

  // Actualizar fórmulas de Vendidos con la nueva semana
  setupProductosFormulas();

  SS.toast('Stock Inicial y formulas actualizados para la nueva semana', 'Reset semanal', 5);
}

/** NUCLEAR: elimina TODOS los triggers y crea uno solo limpio */
function resetTriggers() {
  // Paso 1: borrar absolutamente todos los triggers
  var all = ScriptApp.getProjectTriggers();
  for (var i = 0; i < all.length; i++) {
    ScriptApp.deleteTrigger(all[i]);
  }

  // Paso 2: crear UN solo trigger onEdit
  ScriptApp.newTrigger('onEditHandler')
    .forSpreadsheet(SS)
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert('Triggers reseteados: ' + all.length + ' eliminados, 1 nuevo creado (onEditHandler).');
}

/** Muestra/oculta la hoja Config para edición manual */
function mostrarConfig() {
  const sh = SS.getSheetByName('Config');
  if (!sh) { setupConfig(); return; }
  if (sh.isSheetHidden()) {
    sh.showSheet();
    SS.setActiveSheet(sh);
    SS.toast('Config visible. Cuando termines, podés ocultarla desde el menú de la hoja.', 'Config', 5);
  } else {
    sh.hideSheet();
    SS.toast('Config ocultada.', 'Config', 3);
  }
}

// ── Pedidos (legacy): sync stock cuando cambia Estado ────────
function _onEditPedidos(e) {
  if (e.range.getColumn() !== 12) return; // col 12 = Estado

  const nuevoEstado = e.value;
  if (nuevoEstado !== 'entregado' && nuevoEstado !== 'cancelado') return;

  const idPedido   = e.range.getSheet().getRange(e.range.getRow(), 1).getValue();
  const barrio     = e.range.getSheet().getRange(e.range.getRow(), 5).getValue();
  const esClub     = String(barrio).startsWith('Club-');
  const hDetalle   = SS.getSheetByName('Detalle_Pedidos');
  const hProductos = SS.getSheetByName('Productos');
  if (!hDetalle || !hProductos) return;

  const detalleData = hDetalle.getDataRange().getValues();
  const prodData    = hProductos.getDataRange().getValues();

  if (!esClub) {
    detalleData.slice(1).forEach(row => {
      if (row[0] !== idPedido) return;
      const idProd = row[2];
      const qty    = row[4];

      for (let r = 1; r < prodData.length; r++) {
        if (prodData[r][0] === idProd) {
          const celdaRes = hProductos.getRange(r + 1, 5); // E = Reservado
          const celdaFis = hProductos.getRange(r + 1, 4); // D = Stock Físico
          if (nuevoEstado === 'entregado') {
            celdaFis.setValue(Math.max(0, celdaFis.getValue() - qty));
            celdaRes.setValue(Math.max(0, celdaRes.getValue() - qty));
          } else {
            celdaRes.setValue(Math.max(0, celdaRes.getValue() - qty));
          }
          break;
        }
      }
    });
  }

  if (nuevoEstado === 'cancelado') {
    for (let r = detalleData.length - 1; r >= 1; r--) {
      if (detalleData[r][0] === idPedido) {
        hDetalle.deleteRow(r + 1);
      }
    }
  }
}

// ════════════════════════════════════════════════════════════
//  MENÚ CUSTOM — Maleu → Gestión de Stock
// ════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Maleu')
    .addItem('Nueva compra a proveedor', 'abrirSidebarCompra')
    .addItem('Resumen para WhatsApp (pendientes)', 'generarResumenWA')
    .addItem('Hoja de Ruta (búsqueda)', 'generarHojaDeRuta')
    .addItem('Marcar todas las OC como Recibidas', 'marcarOCRecibidas')
    .addSeparator()
    .addItem('Archivar semana (Home + Clubes)', 'archivarSemana')
    .addSeparator()
    .addItem('Actualizar fórmulas Productos', 'setupProductosFormulas')
    .addItem('Reset stock semanal', 'resetStockSemanal')
    .addSeparator()
    .addItem('Ver/Editar Config', 'mostrarConfig')
    .addToUi();
}

/** Marca TODAS las OC en estado "Pedido" como "Recibido" de un solo click.
 *  Solo suma stock en Productos si Canal = "Depósito".
 *  Las de otros canales (Home/Clubes/Red) se marcan Recibido pero NO tocan stock. */
function marcarOCRecibidas() {
  var shOC = SS.getSheetByName('Orden de Compra');
  var hProd = SS.getSheetByName('Productos');
  if (!shOC || !hProd) return;

  var ocData = shOC.getDataRange().getValues();

  // Contar cuántas hay en "Pedido" (col U=21, 0-based=20)
  var filasPedido = [];
  for (var r = 1; r < ocData.length; r++) {
    if (String(ocData[r][20]).trim() === 'Pedido') filasPedido.push(r);
  }

  if (filasPedido.length === 0) {
    SpreadsheetApp.getUi().alert('No hay órdenes en estado "Pedido"');
    return;
  }

  // Confirmar
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Marcar como Recibidas',
    filasPedido.length + ' órdenes en estado "Pedido" se van a marcar como "Recibido".\n\n' +
    'Las de Canal "Depósito" van a sumar stock en Productos.\n' +
    'Las de otros canales (Home/Clubes/Red) solo cambian de estado.\n\n' +
    '¿Confirmar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  // Timestamp
  var ahora = new Date();
  var argDate = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var dd = String(argDate.getDate()).padStart(2, '0');
  var mm = String(argDate.getMonth() + 1).padStart(2, '0');
  var yyyy = argDate.getFullYear();
  var fechaHoy = dd + '/' + mm + '/' + yyyy;

  var prodData = hProd.getDataRange().getValues();
  var stockSumado = {};
  var marcadas = 0;

  filasPedido.forEach(function(r) {
    var row = r + 1; // 1-based
    var canal  = String(ocData[r][4]).trim();   // E = Canal
    var origen = String(ocData[r][19]).trim();  // T = Origen
    var abbr   = String(ocData[r][11]).trim();  // L = Abreviatura
    var qty    = Number(ocData[r][12]) || 0;    // M = Cantidad

    // Marcar como Recibido + fecha
    shOC.getRange(row, 21).setValue('Recibido');  // U = Estado OC
    shOC.getRange(row, 23).setValue(fechaHoy);    // W = Fecha Recibido

    // Solo sumar stock si Canal = Depósito Y Origen = Orden de Compra
    if (canal === 'Depósito' && origen === 'Orden de Compra' && abbr && qty > 0) {
      if (!stockSumado[abbr]) stockSumado[abbr] = 0;
      stockSumado[abbr] += qty;
    }
    marcadas++;
  });

  // Sumar stock en Productos + log Kardex
  var prodActualizados = 0;
  Object.keys(stockSumado).forEach(function(abbr) {
    for (var p = 1; p < prodData.length; p++) {
      if (String(prodData[p][2]).trim() === abbr) {
        var celda = hProd.getRange(p + 1, 6); // F = Stock Físico
        var actual = Number(celda.getValue()) || 0;
        var nuevoStock = actual + stockSumado[abbr];
        celda.setValue(nuevoStock);
        _logKardex(abbr, '+REC', stockSumado[abbr], actual, nuevoStock, 'OC', 'Bulk Recibido');
        prodActualizados++;
        break;
      }
    }
  });

  var msg = marcadas + ' órdenes marcadas como Recibido';
  if (prodActualizados > 0) msg += '\n' + prodActualizados + ' productos actualizados en stock (Depósito)';
  SS.toast(msg, 'Listo', 5);
}

/** Genera Hoja de Ruta: Sección 1 = Búsqueda (por proveedor), Sección 2 = Entregas (por cliente con checkbox) */
function generarHojaDeRuta() {
  var shOC = SS.getSheetByName('Orden de Compra');
  if (!shOC) { SpreadsheetApp.getUi().alert('No se encontró la hoja "Orden de Compra"'); return; }

  var data = shOC.getDataRange().getValues();
  if (data.length <= 1) { SpreadsheetApp.getUi().alert('No hay órdenes de compra'); return; }

  // ── Recopilar datos ──
  var porProv = {};
  var porCliente = {};
  var totalGeneral = 0;

  for (var r = 1; r < data.length; r++) {
    var estado = String(data[r][20]).trim();  // U = Estado OC (0-based 20)
    if (estado !== 'Pendiente' && estado !== 'Pedido') continue;
    var origen = String(data[r][19]).trim();  // T = Origen (0-based 19)
    if (origen === 'Depósito') continue;

    var prov     = String(data[r][9]).trim();   // J = Proveedor
    var producto = String(data[r][10]).trim();  // K = Producto
    var qty      = Number(data[r][12]) || 0;    // M = Cantidad
    var costoU   = Number(String(data[r][13]).replace(/[$.]/g,'').replace(/,/g,'')) || 0; // N = Costo Unit
    var costoT   = costoU * qty;
    var cliente  = String(data[r][6]).trim();   // G = Cliente
    var canal    = String(data[r][4]).trim();   // E = Canal
    var dir      = String(data[r][8]).trim();   // I = Dirección

    if (!prov || qty === 0) continue;

    // Agrupar por proveedor
    if (!porProv[prov]) porProv[prov] = { totalCosto: 0, productos: {} };
    if (!porProv[prov].productos[producto]) porProv[prov].productos[producto] = 0;
    porProv[prov].productos[producto] += qty;
    porProv[prov].totalCosto += costoT;
    totalGeneral += costoT;

    // Agrupar por cliente
    var ck = cliente;
    if (!porCliente[ck]) porCliente[ck] = { canal: canal, dir: dir, items: [], total: 0 };
    porCliente[ck].items.push({ producto: producto, qty: qty, prov: prov });
    porCliente[ck].total += costoT;
  }

  var provs = Object.keys(porProv);
  var clientes = Object.keys(porCliente);
  if (provs.length === 0) { SpreadsheetApp.getUi().alert('No hay órdenes pendientes para buscar'); return; }

  // ── Construir HTML ──
  var html = '<style>' +
    'body{font-family:system-ui,sans-serif;font-size:13px;color:#331C1C;padding:20px;max-width:750px;margin:0 auto;}' +
    'h2{font-size:20px;font-weight:800;margin-bottom:2px;}' +
    '.sub{font-size:12px;color:#7A4A4A;margin-bottom:20px;}' +
    '.section-title{font-size:14px;font-weight:800;text-transform:uppercase;letter-spacing:.08em;color:#F07D47;margin:24px 0 12px;padding-bottom:6px;border-bottom:2px solid #F07D47;}' +
    '.prov{background:#f8f6f0;border-radius:12px;padding:14px 16px;margin-bottom:12px;}' +
    '.prov-t{font-size:14px;font-weight:800;display:flex;justify-content:space-between;margin-bottom:8px;}' +
    '.prov-cost{color:#F07D47;}' +
    '.pl-r{display:flex;justify-content:space-between;padding:3px 0;border-bottom:1px solid #e8dfc4;font-size:13px;}' +
    '.pl-r:last-child{border-bottom:none;}' +
    '.pl-n{font-weight:600;}.pl-q{font-weight:800;color:#F07D47;}' +
    '.cliente{background:#fff;border:1.5px solid #e8dfc4;border-radius:12px;padding:14px 16px;margin-bottom:12px;page-break-inside:avoid;}' +
    '.cliente-head{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:8px;}' +
    '.cliente-name{font-size:14px;font-weight:800;}' +
    '.cliente-canal{font-size:10px;font-weight:700;background:#F07D47;color:#fff;padding:2px 8px;border-radius:50px;}' +
    '.cliente-dir{font-size:11px;color:#7A4A4A;margin-bottom:8px;}' +
    '.item-row{display:flex;align-items:center;gap:8px;padding:5px 0;border-bottom:1px solid #f2e8c7;font-size:13px;}' +
    '.item-row:last-child{border-bottom:none;}' +
    '.checkbox{width:18px;height:18px;border:2px solid #331C1C;border-radius:4px;flex-shrink:0;}' +
    '.item-name{flex:1;font-weight:600;}' +
    '.item-qty{font-weight:800;color:#F07D47;min-width:30px;text-align:right;}' +
    '.item-prov{font-size:10px;color:#7A4A4A;}' +
    '.tb{background:#331C1C;color:#F2E8C7;border-radius:12px;padding:14px 16px;display:flex;justify-content:space-between;font-size:15px;font-weight:800;margin-top:20px;}' +
    '.tb span:last-child{color:#F07D47;}' +
    '.pb{display:block;width:100%;margin-top:12px;padding:12px;background:#F07D47;color:#fff;border:none;border-radius:10px;font-size:14px;font-weight:700;cursor:pointer;}' +
    '@media print{.pb{display:none;} body{padding:10px;font-size:12px;} .cliente{border:1px solid #ccc;} .checkbox{border:1.5px solid #000;}}' +
    '</style>';

  html += '<h2>Hoja de Ruta — Maleu</h2>';
  html += '<p class="sub">' + new Date().toLocaleDateString('es-AR', {weekday:'long', day:'numeric', month:'long', year:'numeric'}) + '</p>';

  // ── SECCIÓN 1: BÚSQUEDA POR PROVEEDOR ──
  html += '<div class="section-title">1. Búsqueda por proveedor</div>';

  provs.forEach(function(prov) {
    var d = porProv[prov];
    html += '<div class="prov"><div class="prov-t"><span>' + prov + '</span><span class="prov-cost">$' + d.totalCosto.toLocaleString('es-AR') + '</span></div>';
    Object.keys(d.productos).forEach(function(prod) {
      html += '<div class="pl-r"><span class="pl-n">' + prod + '</span><span class="pl-q">x' + d.productos[prod] + '</span></div>';
    });
    html += '</div>';
  });

  html += '<div class="tb"><span>Total a pagar proveedores</span><span>$' + totalGeneral.toLocaleString('es-AR') + '</span></div>';

  // ── SECCIÓN 2: ENTREGAS POR CLIENTE ──
  html += '<div class="section-title">2. Entregas por cliente (' + clientes.length + ')</div>';

  clientes.forEach(function(nombre) {
    var c = porCliente[nombre];
    html += '<div class="cliente">';
    html += '<div class="cliente-head"><span class="cliente-name">' + nombre + '</span><span class="cliente-canal">' + c.canal + '</span></div>';
    if (c.dir) html += '<div class="cliente-dir">' + c.dir + '</div>';
    c.items.forEach(function(item) {
      html += '<div class="item-row">' +
        '<div class="checkbox"></div>' +
        '<span class="item-name">' + item.producto + '</span>' +
        '<span class="item-qty">x' + item.qty + '</span>' +
        '</div>';
    });
    html += '</div>';
  });

  html += '<button class="pb" onclick="window.print()">Imprimir Hoja de Ruta</button>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(750).setHeight(600),
    'Hoja de Ruta — Maleu'
  );
}

/** Genera resumen de OC pendientes agrupado por proveedor, listo para copiar a WhatsApp */
function generarResumenWA() {
  const shOC = SS.getSheetByName('Orden de Compra');
  if (!shOC) { SpreadsheetApp.getUi().alert('No se encontró la hoja "Orden de Compra"'); return; }

  const data = shOC.getDataRange().getValues();
  if (data.length <= 1) { SpreadsheetApp.getUi().alert('No hay órdenes de compra'); return; }

  // Agrupar pendientes por proveedor
  const porProveedor = {};
  for (var r = 1; r < data.length; r++) {
    var estado = String(data[r][20]).trim(); // U = Estado OC (0-based 20)
    if (estado !== 'Pendiente') continue;

    var origen = String(data[r][19]).trim();   // T = Origen (0-based 19)
    if (origen !== 'Orden de Compra') continue; // Solo las que hay que comprar

    var proveedor = String(data[r][9]).trim();  // J = Proveedor
    var producto  = String(data[r][10]).trim(); // K = Producto
    var abbr      = String(data[r][11]).trim(); // L = Abreviatura
    var qty       = Number(data[r][12]) || 0;   // M = Cantidad
    var costoUnit = Number(String(data[r][13]).replace(/[$.]/g,'').replace(/,/g,'')) || 0; // N = Costo Unit
    var canal     = String(data[r][4]).trim();  // E = Canal

    if (!proveedor || qty === 0) continue;

    if (!porProveedor[proveedor]) porProveedor[proveedor] = { items: {}, totalCosto: 0 };

    // Agrupar mismo producto sumando cantidades
    var key = abbr || producto;
    if (!porProveedor[proveedor].items[key]) {
      porProveedor[proveedor].items[key] = { nombre: producto, qty: 0, costoUnit: costoUnit };
    }
    porProveedor[proveedor].items[key].qty += qty;
    porProveedor[proveedor].totalCosto += costoUnit * qty;
  }

  var proveedores = Object.keys(porProveedor);
  if (proveedores.length === 0) {
    SpreadsheetApp.getUi().alert('No hay órdenes pendientes');
    return;
  }

  // Generar mensaje por proveedor
  var mensajes = proveedores.map(function(prov) {
    var data = porProveedor[prov];
    var items = Object.values(data.items);
    var lineas = items.map(function(item) {
      return '  · ' + item.nombre + ' × ' + item.qty;
    }).join('\n');

    var totalStr = '$' + data.totalCosto.toLocaleString('es-AR');

    return '━━━━━━━━━━━━━━━\n' +
      '📦 *Pedido para ' + prov + '*\n' +
      '━━━━━━━━━━━━━━━\n\n' +
      lineas + '\n\n' +
      '💰 *Total: ' + totalStr + '*\n\n' +
      '📅 Búsqueda: viernes 27/03\n' +
      '🧡 _Maleu_';
  });

  // Mostrar en un dialog para copiar
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:system-ui,sans-serif;font-size:13px;color:#331C1C;padding:16px;}' +
    '.prov-block{background:#f8f6f0;border-radius:10px;padding:14px;margin-bottom:14px;white-space:pre-wrap;font-family:monospace;font-size:12px;line-height:1.6;}' +
    '.prov-title{font-size:14px;font-weight:800;margin-bottom:8px;font-family:system-ui;}' +
    '.copy-btn{display:block;width:100%;padding:10px;margin-top:8px;background:#25D366;color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer;font-family:system-ui;}' +
    '.copy-btn:hover{background:#1bb954;}' +
    '.copied{background:#331C1C !important;}' +
    'h2{font-size:16px;margin-bottom:14px;}' +
    '</style>' +
    '<h2>Resumen para WhatsApp (' + proveedores.length + ' proveedor' + (proveedores.length > 1 ? 'es' : '') + ')</h2>' +
    mensajes.map(function(msg, i) {
      var prov = proveedores[i];
      var escaped = msg.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
      return '<div class="prov-title">' + prov + '</div>' +
        '<div class="prov-block" id="msg-' + i + '">' + escaped + '</div>' +
        '<button class="copy-btn" onclick="copyMsg(' + i + ',this)">Copiar mensaje de ' + prov + '</button>';
    }).join('<br>') +
    '<script>' +
    'function copyMsg(i,btn){' +
    '  var text=document.getElementById("msg-"+i).textContent;' +
    '  navigator.clipboard.writeText(text).then(function(){' +
    '    btn.textContent="✓ Copiado!";btn.classList.add("copied");' +
    '    setTimeout(function(){btn.textContent="Copiar de nuevo";btn.classList.remove("copied");},2000);' +
    '  });' +
    '}' +
    '</script>'
  ).setTitle('Resumen para WhatsApp').setWidth(420).setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Pedidos a Proveedores');
}

function abrirSidebarCompra() {
  const html = HtmlService.createHtmlOutput(_getSidebarHTML())
    .setTitle('Compra a Proveedor')
    .setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── Datos dinámicos para la sidebar ──────────────────────────

/** Devuelve proveedores únicos desde hoja Proveedores */
function getProveedores() {
  const hProv = SS.getSheetByName('Proveedores');
  if (!hProv) return [];
  const data = hProv.getDataRange().getValues();
  const provs = [];
  const seen = {};
  for (let r = 1; r < data.length; r++) {
    const prov = String(data[r][2]).trim();
    if (prov && !seen[prov]) {
      seen[prov] = true;
      provs.push(prov);
    }
  }
  return provs;
}

/** Devuelve productos de un proveedor con costo desde Productos */
function getProductosPorProveedor(proveedor) {
  const hProv = SS.getSheetByName('Proveedores');
  const hProd = SS.getSheetByName('Productos');
  if (!hProv) return [];

  const provData = hProv.getDataRange().getValues();
  const costoMap = {};
  const stockMap = {};
  if (hProd) {
    const prodData = hProd.getDataRange().getValues();
    for (let r = 1; r < prodData.length; r++) {
      const abbr = String(prodData[r][2]).trim();
      const costo = Number(String(prodData[r][9]).replace(/[$.]/g,'').replace(/,/g,'')) || 0;
      const stock = Number(prodData[r][5]) || 0; // col F = Stock Físico
      if (abbr) { costoMap[abbr] = costo; stockMap[abbr] = stock; }
    }
  }

  const items = [];
  let lastProv = '', lastProd = '';
  for (let r = 1; r < provData.length; r++) {
    if (provData[r][2] && String(provData[r][2]).trim()) lastProv = String(provData[r][2]).trim();
    if (provData[r][1] && String(provData[r][1]).trim()) lastProd = String(provData[r][1]).trim();
    if (lastProv !== proveedor) continue;

    const abbr  = String(provData[r][4]).trim();
    const gusto = String(provData[r][3]).trim();
    if (!abbr) continue;

    items.push({
      abbr: abbr,
      nombre: lastProd + ' — ' + gusto,
      costo: costoMap[abbr] || 0,
      stock: stockMap[abbr] !== undefined ? stockMap[abbr] : null
    });
  }
  return items;
}

/** Genera filas en Orden de Compra para compra de Depósito / Red (sidebar) */
function confirmarCompraDeposito(proveedor, items, fechaBusqueda, vendedor) {
  const shOC = SS.getSheetByName('Orden de Compra');
  if (!shOC) throw new Error('Hoja "Orden de Compra" no encontrada');

  var esMarcos = vendedor === 'Marcos Bottcher';

  // Validar que haya items con cantidad > 0
  const itemsValidos = items.filter(function(item) { return item.qty > 0; });
  if (itemsValidos.length === 0) throw new Error('No hay productos con cantidad > 0');

  // Timestamp Argentina
  const ahora   = new Date();
  const argDate = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  const MESES   = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  const dd   = String(argDate.getDate()).padStart(2, '0');
  const mm   = String(argDate.getMonth() + 1).padStart(2, '0');
  const yyyy = argDate.getFullYear();
  const hh   = String(argDate.getHours()).padStart(2, '0');
  const mi   = String(argDate.getMinutes()).padStart(2, '0');
  const fechaStr  = dd + '/' + mm + '/' + yyyy + ' ' + hh + ':' + mi;
  const semana    = _isoWeek(argDate);
  const mesNombre = MESES[argDate.getMonth()];

  // ── Construir filas (25 columnas A-Y) ──
  const newRows = [];
  itemsValidos.forEach(function(item) {
    const ocId = _nextId('OC-');
    const costoTotal = (item.costo || 0) * item.qty;
    newRows.push([
      ocId,                                       // A  N° Orden
      fechaStr,                                   // B  Fecha Creación
      semana,                                     // C  Semana
      mesNombre,                                  // D  Mes
      esMarcos ? 'Red' : 'Depósito',             // E  Canal
      '',                                         // F  N° Pedido Origen (no aplica)
      vendedor || 'Tadeo — Stock',                // G  Cliente
      '',                                         // H  Teléfono
      esMarcos ? 'Tortugas, Garín' : 'Depósito Maleu', // I  Dirección
      proveedor,                                  // J  Proveedor
      item.nombre,                                // K  Producto
      item.abbr,                                  // L  Abreviatura
      item.qty,                                   // M  Cantidad
      item.costo || 0,                            // N  Costo Unitario
      costoTotal,                                 // O  Costo Total
      0,                                          // P  Precio Venta (0 = reposición)
      0,                                          // Q  Ingreso Total (fórmula)
      0,                                          // R  Margen Bruto $ (fórmula)
      0,                                          // S  Margen % (fórmula)
      item.origen || 'Orden de Compra',           // T  Origen
      'Pendiente',                                // U  Estado OC
      fechaBusqueda || '',                         // V  Fecha Pedido Prov
      '',                                         // W  Fecha Recibido
      'No',                                       // X  Pagado Proveedor
      '',                                         // Y  Cobrado Cliente (vacío = no aplica)
    ]);
  });

  // Escribir filas
  const startRow = shOC.getLastRow() + 1;
  shOC.getRange(startRow, 1, newRows.length, 25).setValues(newRows);

  // Fórmulas financieras (setFormula individual, compatible con locale español)
  for (let i = 0; i < newRows.length; i++) {
    const r = startRow + i;
    shOC.getRange(r, 17).setFormula('=P' + r + '*M' + r);    // Q = Ingreso
    shOC.getRange(r, 18).setFormula('=Q' + r + '-O' + r);    // R = Margen $
    shOC.getRange(r, 19).setFormula('=R' + r + '/Q' + r);    // S = Margen %
  }

  // Formato moneda y porcentaje
  shOC.getRange(startRow, 13, newRows.length, 1).setNumberFormat('0');       // M Cantidad
  shOC.getRange(startRow, 14, newRows.length, 2).setNumberFormat('$#,##0');  // N-O
  shOC.getRange(startRow, 16, newRows.length, 3).setNumberFormat('$#,##0');  // P-R
  shOC.getRange(startRow, 19, newRows.length, 1).setNumberFormat('0.0%');    // S

  // Retornar resumen para confirmación
  const totalCosto = newRows.reduce(function(s, r) { return s + (r[14] || 0); }, 0);
  return {
    ok: true,
    filas: newRows.length,
    totalCosto: totalCosto,
    primerOC: newRows[0][0],
    ultimoOC: newRows[newRows.length - 1][0]
  };
}

// ── Sidebar HTML ─────────────────────────────────────────────

function _getSidebarHTML() {
  return `<!DOCTYPE html>
<html>
<head>
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { font-family:'Segoe UI',system-ui,sans-serif; font-size:13px; color:#331C1C; padding:16px; }
  h2 { font-size:16px; font-weight:800; margin-bottom:4px; }
  .subtitle { font-size:12px; color:#7A4A4A; margin-bottom:18px; }
  label { display:block; font-size:12px; font-weight:700; margin-bottom:4px; margin-top:14px; }
  select, input[type="date"], input[type="number"] {
    width:100%; padding:10px; border:1.5px solid #D4C8A0; border-radius:8px;
    font-family:inherit; font-size:13px; color:#331C1C; background:#fff;
    outline:none; appearance:none;
  }
  select:focus, input:focus { border-color:#F07D47; box-shadow:0 0 0 3px rgba(240,125,71,.12); }
  .products-list { margin-top:12px; }
  .prod-row {
    display:flex; align-items:center; gap:8px; padding:8px 0;
    border-bottom:1px solid #F2E8C7;
  }
  .prod-name { flex:1; font-size:12px; font-weight:600; line-height:1.3; }
  .prod-costo { font-size:11px; color:#7A4A4A; }
  .prod-stock { font-size:10px; font-weight:700; padding:2px 6px; border-radius:50px; display:inline-block; margin-top:2px; }
  .stock-ok  { background:#edf7ee; color:#2e7d32; }
  .stock-low { background:#fff3e0; color:#e67e22; }
  .stock-out { background:#fce8e8; color:#c0392b; }
  .prod-qty { width:55px; text-align:center; padding:8px 4px; }
  .prod-origen { width:55px; padding:6px 2px; font-size:10px; border:1px solid #D4C8A0; border-radius:6px; text-align:center; background:#fff; color:#331C1C; }
  .empty-msg { font-size:12px; color:#7A4A4A; font-style:italic; padding:20px 0; text-align:center; }
  .summary {
    background:#F2E8C7; border-radius:10px; padding:12px; margin-top:16px; display:none;
  }
  .summary.visible { display:block; }
  .summary h4 { font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:.06em; color:#F07D47; margin-bottom:8px; }
  .summary-line { display:flex; justify-content:space-between; font-size:12px; padding:2px 0; }
  .summary-total { font-weight:800; border-top:1.5px solid #D4C8A0; margin-top:6px; padding-top:8px; font-size:13px; }
  .summary-total span:last-child { color:#F07D47; }
  .btn-confirm {
    display:block; width:100%; margin-top:16px; padding:14px;
    background:#331C1C; color:#F2E8C7; border:none; border-radius:10px;
    font-family:inherit; font-size:14px; font-weight:700; cursor:pointer;
    transition:background .15s;
  }
  .btn-confirm:hover { background:#1E1010; }
  .btn-confirm:disabled { background:#ccc; color:#888; cursor:not-allowed; }
  .result {
    margin-top:14px; padding:12px; border-radius:10px; font-size:12px;
    font-weight:600; display:none;
  }
  .result.ok { display:block; background:#e8f5e8; color:#2d6a2d; }
  .result.err { display:block; background:#fce8e8; color:#c0392b; }
  .fecha-group { margin-top:14px; }
</style>
</head>
<body>

<h2>Compra a Proveedor</h2>
<p class="subtitle">Reposición de stock para Depósito</p>

<label for="sel-vendedor">¿Para quién es la compra?</label>
<select id="sel-vendedor">
  <option value="Tadeo — Stock">Tadeo — Stock (Depósito)</option>
  <option value="Marcos Bottcher">Marcos Bottcher (Red)</option>
</select>

<label for="sel-prov" style="margin-top:14px">Proveedor</label>
<select id="sel-prov" onchange="onProvChange()">
  <option value="">Seleccioná un proveedor</option>
</select>

<div id="products-container">
  <div class="empty-msg">Elegí un proveedor para ver sus productos</div>
</div>

<div class="fecha-group">
  <label for="fecha-busqueda">Fecha estimada de búsqueda (dd/mm/aaaa)</label>
  <input type="text" id="fecha-busqueda" placeholder="27/03/2026">
</div>

<div class="summary" id="summary"></div>

<button class="btn-confirm" id="btn-confirm" onclick="confirmar()" disabled>
  Confirmar compra →
</button>

<div class="result" id="result"></div>

<script>
let productos = [];
let proveedorActual = '';

// Cargar proveedores al abrir
google.script.run.withSuccessHandler(function(provs) {
  const sel = document.getElementById('sel-prov');
  provs.forEach(function(p) {
    const opt = document.createElement('option');
    opt.value = p; opt.textContent = p;
    sel.appendChild(opt);
  });
}).getProveedores();

// Pre-cargar fecha de mañana
var manana = new Date();
manana.setDate(manana.getDate() + 1);
var dd = String(manana.getDate()).padStart(2,'0');
var mm = String(manana.getMonth()+1).padStart(2,'0');
var yy = manana.getFullYear();
document.getElementById('fecha-busqueda').value = dd+'/'+mm+'/'+yy;

function onProvChange() {
  proveedorActual = document.getElementById('sel-prov').value;
  const container = document.getElementById('products-container');
  if (!proveedorActual) {
    container.innerHTML = '<div class="empty-msg">Elegí un proveedor para ver sus productos</div>';
    productos = [];
    updateSummary();
    return;
  }
  container.innerHTML = '<div class="empty-msg">Cargando...</div>';
  google.script.run.withSuccessHandler(function(items) {
    productos = items;
    if (items.length === 0) {
      container.innerHTML = '<div class="empty-msg">Sin productos para este proveedor</div>';
      return;
    }
    container.innerHTML = '<div class="products-list">' + items.map(function(item, i) {
      var stockHtml = '';
      if (item.stock !== null) {
        if (item.stock === 0) stockHtml = '<span class="prod-stock stock-out">Sin stock</span>';
        else if (item.stock <= 5) stockHtml = '<span class="prod-stock stock-low">Stock: ' + item.stock + '</span>';
        else stockHtml = '<span class="prod-stock stock-ok">Stock: ' + item.stock + '</span>';
      }
      return '<div class="prod-row">' +
        '<div class="prod-name">' + item.nombre + '<br><span class="prod-costo">' +
        (item.costo > 0 ? '$' + item.costo.toLocaleString('es-AR') + ' c/u' : 'Sin costo cargado') +
        '</span> ' + stockHtml + '</div>' +
        '<input type="number" class="prod-qty" min="0" value="0" ' +
        'data-index="' + i + '" onchange="updateSummary()" oninput="updateSummary()">' +
        '<select class="prod-origen" data-index="' + i + '" onchange="updateSummary()">' +
        '<option value="Orden de Compra">OC</option>' +
        '<option value="Depósito">Dep</option>' +
        '</select>' +
        '</div>';
    }).join('') + '</div>';
    updateSummary();
  }).getProductosPorProveedor(proveedorActual);
}

function getItems() {
  const inputs = document.querySelectorAll('.prod-qty');
  const origenes = document.querySelectorAll('.prod-origen');
  const items = [];
  inputs.forEach(function(input) {
    const idx = parseInt(input.dataset.index);
    const qty = parseInt(input.value) || 0;
    const origenSel = origenes[idx] ? origenes[idx].value : 'Orden de Compra';
    if (qty > 0 && productos[idx]) {
      items.push({
        abbr: productos[idx].abbr,
        nombre: productos[idx].nombre,
        costo: productos[idx].costo,
        qty: qty,
        origen: origenSel
      });
    }
  });
  return items;
}

function updateSummary() {
  const items = getItems();
  const sumDiv = document.getElementById('summary');
  const btn = document.getElementById('btn-confirm');

  if (items.length === 0) {
    sumDiv.classList.remove('visible');
    btn.disabled = true;
    return;
  }

  let total = 0;
  const lines = items.map(function(item) {
    const sub = item.costo * item.qty;
    total += sub;
    var origenTag = item.origen === 'Depósito' ? ' <small style="color:#e67e22">(Dep)</small>' : '';
    return '<div class="summary-line"><span>' + item.nombre + ' x' + item.qty + origenTag +
      '</span><span>$' + sub.toLocaleString('es-AR') + '</span></div>';
  }).join('');

  sumDiv.innerHTML = '<h4>Resumen de compra — ' + proveedorActual + '</h4>' +
    lines +
    '<div class="summary-line summary-total"><span>Total a pagar</span><span>$' +
    total.toLocaleString('es-AR') + '</span></div>';
  sumDiv.classList.add('visible');
  btn.disabled = false;
}

function confirmar() {
  const items = getItems();
  if (items.length === 0) return;

  const fecha = document.getElementById('fecha-busqueda').value || '';
  const btn = document.getElementById('btn-confirm');
  const result = document.getElementById('result');

  btn.disabled = true;
  btn.textContent = 'Generando...';
  result.className = 'result';

  google.script.run
    .withSuccessHandler(function(resp) {
      result.className = 'result ok';
      result.textContent = 'Compra registrada: ' + resp.filas + ' productos (' +
        resp.primerOC + ' a ' + resp.ultimoOC + ') — Total: $' +
        resp.totalCosto.toLocaleString('es-AR');
      btn.textContent = 'Confirmar compra →';
      // Resetear cantidades
      document.querySelectorAll('.prod-qty').forEach(function(i) { i.value = 0; });
      updateSummary();
    })
    .withFailureHandler(function(err) {
      result.className = 'result err';
      result.textContent = 'Error: ' + err.message;
      btn.disabled = false;
      btn.textContent = 'Confirmar compra →';
    })
    .confirmarCompraDeposito(proveedorActual, items, fecha, document.getElementById('sel-vendedor').value);
}
</script>
</body>
</html>`;
}

// ════════════════════════════════════════════════════════════
//  setupSheets — ejecutar UNA sola vez para formatear todo
// ════════════════════════════════════════════════════════════
function setupSheets() {
  setupConfig();
  _setupPedidos();
  _setupProductos();
  _setupDetalle();
  _setupPanel();
  _setupProveedores();
  _setupEgresos();
  _setupHome();
  _setupOrdenDeCompra();
  setupKardex();
  SS.toast('Sheets de Maleu configurados correctamente', 'Setup completo', 5);
}

// ── Hoja: Home — diseño, dropdowns y colores ────────────────
function _setupHome() {
  let sh = SS.getSheetByName('Home');
  if (!sh) sh = SS.insertSheet('Home');

  // ── Encabezados ─────────────────────────────────────────────
  const headers = [
    'Hora Pedido','N° Pedido','Día Pedido','Fecha Pedido','Mes Pedido','Semana Pedido','Año Pedido',
    'Cliente','Origen','Día de Entrega Elegido','Estado de Entrega','Forma de Pago',
    'Estado de Pago','Total ($)','Efectivo','Transferencia',
    'Propina Efectivo','Propina Transferencia',
    'PPM','PPJyQ','PPCyQ','SCo','SJyQ','SCa','ECaC','EJyQ',
    'TG','TLC','TC','F','PMa','PJyQ','PCC','PJyM',
    'Costo','Margen Bruto','Barrio','Sub Barrio','Domicilio - Lote','Teléfono',
    'Hora Entrega','Día Entrega','Fecha Entrega','Mes Entrega','Semana Entrega','Año Entrega'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(BROWN).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.setFrozenRows(1);
  sh.setRowHeight(1, 40);

  // ── Ancho de columnas ──────────────────────────────────────
  const widths = [
    65,  // A  Hora
    95,  // B  N° Pedido
    85,  // C  Día
    100, // D  Fecha
    50,  // E  Mes
    65,  // F  Semana
    55,  // G  Año
    170, // H  Cliente
    110, // I  Origen
    110, // J  Día de Entrega
    130, // K  Estado de Entrega
    120, // L  Forma de Pago
    120, // M  Estado de Pago
    95,  // N  Total ($)
    90,  // O  Efectivo
    105, // P  Transferencia
    90,  // Q  Propina Efec
    90,  // R  Propina Trans
  ];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  // Productos (S–AH) → compactas
  for (let c = 19; c <= 34; c++) sh.setColumnWidth(c, 55);
  // AI–AN
  [95, 100, 140, 140, 130, 130].forEach((w, i) => sh.setColumnWidth(35 + i, w));
  // AO–AT (Entrega)
  [65, 90, 100, 90, 75, 60].forEach((w, i) => sh.setColumnWidth(41 + i, w));

  // ── Formato numérico ──────────────────────────────────────
  sh.getRange('N2:P5000').setNumberFormat('$#,##0');
  sh.getRange('AI2:AJ5000').setNumberFormat('$#,##0');
  sh.getRange('Q2:R5000').setNumberFormat('$#,##0');

  // Centrar columnas numéricas y productos
  sh.getRange('A2:A5000').setHorizontalAlignment('center');
  sh.getRange('E2:G5000').setHorizontalAlignment('center');
  sh.getRange('S2:AH5000').setHorizontalAlignment('center');

  // ── Dropdown: I — Origen ──────────────────────────────────
  const origenRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pendiente','Depósito','Orden de Compra'], true)
    .setAllowInvalid(false).build();
  sh.getRange('I2:I5000').setDataValidation(origenRule);

  // ── Dropdown: K — Estado de Entrega ───────────────────────
  const entregaRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pendiente','Reservado','Entregado','Cancelado'], true)
    .setAllowInvalid(false).build();
  sh.getRange('K2:K5000').setDataValidation(entregaRule);

  // ── Dropdown: L — Forma de Pago ───────────────────────────
  const pagoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Efectivo','Transferencia'], true)
    .setAllowInvalid(false).build();
  sh.getRange('L2:L5000').setDataValidation(pagoRule);

  // ── Dropdown: M — Estado de Pago ──────────────────────────
  const cobroRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['No Cobrado','Cobrado'], true)
    .setAllowInvalid(false).build();
  sh.getRange('M2:M5000').setDataValidation(cobroRule);

  // ── Conditional Formatting ────────────────────────────────
  const rules = [];

  // I — Origen
  const iRange = sh.getRange('I2:I5000');
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Pendiente')
    .setBackground('#FFF9C4').setFontColor('#7A6000').setBold(true)
    .setRanges([iRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Depósito')
    .setBackground('#C8E6C9').setFontColor('#1B5E20').setBold(true)
    .setRanges([iRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Orden de Compra')
    .setBackground('#BBDEFB').setFontColor('#0D47A1').setBold(true)
    .setRanges([iRange]).build());

  // K — Estado de Entrega
  const kRange = sh.getRange('K2:K5000');
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Pendiente')
    .setBackground('#FFF9C4').setFontColor('#7A6000').setBold(true)
    .setRanges([kRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Reservado')
    .setBackground('#BBDEFB').setFontColor('#0D47A1').setBold(true)
    .setRanges([kRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Entregado')
    .setBackground('#C8E6C9').setFontColor('#1B5E20').setBold(true)
    .setRanges([kRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Cancelado')
    .setBackground('#FFCDD2').setFontColor('#B71C1C').setBold(true)
    .setRanges([kRange]).build());

  // M — Estado de Pago
  const mRange = sh.getRange('M2:M5000');
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('No Cobrado')
    .setBackground('#FFE0B2').setFontColor('#E65100').setBold(true)
    .setRanges([mRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Cobrado')
    .setBackground('#C8E6C9').setFontColor('#1B5E20').setBold(true)
    .setRanges([mRange]).build());

  // Margen Bruto positivo/negativo
  const ajRange = sh.getRange('AJ2:AJ5000');
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#E8F5E9').setFontColor('#1B5E20').setBold(true)
    .setRanges([ajRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThanOrEqualTo(0)
    .setBackground('#FFEBEE').setFontColor('#B71C1C').setBold(true)
    .setRanges([ajRange]).build());

  sh.setConditionalFormatRules(rules);

  // ── Color de fondo alterno (banding) ──────────────────────
  try {
    sh.getRange('A2:AT5000').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  } catch(ex) {}

  // ── Tab color ─────────────────────────────────────────────
  sh.setTabColor(ORANGE);
}

// ── Hoja: Orden de Compra — estructura profesional v2 (25 cols A-Y) ─────
function _setupOrdenDeCompra() {
  let sh = SS.getSheetByName('Orden de Compra');
  if (!sh) sh = SS.insertSheet('Orden de Compra');

  // ── Encabezados (25 columnas A-Y) ──────────────────────────
  const headers = [
    'N° Orden',             // A
    'Fecha Creación',       // B
    'Semana',               // C
    'Mes',                  // D
    'Canal',                // E
    'N° Pedido Origen',     // F
    'Cliente',              // G
    'Teléfono',             // H
    'Dirección',            // I
    'Proveedor',            // J
    'Producto',             // K
    'Abreviatura',          // L
    'Cantidad',             // M
    'Costo Unitario',       // N
    'Costo Total',          // O
    'Precio Venta Unit.',   // P
    'Ingreso Total',        // Q
    'Margen Bruto ($)',     // R
    'Margen (%)',           // S
    'Origen',               // T
    'Estado OC',            // U
    'Fecha Pedido Prov',    // V
    'Fecha Recibido',       // W
    'Pagado Proveedor',     // X
    'Cobrado Cliente',      // Y
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(BROWN).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true);

  sh.setFrozenRows(1);
  sh.setRowHeight(1, 44);

  // ── Ancho de columnas ──────────────────────────────────────
  const widths = [
    95,   // A  N° Orden
    140,  // B  Fecha Creación
    65,   // C  Semana
    85,   // D  Mes
    95,   // E  Canal
    105,  // F  N° Pedido Origen
    170,  // G  Cliente
    120,  // H  Teléfono
    160,  // I  Dirección
    140,  // J  Proveedor
    200,  // K  Producto
    75,   // L  Abreviatura
    75,   // M  Cantidad
    105,  // N  Costo Unitario
    105,  // O  Costo Total
    115,  // P  Precio Venta Unit.
    105,  // Q  Ingreso Total
    115,  // R  Margen Bruto ($)
    85,   // S  Margen (%)
    120,  // T  Origen
    110,  // U  Estado OC
    120,  // V  Fecha Pedido Prov
    120,  // W  Fecha Recibido
    120,  // X  Pagado Proveedor
    120,  // Y  Cobrado Cliente
  ];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // ── Formato numérico ──────────────────────────────────────
  sh.getRange('N2:P5000').setNumberFormat('$#,##0');    // Costo Unit, Costo Total, Precio Venta
  sh.getRange('Q2:R5000').setNumberFormat('$#,##0');    // Ingreso, Margen $
  sh.getRange('S2:S5000').setNumberFormat('0.0%');      // Margen %

  // ── Centrar columnas compactas ─────────────────────────────
  sh.getRange('A2:A5000').setHorizontalAlignment('center');
  sh.getRange('C2:C5000').setHorizontalAlignment('center');
  sh.getRange('L2:M5000').setHorizontalAlignment('center');

  // ── Dropdown: T — Origen ───────────────────────────────────
  const origenRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Orden de Compra','Depósito'], true)
    .setAllowInvalid(false).build();
  sh.getRange('T2:T5000').setDataValidation(origenRule);

  // ── Dropdown: U — Estado OC ────────────────────────────────
  const estadoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pendiente','Pedido','Recibido'], true)
    .setAllowInvalid(false).build();
  sh.getRange('U2:U5000').setDataValidation(estadoRule);

  // ── Dropdown: X — Pagado Proveedor ─────────────────────────
  const pagadoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['No','Sí'], true)
    .setAllowInvalid(false).build();
  sh.getRange('X2:X5000').setDataValidation(pagadoRule);

  // ── Dropdown: Y — Cobrado Cliente ──────────────────────────
  const cobradoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['No','Sí',''], true)
    .setAllowInvalid(false).build();
  sh.getRange('Y2:Y5000').setDataValidation(cobradoRule);

  // ── Conditional Formatting ────────────────────────────────
  const rules = [];

  // U — Estado OC
  const uRange = sh.getRange('U2:U5000');
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Pendiente')
    .setBackground('#FFF9C4').setFontColor('#7A6000').setBold(true)
    .setRanges([uRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Pedido')
    .setBackground('#BBDEFB').setFontColor('#0D47A1').setBold(true)
    .setRanges([uRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Recibido')
    .setBackground('#C8E6C9').setFontColor('#1B5E20').setBold(true)
    .setRanges([uRange]).build());

  // T — Origen
  const tRange = sh.getRange('T2:T5000');
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Orden de Compra')
    .setBackground('#BBDEFB').setFontColor('#0D47A1').setBold(true)
    .setRanges([tRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Depósito')
    .setBackground('#C8E6C9').setFontColor('#1B5E20').setBold(true)
    .setRanges([tRange]).build());

  // R — Margen Bruto ($): verde positivo, rojo negativo
  const rRange = sh.getRange('R2:R5000');
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#E8F5E9').setFontColor('#1B5E20').setBold(true)
    .setRanges([rRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThanOrEqualTo(0)
    .setBackground('#FFEBEE').setFontColor('#B71C1C').setBold(true)
    .setRanges([rRange]).build());

  // X — Pagado Proveedor
  const xRange = sh.getRange('X2:X5000');
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('No')
    .setBackground('#FFE0B2').setFontColor('#E65100').setBold(true)
    .setRanges([xRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Sí')
    .setBackground('#C8E6C9').setFontColor('#1B5E20').setBold(true)
    .setRanges([xRange]).build());

  // Y — Cobrado Cliente
  const yRange = sh.getRange('Y2:Y5000');
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('No')
    .setBackground('#FFE0B2').setFontColor('#E65100').setBold(true)
    .setRanges([yRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Sí')
    .setBackground('#C8E6C9').setFontColor('#1B5E20').setBold(true)
    .setRanges([yRange]).build());

  sh.setConditionalFormatRules(rules);

  // ── Banding (filas alternas) ───────────────────────────────
  try {
    sh.getRange('A2:Y5000').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  } catch(ex) {}

  // ── Tab color ──────────────────────────────────────────────
  sh.setTabColor('#0D47A1');
}

// ════════════════════════════════════════════════════════════
//  migrateDetalleNombre — ejecutar UNA sola vez para migrar
//  filas existentes al nuevo esquema con columna Nombre (col B)
// ════════════════════════════════════════════════════════════
function migrateDetalleNombre() {
  const hDetalle  = SS.getSheetByName('Detalle_Pedidos');
  const hPedidos  = SS.getSheetByName('Pedidos');
  if (!hDetalle || !hPedidos) { SS.toast('Faltan hojas', 'Error', 4); return; }

  const detalleData = hDetalle.getDataRange().getValues();
  const headers     = detalleData[0];

  // Si ya tiene 7 columnas con Nombre en col B, no hay nada que migrar
  if (headers.length >= 7 && headers[1] === 'Nombre') {
    SS.toast('Ya migrado — col B ya es Nombre', 'migrateDetalleNombre', 5);
    return;
  }

  // Construir mapa idPedido → Nombre desde la hoja Pedidos
  const pedidosData = hPedidos.getDataRange().getValues();
  const pedidosMap  = {};
  pedidosData.slice(1).forEach(row => { if (row[0]) pedidosMap[String(row[0])] = String(row[2] || ''); });

  // Insertar columna B (posición 2) en Detalle_Pedidos
  hDetalle.insertColumnBefore(2);

  // Escribir encabezado Nombre en B1
  hDetalle.getRange(1, 2).setValue('Nombre')
    .setBackground(BROWN).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Rellenar B2:B{n} con el nombre del pedido correspondiente
  const lastRow = hDetalle.getLastRow();
  if (lastRow >= 2) {
    const idCol    = hDetalle.getRange(2, 1, lastRow - 1, 1).getValues();
    const nombres  = idCol.map(([id]) => [pedidosMap[String(id)] || '']);
    hDetalle.getRange(2, 2, lastRow - 1, 1).setValues(nombres);
  }

  // Re-aplicar formato de columnas
  hDetalle.setColumnWidth(2, 150);
  hDetalle.getRange('F2:G1000').setNumberFormat('$#,##0');
  hDetalle.getRange('E2:E1000').setHorizontalAlignment('center');

  SS.toast(`✅ Migración completa — ${lastRow - 1} filas actualizadas`, 'migrateDetalleNombre', 6);
}

// ── Hoja: Pedidos ────────────────────────────────────────────
function _setupPedidos() {
  let sh = SS.getSheetByName('Pedidos');
  if (!sh) sh = SS.insertSheet('Pedidos');

  // Nueva estructura con Canal (col 2) y Depósito (col 14)
  const headers = ['ID Pedido','Canal','Fecha','Nombre','Barrio','Lote','Teléfono',
                   'Día entrega','Horario','Pago','Total','Estado','Fecha solo','Depósito'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(BROWN).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.setFrozenRows(1);
  sh.setRowHeight(1, 36);

  [130,110,150,170,140,70,130,110,90,120,90,110,0,120].forEach((w, i) => {
    if (w > 0) sh.setColumnWidth(i + 1, w);
  });
  sh.hideColumns(13); // Fecha solo — solo para fórmulas del Panel

  // Dropdown de estado (col 12)
  const dropRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['pendiente','confirmado','entregado','cancelado'], true)
    .setAllowInvalid(false).build();
  sh.getRange(2, 12, 1000, 1).setDataValidation(dropRule);

  // Formato moneda en Total (col K)
  sh.getRange('K2:K1000').setNumberFormat('$#,##0');

  // Formato de fecha (col C)
  sh.getRange('C2:C1000').setNumberFormat('@'); // texto

  // Conditional formatting — Estado (col L)
  const lRange = sh.getRange('L2:L1000');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('pendiente')
      .setBackground('#FFF9C4').setFontColor('#7A6000').setBold(true)
      .setRanges([lRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('confirmado')
      .setBackground('#BBDEFB').setFontColor('#0D47A1').setBold(true)
      .setRanges([lRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('entregado')
      .setBackground('#C8E6C9').setFontColor('#1B5E20').setBold(true)
      .setRanges([lRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('cancelado')
      .setBackground('#FFCDD2').setFontColor('#B71C1C').setBold(true)
      .setRanges([lRange]).build(),
  ]);

  // Banding (filas alternas)
  try { sh.getRange('A2:N1000').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false); }
  catch(ex) {}
}

// ── Hoja: Productos ──────────────────────────────────────────
function _setupProductos() {
  let sh = SS.getSheetByName('Productos');
  if (!sh) return;

  // ── Encabezados ─────────────────────────────────────────────
  const headers = [
    'ID','Producto','Abreviatura',
    'Stock Inicial\nSemana','Vendidos\nSemana','Stock Físico',
    'Reservado','Stock Disponible',
    'Precio','Costo','Margen Unit.','Check\n(D-E)'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(BROWN).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true);

  sh.setFrozenRows(1);
  sh.setRowHeight(1, 44);

  // ── Ancho de columnas ──────────────────────────────────────
  [40, 230, 70, 100, 90, 95, 90, 115, 90, 90, 95, 80].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // ── Formato numérico ──────────────────────────────────────
  sh.getRange('I2:K100').setNumberFormat('$#,##0');

  // ── Centrar columnas numéricas ─────────────────────────────
  sh.getRange('A2:A100').setHorizontalAlignment('center');
  sh.getRange('C2:H100').setHorizontalAlignment('center');

  // ── Conditional formatting — Stock Disponible (col H) ──────
  const hRange = sh.getRange('H2:H100');
  const rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setBackground('#FFCDD2').setFontColor('#B71C1C').setBold(true)
    .setRanges([hRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(1, 5)
    .setBackground('#FFE0B2').setFontColor('#E65100').setBold(true)
    .setRanges([hRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(5)
    .setBackground('#C8E6C9').setFontColor('#1B5E20')
    .setRanges([hRange]).build());

  // Vendidos Semana (col E) — resaltar si vendió algo
  const eRange = sh.getRange('E2:E100');
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#E3F2FD').setFontColor('#0D47A1').setBold(true)
    .setRanges([eRange]).build());

  // Margen Unitario (col K) — verde positivo, rojo negativo
  const kRange = sh.getRange('K2:K100');
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#E8F5E9').setFontColor('#1B5E20')
    .setRanges([kRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThanOrEqualTo(0)
    .setBackground('#FFEBEE').setFontColor('#B71C1C')
    .setRanges([kRange]).build());

  // Check D-E (col L) — centrar
  sh.getRange('L2:L100').setHorizontalAlignment('center');

  sh.setConditionalFormatRules(rules);

  // ── Banding ────────────────────────────────────────────────
  try { sh.getRange('A2:L100').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false); }
  catch(ex) {}

  sh.setTabColor('#1B5E20');
}

// ── Hoja: Detalle_Pedidos ────────────────────────────────────
function _setupDetalle() {
  let sh = SS.getSheetByName('Detalle_Pedidos');
  if (!sh) sh = SS.insertSheet('Detalle_Pedidos');

  sh.getRange(1, 1, 1, 7).setValues([['ID Pedido','Nombre','ID Producto','Producto','Cantidad','Precio unit.','Subtotal']])
    .setBackground(BROWN).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.setFrozenRows(1);
  sh.setRowHeight(1, 36);
  [130, 150, 90, 230, 80, 110, 110].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  sh.getRange('F2:G1000').setNumberFormat('$#,##0');
  sh.getRange('E2:E1000').setHorizontalAlignment('center');

  try { sh.getRange('A2:G1000').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false); }
  catch(ex) {}
}

// ── Hoja: Panel ──────────────────────────────────────────────
function _setupPanel() {
  const existing = SS.getSheetByName('Panel');
  if (existing) SS.deleteSheet(existing);
  const sh = SS.insertSheet('Panel', 0);
  sh.setTabColor(ORANGE);
  sh.setHiddenGridlines(true);
  sh.getRange('A1:H80').setBackground('#F7F3EE');

  sh.setColumnWidth(1, 24);
  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 165);
  sh.setColumnWidth(4, 165);
  sh.setColumnWidth(5, 165);
  sh.setColumnWidth(6, 165);
  sh.setColumnWidth(7, 24);

  // ── TÍTULO ──
  sh.setRowHeight(1, 16);
  sh.setRowHeight(2, 64);
  sh.getRange('B2:F2').merge()
    .setValue('🍕  MALEU — Panel de Operaciones')
    .setBackground(BROWN).setFontColor('#FFFFFF')
    .setFontSize(18).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(3, 16);

  // ── ESTADO DE PEDIDOS ──
  sh.setRowHeight(4, 28);
  sh.getRange('B4').setValue('ESTADO DE PEDIDOS')
    .setFontSize(9).setFontWeight('bold').setFontColor(ORANGE);
  sh.setRowHeight(5, 28);
  sh.setRowHeight(6, 72);

  sh.getRange('C5:F5').setValues([['Pendientes','Confirmados','Entregados','Cancelados']])
    .setFontSize(10).setFontColor('#666666').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.getRange('C6').setFormulaLocal('=SI.ERROR(CONTAR.SI(Pedidos!L:L;"pendiente");0)')
    .setFontSize(30).setFontWeight('bold').setFontColor('#7A6000')
    .setBackground('#FFFDE7').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange('D6').setFormulaLocal('=SI.ERROR(CONTAR.SI(Pedidos!L:L;"confirmado");0)')
    .setFontSize(30).setFontWeight('bold').setFontColor('#0D47A1')
    .setBackground('#E3F2FD').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange('E6').setFormulaLocal('=SI.ERROR(CONTAR.SI(Pedidos!L:L;"entregado");0)')
    .setFontSize(30).setFontWeight('bold').setFontColor('#1B5E20')
    .setBackground('#E8F5E9').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange('F6').setFormulaLocal('=SI.ERROR(CONTAR.SI(Pedidos!L:L;"cancelado");0)')
    .setFontSize(30).setFontWeight('bold').setFontColor('#B71C1C')
    .setBackground('#FFEBEE').setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.setRowHeight(7, 16);

  // ── CANAL ──
  sh.setRowHeight(8, 28);
  sh.getRange('B8').setValue('PEDIDOS ACTIVOS POR CANAL')
    .setFontSize(9).setFontWeight('bold').setFontColor(ORANGE);
  sh.setRowHeight(9, 28);
  sh.setRowHeight(10, 64);

  sh.getRange('C9:E9').setValues([['🏡 Estancias','🏃 Clubes','🛵 Delivery']])
    .setFontSize(10).setFontColor('#666666').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  const activos  = '(CONTAR.SI.CONJUNTO(Pedidos!L:L;"pendiente")+CONTAR.SI.CONJUNTO(Pedidos!L:L;"confirmado"))';
  const clubes   = '(CONTAR.SI.CONJUNTO(Pedidos!E:E;"Club-*";Pedidos!L:L;"pendiente")+CONTAR.SI.CONJUNTO(Pedidos!E:E;"Club-*";Pedidos!L:L;"confirmado"))';
  const delivery = '(CONTAR.SI.CONJUNTO(Pedidos!E:E;"Delivery-*";Pedidos!L:L;"pendiente")+CONTAR.SI.CONJUNTO(Pedidos!E:E;"Delivery-*";Pedidos!L:L;"confirmado"))';

  sh.getRange('C10').setFormulaLocal(`=SI.ERROR(${activos}-${clubes}-${delivery};0)`)
    .setFontSize(30).setFontWeight('bold').setFontColor(BROWN)
    .setBackground('#FFF8F0').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange('D10').setFormulaLocal(`=SI.ERROR(${clubes};0)`)
    .setFontSize(30).setFontWeight('bold').setFontColor('#1E40AF')
    .setBackground('#EFF6FF').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange('E10').setFormulaLocal(`=SI.ERROR(${delivery};0)`)
    .setFontSize(30).setFontWeight('bold').setFontColor('#065F46')
    .setBackground('#ECFDF5').setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.setRowHeight(11, 16);

  // ── INGRESOS ──
  sh.setRowHeight(12, 28);
  sh.getRange('B12').setValue('INGRESOS')
    .setFontSize(9).setFontWeight('bold').setFontColor(ORANGE);
  sh.setRowHeight(13, 28);
  sh.setRowHeight(14, 64);

  sh.getRange('C13:D13').merge().setValue('Total facturado (entregados)')
    .setFontSize(10).setFontColor('#666666').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange('E13:F13').merge().setValue('Pendiente de cobro')
    .setFontSize(10).setFontColor('#666666').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.getRange('C14:D14').merge()
    .setFormulaLocal('=SI.ERROR(SUMAR.SI(Pedidos!L:L;"entregado";Pedidos!K:K);0)')
    .setNumberFormat('$#,##0')
    .setFontSize(24).setFontWeight('bold').setFontColor('#1B5E20')
    .setBackground('#E8F5E9').setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.getRange('E14:F14').merge()
    .setFormulaLocal('=SI.ERROR(SUMAR.SI.CONJUNTO(Pedidos!K:K;Pedidos!L:L;"pendiente")+SUMAR.SI.CONJUNTO(Pedidos!K:K;Pedidos!L:L;"confirmado");0)')
    .setNumberFormat('$#,##0')
    .setFontSize(24).setFontWeight('bold').setFontColor('#0D47A1')
    .setBackground('#E3F2FD').setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.setRowHeight(15, 16);

  // ── PEDIDOS POR DÍA ──
  sh.setRowHeight(16, 28);
  sh.getRange('B16').setValue('PEDIDOS ACTIVOS POR DÍA DE ENTREGA')
    .setFontSize(9).setFontWeight('bold').setFontColor(ORANGE);
  sh.setRowHeight(17, 28);
  sh.setRowHeight(18, 64);

  const dias = ['Miércoles','Viernes','Sábado','Domingo'];
  sh.getRange('C17:F17').setValues([dias])
    .setFontSize(10).setFontColor('#666666').setFontWeight('bold')
    .setHorizontalAlignment('center').setBackground(CREAM);

  dias.forEach((dia, i) => {
    sh.getRange(18, 3 + i)
      .setFormulaLocal(`=SI.ERROR(CONTAR.SI.CONJUNTO(Pedidos!H:H;"${dia}";Pedidos!L:L;"pendiente")+CONTAR.SI.CONJUNTO(Pedidos!H:H;"${dia}";Pedidos!L:L;"confirmado");0)`)
      .setFontSize(28).setFontWeight('bold').setFontColor(BROWN)
      .setBackground('#FFF8F0').setHorizontalAlignment('center').setVerticalAlignment('middle');
  });

  sh.setRowHeight(19, 16);

  // ── ALERTAS DE STOCK ──
  sh.setRowHeight(20, 28);
  sh.getRange('B20').setValue('⚠️  ALERTAS DE STOCK  (≤ 5 unidades disponibles)')
    .setFontSize(9).setFontWeight('bold').setFontColor('#B71C1C');
  sh.setRowHeight(21, 30);
  sh.getRange('C21:D21').setValues([['Producto','Stock disponible']])
    .setBackground(BROWN).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');

  sh.getRange('C22')
    .setFormulaLocal('=SI.ERROR(QUERY(Productos!B:E;"SELECT B, E WHERE E <= 5 AND E IS NOT NULL ORDER BY E ASC";0);"✅ Sin stock crítico")');
  sh.getRange('C22:C42').setFontSize(10).setFontColor(BROWN);
  sh.getRange('D22:D42').setFontSize(10).setFontColor('#E65100').setFontWeight('bold').setHorizontalAlignment('center');

  // ── PIE ──
  sh.setRowHeight(46, 24);
  sh.getRange('B46')
    .setFormulaLocal('="Actualizado: "&TEXTO(AHORA();"dd/mm/yyyy HH:mm")')
    .setFontSize(8).setFontColor('#AAAAAA').setFontStyle('italic');
}

// ── Hoja: Pedidos_Proveedores ─────────────────────────────
function _setupProveedores() {
  let sh = SS.getSheetByName('Pedidos_Proveedores');
  if (!sh) sh = SS.insertSheet('Pedidos_Proveedores');

  const headers = ['ID','Fecha','Proveedor','Producto','Cantidad','Unidad',
                   'Precio unit.','Total','Estado','Notas','Fecha entrega'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(BROWN).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.setFrozenRows(1);
  sh.setRowHeight(1, 36);

  [120, 140, 160, 180, 80, 80, 110, 110, 100, 220, 110].forEach((w, i) => {
    sh.setColumnWidth(i + 1, w);
  });

  sh.getRange('G2:H1000').setNumberFormat('$#,##0.##');
  sh.getRange('E2:E1000').setHorizontalAlignment('center');

  // Dropdown de estado
  const dropRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pendiente', 'Pedido', 'Entregado'], true)
    .setAllowInvalid(false).build();
  sh.getRange(2, 9, 1000, 1).setDataValidation(dropRule);

  // Conditional formatting — Estado (col I)
  const iRange = sh.getRange('I2:I1000');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Pendiente')
      .setBackground('#FFF9C4').setFontColor('#7A6000').setBold(true)
      .setRanges([iRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Pedido')
      .setBackground('#BBDEFB').setFontColor('#0D47A1').setBold(true)
      .setRanges([iRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Entregado')
      .setBackground('#C8E6C9').setFontColor('#1B5E20').setBold(true)
      .setRanges([iRange]).build(),
  ]);

  try { sh.getRange('A2:K1000').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false); }
  catch(ex) {}
}

// ── Hoja: Egresos ──────────────────────────────────────────
function _setupEgresos() {
  let sh = SS.getSheetByName('Egresos');
  if (!sh) sh = SS.insertSheet('Egresos');

  const headers = ['Fecha','Mes','Semana','Categoría','Concepto / Detalle','Método','Monto','Nota'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(BROWN).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.setFrozenRows(1);
  sh.setRowHeight(1, 36);
  [100, 80, 70, 140, 230, 120, 100, 210].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  sh.getRange('G2:G1000').setNumberFormat('$#,##0');
  sh.getRange('A2:A1000').setNumberFormat('dd/mm/yyyy');

  // Dropdown categorías
  const catRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Ingredientes','Packaging','Marketing','Transporte','Tecnología','Sueldos','Impuestos','Otros'], true)
    .setAllowInvalid(true).build();
  sh.getRange(2, 4, 1000, 1).setDataValidation(catRule);

  // Dropdown método pago
  const metRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Efectivo','Transferencia','Tarjeta','Otro'], true)
    .setAllowInvalid(false).build();
  sh.getRange(2, 6, 1000, 1).setDataValidation(metRule);

  sh.setTabColor(ORANGE);

  try { sh.getRange('A2:H1000').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false); }
  catch(ex) {}
}
