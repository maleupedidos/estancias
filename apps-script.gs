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
  // Escanear la hoja real para encontrar el máximo ID existente
  // Esto previene saltos cuando se eliminan filas manualmente
  const sheetMap = { 'H-': ['Home', 'B'], 'C-': ['Clubes', 'B'], 'OC-': ['Orden de Compra', 'A'], 'P-': ['Pilar', 'B'], 'CF-': ['Capital Federal', 'B'], 'R-': ['Red', 'B'] };
  const info = sheetMap[prefix];
  if (!info) throw new Error('Prefix desconocido: ' + prefix);

  var max = 0;
  var regex = new RegExp('^' + prefix.replace('-', '\\-') + '(\\d+)$');

  // Buscar máximo en la hoja operativa
  var shData = SS.getSheetByName(info[0]);
  if (shData && shData.getLastRow() > 1) {
    var col = info[1];
    var data = shData.getRange(col + '2:' + col + shData.getLastRow()).getValues();
    for (var i = 0; i < data.length; i++) {
      var match = String(data[i][0]).match(regex);
      if (match) { var n = parseInt(match[1], 10); if (n > max) max = n; }
    }
  }

  // También revisar archivos por si el máximo está ahí
  var archName = 'Archivo ' + info[0];
  var shArch = SS.getSheetByName(archName);
  if (shArch && shArch.getLastRow() > 1) {
    var col2 = info[1];
    var data2 = shArch.getRange(col2 + '2:' + col2 + shArch.getLastRow()).getValues();
    for (var j = 0; j < data2.length; j++) {
      var match2 = String(data2[j][0]).match(regex);
      if (match2) { var n2 = parseInt(match2[1], 10); if (n2 > max) max = n2; }
    }
  }

  // También comparar con Config para nunca retroceder
  var shConfig = SS.getSheetByName('Config');
  if (shConfig) {
    var rowMap = { 'H-': 2, 'C-': 3, 'OC-': 4, 'P-': 5, 'CF-': 6, 'R-': 7 };
    var configVal = Number(shConfig.getRange(rowMap[prefix], 2).getValue()) || 0;
    if (configVal > max) max = configVal;
  }

  var nuevo = max + 1;

  // Actualizar Config
  if (shConfig) {
    var rowMap2 = { 'H-': 2, 'C-': 3, 'OC-': 4, 'P-': 5, 'CF-': 6, 'R-': 7 };
    shConfig.getRange(rowMap2[prefix], 2).setValue(nuevo);
  }

  return prefix + String(nuevo).padStart(3, '0');
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
  var shRed    = SS.getSheetByName('Red');
  var archHome = 0, archPilar = 0, archCaba = 0, archClubes = 0, archRed = 0;

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
  archRed    = _countArchivable(shRed, 11); // Red: Estado Entrega en col L (index 11)

  if (archHome === 0 && archPilar === 0 && archCaba === 0 && archClubes === 0 && archRed === 0) {
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
    '  Clubes: ' + archClubes + ' pedido(s)\n' +
    '  Red: ' + archRed + ' pedido(s)\n\n' +
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

  // ── Archivar Red ──
  if (shRed && archRed > 0) {
    totalMovidas += _archivarHoja(shRed, 'Archivo Red', 11);
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
  if (action === 'entregas') return _doGetEntregas(e);
  if (action === 'busqueda') return _doGetBusqueda();
  if (action === 'catalogo') return _doGetCatalogo();
  if (action === 'admin') return _doGetAdmin();
  if (action === 'ventas') return _doGetVentas();
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

// ══════════════════════════════════════════════════════════════
//  ENTREGAS — API para la página de ruteo (ruta.html)
// ══════════════════════════════════════════════════════════════

/** GET ?action=entregas[&dia=Viernes]
 *  Devuelve pedidos pendientes de entrega de Home, Pilar y Capital Federal.
 *  Formato compacto para señal floja. */
function _doGetEntregas(e) {
  var dia = e && e.parameter && e.parameter.dia;

  var ABBRS = ['PPM','PPJyQ','PPCyQ','SCo','SJyQ','SCa','ECaC','EJyQ','ECyQ','EV',
               'TG','TLC','TC','F','PMu','PMa','PJyQ','PCC','PJyM'];

  var entregas = [];

  // ── Home, Pilar, Capital Federal ──
  ['Home', 'Pilar', 'Capital Federal'].forEach(function(hoja) {
    var sh = SS.getSheetByName(hoja);
    if (!sh || sh.getLastRow() <= 1) return;
    var data = sh.getDataRange().getValues();

    for (var r = 1; r < data.length; r++) {
      var estado = String(data[r][10]).trim();
      if (estado === 'Entregado' || estado === 'Cancelado') continue;

      var diaEntrega = String(data[r][9]).trim();
      if (dia && diaEntrega !== dia) continue;

      var productos = [];
      for (var p = 0; p < 19; p++) {
        var qty = Number(data[r][20 + p]) || 0;
        if (qty > 0) productos.push({ a: ABBRS[p], q: qty });
      }
      if (productos.length === 0) continue;

      var direccion = '', barrio = '', subBarrio = '', lote = '', telefono = '';
      if (hoja === 'Home') {
        barrio = String(data[r][41] || '').trim();
        subBarrio = String(data[r][42] || '').trim();
        lote = String(data[r][43] || '').trim();
        direccion = (subBarrio || barrio) + (lote ? ' · Lote ' + lote : '');
        telefono = String(data[r][44] || '');
      } else if (hoja === 'Pilar') {
        barrio = 'Pilar';
        direccion = [data[r][41], data[r][42]].filter(Boolean).join(' · ');
        telefono = String(data[r][43] || '');
      } else {
        barrio = 'Capital Federal';
        direccion = [data[r][41], [data[r][42], data[r][43]].filter(Boolean).join(' '), data[r][44]].filter(Boolean).join(' · ');
        telefono = String(data[r][45] || '');
      }

      entregas.push({
        id: Number(data[r][1]) || 0,
        h: hoja,
        r: r + 1,
        c: String(data[r][7] || '').trim(),
        t: telefono.trim(),
        d: direccion,
        b: barrio,
        sb: subBarrio || barrio,
        l: lote,
        de: diaEntrega,
        es: estado,
        o: String(data[r][8] || '').trim(),
        fp: String(data[r][11] || '').trim(),
        ep: String(data[r][12] || '').trim(),
        $: Number(data[r][13]) || 0,
        p: productos
      });
    }
  });

  // ── Clubes ──
  var ABBRS_CLUB = ['PMu','PMa','PJyQ','PCC','PJyM','PPM','PPJyQ','PPCyQ'];
  var shClub = SS.getSheetByName('Clubes');
  if (shClub && shClub.getLastRow() > 1) {
    var clubData = shClub.getDataRange().getValues();
    for (var rc = 1; rc < clubData.length; rc++) {
      var estadoC = String(clubData[rc][13]).trim(); // N = Estado de Entrega
      if (estadoC === 'Entregado' || estadoC === 'Cancelado') continue;

      var diaC = String(clubData[rc][12]).trim(); // M = Día de Entrega
      if (dia && diaC !== dia) continue;

      var prodsC = [];
      for (var pc = 0; pc < 8; pc++) {
        var qtyC = Number(clubData[rc][23 + pc]) || 0; // X-AE = cols 23-30
        if (qtyC > 0) prodsC.push({ a: ABBRS_CLUB[pc], q: qtyC });
      }
      if (prodsC.length === 0) continue;

      var club = String(clubData[rc][8] || '').trim();
      var deporte = String(clubData[rc][9] || '').trim();
      var grupo = String(clubData[rc][10] || '').trim();
      var dirClub = [club, deporte, grupo].filter(Boolean).join(' · ');

      entregas.push({
        id: Number(clubData[rc][1]) || 0,
        h: 'Clubes',
        r: rc + 1,
        c: String(clubData[rc][7] || '').trim(),
        t: String(clubData[rc][33] || '').trim(), // AH = Teléfono
        d: dirClub,
        b: club,
        sb: club,
        l: '',
        de: diaC,
        es: estadoC,
        o: String(clubData[rc][11] || '').trim(), // L = Origen
        fp: String(clubData[rc][14] || '').trim(), // O = Forma de Pago
        ep: String(clubData[rc][15] || '').trim(), // P = Estado de Pago
        $: Number(clubData[rc][16]) || 0, // Q = Total
        p: prodsC
      });
    }
  }

  // ── Red ──
  var ABBRS_RED = ['PPM','PPJyQ','PPCyQ','SCo','SJyQ','SCa','ECaC','EJyQ','ECyQ','EV',
                   'TG','TLC','TC','F','SQB','SL','SPyP','SE','PMu','PMa','PJyQ','PCC','PJyM'];
  var shRed = SS.getSheetByName('Red');
  if (shRed && shRed.getLastRow() > 1) {
    var redData = shRed.getDataRange().getValues();
    for (var rr = 1; rr < redData.length; rr++) {
      var estadoR = String(redData[rr][11]).trim(); // L = Estado de Entrega
      if (estadoR === 'Entregado' || estadoR === 'Cancelado') continue;

      var diaR = String(redData[rr][10]).trim(); // K = Día de Entrega
      if (dia && diaR !== dia) continue;

      var prodsR = [];
      for (var pr = 0; pr < 23; pr++) {
        var qtyR = Number(redData[rr][21 + pr]) || 0; // V-AR = cols 21-43
        if (qtyR > 0) prodsR.push({ a: ABBRS_RED[pr], q: qtyR });
      }
      if (prodsR.length === 0) continue;

      var vendedorR = String(redData[rr][7] || '').trim();
      var partidoR = String(redData[rr][48] || '').trim();
      var localidadR = String(redData[rr][49] || '').trim();
      var barrioR = String(redData[rr][50] || '').trim();
      var domicilioR = String(redData[rr][51] || '').trim();
      var dirRed = [vendedorR, partidoR, localidadR, barrioR, domicilioR].filter(Boolean).join(' · ');

      entregas.push({
        id: Number(redData[rr][1]) || 0,
        h: 'Red',
        r: rr + 1,
        c: String(redData[rr][8] || '').trim(), // I = Cliente
        t: String(redData[rr][52] || '').trim(), // BA = Teléfono
        d: dirRed,
        b: 'Red · ' + partidoR,
        sb: localidadR || partidoR,
        l: domicilioR,
        de: diaR,
        es: estadoR,
        o: String(redData[rr][9] || '').trim(), // J = Origen
        fp: String(redData[rr][12] || '').trim(), // M = Forma de Pago
        ep: String(redData[rr][13] || '').trim(), // N = Estado de Pago
        $: Number(redData[rr][14]) || 0, // O = Total
        p: prodsR,
        v: vendedorR // Vendedor (extra para Red)
      });
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ts: Date.now(), e: entregas }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════════
//  BÚSQUEDA — API para la página de búsqueda de productos (busqueda.html)
// ══════════════════════════════════════════════════════════════

/** GET ?action=busqueda
 *  Devuelve OC pendientes/pedidas agrupadas por proveedor y por cliente.
 *  Formato compacto para PWA offline-first. */
function _doGetBusqueda() {
  var shOC = SS.getSheetByName('Orden de Compra');
  if (!shOC || shOC.getLastRow() <= 1) {
    return ContentService.createTextOutput(JSON.stringify({ ts: Date.now(), provs: [], clientes: [], total: 0 }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = shOC.getDataRange().getValues();

  function splitProducto(nombre) {
    var parts = nombre.split(' \u2014 ');
    return { cat: (parts[0] || nombre).trim(), var: (parts[1] || '').trim() };
  }

  var porProv = {};
  var porCliente = {};
  var totalGeneral = 0;
  var ocRows = []; // para poder marcar estado desde la PWA

  for (var r = 1; r < data.length; r++) {
    var estado = String(data[r][20]).trim();
    if (estado !== 'Pendiente' && estado !== 'Pedido' && estado !== 'Recibido') continue;
    var origen = String(data[r][19]).trim();
    if (origen === 'Dep\u00f3sito') continue;

    var esRecibido = (estado === 'Recibido');

    var ocNum    = String(data[r][0]).trim();
    var prov     = String(data[r][9]).trim();
    var producto = String(data[r][10]).trim();
    var abbr     = String(data[r][11]).trim();
    var qty      = Number(data[r][12]) || 0;
    var costoU   = Number(data[r][13]) || 0;
    var costoT   = costoU * qty;
    var precioV  = Number(data[r][15]) || 0;
    var ingresoT = precioV * qty;
    var cliente  = String(data[r][6]).trim();
    var canal    = String(data[r][4]).trim();
    var dir      = String(data[r][8]).trim();
    var tel      = String(data[r][7]).trim();
    var nPedido  = String(data[r][5]).trim();

    if (!prov || qty === 0) continue;

    // Proveedor agrupado (solo Pendiente/Pedido — lo que falta comprar/recibir)
    if (!esRecibido) {
      if (!porProv[prov]) porProv[prov] = { costo: 0, cats: {} };
      var sp = splitProducto(producto);
      if (!porProv[prov].cats[sp.cat]) porProv[prov].cats[sp.cat] = [];
      var found = false;
      for (var v = 0; v < porProv[prov].cats[sp.cat].length; v++) {
        if (porProv[prov].cats[sp.cat][v].v === sp.var) {
          porProv[prov].cats[sp.cat][v].q += qty;
          porProv[prov].cats[sp.cat][v].ct += costoT;
          found = true; break;
        }
      }
      if (!found) porProv[prov].cats[sp.cat].push({ v: sp.var, q: qty, ct: costoT });
      porProv[prov].costo += costoT;
      totalGeneral += costoT;
    }

    // Cliente agrupado (Pendiente + Pedido + Recibido — todo lo que hay que armar/entregar)
    var cKey = cliente + '|' + nPedido;
    if (!porCliente[cKey]) porCliente[cKey] = { n: cliente, canal: canal, dir: dir, tel: tel, ped: nPedido, items: [], ing: 0 };
    porCliente[cKey].items.push({ prod: producto, q: qty, pv: precioV * qty, est: estado });
    porCliente[cKey].ing += ingresoT;

    // OC rows para control de estado
    ocRows.push({ oc: ocNum, r: r + 1, prov: prov, prod: producto, abbr: abbr, q: qty, est: estado, canal: canal });
  }

  // Armar arrays
  var provsArr = [];
  var provNames = Object.keys(porProv);
  provNames.forEach(function(prov) {
    var d = porProv[prov];
    var catsArr = [];
    Object.keys(d.cats).forEach(function(cat) {
      catsArr.push({ cat: cat, items: d.cats[cat] });
    });
    // Generar texto WhatsApp
    var waLines = ['Hola! Te paso el pedido de esta semana:', ''];
    catsArr.forEach(function(c) {
      waLines.push(c.cat + ':');
      c.items.forEach(function(it) { waLines.push('  ' + it.q + ' \u00d7 ' + (it.v || c.cat)); });
      waLines.push('');
    });
    waLines.push('Total: $' + d.costo.toLocaleString('es-AR'));
    waLines.push('');
    waLines.push('Gracias!');
    provsArr.push({ n: prov, costo: d.costo, cats: catsArr, wa: waLines.join('\n') });
  });

  var clientesArr = [];
  Object.keys(porCliente).forEach(function(k) {
    clientesArr.push(porCliente[k]);
  });

  // ── DEUDAS: OC Recibidas no pagadas, agrupadas por proveedor y semana ──
  var deudas = {};
  for (var rd = 1; rd < data.length; rd++) {
    var estD = String(data[rd][20]).trim();
    if (estD !== 'Recibido') continue;
    var pagado = String(data[rd][23]).trim(); // X = Pagado Proveedor
    if (pagado === 'Sí' || pagado === 'Si') continue;

    var provD   = String(data[rd][9]).trim();
    var semD    = String(data[rd][2]).trim(); // C = Semana
    var costoTD = Number(data[rd][14]) || 0;  // O = Costo Total
    var prodD   = String(data[rd][10]).trim();
    var qtyD    = Number(data[rd][12]) || 0;

    if (!provD || costoTD === 0) continue;
    if (!deudas[provD]) deudas[provD] = { total: 0, semanas: {}, rows: [] };
    deudas[provD].total += costoTD;
    if (!deudas[provD].semanas[semD]) deudas[provD].semanas[semD] = 0;
    deudas[provD].semanas[semD] += costoTD;
    deudas[provD].rows.push({ r: rd + 1, prod: prodD, q: qtyD, costo: costoTD, sem: semD });
  }

  var deudasArr = [];
  Object.keys(deudas).forEach(function(prov) {
    var d = deudas[prov];
    var semsArr = [];
    Object.keys(d.semanas).sort().forEach(function(s) {
      semsArr.push({ sem: s, monto: d.semanas[s] });
    });
    deudasArr.push({ n: prov, total: d.total, semanas: semsArr, rows: d.rows });
  });

  return ContentService
    .createTextOutput(JSON.stringify({
      ts: Date.now(),
      provs: provsArr,
      clientes: clientesArr,
      total: totalGeneral,
      ocs: ocRows,
      deudas: deudasArr
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** POST action=pagarProveedor — registra pago a proveedor.
 *  Recibe { action:'pagarProveedor', proveedor:'Sevuchitas', efectivo:3000, mp:2000, rows:[5,6,7], notas:'...' }
 *  - Crea 1-2 filas en Egresos (una por método con monto > 0)
 *  - Crea 1-2 filas en Movimientos Historicos
 *  - Marca OC rows como Pagado Proveedor = "Sí" */
function _doPostPagarProveedor(postData) {
  var proveedor = String(postData.proveedor || '').trim();
  var montoEf = Number(postData.efectivo) || 0;
  var montoMp = Number(postData.mp) || 0;
  var rows = postData.rows || [];
  var notas = String(postData.notas || '').trim();
  var totalPago = montoEf + montoMp;

  if (!proveedor || totalPago <= 0) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'Sin monto' })).setMimeType(ContentService.MimeType.JSON);
  }

  var ahora = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var semana = _isoWeek(argNow);
  var mes = MESES[argNow.getMonth()];

  // ── Registrar en Egresos ──
  var shEg = SS.getSheetByName('Egresos');
  if (!shEg) {
    shEg = SS.insertSheet('Egresos');
    shEg.getRange(1, 1, 1, 8).setValues([['Fecha', 'Semana', 'Mes', 'Categoria', 'Concepto', 'Metodo', 'Monto', 'Notas']]);
    shEg.setFrozenRows(1);
  }

  var concepto = 'Pago a ' + proveedor;
  var pagos = [];
  if (montoEf > 0) pagos.push({ metodo: 'Efectivo', monto: montoEf });
  if (montoMp > 0) pagos.push({ metodo: 'Mercado Pago', monto: montoMp });

  pagos.forEach(function(p) {
    shEg.appendRow([argNow, semana, mes, 'Proveedor', concepto, p.metodo, p.monto, notas]);
    var lastRow = shEg.getLastRow();
    shEg.getRange(lastRow, 1).setNumberFormat('dd/MM/yyyy');
    shEg.getRange(lastRow, 7).setNumberFormat('$#,##0');
  });

  // ── Registrar en Movimientos Historicos ──
  var shMH = SS.getSheetByName('Movimientos Historicos');
  if (shMH) {
    pagos.forEach(function(p) {
      shMH.appendRow([argNow, mes, semana, 'Proveedor', concepto, p.metodo, p.monto, notas]);
      var lastRow = shMH.getLastRow();
      shMH.getRange(lastRow, 1).setNumberFormat('dd/MM/yyyy');
      shMH.getRange(lastRow, 7).setNumberFormat('$#,##0');
    });
  }

  // ── Marcar OC rows como Pagado ──
  if (rows.length > 0) {
    var shOC = SS.getSheetByName('Orden de Compra');
    if (shOC) {
      rows.forEach(function(r) {
        var row = Number(r);
        if (row > 1) shOC.getRange(row, 24).setValue('Sí'); // X = Pagado Proveedor
      });
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true, total: totalPago, filas: pagos.length })).setMimeType(ContentService.MimeType.JSON);
}

/** POST action=marcarOC — cambia estado de OC desde busqueda.html.
 *  Recibe { action:'marcarOC', rows:[{r:5, estado:'Pedido'}, ...] } */
function _doPostMarcarOC(postData) {
  var shOC = SS.getSheetByName('Orden de Compra');
  if (!shOC) return ContentService.createTextOutput(JSON.stringify({ ok: false })).setMimeType(ContentService.MimeType.JSON);

  var rows = postData.rows || [];
  var ahora = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));

  rows.forEach(function(item) {
    var row = Number(item.r);
    var nuevoEstado = String(item.estado || '').trim();
    if (!row || !nuevoEstado) return;

    // Col U (21) = Estado OC
    shOC.getRange(row, 21).setValue(nuevoEstado);

    // Si "Pedido" → Col V (22) = Fecha Pedido Prov
    if (nuevoEstado === 'Pedido') {
      shOC.getRange(row, 22).setValue(argNow);
    }
    // Si "Recibido" → Col W (23) = Fecha Recibido
    if (nuevoEstado === 'Recibido') {
      shOC.getRange(row, 23).setValue(argNow);
    }
  });

  return ContentService.createTextOutput(JSON.stringify({ ok: true, updated: rows.length })).setMimeType(ContentService.MimeType.JSON);
}

/** POST action=recibirMercaderia — marca OC como Recibido con cantidad real.
 *  Recibe { action:'recibirMercaderia', items:[{r:5, qtyRecibida:7}, {r:6, qtyRecibida:0}] }
 *  Si qtyRecibida = 0, la fila se marca como "Cancelado" (no llegó).
 *  Si qtyRecibida < qtyOriginal, se actualiza la cantidad y recalculan costos. */
function _doPostRecibirMercaderia(postData) {
  var shOC = SS.getSheetByName('Orden de Compra');
  if (!shOC) return ContentService.createTextOutput(JSON.stringify({ ok: false })).setMimeType(ContentService.MimeType.JSON);

  var items = postData.items || [];
  var ahora = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var updated = 0;

  items.forEach(function(item) {
    var row = Number(item.r);
    var qtyRecibida = Number(item.qtyRecibida);
    if (!row) return;

    if (qtyRecibida <= 0) {
      // No llegó nada: cancelar esta línea de OC
      shOC.getRange(row, 21).setValue('Cancelado'); // U = Estado OC
      shOC.getRange(row, 23).setValue(argNow);      // W = Fecha Recibido
      shOC.getRange(row, 13).setValue(0);            // M = Cantidad → 0
      shOC.getRange(row, 15).setValue(0);            // O = Costo Total → 0
    } else {
      // Actualizar cantidad recibida
      var costoUnit = Number(shOC.getRange(row, 14).getValue()) || 0; // N = Costo Unitario
      shOC.getRange(row, 13).setValue(qtyRecibida);                   // M = Cantidad
      shOC.getRange(row, 15).setValue(costoUnit * qtyRecibida);       // O = Costo Total
      shOC.getRange(row, 21).setValue('Recibido');                    // U = Estado OC
      shOC.getRange(row, 23).setValue(argNow);                        // W = Fecha Recibido
    }
    updated++;
  });

  return ContentService.createTextOutput(JSON.stringify({ ok: true, updated: updated })).setMimeType(ContentService.MimeType.JSON);
}

/** GET ?action=catalogo
 *  Devuelve proveedores, productos por proveedor, costos y stock.
 *  Para el formulario de "Nuevo Pedido" en busqueda.html. */
function _doGetCatalogo() {
  var hProv = SS.getSheetByName('Proveedores');
  var hProd = SS.getSheetByName('Productos');
  if (!hProv) return ContentService.createTextOutput(JSON.stringify({ proveedores: [], productos: {} })).setMimeType(ContentService.MimeType.JSON);

  // Costos y stock desde Productos
  var costoMap = {}, stockMap = {};
  if (hProd) {
    var prodData = hProd.getDataRange().getValues();
    for (var r = 1; r < prodData.length; r++) {
      var ab = String(prodData[r][2]).trim();
      if (!ab) continue;
      costoMap[ab] = parseFloat(String(prodData[r][9]).replace(/[$.]/g,'').replace(/,/g,'')) || 0;
      stockMap[ab] = Number(prodData[r][5]) || 0;
    }
  }

  // Proveedores y productos
  var provData = hProv.getDataRange().getValues();
  var proveedores = [], seenProv = {}, productosPorProv = {};
  var lastProv = '', lastProd = '';
  for (var r2 = 1; r2 < provData.length; r2++) {
    if (provData[r2][2] && String(provData[r2][2]).trim()) lastProv = String(provData[r2][2]).trim();
    if (provData[r2][1] && String(provData[r2][1]).trim()) lastProd = String(provData[r2][1]).trim();
    if (!lastProv) continue;
    if (!seenProv[lastProv]) { seenProv[lastProv] = true; proveedores.push(lastProv); }

    var abbr = String(provData[r2][4]).trim();
    var gusto = String(provData[r2][3]).trim();
    if (!abbr) continue;

    if (!productosPorProv[lastProv]) productosPorProv[lastProv] = [];
    productosPorProv[lastProv].push({
      a: abbr,
      n: lastProd + (gusto ? ' \u2014 ' + gusto : ''),
      c: costoMap[abbr] || 0,
      s: stockMap[abbr] !== undefined ? stockMap[abbr] : null
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ts: Date.now(), proveedores: proveedores, productos: productosPorProv }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** POST action=compraManual — crea filas en Orden de Compra desde busqueda.html.
 *  Replica la lógica de confirmarCompraDeposito pero vía API web.
 *  Recibe { action:'compraManual', vendedor:'Tadeo — Stock', proveedor:'Sevuchitas', items:[{abbr,nombre,costo,qty,origen}] } */
function _doPostCompraManual(postData) {
  var shOC = SS.getSheetByName('Orden de Compra');
  if (!shOC) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Hoja OC no encontrada' })).setMimeType(ContentService.MimeType.JSON);

  var proveedor = String(postData.proveedor || '').trim();
  var vendedor = String(postData.vendedor || 'Tadeo \u2014 Stock').trim();
  var items = postData.items || [];
  var esMarcos = vendedor === 'Marcos Bottcher';

  var itemsValidos = items.filter(function(it) { return (Number(it.qty) || 0) > 0; });
  if (itemsValidos.length === 0) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Sin productos' })).setMimeType(ContentService.MimeType.JSON);

  // Precios de venta desde Productos
  var precioVentaMap = {};
  var hProd = SS.getSheetByName('Productos');
  if (hProd) {
    var prData = hProd.getDataRange().getValues();
    for (var p = 1; p < prData.length; p++) {
      var ab = String(prData[p][2]).trim();
      if (ab) precioVentaMap[ab] = parseFloat(String(prData[p][8]).replace(/[$.]/g,'').replace(/,/g,'')) || 0;
    }
  }

  // Timestamp Argentina
  var ahora = new Date();
  var argDate = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var dd = String(argDate.getDate()).padStart(2, '0');
  var mm = String(argDate.getMonth() + 1).padStart(2, '0');
  var yyyy = argDate.getFullYear();
  var hh = String(argDate.getHours()).padStart(2, '0');
  var mi = String(argDate.getMinutes()).padStart(2, '0');
  var fechaStr = dd + '/' + mm + '/' + yyyy + ' ' + hh + ':' + mi;
  var semana = _isoWeek(argDate);
  var mesNombre = MESES[argDate.getMonth()];

  var newRows = [];
  itemsValidos.forEach(function(item) {
    var ocId = _nextId('OC-');
    var costo = Number(item.costo) || 0;
    var qty = Number(item.qty) || 0;
    var costoTotal = costo * qty;
    newRows.push([
      ocId,                                         // A  N° Orden
      fechaStr,                                     // B  Fecha Creación
      semana,                                       // C  Semana
      mesNombre,                                    // D  Mes
      esMarcos ? 'Red' : 'Dep\u00f3sito',          // E  Canal
      '',                                           // F  N° Pedido Origen
      vendedor,                                     // G  Cliente
      '',                                           // H  Teléfono
      esMarcos ? 'Tortugas, Gar\u00edn' : 'Dep\u00f3sito Maleu', // I  Dirección
      proveedor,                                    // J  Proveedor
      String(item.nombre || ''),                    // K  Producto
      String(item.abbr || ''),                      // L  Abreviatura
      qty,                                          // M  Cantidad
      costo,                                        // N  Costo Unitario
      costoTotal,                                   // O  Costo Total
      precioVentaMap[item.abbr] || 0,               // P  Precio Venta Unit.
      0,                                            // Q  Ingreso Total (fórmula)
      0,                                            // R  Margen Bruto $ (fórmula)
      0,                                            // S  Margen % (fórmula)
      String(item.origen || 'Orden de Compra'),     // T  Origen
      'Pedido',                                     // U  Estado OC
      dd + '/' + mm + '/' + yyyy,                   // V  Fecha Pedido Prov
      '',                                           // W  Fecha Recibido
      'No',                                         // X  Pagado Proveedor
      'No',                                         // Y  Cobrado Cliente
    ]);
  });

  var startRow = shOC.getLastRow() + 1;
  shOC.getRange(startRow, 1, newRows.length, 25).setValues(newRows);

  // Fórmulas financieras
  for (var i = 0; i < newRows.length; i++) {
    var r = startRow + i;
    shOC.getRange(r, 17).setFormula('=P' + r + '*M' + r);
    shOC.getRange(r, 18).setFormula('=Q' + r + '-O' + r);
    shOC.getRange(r, 19).setFormula('=R' + r + '/Q' + r);
  }

  // Formato
  shOC.getRange(startRow, 13, newRows.length, 1).setNumberFormat('0');
  shOC.getRange(startRow, 14, newRows.length, 2).setNumberFormat('$#,##0');
  shOC.getRange(startRow, 16, newRows.length, 3).setNumberFormat('$#,##0');
  shOC.getRange(startRow, 19, newRows.length, 1).setNumberFormat('0.0%');

  var totalCosto = newRows.reduce(function(s, row) { return s + (row[14] || 0); }, 0);
  return ContentService.createTextOutput(JSON.stringify({
    ok: true, filas: newRows.length, totalCosto: totalCosto,
    primerOC: newRows[0][0], ultimoOC: newRows[newRows.length - 1][0]
  })).setMimeType(ContentService.MimeType.JSON);
}

/** POST action=marcarEntregado — marca un pedido como Entregado desde ruta.html.
 *  Replica la lógica de _onEditHome: fecha de entrega + descuento de stock si Depósito. */
function _doPostMarcarEntregado(data) {
  var hoja = String(data.hoja || '');
  var pedidoId = String(data.id || '');
  var cobrado = !!data.cobrado;

  var sh = SS.getSheetByName(hoja);
  if (!sh) return ContentService.createTextOutput(JSON.stringify({ ok: false })).setMimeType(ContentService.MimeType.JSON);

  // Buscar fila por N° Pedido (col B=2). Usar row hint si coincide, sino buscar.
  var row = Number(data.row) || 0;
  if (row > 1 && String(sh.getRange(row, 2).getValue()) === pedidoId) {
    // row hint es correcto
  } else {
    var allData = sh.getDataRange().getValues();
    row = -1;
    for (var r = 1; r < allData.length; r++) {
      if (String(allData[r][1]) === pedidoId) { row = r + 1; break; }
    }
    if (row === -1) return ContentService.createTextOutput(JSON.stringify({ ok: false })).setMimeType(ContentService.MimeType.JSON);
  }

  // Columnas según hoja (Clubes y Red tienen layout diferente)
  var isClub = (hoja === 'Clubes');
  var isRed  = (hoja === 'Red');
  var colEstadoEntrega = isClub ? 14 : isRed ? 12 : 11;   // N / L / K
  var colOrigen        = isClub ? 12 : isRed ? 10 : 9;    // L / J / I
  var colEstadoPago    = isClub ? 16 : isRed ? 14 : 13;   // P / N / M
  var colFormaPago     = isClub ? 15 : isRed ? 13 : 12;   // O / M / L
  var colPropinaEf     = isClub ? 21 : isRed ? 19 : 18;   // U / S / R
  var colPropinaTr     = isClub ? 22 : isRed ? 20 : 19;   // V / T / S

  // Verificar que no esté ya Entregado
  var estadoActual = String(sh.getRange(row, colEstadoEntrega).getValue()).trim();
  if (estadoActual === 'Entregado') {
    return ContentService.createTextOutput(JSON.stringify({ ok: true, ya: true })).setMimeType(ContentService.MimeType.JSON);
  }

  // Marcar Entregado
  sh.getRange(row, colEstadoEntrega).setValue('Entregado');

  // Registrar fecha de entrega (solo Home/Pilar/CF tienen columnas de fecha entrega)
  if (!isClub && !isRed) {
    var colEntrega = hoja === 'Pilar' ? 45 : hoja === 'Capital Federal' ? 47 : 46;
    _registrarFechaEntrega(sh, row, colEntrega);
  }

  // Descontar stock según Origen
  var origen = String(sh.getRange(row, colOrigen).getValue()).trim();
  if (origen === 'Depósito') {
    var hProd = SS.getSheetByName('Productos');
    if (hProd) _homeStockFisico(sh, row, hProd, -1);
  } else if (origen === 'Mixto') {
    var hProd = SS.getSheetByName('Productos');
    if (hProd) _homeStockFisicoMixto(sh, row, hProd, -1);
  }

  // Marcar cobrado si se indicó
  if (cobrado) sh.getRange(row, colEstadoPago).setValue('Cobrado');

  // Escribir propina si existe
  var propina = Number(data.propina) || 0;
  if (propina > 0) {
    var formaPago = String(sh.getRange(row, colFormaPago).getValue()).trim();
    if (formaPago === 'Efectivo') {
      sh.getRange(row, colPropinaEf).setValue(propina);
    } else {
      sh.getRange(row, colPropinaTr).setValue(propina);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  doPost — recibe pedidos desde la página + acciones internas
// ════════════════════════════════════════════════════════════
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'compra')          return _doPostCompra(data);
    if (data.action === 'compraLote')     return _doPostCompraLote(data);
    if (data.action === 'updateCompra')   return _doUpdateCompra(data);
    if (data.action === 'marcarEntregado') return _doPostMarcarEntregado(data);
    if (data.action === 'gasto')           return _doPostGasto(data);
    if (data.action === 'ingreso')         return _doPostIngreso(data);
    if (data.action === 'ajusteSaldo')     return _doPostAjusteSaldo(data);
    if (data.action === 'marcarCobrado')   return _doPostMarcarCobrado(data);
    if (data.action === 'setOrigenProductos') return _doPostSetOrigenProductos(data);
    if (data.action === 'marcarOC')        return _doPostMarcarOC(data);
    if (data.action === 'recibirMercaderia') return _doPostRecibirMercaderia(data);
    if (data.action === 'pagarProveedor')  return _doPostPagarProveedor(data);
    if (data.action === 'compraManual')    return _doPostCompraManual(data);
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
    else if (canal === 'Red')        _doPostRed(data);
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

  // ── N° de pedido = cantidad de pedidos existentes + 1 (simple: 1, 2, 3...)
  const orderNum = sh.getLastRow(); // fila 2 = pedido 1, fila 3 = pedido 2, etc.

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
  row[40] = 0;                                   // AO  Margen Bruto (fórmula se pone después)

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
    var tel = String(data.telefono || '');
    row[44] = tel;                                  // AS  Teléfono
  }

  // Subtotal sin descuento y Descuento al final de la fila
  // Se agregan después de las columnas de entrega (AZ y BA, índices 51 y 52)
  while (row.length < 53) row.push('');
  row[51] = Number(data.subtotalSinDescuento) || total; // AZ = Subtotal sin Descuento
  row[52] = Number(data.descuento) || 0;                // BA = Descuento ($)

  sh.appendRow(row);
  var newRow = sh.getLastRow();

  // Fórmula Facturado en T (col 20) = Total + Propinas
  sh.getRange(newRow, 20).setFormula('=N' + newRow + '+R' + newRow + '+S' + newRow);
  // Fórmula Margen Bruto en AO (col 41) = Facturado - Costo
  sh.getRange(newRow, 41).setFormula('=T' + newRow + '-AN' + newRow);

  // Forzar teléfono como texto (evitar que Sheets lo interprete como fórmula/número)
  var telCol = (sheetName === 'Capital Federal') ? 46 : (sheetName === 'Pilar') ? 44 : 45;
  var telVal = String(data.telefono || '');
  if (telVal) sh.getRange(newRow, telCol).setNumberFormat('@').setValue(telVal);

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

  // ── N° de pedido = cantidad de pedidos existentes + 1
  const orderNum = sh.getLastRow();

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
  row[32] = 0;                                 // AG  Margen Bruto (fórmula se pone después)
  row[33] = String(data.telefono || '');       // AH  Teléfono

  sh.appendRow(row);
  var newRow = sh.getLastRow();
  // Fórmula Facturado en W (col 23) = Total + Propinas
  sh.getRange(newRow, 23).setFormula('=Q' + newRow + '+U' + newRow + '+V' + newRow);
  // Fórmula Margen Bruto en AG (col 33) = Facturado - Costo
  sh.getRange(newRow, 33).setFormula('=W' + newRow + '-AF' + newRow);
  // Forzar teléfono como texto
  var telClubes = String(data.telefono || '');
  if (telClubes) sh.getRange(newRow, 34).setNumberFormat('@').setValue(telClubes);
}

// ════════════════════════════════════════════════════════════
//  _doPostRed — escribe en la hoja "Red" (canal vendedores independientes)
//  55 columnas A–BC: incluye Vendedor, 23 productos, comisión 17%
// ════════════════════════════════════════════════════════════

// Mapeo ID producto (web) → columna 1-based hoja Red
// Productos empiezan en col V(22), 23 productos hasta AR(44)
const RED_PRODUCT_COLS = {
  5:  22,  // PPM
  6:  23,  // PPJyQ
  7:  24,  // PPCyQ
  8:  25,  // SCo
  9:  26,  // SJyQ
  10: 27,  // SCa
  11: 28,  // ECaC
  12: 29,  // EJyQ
  17: 30,  // ECyQ
  18: 31,  // EV
  14: 32,  // TG
  15: 33,  // TLC
  16: 34,  // TC
  13: 35,  // F
  20: 36,  // SQB
  21: 37,  // SL
  22: 38,  // SPyP
  23: 39,  // SE
  19: 40,  // PMu
  1:  41,  // PMa
  2:  42,  // PJyQ
  3:  43,  // PCC
  4:  44,  // PJyM
};

// Mapeo ID producto (web Red) → abreviatura en hoja Productos
const RED_ID_TO_ABBR = {
  5:'PPM', 6:'PPJyQ', 7:'PPCyQ',
  8:'SCo', 9:'SJyQ', 10:'SCa',
  11:'ECaC', 12:'EJyQ', 17:'ECyQ', 18:'EV',
  14:'TG', 15:'TLC', 16:'TC', 13:'F',
  20:'SQB', 21:'SL', 22:'SPyP', 23:'SE',
  19:'PMu', 1:'PMa', 2:'PJyQ', 3:'PCC', 4:'PJyM',
};

function _doPostRed(data) {
  const sh = SS.getSheetByName('Red');
  if (!sh) return;

  const orderNum = _nextId('R-');

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
  const envio         = Number(data.envio) || 0;
  const total         = Number(data.total) || 0;
  const pago          = String(data.pago || '');
  const totalSinEnvio = total - envio;
  const efectivo      = pago === 'Efectivo'      ? totalSinEnvio : 0;
  const transferencia = pago === 'Transferencia' ? totalSinEnvio : 0;

  // Cantidades
  const qtys = {};
  (data.items || []).forEach(function(item) {
    qtys[Number(item.id)] = Number(item.qty) || 0;
  });

  // Costo desde hoja Productos
  let costoTotal = 0;
  const hProductos = SS.getSheetByName('Productos');
  if (hProductos) {
    const prodData = hProductos.getDataRange().getValues();
    (data.items || []).forEach(function(item) {
      const abbr = RED_ID_TO_ABBR[Number(item.id)];
      if (!abbr) return;
      for (let r = 1; r < prodData.length; r++) {
        if (String(prodData[r][2]).trim() === abbr) {
          const costoUnit = Number(prodData[r][9]) || 0; // col J = Costo
          costoTotal += costoUnit * (Number(item.qty) || 0);
          break;
        }
      }
    });
  }

  // Construir fila de 55 columnas (A a BC)
  const row = new Array(55).fill('');
  row[0]  = horaStr;                                // A  Hora
  row[1]  = orderNum;                               // B  N° Pedido
  row[2]  = diaNombre;                              // C  Día
  row[3]  = fechaStr;                               // D  Fecha
  row[4]  = mes;                                    // E  Mes
  row[5]  = semana;                                 // F  Semana
  row[6]  = yyyy;                                   // G  Año
  row[7]  = String(data.vendedor || '');            // H  Vendedor
  row[8]  = String(data.nombre || '');              // I  Cliente
  row[9]  = 'Pendiente';                            // J  Origen
  row[10] = String(data.dia || '');                 // K  Día de Entrega
  row[11] = 'Pendiente';                            // L  Estado de Entrega
  row[12] = pago;                                   // M  Forma de Pago
  row[13] = 'No Cobrado';                           // N  Estado de Pago
  row[14] = total;                                  // O  Total ($)
  row[15] = envio;                                  // P  Envío ($)
  row[16] = efectivo;                               // Q  Efectivo
  row[17] = transferencia;                          // R  Transferencia
  row[18] = 0;                                      // S  Propina Ef
  row[19] = 0;                                      // T  Propina Trans
  // U (col 21) = Facturado → fórmula después

  // Productos: cols V–AR (índices 21–43 en base-0)
  Object.keys(RED_PRODUCT_COLS).forEach(function(id) {
    row[RED_PRODUCT_COLS[id] - 1] = qtys[Number(id)] || 0;
  });

  row[44] = costoTotal;                             // AS  Costo
  // AT (col 46) = Margen Bruto → fórmula después
  // AU (col 47) = Comisión 17% → fórmula después
  // AV (col 48) = Margen Neto → fórmula después
  row[48] = String(data.partido || '');             // AW  Partido
  row[49] = String(data.localidad || '');           // AX  Localidad
  row[50] = String(data.barrioRed || '');           // AY  Barrio
  row[51] = String(data.domicilioRed || '');        // AZ  Domicilio
  row[52] = String(data.telefono || '');            // BA  Teléfono
  row[53] = Number(data.subtotalSinDescuento) || total; // BB  Subtotal sin Descuento
  row[54] = Number(data.descuento) || 0;            // BC  Descuento

  sh.appendRow(row);
  var newRow = sh.getLastRow();

  // Fórmula Facturado en U (col 21) = Total + Propinas
  sh.getRange(newRow, 21).setFormula('=O' + newRow + '+S' + newRow + '+T' + newRow);
  // Fórmula Margen Bruto en AT (col 46) = Facturado - Costo
  sh.getRange(newRow, 46).setFormula('=U' + newRow + '-AS' + newRow);
  // Fórmula Comisión 17% en AU (col 47) = Facturado * 17/100
  sh.getRange(newRow, 47).setFormula('=U' + newRow + '*17/100');
  // Fórmula Margen Neto en AV (col 48) = Margen Bruto - Comisión
  sh.getRange(newRow, 48).setFormula('=AT' + newRow + '-AU' + newRow);
  // Forzar teléfono como texto
  var telRed = String(data.telefono || '');
  if (telRed) sh.getRange(newRow, 53).setNumberFormat('@').setValue(telRed);
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

// Red: col V(22) a AR(44) = 23 productos
const RED_COL_TO_ABBR = {
  22: 'PPM',   // V
  23: 'PPJyQ', // W
  24: 'PPCyQ', // X
  25: 'SCo',   // Y
  26: 'SJyQ',  // Z
  27: 'SCa',   // AA
  28: 'ECaC',  // AB
  29: 'EJyQ',  // AC
  30: 'ECyQ',  // AD
  31: 'EV',    // AE
  32: 'TG',    // AF
  33: 'TLC',   // AG
  34: 'TC',    // AH
  35: 'F',     // AI
  36: 'SQB',   // AJ
  37: 'SL',    // AK
  38: 'SPyP',  // AL
  39: 'SE',    // AM
  40: 'PMu',   // AN
  41: 'PMa',   // AO
  42: 'PJyQ',  // AP
  43: 'PCC',   // AQ
  44: 'PJyM',  // AR
};

// IMPORTANTE: esta función debe configurarse SOLO como trigger instalable.
// NO usar el nombre "onEdit" para evitar doble ejecución (simple + instalable).
// En Activadores: función = onEditHandler, evento = Al editar
function onEditHandler(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  if (sheetName === 'Home' || sheetName === 'Pilar' || sheetName === 'Capital Federal') return _onEditHome(e);
  if (sheetName === 'Red')             return _onEditRed(e);
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

// ── Red: Origen col J(10), Estado Entrega col L(12) ──
function _onEditRed(e) {
  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row <= 1) return;

  const sh = e.range.getSheet();

  // Si cambió Origen (col J=10) a "Orden de Compra" → generar OC
  if (col === 10) {
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
          generarOrdenDeCompra('Red', row);
          SS.toast('Orden de Compra generada para ' + pedidoNum, 'OC', 5);
        } finally {
          lock.releaseLock();
        }
      }
    }
    return;
  }

  if (col !== 12) return; // solo col L (12) = Estado de Entrega

  const origen   = String(sh.getRange(row, 10).getValue()); // col J (10) = Origen
  const nuevo    = String(e.value || '');
  const anterior = String(e.oldValue || '');

  // → Entregado: descontar stock si Depósito (no hay columnas de fecha entrega en Red)
  if (nuevo === 'Entregado' && anterior !== 'Entregado') {
    if (origen === 'Depósito') {
      const hProductos = SS.getSheetByName('Productos');
      if (hProductos) _redStockFisico(sh, row, hProductos, -1);
    }
  }
  // ← Sale de Entregado: devolver stock
  if (anterior === 'Entregado' && nuevo !== 'Entregado') {
    if (origen === 'Depósito') {
      const hProductos = SS.getSheetByName('Productos');
      if (hProductos) _redStockFisico(sh, row, hProductos, +1);
    }
  }
}

// Stock para Red: 23 productos en cols V(22)-AR(44)
function _redStockFisico(shRed, row, hProductos, signo) {
  const cantidades = shRed.getRange(row, 22, 1, 23).getValues()[0]; // cols V–AR
  const prodData   = hProductos.getDataRange().getValues();
  const refPedido  = String(shRed.getRange(row, 2).getValue() || '');

  Object.keys(RED_COL_TO_ABBR).forEach(function(colStr) {
    const colIdx = Number(colStr);
    const abbr   = RED_COL_TO_ABBR[colIdx];
    const qty    = Number(cantidades[colIdx - 22]) || 0;
    if (qty === 0) return;

    for (let r = 1; r < prodData.length; r++) {
      if (String(prodData[r][2]).trim() === abbr) {
        const celdaFis = hProductos.getRange(r + 1, 6);
        const fisico   = Number(celdaFis.getValue()) || 0;
        const nuevoStock = Math.max(0, fisico + (qty * signo));
        celdaFis.setValue(nuevoStock);
        _logKardex(abbr, signo < 0 ? '-SAL' : '+DEV', qty, fisico, nuevoStock, 'Red', refPedido);
        break;
      }
    }
  });
}

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
        if (signo < 0 && nuevoStock <= 3 && nuevoStock >= 0) {
          SS.toast('⚠️ STOCK BAJO: ' + abbr + ' → quedan ' + nuevoStock, 'Alerta Stock', 8);
        }
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
  } else if (canal === 'Red') {
    colCliente = 8; colPedido = 1; colTelefono = 52; // BA(53) = Teléfono
    direccion = [rowData[48], rowData[49], rowData[50], rowData[51]].filter(Boolean).join(' · '); // AW=Partido, AX=Localidad, AY=Barrio, AZ=Domicilio
  } else if (canal === 'Clubes') {
    colCliente = 7; colPedido = 1; colTelefono = 33; // AH(34) = Teléfono
    direccion = [rowData[8], rowData[9], rowData[10]].filter(Boolean).join(' · ');
  }

  // Cargar todos los lookups en una sola lectura por hoja (optimización de velocidad)
  var abbrToProvMap = {}, abbrToNameMap = {}, abbrToCostoMap = {}, abbrToPrecioMap = {};
  var hProvOC = SS.getSheetByName('Proveedores');
  if (hProvOC) {
    var pData = hProvOC.getDataRange().getValues();
    var lastProv = '', lastProd = '';
    for (var p = 1; p < pData.length; p++) {
      if (pData[p][2] && String(pData[p][2]).trim()) lastProv = String(pData[p][2]).trim();
      if (pData[p][1] && String(pData[p][1]).trim()) lastProd = String(pData[p][1]).trim();
      var ab = String(pData[p][4]).trim();
      var gu = String(pData[p][3]).trim();
      if (ab) { abbrToProvMap[ab] = lastProv; abbrToNameMap[ab] = lastProd + (gu ? ' — ' + gu : ''); }
    }
  }
  var hProdOC = SS.getSheetByName('Productos');
  if (hProdOC) {
    var prData = hProdOC.getDataRange().getValues();
    for (var q = 1; q < prData.length; q++) {
      var abr = String(prData[q][2]).trim();
      if (!abr) continue;
      abbrToCostoMap[abr] = parseFloat(String(prData[q][9]).replace(/[$.]/g,'').replace(/,/g,'')) || 0;
      abbrToPrecioMap[abr] = parseFloat(String(prData[q][8]).replace(/[$.]/g,'').replace(/,/g,'')) || 0;
    }
  }

  // Precios especiales para Clubes (menores al retail)
  const CLUBES_PRECIOS = {
    'PMu':7000, 'PMa':7000, 'PJyQ':7000, 'PCC':7000, 'PJyM':7800,
    'PPM':11000, 'PPJyQ':11000, 'PPCyQ':11000
  };

  const colToAbbrMap = (canal === 'Clubes')
    ? {24:'PMu', 25:'PMa', 26:'PJyQ', 27:'PCC', 28:'PJyM', 29:'PPM', 30:'PPJyQ', 31:'PPCyQ'}
    : (canal === 'Red') ? RED_COL_TO_ABBR
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
      'Pedido',                                  // U  Estado OC
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
      shOC.getRange(r, 19).setFormula('=R' + r + '/Q' + r); // S = Margen %
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

  // → Entregado: registrar fecha SIEMPRE + descontar stock según origen
  if (nuevo === 'Entregado' && anterior !== 'Entregado') {
    _registrarFechaEntrega(sh, row, colEntrega);
    if (origen === 'Depósito') {
      const hProductos = SS.getSheetByName('Productos');
      if (hProductos) _homeStockFisico(sh, row, hProductos, -1);
    } else if (origen === 'Mixto') {
      // Solo descontar productos con origen Depósito (leer JSON de Origen Detalle)
      const hProductos = SS.getSheetByName('Productos');
      if (hProductos) _homeStockFisicoMixto(sh, row, hProductos, -1);
    }
  }

  // ← Sale de Entregado: limpiar fecha SIEMPRE + devolver stock según origen
  if (anterior === 'Entregado' && nuevo !== 'Entregado') {
    sh.getRange(row, colEntrega, 1, 6).clearContent();
    if (origen === 'Depósito') {
      const hProductos = SS.getSheetByName('Productos');
      if (hProductos) _homeStockFisico(sh, row, hProductos, +1);
    } else if (origen === 'Mixto') {
      const hProductos = SS.getSheetByName('Productos');
      if (hProductos) _homeStockFisicoMixto(sh, row, hProductos, +1);
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
        // Alerta inline si stock queda bajo
        if (signo < 0 && nuevoStock <= 3 && nuevoStock >= 0) {
          SS.toast('⚠️ STOCK BAJO: ' + abbr + ' → quedan ' + nuevoStock, 'Alerta Stock', 8);
        }
        break;
      }
    }
  });
}

// Versión Mixto: solo descuenta productos que tienen origen "D" en el JSON
function _homeStockFisicoMixto(shHome, row, hProductos, signo) {
  // Buscar columna "Origen Detalle"
  var headers = shHome.getRange(1, 1, 1, shHome.getLastColumn()).getValues()[0];
  var colDetalle = -1;
  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h]).trim() === 'Origen Detalle') { colDetalle = h + 1; break; }
  }
  if (colDetalle === -1) return; // No hay detalle, no descontar nada
  var jsonStr = String(shHome.getRange(row, colDetalle).getValue() || '');
  var detalle = {};
  try { detalle = JSON.parse(jsonStr); } catch(e) { return; }

  var cantidades = shHome.getRange(row, 21, 1, 19).getValues()[0];
  var prodData = hProductos.getDataRange().getValues();
  var refPedido = String(shHome.getRange(row, 2).getValue() || '');

  Object.keys(HOME_COL_TO_ABBR).forEach(function(colStr) {
    var colIdx = Number(colStr);
    var abbr = HOME_COL_TO_ABBR[colIdx];
    if (detalle[abbr] !== 'D') return; // Solo descontar los de Depósito
    var qty = Number(cantidades[colIdx - 21]) || 0;
    if (qty === 0) return;

    for (var r = 1; r < prodData.length; r++) {
      if (String(prodData[r][2]).trim() === abbr) {
        var celdaFis = hProductos.getRange(r + 1, 6);
        var fisico = Number(celdaFis.getValue()) || 0;
        var nuevoStock = Math.max(0, fisico + (qty * signo));
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

// ════════════════════════════════════════════════════════════
//  ALERTA STOCK BAJO — ejecutar con trigger cada 6 horas
//  Envía email gratis a maleucongelados@gmail.com
//  COSTO: $0 (usa MailApp de Google, sin WATI)
// ════════════════════════════════════════════════════════════

/** Verifica productos con stock bajo y envía email si hay alguno.
 *  Configurar como trigger time-driven: cada 6 horas. */
function verificarStockBajo() {
  var hProd = SS.getSheetByName('Productos');
  var hConf = SS.getSheetByName('Config');
  if (!hProd) return;

  // Leer umbral desde Config (default: 3)
  var umbral = 3;
  if (hConf) {
    var confData = hConf.getDataRange().getValues();
    for (var r = 0; r < confData.length; r++) {
      if (String(confData[r][0]).indexOf('Stock Cr') >= 0) {
        umbral = Number(confData[r][1]) || 3;
        break;
      }
    }
  }

  var data = hProd.getDataRange().getValues();
  var alertas = [];
  for (var r = 1; r < data.length; r++) {
    var nombre = String(data[r][1]).trim();
    var disponible = Number(data[r][7]) || 0; // col H = Disponible
    var fisico = Number(data[r][5]) || 0;     // col F = Físico
    if (nombre && disponible <= umbral && fisico > 0) {
      // Solo alerta si el producto tiene stock fisico (no es un producto inactivo con 0)
      alertas.push({ nombre: nombre, disponible: disponible, fisico: fisico });
    } else if (nombre && fisico === 0) {
      alertas.push({ nombre: nombre, disponible: 0, fisico: 0 });
    }
  }

  // Filtrar: solo productos que están activos (tuvieron stock esta semana)
  var alertasReales = alertas.filter(function(a) {
    return a.disponible <= umbral;
  });

  if (alertasReales.length === 0) return; // Todo OK, no alertar

  // Construir email
  var lineas = alertasReales.map(function(a) {
    if (a.fisico === 0) return '🔴 ' + a.nombre + ' — SIN STOCK';
    return '🟡 ' + a.nombre + ' — Quedan ' + a.disponible + ' (físico: ' + a.fisico + ')';
  });

  var sinStock = alertasReales.filter(function(a) { return a.fisico === 0; }).length;
  var bajo = alertasReales.length - sinStock;

  var asunto = '⚠️ Maleu — ';
  if (sinStock > 0) asunto += sinStock + ' sin stock';
  if (sinStock > 0 && bajo > 0) asunto += ' + ';
  if (bajo > 0) asunto += bajo + ' con stock bajo';

  var cuerpo = 'Alerta automática de stock — Maleu\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    lineas.join('\n') + '\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Umbral configurado: ' + umbral + ' unidades\n' +
    'Hora: ' + new Date().toLocaleString('es-AR', {timeZone:'America/Argentina/Buenos_Aires'}) + '\n\n' +
    'Abrir Sheets: https://docs.google.com/spreadsheets/d/1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY/';

  MailApp.sendEmail({
    to: 'maleucongelados@gmail.com',
    subject: asunto,
    body: cuerpo
  });
}

/** Configura el trigger de stock bajo (ejecutar UNA vez).
 *  Si da error de permisos: ir al editor de Apps Script,
 *  ejecutar esta función manualmente desde ahí (Run). */
function setupTriggerStockBajo() {
  try {
    // Eliminar triggers existentes de esta función
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'verificarStockBajo') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    // Crear trigger cada 6 horas
    ScriptApp.newTrigger('verificarStockBajo')
      .timeBased()
      .everyHours(6)
      .create();
    SpreadsheetApp.getUi().alert('Trigger de stock bajo configurado: cada 6 horas → email a maleucongelados@gmail.com');
  } catch(e) {
    // Si falla por permisos, dar instrucciones claras
    SpreadsheetApp.getUi().alert(
      'Error al crear el trigger: ' + e.message + '\n\n' +
      'SOLUCIÓN:\n' +
      '1. Abrí el editor de Apps Script (Extensiones → Apps Script)\n' +
      '2. Seleccioná la función "setupTriggerStockBajo" en el dropdown\n' +
      '3. Hacé click en ▶ (Run)\n' +
      '4. Aceptá los permisos que te pida Google\n\n' +
      'Solo hace falta hacerlo una vez.'
    );
  }
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
    .addItem('Recibir mercadería (por proveedor)', 'abrirRecibirMercaderia')
    .addItem('Marcar TODAS las OC como Recibidas', 'marcarOCRecibidas')
    .addSeparator()
    .addItem('Archivar semana (Home + Clubes)', 'archivarSemana')
    .addSeparator()
    .addItem('Actualizar fórmulas Productos', 'setupProductosFormulas')
    .addItem('Reset stock semanal', 'resetStockSemanal')
    .addSeparator()
    .addItem('Ver/Editar Config', 'mostrarConfig')
    .addSeparator()
    .addItem('⚠️ Activar alerta stock bajo (cada 6hs)', 'setupTriggerStockBajo')
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

// ══════════════════════════════════════════════════════════════
//  RECIBIR MERCADERÍA POR PROVEEDOR — Dialog con checkboxes
// ══════════════════════════════════════════════════════════════

/** Devuelve OC en estado "Pedido" agrupadas por proveedor.
 *  Formato: [{ proveedor, items: [{ row, producto, abbr, qty, costo, canal, origen }], totalCosto }] */
function getOCPendientesPorProveedor() {
  var shOC = SS.getSheetByName('Orden de Compra');
  if (!shOC) return [];
  var data = shOC.getDataRange().getValues();
  var porProv = {};

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][20]).trim() !== 'Pedido') continue; // U=21 (0-based=20)
    var prov = String(data[r][9]).trim();  // J = Proveedor
    if (!prov) prov = '(sin proveedor)';
    if (!porProv[prov]) porProv[prov] = { proveedor: prov, items: [], totalCosto: 0 };
    var costoTotal = Number(data[r][14]) || 0; // O = Costo Total
    porProv[prov].items.push({
      row: r + 1,
      producto: String(data[r][10]).trim(),   // K = Producto
      abbr: String(data[r][11]).trim(),       // L = Abreviatura
      qty: Number(data[r][12]) || 0,          // M = Cantidad
      costo: costoTotal,
      canal: String(data[r][4]).trim(),       // E = Canal
      origen: String(data[r][19]).trim()      // T = Origen
    });
    porProv[prov].totalCosto += costoTotal;
  }

  // Ordenar por cantidad de items desc
  return Object.keys(porProv).map(function(k) { return porProv[k]; })
    .sort(function(a, b) { return b.items.length - a.items.length; });
}

/** Marca como "Recibido" todas las OC de los proveedores seleccionados.
 *  @param {string[]} proveedores - nombres de proveedores a marcar */
function marcarRecibidasPorProveedores(proveedores) {
  var shOC = SS.getSheetByName('Orden de Compra');
  var hProd = SS.getSheetByName('Productos');
  if (!shOC || !hProd) return { ok: false, msg: 'No se encontraron las hojas necesarias' };

  var provSet = {};
  proveedores.forEach(function(p) { provSet[p] = true; });

  var data = shOC.getDataRange().getValues();

  // Timestamp Argentina
  var ahora = new Date();
  var argDate = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var dd = String(argDate.getDate()).padStart(2, '0');
  var mm = String(argDate.getMonth() + 1).padStart(2, '0');
  var yyyy = argDate.getFullYear();
  var fechaHoy = dd + '/' + mm + '/' + yyyy;

  var prodData = hProd.getDataRange().getValues();
  var stockSumado = {};
  var marcadas = 0;
  var provsMarcados = {};

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][20]).trim() !== 'Pedido') continue;
    var prov = String(data[r][9]).trim() || '(sin proveedor)';
    if (!provSet[prov]) continue;

    var row = r + 1;
    var canal  = String(data[r][4]).trim();
    var origen = String(data[r][19]).trim();
    var abbr   = String(data[r][11]).trim();
    var qty    = Number(data[r][12]) || 0;

    shOC.getRange(row, 21).setValue('Recibido');  // U = Estado OC
    shOC.getRange(row, 23).setValue(fechaHoy);    // W = Fecha Recibido
    marcadas++;
    provsMarcados[prov] = true;

    if (canal === 'Depósito' && origen === 'Orden de Compra' && abbr && qty > 0) {
      if (!stockSumado[abbr]) stockSumado[abbr] = 0;
      stockSumado[abbr] += qty;
    }
  }

  // Sumar stock en Productos + Kardex
  var prodActualizados = 0;
  Object.keys(stockSumado).forEach(function(abbr) {
    for (var p = 1; p < prodData.length; p++) {
      if (String(prodData[p][2]).trim() === abbr) {
        var celda = hProd.getRange(p + 1, 6);
        var actual = Number(celda.getValue()) || 0;
        var nuevoStock = actual + stockSumado[abbr];
        celda.setValue(nuevoStock);
        _logKardex(abbr, '+REC', stockSumado[abbr], actual, nuevoStock, 'OC', 'Recibido x Prov');
        prodActualizados++;
        break;
      }
    }
  });

  var provsArr = Object.keys(provsMarcados);
  var msg = marcadas + ' órdenes marcadas como Recibido (' + provsArr.join(', ') + ')';
  if (prodActualizados > 0) msg += '\n' + prodActualizados + ' productos actualizados en stock';
  SS.toast(msg, 'Listo', 5);
  return { ok: true, msg: msg, marcadas: marcadas, proveedores: provsArr };
}

/** Abre el dialog para recibir mercadería por proveedor */
function abrirRecibirMercaderia() {
  var html = HtmlService.createHtmlOutput(_getRecibirMercaderiaHTML())
    .setWidth(420)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Recibir mercadería');
}

function _getRecibirMercaderiaHTML() {
  return `<!DOCTYPE html>
<html>
<head>
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { font-family:'Segoe UI',system-ui,sans-serif; font-size:13px; color:#331C1C; padding:20px; }
  h2 { font-size:16px; font-weight:800; margin-bottom:4px; }
  .subtitle { font-size:12px; color:#7A4A4A; margin-bottom:16px; }
  .loading { text-align:center; padding:40px 0; color:#7A4A4A; font-style:italic; }
  .empty { text-align:center; padding:40px 0; color:#7A4A4A; }
  .prov-list { max-height:340px; overflow-y:auto; }
  .prov-card {
    border:1.5px solid #D4C8A0; border-radius:10px; padding:12px;
    margin-bottom:10px; cursor:pointer; transition:all .15s;
  }
  .prov-card:hover { border-color:#F07D47; }
  .prov-card.selected { border-color:#331C1C; background:#F2E8C7; }
  .prov-header { display:flex; align-items:center; gap:10px; }
  .prov-check { width:20px; height:20px; border-radius:5px; border:2px solid #D4C8A0;
    display:flex; align-items:center; justify-content:center; flex-shrink:0;
    transition:all .15s; font-size:14px; color:#fff; }
  .prov-card.selected .prov-check { background:#331C1C; border-color:#331C1C; }
  .prov-name { font-weight:700; font-size:14px; flex:1; }
  .prov-badge { background:#F07D47; color:#fff; font-size:11px; font-weight:700;
    padding:2px 8px; border-radius:50px; }
  .prov-detail { margin-top:8px; font-size:11px; color:#7A4A4A; line-height:1.6; padding-left:30px; }
  .prov-total { font-size:12px; font-weight:700; color:#331C1C; margin-top:4px; padding-left:30px; }
  .select-all { font-size:12px; color:#F07D47; font-weight:700; cursor:pointer;
    text-decoration:underline; margin-bottom:12px; display:inline-block; }
  .btn-confirm {
    display:block; width:100%; margin-top:16px; padding:14px;
    background:#331C1C; color:#F2E8C7; border:none; border-radius:10px;
    font-family:inherit; font-size:14px; font-weight:700; cursor:pointer;
  }
  .btn-confirm:hover { background:#1E1010; }
  .btn-confirm:disabled { background:#ccc; color:#888; cursor:not-allowed; }
  .result { margin-top:12px; padding:12px; border-radius:10px; font-size:12px;
    font-weight:600; display:none; }
  .result.ok { display:block; background:#e8f5e8; color:#2d6a2d; }
  .result.err { display:block; background:#fce8e8; color:#c0392b; }
</style>
</head>
<body>

<h2>Recibir mercadería</h2>
<p class="subtitle">Seleccioná los proveedores de los que recibiste</p>

<div id="content"><div class="loading">Cargando proveedores...</div></div>
<div id="actions" style="display:none">
  <span class="select-all" onclick="toggleAll()">Seleccionar todos</span>
  <button class="btn-confirm" id="btn" onclick="confirmar()" disabled>Confirmar recibido →</button>
</div>
<div class="result" id="result"></div>

<script>
var proveedores = [];
var selected = {};

google.script.run.withSuccessHandler(function(data) {
  proveedores = data;
  render();
}).withFailureHandler(function(err) {
  document.getElementById('content').innerHTML = '<div class="empty">Error: ' + err.message + '</div>';
}).getOCPendientesPorProveedor();

function render() {
  var el = document.getElementById('content');
  if (proveedores.length === 0) {
    el.innerHTML = '<div class="empty">No hay órdenes en estado "Pedido"</div>';
    return;
  }
  document.getElementById('actions').style.display = '';
  var html = '<div class="prov-list">';
  proveedores.forEach(function(p, i) {
    var items = p.items.map(function(it) {
      return it.producto + ' x' + it.qty;
    }).join(', ');
    html += '<div class="prov-card" data-prov="' + p.proveedor + '" onclick="toggle(this,' + i + ')">' +
      '<div class="prov-header">' +
      '<div class="prov-check" id="chk-' + i + '">✓</div>' +
      '<span class="prov-name">' + p.proveedor + '</span>' +
      '<span class="prov-badge">' + p.items.length + (p.items.length === 1 ? ' item' : ' items') + '</span>' +
      '</div>' +
      '<div class="prov-detail">' + items + '</div>' +
      '<div class="prov-total">Costo: $' + p.totalCosto.toLocaleString('es-AR') + '</div>' +
      '</div>';
  });
  html += '</div>';
  el.innerHTML = html;
}

function toggle(card, idx) {
  var prov = proveedores[idx].proveedor;
  if (selected[prov]) { delete selected[prov]; card.classList.remove('selected'); }
  else { selected[prov] = true; card.classList.add('selected'); }
  document.getElementById('btn').disabled = Object.keys(selected).length === 0;
}

function toggleAll() {
  var allSelected = Object.keys(selected).length === proveedores.length;
  var cards = document.querySelectorAll('.prov-card');
  if (allSelected) {
    selected = {};
    cards.forEach(function(c) { c.classList.remove('selected'); });
  } else {
    proveedores.forEach(function(p, i) {
      selected[p.proveedor] = true;
      cards[i].classList.add('selected');
    });
  }
  document.getElementById('btn').disabled = Object.keys(selected).length === 0;
}

function confirmar() {
  var provsArr = Object.keys(selected);
  if (provsArr.length === 0) return;
  var btn = document.getElementById('btn');
  btn.disabled = true;
  btn.textContent = 'Procesando...';
  google.script.run.withSuccessHandler(function(res) {
    var r = document.getElementById('result');
    if (res.ok) {
      r.className = 'result ok';
      r.textContent = res.msg;
      r.style.display = 'block';
      btn.textContent = 'Listo ✓';
      setTimeout(function() { google.script.host.close(); }, 2000);
    } else {
      r.className = 'result err';
      r.textContent = res.msg;
      r.style.display = 'block';
      btn.textContent = 'Confirmar recibido →';
      btn.disabled = false;
    }
  }).withFailureHandler(function(err) {
    var r = document.getElementById('result');
    r.className = 'result err';
    r.textContent = 'Error: ' + err.message;
    r.style.display = 'block';
    btn.textContent = 'Confirmar recibido →';
    btn.disabled = false;
  }).marcarRecibidasPorProveedores(provsArr);
}
</script>
</body>
</html>`;
}

/** Genera Hoja de Ruta: Sección 1 = Búsqueda (por proveedor, desglosado), Sección 2 = Entregas (por cliente, legible) */
function generarHojaDeRuta() {
  var shOC = SS.getSheetByName('Orden de Compra');
  if (!shOC) { SpreadsheetApp.getUi().alert('No se encontró la hoja "Orden de Compra"'); return; }

  var data = shOC.getDataRange().getValues();
  if (data.length <= 1) { SpreadsheetApp.getUi().alert('No hay órdenes de compra'); return; }

  // Separar "Producto — Gusto" en categoría y variedad
  function splitProducto(nombre) {
    var parts = nombre.split(' — ');
    return { cat: (parts[0] || nombre).trim(), variedad: (parts[1] || '').trim() };
  }

  // ── Recopilar datos ──
  var porProv = {};
  var porCliente = {};
  var totalGeneral = 0;

  for (var r = 1; r < data.length; r++) {
    var estado = String(data[r][20]).trim();
    if (estado !== 'Pendiente' && estado !== 'Pedido') continue;
    var origen = String(data[r][19]).trim();
    if (origen === 'Depósito') continue;

    var prov     = String(data[r][9]).trim();
    var producto = String(data[r][10]).trim();
    var qty      = Number(data[r][12]) || 0;
    var costoU   = Number(String(data[r][13]).replace(/[$.]/g,'').replace(/,/g,'')) || 0;
    var costoT   = costoU * qty;
    var cliente  = String(data[r][6]).trim();
    var canal    = String(data[r][4]).trim();
    var dir      = String(data[r][8]).trim();

    if (!prov || qty === 0) continue;

    // Proveedor → categoría → variedades
    if (!porProv[prov]) porProv[prov] = { totalCosto: 0, categorias: {} };
    var sp = splitProducto(producto);
    if (!porProv[prov].categorias[sp.cat]) porProv[prov].categorias[sp.cat] = [];
    // Sumar si ya existe la variedad
    var found = false;
    for (var v = 0; v < porProv[prov].categorias[sp.cat].length; v++) {
      if (porProv[prov].categorias[sp.cat][v].variedad === sp.variedad) {
        porProv[prov].categorias[sp.cat][v].qty += qty;
        found = true; break;
      }
    }
    if (!found) porProv[prov].categorias[sp.cat].push({ variedad: sp.variedad, qty: qty });
    porProv[prov].totalCosto += costoT;
    totalGeneral += costoT;

    // Cliente → categoría → variedades (sumar si ya existe)
    if (!porCliente[cliente]) porCliente[cliente] = { canal: canal, dir: dir, categorias: {} };
    if (!porCliente[cliente].categorias[sp.cat]) porCliente[cliente].categorias[sp.cat] = [];
    var foundCl = false;
    for (var vc = 0; vc < porCliente[cliente].categorias[sp.cat].length; vc++) {
      if (porCliente[cliente].categorias[sp.cat][vc].variedad === sp.variedad) {
        porCliente[cliente].categorias[sp.cat][vc].qty += qty;
        foundCl = true; break;
      }
    }
    if (!foundCl) porCliente[cliente].categorias[sp.cat].push({ variedad: sp.variedad, qty: qty });
  }

  var provs = Object.keys(porProv);
  var clientes = Object.keys(porCliente);
  if (provs.length === 0) { SpreadsheetApp.getUi().alert('No hay órdenes pendientes para buscar'); return; }

  // ── Fecha de ruta (próximo viernes o día actual) ──
  var ahora = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  // Si es lunes-jueves, la ruta es el viernes. Si es viernes+, es hoy.
  var rutaDate = new Date(argNow);
  var dow = rutaDate.getDay(); // 0=dom, 5=vie
  if (dow >= 1 && dow <= 4) { rutaDate.setDate(rutaDate.getDate() + (5 - dow)); }
  var DIAS_SEMANA = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
  var MESES_LARGO = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var fechaRuta = DIAS_SEMANA[rutaDate.getDay()] + ' ' + rutaDate.getDate() + ' de ' + MESES_LARGO[rutaDate.getMonth()] + ' del ' + rutaDate.getFullYear();

  // ── Construir HTML ──
  var html = '<title>' + fechaRuta + '</title><style>' +
    '@page{margin:0mm;}' +
    'body{font-family:system-ui,sans-serif;font-size:13px;color:#331C1C;padding:12mm 15mm;max-width:750px;margin:0 auto;}' +
    'h2{font-size:20px;font-weight:800;margin-bottom:2px;}' +
    '.sub{font-size:12px;color:#7A4A4A;margin-bottom:20px;}' +
    '.section-title{font-size:14px;font-weight:800;text-transform:uppercase;letter-spacing:.08em;color:#F07D47;margin:24px 0 12px;padding-bottom:6px;border-bottom:2px solid #F07D47;}' +
    '.prov{background:#f8f6f0;border-radius:12px;padding:14px 16px;margin-bottom:12px;}' +
    '.prov-t{font-size:15px;font-weight:800;display:flex;justify-content:space-between;margin-bottom:10px;}' +
    '.prov-cost{color:#F07D47;}' +
    '.cat-label{font-size:12px;font-weight:800;color:#F07D47;text-transform:uppercase;letter-spacing:.04em;margin:10px 0 4px;padding-top:6px;border-top:1px solid #e8dfc4;}' +
    '.cat-label:first-child{border-top:none;margin-top:0;padding-top:0;}' +
    '.var-row{display:flex;align-items:center;gap:6px;padding:2px 0;font-size:13px;}' +
    '.var-qty{font-weight:800;color:#331C1C;min-width:28px;}' +
    '.var-name{font-weight:500;}' +
    '.cliente{background:#fff;border:1.5px solid #e8dfc4;border-radius:12px;padding:14px 16px;margin-bottom:12px;page-break-inside:avoid;}' +
    '.cliente-head{display:flex;align-items:center;gap:10px;margin-bottom:6px;}' +
    '.cb{width:18px;height:18px;border:2px solid #331C1C;border-radius:4px;flex-shrink:0;}' +
    '.cliente-name{font-size:14px;font-weight:800;flex:1;}' +
    '.cliente-canal{font-size:10px;font-weight:700;background:#F07D47;color:#fff;padding:2px 8px;border-radius:50px;}' +
    '.cliente-dir{font-size:11px;color:#7A4A4A;margin:0 0 8px 28px;}' +
    '.cl-cat{font-size:11px;font-weight:800;color:#7A4A4A;margin:8px 0 3px 28px;text-transform:uppercase;letter-spacing:.04em;}' +
    '.cl-item{display:flex;align-items:center;gap:8px;padding:2px 0 2px 28px;font-size:13px;}' +
    '.cl-item .cb{width:15px;height:15px;border-width:1.5px;}' +
    '.cl-qty{font-weight:800;min-width:22px;color:#F07D47;}' +
    '.cl-name{font-weight:500;}' +
    '.tb{background:#331C1C;color:#F2E8C7;border-radius:12px;padding:14px 16px;display:flex;justify-content:space-between;font-size:15px;font-weight:800;margin-top:20px;}' +
    '.tb span:last-child{color:#F07D47;}' +
    '.pb{display:block;width:100%;margin-top:12px;padding:12px;background:#F07D47;color:#fff;border:none;border-radius:10px;font-size:14px;font-weight:700;cursor:pointer;}' +
    '.copy-btn{display:block;width:100%;margin-top:10px;padding:8px;background:#25D366;color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer;}' +
    '@media print{.pb,.copy-btn{display:none!important;} body{padding:10px;font-size:12px;} .cliente{border:1px solid #ccc;} .cb{border:1.5px solid #000;}}' +
    '</style>';

  html += '<h2 style="text-align:center">' + fechaRuta + '</h2>';
  html += '<p class="sub" style="text-align:center">Hoja de Ruta — Maleu</p>';

  // ── SECCIÓN 1: BÚSQUEDA POR PROVEEDOR (desglosado por categoría) ──
  html += '<div class="section-title">1. Búsqueda por proveedor</div>';

  provs.forEach(function(prov, idx) {
    var d = porProv[prov];

    // Texto WA profesional agrupado por categoría
    var copyLines = ['Hola! Te paso el pedido de esta semana:', ''];
    var cats = Object.keys(d.categorias);
    cats.forEach(function(cat) {
      copyLines.push(cat + ':');
      d.categorias[cat].forEach(function(v) {
        copyLines.push('  ' + v.qty + ' × ' + (v.variedad || cat));
      });
      copyLines.push('');
    });
    copyLines.push('Total: $' + d.totalCosto.toLocaleString('es-AR'));
    copyLines.push('');
    copyLines.push('Gracias!');
    var copyText = copyLines.join('\\n');

    // HTML visual
    html += '<div class="prov"><div class="prov-t"><span>' + prov + '</span><span class="prov-cost">$' + d.totalCosto.toLocaleString('es-AR') + '</span></div>';
    cats.forEach(function(cat, ci) {
      html += '<div class="cat-label"' + (ci===0?' style="border-top:none;margin-top:0;padding-top:0"':'') + '>' + cat + '</div>';
      d.categorias[cat].forEach(function(v) {
        html += '<div class="var-row"><span class="var-qty">' + v.qty + ' ×</span><span class="var-name">' + (v.variedad || cat) + '</span></div>';
      });
    });
    html += '<button class="copy-btn" onclick="copyProv(' + idx + ',this)">Copiar para WhatsApp</button>';
    html += '<textarea id="prov-txt-' + idx + '" style="position:absolute;left:-9999px">' + copyText + '</textarea>';
    html += '</div>';
  });

  html += '<div class="tb"><span>Total a pagar proveedores</span><span>$' + totalGeneral.toLocaleString('es-AR') + '</span></div>';

  // ── SECCIÓN 2: ENTREGAS POR CLIENTE (legible, con checkbox por cliente e item) ──
  html += '<div class="section-title">2. Entregas por cliente (' + clientes.length + ')</div>';

  clientes.forEach(function(nombre) {
    var c = porCliente[nombre];
    html += '<div class="cliente">';
    // Checkbox del cliente + nombre + canal
    html += '<div class="cliente-head"><div class="cb"></div><span class="cliente-name">' + nombre + '</span><span class="cliente-canal">' + c.canal + '</span></div>';
    if (c.dir) html += '<div class="cliente-dir">' + c.dir + '</div>';
    // Items agrupados por categoría
    var cats = Object.keys(c.categorias);
    cats.forEach(function(cat) {
      html += '<div class="cl-cat">' + cat + '</div>';
      c.categorias[cat].forEach(function(item) {
        html += '<div class="cl-item"><div class="cb"></div><span class="cl-qty">' + item.qty + ' ×</span><span class="cl-name">' + (item.variedad || cat) + '</span></div>';
      });
    });
    html += '</div>';
  });

  html += '<button class="pb" onclick="window.print()">Imprimir Hoja de Ruta</button>';
  html += '<script>function copyProv(idx,btn){var el=document.getElementById("prov-txt-"+idx);if(!el)return;var txt=el.value.replace(/\\\\n/g,"\\n");navigator.clipboard.writeText(txt).then(function(){btn.textContent="Copiado!";btn.style.background="#331C1C";setTimeout(function(){btn.textContent="Copiar para WhatsApp";btn.style.background="#25D366";},2000);});}</script>';

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

/** Carga TODO de una sola vez: proveedores + productos + costos + stock.
 *  Así la sidebar no necesita hacer requests por cada cambio de proveedor. */
function getAllSidebarData() {
  var hProv = SS.getSheetByName('Proveedores');
  var hProd = SS.getSheetByName('Productos');
  if (!hProv) return { proveedores: [], productosPorProv: {} };

  // Costos y stock desde Productos (1 sola lectura)
  var costoMap = {}, stockMap = {};
  if (hProd) {
    var prodData = hProd.getDataRange().getValues();
    for (var r = 1; r < prodData.length; r++) {
      var ab = String(prodData[r][2]).trim();
      if (!ab) continue;
      costoMap[ab] = parseFloat(String(prodData[r][9]).replace(/[$.]/g,'').replace(/,/g,'')) || 0;
      stockMap[ab] = Number(prodData[r][5]) || 0;
    }
  }

  // Proveedores y productos (1 sola lectura)
  var provData = hProv.getDataRange().getValues();
  var proveedores = [], seenProv = {}, productosPorProv = {};
  var lastProv = '', lastProd = '';
  for (var r2 = 1; r2 < provData.length; r2++) {
    if (provData[r2][2] && String(provData[r2][2]).trim()) lastProv = String(provData[r2][2]).trim();
    if (provData[r2][1] && String(provData[r2][1]).trim()) lastProd = String(provData[r2][1]).trim();
    if (!lastProv) continue;
    if (!seenProv[lastProv]) { seenProv[lastProv] = true; proveedores.push(lastProv); }

    var abbr = String(provData[r2][4]).trim();
    var gusto = String(provData[r2][3]).trim();
    if (!abbr) continue;

    if (!productosPorProv[lastProv]) productosPorProv[lastProv] = [];
    productosPorProv[lastProv].push({
      abbr: abbr,
      nombre: lastProd + ' — ' + gusto,
      costo: costoMap[abbr] || 0,
      stock: stockMap[abbr] !== undefined ? stockMap[abbr] : null
    });
  }

  return { proveedores: proveedores, productosPorProv: productosPorProv };
}

/** Legacy: mantener compatibilidad con sidebar que llama getProductosPorProveedor */
function getProductosPorProveedor(proveedor) {
  var all = getAllSidebarData();
  return all.productosPorProv[proveedor] || [];
}

/** Genera filas en Orden de Compra para compra de Depósito / Red (sidebar) */
function confirmarCompraDeposito(proveedor, items, fechaBusqueda, vendedor) {
  const shOC = SS.getSheetByName('Orden de Compra');
  if (!shOC) throw new Error('Hoja "Orden de Compra" no encontrada');

  var esMarcos = vendedor === 'Marcos Bottcher';

  // Validar que haya items con cantidad > 0
  const itemsValidos = items.filter(function(item) { return item.qty > 0; });
  if (itemsValidos.length === 0) throw new Error('No hay productos con cantidad > 0');

  // Cargar precios de venta desde Productos (para calcular margen)
  var precioVentaMap = {};
  var hProd = SS.getSheetByName('Productos');
  if (hProd) {
    var prData = hProd.getDataRange().getValues();
    for (var p = 1; p < prData.length; p++) {
      var ab = String(prData[p][2]).trim();
      if (ab) precioVentaMap[ab] = parseFloat(String(prData[p][8]).replace(/[$.]/g,'').replace(/,/g,'')) || 0;
    }
  }

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
      precioVentaMap[item.abbr] || 0,             // P  Precio Venta Unit. (retail)
      0,                                          // Q  Ingreso Total (fórmula)
      0,                                          // R  Margen Bruto $ (fórmula)
      0,                                          // S  Margen % (fórmula)
      item.origen || 'Orden de Compra',           // T  Origen
      'Pedido',                                   // U  Estado OC (igual que Home)
      dd + '/' + mm + '/' + yyyy,                 // V  Fecha Pedido Prov (hoy)
      '',                                         // W  Fecha Recibido
      'No',                                       // X  Pagado Proveedor
      'No',                                       // Y  Cobrado Cliente
    ]);
  });

  // Escribir filas
  const startRow = shOC.getLastRow() + 1;
  shOC.getRange(startRow, 1, newRows.length, 25).setValues(newRows);

  // Fórmulas financieras (setFormula individual, compatible con locale español)
  for (let i = 0; i < newRows.length; i++) {
    const r = startRow + i;
    shOC.getRange(r, 17).setFormula('=P' + r + '*M' + r);              // Q = Ingreso
    shOC.getRange(r, 18).setFormula('=Q' + r + '-O' + r);              // R = Margen $
    shOC.getRange(r, 19).setFormula('=R' + r + '/Q' + r); // S = Margen %
  }

  // Formato — batch en lugar de individual (más rápido)
  var fmtRange = shOC.getRange(startRow, 13, newRows.length, 7); // M a S
  shOC.getRange(startRow, 13, newRows.length, 1).setNumberFormat('0');       // M
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
let _allData = null; // Cache de todos los datos

// Cargar TODO de una vez al abrir (1 sola request = más rápido)
google.script.run.withSuccessHandler(function(data) {
  _allData = data;
  const sel = document.getElementById('sel-prov');
  data.proveedores.forEach(function(p) {
    const opt = document.createElement('option');
    opt.value = p; opt.textContent = p;
    sel.appendChild(opt);
  });
}).getAllSidebarData();

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
  // Usar datos cacheados (instantáneo, sin request al server)
  var items = (_allData && _allData.productosPorProv[proveedorActual]) || [];
  (function(items) {
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
  })(items);
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

// ══════════════════════════════════════════════════════════════
//  ADMIN — Dashboard resumen completo de Maleu
// ══════════════════════════════════════════════════════════════

/** GET ?action=admin
 *  Devuelve resumen de pedidos (semana actual) + stock + OC pendientes.
 *  Datos compactos para la PWA admin. */
function _doGetAdmin() {
  var ahora = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));

  // ── Leer pedidos de las 4 hojas operativas ──
  var ABBRS_HOME = ['PPM','PPJyQ','PPCyQ','SCo','SJyQ','SCa','ECaC','EJyQ','ECyQ','EV',
                    'TG','TLC','TC','F','PMu','PMa','PJyQ','PCC','PJyM'];
  var ABBRS_CLUB = ['PMu','PMa','PJyQ','PCC','PJyM','PPM','PPCyQ','PPJyQ'];

  var canales = [];
  var pedidos = [];

  // Home, Pilar, Capital Federal (misma estructura: 53 cols)
  ['Home', 'Pilar', 'Capital Federal'].forEach(function(hoja) {
    var sh = SS.getSheetByName(hoja);
    var stats = { nombre: hoja, pedidos: 0, entregados: 0, pendientes: 0, cancelados: 0, reservados: 0, facturado: 0, cobrado: 0, noCobrado: 0 };
    if (!sh || sh.getLastRow() <= 1) { canales.push(stats); return; }
    var data = sh.getDataRange().getValues();

    for (var r = 1; r < data.length; r++) {
      var estado = String(data[r][10]).trim();
      var estadoPago = String(data[r][12]).trim();
      var total = Number(data[r][13]) || 0;
      var facturado = Number(data[r][19]) || total;
      var origen = String(data[r][8]).trim();
      var cliente = String(data[r][7]).trim();
      var nPedido = String(data[r][1]).trim();
      var diaEntrega = String(data[r][9]).trim();
      var fecha = data[r][3];
      var fechaStr = '';
      if (fecha instanceof Date) {
        fechaStr = Utilities.formatDate(fecha, 'America/Argentina/Buenos_Aires', 'dd/MM');
      } else {
        fechaStr = String(fecha || '');
      }

      stats.pedidos++;
      if (estado === 'Entregado') stats.entregados++;
      else if (estado === 'Cancelado') stats.cancelados++;
      else if (estado === 'Reservado') stats.reservados++;
      else stats.pendientes++;

      stats.facturado += facturado;
      if (estadoPago === 'Cobrado') stats.cobrado += facturado;
      else stats.noCobrado += facturado;

      // Productos del pedido
      var prods = [];
      for (var p = 0; p < 19; p++) {
        var qty = Number(data[r][20 + p]) || 0;
        if (qty > 0) prods.push({ a: ABBRS_HOME[p], q: qty });
      }

      var formaPago = String(data[r][11] || '').trim();
      var costoPed = Number(data[r][39]) || 0;
      var margenPed = Number(data[r][40]) || 0;
      var subBarrio = hoja === 'Home' ? String(data[r][42] || '').trim() : '';

      pedidos.push({
        n: nPedido, h: hoja, c: cliente, f: fechaStr, de: diaEntrega,
        es: estado, o: origen, ep: estadoPago, fp: formaPago,
        $: facturado, co: costoPed, mg: margenPed, br: subBarrio, p: prods
      });
    }
    canales.push(stats);
  });

  // Clubes (34 cols, estructura diferente)
  var shClubes = SS.getSheetByName('Clubes');
  var statsClubes = { nombre: 'Clubes', pedidos: 0, entregados: 0, pendientes: 0, cancelados: 0, reservados: 0, facturado: 0, cobrado: 0, noCobrado: 0 };
  if (shClubes && shClubes.getLastRow() > 1) {
    var dataClubes = shClubes.getDataRange().getValues();
    for (var rc = 1; rc < dataClubes.length; rc++) {
      var estadoC = String(dataClubes[rc][13]).trim();
      var estadoPagoC = String(dataClubes[rc][15]).trim();
      var totalC = Number(dataClubes[rc][16]) || 0;
      var facturadoC = Number(dataClubes[rc][22]) || totalC;
      var origenC = String(dataClubes[rc][11]).trim();
      var clienteC = String(dataClubes[rc][7]).trim();
      var clubC = String(dataClubes[rc][8]).trim();
      var nPedidoC = String(dataClubes[rc][1]).trim();
      var diaEntregaC = String(dataClubes[rc][12]).trim();
      var fechaC = dataClubes[rc][3];
      var fechaStrC = '';
      if (fechaC instanceof Date) {
        fechaStrC = Utilities.formatDate(fechaC, 'America/Argentina/Buenos_Aires', 'dd/MM');
      } else {
        fechaStrC = String(fechaC || '');
      }

      statsClubes.pedidos++;
      if (estadoC === 'Entregado') statsClubes.entregados++;
      else if (estadoC === 'Cancelado') statsClubes.cancelados++;
      else if (estadoC === 'Reservado') statsClubes.reservados++;
      else statsClubes.pendientes++;

      statsClubes.facturado += facturadoC;
      if (estadoPagoC === 'Cobrado') statsClubes.cobrado += facturadoC;
      else statsClubes.noCobrado += facturadoC;

      // Productos Clubes
      var prodsC = [];
      for (var pc = 0; pc < 8; pc++) {
        var qtyC = Number(dataClubes[rc][23 + pc]) || 0;
        if (qtyC > 0) prodsC.push({ a: ABBRS_CLUB[pc], q: qtyC });
      }

      var formaPagoC = String(dataClubes[rc][14] || '').trim();

      var costoCl = Number(dataClubes[rc][31]) || 0;
      var margenCl = Number(dataClubes[rc][32]) || 0;

      pedidos.push({
        n: nPedidoC, h: 'Clubes',
        c: clienteC + (clubC ? ' (' + clubC + ')' : ''),
        f: fechaStrC, de: diaEntregaC, es: estadoC, o: origenC,
        ep: estadoPagoC, fp: formaPagoC, $: facturadoC,
        co: costoCl, mg: margenCl, br: clubC, p: prodsC
      });
    }
  }
  canales.push(statsClubes);

  // Red (55 cols, estructura propia)
  var ABBRS_RED_DASH = ['PPM','PPJyQ','PPCyQ','SCo','SJyQ','SCa','ECaC','EJyQ','ECyQ','EV',
                        'TG','TLC','TC','F','SQB','SL','SPyP','SE','PMu','PMa','PJyQ','PCC','PJyM'];
  var shRedDash = SS.getSheetByName('Red');
  var statsRed = { nombre: 'Red', pedidos: 0, entregados: 0, pendientes: 0, cancelados: 0, reservados: 0, facturado: 0, cobrado: 0, noCobrado: 0 };
  if (shRedDash && shRedDash.getLastRow() > 1) {
    var dataRed = shRedDash.getDataRange().getValues();
    for (var rr = 1; rr < dataRed.length; rr++) {
      var estadoR = String(dataRed[rr][11]).trim(); // L = Estado Entrega
      var estadoPagoR = String(dataRed[rr][13]).trim(); // N = Estado Pago
      var totalR = Number(dataRed[rr][14]) || 0; // O = Total
      var facturadoR = Number(dataRed[rr][20]) || totalR; // U = Facturado
      var origenR = String(dataRed[rr][9]).trim(); // J = Origen
      var clienteR = String(dataRed[rr][8]).trim(); // I = Cliente
      var vendedorR = String(dataRed[rr][7]).trim(); // H = Vendedor
      var nPedidoR = String(dataRed[rr][1]).trim();
      var diaEntregaR = String(dataRed[rr][10]).trim(); // K
      var fechaR = dataRed[rr][3];
      var fechaStrR = fechaR instanceof Date ? Utilities.formatDate(fechaR, 'America/Argentina/Buenos_Aires', 'dd/MM') : String(fechaR || '');

      statsRed.pedidos++;
      if (estadoR === 'Entregado') statsRed.entregados++;
      else if (estadoR === 'Cancelado') statsRed.cancelados++;
      else if (estadoR === 'Reservado') statsRed.reservados++;
      else statsRed.pendientes++;

      statsRed.facturado += facturadoR;
      if (estadoPagoR === 'Cobrado') statsRed.cobrado += facturadoR;
      else statsRed.noCobrado += facturadoR;

      var prodsR = [];
      for (var prr = 0; prr < 23; prr++) {
        var qtyR = Number(dataRed[rr][21 + prr]) || 0;
        if (qtyR > 0) prodsR.push({ a: ABBRS_RED_DASH[prr], q: qtyR });
      }

      var formaPagoR = String(dataRed[rr][12] || '').trim();
      var costoR = Number(dataRed[rr][44]) || 0;
      var margenR = Number(dataRed[rr][45]) || 0;

      pedidos.push({
        n: nPedidoR, h: 'Red',
        c: clienteR + (vendedorR ? ' (Red: ' + vendedorR + ')' : ''),
        f: fechaStrR, de: diaEntregaR, es: estadoR, o: origenR,
        ep: estadoPagoR, fp: formaPagoR, $: facturadoR,
        co: costoR, mg: margenR, br: vendedorR, p: prodsR
      });
    }
  }
  canales.push(statsRed);

  // ── Totales ──
  var totales = { pedidos: 0, entregados: 0, pendientes: 0, cancelados: 0, reservados: 0, facturado: 0, cobrado: 0, noCobrado: 0, ticket: 0 };
  canales.forEach(function(c) {
    totales.pedidos += c.pedidos;
    totales.entregados += c.entregados;
    totales.pendientes += c.pendientes;
    totales.cancelados += c.cancelados;
    totales.reservados += c.reservados;
    totales.facturado += c.facturado;
    totales.cobrado += c.cobrado;
    totales.noCobrado += c.noCobrado;
  });
  if (totales.pedidos > 0) totales.ticket = Math.round(totales.facturado / totales.pedidos);

  // ── Stock (hoja Productos) ──
  var stock = [];
  var shProd = SS.getSheetByName('Productos');
  if (shProd && shProd.getLastRow() > 1) {
    var dataProd = shProd.getDataRange().getValues();
    // Buscar columnas por header
    var hdr = dataProd[0];
    var colNombre = -1, colAbbr = -1, colFisico = -1, colReservado = -1, colDisp = -1;
    for (var h = 0; h < hdr.length; h++) {
      var hv = String(hdr[h]).trim().toLowerCase();
      if (hv === 'producto' || hv === 'nombre') colNombre = h;
      if (hv === 'abreviatura') colAbbr = h;
      if (hv === 'stock físico' || hv === 'stock fisico') colFisico = h;
      if (hv === 'reservado') colReservado = h;
      if (hv === 'disponible') colDisp = h;
    }
    for (var rp = 1; rp < dataProd.length; rp++) {
      var nombre = colNombre >= 0 ? String(dataProd[rp][colNombre]).trim() : '';
      var abbr = colAbbr >= 0 ? String(dataProd[rp][colAbbr]).trim() : '';
      if (!nombre && !abbr) continue;
      var fisico = colFisico >= 0 ? (Number(dataProd[rp][colFisico]) || 0) : 0;
      var reservado = colReservado >= 0 ? (Number(dataProd[rp][colReservado]) || 0) : 0;
      var disponible = colDisp >= 0 ? (Number(dataProd[rp][colDisp]) || 0) : fisico - reservado;
      stock.push({ n: nombre, a: abbr, f: fisico, r: reservado, d: disponible });
    }
  }

  // ── OC pendientes ──
  var ocPend = 0;
  var ocTotal = 0;
  var shOC = SS.getSheetByName('Orden de Compra');
  if (shOC && shOC.getLastRow() > 1) {
    var dataOC = shOC.getDataRange().getValues();
    for (var ro = 1; ro < dataOC.length; ro++) {
      var estOC = String(dataOC[ro][20]).trim();
      if (estOC === 'Pendiente' || estOC === 'Pedido') {
        ocPend++;
        ocTotal += Number(dataOC[ro][14]) || 0;
      }
    }
  }

  // ── Egresos (gastos) ──
  var gastos = [];
  var totalGastos = 0;
  var gastosEf = 0;
  var gastosMP = 0;
  var shEgr = SS.getSheetByName('Egresos');
  if (shEgr && shEgr.getLastRow() > 1) {
    var dataEgr = shEgr.getDataRange().getValues();
    var hdrE = dataEgr[0];
    for (var re = 1; re < dataEgr.length; re++) {
      if (!dataEgr[re][0]) continue;
      var fechaE = dataEgr[re][0];
      var fechaStrE = '';
      if (fechaE instanceof Date) {
        fechaStrE = Utilities.formatDate(fechaE, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
      } else {
        fechaStrE = String(fechaE || '');
      }
      var montoE = Number(dataEgr[re][6]) || 0;
      var metodoE = String(dataEgr[re][5] || '').trim();
      totalGastos += montoE;
      if (metodoE === 'Efectivo') gastosEf += montoE;
      else gastosMP += montoE;

      gastos.push({
        f: fechaStrE,
        sem: String(dataEgr[re][1] || ''),
        mes: String(dataEgr[re][2] || ''),
        cat: String(dataEgr[re][3] || '').trim(),
        con: String(dataEgr[re][4] || '').trim(),
        met: metodoE,
        $: montoE,
        not: String(dataEgr[re][7] || '').trim()
      });
    }
  }

  // ── Cobrado por método de pago ──
  var cobradoEf = 0;
  var cobradoMP = 0;
  pedidos.forEach(function(p) {
    if (p.ep === 'Cobrado') {
      if (p.fp === 'Efectivo') cobradoEf += p.$;
      else cobradoMP += p.$;
    }
  });

  // ── Saldo Base (último ajuste manual) ──
  var saldoBase = { ef: 0, mp: 0, fecha: '' };
  var shSaldo = SS.getSheetByName('Saldo Base');
  if (shSaldo && shSaldo.getLastRow() > 1) {
    var lastRow = shSaldo.getLastRow();
    saldoBase.ef = Number(shSaldo.getRange(lastRow, 2).getValue()) || 0;
    saldoBase.mp = Number(shSaldo.getRange(lastRow, 3).getValue()) || 0;
    var fSaldo = shSaldo.getRange(lastRow, 1).getValue();
    saldoBase.fecha = fSaldo instanceof Date ? Utilities.formatDate(fSaldo, 'America/Argentina/Buenos_Aires', 'dd/MM HH:mm') : '';
  }

  // ── Ingresos manuales ──
  var ingresos = [];
  var ingresosEf = 0;
  var ingresosMP = 0;
  var shIng = SS.getSheetByName('Ingresos');
  if (shIng && shIng.getLastRow() > 1) {
    var dataIng = shIng.getDataRange().getValues();
    for (var ri = 1; ri < dataIng.length; ri++) {
      if (!dataIng[ri][0]) continue;
      var fechaI = dataIng[ri][0];
      var fechaStrI = fechaI instanceof Date ? Utilities.formatDate(fechaI, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : String(fechaI || '');
      var montoI = Number(dataIng[ri][6]) || 0;
      var metodoI = String(dataIng[ri][5] || '').trim();
      if (metodoI === 'Efectivo') ingresosEf += montoI;
      else ingresosMP += montoI;
      ingresos.push({
        f: fechaStrI, cat: String(dataIng[ri][3] || '').trim(),
        con: String(dataIng[ri][4] || '').trim(), met: metodoI,
        $: montoI, not: String(dataIng[ri][7] || '').trim()
      });
    }
  }

  // ── Gastos Historicos (por mes, para P&L) ──
  var gastosHist = {};
  var shGH = SS.getSheetByName('Gastos Historicos');
  if (shGH && shGH.getLastRow() > 1) {
    var dGH = shGH.getDataRange().getValues();
    for (var rg = 1; rg < dGH.length; rg++) {
      if (!dGH[rg][0]) continue;
      var mesGH = String(dGH[rg][1] || '').trim();
      if (!mesGH) continue;
      var montoGH = Number(dGH[rg][6]) || 0;
      if (!gastosHist[mesGH]) gastosHist[mesGH] = 0;
      gastosHist[mesGH] += montoGH;
    }
  }
  // Also add current Egresos to gastosHist
  gastos.forEach(function(g) {
    var m = g.mes || '';
    if (m) { if (!gastosHist[m]) gastosHist[m] = 0; gastosHist[m] += g.$; }
  });

  return ContentService
    .createTextOutput(JSON.stringify({
      ts: Date.now(),
      canales: canales,
      totales: totales,
      pedidos: pedidos,
      stock: stock,
      oc: { pendientes: ocPend, costo: ocTotal },
      caja: { cobradoEf: cobradoEf, cobradoMP: cobradoMP, gastosEf: gastosEf, gastosMP: gastosMP, totalGastos: totalGastos, ingresosEf: ingresosEf, ingresosMP: ingresosMP },
      saldoBase: saldoBase,
      gastos: gastos,
      ingresos: ingresos,
      gastosHist: gastosHist
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** POST action=gasto — guarda un gasto en la hoja Egresos.
 *  Recibe { action:'gasto', fecha, categoria, concepto, metodo, monto, notas } */
function _doPostGasto(data) {
  var shEgr = SS.getSheetByName('Egresos');
  if (!shEgr) {
    shEgr = SS.insertSheet('Egresos');
    shEgr.getRange(1, 1, 1, 8).setValues([['Fecha', 'Semana', 'Mes', 'Categoría', 'Concepto', 'Método', 'Monto', 'Notas']]);
    shEgr.setFrozenRows(1);
  }

  var ahora = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));

  // Parse fecha or use today
  var fecha = data.fecha ? new Date(data.fecha + 'T12:00:00') : argNow;
  var semana = _getWeekNumber(fecha);
  var meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var mes = meses[fecha.getMonth()];

  var cat = String(data.categoria || '').trim();
  var con = String(data.concepto || '').trim();
  var met = String(data.metodo || '').trim();
  var monto = Number(data.monto) || 0;
  var notas = String(data.notas || '').trim();

  shEgr.appendRow([fecha, semana, mes, cat, con, met, monto, notas]);
  shEgr.getRange(shEgr.getLastRow(), 1).setNumberFormat('dd/MM/yyyy');

  // Copiar a Movimientos Historicos
  _appendMovHistorico(fecha, mes, semana, cat, con, met, monto, notas);

  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

/** Agrega una fila a "Movimientos Historicos" (headers en fila 2, data desde fila 3). */
function _appendMovHistorico(fecha, mes, semana, cat, concepto, metodo, monto, notas) {
  var sh = SS.getSheetByName('Movimientos Historicos');
  if (!sh) return;
  sh.appendRow([fecha, mes, semana, cat, concepto, metodo, monto, notas]);
  sh.getRange(sh.getLastRow(), 1).setNumberFormat('dd/MM/yyyy');
  sh.getRange(sh.getLastRow(), 7).setNumberFormat('$#.##0,00');
}

function _getWeekNumber(d) {
  var oneJan = new Date(d.getFullYear(), 0, 1);
  var days = Math.floor((d - oneJan) / 86400000);
  return Math.ceil((days + oneJan.getDay() + 1) / 7);
}

/** POST action=ingreso — guarda un ingreso manual en hoja Ingresos.
 *  { action:'ingreso', fecha, categoria, concepto, metodo, monto, notas } */
function _doPostIngreso(data) {
  var sh = SS.getSheetByName('Ingresos');
  if (!sh) {
    sh = SS.insertSheet('Ingresos');
    sh.getRange(1, 1, 1, 8).setValues([['Fecha', 'Semana', 'Mes', 'Categoría', 'Concepto', 'Método', 'Monto', 'Notas']]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 8).setBackground(BROWN).setFontColor('#FFFFFF').setFontWeight('bold');
  }
  var ahora = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var fecha = data.fecha ? new Date(data.fecha + 'T12:00:00') : argNow;
  var semana = _getWeekNumber(fecha);
  var meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var mes = meses[fecha.getMonth()];
  var cat = String(data.categoria||'').trim();
  var con = String(data.concepto||'').trim();
  var met = String(data.metodo||'').trim();
  var monto = Number(data.monto)||0;
  var notas = String(data.notas||'').trim();

  sh.appendRow([fecha, semana, mes, cat, con, met, monto, notas]);
  sh.getRange(sh.getLastRow(), 1).setNumberFormat('dd/MM/yyyy');

  // Copiar a Movimientos Historicos (como ingreso)
  _appendMovHistorico(fecha, mes, semana, 'INGRESO: '+cat, con, met, monto, notas);

  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

/** POST action=ajusteSaldo — guarda saldo base manual (EF y MP).
 *  { action:'ajusteSaldo', efectivo:number, mp:number } */
/** POST action=setOrigenProductos — define origen por producto dentro de un pedido.
 *  { action:'setOrigenProductos', hoja:'Home', id:'H-045', productos:[{a:'PPM',o:'D'},{a:'SCo',o:'OC'}] }
 *  o='D' (Depósito) / o='OC' (Orden de Compra) */
function _doPostSetOrigenProductos(data) {
  var hoja = String(data.hoja || '');
  var pedidoId = String(data.id || '');
  var prods = data.productos || [];

  var sh = SS.getSheetByName(hoja);
  if (!sh) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'hoja' })).setMimeType(ContentService.MimeType.JSON);

  // Buscar fila
  var allData = sh.getDataRange().getValues();
  var row = -1;
  for (var r = 1; r < allData.length; r++) {
    if (String(allData[r][1]).trim() === pedidoId) { row = r + 1; break; }
  }
  if (row === -1) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'no encontrado' })).setMimeType(ContentService.MimeType.JSON);

  // Determinar resumen
  var allD = true, allOC = true;
  prods.forEach(function(p) {
    if (p.o === 'D') allOC = false;
    if (p.o === 'OC') allD = false;
  });
  var summary = allD ? 'Depósito' : allOC ? 'Orden de Compra' : 'Mixto';

  // Columna Origen según hoja
  var colOrigen = hoja === 'Clubes' ? 12 : hoja === 'Red' ? 10 : 9;
  sh.getRange(row, colOrigen).setValue(summary);

  // Guardar detalle JSON en columna "Origen Detalle"
  var headers = allData[0];
  var colDetalle = -1;
  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h]).trim() === 'Origen Detalle') { colDetalle = h + 1; break; }
  }
  if (colDetalle === -1) {
    colDetalle = headers.length + 1;
    sh.getRange(1, colDetalle).setValue('Origen Detalle');
  }
  var detailObj = {};
  prods.forEach(function(p) { detailObj[p.a] = p.o; });
  sh.getRange(row, colDetalle).setValue(JSON.stringify(detailObj));

  // Generar OC solo para productos con origen OC
  var ocProds = prods.filter(function(p) { return p.o === 'OC'; });
  if (ocProds.length > 0) {
    // Verificar que no existan OC previas para este pedido
    var shOC = SS.getSheetByName('Orden de Compra');
    var yaExiste = false;
    if (shOC && shOC.getLastRow() > 1) {
      var existentes = shOC.getRange(2, 6, shOC.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < existentes.length; i++) {
        if (String(existentes[i][0]).trim() === pedidoId) { yaExiste = true; break; }
      }
    }
    if (!yaExiste) {
      _generarOCSelectiva(hoja, row, ocProds.map(function(p) { return p.a; }));
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true, origen: summary })).setMimeType(ContentService.MimeType.JSON);
}

/** Genera OC solo para un subconjunto de productos (abreviaturas).
 *  Reutiliza la lógica de generarOrdenDeCompra pero filtra por abbrs. */
function _generarOCSelectiva(canal, row, abbrs) {
  var shOrigen = SS.getSheetByName(canal);
  var shOC = SS.getSheetByName('Orden de Compra');
  if (!shOrigen || !shOC) return;

  var rowData = shOrigen.getRange(row, 1, 1, shOrigen.getLastColumn()).getValues()[0];
  var colCliente, colPedido, colTelefono;
  var direccion = '';

  if (canal === 'Home') {
    colCliente = 7; colPedido = 1; colTelefono = 44;
    direccion = [rowData[41], rowData[42], 'Lote ' + rowData[43]].filter(Boolean).join(' · ');
  } else if (canal === 'Pilar') {
    colCliente = 7; colPedido = 1; colTelefono = 43;
    direccion = [rowData[41], rowData[42]].filter(Boolean).join(' · ');
  } else if (canal === 'Capital Federal') {
    colCliente = 7; colPedido = 1; colTelefono = 45;
    direccion = [rowData[41], rowData[42] + ' ' + rowData[43], rowData[44]].filter(Boolean).join(' · ');
  } else if (canal === 'Clubes') {
    colCliente = 7; colPedido = 1; colTelefono = 33;
    direccion = [rowData[8], rowData[9], rowData[10]].filter(Boolean).join(' · ');
  }

  // Lookups
  var abbrToProvMap = {}, abbrToNameMap = {}, abbrToCostoMap = {}, abbrToPrecioMap = {};
  var hProvOC = SS.getSheetByName('Proveedores');
  if (hProvOC) {
    var pData = hProvOC.getDataRange().getValues();
    var lastProv = '', lastProd = '';
    for (var p = 1; p < pData.length; p++) {
      if (pData[p][2] && String(pData[p][2]).trim()) lastProv = String(pData[p][2]).trim();
      if (pData[p][1] && String(pData[p][1]).trim()) lastProd = String(pData[p][1]).trim();
      var ab = String(pData[p][4]).trim();
      var gu = String(pData[p][3]).trim();
      if (ab) { abbrToProvMap[ab] = lastProv; abbrToNameMap[ab] = lastProd + (gu ? ' — ' + gu : ''); }
    }
  }
  var hProdOC = SS.getSheetByName('Productos');
  if (hProdOC) {
    var prData = hProdOC.getDataRange().getValues();
    for (var q = 1; q < prData.length; q++) {
      var abr = String(prData[q][2]).trim();
      if (!abr) continue;
      abbrToCostoMap[abr] = parseFloat(String(prData[q][9]).replace(/[$.]/g,'').replace(/,/g,'')) || 0;
      abbrToPrecioMap[abr] = parseFloat(String(prData[q][8]).replace(/[$.]/g,'').replace(/,/g,'')) || 0;
    }
  }

  var CLUBES_PRECIOS = {'PMu':7000,'PMa':7000,'PJyQ':7000,'PCC':7000,'PJyM':7800,'PPM':11000,'PPJyQ':11000,'PPCyQ':11000};
  var colToAbbrMap = (canal === 'Clubes')
    ? {24:'PMu',25:'PMa',26:'PJyQ',27:'PCC',28:'PJyM',29:'PPM',30:'PPJyQ',31:'PPCyQ'}
    : HOME_COL_TO_ABBR;

  var cliente = String(rowData[colCliente] || '');
  var numPedido = String(rowData[colPedido] || '');
  var telefono = String(rowData[colTelefono] || '');

  var ahora = new Date();
  var argDate = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var dd = String(argDate.getDate()).padStart(2,'0');
  var mm = String(argDate.getMonth()+1).padStart(2,'0');
  var yyyy = argDate.getFullYear();
  var hh = String(argDate.getHours()).padStart(2,'0');
  var mi = String(argDate.getMinutes()).padStart(2,'0');
  var fechaStr = dd+'/'+mm+'/'+yyyy+' '+hh+':'+mi;
  var semana = _isoWeek(argDate);
  var mesNombre = MESES[argDate.getMonth()];

  // Solo generar OC para las abreviaturas indicadas
  var abbrSet = {};
  abbrs.forEach(function(a) { abbrSet[a] = true; });

  var newRows = [];
  Object.keys(colToAbbrMap).forEach(function(colStr) {
    var colIdx = Number(colStr);
    var abbr = colToAbbrMap[colIdx];
    if (!abbrSet[abbr]) return; // Skip si no está en la lista OC
    var qty = Number(rowData[colIdx - 1]) || 0;
    if (qty === 0) return;

    var ocId = _nextId('OC-');
    var costoUnit = abbrToCostoMap[abbr] || 0;
    var costoTotal = costoUnit * qty;
    var precioVenta = (canal === 'Clubes') ? (CLUBES_PRECIOS[abbr] || 0) : (abbrToPrecioMap[abbr] || 0);

    newRows.push([
      ocId, fechaStr, semana, mesNombre, canal, numPedido, cliente, telefono, direccion,
      abbrToProvMap[abbr] || '', abbrToNameMap[abbr] || abbr, abbr, qty,
      costoUnit, costoTotal, precioVenta, 0, 0, 0,
      'Orden de Compra', 'Pendiente', '', '', 'No', 'No'
    ]);
  });

  if (newRows.length > 0) {
    newRows.forEach(function(r) { shOC.appendRow(r); });
    var actualLastRow = shOC.getLastRow();
    var formulaStart = actualLastRow - newRows.length + 1;
    for (var i = 0; i < newRows.length; i++) {
      var rr = formulaStart + i;
      shOC.getRange(rr, 17).setFormula('=P'+rr+'*M'+rr);
      shOC.getRange(rr, 18).setFormula('=Q'+rr+'-O'+rr);
      shOC.getRange(rr, 19).setFormula('=R'+rr+'/Q'+rr);
    }
    shOC.getRange(formulaStart, 13, newRows.length, 1).setNumberFormat('0');
    shOC.getRange(formulaStart, 14, newRows.length, 2).setNumberFormat('$#,##0');
    shOC.getRange(formulaStart, 16, newRows.length, 3).setNumberFormat('$#,##0');
    shOC.getRange(formulaStart, 19, newRows.length, 1).setNumberFormat('0.0%');
  }
}

/** POST action=marcarCobrado — marca un pedido como Cobrado desde el Panel.
 *  { action:'marcarCobrado', hoja:'Clubes', id:'C-005' } */
function _doPostMarcarCobrado(data) {
  var hoja = String(data.hoja || '');
  var pedidoId = String(data.id || '');

  var sh = SS.getSheetByName(hoja);
  if (!sh) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'hoja' })).setMimeType(ContentService.MimeType.JSON);

  // Buscar fila por N° Pedido (col B=2)
  var allData = sh.getDataRange().getValues();
  var row = -1;
  for (var r = 1; r < allData.length; r++) {
    if (String(allData[r][1]).trim() === pedidoId) { row = r + 1; break; }
  }
  if (row === -1) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'no encontrado' })).setMimeType(ContentService.MimeType.JSON);

  // Columna Estado de Pago según hoja
  var colPago = hoja === 'Clubes' ? 16 : hoja === 'Red' ? 14 : 13; // Clubes=P(16), Red=N(14), Home/Pilar/CF=M(13)
  sh.getRange(row, colPago).setValue('Cobrado');

  // Propina (opcional)
  var propina = Number(data.propina) || 0;
  var propMet = String(data.propMet || '').trim();
  if (propina > 0 && propMet) {
    if (hoja === 'Clubes') {
      // Clubes: Propina Ef=U(21), Propina Trans=V(22)
      var colProp = propMet === 'Efectivo' ? 21 : 22;
    } else if (hoja === 'Red') {
      // Red: Propina Ef=S(19), Propina Trans=T(20)
      var colProp = propMet === 'Efectivo' ? 19 : 20;
    } else {
      // Home/Pilar/CF: Propina Ef=R(18), Propina Trans=S(19)
      var colProp = propMet === 'Efectivo' ? 18 : 19;
    }
    var propActual = Number(sh.getRange(row, colProp).getValue()) || 0;
    sh.getRange(row, colProp).setValue(propActual + propina);
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

function _doPostAjusteSaldo(data) {
  var deseadoEf = Number(data.efectivo) || 0;
  var deseadoMP = Number(data.mp) || 0;

  // ── Calcular cobrados actuales (pedidos operativos cobrados) ──
  var cobradoEf = 0, cobradoMP = 0;
  ['Home', 'Pilar', 'Capital Federal'].forEach(function(hoja) {
    var sh = SS.getSheetByName(hoja);
    if (!sh || sh.getLastRow() <= 1) return;
    var d = sh.getDataRange().getValues();
    for (var r = 1; r < d.length; r++) {
      if (String(d[r][12]).trim() === 'Cobrado') {
        var fac = Number(d[r][19]) || Number(d[r][13]) || 0;
        if (String(d[r][11]).trim() === 'Efectivo') cobradoEf += fac;
        else cobradoMP += fac;
      }
    }
  });
  var shCl = SS.getSheetByName('Clubes');
  if (shCl && shCl.getLastRow() > 1) {
    var dCl = shCl.getDataRange().getValues();
    for (var rc = 1; rc < dCl.length; rc++) {
      if (String(dCl[rc][15]).trim() === 'Cobrado') {
        var facC = Number(dCl[rc][22]) || Number(dCl[rc][16]) || 0;
        if (String(dCl[rc][14]).trim() === 'Efectivo') cobradoEf += facC;
        else cobradoMP += facC;
      }
    }
  }

  // ── Calcular gastos actuales ──
  var gastosEf = 0, gastosMP = 0;
  var shEgr = SS.getSheetByName('Egresos');
  if (shEgr && shEgr.getLastRow() > 1) {
    var dE = shEgr.getDataRange().getValues();
    for (var re = 1; re < dE.length; re++) {
      if (!dE[re][0]) continue;
      var mE = Number(dE[re][6]) || 0;
      if (String(dE[re][5]).trim() === 'Efectivo') gastosEf += mE;
      else gastosMP += mE;
    }
  }

  // ── Calcular ingresos actuales ──
  var ingresosEf = 0, ingresosMP = 0;
  var shIng = SS.getSheetByName('Ingresos');
  if (shIng && shIng.getLastRow() > 1) {
    var dI = shIng.getDataRange().getValues();
    for (var ri = 1; ri < dI.length; ri++) {
      if (!dI[ri][0]) continue;
      var mI = Number(dI[ri][6]) || 0;
      if (String(dI[ri][5]).trim() === 'Efectivo') ingresosEf += mI;
      else ingresosMP += mI;
    }
  }

  // ── Back-calculate base so that: base + cobrado + ingresos - gastos = deseado ──
  var baseEf = deseadoEf - cobradoEf - ingresosEf + gastosEf;
  var baseMP = deseadoMP - cobradoMP - ingresosMP + gastosMP;

  // ── Guardar ──
  var shSaldo = SS.getSheetByName('Saldo Base');
  if (!shSaldo) {
    shSaldo = SS.insertSheet('Saldo Base');
    shSaldo.getRange(1, 1, 1, 3).setValues([['Fecha', 'Efectivo', 'Mercado Pago']]);
    shSaldo.setFrozenRows(1);
    shSaldo.getRange(1, 1, 1, 3).setBackground(BROWN).setFontColor('#FFFFFF').setFontWeight('bold');
  }
  var ahora = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  shSaldo.appendRow([argNow, baseEf, baseMP]);
  shSaldo.getRange(shSaldo.getLastRow(), 1).setNumberFormat('dd/MM/yyyy HH:mm');

  return ContentService.createTextOutput(JSON.stringify({ ok: true, ef: deseadoEf, mp: deseadoMP })).setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════════
//  VENTAS — Lee Archivo + Operativas (entregados) + Historico
// ══════════════════════════════════════════════════════════════

/** GET ?action=ventas
 *  Devuelve TODAS las ventas: archivadas + semana actual (entregados) + histórico legacy.
 *  Canales: Venta Directa (Home+Pilar+CF), Clubes, Red, B2B, Catering. */
function _doGetVentas() {
  var ventas = [];
  var MVAL = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

  // ── Helper: parsear fila VD (Home/Pilar/CF) — misma estructura en operativa, archivo e historico ──
  function parseVD(data, r, zona) {
    var cliente = String(data[r][7] || '').trim();
    if (!cliente) return null;
    var facturado = Number(data[r][19]) || Number(data[r][13]) || 0;
    if (facturado === 0) return null;
    var fechaEnt = data[r][47];
    var fecha = fechaEnt instanceof Date ? fechaEnt : data[r][3];
    var fechaStr = fecha instanceof Date ? Utilities.formatDate(fecha, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : String(fecha || '');
    var mesEnt = String(data[r][48] || '').trim();
    var semEnt = Number(data[r][49]) || 0;
    var mesPed = String(data[r][4] || '').trim();
    var semPed = Number(data[r][5]) || 0;
    var mes = MVAL.indexOf(mesEnt) >= 0 ? mesEnt : mesPed;
    var sem = semEnt > 0 ? semEnt : semPed;
    return {
      canal: 'Venta Directa', zona: zona, fecha: fechaStr, mes: mes, sem: sem,
      cliente: cliente, estado: String(data[r][10] || '').trim(),
      fp: String(data[r][11] || '').trim(), ep: String(data[r][12] || '').trim(),
      $: facturado, ef: Number(data[r][15]) || 0, tr: Number(data[r][16]) || 0,
      costo: Number(data[r][39]) || 0, margen: Number(data[r][40]) || 0
    };
  }

  // ── Helper: parsear fila Clubes ──
  function parseClubes(data, r) {
    var cli = String(data[r][7] || '').trim();
    if (!cli) return null;
    var club = String(data[r][8] || '').trim();
    var fac = Number(data[r][22]) || Number(data[r][16]) || 0;
    if (fac === 0) return null;
    var fC = data[r][3];
    var fCS = fC instanceof Date ? Utilities.formatDate(fC, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : String(fC || '');
    return {
      canal: 'Clubes', zona: club, fecha: fCS,
      mes: String(data[r][4] || '').trim(), sem: Number(data[r][5]) || 0,
      cliente: cli + (club ? ' (' + club + ')' : ''),
      estado: String(data[r][13] || '').trim(),
      fp: String(data[r][14] || '').trim(), ep: String(data[r][15] || '').trim(),
      $: fac, ef: Number(data[r][18]) || 0, tr: Number(data[r][19]) || 0,
      costo: Number(data[r][31]) || 0, margen: Number(data[r][32]) || 0
    };
  }

  // ── Helper: leer hoja completa con parser ──
  function readSheet(sheetName, parser, parserArg) {
    var sh = SS.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() <= 1) return;
    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var v = parser(data, r, parserArg);
      if (v) ventas.push(v);
    }
  }

  // ── Helper: leer solo Entregados de hoja operativa ──
  function readOperativa(sheetName, parser, parserArg, colEstado) {
    var sh = SS.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() <= 1) return;
    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var estado = String(data[r][colEstado] || '').trim();
      if (estado !== 'Entregado') continue;
      var v = parser(data, r, parserArg);
      if (v) ventas.push(v);
    }
  }

  // ── Venta Directa: Archivo + Operativa (entregados) + Historico legacy ──
  var vdCfg = [
    { zona: 'Home' },
    { zona: 'Pilar' },
    { zona: 'Capital Federal' }
  ];
  vdCfg.forEach(function(cfg) {
    readSheet('Archivo ' + cfg.zona, parseVD, cfg.zona);
    readOperativa(cfg.zona, parseVD, cfg.zona, 10);
    readSheet('Historico ' + cfg.zona, parseVD, cfg.zona);
  });

  // ── Clubes: Archivo + Operativa (entregados) + Historico legacy ──
  readSheet('Archivo Clubes', parseClubes);
  readOperativa('Clubes', parseClubes, null, 13);
  readSheet('Historico Clubes', parseClubes);

  // ── Red (headers row 1, data from row 2) ──
  var shRd = SS.getSheetByName('Historico Red');
  if (shRd && shRd.getLastRow() > 1) {
    var dRd = shRd.getDataRange().getValues();
    for (var rr = 1; rr < dRd.length; rr++) {
      var cliR = String(dRd[rr][5] || '').trim();
      if (!cliR) continue;
      var facR = Number(dRd[rr][12]) || 0;
      if (facR === 0) continue;
      var fR = dRd[rr][1];
      var fRS = fR instanceof Date ? Utilities.formatDate(fR, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : String(fR || '');
      ventas.push({
        canal: 'Red', zona: String(dRd[rr][7] || '').trim(), fecha: fRS,
        mes: String(dRd[rr][2] || '').trim(), sem: Number(dRd[rr][3]) || 0,
        cliente: cliR, estado: 'Entregado',
        fp: Number(dRd[rr][14]) > 0 ? 'Transferencia' : 'Efectivo', ep: 'Cobrado',
        $: facR, ef: Number(dRd[rr][13]) || 0, tr: Number(dRd[rr][14]) || 0,
        costo: 0, margen: 0
      });
    }
  }

  // ── B2B (headers row 1, data from row 2) ──
  var shB = SS.getSheetByName('Historico B2B');
  if (shB && shB.getLastRow() > 1) {
    var dB = shB.getDataRange().getValues();
    for (var rb = 1; rb < dB.length; rb++) {
      var cliB = String(dB[rb][5] || '').trim();
      if (!cliB) continue;
      var facB = Number(dB[rb][8]) || 0;
      if (facB === 0) continue;
      var fB = dB[rb][1];
      var fBS = fB instanceof Date ? Utilities.formatDate(fB, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : String(fB || '');
      ventas.push({
        canal: 'B2B', zona: '', fecha: fBS,
        mes: String(dB[rb][2] || '').trim(), sem: Number(dB[rb][3]) || 0,
        cliente: cliB, estado: 'Entregado',
        fp: Number(dB[rb][10]) > 0 ? 'Transferencia' : 'Efectivo', ep: 'Cobrado',
        $: facB, ef: Number(dB[rb][9]) || 0, tr: Number(dB[rb][10]) || 0,
        costo: 0, margen: 0
      });
    }
  }

  // ── Catering (headers row 1, data from row 2) ──
  var shCt = SS.getSheetByName('Historico Catering');
  if (shCt && shCt.getLastRow() > 1) {
    var dCt = shCt.getDataRange().getValues();
    for (var rt = 1; rt < dCt.length; rt++) {
      var cliT = String(dCt[rt][5] || '').trim();
      if (!cliT) continue;
      var facT = Number(dCt[rt][8]) || 0;
      if (facT === 0) continue;
      var fT = dCt[rt][1];
      var fTS = fT instanceof Date ? Utilities.formatDate(fT, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : String(fT || '');
      ventas.push({
        canal: 'Catering', zona: String(dCt[rt][7] || '').trim(), fecha: fTS,
        mes: String(dCt[rt][2] || '').trim(), sem: Number(dCt[rt][3]) || 0,
        cliente: cliT, estado: 'Entregado',
        fp: Number(dCt[rt][11]) > 0 ? 'Transferencia' : 'Efectivo', ep: 'Cobrado',
        $: facT, ef: Number(dCt[rt][10]) || 0, tr: Number(dCt[rt][11]) || 0,
        costo: Number(dCt[rt][17]) || 0, margen: Number(dCt[rt][22]) || 0
      });
    }
  }

  // ── Deduplicar por N° Pedido (puede estar en operativa + archivo si justo se archivó) ──
  var seen = {};
  ventas = ventas.filter(function(v) {
    var key = v.canal + '|' + v.cliente + '|' + v.fecha + '|' + v.$;
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });

  return ContentService
    .createTextOutput(JSON.stringify({ ts: Date.now(), v: ventas }))
    .setMimeType(ContentService.MimeType.JSON);
}
