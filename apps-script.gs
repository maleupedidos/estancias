/**
 * MALEU — Apps Script v3
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

const SS      = SpreadsheetApp.getActiveSpreadsheet();
const BROWN   = '#3D1C0A';
const ORANGE  = '#F97035';
const CREAM   = '#E8DFC4';

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
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function _doPostPedido(data) {
    // Escribir en hoja Home — el stock se maneja vía onEdit (J=Reservado/Entregado)
    _doPostHome(data);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  _doPostHome — escribe en la hoja "Home" (canal Estancias)
//  Columnas A–AP según el esquema de análisis de Tadeo
// ════════════════════════════════════════════════════════════

// Mapeo id de producto (página web) → columna 1-based de la hoja Home
// Con col A=Hora, los productos empiezan en col S (19) hasta AJ (36)
const HOME_PRODUCT_COLS = {
  5:  19,  // PPM   — Pack Muzarella x2
  6:  20,  // PPJyQ — Pack Jamón y Queso x2
  7:  21,  // PPCyQ — Pack Cebolla y Queso x2
  8:  22,  // SCo   — Sorrentinos Cordero al Malbec
  9:  23,  // SJyQ  — Sorrentinos Jamón y Queso
  10: 24,  // SCa   — Sorrentinos Calabaza y Queso
  11: 25,  // ECaC  — Empanadas Carne a Cuchillo x8
  12: 26,  // EJyQ  — Empanadas Jamón y Queso x8
  17: 27,  // ECyQ  — Empanadas Cebolla y Queso x8
  18: 28,  // EV    — Empanadas Verdura x8
  14: 29,  // TG    — Torta Golosa
  15: 30,  // TLC   — Torta Lemon Crumble
  16: 31,  // TC    — Torta Coco
  13: 32,  // F     — Franui Leche
  1:  33,  // PMar  — Pizza Margarita
  2:  34,  // PJyQ  — Pizza Jamón y Queso
  3:  35,  // PCC   — Pizza Cebolla Caramelizada
  4:  36,  // PJyM  — Pizza Jamón y Morrón
};

// Mapeo id de producto (página web) → abreviatura en hoja Productos (col C)
const PAGE_ID_TO_ABBR = {
  5:  'PPM',   6:  'PPJyQ', 7:  'PPCyQ',
  8:  'SCo',   9:  'SJyQ',  10: 'SCa',
  11: 'ECaC',  12: 'EJyQ',  17: 'ECyQ', 18: 'EV',
  14: 'TG',    15: 'TLC',   16: 'TC',   13: 'F',
  1:  'PMar',  2:  'PJyQ',  3:  'PCC',  4:  'PJyM',
};

function _doPostHome(data) {
  const sh = SS.getSheetByName('Home');
  if (!sh) return; // Si la hoja no existe, silencio — no romper el flujo

  // ── N° de pedido autoincremental H-XXX (ahora en col B) ──
  const lastRow = sh.getLastRow();
  let maxNum = 0;
  if (lastRow > 1) {
    const colB = sh.getRange(2, 2, lastRow - 1, 1).getValues(); // col B = N° Pedido
    colB.forEach(function(row) {
      const match = String(row[0]).match(/^H-(\d+)$/);
      if (match) {
        const n = parseInt(match[1], 10);
        if (n > maxNum) maxNum = n;
      }
    });
  }
  const orderNum = 'H-' + String(maxNum + 1).padStart(3, '0');

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
  const total         = Number(data.total) || 0;
  const pago          = String(data.pago || '');
  const efectivo      = pago === 'Efectivo'      ? total : 0;
  const transferencia = pago === 'Transferencia' ? total : 0;

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
  const row = new Array(48).fill('');
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
  row[14] = efectivo;                           // O  Efectivo ($)
  row[15] = transferencia;                      // P  Transferencia ($)
  row[16] = 0;                                  // Q  Propina Efectivo (default $0)
  row[17] = 0;                                  // R  Propina Transferencia (default $0)

  // Productos: cols S–AJ (índices 18–35 en base-0)
  Object.keys(HOME_PRODUCT_COLS).forEach(function(id) {
    row[HOME_PRODUCT_COLS[id] - 1] = qtys[Number(id)] || 0;
  });

  row[36] = costoTotal;                         // AK  Costo
  row[37] = total - costoTotal;                 // AL  Margen Bruto (Cobrado - Costo)
  row[38] = barrioPrivado;                      // AM  Barrio
  row[39] = subBarrio;                          // AN  Sub Barrio
  row[40] = String(data.lote || '');            // AO  Domicilio - Lote
  row[41] = String(data.telefono || '');        // AP  Teléfono
  // AQ-AV (indices 42-47) = Fecha/Hora/Día/Semana/Año Entrega → se llenan al marcar K="Entregado"

  sh.appendRow(row);
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
      const celda = hProd.getRange(r + 1, 3); // col 3 = Stock Físico
      celda.setValue((celda.getValue() || 0) + cantidad);
      break;
    }
  }
}

// ════════════════════════════════════════════════════════════
//  onEdit — actualiza stock al cambiar estado de un pedido
// ════════════════════════════════════════════════════════════

// Mapeo inverso: columna Home (1-based) → Abreviatura en Productos (col C)
// Con col A=Hora, productos van de S(19) a AH(34)
const HOME_COL_TO_ABBR = {
  19: 'PPM',   // S
  20: 'PPJyQ', // T
  21: 'PPCyQ', // U
  22: 'SCo',   // V
  23: 'SJyQ',  // W
  24: 'SCa',   // X
  25: 'ECaC',  // Y
  26: 'EJyQ',  // Z
  27: 'ECyQ',  // AA
  28: 'EV',    // AB
  29: 'TG',    // AC
  30: 'TLC',   // AD
  31: 'TC',    // AE
  32: 'F',     // AF
  33: 'PMar',  // AG
  34: 'PJyQ',  // AH
  35: 'PCC',   // AI
  36: 'PJyM',  // AJ
};

// IMPORTANTE: esta función debe configurarse SOLO como trigger instalable.
// NO usar el nombre "onEdit" para evitar doble ejecución (simple + instalable).
// En Activadores: función = onEditHandler, evento = Al editar
function onEditHandler(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  if (sheetName === 'Home')    return _onEditHome(e);
  if (sheetName === 'Pedidos') return _onEditPedidos(e);
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

/**
 * Genera filas en Orden de Compra para un pedido dado.
 * @param {string} canal - 'Home', 'Delivery' o 'Clubes'
 * @param {number} row - fila del pedido en la hoja de origen
 */
function generarOrdenDeCompra(canal, row) {
  const shOrigen = SS.getSheetByName(canal);
  const shOC     = SS.getSheetByName('Orden de Compra');
  if (!shOrigen || !shOC) return;

  const rowData = shOrigen.getRange(row, 1, 1, shOrigen.getLastColumn()).getValues()[0];

  // Determinar columnas de productos según canal
  let prodStartCol, prodEndCol, colCliente, colPedido, colTelefono, colDiaEntrega;
  let direccion = '';

  if (canal === 'Home') {
    prodStartCol = 18; prodEndCol = 35; // S(18) a AJ(35) en 0-based
    colCliente = 7; colPedido = 1; colTelefono = 41; colDiaEntrega = 9;
    direccion = [rowData[38], rowData[39], 'Lote ' + rowData[40]].filter(Boolean).join(' · ');
  } else if (canal === 'Delivery') {
    prodStartCol = 18; prodEndCol = 35;
    colCliente = 7; colPedido = 1; colTelefono = 41; colDiaEntrega = 9;
    direccion = [rowData[38], rowData[39], 'Lote ' + rowData[40]].filter(Boolean).join(' · ');
  } else if (canal === 'Clubes') {
    prodStartCol = 21; prodEndCol = 23; // PPM(21) a PPCyQ(23) en 0-based
    colCliente = 7; colPedido = 1; colTelefono = 26; colDiaEntrega = 12;
    direccion = String(rowData[8] || '') + ' — ' + String(rowData[9] || ''); // Club + Deporte
  }

  const abbrToProvMap = _getAbbrToProveedor();
  const abbrToNameMap = _getAbbrToProductName();
  const abbrToCostoMap = _getAbbrToCosto();

  // Obtener abreviaturas en orden de columna
  const colToAbbrMap = (canal === 'Clubes')
    ? {21:'PPM', 22:'PPJyQ', 23:'PPCyQ'}
    : HOME_COL_TO_ABBR;

  const cliente    = String(rowData[colCliente] || '');
  const numPedido  = String(rowData[colPedido] || '');
  const telefono   = String(rowData[colTelefono] || '');
  const diaEntrega = String(rowData[colDiaEntrega] || '');

  // N° de Orden autoincremental
  const ocData = shOC.getDataRange().getValues();
  let maxOC = 0;
  for (let r = 1; r < ocData.length; r++) {
    const match = String(ocData[r][0]).match(/^OC-(\d+)$/);
    if (match) { const n = parseInt(match[1]); if (n > maxOC) maxOC = n; }
  }

  const ahora   = new Date();
  const argDate = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  const dd      = String(argDate.getDate()).padStart(2, '0');
  const mm      = String(argDate.getMonth() + 1).padStart(2, '0');
  const yyyy    = argDate.getFullYear();
  const fechaStr = dd + '/' + mm + '/' + yyyy;

  const newRows = [];
  Object.keys(colToAbbrMap).forEach(function(colStr) {
    const colIdx = Number(colStr);
    const abbr   = colToAbbrMap[colIdx];
    const qty    = Number(rowData[colIdx - 1]) || 0; // 0-based
    if (qty === 0) return;

    maxOC++;
    const costoUnit  = abbrToCostoMap[abbr] || 0;
    const costoTotal = costoUnit * qty;

    newRows.push([
      'OC-' + String(maxOC).padStart(3, '0'),  // A  N° Orden
      fechaStr,                                   // B  Fecha Generada
      canal,                                      // C  Canal
      numPedido,                                  // D  N° Pedido Origen
      cliente,                                    // E  Cliente
      telefono,                                   // F  Teléfono
      direccion,                                  // G  Dirección Entrega
      abbrToProvMap[abbr] || '',                  // H  Proveedor
      abbrToNameMap[abbr] || abbr,                // I  Producto
      abbr,                                       // J  Abreviatura
      qty,                                        // K  Cantidad
      costoUnit,                                  // L  Costo Unitario
      costoTotal,                                 // M  Costo Total
      'Pendiente',                                // N  Estado
      '',                                         // O  Fecha Pedido Proveedor
      '',                                         // P  Fecha Búsqueda
      '',                                         // Q  Fecha Recibido
      diaEntrega,                                 // R  Día Entrega Cliente
      'No',                                       // S  Pagado
    ]);
  });

  if (newRows.length > 0) {
    shOC.getRange(shOC.getLastRow() + 1, 1, newRows.length, 19).setValues(newRows);
    // Formato moneda en costo
    const startRow = shOC.getLastRow() - newRows.length + 1;
    shOC.getRange(startRow, 12, newRows.length, 2).setNumberFormat('$#,##0');
  }
}

// ── Home: sync stock + auto-generar Orden de Compra ──
// Col I (9) = Origen → si cambia a "Orden de Compra", genera filas en OC
// Col K (11) = Estado de Entrega → si cambia a "Entregado", descuenta stock
function _onEditHome(e) {
  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row <= 1) return;

  // Si cambió Origen (col I=9) a "Orden de Compra" → generar OC
  if (col === 9) {
    const nuevoOrigen = String(e.value || '');
    if (nuevoOrigen === 'Orden de Compra') {
      generarOrdenDeCompra('Home', row);
      SS.toast('Orden de Compra generada para ' + e.range.getSheet().getRange(row, 2).getValue(), 'OC', 5);
    }
    return;
  }

  if (col !== 11) return; // solo col K (11) = Estado de Entrega

  const sh     = e.range.getSheet();
  const origen = String(sh.getRange(row, 9).getValue()); // col I (9) = Origen
  if (origen !== 'Depósito') return;

  const nuevo    = String(e.value || '');
  const anterior = String(e.oldValue || '');

  const hProductos = SS.getSheetByName('Productos');
  if (!hProductos) return;

  // → Entregado: descontar Stock Físico + registrar fecha/hora de entrega
  if (nuevo === 'Entregado' && anterior !== 'Entregado') {
    _homeStockFisico(sh, row, hProductos, -1);
    _registrarFechaEntrega(sh, row);
  }

  // ← Sale de Entregado (corrección manual): devolver Stock Físico + borrar fecha entrega
  if (anterior === 'Entregado' && nuevo !== 'Entregado') {
    _homeStockFisico(sh, row, hProductos, +1);
    sh.getRange(row, 43, 1, 6).clearContent(); // limpiar AQ-AV
  }
}

// Llena cols AO-AT con fecha/hora de entrega en zona Argentina
function _registrarFechaEntrega(sh, row) {
  var ahora   = new Date();
  var argDate = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var DIAS    = ['Domingo','Lunes','Martes','Mi\u00E9rcoles','Jueves','Viernes','S\u00E1bado'];
  var MESES   = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var dd      = String(argDate.getDate()).padStart(2, '0');
  var mm      = String(argDate.getMonth() + 1).padStart(2, '0');
  var yyyy    = argDate.getFullYear();
  var hh      = String(argDate.getHours()).padStart(2, '0');
  var mi      = String(argDate.getMinutes()).padStart(2, '0');

  sh.getRange(row, 43, 1, 6).setValues([[
    hh + ':' + mi,                     // AQ  Hora Entrega
    DIAS[argDate.getDay()],            // AR  Día Entrega
    dd + '/' + mm + '/' + yyyy,        // AS  Fecha Entrega
    MESES[argDate.getMonth()],         // AT  Mes Entrega
    _isoWeek(argDate),                 // AU  Semana Entrega
    yyyy                               // AV  Año Entrega
  ]]);
}

// Ajusta Stock Físico (col F=6) de Productos. signo: -1 = restar, +1 = sumar
function _homeStockFisico(shHome, row, hProductos, signo) {
  const cantidades = shHome.getRange(row, 19, 1, 18).getValues()[0]; // cols S–AJ
  const prodData   = hProductos.getDataRange().getValues();

  Object.keys(HOME_COL_TO_ABBR).forEach(function(colStr) {
    const colIdx = Number(colStr);
    const abbr   = HOME_COL_TO_ABBR[colIdx];
    const qty    = Number(cantidades[colIdx - 19]) || 0;
    if (qty === 0) return;

    for (let r = 1; r < prodData.length; r++) {
      if (String(prodData[r][2]).trim() === abbr) {
        const celdaFis = hProductos.getRange(r + 1, 6); // F = Stock Físico
        const fisico   = Number(celdaFis.getValue()) || 0;
        celdaFis.setValue(Math.max(0, fisico + (qty * signo)));
        break;
      }
    }
  });
}

// ── Fórmulas en Productos: Reservado (E) y Disponible (F) ───
// Ejecutar UNA vez — pone fórmulas SUMPRODUCT que se auto-actualizan.
// Abreviatura (col C de Productos) → letra de columna en Home
const ABBR_TO_HOME_COL = {
  'PPM':'S', 'PPJyQ':'T', 'PPCyQ':'U',
  'SCo':'V', 'SJyQ':'W', 'SCa':'X',
  'ECaC':'Y', 'EJyQ':'Z', 'ECyQ':'AA', 'EV':'AB',
  'TG':'AC', 'TLC':'AD', 'TC':'AE', 'F':'AF',
  'PMar':'AG', 'PJyQ':'AH', 'PCC':'AI', 'PJyM':'AJ',
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

    // Col E (Vendidos Semana) = SUMPRODUCT: Entregados por semana de ENTREGA (AR/AS)
    hProd.getRange(rowNum, 5).setFormula(
      '=SUMPRODUCT((Home!$I$2:$I$10000="' + dep + '")*(Home!$K$2:$K$10000="Entregado")' +
      '*(Home!$AS$2:$AS$10000=' + semanaActual + ')*(Home!$AT$2:$AT$10000=' + anioActual + ')' +
      '*(Home!' + homeCol + '$2:' + homeCol + '$10000))'
    );

    // Col G (Reservado) = SUMPRODUCT: Reservados activos
    hProd.getRange(rowNum, 7).setFormula(
      '=SUMPRODUCT((Home!$I$2:$I$10000="' + dep + '")*(Home!$K$2:$K$10000="Reservado")' +
      '*(Home!' + homeCol + '$2:' + homeCol + '$10000))'
    );

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
//  setupSheets — ejecutar UNA sola vez para formatear todo
// ════════════════════════════════════════════════════════════
function setupSheets() {
  _setupPedidos();
  _setupProductos();
  _setupDetalle();
  _setupPanel();
  _setupProveedores();
  _setupEgresos();
  _setupHome();
  _setupOrdenDeCompra();
  SS.toast('✅  Sheets de Maleu configurados correctamente', 'Setup completo', 5);
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
    'TG','TLC','TC','F','PMar','PJyQ','PCC','PJyM',
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

// ── Hoja: Orden de Compra — vista automática desde Home ─────
function _setupOrdenDeCompra() {
  let sh = SS.getSheetByName('Orden de Compra');
  if (!sh) sh = SS.insertSheet('Orden de Compra');

  // Limpiar contenido previo
  sh.clear();

  // ── Título ──────────────────────────────────────────────────
  sh.setRowHeight(1, 48);
  sh.getRange('A1:H1').merge()
    .setValue('📋  ÓRDENES DE COMPRA — Vista automática desde Home')
    .setBackground(BROWN).setFontColor('#FFFFFF')
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── QUERY automática ────────────────────────────────────────
  // Trae las columnas más útiles de Home donde Origen = "Orden de Compra"
  // A=Hora, B=N°Pedido, D=Fecha, H=Cliente, K=EstadoEntrega, L=FormaPago,
  // M=EstadoPago, N=Total, AK=Barrio, AL=SubBarrio, AM=Lote, AN=Teléfono
  sh.getRange('A3').setFormula(
    '=IFERROR(' +
      'QUERY(Home!A:AN,' +
        '"SELECT B, A, D, H, K, L, M, N, AK, AL, AM, AN ' +
        'WHERE I = \'Orden de Compra\' ' +
        'ORDER BY D DESC, A DESC",1)' +
    ',"No hay órdenes de compra todavía")'
  );

  // ── Formato de la zona de datos ─────────────────────────────
  // Encabezados del QUERY (fila 3) se formatean automáticamente
  sh.getRange('A3:L3')
    .setBackground('#E8DFC4').setFontColor(BROWN)
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');

  // Ancho de columnas
  const widths = [
    95,  // A  N° Pedido
    65,  // B  Hora
    100, // C  Fecha
    170, // D  Cliente
    130, // E  Estado de Entrega
    120, // F  Forma de Pago
    120, // G  Estado de Pago
    95,  // H  Total ($)
    140, // I  Barrio
    140, // J  Sub Barrio
    130, // K  Domicilio
    130, // L  Teléfono
  ];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // Formato moneda en Total (col H en esta hoja)
  sh.getRange('H4:H5000').setNumberFormat('$#,##0');

  sh.setFrozenRows(3);
  sh.setTabColor('#0D47A1'); // azul — distinguir de Home (naranja)

  // ── Nota informativa ────────────────────────────────────────
  sh.getRange('A2').setValue('⚡ Esta hoja se actualiza automáticamente. Para gestionar pedidos, editá la hoja Home.')
    .setFontSize(9).setFontColor('#666666').setFontStyle('italic');
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
