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
  if (action === 'compras') {
    return _doGetCompras();
  }
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

// ════════════════════════════════════════════════════════════
//  doPost — recibe pedidos desde la página + acciones internas
// ════════════════════════════════════════════════════════════
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'compra')       return _doPostCompra(data);
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
    const ahora    = new Date();
    const idPedido = 'P-' + ahora.getTime();
    const fecha    = Utilities.formatDate(ahora, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');
    const fechaDate = Utilities.formatDate(ahora, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
    const items    = data.items || [];

    // Hoja Pedidos
    const hPedidos = SS.getSheetByName('Pedidos');
    hPedidos.appendRow([
      idPedido,
      fecha,
      data.nombre   || '',
      data.barrio   || '',
      data.lote     || '',
      data.telefono || '',
      data.dia      || '',
      data.horario  || '',
      data.pago     || '',
      data.total    || 0,
      'pendiente',
      fechaDate
    ]);

    // Hoja Detalle_Pedidos + reservar stock
    // Clubes es por encargo: NO toca el stock del depósito
    const esClub     = String(data.barrio || '').startsWith('Club-');
    const hDetalle   = SS.getSheetByName('Detalle_Pedidos');
    const hProductos = SS.getSheetByName('Productos');
    const prodData   = hProductos.getDataRange().getValues();

    items.forEach(item => {
      hDetalle.appendRow([
        idPedido, item.id, item.nombre, item.qty, item.precio,
        item.qty * item.precio
      ]);

      if (!esClub) {
        for (let r = 1; r < prodData.length; r++) {
          if (prodData[r][0] === item.id) {
            const cell = hProductos.getRange(r + 1, 4);
            cell.setValue((cell.getValue() || 0) + item.qty);
            break;
          }
        }
      }
    });

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, id: idPedido }))
      .setMimeType(ContentService.MimeType.JSON);
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

function _doUpdateCompra(data) {
  const sh = SS.getSheetByName('Pedidos_Proveedores');
  if (!sh) throw new Error('Hoja Pedidos_Proveedores no encontrada.');

  const shData = sh.getDataRange().getValues();
  for (let r = 1; r < shData.length; r++) {
    if (String(shData[r][0]) === String(data.id)) {
      sh.getRange(r + 1, 9).setValue(data.estado); // col 9 = Estado
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: 'ID no encontrado' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  onEdit — actualiza stock al cambiar estado de un pedido
// ════════════════════════════════════════════════════════════
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'Pedidos') return;
  if (e.range.getColumn() !== 11) return;

  const nuevoEstado = e.value;
  if (nuevoEstado !== 'entregado' && nuevoEstado !== 'cancelado') return;

  const idPedido   = sheet.getRange(e.range.getRow(), 1).getValue();
  const barrio     = sheet.getRange(e.range.getRow(), 4).getValue(); // col 4 = Barrio
  const esClub     = String(barrio).startsWith('Club-'); // Clubes no toca stock
  const hDetalle   = SS.getSheetByName('Detalle_Pedidos');
  const hProductos = SS.getSheetByName('Productos');
  const detalleData = hDetalle.getDataRange().getValues();
  const prodData    = hProductos.getDataRange().getValues();

  if (!esClub) {
    detalleData.slice(1).forEach(row => {
      if (row[0] !== idPedido) return;
      const idProd = row[1];
      const qty    = row[3];

      for (let r = 1; r < prodData.length; r++) {
        if (prodData[r][0] === idProd) {
          const celdaRes = hProductos.getRange(r + 1, 4);
          const celdaFis = hProductos.getRange(r + 1, 3);
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

  // Cancelado: eliminar filas del Detalle (de abajo hacia arriba para no correr índices)
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
  SS.toast('✅  Sheets de Maleu configurados correctamente', 'Setup completo', 5);
}

// ── Hoja: Pedidos ────────────────────────────────────────────
function _setupPedidos() {
  let sh = SS.getSheetByName('Pedidos');
  if (!sh) sh = SS.insertSheet('Pedidos');

  const headers = ['ID Pedido','Fecha','Nombre','Barrio','Lote','Teléfono',
                   'Día entrega','Horario','Pago','Total','Estado','Fecha solo'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(BROWN).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.setFrozenRows(1);
  sh.setRowHeight(1, 36);

  [130,150,170,140,70,130,110,90,120,90,110,0].forEach((w, i) => {
    if (w > 0) sh.setColumnWidth(i + 1, w);
  });
  sh.hideColumns(12); // Fecha solo — solo para fórmulas del Panel

  // Dropdown de estado
  const dropRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['pendiente','confirmado','entregado','cancelado'], true)
    .setAllowInvalid(false).build();
  sh.getRange(2, 11, 1000, 1).setDataValidation(dropRule);

  // Formato moneda en Total
  sh.getRange('J2:J1000').setNumberFormat('$#,##0');

  // Formato de fecha
  sh.getRange('B2:B1000').setNumberFormat('@'); // texto

  // Conditional formatting — Estado (col K)
  const kRange = sh.getRange('K2:K1000');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('pendiente')
      .setBackground('#FFF9C4').setFontColor('#7A6000').setBold(true)
      .setRanges([kRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('confirmado')
      .setBackground('#BBDEFB').setFontColor('#0D47A1').setBold(true)
      .setRanges([kRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('entregado')
      .setBackground('#C8E6C9').setFontColor('#1B5E20').setBold(true)
      .setRanges([kRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('cancelado')
      .setBackground('#FFCDD2').setFontColor('#B71C1C').setBold(true)
      .setRanges([kRange]).build(),
  ]);

  // Banding (filas alternas)
  try { sh.getRange('A2:L1000').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false); }
  catch(ex) {}
}

// ── Hoja: Productos ──────────────────────────────────────────
function _setupProductos() {
  let sh = SS.getSheetByName('Productos');
  if (!sh) return;

  sh.getRange(1, 1, 1, 6).setValues([['ID','Nombre','Stock Físico','Reservado','Stock Disponible','Precio']])
    .setBackground(BROWN).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.setFrozenRows(1);
  sh.setRowHeight(1, 36);
  [50, 230, 100, 90, 140, 90].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  sh.getRange('F2:F100').setNumberFormat('$#,##0');
  sh.getRange('C2:E100').setHorizontalAlignment('center');

  // Conditional formatting — Stock Disponible (col E)
  const eRange = sh.getRange('E2:E100');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setBackground('#FFCDD2').setFontColor('#B71C1C').setBold(true)
      .setRanges([eRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(1, 5)
      .setBackground('#FFE0B2').setFontColor('#E65100').setBold(true)
      .setRanges([eRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(5)
      .setBackground('#C8E6C9').setFontColor('#1B5E20')
      .setRanges([eRange]).build(),
  ]);
}

// ── Hoja: Detalle_Pedidos ────────────────────────────────────
function _setupDetalle() {
  let sh = SS.getSheetByName('Detalle_Pedidos');
  if (!sh) sh = SS.insertSheet('Detalle_Pedidos');

  sh.getRange(1, 1, 1, 6).setValues([['ID Pedido','ID Producto','Producto','Cantidad','Precio unit.','Subtotal']])
    .setBackground(BROWN).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.setFrozenRows(1);
  sh.setRowHeight(1, 36);
  [130, 90, 230, 80, 110, 110].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  sh.getRange('E2:F1000').setNumberFormat('$#,##0');
  sh.getRange('D2:D1000').setHorizontalAlignment('center');

  try { sh.getRange('A2:F1000').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false); }
  catch(ex) {}
}

// ── Hoja: Panel ──────────────────────────────────────────────
function _setupPanel() {
  const existing = SS.getSheetByName('Panel');
  if (existing) SS.deleteSheet(existing);
  const sh = SS.insertSheet('Panel', 0);
  sh.setTabColor(ORANGE);
  sh.setHiddenGridlines(true);

  // Fondo general
  sh.getRange('A1:H60').setBackground('#F7F3EE');

  // Columnas
  sh.setColumnWidth(1, 24);
  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 165);
  sh.setColumnWidth(4, 165);
  sh.setColumnWidth(5, 165);
  sh.setColumnWidth(6, 165);
  sh.setColumnWidth(7, 24);

  // ── TÍTULO ──────────────────────────────────────────────
  sh.setRowHeight(1, 16);
  sh.setRowHeight(2, 64);
  sh.getRange('B2:F2').merge()
    .setValue('🍕  MALEU — Panel de Operaciones')
    .setBackground(BROWN).setFontColor('#FFFFFF')
    .setFontSize(18).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(3, 16);

  // ── SECCIÓN: ESTADO DE PEDIDOS ───────────────────────────
  sh.setRowHeight(4, 28);
  sh.getRange('B4').setValue('ESTADO DE PEDIDOS')
    .setFontSize(9).setFontWeight('bold').setFontColor(ORANGE).setFontStyle('normal');

  sh.setRowHeight(5, 28);
  sh.setRowHeight(6, 72);

  const estadoData = [
    ['Pendientes', 'Confirmados', 'Entregados', 'Cancelados'],
  ];
  sh.getRange('C5:F5').setValues(estadoData)
    .setFontSize(10).setFontColor('#666666').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.getRange('C6').setFormula('=IFERROR(COUNTIF(Pedidos!K:K,"pendiente"),0)')
    .setFontSize(30).setFontWeight('bold').setFontColor('#7A6000')
    .setBackground('#FFFDE7').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange('D6').setFormula('=IFERROR(COUNTIF(Pedidos!K:K,"confirmado"),0)')
    .setFontSize(30).setFontWeight('bold').setFontColor('#0D47A1')
    .setBackground('#E3F2FD').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange('E6').setFormula('=IFERROR(COUNTIF(Pedidos!K:K,"entregado"),0)')
    .setFontSize(30).setFontWeight('bold').setFontColor('#1B5E20')
    .setBackground('#E8F5E9').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange('F6').setFormula('=IFERROR(COUNTIF(Pedidos!K:K,"cancelado"),0)')
    .setFontSize(30).setFontWeight('bold').setFontColor('#B71C1C')
    .setBackground('#FFEBEE').setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.getRange('C5:C6').setBorder(null,null,null,null,null,null);
  sh.setRowHeight(7, 16);

  // ── SECCIÓN: INGRESOS ────────────────────────────────────
  sh.setRowHeight(8, 28);
  sh.getRange('B8').setValue('INGRESOS')
    .setFontSize(9).setFontWeight('bold').setFontColor(ORANGE);

  sh.setRowHeight(9, 28);
  sh.setRowHeight(10, 64);

  sh.getRange('C9:D9').merge().setValue('Total facturado (entregados)')
    .setFontSize(10).setFontColor('#666666').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange('E9:F9').merge().setValue('Pendiente de cobro')
    .setFontSize(10).setFontColor('#666666').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.getRange('C10:D10').merge()
    .setFormula('=IFERROR(SUMIF(Pedidos!K:K,"entregado",Pedidos!J:J),0)')
    .setNumberFormat('$#,##0')
    .setFontSize(24).setFontWeight('bold').setFontColor('#1B5E20')
    .setBackground('#E8F5E9').setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.getRange('E10:F10').merge()
    .setFormula('=IFERROR(SUMIFS(Pedidos!J:J,Pedidos!K:K,"pendiente")+SUMIFS(Pedidos!J:J,Pedidos!K:K,"confirmado"),0)')
    .setNumberFormat('$#,##0')
    .setFontSize(24).setFontWeight('bold').setFontColor('#0D47A1')
    .setBackground('#E3F2FD').setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.setRowHeight(11, 16);

  // ── SECCIÓN: PEDIDOS POR DÍA ─────────────────────────────
  sh.setRowHeight(12, 28);
  sh.getRange('B12').setValue('PEDIDOS ACTIVOS POR DÍA DE ENTREGA')
    .setFontSize(9).setFontWeight('bold').setFontColor(ORANGE);

  sh.setRowHeight(13, 28);
  sh.setRowHeight(14, 64);

  const dias = ['Miércoles','Viernes','Sábado','Domingo'];
  sh.getRange('C13:F13').setValues([dias])
    .setFontSize(10).setFontColor('#666666').setFontWeight('bold')
    .setHorizontalAlignment('center').setBackground(CREAM);

  dias.forEach((dia, i) => {
    sh.getRange(14, 3 + i)
      .setFormula(`=IFERROR(COUNTIFS(Pedidos!G:G,"${dia}",Pedidos!K:K,"pendiente")+COUNTIFS(Pedidos!G:G,"${dia}",Pedidos!K:K,"confirmado"),0)`)
      .setFontSize(28).setFontWeight('bold').setFontColor(BROWN)
      .setBackground('#FFF8F0').setHorizontalAlignment('center').setVerticalAlignment('middle');
  });

  sh.setRowHeight(15, 16);

  // ── SECCIÓN: ALERTA DE STOCK ─────────────────────────────
  sh.setRowHeight(16, 28);
  sh.getRange('B16').setValue('⚠️  ALERTAS DE STOCK  (≤ 5 unidades disponibles)')
    .setFontSize(9).setFontWeight('bold').setFontColor('#B71C1C');

  sh.setRowHeight(17, 30);
  sh.getRange('C17:D17').setValues([['Producto','Stock disponible']])
    .setBackground(BROWN).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');

  sh.getRange('C18')
    .setFormula("=IFERROR(QUERY(Productos!B:E,\"SELECT B, E WHERE E <= 5 AND E IS NOT NULL ORDER BY E ASC\",0),\"Sin stock critico\")");
  sh.getRange('C18:C35').setFontSize(10).setFontColor(BROWN);
  sh.getRange('D18:D35').setFontSize(10).setFontColor('#E65100').setFontWeight('bold').setHorizontalAlignment('center');

  // ── PIE ──────────────────────────────────────────────────
  sh.setRowHeight(38, 24);
  sh.getRange('B38')
    .setFormula('="Actualizado: "&TEXT(NOW(),"dd/mm/yyyy HH:mm")')
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
