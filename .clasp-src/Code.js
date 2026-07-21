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

const DIAS_SEMANA_ES = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];

/** Parsea un valor "Día de Entrega Elegido" a Date local (Date o null).
 *  Acepta: Date, "yyyy-MM-dd", "dd/MM/yyyy", "dd/MM/yy". */
function _parseDiaEntregaDate(v) {
  if (v instanceof Date) return v;
  var s = String(v == null ? '' : v).trim();
  if (!s) return null;
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) { var y = Number(m[3]); if (y < 100) y += 2000; return new Date(y, Number(m[2]) - 1, Number(m[1])); }
  return null;
}
/** Convierte un valor de celda "Día de Entrega" al nombre del día en español. */
function _fechaADiaSemana(v) {
  var d = _parseDiaEntregaDate(v);
  if (d) return DIAS_SEMANA_ES[d.getDay()];
  return String(v == null ? '' : v).trim();
}
/** Devuelve fecha ISO "yyyy-MM-dd" del valor o '' si no se puede parsear. */
function _fechaAISO(v) {
  var d = _parseDiaEntregaDate(v);
  if (!d) return '';
  var y = d.getFullYear(), m = String(d.getMonth() + 1).padStart(2, '0'), dd = String(d.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + dd;
}

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
  const sheetMap = { 'H-': ['Home', 'B'], 'C-': ['Clubes', 'B'], 'OC-': ['Orden de Compra', 'A'], 'P-': ['Pilar', 'B'], 'R-': ['Red', 'B'] };
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

  // Comparar con Config para nunca retroceder
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
 * @param {string} canal - 'Home', 'Clubes', 'OC', 'Deposito', 'Manual'
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
  // Fix bug #NAME?: forzar columna D (Tipo) como texto plano
  // para que Sheets no interprete el '+' de '+REC' como fórmula.
  var lastRow = sh.getLastRow();
  sh.getRange(lastRow, 4).setNumberFormat('@').setValue(tipo);
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
//  doGet — lectura de datos (compras)
// ════════════════════════════════════════════════════════════
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'compras') return _doGetCompras();
  if (action === 'egresos') return _doGetEgresos();
  if (action === 'entregas') return _doGetEntregas(e);
  if (action === 'vendedores') return _doGetVendedores();
  if (action === 'dashboardVendedor') return _doGetDashboardVendedor(e);
  if (action === 'resolverVendedor') return _doGetResolverVendedor(e);
  if (action === 'busqueda') return _doGetBusqueda();
  if (action === 'catalogo') return _doGetCatalogo();
  if (action === 'admin') return _doGetAdmin();
  if (action === 'pedidosLight') return _doGetAdmin({ light: true });
  if (action === 'pedidosNew') return _doGetAdmin({ light: true, tail: Math.max(3, Math.min(50, parseInt(e.parameter.tail) || 10)) });
  if (action === 'cajaLight') return _doGetAdmin({ mode: 'caja' });
  if (action === 'ventas') return _doGetVentas();
  if (action === 'stock') return _doGetStock();
  if (action === 'stock_full') return _doGetStockFull();
  if (action === 'precios') return _doGetPrecios();
  if (action === 'validarCupon') return _doGetValidarCupon(e);
  if (action === 'cobrosPendientes') return _doGetCobrosPendientes(e);
  if (action === 'billetera') return _doGetBilletera();
  if (action === 'pendientesGuardarStock') return _doGetPendientesGuardarStock();
  if (action === 'saldoCliente') return _doGetSaldoCliente(e);
  if (action === 'resumenSemanal') return _doGetResumenSemanal(e);
  if (action === 'crmClientes') return _doGetCrmClientes(e);
  if (action === 'crmCliente') return _doGetCrmCliente(e);
  if (action === 'crmProductos') return _doGetCrmProductos(e);
  if (action === 'crmProducto') return _doGetCrmProducto(e);
  if (action === 'crmZonas') return _doGetCrmZonas(e);
  if (action === 'crmLotes') return _doGetCrmLotes(e);
  if (action === 'hogaresMerge') return _doGetHogaresMerge(e);
  if (action === 'crmPuntos') return _doGetCrmPuntos(e);
  if (action === 'crmInteracciones') return _doGetCrmInteracciones(e);
  if (action === 'analisisProveedor') return _doGetAnalisisProveedor(e);
  if (action === 'productosAnalytics') return _doGetProductosAnalytics(e);
  if (action === 'combosEval') return _doGetCombosEval(e);
  if (action === 'miReparto') return _doGetMiReparto(e);
  if (action === 'miRepartoHistorial') return _doGetMiRepartoHistorial(e);
  if (action === 'cierreMensual') return _doGetCierreMensual(e);
  if (action === 'planMes') return _doGetPlanMes(e);
  if (action === 'repartidoresList') return _doGetRepartidoresList();
  if (action === 'ajustesData') return _doGetAjustesData();
  return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
}

// ════════════════════════════════════════════════════════════
//  Mi Reparto - Historial: días trabajados del repartidor.
//  Params: usuario (obligatorio), desde/hasta (ISO, opcionales)
//  Devuelve por día: cantidad entregas, propinas, fijo $50k, total.
//  Solo días con al menos 1 entrega. Ordenado descendente.
// ════════════════════════════════════════════════════════════
function _doGetMiRepartoHistorial(e) {
  var usuario = String((e && e.parameter && e.parameter.usuario) || '').trim();
  if (!usuario) {
    return ContentService.createTextOutput(JSON.stringify({ok:false, err:'Falta usuario'})).setMimeType(ContentService.MimeType.JSON);
  }
  var desde = String((e && e.parameter && e.parameter.desde) || '').trim();
  var hasta = String((e && e.parameter && e.parameter.hasta) || '').trim();
  var FIJO_DIA = 50000;

  var sh = SS.getSheetByName('Home');
  if (!sh || sh.getLastRow() <= 1) {
    return ContentService.createTextOutput(JSON.stringify({
      ok:true, usuario:usuario, fijo:FIJO_DIA,
      diasTrabajados:0, entregasTot:0, propinasTot:0, totalGanado:0, dias:[]
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = {};
  for (var h = 0; h < headers.length; h++) idx[String(headers[h]).trim()] = h;

  var iEstadoEnt = idx['Estado de Entrega'] != null ? idx['Estado de Entrega'] : 10;
  var iEstadoPg  = idx['Estado de Pago'] != null ? idx['Estado de Pago'] : 12;
  var iEf        = idx['Efectivo'] != null ? idx['Efectivo'] : 17;
  var iTr        = idx['Transferencia'] != null ? idx['Transferencia'] : 18;
  var iPropEf    = idx['Propina Efectivo'] != null ? idx['Propina Efectivo'] : 19;
  var iPropTr    = idx['Propina Transferencia'] != null ? idx['Propina Transferencia'] : 20;
  var iFact      = idx['Facturado'] != null ? idx['Facturado'] : 21;
  var iRep       = idx['Repartidor'];
  var iFEnt = -1;
  for (var hf = 0; hf < headers.length; hf++) {
    var hn = String(headers[hf]).trim();
    if (hn === 'Fecha Entrega' || hn === 'Fecha de Entrega') { iFEnt = hf; break; }
  }
  if (iRep == null) {
    return ContentService.createTextOutput(JSON.stringify({ok:false, err:'Columna Repartidor no encontrada'})).setMimeType(ContentService.MimeType.JSON);
  }

  var data = sh.getDataRange().getValues();
  var usuarioLow = usuario.toLowerCase();
  var porDia = {}; // fecha ISO → metrics

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[iEstadoEnt] || '').trim() !== 'Entregado') continue;
    var rep = String(row[iRep] || '').trim();
    if (!rep || rep.toLowerCase() !== usuarioLow) continue;

    var fRaw = (iFEnt >= 0) ? row[iFEnt] : null;
    var fISO = '';
    if (fRaw instanceof Date) {
      fISO = Utilities.formatDate(fRaw, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
    } else if (fRaw) {
      var s = String(fRaw).trim().split(' ')[0];
      var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (m) {
        fISO = m[3] + '-' + (m[2].length===1?'0'+m[2]:m[2]) + '-' + (m[1].length===1?'0'+m[1]:m[1]);
      } else if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
        fISO = s.substring(0,10);
      }
    }
    if (!fISO) continue;
    if (desde && fISO < desde) continue;
    if (hasta && fISO > hasta) continue;

    if (!porDia[fISO]) porDia[fISO] = { entregas:0, propinas:0, ef:0, tr:0, propEf:0, propTr:0, movido:0, pendiente:0, nPend:0 };
    var d = porDia[fISO];
    d.entregas++;
    var ef = Number(row[iEf]) || 0;
    var tr = Number(row[iTr]) || 0;
    var pEf = Math.max(0, Number(row[iPropEf]) || 0);
    var pTr = Math.max(0, Number(row[iPropTr]) || 0);
    var fact = Number(row[iFact]) || 0;
    d.ef += ef;
    d.tr += tr;
    d.propEf += pEf;
    d.propTr += pTr;
    d.propinas += pEf + pTr;
    d.movido += fact;
    var estadoPago = String(row[iEstadoPg] || '').trim();
    if (estadoPago !== 'Cobrado') { d.pendiente += fact; d.nPend++; }
  }

  var dias = [];
  var entregasTot = 0, propinasTot = 0, movidoTot = 0, pendienteTot = 0;
  Object.keys(porDia).forEach(function(f) {
    var d = porDia[f];
    var ganadoDia = FIJO_DIA + d.propinas;
    dias.push({
      f:f, n:d.entregas, prop:d.propinas, fijo:FIJO_DIA, tot:ganadoDia,
      ef:d.ef, tr:d.tr, propEf:d.propEf, propTr:d.propTr,
      movido:d.movido, pendiente:d.pendiente, nPend:d.nPend
    });
    entregasTot += d.entregas;
    propinasTot += d.propinas;
    movidoTot += d.movido;
    pendienteTot += d.pendiente;
  });
  dias.sort(function(a,b){ return a.f < b.f ? 1 : -1; });

  var diasTrabajados = dias.length;
  var totalGanado = diasTrabajados * FIJO_DIA + propinasTot;

  return ContentService.createTextOutput(JSON.stringify({
    ok:true, usuario:usuario, fijo:FIJO_DIA,
    diasTrabajados:diasTrabajados, entregasTot:entregasTot,
    propinasTot:propinasTot, totalGanado:totalGanado,
    movidoTot:movidoTot, pendienteTot:pendienteTot,
    dias:dias
  })).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  POST cerrarDiaReparto — el repartidor cierra el día.
//  Body: { usuario, fecha (ISO), billeteraIni, cobradoEf, esperado, real, notas? }
//  Escribe fila en hoja "Cierres Reparto". Si para esa combinación (usuario, fecha)
//  ya hay cierre previo, lo sobrescribe (no duplica).
// ════════════════════════════════════════════════════════════
function _doPostCerrarDiaReparto(data) {
  var usuario = String(data.usuario || '').trim();
  var fechaISO = String(data.fecha || '').trim();
  if (!usuario || !fechaISO) {
    return ContentService.createTextOutput(JSON.stringify({ok:false, err:'Faltan usuario o fecha'})).setMimeType(ContentService.MimeType.JSON);
  }
  var billeteraIni = Number(data.billeteraIni) || 0;
  var cobradoEf    = Number(data.cobradoEf) || 0;
  var esperado     = Number(data.esperado) || 0;
  var real         = Number(data.real) || 0;
  var notas        = String(data.notas || '').trim();
  var diferencia   = real - esperado;

  var sh = SS.getSheetByName('Cierres Reparto');
  if (!sh) {
    sh = SS.insertSheet('Cierres Reparto');
    sh.getRange(1, 1, 1, 9).setValues([[
      'Fecha Cierre','Usuario','Fecha Reparto','Billetera Inicial','Cobrado Efectivo',
      'Esperado','Real','Diferencia','Notas'
    ]]);
    sh.setFrozenRows(1);
    sh.getRange('A1:I1').setFontWeight('bold').setBackground('#331C1C').setFontColor('#fff');
    sh.setColumnWidths(1, 9, 130);
  }

  var ahora = new Date();
  var fechaRepDate = (function(){
    var p = fechaISO.split('-');
    return new Date(Number(p[0]), Number(p[1])-1, Number(p[2]));
  })();

  // Upsert: sobreescribir cierre previo del mismo (usuario, fecha) en vez de duplicar.
  var rowToWrite = -1;
  if (sh.getLastRow() > 1) {
    var vals = sh.getRange(2, 1, sh.getLastRow()-1, 3).getValues();
    for (var i = 0; i < vals.length; i++) {
      var u = String(vals[i][1] || '').trim().toLowerCase();
      var fr = vals[i][2];
      var frISO = '';
      if (fr instanceof Date) {
        frISO = Utilities.formatDate(fr, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
      } else if (fr) {
        var sFR = String(fr).trim().split(' ')[0];
        var mFR = sFR.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (mFR) frISO = mFR[3]+'-'+(mFR[2].length===1?'0'+mFR[2]:mFR[2])+'-'+(mFR[1].length===1?'0'+mFR[1]:mFR[1]);
      }
      if (u === usuario.toLowerCase() && frISO === fechaISO) {
        rowToWrite = i + 2;
        break;
      }
    }
  }

  var rowData = [[ahora, usuario, fechaRepDate, billeteraIni, cobradoEf, esperado, real, diferencia, notas]];
  if (rowToWrite > 0) {
    sh.getRange(rowToWrite, 1, 1, 9).setValues(rowData);
  } else {
    sh.appendRow(rowData[0]);
  }

  return ContentService.createTextOutput(JSON.stringify({
    ok: true, usuario: usuario, fecha: fechaISO,
    billeteraIni: billeteraIni, cobradoEf: cobradoEf,
    esperado: esperado, real: real, diferencia: diferencia
  })).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  Mi Reparto — resumen del día para rol Repartidor.
//  Params: usuario (nombre exacto col Repartidor), fecha (ISO YYYY-MM-DD).
//  Solo Home (los Pilar/Clubes/Red no van con repartidor delegado por ahora).
//  Devuelve: entregas del día con propinas + fijo $50.000/día.
// ════════════════════════════════════════════════════════════
function _doGetMiReparto(e) {
  var usuario = String((e && e.parameter && e.parameter.usuario) || '').trim();
  var fechaISO = String((e && e.parameter && e.parameter.fecha) || '').trim();
  if (!usuario) {
    return ContentService.createTextOutput(JSON.stringify({ok:false, err:'Falta usuario'})).setMimeType(ContentService.MimeType.JSON);
  }
  if (!fechaISO || !/^\d{4}-\d{2}-\d{2}$/.test(fechaISO)) {
    return ContentService.createTextOutput(JSON.stringify({ok:false, err:'Fecha inválida (YYYY-MM-DD)'})).setMimeType(ContentService.MimeType.JSON);
  }

  var FIJO_DIA = 50000;

  // ── Saldo billetera (caja Maleu) ── desde último snapshot Saldo Base col D.
  // Tadeo lo edita en Panel→Caja→Ajustar. El repartidor ve este valor al refrescar.
  var saldoBilletera = 0;
  var shSB = SS.getSheetByName('Saldo Base');
  if (shSB && shSB.getLastRow() > 1 && shSB.getLastColumn() >= 4) {
    saldoBilletera = Number(shSB.getRange(shSB.getLastRow(), 4).getValue()) || 0;
  }

  // ── Reconstruir la billetera INICIAL del día + ajustes de efectivo en mano ──
  // El arqueo necesita la billetera con la que ARRANCASTE el día, no la actual: cada
  // vuelto de billetera bajó la col D, así que lo sumamos de vuelta. Además los cambios
  // cruzados mueven efectivo en mano sin pasar por el cobro (CruzadoEf resta, CambioMP suma).
  // Fuente: hoja "Cambios Billetera" (Fecha|Hoja|Pedido|Cliente|Monto|Repartidor|Tipo),
  // filtrada por la fecha de reparto y el repartidor (filas viejas sin repartidor cuentan).
  var vueltosBilletera = 0; // Σ vuelto de billetera de hoy → reconstruye la billetera inicial
  var cambioMPTot = 0;      // Σ CambioMP: efectivo extra que te quedó (devolviste por MP)
  var cruzadoEfTot = 0;     // Σ CruzadoEf: efectivo que diste de vuelto por un pago MP
  var usuarioLowBil = usuario.toLowerCase();
  var shCBil = SS.getSheetByName('Cambios Billetera');
  var vueltoPorPed = {}; // { pedidoId: vuelto efectivo dado en ese pedido } para el desglose
  if (shCBil && shCBil.getLastRow() > 1) {
    var cbVals = shCBil.getDataRange().getValues();
    for (var icb = 1; icb < cbVals.length; icb++) {
      var cbF = cbVals[icb][0];
      var cbISO = '';
      if (cbF instanceof Date) {
        cbISO = Utilities.formatDate(cbF, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
      } else if (cbF) {
        var sCB = String(cbF).trim().split(' ')[0];
        var mCB = sCB.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (mCB) cbISO = mCB[3] + '-' + (mCB[2].length === 1 ? '0' + mCB[2] : mCB[2]) + '-' + (mCB[1].length === 1 ? '0' + mCB[1] : mCB[1]);
        else if (/^\d{4}-\d{2}-\d{2}/.test(sCB)) cbISO = sCB.substring(0, 10);
      }
      if (cbISO !== fechaISO) continue;
      var cbRep = String(cbVals[icb][5] || '').trim().toLowerCase(); // col F (puede faltar en filas viejas)
      if (cbRep && cbRep !== usuarioLowBil) continue;
      var cbTipo = String(cbVals[icb][6] || 'Billetera').trim();     // col G
      var cbMonto = Number(cbVals[icb][4]) || 0;                      // col E
      if (cbTipo === 'CruzadoEf') cruzadoEfTot += cbMonto;
      else if (cbTipo === 'CambioMP') cambioMPTot += cbMonto;
      else vueltosBilletera += cbMonto; // 'Billetera' (default)
      // Vuelto en efectivo por pedido (billetera + cruzado; CambioMP es vuelto por MP, no efectivo).
      if (cbTipo !== 'CambioMP') {
        var cbPed = String(cbVals[icb][2] || '').trim();
        if (cbPed) vueltoPorPed[cbPed] = (vueltoPorPed[cbPed] || 0) + cbMonto;
      }
    }
  }
  var ajusteCruzado = cambioMPTot - cruzadoEfTot; // efectivo en mano que no quedó en el cobro
  var billeteraIni = saldoBilletera + vueltosBilletera;
  // Vuelto efectivo total dado en mano (de billetera + cruzados por pago MP).
  var vueltoEfDia = vueltosBilletera + cruzadoEfTot;

  // ── Cierre del día (si ya cerró). Hoja "Cierres Reparto": Fecha cierre, Usuario, Fecha reparto, Billetera ini, Cobrado Ef, Esperado, Real, Diferencia, Notas.
  var cierreDia = null;
  var shCR = SS.getSheetByName('Cierres Reparto');
  if (shCR && shCR.getLastRow() > 1) {
    var crVals = shCR.getRange(2, 1, shCR.getLastRow()-1, 9).getValues();
    for (var ic = crVals.length - 1; ic >= 0; ic--) {
      var uCR = String(crVals[ic][1] || '').trim().toLowerCase();
      var fRep = crVals[ic][2];
      var fRepISO = '';
      if (fRep instanceof Date) {
        fRepISO = Utilities.formatDate(fRep, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
      } else if (fRep) {
        var sFR = String(fRep).trim().split(' ')[0];
        var mFR = sFR.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (mFR) fRepISO = mFR[3]+'-'+(mFR[2].length===1?'0'+mFR[2]:mFR[2])+'-'+(mFR[1].length===1?'0'+mFR[1]:mFR[1]);
        else if (/^\d{4}-\d{2}-\d{2}/.test(sFR)) fRepISO = sFR.substring(0,10);
      }
      if (uCR === usuario.toLowerCase() && fRepISO === fechaISO) {
        cierreDia = {
          billeteraIni: Number(crVals[ic][3]) || 0,
          cobradoEf:    Number(crVals[ic][4]) || 0,
          esperado:     Number(crVals[ic][5]) || 0,
          real:         Number(crVals[ic][6]) || 0,
          diferencia:   Number(crVals[ic][7]) || 0,
          notas:        String(crVals[ic][8] || '')
        };
        break;
      }
    }
  }

  var sh = SS.getSheetByName('Home');
  if (!sh || sh.getLastRow() <= 1) {
    return ContentService.createTextOutput(JSON.stringify({
      ok:true, usuario:usuario, fecha:fechaISO, fijo:FIJO_DIA, propinas:0, total:FIJO_DIA,
      saldoBilletera:saldoBilletera, billeteraIni:billeteraIni, vueltosBilletera:vueltosBilletera,
      ajusteCruzado:ajusteCruzado, recibidoEf:cambioMPTot+vueltosBilletera, vueltoEf:vueltoEfDia,
      cobradoEfHoy:0, cobradoTrHoy:0,
      plataEnMano:billeteraIni+ajusteCruzado, cierreDia:cierreDia, entregas:[]
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = {};
  for (var h = 0; h < headers.length; h++) {
    idx[String(headers[h]).trim()] = h;
  }

  var iN          = idx['N° Pedido'] != null ? idx['N° Pedido'] : 1;
  var iCliente    = idx['Cliente'] != null ? idx['Cliente'] : 7;
  var iEstadoEnt  = idx['Estado de Entrega'] != null ? idx['Estado de Entrega'] : 10;
  var iEstadoPago = idx['Estado de Pago'] != null ? idx['Estado de Pago'] : 12;
  var iTotal      = idx['Total ($)'] != null ? idx['Total ($)'] : (idx['Total'] != null ? idx['Total'] : 13);
  var iEf         = idx['Efectivo'] != null ? idx['Efectivo'] : 15;
  var iTr         = idx['Transferencia'] != null ? idx['Transferencia'] : 16;
  var iPropEf     = idx['Propina Efectivo'] != null ? idx['Propina Efectivo'] : 19;
  var iPropTr     = idx['Propina Transferencia'] != null ? idx['Propina Transferencia'] : 20;
  var iRep        = idx['Repartidor'];
  var iSubBarrio  = idx['Sub Barrio'];
  var iDomLote    = idx['Domicilio-Lote'] != null ? idx['Domicilio-Lote'] : (idx['Domicilio'] != null ? idx['Domicilio'] : -1);
  // Buscar "Fecha Entrega" real (auto-llenada al marcar Entregado). Fallback "Fecha de Entrega".
  var iFEnt = -1;
  for (var hf = 0; hf < headers.length; hf++) {
    var hn = String(headers[hf]).trim();
    if (hn === 'Fecha Entrega' || hn === 'Fecha de Entrega') { iFEnt = hf; break; }
  }

  if (iRep == null) {
    return ContentService.createTextOutput(JSON.stringify({ok:false, err:'Columna Repartidor no encontrada en hoja Home'})).setMimeType(ContentService.MimeType.JSON);
  }

  var data = sh.getDataRange().getValues();
  var usuarioLow = usuario.toLowerCase();
  var entregas = [];
  var sumPropEf = 0, sumPropTr = 0;
  var sumCobEf = 0, sumCobTr = 0;

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var estEnt = String(row[iEstadoEnt] || '').trim();
    if (estEnt !== 'Entregado') continue;
    var rep = String(row[iRep] || '').trim();
    if (!rep || rep.toLowerCase() !== usuarioLow) continue;

    // Comparar fecha entrega real con la pedida (ISO)
    var fRaw = (iFEnt >= 0) ? row[iFEnt] : null;
    var fISO = '';
    if (fRaw instanceof Date) {
      fISO = Utilities.formatDate(fRaw, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
    } else if (fRaw) {
      var s = String(fRaw).trim().split(' ')[0];
      var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (m) {
        fISO = m[3] + '-' + (m[2].length===1?'0'+m[2]:m[2]) + '-' + (m[1].length===1?'0'+m[1]:m[1]);
      } else if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
        fISO = s.substring(0,10);
      }
    }
    if (fISO !== fechaISO) continue;

    var nPed   = String(row[iN] || '').trim();
    var cliente= String(row[iCliente] || '').trim();
    var total  = Number(row[iTotal]) || 0;
    var propEf = Math.max(0, Number(row[iPropEf]) || 0);
    var propTr = Math.max(0, Number(row[iPropTr]) || 0);
    var ef     = Number(row[iEf]) || 0;
    var tr     = Number(row[iTr]) || 0;
    var subBarrio = (iSubBarrio != null) ? String(row[iSubBarrio] || '').trim() : '';
    var lote   = (iDomLote >= 0) ? String(row[iDomLote] || '').trim() : '';
    var estPag = String(row[iEstadoPago] || '').trim();
    var cobrado = (estPag === 'Cobrado');

    sumPropEf += propEf;
    sumPropTr += propTr;
    // Solo cuenta para la billetera lo cobrado en efectivo.
    // Propina ef se incluye aparte porque se suma al cobro físico que recibe el repartidor.
    if (cobrado) { sumCobEf += ef + propEf; sumCobTr += tr + propTr; }

    entregas.push({
      n: nPed, c: cliente, $: total,
      pEf: propEf, pTr: propTr,
      ef: ef, tr: tr,
      vto: vueltoPorPed[nPed] || 0,   // vuelto en efectivo dado en este pedido
      sb: subBarrio, l: lote,
      ep: estPag, cb: cobrado
    });
  }

  var propinasTot = sumPropEf + sumPropTr;
  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    usuario: usuario,
    fecha: fechaISO,
    fijo: FIJO_DIA,
    propinasEf: sumPropEf,
    propinasTr: sumPropTr,
    propinas: propinasTot,
    total: FIJO_DIA + propinasTot,
    cantidad: entregas.length,
    saldoBilletera: saldoBilletera,        // billetera ACTUAL (la usa el modal de cobro)
    billeteraIni: billeteraIni,            // billetera INICIAL del día (la usa el arqueo)
    vueltosBilletera: vueltosBilletera,    // vueltos de billetera dados hoy
    ajusteCruzado: ajusteCruzado,          // +CambioMP −CruzadoEf (efectivo en mano fuera del cobro)
    recibidoEf: sumCobEf + cambioMPTot + vueltosBilletera, // efectivo BRUTO recibido de clientes
    vueltoEf: vueltoEfDia,                 // efectivo total devuelto de vuelto (billetera + cruzado)
    cobradoEfHoy: sumCobEf,
    cobradoTrHoy: sumCobTr,
    plataEnMano: billeteraIni + sumCobEf + ajusteCruzado,
    cierreDia: cierreDia,
    entregas: entregas
  })).setMimeType(ContentService.MimeType.JSON);
}

// Cobros pendientes: pedidos Entregados + No Cobrados de TODAS las hojas
// Red: agrupado por Vendedor (aplica 17% de comision)
//
// Cache server-side 30s (TTL): la PWA Ruta refresca cada 90s y el panel cada
// poco más. Sin cache, cada llamada toma ~7s leyendo 4 hojas. Con cache, el
// 2do click responde instantáneo. Para forzar refresh (post-cobro), el cliente
// pasa &fresh=1 — eso saltea cache y regenera (y guarda el nuevo valor).
function _doGetCobrosPendientes(e) {
  var _ck = 'cobros_pendientes_v1';
  var _fresh = !!(e && e.parameter && e.parameter.fresh);
  if (!_fresh) {
    try {
      var _c = CacheService.getScriptCache();
      var _h = _c.get(_ck);
      if (_h) return ContentService.createTextOutput(_h).setMimeType(ContentService.MimeType.JSON);
    } catch (_e) { /* sin cache no rompe */ }
  }
  // colSub/colEnv/colDesc: necesarios para recalcular el total al cambiar forma de pago
  // (aplicar/quitar 10% OFF Efectivo). Home/Pilar v2: N(13)=Subtotal, O(14)=Envío, P(15)=Descuento.
  // Clubes no tiene descuento efectivo.
  var hojas = [
    {name:'Home', colCliente:7, colEst:10, colPago:12, colTotal:21, colTel:46, colFp:11, colDia:9, colFecha:3, colSub:13, colEnv:14, colDesc:15},
    {name:'Pilar', colCliente:7, colEst:10, colPago:12, colTotal:21, colTel:49, colFp:11, colDia:9, colFecha:3, colSub:13, colEnv:14, colDesc:15},
    {name:'Clubes', colCliente:7, colEst:13, colPago:15, colTotal:16, colTel:null, colFp:14, colDia:12, colFecha:3, colSub:null, colEnv:null, colDesc:null}
  ];
  // ── Cargar parciales acumulados por pedido (clave: hoja|id) ──
  var parcialesPorPed = {}; // { 'Clubes|33': { total: 1170983, items: [{fp,monto,fecha}, ...] } }
  var shCP = SS.getSheetByName('Cobros Parciales');
  if (shCP && shCP.getLastRow() > 1) {
    var cpData = shCP.getDataRange().getValues();
    for (var rcp = 1; rcp < cpData.length; rcp++) {
      var cpHoja = String(cpData[rcp][1] || '').trim();
      var cpId   = String(cpData[rcp][2] || '').trim();
      if (!cpHoja || !cpId) continue;
      var cpFp = String(cpData[rcp][4] || '').trim();
      var cpMonto = Number(cpData[rcp][5]) || 0;
      var cpFecha = cpData[rcp][0];
      var cpFechaStr = (cpFecha instanceof Date)
        ? Utilities.formatDate(cpFecha, 'America/Argentina/Buenos_Aires', 'dd/MM HH:mm')
        : String(cpFecha || '');
      var k = cpHoja + '|' + cpId;
      if (!parcialesPorPed[k]) parcialesPorPed[k] = { total: 0, items: [] };
      parcialesPorPed[k].total += cpMonto;
      parcialesPorPed[k].items.push({ fp: cpFp, monto: cpMonto, fecha: cpFechaStr });
    }
  }

  // Productos del pedido (para mostrar en el cuadro de Cobros: qué pidió el cliente).
  function _prodsDelCobro(row, hojaName) {
    var ABBRS_H = ['PPM','PPJyQ','PPCyQ','SCo','SJyQ','SCa','ECaC','EJyQ','ECyQ','EV','TG','TLC','TC','F','PMu','PMa','PJyQ','PCC','PJyM'];
    var ABBRS_P = ['PPM','PPJyQ','PPCyQ','SQB','SL','SCo','SPyP','SJyQ','SE','SCa','ECaC','EJyQ','ECyQ','EV','TG','TLC','TC','F','PMu','PMa','PJyQ','PCC','PJyM'];
    var ABBRS_C = ['PMu','PMa','PJyQ','PCC','PJyM','PPM','PPJyQ','PPCyQ'];
    var res = [];
    function add(a, q) { if (q > 0) res.push({ a: a, q: q }); }
    if (hojaName === 'Home' || hojaName === 'Pilar') {
      var abbrs = (hojaName === 'Pilar') ? ABBRS_P : ABBRS_H;
      for (var p = 0; p < abbrs.length; p++) add(abbrs[p], Number(row[22 + p]) || 0);
      var tStart = (hojaName === 'Pilar') ? 59 : 56;
      var tAbbrs = ['TP','TJyQ','TCa','TV'];
      for (var t = 0; t < 4; t++) add(tAbbrs[t], Number(row[tStart + t]) || 0);
      var wStart = (hojaName === 'Pilar') ? 64 : 61;
      var wAbbrs = ['RC','RP'];
      for (var w = 0; w < 2; w++) add(wAbbrs[w], Number(row[wStart + w]) || 0);
    } else if (hojaName === 'Clubes') {
      for (var c = 0; c < ABBRS_C.length; c++) add(ABBRS_C[c], Number(row[23 + c]) || 0);
      var eAbbrs = ['ECaC','EJyQ','ECyQ','EV'];
      for (var e = 0; e < 4; e++) add(eAbbrs[e], Number(row[37 + e]) || 0);
    }
    return res;
  }
  var out = [];
  hojas.forEach(function(cfg) {
    var sh = SS.getSheetByName(cfg.name);
    if (!sh || sh.getLastRow() <= 1) return;
    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var estado = String(data[r][cfg.colEst] || '').trim();
      var pago = String(data[r][cfg.colPago] || '').trim();
      if (estado === 'Cancelado') continue;
      if (pago === 'Cobrado') continue;
      var id = data[r][1];
      if (!id || String(id).trim() === '-') continue;
      var DIAS_COB = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
      // Fecha del pedido (col D) — solo para fallback si no hay fecha entrega
      var fecha = data[r][cfg.colFecha];
      var fechaPedStr = '';
      if (fecha instanceof Date) fechaPedStr = Utilities.formatDate(fecha, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
      else fechaPedStr = String(fecha || '').trim();
      // Día de Entrega Elegido (col J Home/Pilar, col M Clubes) — fuente de orden
      // Puede venir como Date, como texto "dd/MM/yyyy", o como nombre del día (pedidos viejos)
      var diaRawCob = data[r][cfg.colDia];
      var diaCob = '', feEntStr = '';
      if (diaRawCob instanceof Date) {
        diaCob = DIAS_COB[diaRawCob.getDay()];
        feEntStr = Utilities.formatDate(diaRawCob, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
      } else {
        var raw = String(diaRawCob || '').trim();
        var mCob = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (mCob) {
          var dCob = new Date(Number(mCob[3]), Number(mCob[2]) - 1, Number(mCob[1]));
          diaCob = DIAS_COB[dCob.getDay()];
          feEntStr = raw;
        } else {
          // Solo tenemos nombre del día (ej. "Viernes") — dejamos feEntStr vacío
          diaCob = raw;
        }
      }
      // fe = fecha usada para ORDENAR/AGRUPAR en el frontend. Preferencia: día de
      // entrega elegido; fallback: fecha de pedido (pedidos viejos sin col J).
      var feSort = feEntStr || fechaPedStr;
      var totalPed = Number(data[r][cfg.colTotal]) || 0;
      var keyPed = cfg.name + '|' + id;
      var parc = parcialesPorPed[keyPed];
      var parcialTotal = parc ? parc.total : 0;
      var restantePed = parcialTotal > 0 ? Math.max(0, Math.round(totalPed - parcialTotal)) : totalPed;
      // Si los parciales ya cubren el total, el pedido debería estar Cobrado.
      // Lo dejamos pasar al frontend para que se cierre vía cobrar (defensivo).
      if (restantePed <= 0) continue;

      out.push({
        key: cfg.name + '|' + id,
        h: cfg.name,
        id: String(id),
        r: r + 1,
        c: String(data[r][cfg.colCliente] || '').trim(),
        t: cfg.colTel !== null ? String(data[r][cfg.colTel] || '').trim() : '',
        '$': restantePed,                 // saldo pendiente real
        totalOriginal: totalPed,           // total facturado del pedido
        cobradoParcial: parcialTotal,      // ya cobrado en parciales
        parciales: parc ? parc.items : [], // detalle
        fp: String(data[r][cfg.colFp] || '').trim(),
        de: diaCob,
        fe: feSort,
        fePed: fechaPedStr,
        feEnt: feEntStr,
        es: estado,
        // Datos para recalcular total si cambia la forma de pago (aplicar/quitar 10% Efectivo)
        sub: cfg.colSub !== null ? (Number(data[r][cfg.colSub]) || 0) : 0,
        env: cfg.colEnv !== null ? (Number(data[r][cfg.colEnv]) || 0) : 0,
        desc: cfg.colDesc !== null ? (Number(data[r][cfg.colDesc]) || 0) : 0,
        p: _prodsDelCobro(data[r], cfg.name)   // productos del pedido (para el cuadro de Cobros)
      });
    }
  });

  // Red: agrupar por Vendedor. Lo que Tadeo le cobra al vendedor es lo que el
  // vendedor ya cobró (o adeuda) al cliente — el filtro correcto es por
  // "Estado Pago a Maleu" (BB, idx 53), NO por "Estado de Pago" al cliente
  // (col N, idx 13). Un pedido con Cliente=Cobrado pero Maleu=Pendiente
  // significa que Marcos cobró pero todavía no rindió a Maleu, y debe
  // aparecer en cobros pendientes igual.
  // Comisión: leer de hoja Vendedores por nombre (no hardcodear 17%).
  var comisionPctVendedor = {};
  var shVendCo = SS.getSheetByName('Vendedores');
  if (shVendCo && shVendCo.getLastRow() > 1) {
    var dV = shVendCo.getDataRange().getValues();
    for (var rv2 = 1; rv2 < dV.length; rv2++) {
      var nm2 = String(dV[rv2][0] || '').trim();
      if (!nm2) continue;
      comisionPctVendedor[nm2] = Number(dV[rv2][9]) || 17;  // col J = Comisión %
    }
  }
  var shRed = SS.getSheetByName('Red');
  if (shRed && shRed.getLastRow() > 1) {
    var dRed = shRed.getDataRange().getValues();
    // Ubicar columnas por header (Red puede crecer en columnas)
    var hdrRed = dRed[0];
    var idxEstPagoMaleu = -1, idxFpMaleu = -1, idxEntVend = -1;
    for (var hh = 0; hh < hdrRed.length; hh++) {
      var nm = String(hdrRed[hh]).trim();
      if (nm === 'Estado Pago a Maleu') idxEstPagoMaleu = hh;
      else if (nm === 'Forma Pago a Maleu') idxFpMaleu = hh;
      else if (nm === 'Entregado a Vendedor') idxEntVend = hh;
    }
    // Index de "A Pagar" (col real en Sheets): es lo que el vendedor le debe a
    // Maleu por pedido — ya tiene aplicada la comisión que corresponda. Se prefiere
    // esto al cálculo 17%-plano porque algunos pedidos (familia del vendedor,
    // ajustes manuales) tienen comisión 0 hardcodeada. Caso 30/06/2026: Marcos
    // R-45 y R-46 a Juan Cruz Bottcher con comisión $0 — el Panel cobraba $11.900
    // de más al aplicar 17% al bruto total.
    var idxAPagar = -1;
    for (var hhA = 0; hhA < hdrRed.length; hhA++) {
      if (String(hdrRed[hhA]).trim() === 'A Pagar') { idxAPagar = hhA; break; }
    }
    var porVendedor = {};
    for (var r = 1; r < dRed.length; r++) {
      var estado = String(dRed[r][11] || '').trim(); // col 12 Estado Entrega
      if (estado === 'Cancelado') continue;
      // Solo contamos pedidos que Tadeo YA le entregó al vendedor (col BE).
      // Si Tadeo todavía no le pasó la mercadería (entrega futura), el vendedor
      // no podría haber cobrado nada — no debe aparecer en su saldo a Maleu.
      // Reportado por Tadeo (17/05/2026): Marcos sumaba pedidos de la semana
      // siguiente que aún no había recibido de Tadeo.
      var entVend = idxEntVend >= 0 ? String(dRed[r][idxEntVend] || '').trim() : '';
      if (entVend !== 'Entregado') continue;
      var estPagoMaleu = idxEstPagoMaleu >= 0 ? String(dRed[r][idxEstPagoMaleu] || '').trim() : '';
      if (estPagoMaleu === 'Pagado' || estPagoMaleu === 'Sí' || estPagoMaleu === 'Si') continue;
      var vendedor = String(dRed[r][7] || '').trim(); // col 8 Vendedor
      if (!vendedor) continue;
      var pedId = dRed[r][1];
      if (!pedId) continue;
      var total = Number(dRed[r][14]) || 0; // col 15 Total
      // A Pagar real por pedido (col 49). Fallback al cálculo 83%-plano si la
      // columna no existe o está vacía (pedidos viejos sin la col).
      var aPagarPed = idxAPagar >= 0 ? (Number(dRed[r][idxAPagar]) || 0) : 0;
      var fechaV = dRed[r][3];
      var fechaStr = fechaV instanceof Date
        ? Utilities.formatDate(fechaV, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy')
        : String(fechaV || '').trim();
      if (!porVendedor[vendedor]) {
        porVendedor[vendedor] = {
          vendedor: vendedor,
          ids: [],
          totalBruto: 0,
          totalAPagar: 0,
          fechas: [],
          tel: '',
          fps: {}
        };
      }
      var v = porVendedor[vendedor];
      v.ids.push(String(pedId));
      v.totalBruto += total;
      v.totalAPagar += aPagarPed;
      v.fechas.push(fechaStr);
      // Forma Pago: priorizar "Forma Pago a Maleu" (BA) si está; fallback a Forma Pago al cliente (M)
      var fpRed = idxFpMaleu >= 0 ? String(dRed[r][idxFpMaleu] || '').trim() : '';
      if (!fpRed) fpRed = String(dRed[r][12] || '').trim();
      if (fpRed) v.fps[fpRed] = (v.fps[fpRed] || 0) + 1;
    }
    // Convertir en cobros (uno por vendedor)
    Object.keys(porVendedor).forEach(function(vName) {
      var v = porVendedor[vName];
      var comPct = comisionPctVendedor[vName] !== undefined ? comisionPctVendedor[vName] : 17;
      // Si todos los pedidos del vendedor tienen "A Pagar" cargado, usar la suma
      // real. Si no, fallback al cálculo plano por % de comisión.
      var neto, comision;
      if (v.totalAPagar > 0) {
        neto = v.totalAPagar;
        comision = v.totalBruto - v.totalAPagar;
      } else {
        comision = Math.round(v.totalBruto * (comPct / 100));
        neto = v.totalBruto - comision;
      }
      // Fecha mas vieja
      var fechasValidas = v.fechas.filter(Boolean).sort(function(a,b){
        var pa=a.split('/'),pb=b.split('/');
        var ta=new Date(+pa[2],+pa[1]-1,+pa[0]).getTime();
        var tb=new Date(+pb[2],+pb[1]-1,+pb[0]).getTime();
        return ta-tb;
      });
      var fpDom = Object.keys(v.fps).sort(function(a,b){return v.fps[b]-v.fps[a]})[0] || 'Transferencia';
      out.push({
        key: 'Red|vendedor:' + vName,
        h: 'Red',
        id: v.ids.join(','),
        c: vName + ' (Vendedor Red)',
        t: '',
        '$': neto,
        fp: fpDom,
        de: '',
        fe: fechasValidas[0] || '',
        isVendedor: true,
        pedidosIds: v.ids,
        totalBruto: v.totalBruto,
        comision: comision,
        comisionPct: comPct
      });
    });
  }

  // Billetera (fondo de cambio): la PWA Ruta la muestra en el cuadro de Cobros para
  // saber cuánto tenés para dar vuelto y descontarla al devolver cambio.
  var _bilRuta = 0;
  try {
    var _shSBb = SS.getSheetByName('Saldo Base');
    if (_shSBb && _shSBb.getLastRow() > 1 && _shSBb.getLastColumn() >= 4) {
      _bilRuta = Number(_shSBb.getRange(_shSBb.getLastRow(), 4).getValue()) || 0;
    }
  } catch (_eb) { /* nada */ }
  var _resp = JSON.stringify({ts: Date.now(), cobros: out, billetera: _bilRuta});
  try { CacheService.getScriptCache().put(_ck, _resp, 30); } catch (_e) { /* nada */ }
  return ContentService.createTextOutput(_resp).setMimeType(ContentService.MimeType.JSON);
}

// Saldo billetera (fondo de cambio) — endpoint mini sin cache. Lo usa ruta.html
// al abrir el modal de Cobros para garantizar el saldo del momento (el modal
// necesita saber cuánto vuelto físico podés dar). El de cobrosPendientes está
// cacheado 30s y el sub-saldo billetera cambia con cada vuelto entregado.
function _doGetBilletera() {
  var bil = 0;
  try {
    var sh = SS.getSheetByName('Saldo Base');
    if (sh && sh.getLastRow() > 1 && sh.getLastColumn() >= 4) {
      bil = Number(sh.getRange(sh.getLastRow(), 4).getValue()) || 0;
    }
  } catch (_e) { /* nada */ }
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, billetera: bil, ts: Date.now() }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  CUPONES (campañas WATI con descuento en la tienda online)
// ════════════════════════════════════════════════════════════
//
//  Hoja 'Cupones' columnas:
//    A Codigo · B Tipo · C Valor · D Scope · E Vence · F UsosMax ·
//    G UsosActuales · H Segmento · I Mensaje · J Activo · K Stack · L Notas
//
//  Tipo:  PCT  → % sobre el scope (valor 0-100)
//         ARS  → $ fijo de descuento
//         ENVIO → anula el envío (valor ignorado)
//
//  Scope: TODO  → todo el subtotal
//         CATEGORIA:Sorrentinos → solo productos de esa categoría (cat del item)
//         PRODUCTO:SCo → solo ese abbr
//
//  Lee 'codigo' (case-insensitive) y opcionalmente 'itemsJson' (JSON con
//  [{abbr, cat, precio, qty}]). Si vienen items, calcula el descuento concreto;
//  si no vienen, solo valida que el cupón existe y está activo.

function _leerCupon(codigo) {
  var hCup = SS.getSheetByName('Cupones');
  if (!hCup) return null;
  var data = hCup.getDataRange().getValues();
  var key = String(codigo || '').trim().toUpperCase();
  if (!key) return null;
  for (var r = 1; r < data.length; r++) {
    var c = String(data[r][0] || '').trim().toUpperCase();
    if (c === key) {
      return {
        row:     r + 1,
        codigo:  c,
        tipo:    String(data[r][1] || '').trim().toUpperCase(),
        valor:   Number(data[r][2]) || 0,
        scope:   String(data[r][3] || 'TODO').trim(),
        vence:   data[r][4],
        usosMax: data[r][5] === '' || data[r][5] === null ? null : Number(data[r][5]),
        usosAct: Number(data[r][6]) || 0,
        segmento: String(data[r][7] || ''),
        mensaje: String(data[r][8] || ''),
        activo:  /^s[ií]|^yes|^true|^1/i.test(String(data[r][9] || '')),
        stack:   /^s[ií]|^yes|^true|^1/i.test(String(data[r][10] || ''))
      };
    }
  }
  return null;
}

function _doGetValidarCupon(e) {
  var codigo = (e && e.parameter && e.parameter.codigo) || '';
  var itemsJson = (e && e.parameter && e.parameter.itemsJson) || '';
  var resp = { ok: false, codigo: codigo };

  var cup = _leerCupon(codigo);
  if (!cup) {
    resp.error = 'Código no encontrado';
    return ContentService.createTextOutput(JSON.stringify(resp)).setMimeType(ContentService.MimeType.JSON);
  }
  if (!cup.activo) {
    resp.error = 'Cupón no disponible';
    return ContentService.createTextOutput(JSON.stringify(resp)).setMimeType(ContentService.MimeType.JSON);
  }
  if (cup.vence instanceof Date && cup.vence.getTime() < Date.now()) {
    resp.error = 'Cupón vencido';
    return ContentService.createTextOutput(JSON.stringify(resp)).setMimeType(ContentService.MimeType.JSON);
  }
  if (cup.usosMax !== null && cup.usosAct >= cup.usosMax) {
    resp.error = 'Cupón agotado';
    return ContentService.createTextOutput(JSON.stringify(resp)).setMimeType(ContentService.MimeType.JSON);
  }

  // Validación + cálculo del descuento si vienen items
  var items = [];
  if (itemsJson) { try { items = JSON.parse(itemsJson) || []; } catch(_e) { items = []; } }

  var scope = cup.scope || 'TODO';
  var scopeKey = '', scopeVal = '';
  var ix = scope.indexOf(':');
  if (ix >= 0) { scopeKey = scope.substring(0, ix).trim().toUpperCase(); scopeVal = scope.substring(ix + 1).trim(); }
  else scopeKey = scope.trim().toUpperCase();

  // Subtotal de los items que matchean el scope
  var subtotalScope = 0;
  items.forEach(function(it) {
    var pr = Number(it.precio) || 0;
    var qt = Number(it.qty) || 0;
    var monto = pr * qt;
    if (scopeKey === 'TODO') { subtotalScope += monto; return; }
    if (scopeKey === 'CATEGORIA' && String(it.cat || '') === scopeVal) { subtotalScope += monto; return; }
    if (scopeKey === 'PRODUCTO' && String(it.abbr || '') === scopeVal) { subtotalScope += monto; return; }
  });

  var descuento = 0;
  if (cup.tipo === 'PCT') descuento = Math.round(subtotalScope * (cup.valor / 100));
  else if (cup.tipo === 'ARS') descuento = Math.min(cup.valor, subtotalScope);
  else if (cup.tipo === 'ENVIO') descuento = 0; // la tienda lo maneja anulando shipping

  // Si todavía no hay productos del scope, el cupón se acepta igual (queda "pending").
  // La tienda muestra un card invitando a sumar el producto. Cuando el cliente lo
  // agrega, updateUI recalcula el descuento sin re-llamar al backend.
  resp.ok = true;
  resp.codigo = cup.codigo;
  resp.tipo = cup.tipo;
  resp.valor = cup.valor;
  resp.scope = cup.scope;
  resp.mensaje = cup.mensaje;
  resp.stack = cup.stack;
  resp.subtotalScope = subtotalScope;
  resp.descuento = descuento;
  resp.segmento = cup.segmento;
  resp.pending = (subtotalScope === 0 && cup.tipo !== 'ENVIO');
  return ContentService.createTextOutput(JSON.stringify(resp)).setMimeType(ContentService.MimeType.JSON);
}

// Suma un uso al cupón. Se llama desde la tienda al confirmar pedido. No falla
// si el código no existe (silent ok) — el tracking es best-effort.
function _doPostUsarCupon(data) {
  var resp = { ok: true };
  try {
    var cup = _leerCupon(data && data.codigo);
    if (cup) {
      var hCup = SS.getSheetByName('Cupones');
      hCup.getRange(cup.row, 7).setValue(cup.usosAct + 1);
      resp.codigo = cup.codigo;
      resp.usosActuales = cup.usosAct + 1;
    }
  } catch (_e) { /* best-effort */ }
  return ContentService.createTextOutput(JSON.stringify(resp)).setMimeType(ContentService.MimeType.JSON);
}

// Precios y costos por abreviatura — consumido por ruta.html (solapa +)
function _doGetPrecios() {
  var hProd = SS.getSheetByName('Productos');
  var out = {};
  if (hProd) {
    var data = hProd.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var abbr = String(data[r][2]).trim();
      if (!abbr) continue;
      var precio = parseFloat(String(data[r][8]).replace(/[$.]/g,'').replace(/,/g,'')) || 0;
      var costo = parseFloat(String(data[r][9]).replace(/[$.]/g,'').replace(/,/g,'')) || 0;
      out[abbr] = {p: precio, c: costo};
    }
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// Stock disponible en tiempo real — consumido por la tienda online
function _doGetStock() {
  var hProd = SS.getSheetByName('Productos');
  var out = {};
  if (hProd) {
    var data = hProd.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var abbr = String(data[r][2]).trim();
      if (!abbr) continue;
      var disp = Number(data[r][7]);
      out[abbr] = isNaN(disp) ? 0 : disp;
    }
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/* Devuelve para cada abreviatura de producto: f = stock físico actual,
   p = stock proyectado (físico + Σ cantidad de OCs en estado "Pedido"
   pendientes de recibir). Se usa cuando la tienda quiere mostrar al
   cliente "lo que hay + lo que viene en camino" — modo proyectado, que
   aplica los Jueves después de las 12hs (cuando ya cerramos la OC con
   el proveedor para el fin de semana). */
function _doGetStockFull() {
  var hProd = SS.getSheetByName('Productos');
  var out = {};
  if (hProd) {
    var data = hProd.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var abbr = String(data[r][2]).trim();
      if (!abbr) continue;
      var disp = Number(data[r][7]);
      var f = isNaN(disp) ? 0 : disp;
      out[abbr] = { f: f, p: f };
    }
  }
  var hOC = SS.getSheetByName('Orden de Compra');
  if (hOC) {
    var ocData = hOC.getDataRange().getValues();
    // L (idx 11) = Abbr, M (idx 12) = Cantidad, U (idx 20) = Estado OC
    for (var r = 1; r < ocData.length; r++) {
      var estado = String(ocData[r][20] || '').trim();
      if (estado !== 'Pedido') continue;
      var abbr = String(ocData[r][11] || '').trim();
      if (!abbr) continue;
      var qty = Number(ocData[r][12]);
      if (isNaN(qty) || qty <= 0) continue;
      if (!out[abbr]) out[abbr] = { f: 0, p: 0 };
      out[abbr].p += qty;
    }
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
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
 *  Devuelve pedidos pendientes de entrega de Home, Pilar y Clubes.
 *  Formato compacto para señal floja. */
function _invalidateEntregasCache() {
  try {
    var c = CacheService.getScriptCache();
    c.remove('entregas_v1');
    c.remove('entregas_v1_dia');  // se borran todas las variantes potenciales
  } catch(e) {}
}

function _doGetEntregas(e) {
  var dia = e && e.parameter && e.parameter.dia;
  var skipCache = e && e.parameter && e.parameter.fresh === '1';

  // ── CACHE: 20s, invalidado en cada POST exitoso (ver doPost finally) ──
  var cacheKey = 'entregas_v1' + (dia ? '_' + dia : '');
  var cache = null;
  if (!skipCache) {
    try {
      cache = CacheService.getScriptCache();
      var hit = cache.get(cacheKey);
      if (hit) {
        return ContentService
          .createTextOutput(hit)
          .setMimeType(ContentService.MimeType.JSON);
      }
    } catch(eC) { cache = null; }
  }

  var ABBRS = ['PPM','PPJyQ','PPCyQ','SCo','SJyQ','SCa','ECaC','EJyQ','ECyQ','EV',
               'TG','TLC','TC','F','PMu','PMa','PJyQ','PCC','PJyM'];

  // ── Pre-compute: items en OC por pedido (hoja|id) → { abbr: qty } ──
  // Source of truth de qué es OC: la hoja "Orden de Compra" donde cada fila representa
  // un producto a comprar al proveedor para un pedido específico. Así el frontend no
  // depende del JSON oD potencialmente vacío/desactualizado.
  var ocByPedido = {};
  var shOCEntregas = SS.getSheetByName('Orden de Compra');
  if (shOCEntregas && shOCEntregas.getLastRow() > 1) {
    // Cols: E(5)=Canal, F(6)=N° Pedido Origen, L(12)=Abbr, M(13)=Cantidad, U(21)=Estado OC
    var ocVals = shOCEntregas.getRange(2, 1, shOCEntregas.getLastRow() - 1, 25).getValues();
    for (var oi = 0; oi < ocVals.length; oi++) {
      var canalOC = String(ocVals[oi][4] || '').trim();   // E
      var pedidoOC = String(ocVals[oi][5] || '').trim();   // F
      var abbrOC = String(ocVals[oi][11] || '').trim();    // L
      var qtyOC = Number(ocVals[oi][12]) || 0;             // M
      var estadoOC = String(ocVals[oi][20] || '').trim();  // U
      if (!canalOC || !pedidoOC || !abbrOC || qtyOC <= 0) continue;
      // OC ya recibida sigue siendo OC para fines de armado: el producto vino del proveedor,
      // no del freezer; lo "armado" es separar lo del freezer. Solo excluimos canceladas.
      if (estadoOC === 'Cancelada' || estadoOC === 'Cancelado') continue;
      var key = canalOC + '|' + pedidoOC;
      if (!ocByPedido[key]) ocByPedido[key] = {};
      ocByPedido[key][abbrOC] = (ocByPedido[key][abbrOC] || 0) + qtyOC;
    }
  }

  var entregas = [];

  // ── Home, Pilar ──
  var ABBRS_PILAR = ['PPM','PPJyQ','PPCyQ','SQB','SL','SCo','SPyP','SJyQ','SE','SCa',
                     'ECaC','EJyQ','ECyQ','EV','TG','TLC','TC','F',
                     'PMu','PMa','PJyQ','PCC','PJyM'];
  ['Home', 'Pilar'].forEach(function(hoja) {
    var sh = SS.getSheetByName(hoja);
    if (!sh || sh.getLastRow() <= 1) return;
    var data = sh.getDataRange().getValues();
    var isPilar = (hoja === 'Pilar');
    var abbrsList = isPilar ? ABBRS_PILAR : ABBRS;
    var prodCount = isPilar ? 23 : 19;
    var colOriDet = isPilar ? 56 : 53; // Pilar BE / Home BB

    for (var r = 1; r < data.length; r++) {
      var estado = String(data[r][10]).trim();
      if (estado === 'Entregado' || estado === 'Cancelado') continue;

      var diaEntrega = _fechaADiaSemana(data[r][9]);
      var feISO = _fechaAISO(data[r][9]);
      if (dia && diaEntrega !== dia) continue;

      // Productos: Home y Pilar ambos comienzan en col W (23 = idx 22)
      var prodStart = 22;
      var productos = [];
      for (var p = 0; p < prodCount; p++) {
        var qty = Number(data[r][prodStart + p]) || 0;
        if (qty > 0) productos.push({ a: abbrsList[p], q: qty });
      }
      // Tartas (15/05/2026): cols al final, no consecutivas con productos viejos.
      // Home: BE-BH (idx 56-59) · Pilar: BH-BK (idx 59-62).
      var tartaStart = isPilar ? 59 : 56;
      var tartaAbbrs = ['TP', 'TJyQ', 'TCa', 'TV'];
      for (var tp = 0; tp < 4; tp++) {
        var qtyT = Number(data[r][tartaStart + tp]) || 0;
        if (qtyT > 0) productos.push({ a: tartaAbbrs[tp], q: qtyT });
      }
      // Wraps Claudia Polito (27/05/2026): cols al final tras "A Favor / Aplicado".
      // Home: BJ-BK (idx 61-62) · Pilar: BL-BM (idx 64-65).
      // Sin este loop, los pedidos que SOLO tienen wraps (caso Simón Panelo 05/06)
      // pasaban el filtro `productos.length === 0` → se saltean → no aparecen en RUTA.
      var wrapStart = isPilar ? 64 : 61;
      var wrapAbbrs = ['RC', 'RP'];
      for (var wp = 0; wp < 2; wp++) {
        var qtyW = Number(data[r][wrapStart + wp]) || 0;
        if (qtyW > 0) productos.push({ a: wrapAbbrs[wp], q: qtyW });
      }
      if (productos.length === 0) continue;

      var direccion = '', barrio = '', subBarrio = '', lote = '', telefono = '';
      if (hoja === 'Home') {
        // Home v2: AR(43)=Barrio, AS(44)=Sub Barrio, AT(45)=Domicilio-Lote, AU(46)=Tel
        barrio = String(data[r][43] || '').trim();
        subBarrio = String(data[r][44] || '').trim();
        lote = String(data[r][45] || '').trim();
        direccion = (subBarrio || barrio) + (lote ? ' · Lote ' + lote : '');
        telefono = String(data[r][46] || '');
      } else {
        // Pilar v2: AV(47)=Barrio/Dirección, AW(48)=Domicilio/Lote, AX(49)=Tel
        barrio = 'Pilar';
        subBarrio = String(data[r][47] || '').trim();
        lote = String(data[r][48] || '').trim();
        direccion = [subBarrio, lote].filter(Boolean).join(' · ');
        telefono = String(data[r][49] || '');
      }

      // Monto: ambas hojas usan Facturado V(21) con fallback a Total a cobrar Q(16)
      var montoEnt = Number(data[r][21]) || Number(data[r][16]) || 0;

      // Origen Detalle: parsear JSON {abbr:"D"|"OC"} para que el frontend sepa qué sale del depósito
      var oD = {};
      var rawOD = String(data[r][colOriDet] || '').trim();
      if (rawOD) { try { oD = JSON.parse(rawOD) || {}; } catch(eOD) { oD = {}; } }

      // OC items autoritativos: leídos directo de la hoja "Orden de Compra"
      var pedIdStr = String(Number(data[r][1]) || 0);
      var ocItems = [];
      var ocMap = ocByPedido[hoja + '|' + pedIdStr];
      if (ocMap) {
        Object.keys(ocMap).forEach(function(ab){ ocItems.push({ a: ab, q: ocMap[ab] }); });
      }

      // Hora del pedido (col A): para ordenar RUTA por "hora del pedido" (FIFO).
      var horaPed = data[r][0];
      var hrStr = '';
      if (horaPed instanceof Date) hrStr = Utilities.formatDate(horaPed, 'America/Argentina/Buenos_Aires', 'HH:mm');
      else if (horaPed) hrStr = String(horaPed).trim().slice(0,5);
      // Fecha del pedido (col D): junto con hr permite ordenar absoluto entre días.
      var fechaPed = data[r][3];
      var fpStr = '';
      if (fechaPed instanceof Date) fpStr = Utilities.formatDate(fechaPed, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
      else if (fechaPed) fpStr = String(fechaPed).trim();

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
        fe: feISO,
        es: estado,
        o: String(data[r][8] || '').trim(),
        fp: String(data[r][11] || '').trim(),
        ep: String(data[r][12] || '').trim(),
        $: montoEnt,
        p: productos,
        oD: oD,
        oc: ocItems,
        hr: hrStr,
        f: fpStr
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

      var diaC = _fechaADiaSemana(clubData[rc][12]); // M = Día de Entrega
      var feISOc = _fechaAISO(clubData[rc][12]);
      if (dia && diaC !== dia) continue;

      var prodsC = [];
      for (var pc = 0; pc < 8; pc++) {
        var qtyC = Number(clubData[rc][23 + pc]) || 0; // X-AE = cols 23-30
        if (qtyC > 0) prodsC.push({ a: ABBRS_CLUB[pc], q: qtyC });
      }
      // Empanadas Clubes (28/05/2026): cols AL-AO = idx 37-40
      var ABBRS_CLUB_EMP = ['ECaC','EJyQ','ECyQ','EV'];
      for (var ec = 0; ec < 4; ec++) {
        var qtyEc = Number(clubData[rc][37 + ec]) || 0;
        if (qtyEc > 0) prodsC.push({ a: ABBRS_CLUB_EMP[ec], q: qtyEc });
      }
      if (prodsC.length === 0) continue;

      var club = String(clubData[rc][8] || '').trim();
      var deporte = String(clubData[rc][9] || '').trim();
      var grupo = String(clubData[rc][10] || '').trim();
      var dirClub = [club, deporte, grupo].filter(Boolean).join(' · ');

      // Origen Detalle Clubes idx 35 (col AJ)
      var oDc = {};
      var rawODc = String(clubData[rc][35] || '').trim();
      if (rawODc) { try { oDc = JSON.parse(rawODc) || {}; } catch(eC) { oDc = {}; } }

      var pedIdStrC = String(Number(clubData[rc][1]) || 0);
      var ocItemsC = [];
      var ocMapC = ocByPedido['Clubes|' + pedIdStrC];
      if (ocMapC) {
        Object.keys(ocMapC).forEach(function(ab){ ocItemsC.push({ a: ab, q: ocMapC[ab] }); });
      }

      var horaPedC = clubData[rc][0];
      var hrStrC = '';
      if (horaPedC instanceof Date) hrStrC = Utilities.formatDate(horaPedC, 'America/Argentina/Buenos_Aires', 'HH:mm');
      else if (horaPedC) hrStrC = String(horaPedC).trim().slice(0,5);
      var fechaPedC = clubData[rc][3];
      var fpStrC = '';
      if (fechaPedC instanceof Date) fpStrC = Utilities.formatDate(fechaPedC, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
      else if (fechaPedC) fpStrC = String(fechaPedC).trim();
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
        fe: feISOc,
        es: estadoC,
        o: String(clubData[rc][11] || '').trim(), // L = Origen
        fp: String(clubData[rc][14] || '').trim(), // O = Forma de Pago
        ep: String(clubData[rc][15] || '').trim(), // P = Estado de Pago
        $: Number(clubData[rc][16]) || 0, // Q = Total
        p: prodsC,
        oD: oDc,
        oc: ocItemsC,
        hr: hrStrC,
        f: fpStrC
      });
    }
  }

  // ── Red ── (solo armado: items con Origen Detalle = "D" — los que salen del depósito Maleu)
  // Pedidos Red NO los entrega Tadeo, los retira el vendedor (Marcos). Pero Tadeo arma
  // la parte que sale del freezer antes de que Marcos venga a buscarla.
  // Para saber si Tadeo ya le pasó la mercadería al vendedor usamos col BE (idx 56) =
  // "Entregado a Vendedor". La col L "Estado de Entrega" es Marcos→cliente final
  // (la maneja Marcos desde red.html), NO es de armado para Tadeo.
  var ABBRS_RED = ['PPM','PPJyQ','PPCyQ','SQB','SL','SCo','SPyP','SJyQ','SE','SCa',
                   'ECaC','EJyQ','ECyQ','EV','TG','TLC','TC','F',
                   'PMu','PMa','PJyQ','PCC','PJyM'];
  var shRed = SS.getSheetByName('Red');
  if (shRed && shRed.getLastRow() > 1) {
    var redData = shRed.getDataRange().getValues();
    for (var rr = 1; rr < redData.length; rr++) {
      var estadoR = String(redData[rr][11]).trim();          // L = Estado entrega Marcos→cliente
      if (estadoR === 'Cancelado') continue;
      // Bug 05/06/2026: si Tadeo marca pedidos Red individualmente como Entregado (sin
      // usar el botón "ENTREGAR A VENDEDOR"), seguían apareciendo en ARMADO/RUTA del
      // Tadeo. Ahora también filtra por Estado de Entrega = Entregado.
      if (estadoR === 'Entregado') continue;
      var entVend = String(redData[rr][56] || '').trim();    // BE = Entregado a Vendedor (Tadeo→Marcos)
      if (entVend === 'Entregado') continue;                 // ya se lo entregamos a Marcos, no aparece en Armado

      var diaR = _fechaADiaSemana(redData[rr][10]) || ''; // K = Día de Entrega
      var feISOr = _fechaAISO(redData[rr][10]);
      if (dia && diaR !== dia) continue;

      // Origen Detalle Red idx 55 (col BD). Se devuelve al frontend para que
      // _itemsDepDe(e) pueda calcular qué parte sale del freezer.
      var oDr = {};
      var rawODr = String(redData[rr][55] || '').trim();
      if (rawODr) { try { oDr = JSON.parse(rawODr) || {}; } catch(eR) { oDr = {}; } }

      // TODOS los productos del pedido Red (sin filtrar por origen). Marcos retira
      // la tanda completa — sea del freezer (D) o del proveedor (OC). El frontend
      // separa con _itemsDepDe(e) usando oDr para mostrar solo D en Armado, y usa
      // p completo en la card violeta de RUTA (entrega total al vendedor).
      var prodsR = [];
      for (var pr = 0; pr < ABBRS_RED.length; pr++) {
        var qR = Number(redData[rr][21 + pr]) || 0;       // V..AR = idx 21..43
        if (qR > 0) prodsR.push({ a: ABBRS_RED[pr], q: qR });
      }
      // Tartas Red (15/05/2026): cols BF-BI = idx 57-60.
      var TARTA_ABBRS_RED = ['TP', 'TJyQ', 'TCa', 'TV'];
      for (var ptr = 0; ptr < 4; ptr++) {
        var qRT = Number(redData[rr][57 + ptr]) || 0;
        if (qRT > 0) prodsR.push({ a: TARTA_ABBRS_RED[ptr], q: qRT });
      }
      // Wraps Red: cols BJ-BK = idx 61-62. Faltaban acá también (mismo bug que el Panel:
      // caso Alejo Acuña 17/07). Va ANTES del check de pedido vacío para no saltar un
      // pedido que sea SOLO de wraps.
      var WRAP_ABBRS_RED = ['RC', 'RP'];
      for (var pwr2 = 0; pwr2 < 2; pwr2++) {
        var qRW = Number(redData[rr][61 + pwr2]) || 0;
        if (qRW > 0) prodsR.push({ a: WRAP_ABBRS_RED[pwr2], q: qRW });
      }
      if (prodsR.length === 0) continue;                   // pedido vacío (defensivo, no debería pasar)

      var vendedor = String(redData[rr][7] || '').trim();
      var clienteFinal = String(redData[rr][8] || '').trim();
      var horaPedR = redData[rr][0];
      var hrStrR = '';
      if (horaPedR instanceof Date) hrStrR = Utilities.formatDate(horaPedR, 'America/Argentina/Buenos_Aires', 'HH:mm');
      else if (horaPedR) hrStrR = String(horaPedR).trim().slice(0,5);
      var fechaPedR = redData[rr][3];
      var fpStrR = '';
      if (fechaPedR instanceof Date) fpStrR = Utilities.formatDate(fechaPedR, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
      else if (fechaPedR) fpStrR = String(fechaPedR).trim();

      entregas.push({
        id: Number(redData[rr][1]) || 0,
        h: 'Red',
        r: rr + 1,
        c: clienteFinal || vendedor,
        t: String(redData[rr][51] || '').trim(),
        d: vendedor + (clienteFinal ? ' · ' + clienteFinal : ''),
        b: vendedor,
        sb: vendedor,
        l: String(redData[rr][50] || '').trim(),
        de: diaR,
        fe: feISOr,
        es: estadoR,
        o: String(redData[rr][9] || '').trim(),           // J = Origen
        fp: String(redData[rr][12] || '').trim(),
        ep: String(redData[rr][13] || '').trim(),
        $: Number(redData[rr][14]) || 0,                  // O = Total
        p: prodsR,
        oD: oDr,
        retira: vendedor,                                 // Indicador: lo retira el vendedor, no lo entrega Tadeo
        hr: hrStrR,
        f: fpStrR
      });
    }
  }

  var jsonStr = JSON.stringify({ ts: Date.now(), e: entregas });

  // Guardar en cache 5s. Apps Script CacheService limita 100KB por key; si excede, no cachea.
  // TTL corto: rapido para refresh sucesivos pero stale data dura poco si edito el Sheet manualmente.
  if (cache && jsonStr.length < 100000) {
    try { cache.put(cacheKey, jsonStr, 5); } catch(eP) {}
  }

  return ContentService
    .createTextOutput(jsonStr)
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════════
//  BÚSQUEDA — API para la página de búsqueda de productos (busqueda.html)
// ══════════════════════════════════════════════════════════════

/** GET ?action=busqueda
 *  Devuelve OC pendientes/pedidas agrupadas por proveedor y por cliente.
 *  Formato compacto para PWA offline-first. */
/** GET ?action=vendedores — devuelve vendedores activos con sus barrios cubiertos
 *  Formato: { ts, vendedores: [{nombre, wa, barrios:[...], partido, localidad}] } */
function _doGetVendedores() {
  var sh = SS.getSheetByName('Vendedores');
  if (!sh || sh.getLastRow() <= 1) {
    return ContentService.createTextOutput(JSON.stringify({ ts: Date.now(), vendedores: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var data = sh.getDataRange().getValues();
  var vendedores = [];
  for (var r = 1; r < data.length; r++) {
    var estado = String(data[r][6] || '').trim();
    if (estado !== 'Activo') continue;
    var nombre = String(data[r][0] || '').trim();
    if (!nombre) continue;
    var barrios = String(data[r][3] || '').split(',').map(function(b) { return b.trim(); }).filter(Boolean);
    vendedores.push({
      nombre: nombre,
      wa: String(data[r][1] || '').trim(),
      alias: String(data[r][2] || '').trim(),           // col C — alias MP (opcional)
      barrios: barrios,
      partido: String(data[r][4] || '').trim(),
      localidad: String(data[r][5] || '').trim()
    });
  }
  return ContentService.createTextOutput(JSON.stringify({ ts: Date.now(), vendedores: vendedores }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Devuelve la matriz de permisos {rol: [tabs permitidas]} leyendo la hoja Permisos.
 *  Si la hoja no existe, la crea con valores default y la devuelve. */
function _getPermisos() {
  var sh = SS.getSheetByName('Permisos');
  if (!sh) {
    sh = SS.insertSheet('Permisos');
    sh.appendRow(['Rol', 'Tab', 'Acceso']);
    var defaults = [
      ['admin','inicio','Sí'],['admin','ventas','Sí'],['admin','pedidos','Sí'],
      ['admin','caja','Sí'],['admin','egresos','Sí'],['admin','stock','Sí'],
      ['admin','bbdd','Sí'],['admin','ruta','Sí'],['admin','miportal','Sí'],['admin','busqueda','Sí'],
      ['empleado','inicio','Sí'],['empleado','ventas','Sí'],['empleado','pedidos','Sí'],
      ['empleado','caja','No'],['empleado','egresos','No'],['empleado','stock','Sí'],
      ['empleado','bbdd','No'],['empleado','ruta','Sí'],['empleado','miportal','No'],['empleado','busqueda','Sí'],
      ['repartidor','inicio','No'],['repartidor','ventas','No'],['repartidor','pedidos','No'],
      ['repartidor','caja','No'],['repartidor','egresos','No'],['repartidor','stock','No'],
      ['repartidor','bbdd','No'],['repartidor','ruta','Sí'],['repartidor','miportal','No'],['repartidor','busqueda','No'],
      ['vendedor','inicio','No'],['vendedor','ventas','No'],['vendedor','pedidos','No'],
      ['vendedor','caja','No'],['vendedor','egresos','No'],['vendedor','stock','No'],
      ['vendedor','bbdd','No'],['vendedor','ruta','No'],['vendedor','miportal','Sí'],['vendedor','busqueda','No']
    ];
    sh.getRange(2, 1, defaults.length, 3).setValues(defaults);
    sh.getRange('A1:C1').setFontWeight('bold').setBackground('#331C1C').setFontColor('#fff');
    sh.setColumnWidths(1, 3, 130);
  }
  // Migración idempotente: tabs nuevas que deben existir para ciertos roles.
  // Si la combinación rol+tab falta en la hoja, se agrega con su valor default.
  // Esto evita tener que tocar Permisos a mano cada vez que se agrega una tab.
  var NEW_TABS = [
    { rol:'admin',      tab:'mireparto',   def:'Sí' },
    { rol:'empleado',   tab:'mireparto',   def:'Sí' },
    { rol:'repartidor', tab:'mireparto',   def:'Sí' },
    { rol:'vendedor',   tab:'mireparto',   def:'No' },
    { rol:'admin',      tab:'pedidoshome', def:'Sí' },
    { rol:'empleado',   tab:'pedidoshome', def:'Sí' },
    { rol:'repartidor', tab:'pedidoshome', def:'Sí' },
    { rol:'vendedor',   tab:'pedidoshome', def:'No' }
  ];
  var existing = {};
  if (sh.getLastRow() > 1) {
    var allRows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
    for (var er = 0; er < allRows.length; er++) {
      var rE = String(allRows[er][0] || '').trim().toLowerCase();
      var tE = String(allRows[er][1] || '').trim().toLowerCase();
      if (rE && tE) existing[rE + '|' + tE] = true;
    }
  }
  var toAdd = [];
  for (var nt = 0; nt < NEW_TABS.length; nt++) {
    var k = NEW_TABS[nt].rol + '|' + NEW_TABS[nt].tab;
    if (!existing[k]) toAdd.push([NEW_TABS[nt].rol, NEW_TABS[nt].tab, NEW_TABS[nt].def]);
  }
  if (toAdd.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAdd.length, 3).setValues(toAdd);
  }

  var matrix = {};
  if (sh.getLastRow() > 1) {
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < rows.length; i++) {
      var rol = String(rows[i][0] || '').trim().toLowerCase();
      var tab = String(rows[i][1] || '').trim().toLowerCase();
      var ac  = String(rows[i][2] || '').trim().toLowerCase();
      var ok = ac === 'sí' || ac === 'si' || ac === 'true' || ac === '1' || ac === 'yes';
      if (!rol || !tab) continue;
      if (!matrix[rol]) matrix[rol] = [];
      if (ok) matrix[rol].push(tab);
    }
  }
  return matrix;
}

/** POST action=loginUsuario — valida credenciales del Panel.
 *  Busca primero en hoja "Usuarios" (admin/empleado/repartidor).
 *  Si no encuentra, fallback a hoja "Vendedores" (rol implícito = vendedor).
 *  Si no existe la hoja Usuarios, la crea automáticamente con un admin "tadeo"/"1234".
 *  Recibe { action:'loginUsuario', usuario, pin }
 *  Devuelve { ok, rol, nombre, usuario, tabs:[...], ...extra } o { ok:false, err } */
function _doPostLoginUsuario(data) {
  var usuario = String(data.usuario || '').trim().toLowerCase();
  var pin     = String(data.pin || '').trim();
  if (!usuario || !pin) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'Usuario y PIN requeridos' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // 1) Buscar en hoja Usuarios (creandola si no existe)
  var sh = SS.getSheetByName('Usuarios');
  if (!sh) {
    sh = SS.insertSheet('Usuarios');
    sh.appendRow(['Usuario','PIN','Rol','Nombre','Activo','Notas']);
    sh.appendRow(['tadeo','1234','admin','Tadeo Ustariz','Sí','Cambiar PIN al primer login']);
    sh.getRange('A1:F1').setFontWeight('bold').setBackground('#331C1C').setFontColor('#fff');
    sh.setColumnWidths(1, 6, 130);
  }
  if (sh.getLastRow() > 1) {
    var perms = _getPermisos();
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
    for (var r = 0; r < rows.length; r++) {
      var u = String(rows[r][0] || '').trim().toLowerCase();
      var p = String(rows[r][1] || '').trim();
      var rol = String(rows[r][2] || '').trim().toLowerCase() || 'empleado';
      var nombre = String(rows[r][3] || '').trim();
      var activo = String(rows[r][4] || '').trim().toLowerCase();
      var act = activo === 'sí' || activo === 'si' || activo === 'true' || activo === '1' || activo === '';
      if (u === usuario && p === pin && act) {
        return ContentService.createTextOutput(JSON.stringify({
          ok: true, rol: rol, nombre: nombre || u, usuario: u, tabs: perms[rol] || []
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
  }
  // 2) Fallback: hoja Vendedores (rol vendedor)
  var shV = SS.getSheetByName('Vendedores');
  if (shV && shV.getLastRow() > 1) {
    var perms2 = _getPermisos();
    var rowsV = shV.getDataRange().getValues();
    for (var rv = 1; rv < rowsV.length; rv++) {
      var uv = String(rowsV[rv][7] || '').trim().toLowerCase();
      var pv = String(rowsV[rv][8] || '').trim();
      var estado = String(rowsV[rv][6] || '').trim();
      if (uv === usuario && pv === pin && estado === 'Activo') {
        return ContentService.createTextOutput(JSON.stringify({
          ok: true, rol: 'vendedor', nombre: String(rowsV[rv][0] || ''), usuario: uv,
          tabs: perms2['vendedor'] || [],
          wa: String(rowsV[rv][1] || ''),
          aliasMp: String(rowsV[rv][2] || '').trim(),
          comision: Number(rowsV[rv][9]) || 17,
          barrios: String(rowsV[rv][3] || '').split(',').map(function(b){return b.trim();}).filter(Boolean),
          partido: String(rowsV[rv][4] || ''),
          localidad: String(rowsV[rv][5] || '')
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'Usuario o PIN incorrectos' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** POST action=loginVendedor — valida credenciales contra hoja Vendedores
 *  Recibe { action:'loginVendedor', usuario, pin }
 *  Devuelve { ok, nombre, wa, comision } o { ok:false, err } */
function _doPostLoginVendedor(data) {
  var usuario = String(data.usuario || '').trim().toLowerCase();
  var pin     = String(data.pin || '').trim();

  if (!usuario || !pin) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'Usuario y PIN requeridos' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var sh = SS.getSheetByName('Vendedores');
  if (!sh || sh.getLastRow() <= 1) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'Sin vendedores' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var rows = sh.getDataRange().getValues();
  for (var r = 1; r < rows.length; r++) {
    var u = String(rows[r][7] || '').trim().toLowerCase(); // col H = Usuario
    var p = String(rows[r][8] || '').trim();               // col I = PIN
    var estado = String(rows[r][6] || '').trim();          // col G = Estado
    if (u === usuario && p === pin && estado === 'Activo') {
      return ContentService.createTextOutput(JSON.stringify({
        ok: true,
        nombre: String(rows[r][0] || ''),
        wa: String(rows[r][1] || ''),
        aliasMp: String(rows[r][2] || '').trim(),  // col C — alias MP del vendedor (vacio si no cargado)
        comision: Number(rows[r][9]) || 17,
        barrios: String(rows[r][3] || '').split(',').map(function(b){return b.trim();}).filter(Boolean),
        partido: String(rows[r][4] || ''),
        localidad: String(rows[r][5] || '')
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'Usuario o PIN incorrectos' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** GET ?action=dashboardVendedor&nombre=X — stats del vendedor
 *  Devuelve: { ok, stats: {semana, mes, comisionTotal, totalPedidos}, pedidos: [], clientes: [] } */
/** POST action=updatePedidoRed — vendedor actualiza estado de UN pedido suyo
 *  Recibe { action:'updatePedidoRed', pedidoId, vendedor, updates: {...} }
 *  updates puede tener: entrega, cobroCliente, formaPagoCliente, propinaEf, propinaTr, formaPagoMaleu, estadoPagoMaleu
 *  Valida que el pedido sea del vendedor (prevención cross-vendor). */
function _doPostUpdatePedidoRed(data) {
  var pedidoId = String(data.pedidoId || '').trim();
  var vendedor = String(data.vendedor || '').trim();
  var updates = data.updates || {};

  if (!pedidoId || !vendedor) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'pedidoId y vendedor requeridos' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var sh = SS.getSheetByName('Red');
  if (!sh) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'Hoja Red no encontrada' }))
      .setMimeType(ContentService.MimeType.JSON);

  // Buscar fila por N° Pedido (col B=2) — admite "R-001", "007", "7"
  // IMPORTANTE: los N° se reinician por semana ISO, así que iteramos de ATRÁS hacia adelante
  // y nos quedamos con el último match (el más reciente) que pertenezca al vendedor.
  var pedidoNum = pedidoId.replace(/^R-/i, '').replace(/^0+/, '') || '0';
  var allData = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues(); // A-H para buscar B (N°) y H (Vendedor)
  var row = -1;
  for (var r = allData.length - 1; r >= 0; r--) {
    var n = String(allData[r][1]).trim().replace(/^R-/i, '').replace(/^0+/, '') || '0';
    if (n !== pedidoNum) continue;
    var v = String(allData[r][7]).trim();
    if (v !== vendedor) continue; // saltar pedidos viejos de otro vendedor con mismo N°
    row = r + 2; // +2 porque allData empieza en fila 2 y r es 0-indexed
    break;
  }

  if (row === -1) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'Pedido no encontrado' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Aplicar updates (cada acción tiene su col)
  // L(12)=Estado Entrega, M(13)=Forma Pago, N(14)=Estado Pago, S(19)=Propina Ef, T(20)=Propina Tr
  // Q(17)=Efectivo, R(18)=Transferencia (para cobro mixto)
  // BA(53)=Forma Pago a Maleu, BB(54)=Estado Pago a Maleu, BC(55)=Fecha Pago a Maleu
  var applied = [];
  var origen = String(sh.getRange(row, 10).getValue() || '').trim(); // col J = Origen

  if (updates.hasOwnProperty('entrega')) {
    var estadoPrev = String(sh.getRange(row, 12).getValue() || '').trim();
    var nuevoEstado = updates.entrega ? 'Entregado' : 'Pendiente';
    sh.getRange(row, 12).setValue(nuevoEstado);
    applied.push('entrega=' + nuevoEstado);
    // Descontar stock si Origen=Depósito y pasa a Entregado
    if (nuevoEstado === 'Entregado' && estadoPrev !== 'Entregado' && origen === 'Deposito') {
      var hProd = SS.getSheetByName('Productos');
      if (hProd) _redStockFisico(sh, row, hProd, -1);
    } else if (nuevoEstado === 'Entregado' && estadoPrev !== 'Entregado' && origen === 'Mixto') {
      var hProdM = SS.getSheetByName('Productos');
      if (hProdM) _redStockFisicoMixto(sh, row, hProdM, -1);
    }
    if (nuevoEstado !== 'Entregado' && estadoPrev === 'Entregado' && origen === 'Deposito') {
      var hProd2 = SS.getSheetByName('Productos');
      if (hProd2) _redStockFisico(sh, row, hProd2, +1);
    } else if (nuevoEstado !== 'Entregado' && estadoPrev === 'Entregado' && origen === 'Mixto') {
      var hProd2M = SS.getSheetByName('Productos');
      if (hProd2M) _redStockFisicoMixto(sh, row, hProd2M, +1);
    }
  }

  if (updates.hasOwnProperty('formaPagoCliente')) {
    sh.getRange(row, 13).setValue(String(updates.formaPagoCliente || ''));
    applied.push('formaPagoCliente=' + updates.formaPagoCliente);
  }

  if (updates.hasOwnProperty('cobroCliente')) {
    sh.getRange(row, 14).setValue(updates.cobroCliente ? 'Cobrado' : 'No Cobrado');
    applied.push('cobroCliente=' + (updates.cobroCliente ? 'Cobrado' : 'No Cobrado'));
  }

  if (updates.hasOwnProperty('propinaEf')) {
    sh.getRange(row, 19).setValue(Number(updates.propinaEf) || 0);
    applied.push('propinaEf=' + updates.propinaEf);
  }

  if (updates.hasOwnProperty('propinaTr')) {
    sh.getRange(row, 20).setValue(Number(updates.propinaTr) || 0);
    applied.push('propinaTr=' + updates.propinaTr);
  }

  // Cobro mixto: recibe efectivoMonto y transferenciaMonto. Setea Q=17, R=18,
  // marca M=Mixto y N=Cobrado si la suma > 0.
  if (updates.hasOwnProperty('efectivoMonto') || updates.hasOwnProperty('transferenciaMonto')) {
    var efMonto = Number(updates.efectivoMonto) || 0;
    var trMonto = Number(updates.transferenciaMonto) || 0;
    sh.getRange(row, 17).setValue(efMonto);
    sh.getRange(row, 18).setValue(trMonto);
    if (efMonto > 0 && trMonto > 0) {
      sh.getRange(row, 13).setValue('Mixto');
    } else if (efMonto > 0) {
      sh.getRange(row, 13).setValue('Efectivo');
    } else if (trMonto > 0) {
      sh.getRange(row, 13).setValue('Transferencia');
    }
    if (efMonto + trMonto > 0) {
      sh.getRange(row, 14).setValue('Cobrado');
    }
    applied.push('cobroMixto=ef:' + efMonto + ',tr:' + trMonto);
  }

  if (updates.hasOwnProperty('cancelar') && updates.cancelar) {
    var estadoPrevC = String(sh.getRange(row, 12).getValue() || '').trim();
    sh.getRange(row, 12).setValue('Cancelado');
    applied.push('cancelado');
    // Si estaba Entregado y origen=Deposito/Mixto, devolver el stock al freezer
    if (estadoPrevC === 'Entregado' && (origen === 'Deposito' || origen === 'Mixto')) {
      var hProdC = SS.getSheetByName('Productos');
      if (hProdC) {
        if (origen === 'Deposito') _redStockFisico(sh, row, hProdC, +1);
        else _redStockFisicoMixto(sh, row, hProdC, +1);
      }
    }
  }

  if (updates.hasOwnProperty('formaPagoMaleu')) {
    sh.getRange(row, 53).setValue(String(updates.formaPagoMaleu || ''));
    applied.push('formaPagoMaleu=' + updates.formaPagoMaleu);
  }

  if (updates.hasOwnProperty('estadoPagoMaleu')) {
    var nuevoEst = updates.estadoPagoMaleu ? 'Pagado' : 'Pendiente';
    sh.getRange(row, 54).setValue(nuevoEst);
    applied.push('estadoPagoMaleu=' + nuevoEst);
    // Fecha automática cuando marca Pagado
    if (updates.estadoPagoMaleu) {
      var argDate = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
      var dd = String(argDate.getDate()).padStart(2, '0');
      var mm = String(argDate.getMonth() + 1).padStart(2, '0');
      var yyyy = argDate.getFullYear();
      sh.getRange(row, 55).setValue(dd + '/' + mm + '/' + yyyy);
    } else {
      sh.getRange(row, 55).setValue('');
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true, applied: applied }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** POST action=editarPedidoRed — vendedor reescribe productos y total de UN pedido suyo
 *  Body: { action:'editarPedidoRed', pedidoId, vendedor, items:[{id,qty}], total }
 *  Solo se permite si el pedido NO está Entregado ni Cobrado. Recalcula total, costo,
 *  Q/R según forma de pago, y reescribe las cols de productos V-AR. Las fórmulas
 *  U, AT, AU, AV, AW se actualizan solas porque referencian a O y AS.
 */
function _doPostEditarPedidoRed(data) {
  var pedidoId = String(data.pedidoId || '').trim();
  var vendedor = String(data.vendedor || '').trim();
  var items    = data.items || [];
  var totalNew = Number(data.total) || 0;

  if (!pedidoId || !vendedor) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'pedidoId y vendedor requeridos' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (!items.length || totalNew <= 0) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'el pedido no puede quedar vacío' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var sh = SS.getSheetByName('Red');
  if (!sh) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'Hoja Red no encontrada' }))
      .setMimeType(ContentService.MimeType.JSON);

  // Buscar fila por N° Pedido + Vendedor (mismo criterio que updatePedidoRed)
  var pedidoNum = pedidoId.replace(/^R-/i, '').replace(/^0+/, '') || '0';
  var allData   = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  var row = -1;
  for (var r = allData.length - 1; r >= 0; r--) {
    var n = String(allData[r][1]).trim().replace(/^R-/i, '').replace(/^0+/, '') || '0';
    if (n !== pedidoNum) continue;
    var v = String(allData[r][7]).trim();
    if (v !== vendedor) continue;
    row = r + 2;
    break;
  }
  if (row === -1) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'Pedido no encontrado' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var estadoEntrega = String(sh.getRange(row, 12).getValue() || '').trim();
  var estadoPago    = String(sh.getRange(row, 14).getValue() || '').trim();
  if (estadoEntrega === 'Entregado') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'No se puede editar un pedido ya entregado' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (estadoEntrega === 'Cancelado') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'No se puede editar un pedido cancelado' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (estadoPago === 'Cobrado') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'No se puede editar un pedido ya cobrado' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var formaPago = String(sh.getRange(row, 13).getValue() || '').trim();
  var envio     = Number(sh.getRange(row, 16).getValue()) || 0;

  // Reset todas las cols de productos a 0 y aplicar las nuevas cantidades
  var qtys = {};
  items.forEach(function(it) {
    var id = Number(it.id);
    var q  = Number(it.qty) || 0;
    if (q > 0 && RED_PRODUCT_COLS[id]) qtys[id] = q;
  });

  Object.keys(RED_PRODUCT_COLS).forEach(function(idStr) {
    var id  = Number(idStr);
    var col = RED_PRODUCT_COLS[id];
    sh.getRange(row, col).setValue(qtys[id] || 0);
  });

  // Recalcular costo desde hoja Productos
  var costoTotal = 0;
  var hProductos = SS.getSheetByName('Productos');
  if (hProductos) {
    var prodData = hProductos.getDataRange().getValues();
    Object.keys(qtys).forEach(function(idStr) {
      var id = Number(idStr);
      var abbr = RED_ID_TO_ABBR[id];
      if (!abbr) return;
      for (var rp = 1; rp < prodData.length; rp++) {
        if (String(prodData[rp][2]).trim() === abbr) {
          var costoUnit = Number(prodData[rp][9]) || 0;
          costoTotal += costoUnit * qtys[id];
          break;
        }
      }
    });
  }

  // O (15) = Total, AS (45) = Costo. El frontend de edición manda `total` = SOLO
  // productos (sin envío). Sumamos el envío existente (col P) para que O = productos
  // + envío y Facturado (=O-P) siga dando productos. El envío es 100% del vendedor.
  sh.getRange(row, 15).setValue(totalNew + envio);
  sh.getRange(row, 45).setValue(costoTotal);

  // Q (17) Efectivo / R (18) Transferencia segun forma de pago vigente
  var totalSinEnvio = totalNew;
  if (formaPago === 'Efectivo') {
    sh.getRange(row, 17).setValue(totalSinEnvio);
    sh.getRange(row, 18).setValue(0);
  } else if (formaPago === 'Transferencia') {
    sh.getRange(row, 17).setValue(0);
    sh.getRange(row, 18).setValue(totalSinEnvio);
  }
  // Mixto: como solo se setea al cobrar y editar requiere no-cobrado, no debería caer acá.

  // Garantizar fórmulas vivas (por si la fila vino sin ellas)
  // Facturado Red = O − P (productos, SIN envío): el envío es 100% del vendedor,
  // no entra en la base de comisión (17%) ni en lo que rinde a Maleu (83%).
  if (!sh.getRange(row, 21).getFormula()) sh.getRange(row, 21).setFormula('=O' + row + '-P' + row);
  if (!sh.getRange(row, 46).getFormula()) sh.getRange(row, 46).setFormula('=U' + row + '-AS' + row);
  if (!sh.getRange(row, 47).getFormula()) sh.getRange(row, 47).setFormula('=U' + row + '*17/100');
  if (!sh.getRange(row, 48).getFormula()) sh.getRange(row, 48).setFormula('=AT' + row + '-AU' + row);
  if (!sh.getRange(row, 49).getFormula()) sh.getRange(row, 49).setFormula('=U' + row + '*83/100');
  // Ganancia Vendedor (col 65 = BM) = Comisión 17% + Envío
  if (!sh.getRange(row, 65).getFormula()) sh.getRange(row, 65).setFormula('=AU' + row + '+P' + row);

  SpreadsheetApp.flush();

  return ContentService.createTextOutput(JSON.stringify({
    ok: true, total: totalNew, costo: costoTotal, row: row
  })).setMimeType(ContentService.MimeType.JSON);
}

/** POST action=marcarSemanaPagadaRed — marca como Pagado a Maleu todos los pedidos
 *  cobrados de un vendedor en una semana (ISO) dada.
 *  Body: { action:'marcarSemanaPagadaRed', vendedor, year, semana }
 */
/**
 * POST action=marcarSemanaPagadaRed — pagos parciales a Maleu
 * Body: { vendedor, year, semana, ef, tr }
 *
 * Flujo:
 *  1. Registra SIEMPRE el pago (parcial o total) en hoja "Pagos Red Liq".
 *  2. Calcula acumulado_pagado = sum(EF+TR) de todos los pagos de vendedor+año+semana.
 *  3. Calcula a_pagar_total = sum(col AW) de pedidos Cobrados de esa semana.
 *  4. Si acumulado_pagado >= a_pagar_total → marca todos los pedidos como Pagado.
 *  5. Sino → quedan pendientes (la semana sigue abierta).
 */
function _doPostMarcarSemanaPagadaRed(data) {
  var vendedor = String(data.vendedor || '').trim();
  var year     = Number(data.year) || 0;
  var semana   = Number(data.semana) || 0;
  var efPago   = Number(data.ef) || 0;
  var trPago   = Number(data.tr) || 0;
  if (!vendedor || !year || !semana) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'vendedor, year y semana requeridos' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (efPago < 0 || trPago < 0 || (efPago + trPago) <= 0) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'monto inválido' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var sh = SS.getSheetByName('Red');
  if (!sh || sh.getLastRow() <= 1) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'Hoja Red vacía' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var numRows = sh.getLastRow() - 1;
  var rng = sh.getRange(2, 1, numRows, 55).getValues();

  var argDate = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var dd = String(argDate.getDate()).padStart(2, '0');
  var mm = String(argDate.getMonth() + 1).padStart(2, '0');
  var yyyy = argDate.getFullYear();
  var fechaStr = dd + '/' + mm + '/' + yyyy;

  // 1) Calcular a_pagar_total de la semana (suma AW de todos los Cobrados)
  var totalAPagar = 0;
  var pedidosIds = [];
  var filasPendientes = []; // filas a marcar como Pagado si cerramos la semana
  for (var r = 0; r < rng.length; r++) {
    var v = String(rng[r][7] || '').trim();
    if (v !== vendedor) continue;
    var rowSem = Number(rng[r][5]) || 0;
    var rowYr  = Number(rng[r][6]) || 0;
    if (rowYr !== year || rowSem !== semana) continue;
    var estadoPago = String(rng[r][13] || '').trim();
    if (estadoPago !== 'Cobrado') continue;
    totalAPagar += Number(rng[r][48]) || 0;
    pedidosIds.push(String(rng[r][1] || ''));
    var estadoPagoMaleu = String(rng[r][53] || '').trim();
    var yaPagado = (estadoPagoMaleu === 'Pagado' || estadoPagoMaleu === 'Sí' || estadoPagoMaleu === 'Si');
    if (!yaPagado) filasPendientes.push(r + 2);
  }

  if (totalAPagar <= 0) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'No hay pedidos cobrados en esa semana' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 2) Registrar pago en hoja "Pagos Red Liq"
  var shLiq = SS.getSheetByName('Pagos Red Liq');
  if (!shLiq) {
    shLiq = SS.insertSheet('Pagos Red Liq');
    shLiq.appendRow(['Timestamp','Fecha','Vendedor','Año','Semana','Efectivo','Transferencia','Total','A Pagar Esperado','Diferencia','Pedidos','# Pedidos']);
    shLiq.setFrozenRows(1);
    shLiq.getRange('A1:L1').setFontWeight('bold').setBackground('#f3e5ab');
  }
  shLiq.appendRow([
    new Date(), fechaStr, vendedor, year, semana,
    efPago, trPago, efPago + trPago, totalAPagar, 0,
    pedidosIds.join(', '), pedidosIds.length
  ]);

  // 3) Acumulado pagado (incluye el que acabamos de registrar)
  var acumEf = 0, acumTr = 0;
  var dLiq = shLiq.getDataRange().getValues();
  for (var rl = 1; rl < dLiq.length; rl++) {
    if (String(dLiq[rl][2]).trim() !== vendedor) continue;
    if ((Number(dLiq[rl][3]) || 0) !== year) continue;
    if ((Number(dLiq[rl][4]) || 0) !== semana) continue;
    acumEf += Number(dLiq[rl][5]) || 0;
    acumTr += Number(dLiq[rl][6]) || 0;
  }
  var acumTotal = acumEf + acumTr;
  var saldo = totalAPagar - acumTotal; // cuánto falta
  var cerrada = (acumTotal + 0.01 >= totalAPagar);

  // 4) Si la semana quedó cubierta → marcar pedidos como Pagado
  var marcados = 0;
  if (cerrada) {
    // Forma a escribir en BA: si ambos acumulados > 0 → Mixto; sino el método usado
    var formaPago = (acumEf > 0 && acumTr > 0) ? 'Mixto' : (acumEf > 0 ? 'Efectivo' : 'Transferencia');
    for (var k = 0; k < filasPendientes.length; k++) {
      var fila = filasPendientes[k];
      sh.getRange(fila, 53).setValue(formaPago);
      sh.getRange(fila, 54).setValue('Pagado');
      sh.getRange(fila, 55).setValue(fechaStr);
      marcados++;
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    fecha: fechaStr,
    pagoEf: efPago, pagoTr: trPago, pagoTotal: efPago + trPago,
    acumEf: acumEf, acumTr: acumTr, acumTotal: acumTotal,
    aPagarTotal: totalAPagar,
    saldoRestante: Math.max(0, saldo),
    semanaCerrada: cerrada,
    marcados: marcados
  })).setMimeType(ContentService.MimeType.JSON);
}

/** GET ?action=resolverVendedor&usuario=X[&wa=Y] — devuelve el nombre canónico
 *  ACTUAL de un vendedor a partir de su identidad estable (usuario o WhatsApp).
 *  El portal lo usa para auto-corregir la sesión si el vendedor fue renombrado
 *  en la hoja (bug Fini 16/07/2026: sesión vieja "Josefina" vs pedidos "Fini").
 *  Solo devuelve nombre/comisión — no expone PIN. */
function _doGetResolverVendedor(e) {
  var usuario = String(e && e.parameter && e.parameter.usuario || '').trim().toLowerCase();
  var wa      = String(e && e.parameter && e.parameter.wa || '').replace(/\D/g, '');
  if (!usuario && !wa) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'usuario o wa requerido' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var sh = SS.getSheetByName('Vendedores');
  if (!sh || sh.getLastRow() <= 1) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'sin vendedores' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var rows = sh.getDataRange().getValues();
  for (var r = 1; r < rows.length; r++) {
    var estado = String(rows[r][6] || '').trim();
    if (estado !== 'Activo') continue;
    var u = String(rows[r][7] || '').trim().toLowerCase();      // col H = Usuario
    var w = String(rows[r][1] || '').replace(/\D/g, '');        // col B = WhatsApp
    if ((usuario && u === usuario) || (wa && w && w === wa)) {
      return ContentService.createTextOutput(JSON.stringify({
        ok: true,
        nombre: String(rows[r][0] || ''),
        usuario: u,
        wa: w,
        comision: Number(rows[r][9]) || 17
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'no encontrado' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function _doGetDashboardVendedor(e) {
  var nombre = String(e && e.parameter && e.parameter.nombre || '').trim();
  if (!nombre) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'nombre requerido' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Leer hoja Red (operativa)
  var sh = SS.getSheetByName('Red');
  var pedidos = [];
  var clientesMap = {}; // {nombre: {count, total, ultFecha}}

  // Fecha actual Argentina
  var ahora = new Date();
  var argDate = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var semanaActual = _isoWeek(argDate);
  var mesActual    = argDate.getMonth() + 1; // 1-12
  var yearActual   = argDate.getFullYear();
  var MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var nombreMesActual = MESES[mesActual - 1];

  var semanaFacturado = 0, semanaPedidos = 0;
  var mesFacturado = 0, mesPedidos = 0;
  var totalFacturado = 0, totalPedidos = 0;
  // Envío = 100% del vendedor. Su ganancia total = comisión 17% + envío.
  var semanaEnvio = 0, mesEnvio = 0, totalEnvio = 0;
  // Stats operativas viernes
  var viernesHoy = 0, viernesEntregados = 0; // pedidos de este viernes
  var pendCobrar = 0;       // plata sin cobrar al cliente
  var pendLiquidar = 0;     // plata cobrada pero no liquidada a Maleu
  // Breakdown por (año, mes) para desglose de meses anteriores
  var mesesMap = {}; // key: 'YYYY-MM' -> {n, y, m, facturado, pedidos}
  // Breakdown por (año, semana ISO) para desglose semanal
  var semanasMap = {}; // key: 'YYYY-WW' -> {y, w, facturado, pedidos}

  if (sh && sh.getLastRow() > 1) {
    // Leer cols A-BC (55): todo lo que el portal necesita
    var numRows = sh.getLastRow() - 1;
    var data = sh.getRange(2, 1, numRows, 63).getValues(); // hasta BK: incluye Tartas (BF-BI) y Wraps (BJ-BK)
    for (var r = 0; r < data.length; r++) {
      var vendedor = String(data[r][7] || '').trim();
      if (vendedor !== nombre) continue;

      var nPedido      = String(data[r][1] || '').trim();
      var fechaCell    = data[r][3];
      var cliente      = String(data[r][8] || '').trim();
      var estado       = String(data[r][11] || '').trim();
      // Día de Entrega (col K = idx 10): parseamos a Date y de ahí sacamos año/mes/semana.
      // Si no hay día de entrega válido, caemos a las cols E/F/G (mes/sem/año creación) como fallback.
      var diaRawDash = data[r][10];
      var fechaEntrega = null;
      var DIAS_DASH = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
      if (diaRawDash instanceof Date) {
        fechaEntrega = diaRawDash;
      } else {
        var mDash = String(diaRawDash || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (mDash) {
          fechaEntrega = new Date(Number(mDash[3]), Number(mDash[2]) - 1, Number(mDash[1]));
        }
      }
      var MESES_FULL = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
      var diaEntrega, mes, semana, year;
      if (fechaEntrega) {
        diaEntrega = DIAS_DASH[fechaEntrega.getDay()];
        mes        = MESES_FULL[fechaEntrega.getMonth()];
        semana     = _isoWeek(fechaEntrega);
        year       = fechaEntrega.getFullYear();
      } else {
        // Fallback a fecha de creación
        diaEntrega = '';
        mes        = String(data[r][4] || '').trim();
        semana     = Number(data[r][5]) || 0;
        year       = Number(data[r][6]) || 0;
      }
      var formaPago    = String(data[r][12] || '').trim();
      var estadoPago   = String(data[r][13] || '').trim();
      var total        = Number(data[r][14]) || 0;
      var envio        = Number(data[r][15]) || 0; // P = Envío (100% del vendedor)
      var efectivoMonto = Number(data[r][16]) || 0; // Q = Efectivo
      var transferMonto = Number(data[r][17]) || 0; // R = Transferencia
      var propinaEf    = Number(data[r][18]) || 0;
      var propinaTr    = Number(data[r][19]) || 0;
      var facturado    = Number(data[r][20]) || total;
      var comision     = Number(data[r][46]) || Math.round(facturado * 17 / 100);
      var aPagarMaleu  = Number(data[r][48]) || Math.round(facturado * 83 / 100); // AW (49) = A Pagar
      // Cols reales (1-based / 0-based): AX Barrio=50/49, AY Lote=51/50, AZ Teléfono=52/51,
      // BA Forma Pago Maleu=53/52, BB Estado Pago Maleu=54/53, BC Fecha Pago Maleu=55/54
      var barrio       = String(data[r][49] || '').trim();
      var lote         = String(data[r][50] || '').trim();
      var telefono     = String(data[r][51] || '').trim();
      var formaPagoMaleu  = String(data[r][52] || '').trim();
      var estadoPagoMaleu = String(data[r][53] || '').trim();
      var fechaPagoMaleu  = String(data[r][54] || '').trim();

      var fechaStr = '';
      if (fechaCell instanceof Date) {
        fechaStr = Utilities.formatDate(fechaCell, 'America/Argentina/Buenos_Aires', 'dd/MM');
      } else fechaStr = String(fechaCell || '');

      // Stats semana/mes/total (solo pedidos no cancelados)
      if (estado !== 'Cancelado') {
        totalFacturado += facturado; totalPedidos++; totalEnvio += envio;
        if (year === yearActual && semana === semanaActual) {
          semanaFacturado += facturado; semanaPedidos++; semanaEnvio += envio;
          if (diaEntrega === 'Viernes') {
            viernesHoy++;
            if (estado === 'Entregado') viernesEntregados++;
          }
        }
        if (year === yearActual && mes === nombreMesActual) {
          mesFacturado += facturado; mesPedidos++; mesEnvio += envio;
        }
        // Acumulado por (año, mes) para desglose. Solo si tenemos fechaEntrega válida.
        if (fechaEntrega && mes) {
          var mNum = fechaEntrega.getMonth() + 1; // 1-12
          var mKey = year + '-' + (mNum < 10 ? '0' + mNum : mNum);
          if (!mesesMap[mKey]) mesesMap[mKey] = { n: mes, y: year, m: mNum, facturado: 0, pedidos: 0, envio: 0 };
          mesesMap[mKey].facturado += facturado;
          mesesMap[mKey].pedidos++;
          mesesMap[mKey].envio += envio;
        }
        // Acumulado por (año, semana ISO) para desglose semanal
        if (fechaEntrega && semana) {
          var wKey = year + '-' + (semana < 10 ? '0' + semana : semana);
          if (!semanasMap[wKey]) semanasMap[wKey] = { y: year, w: semana, facturado: 0, pedidos: 0, envio: 0 };
          semanasMap[wKey].facturado += facturado;
          semanasMap[wKey].pedidos++;
          semanasMap[wKey].envio += envio;
        }
        // Plata pendiente de cobrar al cliente (entregado pero no cobrado).
        // El cliente paga productos + envío ($3.000, del vendedor), así que la
        // "plata por cobrar" real incluye el envío. Facturado (col U) = solo productos.
        if (estado === 'Entregado' && estadoPago !== 'Cobrado') {
          pendCobrar += facturado + envio;
        }
        // Plata cobrada pero no liquidada a Maleu (cobrado y no pagado a Maleu)
        // Marcos se queda con comisión 17%, el resto (83%) es para Maleu
        // Acepta "Pagado" (nuevo) o "Sí" (formato viejo) como liquidado
        var liquidado = (estadoPagoMaleu === 'Pagado' || estadoPagoMaleu === 'Sí' || estadoPagoMaleu === 'Si');
        if (estadoPago === 'Cobrado' && !liquidado) {
          pendLiquidar += aPagarMaleu;
        }
      }

      // Productos del pedido (cols V-AR = idx 21-43)
      var prods = [];
      var ABBRS = ['PPM','PPJyQ','PPCyQ','SQB','SL','SCo','SPyP','SJyQ','SE','SCa',
                   'ECaC','EJyQ','ECyQ','EV','TG','TLC','TC','F','PMu','PMa','PJyQ','PCC','PJyM'];
      for (var p = 0; p < 23; p++) {
        var qty = Number(data[r][21 + p]) || 0;
        if (qty > 0) prods.push({ a: ABBRS[p], q: qty });
      }
      // Tartas (BF-BI = idx 57-60) y Wraps (BJ-BK = idx 61-62): están al final de la
      // fila, fuera del bloque contiguo V-AR. Sin leerlas acá, el portal del vendedor
      // no las mostraba aunque el pedido las tuviera.
      var EXTRA = [[57,'TP'],[58,'TJyQ'],[59,'TCa'],[60,'TV'],[61,'RC'],[62,'RP']];
      for (var pe = 0; pe < EXTRA.length; pe++) {
        var qtyE = Number(data[r][EXTRA[pe][0]]) || 0;
        if (qtyE > 0) prods.push({ a: EXTRA[pe][1], q: qtyE });
      }

      // Calcular si es "este viernes" (mismo año+semana y día=Viernes)
      var esViernesActual = (year === yearActual && semana === semanaActual && diaEntrega === 'Viernes');

      pedidos.push({
        n: nPedido, c: cliente, f: fechaStr, de: diaEntrega,
        $: facturado, com: comision, env: envio, gan: comision + envio, aPg: aPagarMaleu, es: estado,
        fp: formaPago,       // Efectivo/Transferencia/Mixto
        ep: estadoPago,      // Cobrado/No Cobrado
        ef: efectivoMonto, tr: transferMonto, // cobrado al cliente (para Mixto)
        pEf: propinaEf, pTr: propinaTr,
        fpm: formaPagoMaleu, // Forma Pago a Maleu
        epm: estadoPagoMaleu,// Pagado/Pendiente
        fpmF: fechaPagoMaleu,// Fecha Pago a Maleu
        b: barrio, l: lote, t: telefono,
        prods: prods,
        vi: esViernesActual, // flag "es este viernes"
        _y: year,            // para sort
        _sem: semana         // para sort
      });

      // Clientes agrupados (enriquecido con barrio/teléfono del último pedido)
      if (cliente && estado !== 'Cancelado') {
        if (!clientesMap[cliente]) clientesMap[cliente] = { n: cliente, count: 0, total: 0, ult: fechaStr, b: barrio, t: telefono };
        clientesMap[cliente].count++;
        clientesMap[cliente].total += facturado;
        clientesMap[cliente].ult = fechaStr;
        if (barrio)   clientesMap[cliente].b = barrio;
        if (telefono) clientesMap[cliente].t = telefono;
      }
    }
  }

  // Sort pedidos por año + semana ISO + N° desc (más recientes primero)
  pedidos.sort(function(a, b) {
    if (a._y !== b._y) return b._y - a._y;
    if (a._sem !== b._sem) return b._sem - a._sem;
    var na = Number(a.n) || 0, nb = Number(b.n) || 0;
    return nb - na;
  });
  // Slicing: nunca podar pedidos abiertos. Un pedido es "abierto" si no está
  // Cancelado y todavía falta entregarlo o cobrarlo (al cliente). Esos
  // siempre van al payload — para que el portal no pierda los deudores al
  // cambiar de semana. Cerrados: top 20 más recientes.
  var _abierto = function(p) {
    return p.es !== 'Cancelado' && (p.es !== 'Entregado' || p.ep !== 'Cobrado');
  };
  var pAbiertos = pedidos.filter(_abierto);
  var pCerrados = pedidos.filter(function(p) { return !_abierto(p); });
  pedidos = pAbiertos.concat(pCerrados.slice(0, 20));

  // Clientes: convertir map a array, sort por count desc
  var clientes = Object.keys(clientesMap).map(function(k) { return clientesMap[k]; });
  clientes.sort(function(a, b) { return b.count - a.count; });

  // Rango de la semana ISO actual (lunes a domingo) en formato dd/MM
  var diaSem = argDate.getDay(); // 0=Dom, 1=Lun, ..., 6=Sab
  var offsetALun = (diaSem === 0) ? -6 : (1 - diaSem);
  var lunes = new Date(argDate);
  lunes.setDate(argDate.getDate() + offsetALun);
  var domingo = new Date(lunes);
  domingo.setDate(lunes.getDate() + 6);
  var fmt = function(d) {
    return Utilities.formatDate(d, 'America/Argentina/Buenos_Aires', 'dd/MM');
  };

  // Helper: semanas ISO cuyo jueves cae en (year, month). month 1-12.
  // Devuelve array de {n, y, lun, dom, facturado, pedidos, comision}
  function _semanasDelMes(y, mm) {
    var res = [];
    var d = new Date(y, mm - 1, 1);
    var lastDay = new Date(y, mm, 0);
    // Ir al primer jueves del mes (getDay: 0=Dom..4=Jue..6=Sab)
    while (d.getDay() !== 4) d.setDate(d.getDate() + 1);
    while (d <= lastDay) {
      var mon = new Date(d); mon.setDate(d.getDate() - 3);
      var dom = new Date(d); dom.setDate(d.getDate() + 3);
      var wNum = _isoWeek(d);
      var wYear = d.getFullYear(); // jueves está en este año — ISO year == calendar year en los jueves
      var key = wYear + '-' + (wNum < 10 ? '0' + wNum : wNum);
      var bucket = semanasMap[key] || { facturado: 0, pedidos: 0, envio: 0 };
      res.push({
        n: wNum, y: wYear,
        lun: fmt(mon), dom: fmt(dom),
        facturado: bucket.facturado,
        pedidos: bucket.pedidos,
        comision: Math.round(bucket.facturado * 17 / 100),
        envio: bucket.envio || 0,
        ganancia: Math.round(bucket.facturado * 17 / 100) + (bucket.envio || 0)
      });
      d.setDate(d.getDate() + 7);
    }
    return res;
  }

  // Meses anteriores (excluye el mes actual), ordenados desc por (año, mes), cada uno con desglose semanal
  var mesesAnt = [];
  Object.keys(mesesMap).forEach(function(k) {
    var row = mesesMap[k];
    if (row.y === yearActual && row.m === mesActual) return; // skip mes actual
    mesesAnt.push({
      n: row.n, y: row.y, m: row.m,
      facturado: row.facturado, pedidos: row.pedidos,
      comision: Math.round(row.facturado * 17 / 100),
      envio: row.envio || 0,
      ganancia: Math.round(row.facturado * 17 / 100) + (row.envio || 0),
      semanas: _semanasDelMes(row.y, row.m)
    });
  });
  mesesAnt.sort(function(a, b) {
    if (a.y !== b.y) return b.y - a.y;
    return b.m - a.m;
  });

  // Índice de pagos parciales del vendedor por (año, semana): { "YYYY-WW": { ef, tr, total, pagos:[{f,ef,tr}] } }
  var liqBySem = {};
  var shLiqD = SS.getSheetByName('Pagos Red Liq');
  if (shLiqD && shLiqD.getLastRow() > 1) {
    var dLiqD = shLiqD.getDataRange().getValues();
    for (var rL = 1; rL < dLiqD.length; rL++) {
      if (String(dLiqD[rL][2]).trim() !== nombre) continue;
      var yL = Number(dLiqD[rL][3]) || 0;
      var wL = Number(dLiqD[rL][4]) || 0;
      if (!yL || !wL) continue;
      var kL = yL + '-' + (wL < 10 ? '0' + wL : wL);
      if (!liqBySem[kL]) liqBySem[kL] = { ef: 0, tr: 0, total: 0, pagos: [] };
      var efL = Number(dLiqD[rL][5]) || 0;
      var trL = Number(dLiqD[rL][6]) || 0;
      liqBySem[kL].ef += efL;
      liqBySem[kL].tr += trL;
      liqBySem[kL].total += (efL + trL);
      var fRaw = dLiqD[rL][1];
      var fStr = '';
      if (fRaw instanceof Date) {
        fStr = Utilities.formatDate(fRaw, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
      } else {
        fStr = String(fRaw || '').trim();
      }
      liqBySem[kL].pagos.push({ f: fStr, ef: efL, tr: trL });
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    stats: {
      semana:   { n: semanaActual, lun: fmt(lunes), dom: fmt(domingo),
                  facturado: semanaFacturado, pedidos: semanaPedidos,
                  comision: Math.round(semanaFacturado * 17 / 100),
                  envio: semanaEnvio,
                  ganancia: Math.round(semanaFacturado * 17 / 100) + semanaEnvio },
      mes:      { facturado: mesFacturado, pedidos: mesPedidos, nombre: nombreMesActual,
                  y: yearActual, m: mesActual,
                  comision: Math.round(mesFacturado * 17 / 100),
                  envio: mesEnvio,
                  ganancia: Math.round(mesFacturado * 17 / 100) + mesEnvio,
                  semanas: _semanasDelMes(yearActual, mesActual) },
      mesesAnt: mesesAnt,
      total:    { facturado: totalFacturado,  pedidos: totalPedidos, envio: totalEnvio },
      comision: Math.round(totalFacturado * 17 / 100),
      ganancia: Math.round(totalFacturado * 17 / 100) + totalEnvio,
      viernes:      { total: viernesHoy, entregados: viernesEntregados },
      pendCobrar:   pendCobrar,
      pendLiquidar: pendLiquidar
    },
    pedidos: pedidos,
    clientes: clientes,
    liqBySem: liqBySem
  })).setMimeType(ContentService.MimeType.JSON);
}

function _doGetBusqueda() {
  var shOC = SS.getSheetByName('Orden de Compra');
  if (!shOC || shOC.getLastRow() <= 1) {
    return ContentService.createTextOutput(JSON.stringify({ ts: Date.now(), provs: [], clientes: [], total: 0 }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Parser robusto: acepta números crudos o strings "$8.000". Antes de este parche,
  // Number("$8.000") → NaN → 0 y costoT se calculaba como 0 para esas filas (bug 03/06/26).
  function _toMoney(v) {
    if (typeof v === 'number') return v;
    var s = String(v == null ? '' : v).replace(/[$\s]/g, '').replace(/\./g, '').replace(/,/g, '.');
    var n = Number(s);
    return isFinite(n) ? n : 0;
  }

  // Una sola lectura del rango usado (sin getDataRange que sobre-lee columnas vacías)
  var lastRow = shOC.getLastRow();
  var data = shOC.getRange(1, 1, lastRow, 25).getValues();

  // Mapa de display: nombres unificados para el WhatsApp del proveedor y las cards.
  // Hay variantes históricas en el Sheets ("Pack Pizza", "Pack Pizzas", "Pack Pizzas x2"
  // todas son lo mismo). Las agrupamos bajo "Pack Pizzas". Las pizzas individuales
  // de Bernardo Pisano van como "Pizzas".
  var CAT_DISPLAY = {
    'Tartas Gourmet': 'Tartas',
    'Tartas': 'Tartas',
    'Pizza Premium': 'Pizzas Individuales',
    'Pizzas Premium': 'Pizzas Individuales',
    'Pizza Individual': 'Pizzas Individuales',
    'Pizzas Individuales': 'Pizzas Individuales',
    'Pack Pizza': 'Pack Pizzas',
    'Pack Pizzas': 'Pack Pizzas',
    'Pack Pizzas x2': 'Pack Pizzas',
    'Packs Pizza': 'Pack Pizzas',
    'Pack Pizza x2': 'Pack Pizzas'
  };
  function splitProducto(nombre) {
    var parts = nombre.split(' — ');
    var catRaw = (parts[0] || nombre).trim();
    var cat = CAT_DISPLAY[catRaw] || catRaw;
    return { cat: cat, var: (parts[1] || '').trim() };
  }
  // Renombrar el `prod` completo para que tambien el detalle por OC se vea bien
  // ("Pack Pizza — Muzzarella" → "Pizzas — Muzzarella"). Se usa al guardar items
  // en porCliente/deudas/rows.
  function _renameProd(nombre) {
    var idx = nombre.indexOf(' — ');
    if (idx < 0) return nombre;
    var catRaw = nombre.substring(0, idx).trim();
    var sabor = nombre.substring(idx + 3);
    var cat = CAT_DISPLAY[catRaw] || catRaw;
    return cat + ' — ' + sabor;
  }

  // Semana ISO actual (lun-dom) en zona horaria AR — para filtrar OCs vivas
  var _nowArg = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var semanaActualBusq = _isoWeek(_nowArg);
  var anioActualBusq = _nowArg.getFullYear();

  // ── Pre-loop: leer Home/Pilar/Clubes/Red para mapear cada pedido a su Estado de Entrega
  // y semana ISO de entrega. Tadeo cuenta una venta por su fecha de entrega; agrupar
  // BUSQUEDA y ARMADO por la semana de entrega evita mezclar ciclos (ej: hoy ya recibido +
  // pedido nuevo para la semana siguiente).
  // Parser robusto: las cols pueden venir como Date o como texto "22/5/2026" / "22/05/2026".
  function _parseFechaCualquiera(raw) {
    if (raw instanceof Date) return raw;
    var s = String(raw || '').trim();
    if (!s) return null;
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return null;
  }
  var estadosEntrega = {}; // { 'Home|3': { est: 'Entregado', semEnt: 20, anioEnt: 2026 } }
  // colEntVend (Red): "Entregado a Vendedor" (col 57, idx 56). En Red, lo que importa
  // para que un pedido salga de ARMADO es que Maleu YA le entregó la mercadería al
  // vendedor (Marcos), NO que el cliente final la haya recibido (eso lo hace el vendedor).
  var hojasCanal = [
    { name: 'Home',   colId: 2, colEst: 11, colFEnt: 50 }, // Home: AX = Fecha Entrega (col 50, idx 49)
    { name: 'Pilar',  colId: 2, colEst: 11, colFEnt: 53 }, // Pilar: BA = Fecha Entrega (col 53, idx 52)
    { name: 'Clubes', colId: 2, colEst: 14, colFEnt: 13 }, // Clubes: M = Dia de Entrega Elegido (col 13, idx 12)
    { name: 'Red',    colId: 2, colEst: 12, colFEnt: 11, colEntVend: 57, colVend: 8 }  // Red: K=Dia Entrega, BE(57)=Entregado a Vendedor, H(8)=Vendedor
  ];
  hojasCanal.forEach(function(h) {
    var shCanal = SS.getSheetByName(h.name);
    if (!shCanal || shCanal.getLastRow() <= 1) return;
    var nCols = Math.max(h.colId, h.colEst, h.colFEnt, h.colEntVend || 0);
    var dataCanal = shCanal.getRange(2, 1, shCanal.getLastRow() - 1, nCols).getValues();
    for (var ri = 0; ri < dataCanal.length; ri++) {
      var idPed = String(dataCanal[ri][h.colId - 1] || '').trim();
      if (!idPed) continue;
      var estEnt = String(dataCanal[ri][h.colEst - 1] || '').trim();
      var entVend = h.colEntVend ? String(dataCanal[ri][h.colEntVend - 1] || '').trim() : '';
      // Red: capturar el vendedor (col H) para que ARMADO agrupe por vendedor real
      // (Marcos/Fini/Rufino) en vez de caer al default 'Marcos Bottcher' en el frontend.
      var vendPed = h.colVend ? String(dataCanal[ri][h.colVend - 1] || '').trim() : '';
      // Fecha entrega: parsear el campo correspondiente. Si esta vacio o no parsea,
      // caer a col D (fecha pedido = creacion).
      var fEntDate = _parseFechaCualquiera(dataCanal[ri][h.colFEnt - 1]);
      if (!fEntDate) fEntDate = _parseFechaCualquiera(dataCanal[ri][3]);
      var semEnt = 0, anioEnt = 0;
      if (fEntDate instanceof Date) {
        semEnt = _isoWeek(fEntDate);
        anioEnt = fEntDate.getFullYear();
      }
      estadosEntrega[h.name + '|' + idPed] = { est: estEnt, entVend: entVend, semEnt: semEnt, anioEnt: anioEnt, vendedor: vendPed };
    }
  });

  var porProv = {};
  var porCliente = {};
  var totalGeneral = 0;
  var ocRows = [];
  var deudas = {};

  // ── 1 SOLA PASADA: clasifica cada fila como "OC viva" o "deuda" o ambas ──
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var estado = String(row[20]).trim();
    if (estado !== 'Pendiente' && estado !== 'Pedido' && estado !== 'Recibido') continue;
    var origen = String(row[19]).trim();
    var esDeposito = origen.indexOf('Dep') === 0;
    var esRecibido = (estado === 'Recibido');

    var prov = String(row[9]).trim();
    var qty = Number(row[12]) || 0;
    if (!prov || qty === 0) continue;

    var costoT = _toMoney(row[13]) * qty;

    // Datos del cliente/canal para esta OC (los necesitamos antes del calculo
    // de deudas, porque la semana de la deuda DEBE ser la semana de entrega
    // del pedido al cliente, no la semana de creacion de la OC. Ejemplo del
    // bug: pedido Red creado sab 16/05 (sem 20) para entregar vie 22/05 (sem 21)
    // — la OC contra Sevuchitas debe sumar a la semana 21, no 20.)
    var producto = String(row[10]).trim();
    var abbr = String(row[11]).trim();
    var precioV = Number(row[15]) || 0;
    var ingresoT = precioV * qty;
    var cliente = String(row[6]).trim();
    var canal = String(row[4]).trim();
    var dir = String(row[8]).trim();
    var tel = String(row[7]).trim();
    var nPedido = String(row[5]).trim();
    var ocNum = String(row[0]).trim();

    // ── Normalizar canal Deposito ──
    // El Sheets tiene mezcla historica: canal 'Deposito' (sin tilde, viejo) y 'Depósito' (con tilde, actual),
    // y cliente 'Reposicion stock' vs 'Tadeo — Stock'. Los unificamos para que aparezcan como UNA sola card
    // en ARMADO/BUSQUEDA, agrupada por semana.
    var _canalLower = canal.toLowerCase();
    var esDepCanal = (_canalLower === 'deposito' || _canalLower === 'depósito');
    if (esDepCanal) {
      canal = 'Depósito';
      cliente = 'Depósito Maleu';
      nPedido = '';
    }

    // Lookup semana de entrega del pedido cliente. Si no la encontramos, fallback
    // a la semana de creacion de la OC (col C).
    var entKey = canal + '|' + nPedido;
    var entInfo = estadosEntrega[entKey];
    var semEntPed, anioEntPed;
    if (esDepCanal) {
      // Para Depósito no hay pedido cliente; usar siempre la semana/año de la OC.
      semEntPed = Number(row[2]) || semanaActualBusq;
      var fCreDep = row[1];
      anioEntPed = (fCreDep instanceof Date) ? fCreDep.getFullYear() : anioActualBusq;
    } else {
      semEntPed = (entInfo && entInfo.semEnt) ? entInfo.semEnt : (Number(row[2]) || semanaActualBusq);
      anioEntPed = (entInfo && entInfo.anioEnt) ? entInfo.anioEnt : anioActualBusq;
    }

    // ── DEUDAS: Recibido + no-Depósito + Pagado != Sí ──
    // La semana de la deuda es la semana en la que Tadeo RECIBIÓ la mercadería
    // (col W "Fecha Recibido"). Si está vacía, fallback a la semana de creación
    // de la OC (col C "Semana"). NO usamos semEntPed porque el pedido cliente
    // puede tener Día de Entrega muy viejo (zombi) y arruinar la imputación.
    if (esRecibido && !esDeposito) {
      var pagado = String(row[23]).trim();
      if (pagado !== 'Sí' && pagado !== 'Si' && costoT > 0) {
        var semRec = 0;
        var fRecRaw = row[22]; // col W "Fecha Recibido"
        var fRecDate = _parseFechaCualquiera(fRecRaw);
        if (fRecDate instanceof Date) semRec = _isoWeek(fRecDate);
        var semD = String(semRec || row[2] || '').trim();
        if (!deudas[prov]) deudas[prov] = { total: 0, semanas: {}, rows: [] };
        deudas[prov].total += costoT;
        if (!deudas[prov].semanas[semD]) deudas[prov].semanas[semD] = 0;
        deudas[prov].semanas[semD] += costoT;
        // Detalle por OC: canal/cliente/nped para que el cliente pueda mostrar
        // un breakdown legible ("Red / Marína Donati #24 — Empanadas CaC x1 $14.700").
        deudas[prov].rows.push({
          r: r + 1, prod: _renameProd(producto), abbr: abbr, q: qty,
          costoU: _toMoney(row[13]), costo: costoT,
          sem: semD, canal: canal, cliente: cliente, nped: nPedido
        });
      }
    }

    // ── Filas activas para porProv/porCliente/ocRows: solo no-Depósito ──
    if (esDeposito) continue;

    // Filtrar OCs cuya entrega NO sea esta semana ni la proxima.
    // Para canales con cliente real: atrasados se incluyen como "Esta semana" para no perderlos.
    // Para Depósito: NO traer atrasados RECIBIDOS — son cosas que Tadeo ya guardó.
    // PERO sí traer atrasados Pendiente/Pedido — son compras viejas que aún no fue a buscar.
    var diffSem = (anioEntPed - anioActualBusq) * 53 + (semEntPed - semanaActualBusq);
    if (diffSem > 1) continue;
    if (esDepCanal && diffSem < 0 && esRecibido) continue;

    // Pedidos entregados ya: no incluir en ARMADO (cliente). Sigue contando para BUSQUEDA si la OC esta Pendiente/Pedida.
    // Un pedido sale de ARMADO cuando ya fue entregado. Para Red la entrega
    // relevante es "Entregado a Vendedor" (Maleu → Marcos), NO el Estado de Entrega
    // al cliente final (eso lo gestiona el vendedor después). Para el resto de
    // canales, el Estado de Entrega del pedido es lo que vale.
    var yaEntregado;
    if (canal === 'Red') {
      yaEntregado = entInfo && (entInfo.entVend === 'Entregado' || entInfo.est === 'Entregado');
    } else {
      yaEntregado = entInfo && entInfo.est === 'Entregado';
    }

    // Proveedor agrupado (solo Pendiente/Pedido). Agrupacion incluye semana
    // de entrega: un proveedor con OCs para 2 semanas distintas tiene items separados.
    // qDep = cuántas unidades del item son para reposición de Depósito (cliente
    // "Tadeo — Stock"). Permite al usuario ver "17 × Muzz (10 cli + 7 depo)".
    if (!esRecibido) {
      if (!porProv[prov]) porProv[prov] = { costo: 0, cats: {}, semsEnt: {}, qDepTotal: 0 };
      var sp = splitProducto(producto);
      var arr = porProv[prov].cats[sp.cat] || (porProv[prov].cats[sp.cat] = []);
      var qDepThis = esDepCanal ? qty : 0;
      var found = false;
      for (var v = 0; v < arr.length; v++) {
        if (arr[v].v === sp.var && arr[v].sem === semEntPed && arr[v].anio === anioEntPed) {
          arr[v].q += qty; arr[v].ct += costoT; arr[v].qDep += qDepThis; found = true; break;
        }
      }
      if (!found) arr.push({ v: sp.var, q: qty, qDep: qDepThis, ct: costoT, sem: semEntPed, anio: anioEntPed });
      porProv[prov].costo += costoT;
      porProv[prov].qDepTotal += qDepThis;
      porProv[prov].semsEnt[semEntPed] = anioEntPed; // tracking de semanas activas
      totalGeneral += costoT;
    }

    // Cliente agrupado (solo si NO entregado: no debe aparecer en ARMADO).
    // Para Depósito agregamos la semana al cKey: cada semana tiene su propia card.
    if (!yaEntregado) {
      var cKey = cliente + '|' + nPedido + (esDepCanal ? ('|sem' + semEntPed + '-' + anioEntPed) : '');
      if (!porCliente[cKey]) porCliente[cKey] = { n: cliente, canal: canal, dir: dir, tel: tel, ped: nPedido, items: [], ing: 0, semEnt: semEntPed, anioEnt: anioEntPed, vendedor: (canal === 'Red' && entInfo ? (entInfo.vendedor || '') : '') };
      porCliente[cKey].items.push({ prod: _renameProd(producto), abbr: abbr, q: qty, pv: ingresoT, est: estado });
      porCliente[cKey].ing += ingresoT;
    }

    // _renameProd unifica variantes "Pack Pizza" / "Pack Pizzas" / "Pack Pizzas x2"
    // bajo un solo nombre, así el panel "Recibir mercadería" agrupa correctamente
    // por sabor en vez de tener múltiples entradas por la misma cosa.
    ocRows.push({ oc: ocNum, r: r + 1, prov: prov, prod: _renameProd(producto), abbr: abbr, q: qty, est: estado, canal: canal, cliente: cliente, nped: nPedido, semEnt: semEntPed });
  }

  // Programaciones de búsqueda por proveedor (hoja "Programaciones Búsqueda")
  var fProgMap = {};
  var shProg = SS.getSheetByName('Programaciones Búsqueda');
  if (shProg && shProg.getLastRow() > 1) {
    var progData = shProg.getRange(2, 1, shProg.getLastRow() - 1, 2).getValues();
    for (var rpg = 0; rpg < progData.length; rpg++) {
      var pName = String(progData[rpg][0]).trim();
      var pFec = progData[rpg][1];
      if (!pName) continue;
      if (pFec instanceof Date) {
        fProgMap[pName] = Utilities.formatDate(pFec, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
      } else {
        fProgMap[pName] = String(pFec || '').trim();
      }
    }
  }

  // Armar provsArr. Cada prov lleva tambien semsEnt (lista de semanas en las que
  // tiene OCs activas) — el cliente lo usa para mostrarlo solo en la pestaña
  // semanal correspondiente.
  var provsArr = [];
  Object.keys(porProv).forEach(function(prov) {
    var d = porProv[prov];
    var catsArr = [];
    Object.keys(d.cats).forEach(function(cat) { catsArr.push({ cat: cat, items: d.cats[cat] }); });
    var waLines = ['Hola! Te paso el pedido de esta semana:', ''];
    catsArr.forEach(function(c) {
      waLines.push(c.cat + ':');
      c.items.forEach(function(it) { waLines.push('  ' + it.q + ' × ' + (it.v || c.cat)); });
      waLines.push('');
    });
    waLines.push('Total: $' + d.costo.toLocaleString('es-AR'));
    waLines.push('');
    waLines.push('Gracias!');
    var semsEntArr = Object.keys(d.semsEnt).map(function(s){ return Number(s); }).sort(function(a,b){return a-b;});
    provsArr.push({ n: prov, costo: d.costo, cats: catsArr, wa: waLines.join('\n'), fProg: fProgMap[prov] || '', semsEnt: semsEntArr, qDepTotal: d.qDepTotal || 0 });
  });

  // clientesArr: ya esta filtrado en el loop (no incluye entregados).
  var clientesArr = Object.keys(porCliente).map(function(k) { return porCliente[k]; });

  // ── Pagos Proveedores (ledger) ──
  var pagosImputados = {};
  var pagosLibres = {};
  var shPagos = SS.getSheetByName('Pagos Proveedores');
  if (shPagos && shPagos.getLastRow() > 1) {
    var pagosCols = Math.max(8, shPagos.getLastColumn());
    var pagosData = shPagos.getRange(1, 1, shPagos.getLastRow(), pagosCols).getValues();
    for (var rp = 1; rp < pagosData.length; rp++) {
      var pp = String(pagosData[rp][1]).trim();
      if (!pp) continue;
      var ef = Number(pagosData[rp][2]) || 0;
      var mp = Number(pagosData[rp][3]) || 0;
      var bonif = Number(pagosData[rp][7]) || 0; // col H Bonificación
      var tot = Number(pagosData[rp][4]) || (ef + mp + bonif);
      if (tot <= 0) continue;
      var semImp = String(pagosData[rp][5] || '').trim();
      var notas = String(pagosData[rp][6] || '').trim();
      var fechaRaw = pagosData[rp][0];
      // Prefijo de dia (Lun/Mar/Mié/Jue/Vie/Sáb/Dom) ayuda a ubicar el pago rapido.
      var fechaStr;
      if (fechaRaw instanceof Date) {
        var diasCortos = ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'];
        var diaCorto = diasCortos[fechaRaw.getDay()];
        fechaStr = diaCorto + ' ' + Utilities.formatDate(fechaRaw, 'America/Argentina/Buenos_Aires', 'dd/MM HH:mm');
      } else {
        fechaStr = String(fechaRaw || '').trim();
      }
      var pagoObj = { fecha: fechaStr, ef: ef, mp: mp, bonif: bonif, tot: tot, notas: notas };
      if (semImp) {
        if (!pagosImputados[pp]) pagosImputados[pp] = {};
        if (!pagosImputados[pp][semImp]) pagosImputados[pp][semImp] = [];
        pagosImputados[pp][semImp].push(pagoObj);
      } else {
        if (!pagosLibres[pp]) pagosLibres[pp] = [];
        pagosLibres[pp].push(pagoObj);
      }
    }
  }

  var deudasArr = [];
  var allProvs = {};
  Object.keys(deudas).forEach(function(p) { allProvs[p] = true; });
  Object.keys(pagosImputados).forEach(function(p) { allProvs[p] = true; });
  Object.keys(pagosLibres).forEach(function(p) { allProvs[p] = true; });

  Object.keys(allProvs).forEach(function(prov) {
    var d = deudas[prov] || { total: 0, semanas: {}, rows: [] };
    var libres = pagosLibres[prov] || [];
    var imp = pagosImputados[prov] || {};
    var saldoLibre = libres.reduce(function(a, p) { return a + p.tot; }, 0);

    var semSet = {};
    Object.keys(d.semanas).forEach(function(s) { semSet[s] = true; });
    Object.keys(imp).forEach(function(s) { semSet[s] = true; });
    var semKeys = Object.keys(semSet).sort();

    // Indexar los items detallados por semana — el cliente los muestra en el
    // breakdown de cada card (qué fue lo que se pidió esa semana).
    var rowsBySem = {};
    (d.rows || []).forEach(function(rw){
      var sk = String(rw.sem || '').trim();
      if (!sk) return;
      if (!rowsBySem[sk]) rowsBySem[sk] = [];
      rowsBySem[sk].push(rw);
    });

    var semsArr = [];
    var totalOriginal = 0;
    var totalPagado = 0;
    var totalPendiente = 0;
    for (var i = 0; i < semKeys.length; i++) {
      var s = semKeys[i];
      var bruto = d.semanas[s] || 0;
      var pagosSem = imp[s] || [];
      var pagadoImp = pagosSem.reduce(function(a, p) { return a + p.tot; }, 0);
      var residuo = Math.max(0, bruto - pagadoImp);
      var aplicarLibre = Math.min(saldoLibre, residuo);
      saldoLibre -= aplicarLibre;
      var pagadoTotalSem = pagadoImp + aplicarLibre;
      var pendiente = Math.max(0, bruto - pagadoTotalSem);
      semsArr.push({
        sem: s, original: bruto, pagado: pagadoTotalSem, pendiente: pendiente,
        pagosImp: pagosSem, pagadoFifo: aplicarLibre,
        items: rowsBySem[s] || [] // detalle de OCs imputadas a esta semana
      });
      totalOriginal += bruto;
      totalPagado += pagadoTotalSem;
      totalPendiente += pendiente;
    }

    if (totalPendiente > 0.0001) {
      deudasArr.push({
        n: prov,
        total: totalPendiente,
        original: totalOriginal,
        pagado: totalPagado,
        semanas: semsArr,
        pagosLibres: libres,
        saldoLibreSobrante: saldoLibre,
        rows: d.rows
      });
    }
  });

  // Stocks por abreviatura (físico actual). Permite mostrar en ARMADO/Depósito
  // "14 × Muzz (8 actual + 14 = 22 final)" para que Tadeo vea cómo queda el freezer.
  var stocksProductos = {};
  var hProdSt = SS.getSheetByName('Productos');
  if (hProdSt && hProdSt.getLastRow() > 1) {
    var stData = hProdSt.getRange(2, 1, hProdSt.getLastRow() - 1, 8).getValues();
    for (var rs = 0; rs < stData.length; rs++) {
      var abr = String(stData[rs][2] || '').trim();
      if (!abr) continue;
      stocksProductos[abr] = Number(stData[rs][5]) || 0; // col F = Stock Físico
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      ts: Date.now(),
      provs: provsArr,
      clientes: clientesArr,
      total: totalGeneral,
      ocs: ocRows,
      deudas: deudasArr,
      semActual: semanaActualBusq,
      anioActual: anioActualBusq,
      stocksProductos: stocksProductos
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** POST action=pagarProveedor — registra pago a proveedor (ledger model).
 *  Recibe { action:'pagarProveedor', proveedor:'Sevuchitas', efectivo:3000, mp:2000, semana:'17', notas:'...' }
 *  - Crea fila en hoja "Pagos Proveedores" (ledger único de pagos parciales)
 *  - Crea 1-2 filas en Egresos (una por método con monto > 0)
 *  - Si viene `semana` → pago imputado a esa semana (col F). Sin semana → libre, FIFO desde más vieja.
 *  - Ya NO marca filas de OC como Pagado=Sí. Deuda = costos − pagos imputados − libres FIFO. */
function _doPostPagarProveedor(postData) {
  var proveedor = String(postData.proveedor || '').trim();
  var montoEf = Number(postData.efectivo) || 0;
  var montoMp = Number(postData.mp) || 0;
  // Bonificación del proveedor (ej. te regaló 1 J&Morrón). Salda deuda pero
  // NO es plata que sale de caja → no crea egreso. Total cubre ef+mp+bonif.
  var montoBonif = Number(postData.bonif) || 0;
  var notas = String(postData.notas || '').trim();
  var semanaImp = String(postData.semana || '').trim();
  var totalPago = montoEf + montoMp + montoBonif;

  if (!proveedor || totalPago <= 0) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'Sin monto' })).setMimeType(ContentService.MimeType.JSON);
  }

  var ahora = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var semana = _isoWeek(argNow);
  var mes = MESES[argNow.getMonth()];

  // ── Ledger: hoja "Pagos Proveedores" ──
  // Layout: A Fecha | B Proveedor | C Efectivo | D MP | E Total | F Semana Imputada | G Notas | H Bonificación
  var shPag = SS.getSheetByName('Pagos Proveedores');
  if (!shPag) {
    shPag = SS.insertSheet('Pagos Proveedores');
    shPag.getRange(1, 1, 1, 8).setValues([['Fecha', 'Proveedor', 'Efectivo', 'Mercado Pago', 'Total', 'Semana Imputada', 'Notas', 'Bonificación']]);
    shPag.setFrozenRows(1);
  } else if (shPag.getLastColumn() < 8) {
    // Migración suave: agregar columnas faltantes (Notas col 7 si era 6, Bonif col 8 si era 7)
    if (shPag.getLastColumn() < 7) shPag.getRange(1, 6, 1, 2).setValues([['Semana Imputada', 'Notas']]);
    shPag.getRange(1, 8).setValue('Bonificación');
  }
  shPag.appendRow([argNow, proveedor, montoEf, montoMp, totalPago, semanaImp, notas, montoBonif]);
  var rPag = shPag.getLastRow();
  shPag.getRange(rPag, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  shPag.getRange(rPag, 3, 1, 3).setNumberFormat('$#,##0');
  shPag.getRange(rPag, 8).setNumberFormat('$#,##0');

  // ── Egresos (impacto en caja) ──
  // Solo Efectivo y Mercado Pago salen de caja. La bonificación NO genera egreso.
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

  // El pago en EFECTIVO a un proveedor consume el/los sobre(s) que Tadeo dejó
  // preparado(s) para ese proveedor: los marca "Pagado" en la hoja Sobres y baja el
  // sub-saldo Sobres (col E) por el monto de esos sobres. Si no había sobre, no baja
  // (el gasto sale de la caja fuerte). El total de efectivo ya lo descuenta el Egreso.
  if (montoEf > 0) {
    var shSob = SS.getSheetByName('Sobres');
    var sobreConsumido = 0;
    if (shSob && shSob.getLastRow() > 1) {
      var normProv = _normNombre(proveedor);
      var dSob = shSob.getDataRange().getValues();
      for (var rsob = 1; rsob < dSob.length; rsob++) {
        if (String(dSob[rsob][3] || '').trim() !== 'Activo') continue;
        if (_normNombre(String(dSob[rsob][1] || '')) !== normProv) continue;
        shSob.getRange(rsob + 1, 4).setValue('Pagado');
        shSob.getRange(rsob + 1, 5).setValue(argNow);
        sobreConsumido += Number(dSob[rsob][2]) || 0;
      }
    }
    if (sobreConsumido > 0) _ajustarColSobres(-sobreConsumido);
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true, total: totalPago, bonif: montoBonif, filas: pagos.length })).setMimeType(ContentService.MimeType.JSON);
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

/** POST action=programarBusqueda — agenda fecha de búsqueda para un proveedor.
 *  Recibe { action:'programarBusqueda', proveedor:'Sevuchitas', fecha:'2026-04-30' (o '' para limpiar) }
 *  Hoja "Programaciones Búsqueda": Proveedor | Fecha (YYYY-MM-DD) | Actualizado
 *  Upsert: si ya hay fila para ese proveedor, la actualiza; sino la crea. */
function _doPostProgramarBusqueda(postData) {
  var prov = String(postData.proveedor || '').trim();
  var fecha = String(postData.fecha || '').trim();
  if (!prov) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'Sin proveedor' })).setMimeType(ContentService.MimeType.JSON);
  }
  var sh = SS.getSheetByName('Programaciones Búsqueda');
  if (!sh) {
    sh = SS.insertSheet('Programaciones Búsqueda');
    sh.getRange(1, 1, 1, 3).setValues([['Proveedor', 'Fecha Programada', 'Actualizado']]);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160);
    sh.setColumnWidth(2, 130);
    sh.setColumnWidth(3, 160);
  }
  var ahora = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var lastR = sh.getLastRow();
  var rowFound = -1;
  if (lastR > 1) {
    var data = sh.getRange(2, 1, lastR - 1, 1).getValues();
    for (var r = 0; r < data.length; r++) {
      if (String(data[r][0]).trim().toLowerCase() === prov.toLowerCase()) { rowFound = r + 2; break; }
    }
  }
  if (rowFound > 0) {
    sh.getRange(rowFound, 2, 1, 2).setValues([[fecha, argNow]]);
  } else if (fecha) {
    sh.appendRow([prov, fecha, argNow]);
    var nr = sh.getLastRow();
    sh.getRange(nr, 3).setNumberFormat('dd/MM/yyyy HH:mm');
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true, proveedor: prov, fecha: fecha })).setMimeType(ContentService.MimeType.JSON);
}

/** POST action=recibirMercaderia — marca OC como Recibido.
 *  Formato NUEVO (preferido, agrupa por producto y absorbe diff en "Tadeo — Stock"):
 *    { action:'recibirMercaderia',
 *      groups:[{ prov, prod, abbr, qtyRecibida, rows:[5,12,18] }] }
 *  Formato LEGACY (cada fila con su qty ya distribuida):
 *    { action:'recibirMercaderia', items:[{r:5, qtyRecibida:7}] }
 *
 *  Lógica con groups:
 *    - totalPedido = suma de cantidades en rows del grupo
 *    - diff = qtyRecibida - totalPedido
 *    - Si diff != 0 y existe fila con Cliente='Tadeo — Stock' (reposición): absorbe ahí
 *      (ajusta Cantidad, Costo Total, y vía fórmulas Ingreso/Margen).
 *    - Si no hay fila Stock y diff<0 → recorta desde el último row (aviso).
 *    - Si no hay fila Stock y diff>0 → suma el sobrante al stock físico (aviso).
 *    - Marca todas las filas como Recibido (con su nueva qty) y suma al stock físico
 *      solo las filas Canal=Depósito + Origen=Orden de Compra. */
function _doPostRecibirMercaderia(postData) {
  // ── IDEMPOTENCIA ──
  // Si el cliente reintenta (timeout en cliente pero el server YA procesó),
  // devolvemos la respuesta cacheada para que no duplique stock/marcas.
  // El frontend genera clientOpId al confirmar; mismo ID en retries.
  var coid = String(postData.clientOpId || '').trim();
  if (coid) {
    try {
      var _cacheIdem = CacheService.getScriptCache();
      var prev = _cacheIdem.get('rm_' + coid);
      if (prev) {
        // Ya procesado: devolver la respuesta original (no re-ejecutar)
        return ContentService.createTextOutput(prev).setMimeType(ContentService.MimeType.JSON);
      }
    } catch (_eIdem) { /* CacheService falla → seguir sin idempotencia (mejor un duplicado raro que perderse el cobro) */ }
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, retry: true, err: 'lock' })).setMimeType(ContentService.MimeType.JSON);
  }
  try {
  var shOC = SS.getSheetByName('Orden de Compra');
  if (!shOC) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'sheet' })).setMimeType(ContentService.MimeType.JSON);

  var hProd = SS.getSheetByName('Productos');

  var ahora = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var updated = 0;
  var avisos = [];

  // ── BATCH READ ──
  // Antes cargarInfo hacía 7 getRange().getValue() por fila = 7 round-trips por OC.
  // Con 20 OCs eso son 140 round-trips → ~5-10s solo en lectura.
  // Ahora leemos toda la hoja OC una sola vez e indexamos por fila.
  var _ocLastRow = shOC.getLastRow();
  var _ocData = _ocLastRow > 1 ? shOC.getRange(1, 1, _ocLastRow, 25).getValues() : [];
  // _ocData[r-1] tiene la fila r (1-indexed). Usamos esto en cargarInfo y para los reads costoU/precio en setRowQty.

  // Reconoce filas Depósito en cualquiera de sus variantes históricas/normalizadas:
  // "Tadeo — Stock", "Depósito Maleu", "Reposicion stock", o canal "Depósito".
  function isTadeoStock(cliente, canal) {
    var s = String(cliente || '').toLowerCase().replace(/\s+/g, ' ').trim();
    var ca = String(canal || '').toLowerCase();
    if (ca === 'depósito' || ca === 'deposito') return true;
    if (s.indexOf('tadeo') === 0 && s.indexOf('stock') !== -1) return true;
    if (s.indexOf('dep') === 0 && (s.indexOf('maleu') !== -1 || s === 'depósito' || s === 'deposito')) return true;
    if (s.indexOf('reposicion') !== -1) return true;
    return false;
  }

  function setRowQty(row, newQty) {
    // Lee costoU y precioV del batch en memoria — 0 round-trips.
    var rowData = _ocData[row - 1] || [];
    var costoUnit = Number(rowData[13]) || 0; // col N = 14 (idx 13)
    var precioV   = Number(rowData[15]) || 0; // col P = 16 (idx 15)
    shOC.getRange(row, 13).setValue(newQty);                // M Cantidad
    shOC.getRange(row, 15).setValue(costoUnit * newQty);    // O Costo Total
    // Q / R / S: si son fórmula, se auto-actualizan; si son valor, los escribimos.
    var rQ = shOC.getRange(row, 17);
    var rR = shOC.getRange(row, 18);
    var rS = shOC.getRange(row, 19);
    if (!rQ.getFormula()) rQ.setValue(precioV * newQty);
    if (!rR.getFormula()) rR.setValue((precioV * newQty) - (costoUnit * newQty));
    if (!rS.getFormula()) rS.setValue(precioV > 0 ? ((precioV - costoUnit) / precioV) : 0);
  }

  function markRecibido(row) {
    // Skip-flag para que _onEditOC NO duplique el +REC que este POST ya hace
    PropertiesService.getScriptProperties().setProperty('skip_onedit_OC_' + row, String(Date.now()));
    shOC.getRange(row, 21).setValue('Recibido');
    shOC.getRange(row, 23).setValue(argNow);
  }

  function markCancelado(row) {
    PropertiesService.getScriptProperties().setProperty('skip_onedit_OC_' + row, String(Date.now()));
    shOC.getRange(row, 21).setValue('Cancelado');
    shOC.getRange(row, 23).setValue(argNow);
    shOC.getRange(row, 13).setValue(0);
    shOC.getRange(row, 15).setValue(0);
  }

  // Cache de Productos en memoria: leemos UNA vez y reutilizamos.
  // Antes sumarStock leía hoja Productos entera por cada producto = N round-trips.
  var _prodData = null, _prodRowByAbbr = null;
  function _ensureProdCache() {
    if (_prodData !== null || !hProd) return;
    _prodData = hProd.getDataRange().getValues();
    _prodRowByAbbr = {};
    for (var rp = 1; rp < _prodData.length; rp++) {
      var ab = String(_prodData[rp][2] || '').trim();
      if (ab) _prodRowByAbbr[ab] = rp + 1; // row 1-indexed
    }
  }
  function sumarStock(abbr, qty, refOC, canal, origen) {
    if (!abbr || qty <= 0 || !hProd) return;
    if (canal.indexOf('Dep') !== 0) return;
    if (origen !== 'Orden de Compra') return;
    _ensureProdCache();
    var rowProd = _prodRowByAbbr[abbr];
    if (!rowProd) return;
    var celdaFis = hProd.getRange(rowProd, 6);
    var fisico = Number(_prodData[rowProd - 1][5]) || 0;
    var nuevo = fisico + qty;
    celdaFis.setValue(nuevo);
    _prodData[rowProd - 1][5] = nuevo; // sync cache para próximas llamadas del mismo POST
    _logKardex(abbr, '+REC', qty, fisico, nuevo, 'OC', refOC);
  }

  function cargarInfo(row) {
    // Antes: 7 getRange().getValue() = 7 round-trips. Ahora: 0 (lee del batch).
    var rd = _ocData[row - 1] || [];
    return {
      row: row,
      canal:   String(rd[4] || '').trim(),    // col E idx 4
      cliente: String(rd[6] || '').trim(),    // col G idx 6
      abbr:    String(rd[11] || '').trim(),   // col L idx 11
      qty:     Number(rd[12]) || 0,           // col M idx 12
      origen:  String(rd[19] || '').trim(),   // col T idx 19
      estado:  String(rd[20] || '').trim(),   // col U idx 20
      refOC:   String(rd[0] || '')            // col A idx 0
    };
  }

  // ── Formato NUEVO: groups ───────────────────────────────
  var groups = postData.groups || [];
  groups.forEach(function(g) {
    var rows = (g.rows || []).map(Number).filter(function(r){ return r > 0; });
    var qtyRecibida = Number(g.qtyRecibida);
    if (isNaN(qtyRecibida) || qtyRecibida < 0) qtyRecibida = 0;
    if (rows.length === 0) return;

    var infos = rows.map(cargarInfo);
    infos.forEach(function(i){ i.isStock = isTadeoStock(i.cliente, i.canal); });
    var pending = infos.filter(function(i){ return i.estado !== 'Recibido' && i.estado !== 'Cancelado'; });
    if (pending.length === 0) return;

    // ── Overrides manuales: el cliente eligió explícitamente cuánto va en cada row ──
    if (g.overrides && typeof g.overrides === 'object') {
      var hayOverride = false;
      for (var ko in g.overrides) { hayOverride = true; break; }
      if (hayOverride) {
        var prodLabelOv = g.prod || pending[0].abbr;
        pending.forEach(function(i) {
          var newQ = (g.overrides[i.row] !== undefined) ? Number(g.overrides[i.row]) : i.qty;
          if (isNaN(newQ) || newQ < 0) newQ = 0;
          if (newQ === 0) {
            markCancelado(i.row);
          } else {
            if (newQ !== i.qty) setRowQty(i.row, newQ);
            markRecibido(i.row);
            sumarStock(i.abbr, newQ, i.refOC, i.canal, i.origen);
          }
        });
        avisos.push({ prod: prodLabelOv, tipo: 'manual_override', diff: qtyRecibida - pending.reduce(function(s,i){ return s+i.qty; }, 0) });
        updated += pending.length;
        return;
      }
    }

    var totalPedido = pending.reduce(function(s, i){ return s + i.qty; }, 0);
    var diff = qtyRecibida - totalPedido;
    var prodLabel = g.prod || pending[0].abbr;

    // Caso 1 — cantidad 0 recibida: todo se cancela
    if (qtyRecibida === 0) {
      pending.forEach(function(i) { markCancelado(i.row); });
      avisos.push({ prod: prodLabel, tipo: 'todo_cancelado', diff: -totalPedido });
      updated += pending.length;
      return;
    }

    // Caso 2 — coincide exactamente: mark Recibido con qty original
    if (diff === 0) {
      pending.forEach(function(i) {
        markRecibido(i.row);
        sumarStock(i.abbr, i.qty, i.refOC, i.canal, i.origen);
      });
      updated += pending.length;
      return;
    }

    // Caso 3 — hay diff. Buscar fila "Tadeo — Stock" para absorber
    var stockRows = pending.filter(function(i){ return i.isStock; });
    var realRows  = pending.filter(function(i){ return !i.isStock; });
    var stockRow  = stockRows.length > 0 ? stockRows[stockRows.length - 1] : null; // última Stock

    if (stockRow) {
      var newStockQty = stockRow.qty + diff;
      if (newStockQty >= 0) {
        // Stock puede absorber todo el diff (positivo o negativo)
        if (newStockQty === 0) {
          markCancelado(stockRow.row);
        } else {
          setRowQty(stockRow.row, newStockQty);
          markRecibido(stockRow.row);
          sumarStock(stockRow.abbr, newStockQty, stockRow.refOC, stockRow.canal, stockRow.origen);
        }
        // Resto de filas (reales + otras stock) se reciben con qty original
        pending.forEach(function(i) {
          if (i === stockRow) return;
          markRecibido(i.row);
          sumarStock(i.abbr, i.qty, i.refOC, i.canal, i.origen);
        });
        avisos.push({
          prod: prodLabel,
          tipo: diff > 0 ? 'sobrante_absorbido_stock' : 'faltante_absorbido_stock',
          diff: diff,
          stockAntes: stockRow.qty,
          stockDespues: newStockQty
        });
        updated += pending.length;
        return;
      }

      // Stock no alcanza para absorber el faltante entero (diff muy negativo)
      // Cancela Stock y prorratea el resto desde el último real
      markCancelado(stockRow.row);
      var faltanteRest = diff + stockRow.qty; // aún negativo
      for (var k = realRows.length - 1; k >= 0 && faltanteRest < 0; k--) {
        var rr = realRows[k];
        var recorte = Math.min(rr.qty, -faltanteRest);
        var nuevaQ = rr.qty - recorte;
        if (nuevaQ <= 0) {
          markCancelado(rr.row);
        } else {
          setRowQty(rr.row, nuevaQ);
          markRecibido(rr.row);
          sumarStock(rr.abbr, nuevaQ, rr.refOC, rr.canal, rr.origen);
        }
        rr._absorbido = recorte;
        faltanteRest += recorte;
      }
      realRows.forEach(function(rr) {
        if (rr._absorbido !== undefined) return;
        markRecibido(rr.row);
        sumarStock(rr.abbr, rr.qty, rr.refOC, rr.canal, rr.origen);
      });
      avisos.push({
        prod: prodLabel,
        tipo: 'faltante_stock_insuficiente_clientes_afectados',
        diff: diff,
        stockAntes: stockRow.qty,
        stockDespues: 0
      });
      updated += pending.length;
      return;
    }

    // Sin fila Tadeo — Stock en el grupo
    if (diff > 0) {
      // Sobrante — reciben todos con qty original y el sobrante va directo al stock físico
      realRows.forEach(function(rr) {
        markRecibido(rr.row);
        sumarStock(rr.abbr, rr.qty, rr.refOC, rr.canal, rr.origen);
      });
      var sample = pending[0];
      sumarStock(sample.abbr, diff, sample.refOC, 'Dep\u00f3sito', 'Orden de Compra');
      avisos.push({ prod: prodLabel, tipo: 'sobrante_sin_stock_row', diff: diff });
    } else {
      // Faltante sin fila Stock: recorta desde el último real (FIFO inverso) + aviso
      var remaining = qtyRecibida;
      pending.forEach(function(rr) {
        var asignar = Math.min(remaining, rr.qty);
        if (asignar <= 0) {
          markCancelado(rr.row);
        } else {
          if (asignar < rr.qty) setRowQty(rr.row, asignar);
          markRecibido(rr.row);
          sumarStock(rr.abbr, asignar, rr.refOC, rr.canal, rr.origen);
        }
        remaining -= asignar;
      });
      avisos.push({ prod: prodLabel, tipo: 'faltante_sin_stock_row_clientes_afectados', diff: diff });
    }
    updated += pending.length;
  });

  // ── Formato LEGACY: items ───────────────────────────────
  var itemsLegacy = postData.items || [];
  if (groups.length === 0 && itemsLegacy.length > 0) {
    itemsLegacy.forEach(function(item) {
      var row = Number(item.r);
      var qtyRecibida = Number(item.qtyRecibida);
      if (!row) return;
      var info = cargarInfo(row);
      if (qtyRecibida <= 0) {
        markCancelado(row);
      } else {
        if (qtyRecibida !== info.qty) setRowQty(row, qtyRecibida);
        markRecibido(row);
        sumarStock(info.abbr, qtyRecibida, info.refOC, info.canal, info.origen);
      }
      updated++;
    });
  }

  SpreadsheetApp.flush();
  var _respPayload = JSON.stringify({ ok: true, updated: updated, avisos: avisos });
  // Guardar respuesta en cache para retries idempotentes (TTL 6h).
  if (coid) {
    try { CacheService.getScriptCache().put('rm_' + coid, _respPayload, 21600); } catch(_e) {}
  }
  return ContentService.createTextOutput(_respPayload).setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

/** Demanda Home por abreviatura en las últimas 3 semanas ISO COMPLETAS
 *  (excluye la semana en curso, que está incompleta y subestimaría).
 *  Base del sugeridor de compra del depósito en busqueda.html "+ NUEVO".
 *  El depósito responde a la venta de Home (Estancias/Río/Alcanfores).
 *  Devuelve { abbr: { avg: <prom semanal>, weeks: [w-3, w-2, w-1] } }. */
function _homeDemandByAbbr() {
  var out = {};
  var hHome = SS.getSheetByName('Home');

  // Acumuladores por abbr (siempre, aunque no haya hoja Home)
  Object.keys(HOME_COL_TO_ABBR).forEach(function(col){
    out[HOME_COL_TO_ABBR[col]] = { avg: 0, weeks: [0, 0, 0] };
  });
  if (!hHome) return out;

  // 3 semanas anteriores (orden viejo→nuevo): paso atrás de a 7 días desde hoy AR
  var ahora  = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var idxOf = {}; // "año-semana" → 0|1|2
  for (var k = 3; k >= 1; k--) {
    var dt = new Date(argNow.getTime() - k * 7 * 86400000);
    idxOf[dt.getFullYear() + '-' + _isoWeek(dt)] = 3 - k;
  }

  var data = hHome.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    var estado = String(data[r][10] || '').trim();   // K = Estado Entrega
    if (estado === 'Cancelado') continue;
    var semana = Number(data[r][5]);                  // F = Semana ISO
    var anio   = Number(data[r][6]);                  // G = Año
    var slot   = idxOf[anio + '-' + semana];
    if (slot === undefined) continue;
    for (var col in HOME_COL_TO_ABBR) {
      var q = Number(data[r][col - 1]) || 0;          // HOME_COL_TO_ABBR es 1-based
      if (q) out[HOME_COL_TO_ABBR[col]].weeks[slot] += q;
    }
  }

  Object.keys(out).forEach(function(ab){
    var w = out[ab].weeks;
    out[ab].avg = Math.round(((w[0] + w[1] + w[2]) / 3) * 10) / 10;
  });
  return out;
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
      stockMap[ab] = Number(prodData[r][7]) || 0; // col H = Stock Disponible (físico - reservado)
    }
  }

  // Demanda Home (3 sem) para el sugeridor de compra del depósito
  var demMap = _homeDemandByAbbr();

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
      cat: lastProd,                                           // categor\u00eda = producto base (Pizza, Empanadas, etc.)
      c: costoMap[abbr] || 0,
      s: stockMap[abbr] !== undefined ? stockMap[abbr] : null,
      dem: demMap[abbr] ? demMap[abbr].avg : 0,                // prom semanal Home (3 sem)
      wk:  demMap[abbr] ? demMap[abbr].weeks : [0, 0, 0]       // desglose [w-3, w-2, w-1]
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ts: Date.now(), proveedores: proveedores, productos: productosPorProv }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Dirección de retiro/pasamanos para un vendedor Red (col I "Dirección" de OC).
 *  Toma el primer barrio de la hoja Vendedores; fallback a localidad. Genérico
 *  para N vendedores (Marcos, Fini, …) — no hardcodea ningún nombre. */
function _dirVendedorRed(nombre) {
  try {
    var sh = SS.getSheetByName('Vendedores');
    if (sh) {
      var vals = sh.getDataRange().getValues();
      for (var i = 1; i < vals.length; i++) {
        if (String(vals[i][0]).trim() === String(nombre).trim()) {
          var barrios = String(vals[i][3] || '').split(',').map(function(s){ return s.trim(); }).filter(Boolean);
          var localidad = String(vals[i][5] || '').trim();
          return barrios[0] || localidad || 'Retira vendedor';
        }
      }
    }
  } catch (e) {}
  return 'Retira vendedor';
}

/** POST action=compraManual — crea filas en Orden de Compra desde busqueda.html.
 *  Replica la lógica de confirmarCompraDeposito pero vía API web.
 *  Recibe { action:'compraManual', vendedor:'Tadeo — Stock', proveedor:'Sevuchitas', items:[{abbr,nombre,costo,qty,origen}] } */
function _doPostCompraManual(postData) {
  var shOC = SS.getSheetByName('Orden de Compra');
  if (!shOC) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Hoja OC no encontrada' })).setMimeType(ContentService.MimeType.JSON);

  var proveedor = String(postData.proveedor || '').trim();
  var vendedor = String(postData.vendedor || 'Tadeo — Stock').trim();
  var items = postData.items || [];
  // Es compra Red si va dirigida a un vendedor (cualquiera) y no al stock propio.
  var esVendedorRed = vendedor !== 'Tadeo — Stock' && vendedor !== '';

  var itemsValidos = items.filter(function(it) { return (Number(it.qty) || 0) > 0; });
  if (itemsValidos.length === 0) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Sin productos' })).setMimeType(ContentService.MimeType.JSON);

  // Productos: leer UNA vez, indexar por abbr (precio venta + fisico + nombre + rowIdx)
  var prodMap = {};
  var hProd = SS.getSheetByName('Productos');
  var prodData = hProd ? hProd.getDataRange().getValues() : [];
  for (var p = 1; p < prodData.length; p++) {
    var ab = String(prodData[p][2]).trim();
    if (!ab) continue;
    prodMap[ab] = {
      rowIdx: p + 1,
      nombre: String(prodData[p][1]).trim(),
      fisico: Number(prodData[p][5]) || 0,
      precioVenta: parseFloat(String(prodData[p][8]).replace(/[$.]/g,'').replace(/,/g,'')) || 0
    };
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
  var fechaSoloDia = dd + '/' + mm + '/' + yyyy;
  var semana = _isoWeek(argDate);
  var mesNombre = MESES[argDate.getMonth()];

  // Reservar IDs en bulk: 1 escaneo + 1 escritura a Config (en vez de N)
  var nIds = itemsValidos.length;
  var maxOC = 0;
  if (shOC.getLastRow() > 1) {
    var ocIds = shOC.getRange(2, 1, shOC.getLastRow() - 1, 1).getValues();
    var rxOC = /^OC-(\d+)$/;
    for (var ii = 0; ii < ocIds.length; ii++) {
      var m = String(ocIds[ii][0]).match(rxOC);
      if (m) { var nn = parseInt(m[1], 10); if (nn > maxOC) maxOC = nn; }
    }
  }
  var shConfig = SS.getSheetByName('Config');
  if (shConfig) {
    var configVal = Number(shConfig.getRange(4, 2).getValue()) || 0;
    if (configVal > maxOC) maxOC = configVal;
  }
  var firstId = maxOC + 1;
  if (shConfig) shConfig.getRange(4, 2).setValue(firstId + nIds - 1);

  // Construir filas + cambios de stock + filas Kardex
  var newRows = [];
  var stockUpdates = [];
  var kardexRows = [];
  var canalLabel = esVendedorRed ? 'Red' : 'Depósito';
  var dirLabel = esVendedorRed ? _dirVendedorRed(vendedor) : 'Depósito Maleu';

  for (var idx = 0; idx < itemsValidos.length; idx++) {
    var item = itemsValidos[idx];
    var ocId = 'OC-' + String(firstId + idx).padStart(3, '0');
    var costo = Number(item.costo) || 0;
    var qty = Number(item.qty) || 0;
    var costoTotal = costo * qty;
    var origenItem = String(item.origen || 'Orden de Compra');
    var esDeposito = origenItem.indexOf('Dep') === 0;
    var info = prodMap[item.abbr] || {};

    newRows.push([
      ocId,
      fechaStr,
      semana,
      mesNombre,
      canalLabel,
      '',
      vendedor,
      '',
      dirLabel,
      proveedor,
      String(item.nombre || info.nombre || ''),
      String(item.abbr || ''),
      qty,
      costo,
      costoTotal,
      info.precioVenta || 0,
      0, 0, 0,
      origenItem,
      esDeposito ? 'Recibido' : 'Pedido',
      fechaSoloDia,
      esDeposito ? fechaSoloDia : '',
      esDeposito ? 'Sí' : 'No',
      'No'
    ]);

    if (esDeposito && item.abbr && qty > 0 && info.rowIdx) {
      var fisico = info.fisico;
      var nuevoStock = Math.max(0, fisico - qty);
      info.fisico = nuevoStock;
      stockUpdates.push({ rowIdx: info.rowIdx, nuevoStock: nuevoStock });
      kardexRows.push([
        fechaStr,
        info.nombre || item.abbr,
        item.abbr,
        '-DEP',
        qty,
        fisico,
        nuevoStock,
        canalLabel,
        ocId
      ]);
    }
  }

  // Escribir todas las filas en bloque
  var startRow = shOC.getLastRow() + 1;
  shOC.getRange(startRow, 1, newRows.length, 25).setValues(newRows);

  // Formulas financieras: setFormulas batch (1 call por columna en vez de 3*N)
  var formQ = [], formR = [], formS = [];
  for (var i = 0; i < newRows.length; i++) {
    var r = startRow + i;
    formQ.push(['=P' + r + '*M' + r]);
    formR.push(['=Q' + r + '-O' + r]);
    formS.push(['=R' + r + '/Q' + r]);
  }
  shOC.getRange(startRow, 17, newRows.length, 1).setFormulas(formQ);
  shOC.getRange(startRow, 18, newRows.length, 1).setFormulas(formR);
  shOC.getRange(startRow, 19, newRows.length, 1).setFormulas(formS);

  // Formato
  shOC.getRange(startRow, 13, newRows.length, 1).setNumberFormat('0');
  shOC.getRange(startRow, 14, newRows.length, 2).setNumberFormat('$#,##0');
  shOC.getRange(startRow, 16, newRows.length, 3).setNumberFormat('$#,##0');
  shOC.getRange(startRow, 19, newRows.length, 1).setNumberFormat('0.0%');

  // Stock updates
  if (stockUpdates.length > 0 && hProd) {
    stockUpdates.forEach(function(u) { hProd.getRange(u.rowIdx, 6).setValue(u.nuevoStock); });
  }

  // Kardex: append batch (en vez de N appendRow secuenciales)
  if (kardexRows.length > 0) {
    var shKardex = SS.getSheetByName('Kardex');
    if (shKardex) {
      var kStart = shKardex.getLastRow() + 1;
      shKardex.getRange(kStart, 1, kardexRows.length, 9).setValues(kardexRows);
      shKardex.getRange(kStart, 4, kardexRows.length, 1).setNumberFormat('@');
    }
  }

  var totalCosto = newRows.reduce(function(s, row) { return s + (row[14] || 0); }, 0);
  return ContentService.createTextOutput(JSON.stringify({
    ok: true, filas: newRows.length, totalCosto: totalCosto,
    primerOC: newRows[0][0], ultimoOC: newRows[newRows.length - 1][0]
  })).setMimeType(ContentService.MimeType.JSON);
}

/** POST action=marcarEntregado — marca un pedido como Entregado desde ruta.html.
 *  Replica la lógica de _onEditHome: fecha de entrega + descuento de stock si Deposito. */
function _doPostMarcarEntregado(data) {
  var hoja = String(data.hoja || '');
  var pedidoId = String(data.id || '');
  var cobrado = !!data.cobrado;

  var sh = SS.getSheetByName(hoja);
  if (!sh) return ContentService.createTextOutput(JSON.stringify({ ok: false })).setMimeType(ContentService.MimeType.JSON);

  // Buscar fila por N° Pedido (col B=2). Usar row hint si coincide, sino buscar último match.
  var row = Number(data.row) || 0;
  if (row > 1 && String(sh.getRange(row, 2).getValue()) === pedidoId) {
    // row hint es correcto
  } else {
    var allData = sh.getDataRange().getValues();
    row = -1;
    // Buscar último match (en caso de N° reseteados semanalmente)
    for (var r = allData.length - 1; r >= 1; r--) {
      if (String(allData[r][1]) === pedidoId) { row = r + 1; break; }
    }
    if (row === -1) return ContentService.createTextOutput(JSON.stringify({ ok: false })).setMimeType(ContentService.MimeType.JSON);
  }

  // Columnas según hoja (Clubes y Red tienen layout diferente)
  // Home/Pilar v2 (abr/2026): I=9 Origen, K=11 EstEntrega, L=12 FormaPago, M=13 EstPago,
  //                           R=18 Efectivo, S=19 Transferencia, T=20 PropEf, U=21 PropTr
  var isClub = (hoja === 'Clubes');
  var isRed  = (hoja === 'Red');
  var colEstadoEntrega = isClub ? 14 : isRed ? 12 : 11;   // N / L / K
  var colOrigen        = isClub ? 12 : isRed ? 10 : 9;    // L / J / I
  var colEstadoPago    = isClub ? 16 : isRed ? 14 : 13;   // P / N / M
  var colFormaPago     = isClub ? 15 : isRed ? 13 : 12;   // O / M / L
  var colPropinaEf     = isClub ? 21 : isRed ? 19 : 20;   // U / S / T
  var colPropinaTr     = isClub ? 22 : isRed ? 20 : 21;   // V / T / U
  // Col Repartidor (agregada 05/05/2026): última col de cada hoja
  // Home: BD=56 · Pilar: BG=59 · Clubes: AK=37
  var colRepartidor    = (hoja==='Home'?56 : hoja==='Pilar'?59 : isClub?37 : 0);

  // Verificar que no esté ya Entregado (idempotente)
  var estadoActual = String(sh.getRange(row, colEstadoEntrega).getValue()).trim();
  if (estadoActual === 'Entregado') {
    // Si el POST trae cobrado=true y el pedido aún no está cobrado, al menos marcar pago.
    if (cobrado && String(sh.getRange(row, colEstadoPago).getValue()).trim() !== 'Cobrado') {
      sh.getRange(row, colEstadoPago).setValue('Cobrado');
      _stampFechaCobro(sh, row);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true, ya: true })).setMimeType(ContentService.MimeType.JSON);
  }

  // Flag skip-onedit para que el trigger onEdit NO duplique fecha/stock.
  // El POST va a hacer TODO el trabajo (más confiable que depender del trigger).
  var props = PropertiesService.getScriptProperties();
  var skipKey = 'skip_onedit_' + hoja + '_' + row;
  props.setProperty(skipKey, String(Date.now()));

  try {
    // Marcar Entregado (el trigger se dispara pero saldrá por el flag)
    sh.getRange(row, colEstadoEntrega).setValue('Entregado');
    SpreadsheetApp.flush();  // asegurar que el setValue persista antes de seguir

    // Registrar fecha de entrega (solo Home/Pilar/CF)
    if (!isClub && !isRed) {
      var colEntrega = hoja === 'Pilar' ? 51 : 48; // Home (AV=48), Pilar (AY=51)
      _registrarFechaEntrega(sh, row, colEntrega);
    }

    // Descontar stock según Origen (con skip-flag el trigger onEdit no lo hará, así que lo hace el POST)
    var origen = String(sh.getRange(row, colOrigen).getValue()).trim();
    if (origen === 'Deposito') {
      var hProd = SS.getSheetByName('Productos');
      if (hProd) {
        if (isClub)      _clubesStockFisico(sh, row, hProd, -1);
        else if (isRed)  _redStockFisico(sh, row, hProd, -1);
        else             _homeStockFisico(sh, row, hProd, -1); // Home/Pilar
      }
    } else if (origen === 'Mixto') {
      var hProd2 = SS.getSheetByName('Productos');
      if (hProd2) {
        if (isClub)      _clubesStockFisicoMixto(sh, row, hProd2, -1);
        else if (isRed)  _redStockFisicoMixto(sh, row, hProd2, -1);
        else             _homeStockFisicoMixto(sh, row, hProd2, -1);
      }
    }

    // Marcar cobrado si se indicó + estampar fecha (idempotente)
    if (cobrado) {
      if (String(sh.getRange(row, colEstadoPago).getValue()).trim() !== 'Cobrado') {
        sh.getRange(row, colEstadoPago).setValue('Cobrado');
        _stampFechaCobro(sh, row);
      }
    }

    // Escribir propina si existe. Puede ser negativa (ajuste de redondeo:
    // ej. cobre $10.550 de un pedido de $10.580 → propina = -30).
    var propina = Number(data.propina) || 0;
    if (propina !== 0) {
      var formaPago = String(sh.getRange(row, colFormaPago).getValue()).trim();
      if (formaPago === 'Efectivo') sh.getRange(row, colPropinaEf).setValue(propina);
      else                          sh.getRange(row, colPropinaTr).setValue(propina);
    }

    // Registrar nombre del repartidor (quién marcó la entrega)
    var rep = String(data.repartidor || '').trim();
    if (rep && colRepartidor > 0) {
      sh.getRange(row, colRepartidor).setValue(rep);
    }

    SpreadsheetApp.flush();
    return ContentService.createTextOutput(JSON.stringify({ ok: true, row: row })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    // Limpiar flag aunque haya excepción
    props.deleteProperty(skipKey);
  }
}

// ════════════════════════════════════════════════════════════
//  doPost — recibe pedidos desde la página + acciones internas
// ════════════════════════════════════════════════════════════
function doPost(e) {
  const lock = LockService.getScriptLock();
  // Timeout generoso (30s). Si no consigue el lock, respondemos ok:false con err para que el cliente reintente.
  if (!lock.tryLock(30000)) {
    _logPedidoFallido(e, new Error('LockTimeout 30s'));
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'LockTimeout', retry: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'compra')          return _doPostCompra(data);
    if (data.action === 'compraLote')     return _doPostCompraLote(data);
    if (data.action === 'updateCompra')   return _doUpdateCompra(data);
    if (data.action === 'marcarEntregado') return _doPostMarcarEntregado(data);
    if (data.action === 'gasto')           return _doPostGasto(data);
    if (data.action === 'ingreso')         return _doPostIngreso(data);
    if (data.action === 'ajusteSaldo')     return _doPostAjusteSaldo(data);
    if (data.action === 'moverBilletera')  return _doPostMoverBilletera(data);
    if (data.action === 'moverEfectivo')   return _doPostMoverEfectivo(data);
    if (data.action === 'invertir')        return _doPostInvertir(data);
    if (data.action === 'prepararSobre')   return _doPostPrepararSobre(data);
    if (data.action === 'eliminarSobre')   return _doPostEliminarSobre(data);
    if (data.action === 'marcarCobrado')   return _doPostMarcarCobrado(data);
    if (data.action === 'descobrar')       return _doPostDescobrar(data);
    if (data.action === 'cobrarParcial')   return _doPostCobrarParcial(data);
    if (data.action === 'pagarVendedor')   return _doPostPagarVendedor(data);
    if (data.action === 'cancelarPedido')  return _doPostCancelarPedido(data);
    if (data.action === 'cobrarVendedorRed') return _doPostCobrarVendedorRed(data);
    if (data.action === 'setOrigenProductos') return _doPostSetOrigenProductos(data);
    if (data.action === 'deshacerOrigen')     return _doPostDeshacerOrigen(data);
    if (data.action === 'cambiarOrigen')     return _doPostCambiarOrigen(data);
    if (data.action === 'cambiarEstadoEntrega') return _doPostCambiarEstadoEntrega(data);
    if (data.action === 'updateDiaEntrega')  return _doPostUpdateDiaEntrega(data);
    if (data.action === 'marcarOC')        return _doPostMarcarOC(data);
    if (data.action === 'recibirMercaderia') return _doPostRecibirMercaderia(data);
    if (data.action === 'pagarProveedor')  return _doPostPagarProveedor(data);
    if (data.action === 'compraManual')    return _doPostCompraManual(data);
    if (data.action === 'loginUsuario')    return _doPostLoginUsuario(data);
    if (data.action === 'loginVendedor')   return _doPostLoginVendedor(data);
    if (data.action === 'updatePedidoRed') return _doPostUpdatePedidoRed(data);
    if (data.action === 'editarPedidoRed') return _doPostEditarPedidoRed(data);
    if (data.action === 'editarPedido')    return _doPostEditarPedido(data);
    if (data.action === 'marcarSemanaPagadaRed') return _doPostMarcarSemanaPagadaRed(data);
    if (data.action === 'crmUpdateClienteMeta') return _doPostCrmUpdateClienteMeta(data);
    if (data.action === 'guardarCumpleCliente') return _doPostGuardarCumpleCliente(data);
    if (data.action === 'usarCupon')            return _doPostUsarCupon(data);
    if (data.action === 'crmMergeClientes')     return _doPostCrmMergeClientes(data);
    if (data.action === 'crmZonaSave')          return _doPostCrmZonaSave(data);
    if (data.action === 'crmZonaDelete')        return _doPostCrmZonaDelete(data);
    if (data.action === 'crmLoteSave')          return _doPostCrmLoteSave(data);
    if (data.action === 'crmSetVisitante')      return _doPostCrmSetVisitante(data);
    if (data.action === 'crmSetResidencia')     return _doPostCrmSetResidencia(data);
    if (data.action === 'hogaresMergeSave')     return _doPostHogaresMerge(data);
    if (data.action === 'crmPuntoSave')         return _doPostCrmPuntoSave(data);
    if (data.action === 'crmLogInteraccion')    return _doPostCrmLogInteraccion(data);
    if (data.action === 'crmDeleteInteraccion') return _doPostCrmDeleteInteraccion(data);
    if (data.action === 'crmLogCampania')       return _doPostCrmLogCampania(data);
    if (data.action === 'programarBusqueda')    return _doPostProgramarBusqueda(data);
    if (data.action === 'cerrarDiaReparto')     return _doPostCerrarDiaReparto(data);
    if (data.action === 'marcarEntregadoAVendedor') return _doPostMarcarEntregadoAVendedor(data);
    if (data.action === 'marcarGuardadoEnStock') return _doPostMarcarGuardadoEnStock(data);
    if (data.action === 'crearSaldoCliente') return _doPostCrearSaldoCliente(data);
    if (data.action === 'planMetaSet')     return _doPostPlanMetaSet(data);
    if (data.action === 'planAccionAdd')   return _doPostPlanAccionAdd(data);
    if (data.action === 'planAccionUpdate')return _doPostPlanAccionUpdate(data);
    if (data.action === 'planAccionDelete')return _doPostPlanAccionDelete(data);
    if (data.action === 'usuarioSet')      return _doPostUsuarioSet(data);
    if (data.action === 'usuarioDelete')   return _doPostUsuarioDelete(data);
    if (data.action === 'permisoSet')      return _doPostPermisoSet(data);
    if (data.action === 'configMaleuSet')  return _doPostConfigMaleuSet(data);
    if (data.action === 'provisionSet')    return _doPostProvisionSet(data);
    if (data.action === 'provisionDelete') return _doPostProvisionDelete(data);
    if (data.action === 'configOpSet')     return _doPostConfigOpSet(data);
    if (data.action === 'vendedorSet')     return _doPostVendedorSet(data);
    return _doPostPedido(data);
  } catch(err) {
    // LOG DE ERROR: guardar pedido fallido para no perderlo jamás
    _logPedidoFallido(e, err);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
    // Invalidar cache de entregas tras cualquier POST: no servir datos stale despues de escribir.
    _invalidateEntregasCache();
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
    // ── Dedupe por clientOrderId ──
    // El frontend genera un ID único cuando arma el pedido. Si el cliente tiene mala
    // señal, el POST puede llegar al server pero la respuesta no volver al navegador →
    // el cliente reintenta y crearía un duplicado. CacheService con TTL 6h descarta
    // POSTs repetidos. Suficiente porque los retries de cola persisten minutos, no horas.
    var coid = String(data.clientOrderId || '').trim();
    if (coid) {
      try {
        var _cache = CacheService.getScriptCache();
        if (_cache.get('coid_' + coid)) {
          return ContentService
            .createTextOutput(JSON.stringify({ ok: true, dedup: true }))
            .setMimeType(ContentService.MimeType.JSON);
        }
        _cache.put('coid_' + coid, '1', 21600); // 6h
      } catch (_e) { /* si CacheService falla, dejamos pasar — mejor un duplicado que un pedido perdido */ }
    }

    const canal = String(data.canal || 'Home');
    if (canal === 'Clubes')          _doPostClubes(data);
    else if (canal === 'Red')        _doPostRed(data);
    else if (canal === 'Pilar')      _doPostHome(data, 'Pilar', 'P');
    else                             _doPostHome(data, 'Home', 'H');

    // Cumpleaños capturado en el form de la tienda → guardar en Clientes Meta
    // (no destructivo). Falla silenciosa: nunca debe tumbar un pedido.
    if (data.cumple) {
      try { _guardarCumpleCliente(data.telefono, data.cumple); } catch (_e) {}
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  _doPostHome — escribe en la hoja "Home" (canal Estancias)
//  Columnas A–AP según el esquema de análisis de Tadeo
// ════════════════════════════════════════════════════════════

// Mapeo id de producto (página web) → columna 1-based de la hoja Home
// Home v2 abr/2026: productos empiezan en W (23) hasta AO (41)
// 15/05/2026: tartas en BE-BH (57-60), al final tras cols de control (no consecutivas).
const HOME_PRODUCT_COLS = {
  5:  23,  // W  — PPM   — Pack Muzarella x2
  6:  24,  // X  — PPJyQ — Pack Jamón y Queso x2
  7:  25,  // Y  — PPCyQ — Pack Cebolla y Queso x2
  8:  26,  // Z  — SCo   — Sorrentinos Cordero al Malbec
  9:  27,  // AA — SJyQ  — Sorrentinos Jamón y Queso
  10: 28,  // AB — SCa   — Sorrentinos Calabaza y Queso
  11: 29,  // AC — ECaC  — Empanadas Carne a Cuchillo x8
  12: 30,  // AD — EJyQ  — Empanadas Jamón y Queso x8
  17: 31,  // AE — ECyQ  — Empanadas Cebolla y Queso Azul x8
  18: 32,  // AF — EV    — Empanadas Verdura x8
  14: 33,  // AG — TG    — Torta Golosa
  15: 34,  // AH — TLC   — Torta Lemon Crumble
  16: 35,  // AI — TC    — Torta Coco
  13: 36,  // AJ — F     — Franui Leche
  19: 37,  // AK — PMu   — Pizza Muzzarella
  1:  38,  // AL — PMa   — Pizza Margarita
  2:  39,  // AM — PJyQ  — Pizza Jamón y Queso
  3:  40,  // AN — PCC   — Pizza Cebolla Caramelizada
  4:  41,  // AO — PJyM  — Pizza Jamón y Morrón
  24: 57,  // BE — TP    — Tarta Pollo y Verdeo
  25: 58,  // BF — TJyQ  — Tarta Jamón y Queso
  26: 59,  // BG — TCa   — Tarta Calabaza
  27: 60,  // BH — TV    — Tarta Verdura
  // Wraps Claudia Polito (27/05/2026, BJ-BK al final tras "A Favor / Aplicado")
  28: 62,  // BJ — RC    — Wrap Carne
  29: 63,  // BK — RP    — Wrap Pollo
};

// Mapeo id de producto (página web) → abreviatura en hoja Productos (col C)
const PAGE_ID_TO_ABBR = {
  5:  'PPM',   6:  'PPJyQ', 7:  'PPCyQ',
  8:  'SCo',   9:  'SJyQ',  10: 'SCa',
  11: 'ECaC',  12: 'EJyQ',  17: 'ECyQ', 18: 'EV',
  14: 'TG',    15: 'TLC',   16: 'TC',   13: 'F',
  19: 'PMu',  1:  'PMa',  2:  'PJyQ',  3:  'PCC',  4:  'PJyM',
  // Exclusivos de Pilar: sorrentinos adicionales
  20: 'SQB',  21: 'SL',    22: 'SPyP', 23: 'SE',
  // Tartas (Home/Pilar/Red, 15/05/2026)
  24: 'TP', 25: 'TJyQ', 26: 'TCa', 27: 'TV',
  // Wraps Claudia Polito (27/05/2026)
  28: 'RC', 29: 'RP',
};

// ── Layout NUEVO de Pilar (abr/2026 v2): 58 cols. Bloque monetario ampliado con Descuento
// y Total a cobrar en P/Q, productos en W(23)–AS(45), Facturado en V(22).
// Orden productos: PPM, PPJyQ, PPCyQ, SQB, SL, SCo, SPyP, SJyQ, SE, SCa, ECaC, EJyQ, ECyQ, EV, TG, TLC, TC, F, PMu, PMa, PJyQ, PCC, PJyM
const PILAR_PRODUCT_COLS = {
  5:  23, // W  — PPM
  6:  24, // X  — PPJyQ
  7:  25, // Y  — PPCyQ
  20: 26, // Z  — SQB   (exclusivo Pilar)
  21: 27, // AA — SL    (exclusivo Pilar)
  8:  28, // AB — SCo
  22: 29, // AC — SPyP  (exclusivo Pilar)
  9:  30, // AD — SJyQ
  23: 31, // AE — SE    (exclusivo Pilar)
  10: 32, // AF — SCa
  11: 33, // AG — ECaC
  12: 34, // AH — EJyQ
  17: 35, // AI — ECyQ
  18: 36, // AJ — EV
  14: 37, // AK — TG
  15: 38, // AL — TLC
  16: 39, // AM — TC
  13: 40, // AN — F
  19: 41, // AO — PMu
  1:  42, // AP — PMa
  2:  43, // AQ — PJyQ
  3:  44, // AR — PCC
  4:  45, // AS — PJyM
  // Tartas en BH-BK (60-63), al final (no consecutivas tras cols de control)
  24: 60, // BH — TP
  25: 61, // BI — TJyQ
  26: 62, // BJ — TCa
  27: 63, // BK — TV
  // Wraps Claudia Polito (27/05/2026): BM-BN tras "A Favor / Aplicado" en BL
  28: 65, // BM — RC
  29: 66, // BN — RP
};

// Mapeo col → abreviatura para Pilar (inverso de PILAR_PRODUCT_COLS vía PAGE_ID_TO_ABBR)
const PILAR_COL_TO_ABBR = {
  23:'PPM', 24:'PPJyQ', 25:'PPCyQ',
  26:'SQB', 27:'SL', 28:'SCo', 29:'SPyP', 30:'SJyQ', 31:'SE', 32:'SCa',
  33:'ECaC', 34:'EJyQ', 35:'ECyQ', 36:'EV',
  37:'TG', 38:'TLC', 39:'TC', 40:'F',
  41:'PMu', 42:'PMa', 43:'PJyQ', 44:'PCC', 45:'PJyM',
  60:'TP', 61:'TJyQ', 62:'TCa', 63:'TV',
  65:'RC', 66:'RP',
};

// ─── Auto-reserva Home en horario pico (Vie 15hs → Dom 23hs AR) ───
// Si el pedido entra a Home en esa ventana y todos los productos tienen stock
// disponible suficiente, lo damos por Origen=Deposito + Estado=Reservado de movida
// para que Tadeo no tenga que tocarlos uno por uno en horario pico.
function _inVentanaAutoReservaHome(argDate) {
  var day = argDate.getDay();   // 0=Dom, 5=Vie, 6=Sab
  var hour = argDate.getHours();
  if (day === 5 && hour >= 15) return true;  // Viernes desde 15hs
  if (day === 6) return true;                // Sábado completo
  if (day === 0 && hour < 23) return true;   // Domingo hasta 23hs
  return false;
}

// Devuelve true si TODOS los abbrs del map tienen Stock Disponible (col H) >= qty.
// abbrToQty: { 'PMu': 2, 'PMa': 1, ... }
function _stockSuficienteParaAbbrs(hProd, abbrToQty) {
  var prodData = hProd.getDataRange().getValues();
  for (var abbr in abbrToQty) {
    var qty = Number(abbrToQty[abbr]) || 0;
    if (qty <= 0) continue;
    var found = false;
    for (var r = 1; r < prodData.length; r++) {
      if (String(prodData[r][2]).trim() === abbr) {  // col C = Abreviatura
        var disp = Number(prodData[r][7]) || 0;       // col H = Stock Disponible
        if (disp < qty) return false;
        found = true;
        break;
      }
    }
    if (!found) return false;
  }
  return true;
}

// Wrapper que convierte qtys con id_web → abbr antes de chequear.
function _stockSuficienteParaPedidoHome(hProd, qtys) {
  var abbrMap = {};
  for (var idStr in qtys) {
    var q = Number(qtys[idStr]) || 0;
    if (q <= 0) continue;
    var abbr = PAGE_ID_TO_ABBR[Number(idStr)];
    if (!abbr) return false;
    abbrMap[abbr] = (abbrMap[abbr] || 0) + q;
  }
  return _stockSuficienteParaAbbrs(hProd, abbrMap);
}

function _doPostHome(data, sheetName, prefix) {
  sheetName = sheetName || 'Home';
  prefix = prefix || 'H';
  const sh = SS.getSheetByName(sheetName);
  if (!sh) return;

  // ── AUTO-RESERVA Home en ventana Vie 15hs → Dom 23hs ──
  // Solo aplica a Home (no Pilar — los Pilar son a coordinar, no reservamos stock).
  // Solo si el cliente no mandó un origen explícito (default es Pendiente).
  if (sheetName === 'Home' && !data.origen) {
    var _now = new Date();
    var _argNow = new Date(_now.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
    if (_inVentanaAutoReservaHome(_argNow)) {
      var _qtysTmp = {};
      (data.items || []).forEach(function(it) { _qtysTmp[Number(it.id)] = Number(it.qty) || 0; });
      var _hProdChk = SS.getSheetByName('Productos');
      if (_hProdChk && _stockSuficienteParaPedidoHome(_hProdChk, _qtysTmp)) {
        data.origen = 'Deposito';
        data.estadoEntrega = 'Reservado';
      }
    }
  }

  // ── N° de pedido: reinicia por semana ISO según fechaEntrega. Ignora cancelados ("-").
  const orderNum = _nextWeeklyNum(sh, 2, 10, String(data.fechaEntrega || ''), prefix);

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
  let   total           = Number(data.total) || 0; // total a cobrar = subtotal + envio - descuento
  let   descuento       = Number(data.descuento) || 0;
  const subtotalProd    = Number(data.subtotalSinDescuento) || (total + descuento - envio);
  const pago            = String(data.pago || '');

  // ── SALVAGUARDA DESCUENTO (server-side, fuente de verdad) ──
  // El backend SIEMPRE recalcula el descuento desde subtotalProd. No depende del
  // valor que manda el frontend, que puede llegar stale (cache vieja, race con
  // pago no seteado al recalcular, frontend en versión anterior del SW).
  // Caso real que motivó endurecer (08/05/2026): pedido Milagros Pereira Home
  // row 294, subtotal $133.100 con Transferencia → backend guardó descuento $0
  // porque la guarda anterior solo activaba si descuento === 0 estricto y el
  // frontend mandó algún otro valor.
  // Reglas:
  //   • Home: 10% OFF si subtotalProd >= 100k  Ó  pago = Efectivo
  //   • Pilar: idem (los pedidos Pilar con vendedor Red no entran acá, van a hoja Red)
  //   • Clubes: sin descuento
  // ── DESCUENTO MANUAL AUTORITATIVO (PWA Ruta tab "+") ──
  // Cuando Tadeo fija un descuento % explícito (global al pedido o por categoría)
  // desde el alta express, ESE es el descuento final en pesos y REEMPLAZA al 10%
  // automático del efectivo/bulk (no se apila). El frontend manda el flag
  // descuentoManualEsAutoridad=true + descuento ya calculado en pesos.
  // SÓLO este path confía en el front: la tienda online y todos los demás callers
  // nunca setean el flag → la salvaguarda del 10% sigue intacta para ellos.
  var descAutoritativo = (data.descuentoManualEsAutoridad === true || data.descuentoManualEsAutoridad === 'true');
  if (descAutoritativo && sheetName !== 'Clubes') {
    descuento = Number(data.descuento) || 0;
    if (descuento < 0) descuento = 0;
    if (descuento > subtotalProd) descuento = subtotalProd; // nunca descontar más que el subtotal
    total = subtotalProd + envio - descuento;
    try { Logger.log('Descuento autoritativo (tab +): cliente=' + data.nombre + ' subtotal=' + subtotalProd + ' pago=' + pago + ' descuento=' + descuento + ' globalPct=' + (data.descuentoGlobalPct || 0)); } catch(_) {}
  } else if (sheetName !== 'Clubes') {
    // Tipo: 'normal' (retail, aplica 10% OFF auto), 'costo' (familia, sin
    // descuento), 'club' (precio Club, sin descuento). La salvaguarda 10%
    // solo aplica al tipo 'normal'. Bug 12/06/2026: pedido al costo se
    // cobraba con 10% encima del costo (Tadeo se autopedido 1 Franui).
    var _tipoPedido = String(data.tipoPedido || 'normal').trim().toLowerCase();
    if (_tipoPedido === 'normal') {
      const aplicaBulk = subtotalProd >= 100000;
      const aplicaCash = (pago === 'Efectivo');
      const descCorrecto = (aplicaBulk || aplicaCash) ? Math.round(subtotalProd * 0.10) : 0;
      if (descuento !== descCorrecto) {
        try { Logger.log('Descuento recalculado: cliente=' + data.nombre + ' subtotal=' + subtotalProd + ' pago=' + pago + ' frontend_envio=' + descuento + ' backend_corrige=' + descCorrecto); } catch(_) {}
        descuento = descCorrecto;
        total = subtotalProd + envio - descuento;
      }
    } else if (descuento !== 0) {
      // No-normal pero el frontend mandó descuento (legacy): forzar 0 para
      // que "Total al costo" cuadre con lo que vio Tadeo en la pantalla.
      descuento = 0;
      total = subtotalProd + envio;
    }
  }

  // Descuento manual (regalo / bonificación puntual). Se suma al descuento
  // automático del 10% y se resta del total. Cargado desde la PWA Ruta tab "+"
  // por Tadeo cuando regala un producto o hace una bonificación fuera de regla.
  // Caso típico: regalar 1 pizza ($11.000) → producto va igual en el pedido
  // (cuenta para stock) pero se descuenta su valor del total.
  var descuentoManual = Number(data.descuentoManual) || 0;
  if (descuentoManual < 0) descuentoManual = 0;
  if (descuentoManual > 0) {
    descuento = descuento + descuentoManual;
    total = Math.max(0, subtotalProd + envio - descuento);
  }

  // saldoAplicado: si el cliente tenía saldo a favor previo y lo aplicó en checkout,
  // descuenta del cobro real. Total (col Q) se mantiene en el bruto del pedido,
  // pero ef/tr reflejan lo que el cliente paga realmente.
  var saldoAplicadoNum = Number(data.saldoAplicado) || 0;
  if (saldoAplicadoNum < 0) saldoAplicadoNum = 0;
  var cobroReal = Math.max(0, total - saldoAplicadoNum);
  // Mixto: el frontend manda data.efectivo y data.transferencia explícitos
  // (caso "Ya me pagó" desde la tab "+" con pago=Mixto). Si no manda nada,
  // ambos quedan 0 — el pedido se carga como Mixto pero sin cobrar (espera
  // que Tadeo cobre desde la tab COBROS con el modal unificado).
  var efectivo, transferencia;
  if (pago === 'Mixto') {
    efectivo      = Number(data.efectivo)      || 0;
    transferencia = Number(data.transferencia) || 0;
  } else if (pago === 'Efectivo') {
    efectivo      = cobroReal;
    transferencia = 0;
  } else if (pago === 'Transferencia') {
    efectivo      = 0;
    transferencia = cobroReal;
  } else {
    efectivo      = 0;
    transferencia = 0;
  }

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

  // ── Layout Home/Pilar v2 (abr/2026) ──────────────────────
  // Ambas hojas comparten N..V; difieren en nº de productos (Home 19, Pilar 23)
  var isPilar = (sheetName === 'Pilar');

  row[0]  = horaStr;                            // A  Hora Pedido
  row[1]  = orderNum;                           // B  N° Pedido
  row[2]  = diaNombre;                          // C  Día Pedido
  row[3]  = fechaStr;                           // D  Fecha Pedido
  row[4]  = MESES[mes - 1];                      // E  Mes Pedido (nombre)
  row[5]  = semana;                             // F  Semana Pedido
  row[6]  = yyyy;                               // G  Año Pedido
  row[7]  = String(data.nombre || '');          // H  Cliente
  row[8]  = String(data.origen || 'Pendiente');          // I  Origen
  // J Día de entrega elegido: si viene fechaEntrega (ISO YYYY-MM-DD) se guarda como fecha real; si no, fallback al nombre del dia
  (function(){
    var fE = String(data.fechaEntrega || '');
    if (fE && /^\d{4}-\d{2}-\d{2}$/.test(fE)) {
      var p = fE.split('-');
      row[9] = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    } else {
      row[9] = String(data.dia || '');
    }
  })();
  row[10] = String(data.estadoEntrega || 'Pendiente');   // K  Estado de Entrega
  row[11] = pago;                               // L  Forma de Pago
  row[12] = (String(data.estadoPago || '') === 'Cobrado') ? 'Cobrado' : 'No Cobrado';  // M  Estado de Pago
  row[13] = subtotalProd;                       // N  Subtotal Producto
  row[14] = envio;                              // O  Envío
  row[15] = descuento;                          // P  Descuento
  row[16] = total;                              // Q  Total a cobrar
  // Propina + ajuste para Mixto: el frontend manda data.efectivo y data.transferencia
  // como el monto FÍSICO recibido por cada método (incluyendo la parte de propina si
  // hubo). Acá descontamos la propina del método dominante para que col R/S queden
  // SIN propina y T/U tengan la propina aparte (convención del backend).
  // Ejemplo Mixto: pedido $19.970, propina $30, cliente pagó $10k ef + $10k mp.
  //   ef_físico=10000, tr_físico=10000, propina=30 → propMet=Efectivo (empate, default ef)
  //   col R = 9970, col S = 10000, col T = 30, col U = 0.
  //   V = Q + T + U = 19970 + 30 = $20.000 ✓ · Caja: R+T=10000, S+U=10000 ✓
  var propinaNueva = Number(data.propina) || 0;
  var propEfFinal = 0, propTrFinal = 0;
  if (propinaNueva !== 0) {
    if (pago === 'Efectivo')           propEfFinal = propinaNueva;
    else if (pago === 'Transferencia') propTrFinal = propinaNueva;
    else if (pago === 'Mixto') {
      // Asignar al método con más monto físico (en empate, ef por default)
      if (efectivo >= transferencia)   propEfFinal = propinaNueva;
      else                             propTrFinal = propinaNueva;
    }
  }
  // Para Mixto: restar la propina del split físico para que R/S queden netas
  // (consistente con el contrato del backend: R/S sin propina, T/U con propina).
  if (pago === 'Mixto') {
    if (propEfFinal !== 0) efectivo      = Math.max(0, efectivo      - propEfFinal);
    if (propTrFinal !== 0) transferencia = Math.max(0, transferencia - propTrFinal);
  }
  row[17] = efectivo;                           // R  Efectivo (sin propina)
  row[18] = transferencia;                      // S  Transferencia (sin propina)
  row[19] = propEfFinal;                        // T  Propina Efectivo (puede ser negativa = ajuste)
  row[20] = propTrFinal;                        // U  Propina Transferencia
  // V (21) — Facturado: fórmula =Q+T+U (se setea post-append)

  var PRODUCT_COLS = isPilar ? PILAR_PRODUCT_COLS : HOME_PRODUCT_COLS;
  // Cols post-productos: Home productos W(23)–AO(41); Pilar W(23)–AS(45)
  var COL_COSTO    = isPilar ? 46 : 42;  // AT / AP
  var COL_MARGEN   = isPilar ? 47 : 43;  // AU / AQ
  var COL_BARRIO   = isPilar ? 48 : 44;  // AV / AR
  var COL_SUBBAR   = isPilar ? 0  : 45;  // ---- / AS  (Pilar no tiene)
  var COL_LOTE     = isPilar ? 49 : 46;  // AW / AT
  var COL_TEL      = isPilar ? 50 : 47;  // AX / AU
  var COL_ENTREGA  = isPilar ? 51 : 48;  // AY / AV (inicio bloque 6 cols)
  var ROW_LEN      = isPilar ? 58 : 55;

  // Productos (variable por hoja)
  Object.keys(PRODUCT_COLS).forEach(function(id) {
    row[PRODUCT_COLS[id] - 1] = qtys[Number(id)] || 0;
  });

  // Costo + placeholder Margen (fórmulas después del append)
  while (row.length < ROW_LEN) row.push('');
  row[COL_COSTO - 1]  = costoTotal;
  row[COL_MARGEN - 1] = 0;

  // Dirección / Barrio / Lote / Teléfono
  if (isPilar) {
    row[COL_BARRIO - 1] = String(data.barrio || data.direccion || '');
    row[COL_LOTE - 1]   = String(data.lote || '');
    row[COL_TEL - 1]    = String(data.telefono || '');
  } else {
    row[COL_BARRIO - 1] = barrioPrivado;
    row[COL_SUBBAR - 1] = subBarrio;
    row[COL_LOTE - 1]   = String(data.lote || '');
    row[COL_TEL - 1]    = String(data.telefono || '');
  }

  sh.appendRow(row);
  var newRow = sh.getLastRow();

  // ── "Ya me pagó" desde el "+": estampar Fecha de Cobro al crear el pedido ──
  // El frontend (PWA Ruta tab "+") manda data.estadoPago='Cobrado' cuando
  // Tadeo marca el checkbox "Ya me pagó". Como _stampFechaCobro es idempotente
  // (no sobrescribe si ya hay fecha), es seguro llamar incondicionalmente acá.
  if (String(data.estadoPago || '') === 'Cobrado') {
    try { _stampFechaCobro(sh, newRow); } catch (_e) { /* nada */ }
  }

  // ── Repartidor al crear (regla 24/06→06/07/26) ──
  // El Repartidor refleja quién ENTREGÓ. Los autopedidos creados desde el "+" de Ruta
  // mandan data.repartidor = usuario logueado; los pedidos de la tienda (cliente) NO
  // lo mandan. Estampamos el creador SIEMPRE que venga (creado Entregado, Reservado o
  // Pendiente): así los autopedidos que Tadeo carga y entrega él mismo quedan con su
  // nombre aunque no pasen por marcarEntregado. Si más tarde lo entrega OTRA persona
  // desde Ruta, marcarEntregado pisa este valor con el repartidor real (siempre manda
  // _rutaUser()), así que el caso "lo entrega Santos/Agustín" se corrige solo.
  var repCrea = String(data.repartidor || '').trim();
  if (repCrea) {
    var colRepCrea = isPilar ? 59 : 56; // Home col 56 / Pilar col 59 (misma que marcarEntregado)
    sh.getRange(newRow, colRepCrea).setValue(repCrea);
  }

  // ── Saldo a favor aplicado en este pedido (si el cliente tenía crédito previo) ──
  // El frontend tienda detecta saldo via GET saldoCliente y manda saldoAplicado>0.
  // Lo escribimos como NEGATIVO en col BI/BL para que el Facturado refleje lo que
  // el cliente paga realmente (Total - saldoAplicado). También registra movimiento
  // "Aplicación" en hoja Saldos Clientes para llevar la cuenta del crédito restante.
  var saldoAplicado = Number(data.saldoAplicado) || 0;
  if (saldoAplicado > 0) {
    var colAFavIdx = (sheetName === 'Pilar') ? 64 : 61;
    sh.getRange(newRow, colAFavIdx).setValue(-saldoAplicado);
    try {
      _ensureSaldosClientesSheet();
      var shSC = SS.getSheetByName('Saldos Clientes');
      var ahoraSC = new Date();
      var argNowSC = new Date(ahoraSC.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
      shSC.appendRow([
        argNowSC,
        String(data.nombre || ''),
        _normalizarTel(data.telefono),
        'Aplicación',
        saldoAplicado,
        sheetName,
        String(orderNum || ''),
        'Aplicado en checkout (auto-detectado por teléfono)',
        ''
      ]);
      shSC.getRange(shSC.getLastRow(), 1).setNumberFormat('dd/MM/yyyy HH:mm');
    } catch(_) {}
  }

  // Fórmula Facturado (V=22): V = Q (Total a cobrar) + T (Propina Ef) + U (Propina Tr) + BI/BL (A Favor / Aplicado).
  // BI en Home (col 61), BL en Pilar (col 64). Default vacío = 0.
  // A Favor / Aplicado refleja el "extra" cobrado al cliente que queda a favor (positivo)
  // o el saldo aplicado del cliente en una compra futura (negativo).
  // Tadeo (17/05/2026): "facturado pueda ver todo lo que tengo realmente en efectivo".
  // NOTA: BE/BH son cols de Tartas (TP), no usar ahí.
  var colAFavor = (sheetName === 'Pilar') ? 'BL' : 'BI';
  sh.getRange(newRow, 22).setFormula('=Q' + newRow + '+T' + newRow + '+U' + newRow + '+' + colAFavor + newRow);
  // Fórmula Margen Bruto = Facturado (V) - Costo
  var costoLetter = _colLetter(COL_COSTO);
  sh.getRange(newRow, COL_MARGEN).setFormula('=V' + newRow + '-' + costoLetter + newRow);

  // Forzar teléfono como texto
  var telVal = String(data.telefono || '');
  if (telVal) sh.getRange(newRow, COL_TEL).setNumberFormat('@').setValue(telVal);

  // Si se crea como Entregado: stock + fecha entrega
  var estadoEnt = String(data.estadoEntrega || 'Pendiente');
  var origenFinal = String(data.origen || 'Pendiente');
  if (estadoEnt === 'Entregado') {
    var hProd = SS.getSheetByName('Productos');
    if (hProd) {
      if (origenFinal === 'Deposito') _homeStockFisico(sh, newRow, hProd, -1);
      else if (origenFinal === 'Mixto') _homeStockFisicoMixto(sh, newRow, hProd, -1);
    }
    _registrarFechaEntrega(sh, newRow, COL_ENTREGA);
  }

  // Origen Detalle (oD JSON): si origen=Deposito o OC, escribimos el JSON automatico para
  // que la PWA Ruta filtre bien en Armado. Mixto requiere split por producto (no se asume).
  if (origenFinal === 'Deposito' || origenFinal === 'Orden de Compra') {
    var marca = (origenFinal === 'Deposito') ? 'D' : 'OC';
    var oDObj = {};
    Object.keys(PRODUCT_COLS).forEach(function(idStr) {
      var qty = qtys[Number(idStr)] || 0;
      if (qty > 0) {
        var abbr = PAGE_ID_TO_ABBR[Number(idStr)];
        if (abbr) oDObj[abbr] = marca;
      }
    });
    var headersH = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    for (var ihH = 0; ihH < headersH.length; ihH++) {
      if (String(headersH[ihH]).trim() === 'Origen Detalle') {
        sh.getRange(newRow, ihH + 1).setValue(JSON.stringify(oDObj));
        break;
      }
    }
  }

  // ── Descuento Detalle (receta JSON) ──
  // Para que el descuento sobreviva ediciones de productos: si vino una receta con
  // descuento real, la guardamos en la col "Descuento Detalle" (la crea si falta).
  var _recetaObj = _parseReceta(data.descuentoReceta);
  if (_recetaActiva(_recetaObj)) {
    var headersR = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var colRec = 0;
    for (var iR = 0; iR < headersR.length; iR++) {
      if (String(headersR[iR]).trim() === 'Descuento Detalle') { colRec = iR + 1; break; }
    }
    if (!colRec) { colRec = sh.getLastColumn() + 1; sh.getRange(1, colRec).setValue('Descuento Detalle'); }
    sh.getRange(newRow, colRec).setValue(JSON.stringify(_recetaObj));
  }

  // ── Combos: registrar en col "Combo Detalle" + ledger "Combos" ──
  _persistCombos(data, isPilar ? 'Pilar' : 'Home', orderNum, sh, newRow);

  // Sync WATI desactivado (17/05/2026): permiso UrlFetchApp.fetch revocado.
  // Ensuciaba Log Errores con una excepción por pedido sin afectar el grabado.
  // Reactivar cuando se renueve OAuth del Apps Script.
  // _syncWatiContact_(isPilar ? 'Pilar' : 'Home', data.telefono, isPilar ? '' : subBarrio, isPilar ? (data.barrio || data.direccion) : barrioPrivado);
}

// Helper: número de columna → letra (1→A, 27→AA, ...)
function _colLetter(n) {
  var s = '';
  while (n > 0) {
    var rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// ════════════════════════════════════════════════════════════
// COMBOS — persistencia para que el ERP SEPA que un pedido llevó combos.
// La tienda expande el combo a productos reales en las columnas de producto
// (stock/costo OK) pero se perdía la info de que fue un combo. Acá:
//   (1) guardamos la receta JSON (data.comboDetalle) en la col "Combo Detalle"
//       de la hoja del pedido → el Panel puede mostrar "2× Combo Mundialista".
//   (2) agregamos una fila por combo a la hoja ledger "Combos" → evaluación
//       (unidades y facturación por combo, por semana).
// Todo en try/catch: un fallo de logging JAMÁS debe romper el grabado del pedido.
// ════════════════════════════════════════════════════════════
function _ensureCombosSheet() {
  var lg = SS.getSheetByName('Combos');
  if (!lg) {
    lg = SS.insertSheet('Combos');
    lg.appendRow(['Fecha','Semana','Año','Canal','N° Pedido','Cliente','Combo ID','Combo',
                  'Cantidad','Precio Unit','Total Línea','Sabores','Items JSON']);
    lg.setFrozenRows(1);
  }
  return lg;
}

// Evaluación de combos: lee la hoja ledger "Combos" y agrega por combo, por
// semana, y devuelve las últimas ventas. Fuente para el tab Combos del Panel.
function _doGetCombosEval(e) {
  var out = { ok: true, combos: [], totales: { unidades: 0, facturacion: 0, lineas: 0, pedidos: 0 }, semanas: [], recientes: [] };
  try {
    var lg = SS.getSheetByName('Combos');
    if (!lg || lg.getLastRow() < 2) return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
    var vals = lg.getDataRange().getValues();
    // Cols: Fecha0 Semana1 Año2 Canal3 N°4 Cliente5 ComboID6 Combo7 Cantidad8 PrecioUnit9 TotalLínea10 Sabores11 ItemsJSON12
    var byCombo = {}, bySemana = {}, pedidosSet = {};
    for (var r = 1; r < vals.length; r++) {
      var row = vals[r];
      var id = String(row[6] || '').trim(); if (!id) continue;
      var nombre = String(row[7] || id).trim();
      var qty = Number(row[8]) || 0;
      var tot = Number(row[10]) || 0;
      var canal = String(row[3] || '').trim();
      var pedKey = canal + '|' + String(row[4] || '');
      if (!byCombo[id]) byCombo[id] = { id: id, nombre: nombre, unidades: 0, facturacion: 0, lineas: 0, pedidos: {}, canales: {} };
      var c = byCombo[id];
      c.unidades += qty; c.facturacion += tot; c.lineas += 1; c.pedidos[pedKey] = 1;
      c.canales[canal] = (c.canales[canal] || 0) + qty;
      out.totales.unidades += qty; out.totales.facturacion += tot; out.totales.lineas += 1;
      pedidosSet[pedKey] = 1;
      var wk = String(row[2] || '') + '|' + String(row[1] || '');
      if (!bySemana[wk]) bySemana[wk] = { anio: Number(row[2]) || 0, semana: Number(row[1]) || 0, unidades: 0, facturacion: 0 };
      bySemana[wk].unidades += qty; bySemana[wk].facturacion += tot;
    }
    out.totales.pedidos = Object.keys(pedidosSet).length;
    out.combos = Object.keys(byCombo).map(function (k) {
      var c = byCombo[k];
      return { id: c.id, nombre: c.nombre, unidades: c.unidades, facturacion: c.facturacion,
               lineas: c.lineas, pedidos: Object.keys(c.pedidos).length,
               ticket: c.unidades ? Math.round(c.facturacion / c.unidades) : 0, canales: c.canales };
    }).sort(function (a, b) { return b.facturacion - a.facturacion; });
    out.semanas = Object.keys(bySemana).map(function (k) { return bySemana[k]; })
      .sort(function (a, b) { return (a.anio - b.anio) || (a.semana - b.semana); });
    var rec = [];
    for (var rr = vals.length - 1; rr >= 1 && rec.length < 25; rr--) {
      var rw = vals[rr];
      if (!String(rw[6] || '').trim()) continue;
      var _fe = rw[0];
      var _feStr = (_fe instanceof Date) ? Utilities.formatDate(_fe, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : String(_fe || '');
      rec.push({ fecha: _feStr, canal: String(rw[3] || ''), pedido: rw[4],
                 cliente: String(rw[5] || ''), combo: String(rw[7] || ''), qty: Number(rw[8]) || 0,
                 total: Number(rw[10]) || 0, sabores: String(rw[11] || '') });
    }
    out.recientes = rec;
  } catch (err) { out.ok = false; out.err = String(err); }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function _persistCombos(data, canal, orderNum, sh, newRow) {
  var combos = data.combos;
  if (!combos || !combos.length) return;

  // (1) Columna "Combo Detalle" en la hoja del pedido (idempotente: la crea si falta).
  try {
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var colCD = 0;
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim() === 'Combo Detalle') { colCD = i + 1; break; }
    }
    if (!colCD) { colCD = sh.getLastColumn() + 1; sh.getRange(1, colCD).setValue('Combo Detalle'); }
    sh.getRange(newRow, colCD).setValue(data.comboDetalle || JSON.stringify(combos));
  } catch (eCD) { try { Logger.log('Combo Detalle col falló: ' + eCD); } catch (_e) {} }

  // (2) Ledger "Combos" — una fila por combo del pedido.
  try {
    var lg = _ensureCombosSheet();
    var argDate = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
    var dd = String(argDate.getDate()).padStart(2, '0');
    var mm = String(argDate.getMonth() + 1).padStart(2, '0');
    var fechaStr = dd + '/' + mm + '/' + argDate.getFullYear();
    var semana = _isoWeek(argDate);
    var rows = combos.map(function (c) {
      var qty = Number(c.qty) || 0;
      var precio = Number(c.precio) || 0;
      var sabores = (c.picks || []).map(function (pk) { return pk.label + ': ' + pk.nombre; }).join(' · ');
      return [fechaStr, semana, argDate.getFullYear(), canal, orderNum, String(data.nombre || ''),
              String(c.id || ''), String(c.nombre || ''), qty, precio, qty * precio,
              sabores, JSON.stringify(c.items || [])];
    });
    if (rows.length) lg.getRange(lg.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  } catch (eLg) { try { Logger.log('Combos ledger falló: ' + eLg); } catch (_e) {} }
}

// ════════════════════════════════════════════════════════════
// RECETA DE DESCUENTO (col "Descuento Detalle") — tab "+" de Ruta
// Guarda la *receta* del descuento ({g, c, r, auto10}) para que sobreviva a
// ediciones de productos: al editar, col P se recalcula aplicando la receta al
// NUEVO subtotal (en vez de pisar todo con la regla del 10%).
//   g      = % global al pedido
//   c      = { 'Tartas':20, ... } % por categoría
//   r      = regalo $ (bonificación puntual, se suma aparte)
//   auto10 = true si el pedido NO tiene % manual y debe seguir aplicando el 10% auto
// ════════════════════════════════════════════════════════════
// Mapa abbr → categoría (mismo criterio que NP_CAT del front de la tab "+").
var DESC_CAT_OF_ABBR = {
  PPM:'Pack Pizzas x2', PPJyQ:'Pack Pizzas x2', PPCyQ:'Pack Pizzas x2',
  PMu:'Pizzas Individuales', PMa:'Pizzas Individuales', PJyQ:'Pizzas Individuales', PCC:'Pizzas Individuales', PJyM:'Pizzas Individuales',
  RC:'Wraps', RP:'Wraps',
  ECaC:'Empanadas', EJyQ:'Empanadas', ECyQ:'Empanadas', EV:'Empanadas',
  SCo:'Sorrentinos', SJyQ:'Sorrentinos', SCa:'Sorrentinos', SQB:'Sorrentinos', SL:'Sorrentinos', SPyP:'Sorrentinos', SE:'Sorrentinos',
  TP:'Tartas', TJyQ:'Tartas', TCa:'Tartas', TV:'Tartas',
  F:'Franuis',
  TG:'Tortas', TLC:'Tortas', TC:'Tortas'
};

// Parsea la receta desde el JSON guardado o desde el objeto del POST. Devuelve null si vacía.
function _parseReceta(raw) {
  if (!raw) return null;
  var o = (typeof raw === 'object') ? raw : null;
  if (!o) { try { o = JSON.parse(raw); } catch(e) { return null; } }
  if (!o || typeof o !== 'object') return null;
  return o;
}

// ¿La receta es "activa"? (autoritativa → se guarda y manda en la col P).
// True si tiene descuento real (g/c/r) O si suprime el 10% auto (auto10===false,
// caso toggle "Sin 10% OFF" que quiere cobrar precio normal aunque sea Efectivo).
function _recetaActiva(r) {
  if (!r) return false;
  if (r.auto10 === false) return true;
  return (Number(r.g) > 0) || (Number(r.r) > 0) || (r.c && Object.keys(r.c).some(function(k){return Number(r.c[k]) > 0;}));
}

// Calcula el descuento total en pesos (col P) aplicando la receta a las líneas.
// lineas = [{abbr, qty, precio}] (precio unitario ya resuelto). Mirror EXACTO de npDiscountInfo() del front.
function _calcDescuentoReceta(receta, lineas, pago) {
  if (!receta) return null;
  var sub = 0, catDiscP = 0;
  (lineas || []).forEach(function(l){
    var line = (Number(l.precio) || 0) * (Number(l.qty) || 0);
    sub += line;
    var cp = (receta.c && Number(receta.c[DESC_CAT_OF_ABBR[l.abbr]])) || 0;
    if (cp > 0) catDiscP += line * cp / 100;
  });
  catDiscP = Math.round(catDiscP);
  var g = Number(receta.g) || 0;
  var globalDiscP = Math.round(Math.max(0, sub - catDiscP) * g / 100);
  var pctDisc = catDiscP + globalDiscP;
  if (pctDisc > sub) pctDisc = sub;
  var auto = 0;
  if (receta.auto10 && pctDisc === 0 && (pago === 'Efectivo' || sub >= 100000)) auto = Math.round(sub * 0.10);
  var eff = pctDisc > 0 ? pctDisc : auto;
  var regalo = Math.max(0, Number(receta.r) || 0);
  return Math.min(sub, eff + regalo);
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
  // Empanadas habilitadas en Clubes (28/05/2026) — cols AL-AO (0-based 37-40)
  'ecac': 37, // AL — ECaC — Empanadas Carne a Cuchillo x8
  'ejyq': 38, // AM — EJyQ — Empanadas Jamón y Queso x8
  'ecyq': 39, // AN — ECyQ — Empanadas Cebolla y Queso Azul x8
  'evc':  40, // AO — EV   — Empanadas Verdura x8
};

// Mapeo ID producto (web clubes) → abreviatura en hoja Productos
const CLUBES_ID_TO_ABBR = {
  'pmu':'PMu', 'pma':'PMa', 'pjq':'PJyQ', 'pcc':'PCC', 'pjm':'PJyM',
  'pp1':'PPM', 'pp2':'PPJyQ', 'pp3':'PPCyQ',
  'ecac':'ECaC', 'ejyq':'EJyQ', 'ecyq':'ECyQ', 'evc':'EV',
};

function _doPostClubes(data) {
  const sh = SS.getSheetByName('Clubes');
  if (!sh) return;

  // ── N° de pedido: reinicia por semana ISO según fechaEntrega. Ignora cancelados ("-").
  const orderNum = _nextWeeklyNum(sh, 2, 13, String(data.fechaEntrega || ''), 'C');

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
  row[11] = String(data.origen || 'Pendiente');         // L  Origen
  // M Día de Entrega: fechaEntrega como fecha real si viene en ISO, si no nombre del dia
  (function(){
    var fE = String(data.fechaEntrega || '');
    if (fE && /^\d{4}-\d{2}-\d{2}$/.test(fE)) {
      var p = fE.split('-');
      row[12] = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    } else {
      row[12] = String(data.dia || '');
    }
  })();
  row[13] = String(data.estadoEntrega || 'Pendiente');  // N  Estado de Entrega
  row[14] = pago;                              // O  Forma de Pago
  row[15] = (String(data.estadoPago || '') === 'Cobrado') ? 'Cobrado' : 'No Cobrado';  // P  Estado de Pago
  row[16] = total;                             // Q  Total ($)
  row[17] = envio;                             // R  Envío ($)
  row[18] = efectivo;                          // S  Efectivo
  row[19] = transferencia;                     // T  Transferencia
  // Propina: si pago=Efectivo va en U, si pago=Transferencia va en V
  var propinaC = Number(data.propina) || 0;
  row[20] = (pago === 'Efectivo'      && propinaC !== 0) ? propinaC : 0;  // U  Propina Efectivo (puede ser negativa = ajuste)
  row[21] = (pago === 'Transferencia' && propinaC !== 0) ? propinaC : 0;  // V  Propina Transferencia
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

  // Si se crea como Entregado: stock + fecha entrega
  var estadoEntC = String(data.estadoEntrega || 'Pendiente');
  var origenFinalC = String(data.origen || 'Pendiente');
  if (estadoEntC === 'Entregado') {
    var hProdC = SS.getSheetByName('Productos');
    if (hProdC) {
      if (origenFinalC === 'Deposito') _clubesStockFisico(sh, newRow, hProdC, -1);
      else if (origenFinalC === 'Mixto') _clubesStockFisicoMixto(sh, newRow, hProdC, -1);
    }
  }

  // Origen Detalle (oD JSON) auto si origen=Deposito o OC. Mixto requiere split por producto.
  if (origenFinalC === 'Deposito' || origenFinalC === 'Orden de Compra') {
    var marcaC = (origenFinalC === 'Deposito') ? 'D' : 'OC';
    var oDObjC = {};
    Object.keys(CLUBES_PRODUCT_COLS).forEach(function(idC) {
      var qtyC = qtys[idC] || 0;
      if (qtyC > 0) {
        var abbrC = CLUBES_ID_TO_ABBR[idC];
        if (abbrC) oDObjC[abbrC] = marcaC;
      }
    });
    var headersC = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    for (var ihC = 0; ihC < headersC.length; ihC++) {
      if (String(headersC[ihC]).trim() === 'Origen Detalle') {
        sh.getRange(newRow, ihC + 1).setValue(JSON.stringify(oDObjC));
        break;
      }
    }
  }

  // Sync WATI desactivado (17/05/2026): permiso UrlFetchApp.fetch revocado.
  // _syncWatiContact_('Clubes', data.telefono, '', '');
}

// ════════════════════════════════════════════════════════════
//  _doPostRed — escribe en la hoja "Red" (canal vendedores independientes)
//  55 columnas A–BC: Vendedor, 23 productos, comisión 17%, A Pagar, Barrio+Lote+Teléfono, Pago a Maleu
// ════════════════════════════════════════════════════════════

// Mapeo ID producto (web) → columna 1-based hoja Red (20/04/2026: orden actualizado)
// V-X packs · Y-AE sorrentinos · AF-AI empanadas · AJ-AM postres · AN-AR pizzas
const RED_PRODUCT_COLS = {
  5:  22,  // V  — PPM
  6:  23,  // W  — PPJyQ
  7:  24,  // X  — PPCyQ
  20: 25,  // Y  — SQB
  21: 26,  // Z  — SL
  8:  27,  // AA — SCo
  22: 28,  // AB — SPyP
  9:  29,  // AC — SJyQ
  23: 30,  // AD — SE
  10: 31,  // AE — SCa
  11: 32,  // AF — ECaC
  12: 33,  // AG — EJyQ
  17: 34,  // AH — ECyQ
  18: 35,  // AI — EV
  14: 36,  // AJ — TG
  15: 37,  // AK — TLC
  16: 38,  // AL — TC
  13: 39,  // AM — F
  19: 40,  // AN — PMu
  1:  41,  // AO — PMa
  2:  42,  // AP — PJyQ
  3:  43,  // AQ — PCC
  4:  44,  // AR — PJyM
  // Tartas en BF-BI (58-61), al final (no consecutivas tras cols de control)
  24: 58,  // BF — TP
  25: 59,  // BG — TJyQ
  26: 60,  // BH — TCa
  27: 61,  // BI — TV
  // Wraps Claudia Polito (27/05/2026): BJ-BK al final
  28: 62,  // BJ — RC
  29: 63,  // BK — RP
};

// Mapeo ID producto (web Red) → abreviatura en hoja Productos
const RED_ID_TO_ABBR = {
  5:'PPM', 6:'PPJyQ', 7:'PPCyQ',
  8:'SCo', 9:'SJyQ', 10:'SCa',
  11:'ECaC', 12:'EJyQ', 17:'ECyQ', 18:'EV',
  14:'TG', 15:'TLC', 16:'TC', 13:'F',
  20:'SQB', 21:'SL', 22:'SPyP', 23:'SE',
  19:'PMu', 1:'PMa', 2:'PJyQ', 3:'PCC', 4:'PJyM',
  24:'TP', 25:'TJyQ', 26:'TCa', 27:'TV',
  28:'RC', 29:'RP',
};

function _doPostRed(data) {
  const sh = SS.getSheetByName('Red');
  if (!sh) return;

  // ── N° de pedido: reinicia por semana ISO según fechaEntrega. Ignora cancelados ("-").
  const orderNum = _nextWeeklyNum(sh, 2, 11, String(data.fechaEntrega || ''), 'R');

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
  // K Día de Entrega: fechaEntrega como fecha real si viene en ISO, si no nombre del dia
  (function(){
    var fE = String(data.fechaEntrega || '');
    if (fE && /^\d{4}-\d{2}-\d{2}$/.test(fE)) {
      var p = fE.split('-');
      row[10] = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    } else {
      row[10] = String(data.dia || '');
    }
  })();
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
  // AW (col 49) = A Pagar (a Maleu) → fórmula después = Facturado * 83%
  // Soporta ambos nombres: nuevos (barrioPrivado/lote) y viejos (barrioRed/domicilioRed) por retrocompatibilidad
  row[49] = String(data.barrioPrivado || data.barrioRed || '');   // AX  Barrio Privado (col 50)
  row[50] = String(data.lote || data.domicilioRed || '');          // AY  Lote (col 51)
  row[51] = String(data.telefono || '');                           // AZ  Teléfono (col 52)
  // BA (col 53) = Forma Pago a Maleu (vacío hasta que el vendedor marque)
  // BB (col 54) = Estado Pago a Maleu → default "Pendiente"
  // BC (col 55) = Fecha Pago a Maleu (vacío)
  row[53] = 'Pendiente';                                           // BB Estado Pago a Maleu

  sh.appendRow(row);
  var newRow = sh.getLastRow();

  // Fórmula Facturado en U (col 21) = Total + Propinas
  // Facturado Red = O − P (productos, SIN envío): el envío es 100% del vendedor.
  sh.getRange(newRow, 21).setFormula('=O' + newRow + '-P' + newRow);
  // Fórmula Margen Bruto en AT (col 46) = Facturado - Costo
  sh.getRange(newRow, 46).setFormula('=U' + newRow + '-AS' + newRow);
  // Fórmula Comisión 17% en AU (col 47) = Facturado * 17/100
  sh.getRange(newRow, 47).setFormula('=U' + newRow + '*17/100');
  // Fórmula Margen Neto en AV (col 48) = Margen Bruto - Comisión
  sh.getRange(newRow, 48).setFormula('=AT' + newRow + '-AU' + newRow);
  // Fórmula A Pagar en AW (col 49) = Facturado * 83/100 (lo que queda para Maleu)
  sh.getRange(newRow, 49).setFormula('=U' + newRow + '*83/100');
  // Fórmula Ganancia Vendedor en BM (col 65) = Comisión 17% + Envío (lo que gana el vendedor)
  sh.getRange(newRow, 65).setFormula('=AU' + newRow + '+P' + newRow);
  // Forzar teléfono como texto (col AZ = 52)
  var telRed = String(data.telefono || '');
  if (telRed) sh.getRange(newRow, 52).setNumberFormat('@').setValue(telRed);

  // ── Combos: registrar en col "Combo Detalle" + ledger "Combos" ──
  _persistCombos(data, 'Red', orderNum, sh, newRow);

  // Sync WATI desactivado (17/05/2026): permiso UrlFetchApp.fetch revocado.
  // _syncWatiContact_('Red', data.telefono, '', '');
}

// Número de semana ISO (lunes = primer día de la semana)
function _isoWeek(date) {
  const d      = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

/** Lunes (00:00 hora local) de la semana ISO que contiene `date`. */
function _isoWeekMondayLocal(date) {
  var x = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  var day = x.getDay() || 7;            // domingo = 7
  x.setDate(x.getDate() - (day - 1));
  return x;
}

/**
 * Próximo N° de pedido absoluto para una hoja (max + 1).
 * Autoincremental sin reset. Cada hoja lleva su propia secuencia.
 * Ignora filas con N° "-" (canceladas) y vacías.
 * @param {Sheet} sh
 * @param {number} colN — columna 1-based con N° (siempre 2 = B en hojas operativas)
 * @returns {number}
 *
 * Reemplazó a _nextWeeklyNum (23/04/2026): el N° semanal generaba colisiones
 * cuando se movían pedidos manualmente entre semanas, rompiendo identificadores
 * únicos en frontends (PWA Ruta, Búsqueda, Panel). Ahora N° es único de por vida.
 * Los argumentos colDia, fechaEntregaISO y prefix se aceptan por compatibilidad
 * con las llamadas existentes pero se ignoran.
 */
function _nextWeeklyNum(sh, colN, colDia, fechaEntregaISO, prefix) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 1;
  var nValues = sh.getRange(2, colN, lastRow - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < nValues.length; i++) {
    var raw = String(nValues[i][0] || '').trim();
    if (!raw || raw === '-') continue;
    var mm = raw.match(/(\d+)\s*$/);
    if (!mm) continue;
    var nNum = Number(mm[1]);
    if (nNum > max) max = nNum;
  }
  return max + 1;
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
// Home v2 abr/2026: productos van de W(23) a AO(41)
const HOME_COL_TO_ABBR = {
  23: 'PPM',   // W
  24: 'PPJyQ', // X
  25: 'PPCyQ', // Y
  26: 'SCo',   // Z
  27: 'SJyQ',  // AA
  28: 'SCa',   // AB
  29: 'ECaC',  // AC
  30: 'EJyQ',  // AD
  31: 'ECyQ',  // AE
  32: 'EV',    // AF
  33: 'TG',    // AG
  34: 'TLC',   // AH
  35: 'TC',    // AI
  36: 'F',     // AJ
  37: 'PMu',   // AK
  38: 'PMa',   // AL
  39: 'PJyQ',  // AM
  40: 'PCC',   // AN
  41: 'PJyM',  // AO
  // Tartas (15/05/2026, cols BE-BH al final, no consecutivas)
  57: 'TP',    // BE
  58: 'TJyQ',  // BF
  59: 'TCa',   // BG
  60: 'TV',    // BH
  // Wraps Claudia Polito (27/05/2026, BJ-BK tras "A Favor / Aplicado" en BI)
  62: 'RC',    // BJ
  63: 'RP',    // BK
};

// Red: col V(22) a AR(44) = 23 productos (orden actualizado 20/04/2026)
const RED_COL_TO_ABBR = {
  22: 'PPM',   // V
  23: 'PPJyQ', // W
  24: 'PPCyQ', // X
  25: 'SQB',   // Y
  26: 'SL',    // Z
  27: 'SCo',   // AA
  28: 'SPyP',  // AB
  29: 'SJyQ',  // AC
  30: 'SE',    // AD
  31: 'SCa',   // AE
  32: 'ECaC',  // AF
  33: 'EJyQ',  // AG
  34: 'ECyQ',  // AH
  35: 'EV',    // AI
  36: 'TG',    // AJ
  37: 'TLC',   // AK
  38: 'TC',    // AL
  39: 'F',     // AM
  40: 'PMu',   // AN
  41: 'PMa',   // AO
  42: 'PJyQ',  // AP
  43: 'PCC',   // AQ
  44: 'PJyM',  // AR
  // Tartas (15/05/2026, cols BF-BI al final, no consecutivas)
  58: 'TP',    // BF
  59: 'TJyQ',  // BG
  60: 'TCa',   // BH
  61: 'TV',    // BI
  // Wraps Claudia Polito (27/05/2026, BJ-BK al final)
  62: 'RC',    // BJ
  63: 'RP',    // BK
};

// IMPORTANTE: esta función debe configurarse SOLO como trigger instalable.
// NO usar el nombre "onEdit" para evitar doble ejecución (simple + instalable).
// En Activadores: función = onEditHandler, evento = Al editar
function onEditHandler(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // Invalidar cache de entregas si se editó alguna hoja operativa de pedidos.
  // Asi cualquier edicion manual en el Sheet (Estado de Entrega, Origen, etc.) hace
  // que el proximo GET ?action=entregas traiga datos frescos sin esperar el TTL.
  if (sheetName === 'Home' || sheetName === 'Pilar' || sheetName === 'Clubes' || sheetName === 'Red') {
    try { _invalidateEntregasCache(); } catch(eC) {}
  }

  if (sheetName === 'Home' || sheetName === 'Pilar') return _onEditHome(e);
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

  // Skip si el edit viene del POST recibirMercaderia (evita doble +REC / doble -AJU)
  var _skipKey = 'skip_onedit_OC_' + row;
  var _props = PropertiesService.getScriptProperties();
  if (_props.getProperty(_skipKey)) {
    _props.deleteProperty(_skipKey);
    return;
  }

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

  // → Recibido: llenar Fecha Recibido (W=23) + sumar stock SOLO si Canal=Deposito
  if (nuevo === 'Recibido' && anterior !== 'Recibido') {
    sh.getRange(row, 23).setValue(fechaHoy);

    var canal = String(sh.getRange(row, 5).getValue()).trim(); // E = Canal
    if (canal === 'Deposito' && origen === 'Orden de Compra') {
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

  // ← Sale de Recibido (corrección): restar stock SOLO si Canal=Deposito
  if (anterior === 'Recibido' && nuevo !== 'Recibido') {
    var canal2 = String(sh.getRange(row, 5).getValue()).trim(); // E = Canal
    if (canal2 === 'Deposito' && origen === 'Orden de Compra') {
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
  // Empanadas (28/05/2026) — cols AL-AO (1-based 38-41)
  38: 'ECaC',  // AL
  39: 'EJyQ',  // AM
  40: 'ECyQ',  // AN
  41: 'EV',    // AO
};

// ── Red: Origen col J(10), Estado Entrega col L(12) ──
function _onEditRed(e) {
  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row <= 1) return;

  const sh = e.range.getSheet();

  // Skip si el edit viene de un POST que ya hizo todo el trabajo
  if (col === 12) {  // Estado de Entrega
    var _skipKeyR = 'skip_onedit_Red_' + row;
    var _propsR = PropertiesService.getScriptProperties();
    if (_propsR.getProperty(_skipKeyR)) {
      _propsR.deleteProperty(_skipKeyR);
      return;
    }
  }

  // Si cambió Origen (col J=10) a "Orden de Compra" → generar OC
  if (col === 10) {
    const nuevoOrigen = String(e.value || '');
    if (nuevoOrigen === 'Orden de Compra') {
      var lock = LockService.getScriptLock();
      if (lock.tryLock(100)) {
        try {
          var pedidoNum = String(sh.getRange(row, 2).getValue());
          var clienteOrig = String(sh.getRange(row, 9).getValue() || '').trim(); // Red: col 9 = Cliente
          var shOC = SS.getSheetByName('Orden de Compra');
          if (shOC && shOC.getLastRow() > 1) {
            var existentes = shOC.getRange(2, 5, shOC.getLastRow() - 1, 3).getValues();
            for (var i = 0; i < existentes.length; i++) {
              if (String(existentes[i][0]).trim() === 'Red' &&
                  String(existentes[i][1]).trim() === pedidoNum &&
                  String(existentes[i][2]).trim() === clienteOrig) {
                SS.toast('OC ya existe para ' + clienteOrig + ' (' + pedidoNum + ')', 'Duplicado evitado', 4);
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

  // → Cancelado: cancelar OCs vinculadas (antes de cambiar N° a "-") + N° pasa a "-"
  if (nuevo === 'Cancelado' && anterior !== 'Cancelado') {
    var _nPedActualR = String(sh.getRange(row, 2).getValue() || '').trim();
    var _cliActualR = String(sh.getRange(row, 9).getValue() || '').trim(); // Red: col 9 = Cliente
    _cancelarOCsVinculadas('Red', _nPedActualR, _cliActualR);
    sh.getRange(row, 2).setValue('-');
  }
  // ← Sale de Cancelado: re-asignar N° siguiente de la semana (Día Entrega en col K=11)
  if (anterior === 'Cancelado' && nuevo !== 'Cancelado') {
    var _diaStr = String(sh.getRange(row, 11).getValue() || '').trim();
    var _m = _diaStr.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (_m) {
      var _iso = _m[3] + '-' + _m[2].padStart(2, '0') + '-' + _m[1].padStart(2, '0');
      sh.getRange(row, 2).setValue(_nextWeeklyNum(sh, 2, 11, _iso, 'R'));
    }
  }

  // → Entregado: descontar stock si Deposito o Mixto (no hay columnas de fecha entrega en Red)
  if (nuevo === 'Entregado' && anterior !== 'Entregado') {
    if (origen === 'Deposito') {
      const hProductos = SS.getSheetByName('Productos');
      if (hProductos) _redStockFisico(sh, row, hProductos, -1);
    } else if (origen === 'Mixto') {
      const hProductosM = SS.getSheetByName('Productos');
      if (hProductosM) _redStockFisicoMixto(sh, row, hProductosM, -1);
    }
  }
  // ← Sale de Entregado: devolver stock
  if (anterior === 'Entregado' && nuevo !== 'Entregado') {
    if (origen === 'Deposito') {
      const hProductos = SS.getSheetByName('Productos');
      if (hProductos) _redStockFisico(sh, row, hProductos, +1);
    } else if (origen === 'Mixto') {
      const hProductosM2 = SS.getSheetByName('Productos');
      if (hProductosM2) _redStockFisicoMixto(sh, row, hProductosM2, +1);
    }
  }
}

// Stock para Red con origen Mixto: solo descuenta productos con detalle "D"
function _redStockFisicoMixto(shRed, row, hProductos, signo) {
  var headers = shRed.getRange(1, 1, 1, shRed.getLastColumn()).getValues()[0];
  var colDetalle = -1;
  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h]).trim() === 'Origen Detalle') { colDetalle = h + 1; break; }
  }
  if (colDetalle === -1) return;
  var jsonStr = String(shRed.getRange(row, colDetalle).getValue() || '');
  var detalle = {};
  try { detalle = JSON.parse(jsonStr); } catch(e) { return; }

  // BUG 06/06/26: antes 23 cols (V–AR), excluía tartas y wraps. Ampliado a 42 (V–BK).
  var cantidades = shRed.getRange(row, 22, 1, 42).getValues()[0];
  var prodData   = hProductos.getDataRange().getValues();
  var refPedido  = String(shRed.getRange(row, 2).getValue() || '');

  Object.keys(RED_COL_TO_ABBR).forEach(function(colStr) {
    var colIdx = Number(colStr);
    var abbr   = RED_COL_TO_ABBR[colIdx];
    var info   = detalle[abbr];
    if (info == null) return;
    var qty;
    if (typeof info === 'object') qty = Number(info.d) || 0;
    else if (info === 'D')        qty = Number(cantidades[colIdx - 22]) || 0;
    else                          qty = 0;
    if (qty <= 0) return;

    for (var r = 1; r < prodData.length; r++) {
      if (String(prodData[r][2]).trim() === abbr) {
        var celdaFis = hProductos.getRange(r + 1, 6);
        var fisico   = Number(celdaFis.getValue()) || 0;
        var nuevoStock = Math.max(0, fisico + (qty * signo));
        celdaFis.setValue(nuevoStock);
        _logKardex(abbr, signo < 0 ? '-SAL' : '+DEV', qty, fisico, nuevoStock, 'Red', refPedido);
        break;
      }
    }
  });
}

// Stock para Red: 23 productos en cols V(22)-AR(44)
function _redStockFisico(shRed, row, hProductos, signo) {
  // BUG 06/06/26: antes leía 23 cols (V–AR), excluyendo tartas (BF-BI) y wraps (BJ-BK).
  // Ampliado a 42 cols (V=22 hasta BK=63) para cubrir todo el catálogo.
  const cantidades = shRed.getRange(row, 22, 1, 42).getValues()[0]; // cols V–BK
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

  // Skip si el edit viene de un POST que ya hizo todo el trabajo
  if (col === 14) {  // Estado de Entrega
    var _skipKeyC = 'skip_onedit_Clubes_' + row;
    var _propsC = PropertiesService.getScriptProperties();
    if (_propsC.getProperty(_skipKeyC)) {
      _propsC.deleteProperty(_skipKeyC);
      return;
    }
  }

  // Col L (12) = Origen → generar OC con lock
  if (col === 12) {
    const nuevoOrigen = String(e.value || '');
    if (nuevoOrigen === 'Orden de Compra') {
      var lock = LockService.getScriptLock();
      if (lock.tryLock(100)) {
        try {
          var pedidoNum = String(sh.getRange(row, 2).getValue());
          var clienteOrig = String(sh.getRange(row, 8).getValue() || '').trim();
          var shOC = SS.getSheetByName('Orden de Compra');
          if (shOC && shOC.getLastRow() > 1) {
            var existentes = shOC.getRange(2, 5, shOC.getLastRow() - 1, 3).getValues();
            for (var i = 0; i < existentes.length; i++) {
              if (String(existentes[i][0]).trim() === 'Clubes' &&
                  String(existentes[i][1]).trim() === pedidoNum &&
                  String(existentes[i][2]).trim() === clienteOrig) {
                SS.toast('OC ya existe para ' + clienteOrig + ' (' + pedidoNum + ')', 'Duplicado evitado', 4);
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

  // Col P (16) = Estado de Pago → estampar/limpiar Fecha de Cobro
  if (col === 16) {
    var nuevoPagoC = String(e.value || '').trim();
    var anteriorPagoC = String(e.oldValue || '').trim();
    if (nuevoPagoC === 'Cobrado' && anteriorPagoC !== 'Cobrado') _stampFechaCobro(sh, row);
    else if (anteriorPagoC === 'Cobrado' && nuevoPagoC !== 'Cobrado') _clearFechaCobro(sh, row);
    return;
  }

  // Col N (14) = Estado de Entrega → stock solo si Deposito
  if (col === 14) {
    const origen = String(sh.getRange(row, 12).getValue()); // L = Origen
    const nuevo    = String(e.value || '');
    const anterior = String(e.oldValue || '');

    // → Cancelado: cancelar OCs vinculadas (antes de cambiar N° a "-") + N° pasa a "-"
    if (nuevo === 'Cancelado' && anterior !== 'Cancelado') {
      var _nPedActualC = String(sh.getRange(row, 2).getValue() || '').trim();
      var _cliActualC = String(sh.getRange(row, 8).getValue() || '').trim(); // Clubes: col 8 = Cliente
      _cancelarOCsVinculadas('Clubes', _nPedActualC, _cliActualC);
      sh.getRange(row, 2).setValue('-');
    }
    // ← Sale de Cancelado: re-asignar N° siguiente de la semana (Día Entrega en col M=13)
    if (anterior === 'Cancelado' && nuevo !== 'Cancelado') {
      var _diaStrC = String(sh.getRange(row, 13).getValue() || '').trim();
      var _mC = _diaStrC.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (_mC) {
        var _isoC = _mC[3] + '-' + _mC[2].padStart(2, '0') + '-' + _mC[1].padStart(2, '0');
        sh.getRange(row, 2).setValue(_nextWeeklyNum(sh, 2, 13, _isoC, 'C'));
      }
    }

    // → Entregado: descontar Stock Físico si Deposito o Mixto
    if (nuevo === 'Entregado' && anterior !== 'Entregado') {
      if (origen === 'Deposito') {
        const hProductos = SS.getSheetByName('Productos');
        if (hProductos) _clubesStockFisico(sh, row, hProductos, -1);
      } else if (origen === 'Mixto') {
        const hProductosM = SS.getSheetByName('Productos');
        if (hProductosM) _clubesStockFisicoMixto(sh, row, hProductosM, -1);
      }
    }
    // ← Sale de Entregado: devolver Stock Físico si Deposito o Mixto
    if (anterior === 'Entregado' && nuevo !== 'Entregado') {
      if (origen === 'Deposito') {
        const hProductos = SS.getSheetByName('Productos');
        if (hProductos) _clubesStockFisico(sh, row, hProductos, +1);
      } else if (origen === 'Mixto') {
        const hProductosM2 = SS.getSheetByName('Productos');
        if (hProductosM2) _clubesStockFisicoMixto(sh, row, hProductosM2, +1);
      }
    }
  }
}

// Stock para Clubes con origen Mixto: solo descuenta productos con detalle "D"
function _clubesStockFisicoMixto(shClubes, row, hProductos, signo) {
  var headers = shClubes.getRange(1, 1, 1, shClubes.getLastColumn()).getValues()[0];
  var colDetalle = -1;
  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h]).trim() === 'Origen Detalle') { colDetalle = h + 1; break; }
  }
  if (colDetalle === -1) return;
  var jsonStr = String(shClubes.getRange(row, colDetalle).getValue() || '');
  var detalle = {};
  try { detalle = JSON.parse(jsonStr); } catch(e) { return; }

  var cantidades = shClubes.getRange(row, 24, 1, 18).getValues()[0]; // X–AO (incluye empanadas)
  var prodData   = hProductos.getDataRange().getValues();
  var refPedido  = String(shClubes.getRange(row, 2).getValue() || '');

  Object.keys(CLUBES_COL_TO_ABBR).forEach(function(colStr) {
    var colIdx = Number(colStr);
    var abbr   = CLUBES_COL_TO_ABBR[colIdx];
    var info   = detalle[abbr];
    if (info == null) return;
    var qty;
    if (typeof info === 'object') qty = Number(info.d) || 0;
    else if (info === 'D')        qty = Number(cantidades[colIdx - 24]) || 0;
    else                          qty = 0;
    if (qty <= 0) return;

    for (var r = 1; r < prodData.length; r++) {
      if (String(prodData[r][2]).trim() === abbr) {
        var celdaFis = hProductos.getRange(r + 1, 6);
        var fisico   = Number(celdaFis.getValue()) || 0;
        var nuevoStock = Math.max(0, fisico + (qty * signo));
        celdaFis.setValue(nuevoStock);
        _logKardex(abbr, signo < 0 ? '-SAL' : '+DEV', qty, fisico, nuevoStock, 'Clubes', refPedido);
        break;
      }
    }
  });
}

// Ajusta Stock Físico (col F=6) de Productos desde Clubes + log Kardex. signo: -1=restar, +1=sumar
function _clubesStockFisico(shClubes, row, hProductos, signo) {
  // Leer hasta col AO (41) para incluir empanadas (cols AL-AO). X=24 → 18 cols.
  const cantidades = shClubes.getRange(row, 24, 1, 18).getValues()[0]; // cols X–AO
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
 * @param {string} canal - 'Home', 'Pilar', 'Red' o 'Clubes'
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
    // Home v2: AR(43)=Barrio, AS(44)=Sub Barrio, AT(45)=Domicilio-Lote, AU(46)=Teléfono
    colCliente = 7; colPedido = 1; colTelefono = 46;
    direccion = [rowData[43], rowData[44], 'Lote ' + rowData[45]].filter(Boolean).join(' · ');
  } else if (canal === 'Pilar') {
    // Pilar v2: AV(47)=Barrio/Dirección, AW(48)=Domicilio/Lote, AX(49)=Teléfono
    colCliente = 7; colPedido = 1; colTelefono = 49;
    direccion = [rowData[47], rowData[48]].filter(Boolean).join(' · ');
  } else if (canal === 'Red') {
    // Red v3 (55 cols): I(9)=Cliente, B(2)=N°, AZ(52)=Teléfono, AX(50)=Barrio, AY(51)=Lote
    colCliente = 8; colPedido = 1; colTelefono = 51;
    var _barrioR = String(rowData[49] || '').trim();
    var _loteR   = String(rowData[50] || '').trim();
    direccion = [_barrioR, _loteR ? 'Lote ' + _loteR : ''].filter(Boolean).join(' · ');
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
    'PMu':7500, 'PMa':7500, 'PJyQ':7500, 'PCC':7500, 'PJyM':7800,
    'PPM':12000, 'PPJyQ':12000, 'PPCyQ':12000,
    'ECaC':18400, 'EJyQ':16000, 'ECyQ':16000, 'EV':16000
  };

  // BUG 04/06/26: Pilar tiene layout propio (4 sorrentinos extras SQB/SL/SPyP/SE
  // entre cols 26-31) que corre todas las cols de productos. Usar HOME_COL_TO_ABBR
  // para Pilar generaba OCs con el producto equivocado (ej: pedido ECaC qty=1 en
  // col AG/33 generaba OC de TG porque en Home idx 33 = TG). Fix: mapeo propio.
  const colToAbbrMap = (canal === 'Clubes')
    ? CLUBES_COL_TO_ABBR
    : (canal === 'Red')   ? RED_COL_TO_ABBR
    : (canal === 'Pilar') ? PILAR_COL_TO_ABBR
    : HOME_COL_TO_ABBR;

  const cliente    = String(rowData[colCliente] || '').trim();
  const numPedido  = String(rowData[colPedido] || '').trim();
  const telefono   = String(rowData[colTelefono] || '').trim();

  // Diagnóstico: si cliente o numPedido vienen vacíos, abortar antes de ensuciar OC
  Logger.log('generarOC canal=' + canal + ' row=' + row + ' cliente="' + cliente + '" nPed="' + numPedido + '" tel="' + telefono + '" dir="' + direccion + '" rowDataLen=' + rowData.length);
  if (!cliente || !numPedido) {
    SS.toast('OC abortada: fila ' + row + ' sin cliente/N° — revisá el pedido', 'OC error', 8);
    Logger.log('OC abortada: datos insuficientes en fila ' + row);
    return;
  }

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
    // Defensa: solo cantidades sanas (entero positivo razonable). Evita que
    // una celda con formato fecha mal heredado (serial date → número enorme
    // negativo) genere una OC basura. Bug detectado 28/05/2026 con empanadas
    // Clubes recién agregadas.
    if (!(qty > 0 && qty < 10000 && Number.isInteger(qty))) return;

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
      'Pendiente',                               // U  Estado OC (Momento 1: no pasada aún al proveedor)
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
/** Marca como "Cancelado" todas las OCs vinculadas a un pedido cliente
 *  (match por Canal + N° Pedido Origen + Cliente). NO toca OCs Recibidas
 *  (esas ya entraron al stock y no se deben revertir desde acá).
 *  Devuelve cantidad cancelada. Llamado desde _onEditHome/_onEditRed/_onEditClubes
 *  cuando un pedido pasa a Cancelado, ANTES de que el N° se reescriba a "-". */
function _cancelarOCsVinculadas(canal, nPedido, cliente) {
  if (!nPedido || nPedido === '-') return 0;
  var shOC = SS.getSheetByName('Orden de Compra');
  if (!shOC || shOC.getLastRow() <= 1) return 0;
  // Lectura de cols E(Canal=5), F(NPedOrig=6), G(Cliente=7), U(EstadoOC=21)
  var lastRow = shOC.getLastRow();
  var data = shOC.getRange(2, 1, lastRow - 1, 21).getValues();
  var cancelled = 0;
  for (var i = 0; i < data.length; i++) {
    var canalOC = String(data[i][4]).trim();
    var nPedOC  = String(data[i][5]).trim();
    var cliOC   = String(data[i][6]).trim();
    var estadoOC = String(data[i][20]).trim();
    if (canalOC === canal && nPedOC === nPedido && cliOC === cliente &&
        estadoOC !== 'Recibido' && estadoOC !== 'Cancelado') {
      shOC.getRange(i + 2, 21).setValue('Cancelado'); // col U = 21
      cancelled++;
    }
  }
  if (cancelled > 0) {
    SS.toast('Canceladas ' + cancelled + ' OC vinculadas a ' + cliente + ' (' + nPedido + ')', 'OC', 5);
  }
  return cancelled;
}

function _onEditHome(e) {
  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row <= 1) return;

  const sh = e.range.getSheet();

  // Skip si el edit viene de un POST que ya hizo todo el trabajo (evita doble descuento de stock/fecha)
  if (col === 11) {  // solo aplicable para Estado de Entrega
    var _skipKey = 'skip_onedit_' + sh.getName() + '_' + row;
    var _props = PropertiesService.getScriptProperties();
    if (_props.getProperty(_skipKey)) {
      _props.deleteProperty(_skipKey);
      return;
    }
  }

  // Si cambió Origen (col I=9) a "Orden de Compra" → generar OC con lock
  if (col === 9) {
    const nuevoOrigen = String(e.value || '');
    if (nuevoOrigen === 'Orden de Compra') {
      var lock = LockService.getScriptLock();
      if (lock.tryLock(100)) {
        try {
          // Verificar si ya existen OC para este pedido (Canal+N°+Cliente, no solo N°)
          var pedidoNum = String(sh.getRange(row, 2).getValue());
          var clienteOrig = String(sh.getRange(row, 8).getValue() || '').trim();
          var canalActual = sh.getName();
          var shOC = SS.getSheetByName('Orden de Compra');
          if (shOC && shOC.getLastRow() > 1) {
            var existentes = shOC.getRange(2, 5, shOC.getLastRow() - 1, 3).getValues(); // E=Canal, F=N°, G=Cliente
            for (var i = 0; i < existentes.length; i++) {
              if (String(existentes[i][0]).trim() === canalActual &&
                  String(existentes[i][1]).trim() === pedidoNum &&
                  String(existentes[i][2]).trim() === clienteOrig) {
                SS.toast('OC ya existe para ' + clienteOrig + ' (' + pedidoNum + ')', 'Duplicado evitado', 4);
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

  // Estado de Pago (col M=13): estampar/limpiar Fecha de Cobro
  if (col === 13) {
    var nuevoPago = String(e.value || '').trim();
    var anteriorPago = String(e.oldValue || '').trim();
    if (nuevoPago === 'Cobrado' && anteriorPago !== 'Cobrado') _stampFechaCobro(sh, row);
    else if (anteriorPago === 'Cobrado' && nuevoPago !== 'Cobrado') _clearFechaCobro(sh, row);
    return;
  }

  if (col !== 11) return; // solo col K (11) = Estado de Entrega

  const origen = String(sh.getRange(row, 9).getValue()); // col I (9) = Origen
  const nuevo    = String(e.value || '');
  const anterior = String(e.oldValue || '');

  // → Cancelado: cancelar OCs vinculadas (antes de cambiar N° a "-") + N° pasa a "-"
  if (nuevo === 'Cancelado' && anterior !== 'Cancelado') {
    var _nPedActual = String(sh.getRange(row, 2).getValue() || '').trim();
    var _cliActual = String(sh.getRange(row, 8).getValue() || '').trim();
    _cancelarOCsVinculadas(sh.getName(), _nPedActual, _cliActual);
    sh.getRange(row, 2).setValue('-');
  }
  // ← Sale de Cancelado: re-asignar N° siguiente de la semana (prefix según hoja: H=Home, P=Pilar)
  if (anterior === 'Cancelado' && nuevo !== 'Cancelado') {
    var _diaStr = String(sh.getRange(row, 10).getValue() || '').trim();
    var _m = _diaStr.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (_m) {
      var _iso = _m[3] + '-' + _m[2].padStart(2, '0') + '-' + _m[1].padStart(2, '0');
      var _pref = (sh.getName() === 'Pilar') ? 'P' : 'H';
      sh.getRange(row, 2).setValue(_nextWeeklyNum(sh, 2, 10, _iso, _pref));
    }
  }

  // Columna donde empieza "Hora Entrega": Home=48 (AV), Pilar=51 (AY) — layout v2 abr/2026
  const sheetName = sh.getName();
  const colEntrega = sheetName === 'Pilar' ? 51 : 48;

  // → Entregado: registrar fecha SIEMPRE + descontar stock según origen
  if (nuevo === 'Entregado' && anterior !== 'Entregado') {
    _registrarFechaEntrega(sh, row, colEntrega);
    if (origen === 'Deposito') {
      const hProductos = SS.getSheetByName('Productos');
      if (hProductos) _homeStockFisico(sh, row, hProductos, -1);
    } else if (origen === 'Mixto') {
      // Solo descontar productos con origen Deposito (leer JSON de Origen Detalle)
      const hProductos = SS.getSheetByName('Productos');
      if (hProductos) _homeStockFisicoMixto(sh, row, hProductos, -1);
    }
  }

  // ← Sale de Entregado: limpiar fecha SIEMPRE + devolver stock según origen
  if (anterior === 'Entregado' && nuevo !== 'Entregado') {
    sh.getRange(row, colEntrega, 1, 6).clearContent();
    if (origen === 'Deposito') {
      const hProductos = SS.getSheetByName('Productos');
      if (hProductos) _homeStockFisico(sh, row, hProductos, +1);
    } else if (origen === 'Mixto') {
      const hProductos = SS.getSheetByName('Productos');
      if (hProductos) _homeStockFisicoMixto(sh, row, hProductos, +1);
    }
  }
}

// Llena 6 columnas de entrega desde colStart en zona Argentina
/** Devuelve el índice 1-based de la col "Fecha de Cobro". La crea si no existe. */
function _ensureFechaCobroCol(sh) {
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h]).trim() === 'Fecha de Cobro') return h + 1;
  }
  var newCol = sh.getLastColumn() + 1;
  sh.getRange(1, newCol).setValue('Fecha de Cobro')
    .setBackground('#5D4037').setFontColor('#FFFFFF').setFontWeight('bold');
  return newCol;
}

/** Estampa fecha+hora argentina en la col "Fecha de Cobro" de `row`. */
function _stampFechaCobro(sh, row) {
  var col = _ensureFechaCobroCol(sh);
  // No sobreescribir si ya hay fecha (evita que una edición posterior pisé la original)
  var actual = String(sh.getRange(row, col).getValue() || '').trim();
  if (actual) return;
  var ahora = new Date();
  var argDate = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var dd = String(argDate.getDate()).padStart(2, '0');
  var mm = String(argDate.getMonth() + 1).padStart(2, '0');
  var yyyy = argDate.getFullYear();
  var hh = String(argDate.getHours()).padStart(2, '0');
  var mi = String(argDate.getMinutes()).padStart(2, '0');
  sh.getRange(row, col).setValue(dd + '/' + mm + '/' + yyyy + ' ' + hh + ':' + mi);
}

/** Limpia la col "Fecha de Cobro" de `row` (cuando se revierte de Cobrado → No Cobrado). */
function _clearFechaCobro(sh, row) {
  var col = _ensureFechaCobroCol(sh);
  sh.getRange(row, col).clearContent();
}

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
  var isPilar = (shHome.getName() === 'Pilar');
  var prodCount = isPilar ? 23 : 19;
  var COL_MAP   = isPilar ? PILAR_COL_TO_ABBR : HOME_COL_TO_ABBR;
  // Leer hasta la última col de wraps (Home BK=63, Pilar BN=66) para que el
  // mapeo de tartas y wraps también caiga dentro del array de cantidades.
  // BUG 06/06/26: antes era 60/63 → no descontaba wraps RC/RP del stock físico.
  var maxColPC = isPilar ? 66 : 63;
  const cantidades = shHome.getRange(row, 23, 1, maxColPC - 22).getValues()[0];
  const prodData   = hProductos.getDataRange().getValues();
  const refPedido  = String(shHome.getRange(row, 2).getValue() || ''); // B = N° Pedido

  Object.keys(COL_MAP).forEach(function(colStr) {
    const colIdx = Number(colStr);
    const abbr   = COL_MAP[colIdx];
    const qty    = Number(cantidades[colIdx - 23]) || 0;
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

  var isPilarMx = (shHome.getName() === 'Pilar');
  var prodCountMx = isPilarMx ? 23 : 19;
  var COL_MAP_MX  = isPilarMx ? PILAR_COL_TO_ABBR : HOME_COL_TO_ABBR;
  // Leer hasta la última col de wraps (Home BK=63, Pilar BN=66) — antes 60/63.
  var maxColMX = isPilarMx ? 66 : 63;
  var cantidades = shHome.getRange(row, 23, 1, maxColMX - 22).getValues()[0];
  var prodData = hProductos.getDataRange().getValues();
  var refPedido = String(shHome.getRange(row, 2).getValue() || '');

  Object.keys(COL_MAP_MX).forEach(function(colStr) {
    var colIdx = Number(colStr);
    var abbr = COL_MAP_MX[colIdx];
    var info = detalle[abbr];
    if (info == null) return;
    // Soporta:
    //  Nuevo:  { d: N, oc: Y }    → descuenta N
    //  Viejo:  "D"                → descuenta toda la cantidad del pedido
    //  Otro:   "OC" / { d:0,... } → no descuenta
    var qty;
    if (typeof info === 'object') {
      qty = Number(info.d) || 0;
    } else if (info === 'D') {
      qty = Number(cantidades[colIdx - 23]) || 0;
    } else {
      qty = 0;
    }
    if (qty <= 0) return;

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
// Home v2 abr/2026: productos van de W(23) a AO(41)
const ABBR_TO_HOME_COL = {
  'PPM':'W', 'PPJyQ':'X', 'PPCyQ':'Y',
  'SCo':'Z', 'SJyQ':'AA', 'SCa':'AB',
  'ECaC':'AC', 'EJyQ':'AD', 'ECyQ':'AE', 'EV':'AF',
  'TG':'AG', 'TLC':'AH', 'TC':'AI', 'F':'AJ',
  'PMu':'AK', 'PMa':'AL', 'PJyQ':'AM', 'PCC':'AN', 'PJyM':'AO',
  // Tartas (BE-BH) y Wraps Claudia Polito (BJ-BK) — al final
  'TP':'BE', 'TJyQ':'BF', 'TCa':'BG', 'TV':'BH',
  'RC':'BJ', 'RP':'BK',
};

// Mapeo para Clubes: abreviatura → letra de columna en Clubes
// Productos van de X(24) a AE(31)
const ABBR_TO_CLUBES_COL = {
  'PMu':'X', 'PMa':'Y', 'PJyQ':'Z', 'PCC':'AA', 'PJyM':'AB',
  'PPM':'AC', 'PPJyQ':'AD', 'PPCyQ':'AE',
};

// Pilar (layout nuevo abr/2026 v2): 23 productos en cols W(23)–AS(45)
const ABBR_TO_PILAR_COL = {
  'PPM':'W', 'PPJyQ':'X', 'PPCyQ':'Y',
  'SQB':'Z', 'SL':'AA', 'SCo':'AB', 'SPyP':'AC', 'SJyQ':'AD', 'SE':'AE', 'SCa':'AF',
  'ECaC':'AG', 'EJyQ':'AH', 'ECyQ':'AI', 'EV':'AJ',
  'TG':'AK', 'TLC':'AL', 'TC':'AM', 'F':'AN',
  'PMu':'AO', 'PMa':'AP', 'PJyQ':'AQ', 'PCC':'AR', 'PJyM':'AS',
  // Tartas (BH-BK) y Wraps Claudia Polito (BM-BN) — al final
  'TP':'BH', 'TJyQ':'BI', 'TCa':'BJ', 'TV':'BK',
  'RC':'BM', 'RP':'BN',
};

// Mapeo Red: abreviatura → letra de columna en Red (productos van de V(22) a AR(44))
const ABBR_TO_RED_COL = {
  'PPM':'V', 'PPJyQ':'W', 'PPCyQ':'X',
  'SQB':'Y', 'SL':'Z', 'SCo':'AA', 'SPyP':'AB', 'SJyQ':'AC', 'SE':'AD', 'SCa':'AE',
  'ECaC':'AF', 'EJyQ':'AG', 'ECyQ':'AH', 'EV':'AI',
  'TG':'AJ', 'TLC':'AK', 'TC':'AL', 'F':'AM',
  'PMu':'AN', 'PMa':'AO', 'PJyQ':'AP', 'PCC':'AQ', 'PJyM':'AR',
  // Tartas (BF-BI) y Wraps Claudia Polito (BJ-BK) — al final
  'TP':'BF', 'TJyQ':'BG', 'TCa':'BH', 'TV':'BI',
  'RC':'BJ', 'RP':'BK',
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

  // Helper: arma un SUMPRODUCT por hoja con soporte Deposito + Mixto + filtro estado/semana.
  // - colOrigen, colEstado, colDetalle: letras de columna de la hoja
  // - colSem, colAnio: letras de columnas con semana/año a filtrar (null = sin filtro)
  // - colProd: letra de la columna del producto en la hoja
  // - estado: "Entregado" o "Reservado"
  function buildHojaTerm(hoja, colOrigen, colEstado, colDetalle, colSem, colAnio, colProd, estado, abbr) {
    var rng = function(c) { return hoja + '!$' + c + '$2:$' + c + '$10000'; };
    var rngP = hoja + '!' + colProd + '$2:' + colProd + '$10000';
    var origen = '(' + rng(colOrigen) + '="Deposito")';
    // Mixto: contiene "<abbr>":"D" en JSON de Origen Detalle
    var mixto = '(' + rng(colOrigen) + '="Mixto")*IFERROR(REGEXMATCH(' + rng(colDetalle) +
                ';"""' + abbr + '"":""D""");FALSE)';
    var multStock = '((' + origen + '+' + mixto + ')*' + rngP + ')';
    var filtroEstado = '(' + rng(colEstado) + '="' + estado + '")';
    var filtroSem = (colSem && colAnio)
      ? '*(' + rng(colSem) + '=' + semanaActual + ')*(' + rng(colAnio) + '=' + anioActual + ')'
      : '';
    return 'SUMPRODUCT(' + multStock + '*' + filtroEstado + filtroSem + ')';
  }

  for (let r = 1; r < data.length; r++) {
    const abbr    = String(data[r][2]).trim(); // col C = Abreviatura
    const homeCol   = ABBR_TO_HOME_COL[abbr];
    const pilarCol  = ABBR_TO_PILAR_COL[abbr];
    const clubesCol = ABBR_TO_CLUBES_COL[abbr];
    const redCol    = ABBR_TO_RED_COL[abbr];
    if (!homeCol && !pilarCol && !clubesCol && !redCol) continue;

    const rowNum = r + 1;

    // Col E (Vendidos Semana) — Entregados de la semana en curso
    var vendidosTerms = [];
    // Home: Sem Entrega=AZ, Año Entrega=BA, Origen Detalle=BB
    if (homeCol)   vendidosTerms.push(buildHojaTerm('Home',   'I','K','BB','AZ','BA', homeCol,   'Entregado', abbr));
    // Pilar: Sem Entrega=BC, Año Entrega=BD, Origen Detalle=BE
    if (pilarCol)  vendidosTerms.push(buildHojaTerm('Pilar',  'I','K','BE','BC','BD', pilarCol,  'Entregado', abbr));
    // Clubes: usa Semana de creación F y Año G (no tiene cols de entrega), Origen Detalle=AJ
    if (clubesCol) vendidosTerms.push(buildHojaTerm('Clubes', 'L','N','AJ','F','G',   clubesCol, 'Entregado', abbr));
    // Red: usa Semana de creación F y Año G, Origen Detalle=BD
    if (redCol)    vendidosTerms.push(buildHojaTerm('Red',    'J','L','BD','F','G',   redCol,    'Entregado', abbr));
    hProd.getRange(rowNum, 5).setFormula('=' + vendidosTerms.join('+'));

    // Col G (Reservado) — Reservados activos (sin filtro de semana)
    var reservadoTerms = [];
    if (homeCol)   reservadoTerms.push(buildHojaTerm('Home',   'I','K','BB',null,null, homeCol,   'Reservado', abbr));
    if (pilarCol)  reservadoTerms.push(buildHojaTerm('Pilar',  'I','K','BE',null,null, pilarCol,  'Reservado', abbr));
    if (clubesCol) reservadoTerms.push(buildHojaTerm('Clubes', 'L','N','AJ',null,null, clubesCol, 'Reservado', abbr));
    if (redCol)    reservadoTerms.push(buildHojaTerm('Red',    'J','L','BD',null,null, redCol,    'Reservado', abbr));
    hProd.getRange(rowNum, 7).setFormula('=' + reservadoTerms.join('+'));

    // Col H (Stock Disponible) = Stock Físico - Reservado
    hProd.getRange(rowNum, 8).setFormula('=F' + rowNum + '-G' + rowNum);

    // Col K (Margen Unitario) = Precio - Costo
    hProd.getRange(rowNum, 11).setFormula('=I' + rowNum + '-J' + rowNum);

    // Col L (Check) = Stock Inicial + Comprado - Vendidos - Stock Físico (debería ser 0 si todo cuadra)
    hProd.getRange(rowNum, 12).setFormula('=D' + rowNum + '+N' + rowNum + '-E' + rowNum + '-F' + rowNum);

    // Col N (Comprado Semana) — OC recibidas en la semana en curso (Lun 00:00 → Dom 23:59 AR)
    // SOLO reposición de depósito (canal "Deposito"/"Depósito") — excluye OC pasamanos con cliente
    // final (canal Home/Pilar/Clubes/Red). Coincide con sumarStock() que solo suma Físico si canal
    // empieza con "Dep". Así se mantiene la identidad Inicial + Comprado - Vendidos = Físico.
    hProd.getRange(rowNum, 14).setFormula(
      '=SUMPRODUCT(' +
      "('Orden de Compra'!$L$2:$L$10000=$C" + rowNum + ')*' +
      '(\'Orden de Compra\'!$U$2:$U$10000="Recibido")*' +
      '(LEFT(\'Orden de Compra\'!$E$2:$E$10000;3)="Dep")*' +
      "('Orden de Compra'!$W$2:$W$10000>=TODAY()-WEEKDAY(TODAY();3))*" +
      "('Orden de Compra'!$W$2:$W$10000<TODAY()-WEEKDAY(TODAY();3)+7)*" +
      "'Orden de Compra'!$M$2:$M$10000)"
    );
  }

  // Asegurar headers de col L y N
  hProd.getRange(1, 12).setValue('Check (D+N-E-F)');
  hProd.getRange(1, 14).setValue('Comprado Semana');

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

/** Instala el trigger semanal que ejecuta resetStockSemanal cada Lunes 00:00 AR.
 *  Ejecutar UNA sola vez desde el editor de Apps Script. Idempotente: si ya existe, lo reemplaza. */
function installResetStockTrigger() {
  var existentes = ScriptApp.getProjectTriggers();
  var borrados = 0;
  for (var i = 0; i < existentes.length; i++) {
    if (existentes[i].getHandlerFunction() === 'resetStockSemanal') {
      ScriptApp.deleteTrigger(existentes[i]);
      borrados++;
    }
  }
  ScriptApp.newTrigger('resetStockSemanal')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(0)
    .inTimezone('America/Argentina/Buenos_Aires')
    .create();
  SpreadsheetApp.getUi().alert('Trigger resetStockSemanal instalado (Lunes 00:00 AR). Borrados previos: ' + borrados);
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
    .addItem('Actualizar fórmulas Productos', 'setupProductosFormulas')
    .addItem('Reset stock semanal', 'resetStockSemanal')
    .addItem('Instalar trigger reset semanal (Lun 00:00)', 'installResetStockTrigger')
    .addSeparator()
    .addItem('Ver/Editar Config', 'mostrarConfig')
    .addSeparator()
    .addItem('⚠️ Activar alerta stock bajo (cada 6hs)', 'setupTriggerStockBajo')
    .addToUi();
}

/** Marca TODAS las OC en estado "Pedido" como "Recibido" de un solo click.
 *  Solo suma stock en Productos si Canal = "Deposito".
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
    'Las de Canal "Deposito" van a sumar stock en Productos.\n' +
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

    // Solo sumar stock si Canal = Deposito Y Origen = Orden de Compra
    if (canal === 'Deposito' && origen === 'Orden de Compra' && abbr && qty > 0) {
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
  if (prodActualizados > 0) msg += '\n' + prodActualizados + ' productos actualizados en stock (Deposito)';
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

    if (canal === 'Deposito' && origen === 'Orden de Compra' && abbr && qty > 0) {
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
    if (origen === 'Deposito') continue;

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
      // Col H (índice 7) = Stock Disponible
      stockMap[ab] = Number(prodData[r][7]) || 0;
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
      cat: lastProd,
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

/** Genera filas en Orden de Compra para compra de Deposito / Red (sidebar) */
function confirmarCompraDeposito(proveedor, items, fechaBusqueda, vendedor) {
  const shOC = SS.getSheetByName('Orden de Compra');
  if (!shOC) throw new Error('Hoja "Orden de Compra" no encontrada');

  // Es compra Red si va dirigida a un vendedor (cualquiera) y no al stock propio.
  var esVendedorRed = vendedor && vendedor !== 'Tadeo — Stock';
  var dirVendedor = esVendedorRed ? _dirVendedorRed(vendedor) : 'Deposito Maleu';

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
      esVendedorRed ? 'Red' : 'Deposito',        // E  Canal
      '',                                         // F  N° Pedido Origen (no aplica)
      vendedor || 'Tadeo — Stock',                // G  Cliente
      '',                                         // H  Teléfono
      dirVendedor,                                // I  Dirección
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
<p class="subtitle">Reposición de stock para Deposito</p>

<label for="sel-vendedor">¿Para quién es la compra?</label>
<select id="sel-vendedor">
  <option value="Tadeo — Stock">Tadeo — Stock (Deposito)</option>
  <option value="Marcos Bottcher">Marcos Bottcher (Red)</option>
  <option value="Fini Mihailovitch">Fini Mihailovitch (Red)</option>
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
        '<option value="Deposito">Dep</option>' +
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
    var origenTag = item.origen === 'Deposito' ? ' <small style="color:#e67e22">(Dep)</small>' : '';
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
    .requireValueInList(['Pendiente','Deposito','Orden de Compra'], true)
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
    .whenTextEqualTo('Deposito')
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
    .requireValueInList(['Orden de Compra','Deposito'], true)
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
    .whenTextEqualTo('Deposito')
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

  // Nueva estructura con Canal (col 2) y Deposito (col 14)
  const headers = ['ID Pedido','Canal','Fecha','Nombre','Barrio','Lote','Teléfono',
                   'Día entrega','Horario','Pago','Total','Estado','Fecha solo','Deposito'];
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
function _doGetAdmin(opt) {
  var _lightOnly = opt && opt.light === true;
  var _cajaMode = opt && opt.mode === 'caja';
  // _tail: si > 0, leer solo las últimas N filas de cada hoja operativa. Optimización
  // del botón "Actualizar pedidos" — un pedido nuevo cae en las últimas filas, así
  // que pasamos de leer ~407 filas (Home 330 + Pilar + Clubes + Red) a leer ~40.
  // En modo tail, los stats por canal (canales[]) no son representativos → se devuelven
  // vacíos y el frontend conserva los anteriores.
  var _tail = (opt && opt.tail) ? Math.max(3, Math.min(50, opt.tail)) : 0;
  // Cache server-side (TTL 5s) para cajaLight y pedidosLight. La primera lectura
  // sigue tomando ~8s, pero clicks consecutivos del usuario son instantáneos.
  // Se invalida automáticamente cada 5s, así un cobro nuevo aparece en <5s + 8s.
  // En modo tail NO cacheamos: queremos data fresca en cada click del botón.
  var _cacheKey = _tail ? null : (_cajaMode ? 'admin_caja_v1' : (_lightOnly ? 'admin_light_v1' : null));
  if (_cacheKey) {
    try {
      var _cache = CacheService.getScriptCache();
      var _cached = _cache.get(_cacheKey);
      if (_cached) {
        return ContentService.createTextOutput(_cached).setMimeType(ContentService.MimeType.JSON);
      }
    } catch (e) { /* sin cache no rompe el flujo */ }
  }
  var ahora = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));

  // Helper: parsea Date o string "dd/MM/yyyy" o "dd/MM/yyyy HH:mm" → Date (o null)
  function _parseDateAny(v) {
    if (!v) return null;
    if (v instanceof Date) return v;
    var s = String(v).trim();
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
    if (!m) return null;
    var yr = Number(m[3]), mo = Number(m[2]) - 1, dd = Number(m[1]);
    var hh = m[4] ? Number(m[4]) : 0;
    var mi = m[5] ? Number(m[5]) : 0;
    return new Date(yr, mo, dd, hh, mi);
  }

  // ── Saldo Base (último ajuste manual) — leído al principio para filtrar cobros/ingresos/gastos ──
  // En modo light skipeamos esto: Saldo Base solo se usa para los cálculos de
  // caja, no para armar pedidos[]. Ahorra una lectura.
  // bil = sub-saldo "Billetera" (lo que llevas encima para cambio). Es subconjunto
  // del Efectivo total — caja fuerte se deriva como ef - bil. Default 0 si la col
  // todavía no se llenó.
  // sob = sub-saldo "Sobres": efectivo apartado en sobres para pagar proveedores (no
  // está en la caja fuerte ni en la billetera). Caja fuerte = ef - bil - sob.
  // inv = "Inversiones" (col F): plata que salió de MP líquido hacia inversiones
  // (FCI de Mercado Pago, plazo fijo, etc). NO es MP líquido pero sí es de Maleu.
  var saldoBase = { ef: 0, mp: 0, bil: 0, sob: 0, inv: 0, fecha: '', fechaDate: null };
  var shSaldoSnap = (!_lightOnly) ? SS.getSheetByName('Saldo Base') : null;
  if (shSaldoSnap && shSaldoSnap.getLastRow() > 1) {
    var lastRowSB = shSaldoSnap.getLastRow();
    saldoBase.ef = Number(shSaldoSnap.getRange(lastRowSB, 2).getValue()) || 0;
    saldoBase.mp = Number(shSaldoSnap.getRange(lastRowSB, 3).getValue()) || 0;
    // Lectura defensiva de col D/E/F: pueden no existir en hojas viejas
    if (shSaldoSnap.getLastColumn() >= 4) {
      saldoBase.bil = Number(shSaldoSnap.getRange(lastRowSB, 4).getValue()) || 0;
    }
    if (shSaldoSnap.getLastColumn() >= 5) {
      saldoBase.sob = Number(shSaldoSnap.getRange(lastRowSB, 5).getValue()) || 0;
    }
    if (shSaldoSnap.getLastColumn() >= 6) {
      saldoBase.inv = Number(shSaldoSnap.getRange(lastRowSB, 6).getValue()) || 0;
    }
    var fSB = shSaldoSnap.getRange(lastRowSB, 1).getValue();
    var fSBd = _parseDateAny(fSB);
    if (fSBd) {
      saldoBase.fechaDate = fSBd;
      saldoBase.fecha = Utilities.formatDate(fSBd, 'America/Argentina/Buenos_Aires', 'dd/MM HH:mm');
    }
  }
  // Helper: movimiento cuenta solo si ocurrió DESPUÉS del último ajuste de saldo.
  // Si no hay ajuste previo, todo cuenta. Si la fecha es ilegible, se asume anterior (conservador).
  function _afterSaldo(v) {
    if (!saldoBase.fechaDate) return true;
    var d = _parseDateAny(v);
    if (!d) return false;
    return d > saldoBase.fechaDate;
  }

  // ── Leer pedidos de las 4 hojas operativas ──
  var ABBRS_HOME = ['PPM','PPJyQ','PPCyQ','SCo','SJyQ','SCa','ECaC','EJyQ','ECyQ','EV',
                    'TG','TLC','TC','F','PMu','PMa','PJyQ','PCC','PJyM'];
  var ABBRS_CLUB = ['PMu','PMa','PJyQ','PCC','PJyM','PPM','PPJyQ','PPCyQ'];

  var canales = [];
  var pedidos = [];

  if (!_cajaMode) {

  // Home, Pilar
  ['Home', 'Pilar'].forEach(function(hoja) {
    var sh = SS.getSheetByName(hoja);
    var stats = { nombre: hoja, pedidos: 0, entregados: 0, pendientes: 0, cancelados: 0, reservados: 0, facturado: 0, cobrado: 0, noCobrado: 0 };
    if (!sh || sh.getLastRow() <= 1) { canales.push(stats); return; }
    // Lectura con tail si está pedido. data sigue siendo [headers, ...rows] para
    // mantener compat con el código existente; _rOffH = startRow - 1 corrige el
    // "r + 1" usado para guardar la fila real del Sheet.
    var _lastRowH = sh.getLastRow();
    var _lastColH = sh.getLastColumn();
    var _startRowH = (_tail > 0 && _lastRowH > _tail + 1) ? (_lastRowH - _tail + 1) : 2;
    var _rOffH = _startRowH - 1;
    var headersH = sh.getRange(1, 1, 1, _lastColH).getValues()[0];
    var data = [headersH];
    if (_lastRowH >= _startRowH) {
      data = data.concat(sh.getRange(_startRowH, 1, _lastRowH - _startRowH + 1, _lastColH).getValues());
    }

    // Índices comunes Home/Pilar v2 (abr/2026). Difieren solo en nº de productos y cols post-productos.
    var isPilarAdm = (hoja === 'Pilar');
    var IX = isPilarAdm
      ? { total:13, fact:21, ef:17, tr:18, pef:19, ptr:20, costo:45, margen:46,
          prodStart:22, prodCount:23, barrio:47, lote:48, tel:49, idxFechaEnt:52 }
      : { total:13, fact:21, ef:17, tr:18, pef:19, ptr:20, costo:41, margen:42,
          prodStart:22, prodCount:19, barrio:43, subBarrio:44, lote:45, tel:46, idxFechaEnt:49 };

    for (var r = 1; r < data.length; r++) {
      var estado = String(data[r][10]).trim();
      var estadoPago = String(data[r][12]).trim();
      var total = Number(data[r][IX.total]) || 0;
      var facturado = Number(data[r][IX.fact]) || total;
      var origen = String(data[r][8]).trim();
      var cliente = String(data[r][7]).trim();
      var nPedido = String(data[r][1]).trim();
      var diaEntrega = _fechaADiaSemana(data[r][9]);
      var diaEntregaISO = _fechaAISO(data[r][9]);
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

      var ABBRS_PILAR_ADMIN = ['PPM','PPJyQ','PPCyQ','SQB','SL','SCo','SPyP','SJyQ','SE','SCa',
                               'ECaC','EJyQ','ECyQ','EV','TG','TLC','TC','F',
                               'PMu','PMa','PJyQ','PCC','PJyM'];
      var abbrsAdm   = isPilarAdm ? ABBRS_PILAR_ADMIN : ABBRS_HOME;
      var prods = [];
      for (var p = 0; p < IX.prodCount; p++) {
        var qty = Number(data[r][IX.prodStart + p]) || 0;
        if (qty > 0) prods.push({ a: abbrsAdm[p], q: qty });
      }
      // Tartas (15/05/2026): cols al final, no consecutivas con productos viejos.
      // Home: BE-BH (idx 56-59) · Pilar: BH-BK (idx 59-62).
      var tartaStartAdm = isPilarAdm ? 59 : 56;
      var tartaAbbrsAdm = ['TP', 'TJyQ', 'TCa', 'TV'];
      for (var tp2 = 0; tp2 < 4; tp2++) {
        var qtyT2 = Number(data[r][tartaStartAdm + tp2]) || 0;
        if (qtyT2 > 0) prods.push({ a: tartaAbbrsAdm[tp2], q: qtyT2 });
      }
      // Wraps (RC, RP): cols al final tras "A Favor / Aplicado". Home idx 61-62, Pilar 64-65.
      // Sin esto un pedido SOLO de wraps (caso Simón Panelo 05/06) aparece sin productos
      // en el Panel y el editor abre vacío. (El Ruta ya se arregló en _doGetEntregas.)
      var wrapStartAdm = isPilarAdm ? 64 : 61;
      var wrapAbbrsAdm = ['RC', 'RP'];
      for (var wp2 = 0; wp2 < 2; wp2++) {
        var qtyW2 = Number(data[r][wrapStartAdm + wp2]) || 0;
        if (qtyW2 > 0) prods.push({ a: wrapAbbrsAdm[wp2], q: qtyW2 });
      }

      var formaPago = String(data[r][11] || '').trim();
      var costoPed = Number(data[r][IX.costo]) || 0;
      var margenPed = Number(data[r][IX.margen]) || 0;
      // Sub-barrio / dirección: Home usa Sub Barrio (idx 42); Pilar usa "Barrio Privado/Dirección" (idx 47) + "Domicilio/Lote" (idx 48)
      var subBarrio = isPilarAdm
        ? [String(data[r][IX.barrio] || '').trim(), String(data[r][IX.lote] || '').trim()].filter(Boolean).join(' · ')
        : String(data[r][IX.subBarrio] || '').trim();
      var hora = data[r][0];
      var horaStr = hora instanceof Date ? Utilities.formatDate(hora, 'America/Argentina/Buenos_Aires', 'HH:mm') : String(hora || '');
      var diaPedido = String(data[r][2] || '').trim();

      var efR  = Number(data[r][IX.ef]) || 0;
      var trR  = Number(data[r][IX.tr]) || 0;
      var pefR = Number(data[r][IX.pef]) || 0;
      var ptrR = Number(data[r][IX.ptr]) || 0;

      var idxFechaEnt = IX.idxFechaEnt;
      var feRaw = data[r][idxFechaEnt];
      var feStr = '', feDiaStr = '';
      if (feRaw instanceof Date) {
        feStr = Utilities.formatDate(feRaw, 'America/Argentina/Buenos_Aires', 'dd/MM');
        var DIAS_FE = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
        feDiaStr = DIAS_FE[feRaw.getDay()];
      } else if (feRaw) {
        var feParts = String(feRaw).split('/');
        if (feParts.length >= 2) feStr = String(feParts[0]).padStart(2,'0') + '/' + String(feParts[1]).padStart(2,'0');
      }

      // Fecha de Cobro (col nueva dinámica, buscar por header)
      var fcStr = '', fcDiaStr = '';
      for (var hh = 0; hh < headersH.length; hh++) {
        if (String(headersH[hh]).trim() === 'Fecha de Cobro') {
          var fcRaw = data[r][hh];
          if (fcRaw instanceof Date) {
            fcStr = Utilities.formatDate(fcRaw, 'America/Argentina/Buenos_Aires', 'dd/MM');
            var DIAS_FC = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
            fcDiaStr = DIAS_FC[fcRaw.getDay()];
          } else if (fcRaw) {
            var fcStr2 = String(fcRaw).split(' ')[0];
            var fcParts = fcStr2.split('/');
            if (fcParts.length >= 2) fcStr = String(fcParts[0]).padStart(2,'0') + '/' + String(fcParts[1]).padStart(2,'0');
          }
          break;
        }
      }

      // Origen Detalle: leer por header (col puede variar entre Home/Pilar)
      var oDetStr = '';
      for (var hd = 0; hd < headersH.length; hd++) {
        if (String(headersH[hd]).trim() === 'Origen Detalle') { oDetStr = String(data[r][hd] || '').trim(); break; }
      }

      // Descuento Detalle (receta JSON): para que el panel/editor muestren y ajusten el descuento
      var ddStr = '';
      for (var hdd = 0; hdd < headersH.length; hdd++) {
        if (String(headersH[hdd]).trim() === 'Descuento Detalle') { ddStr = String(data[r][hdd] || '').trim(); break; }
      }

      // Combo Detalle (receta JSON de combos): para que el Panel muestre "2× Combo X"
      var cdStr = '';
      for (var hcd = 0; hcd < headersH.length; hcd++) {
        if (String(headersH[hcd]).trim() === 'Combo Detalle') { cdStr = String(data[r][hcd] || '').trim(); break; }
      }

      // Repartidor: Home col 56 (BD, idx 55), Pilar col 59 (idx 58). Leer por header para resiliencia.
      var repStr = '';
      for (var hr2 = 0; hr2 < headersH.length; hr2++) {
        if (String(headersH[hr2]).trim() === 'Repartidor') { repStr = String(data[r][hr2] || '').trim(); break; }
      }

      // Barrio principal (Home: Estancias del Pilar / Los Alcanfores / Estancias del Río · Pilar: barrio privado).
      // Va en campo `bar` para que el panel pueda filtrar por barrio sin meterse con sub-barrios.
      var barStr = String(data[r][IX.barrio] || '').trim();

      // Layout v2 Home/Pilar: N(idx13)=Subtotal Producto · O(idx14)=Envío · P(idx15)=Descuento
      var subtProd = Number(data[r][13]) || 0;
      var envioPed = Number(data[r][14]) || 0;
      var descPed  = Number(data[r][15]) || 0;
      pedidos.push({
        n: nPedido, h: hoja, c: cliente, f: fechaStr, de: diaEntrega, dee: diaEntregaISO,
        es: estado, o: origen, ep: estadoPago, fp: formaPago,
        $: facturado, co: costoPed, mg: margenPed, br: subBarrio, bar: barStr, p: prods,
        subt: subtProd, env: envioPed, desc: descPed,
        hr: horaStr, dia: diaPedido,
        ef: efR, tr: trR, pef: pefR, ptr: ptrR,
        fe: feStr, fed: feDiaStr,
        fc: fcStr, fcd: fcDiaStr,
        oDet: oDetStr,
        dd: ddStr,
        cd: cdStr,
        rep: repStr,
        r: r + _rOffH
      });
    }
    canales.push(stats);
  });

  // Clubes (34 cols, estructura diferente)
  var shClubes = SS.getSheetByName('Clubes');
  var statsClubes = { nombre: 'Clubes', pedidos: 0, entregados: 0, pendientes: 0, cancelados: 0, reservados: 0, facturado: 0, cobrado: 0, noCobrado: 0 };
  if (shClubes && shClubes.getLastRow() > 1) {
    var _lastRowC = shClubes.getLastRow();
    var _lastColC = shClubes.getLastColumn();
    var _startRowC = (_tail > 0 && _lastRowC > _tail + 1) ? (_lastRowC - _tail + 1) : 2;
    var _rOffC = _startRowC - 1;
    var headersClub = shClubes.getRange(1, 1, 1, _lastColC).getValues()[0];
    var dataClubes = [headersClub];
    if (_lastRowC >= _startRowC) {
      dataClubes = dataClubes.concat(shClubes.getRange(_startRowC, 1, _lastRowC - _startRowC + 1, _lastColC).getValues());
    }
    for (var rc = 1; rc < dataClubes.length; rc++) {
      var estadoC = String(dataClubes[rc][13]).trim();
      var estadoPagoC = String(dataClubes[rc][15]).trim();
      var totalC = Number(dataClubes[rc][16]) || 0;
      var facturadoC = Number(dataClubes[rc][22]) || totalC;
      var origenC = String(dataClubes[rc][11]).trim();
      var clienteC = String(dataClubes[rc][7]).trim();
      var clubC = String(dataClubes[rc][8]).trim();
      var nPedidoC = String(dataClubes[rc][1]).trim();
      var diaEntregaC = _fechaADiaSemana(dataClubes[rc][12]);
      var diaEntregaISOC = _fechaAISO(dataClubes[rc][12]);
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
      // Empanadas Clubes (28/05/2026): cols AL-AO = idx 37-40
      var ABBRS_CLUB_EMP2 = ['ECaC','EJyQ','ECyQ','EV'];
      for (var ec2 = 0; ec2 < 4; ec2++) {
        var qtyEc2 = Number(dataClubes[rc][37 + ec2]) || 0;
        if (qtyEc2 > 0) prodsC.push({ a: ABBRS_CLUB_EMP2[ec2], q: qtyEc2 });
      }

      var formaPagoC = String(dataClubes[rc][14] || '').trim();

      var costoCl = Number(dataClubes[rc][31]) || 0;
      var margenCl = Number(dataClubes[rc][32]) || 0;
      var horaC = dataClubes[rc][0];
      var horaStrC = horaC instanceof Date ? Utilities.formatDate(horaC, 'America/Argentina/Buenos_Aires', 'HH:mm') : String(horaC || '');
      var diaPedC = String(dataClubes[rc][2] || '').trim();

      var efC  = Number(dataClubes[rc][18]) || 0; // S = Efectivo
      var trC  = Number(dataClubes[rc][19]) || 0; // T = Transferencia
      var pefC = Number(dataClubes[rc][20]) || 0; // U = Propina Ef
      var ptrC = Number(dataClubes[rc][21]) || 0; // V = Propina Trans

      // Fecha de Cobro (col dinámica)
      var fcStrC = '', fcDiaStrC = '';
      for (var hhC = 0; hhC < headersClub.length; hhC++) {
        if (String(headersClub[hhC]).trim() === 'Fecha de Cobro') {
          var fcRawC = dataClubes[rc][hhC];
          if (fcRawC instanceof Date) {
            fcStrC = Utilities.formatDate(fcRawC, 'America/Argentina/Buenos_Aires', 'dd/MM');
            var DIAS_FCC = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
            fcDiaStrC = DIAS_FCC[fcRawC.getDay()];
          } else if (fcRawC) {
            var fcStr2C = String(fcRawC).split(' ')[0];
            var fcPartsC = fcStr2C.split('/');
            if (fcPartsC.length >= 2) fcStrC = String(fcPartsC[0]).padStart(2,'0') + '/' + String(fcPartsC[1]).padStart(2,'0');
          }
          break;
        }
      }

      var oDetStrC = '';
      for (var hdC = 0; hdC < headersClub.length; hdC++) {
        if (String(headersClub[hdC]).trim() === 'Origen Detalle') { oDetStrC = String(dataClubes[rc][hdC] || '').trim(); break; }
      }

      // Repartidor (Clubes col 37, idx 36). Leer por header para resiliencia.
      var repStrC = '';
      for (var hrC = 0; hrC < headersClub.length; hrC++) {
        if (String(headersClub[hrC]).trim() === 'Repartidor') { repStrC = String(dataClubes[rc][hrC] || '').trim(); break; }
      }

      // Clubes: R(idx17) = Envío. Sin descuento ni subtotal separados (subt = total - env).
      var envioClub = Number(dataClubes[rc][17]) || 0;
      pedidos.push({
        n: nPedidoC, h: 'Clubes',
        c: clienteC + (clubC ? ' (' + clubC + ')' : ''),
        f: fechaStrC, de: diaEntregaC, dee: diaEntregaISOC, es: estadoC, o: origenC,
        ep: estadoPagoC, fp: formaPagoC, $: facturadoC,
        co: costoCl, mg: margenCl, br: clubC, p: prodsC,
        env: envioClub, desc: 0, subt: Math.max(0, totalC - envioClub),
        hr: horaStrC, dia: diaPedC,
        ef: efC, tr: trC, pef: pefC, ptr: ptrC,
        fc: fcStrC, fcd: fcDiaStrC,
        oDet: oDetStrC,
        rep: repStrC,
        r: rc + _rOffC
      });
    }
  }
  canales.push(statsClubes);

  // Red (55 cols, estructura propia). Layout V(22)→AR(44): debe coincidir con RED_COL_TO_ABBR.
  var ABBRS_RED_DASH = ['PPM','PPJyQ','PPCyQ','SQB','SL','SCo','SPyP','SJyQ','SE','SCa',
                        'ECaC','EJyQ','ECyQ','EV','TG','TLC','TC','F','PMu','PMa','PJyQ','PCC','PJyM'];
  var shRedDash = SS.getSheetByName('Red');
  var statsRed = { nombre: 'Red', pedidos: 0, entregados: 0, pendientes: 0, cancelados: 0, reservados: 0, facturado: 0, cobrado: 0, noCobrado: 0 };
  if (shRedDash && shRedDash.getLastRow() > 1) {
    var _lastRowR = shRedDash.getLastRow();
    var _lastColR = shRedDash.getLastColumn();
    var _startRowR = (_tail > 0 && _lastRowR > _tail + 1) ? (_lastRowR - _tail + 1) : 2;
    var _rOffR = _startRowR - 1;
    var _headersRedDash = shRedDash.getRange(1, 1, 1, _lastColR).getValues()[0];
    var dataRed = [_headersRedDash];
    if (_lastRowR >= _startRowR) {
      dataRed = dataRed.concat(shRedDash.getRange(_startRowR, 1, _lastRowR - _startRowR + 1, _lastColR).getValues());
    }
    for (var rr = 1; rr < dataRed.length; rr++) {
      var estadoR = String(dataRed[rr][11]).trim(); // L = Estado Entrega
      var estadoPagoR = String(dataRed[rr][13]).trim(); // N = Estado Pago
      var totalR = Number(dataRed[rr][14]) || 0; // O = Total
      var facturadoR = Number(dataRed[rr][20]) || totalR; // U = Facturado (bruto)
      var aPagarR = Number(dataRed[rr][48]) || Math.round(facturadoR * 83 / 100); // AW = A Pagar a Maleu (neto post-comisión). Es lo que Maleu factura en Red.
      var origenR = String(dataRed[rr][9]).trim(); // J = Origen
      var clienteR = String(dataRed[rr][8]).trim(); // I = Cliente
      var vendedorR = String(dataRed[rr][7]).trim(); // H = Vendedor
      var nPedidoR = String(dataRed[rr][1]).trim();
      var diaEntregaR = _fechaADiaSemana(dataRed[rr][10]); // K
      var diaEntregaISOR = _fechaAISO(dataRed[rr][10]);
      var fechaR = dataRed[rr][3];
      var fechaStrR = fechaR instanceof Date ? Utilities.formatDate(fechaR, 'America/Argentina/Buenos_Aires', 'dd/MM') : String(fechaR || '');

      statsRed.pedidos++;
      if (estadoR === 'Entregado') statsRed.entregados++;
      else if (estadoR === 'Cancelado') statsRed.cancelados++;
      else if (estadoR === 'Reservado') statsRed.reservados++;
      else statsRed.pendientes++;

      // Liquidación de Marcos a Maleu: cols BA/BB/BC (52/53/54). Es la plata
      // que efectivamente entra a la caja de Maleu, no lo que el cliente final
      // le paga a Marcos (cols 12-17). Para el resumen de cobros del panel
      // usamos estos valores como ep/fp/fc del pedido.
      var fpMaleuR     = String(dataRed[rr][52] || '').trim();
      var estadoMaleuR = String(dataRed[rr][53] || '').trim();
      var epMaleuR     = (estadoMaleuR === 'Pagado') ? 'Cobrado' : 'No Cobrado';
      var fcRawR       = dataRed[rr][54];
      var fcMaleuR     = '';
      if (fcRawR) {
        if (fcRawR instanceof Date) {
          fcMaleuR = Utilities.formatDate(fcRawR, 'America/Argentina/Buenos_Aires', 'dd/MM');
        } else {
          var mFc = String(fcRawR).match(/^(\d{1,2})\/(\d{1,2})/);
          if (mFc) fcMaleuR = ('0'+mFc[1]).slice(-2) + '/' + ('0'+mFc[2]).slice(-2);
        }
      }

      statsRed.facturado += aPagarR;
      if (epMaleuR === 'Cobrado') statsRed.cobrado += aPagarR;
      else statsRed.noCobrado += aPagarR;

      var prodsR = [];
      for (var prr = 0; prr < 23; prr++) {
        var qtyR = Number(dataRed[rr][21 + prr]) || 0;
        if (qtyR > 0) prodsR.push({ a: ABBRS_RED_DASH[prr], q: qtyR });
      }
      // Tartas Red: cols BF-BI (idx 57-60), al final tras las cols de control.
      // No son consecutivas con el bloque de productos viejo, por eso van aparte.
      var TARTA_ABBRS_RED_DASH = ['TP', 'TJyQ', 'TCa', 'TV'];
      for (var ptr = 0; ptr < 4; ptr++) {
        var qtyRT = Number(dataRed[rr][57 + ptr]) || 0;
        if (qtyRT > 0) prodsR.push({ a: TARTA_ABBRS_RED_DASH[ptr], q: qtyRT });
      }
      // Wraps Red: cols BJ-BK (idx 61-62). SIN esto el Panel mostraba el pedido
      // sin el wrap → caso Alejo Acuña 17/07: Fini pidió Wrap Carne, el Panel no lo
      // mostró, se descartó por error y quedó mal con el cliente. NUNCA omitir productos.
      var WRAP_ABBRS_RED_DASH = ['RC', 'RP'];
      for (var pwr = 0; pwr < 2; pwr++) {
        var qtyRW = Number(dataRed[rr][61 + pwr]) || 0;
        if (qtyRW > 0) prodsR.push({ a: WRAP_ABBRS_RED_DASH[pwr], q: qtyRW });
      }

      var formaPagoR = String(dataRed[rr][12] || '').trim();
      var costoR = Number(dataRed[rr][44]) || 0;
      // Margen NETO (col AV/48, índice 47) = Margen Bruto − Comisión 17%. Es el margen
      // real de Maleu en Red: el Facturado de la hoja es BRUTO, pero acá mandamos $ neto
      // (A Pagar, 83%), así que el margen también debe ser neto. Antes leía índice 45
      // (col AT Margen Bruto) e ignoraba la comisión → margen Red inflado ~$709k/año.
      var margenR = Number(dataRed[rr][47]) || 0;

      var efRd  = Number(dataRed[rr][16]) || 0; // Q = Efectivo
      var trRd  = Number(dataRed[rr][17]) || 0; // R = Transferencia
      var pefRd = Number(dataRed[rr][18]) || 0; // S = Propina Ef
      var ptrRd = Number(dataRed[rr][19]) || 0; // T = Propina Trans

      var horaRd = dataRed[rr][0]; // A = Hora
      var horaStrRd = horaRd instanceof Date ? Utilities.formatDate(horaRd, 'America/Argentina/Buenos_Aires', 'HH:mm') : String(horaRd || '');
      var diaPedRd = String(dataRed[rr][2] || '').trim(); // C = Día

      // Origen Detalle Red col idx 55 (BD)
      var oDetStrR = String(dataRed[rr][55] || '').trim();

      // Red: P(idx15) = Envío. Sin descuento (Red no lleva descuento). subt = totalR - env.
      var envioRed = Number(dataRed[rr][15]) || 0;
      pedidos.push({
        n: nPedidoR, h: 'Red',
        c: clienteR + (vendedorR ? ' (Red: ' + vendedorR + ')' : ''),
        f: fechaStrR, de: diaEntregaR, dee: diaEntregaISOR, es: estadoR, o: origenR,
        // ep/fp/fc reflejan la liquidación de Marcos a Maleu (BA/BB/BC), no el
        // cobro del cliente final a Marcos. Es lo que entra a la caja de Maleu.
        ep: epMaleuR, fp: fpMaleuR || formaPagoR, fc: fcMaleuR, $: aPagarR,
        co: costoR, mg: margenR, br: vendedorR, p: prodsR,
        env: envioRed, desc: 0, subt: Math.max(0, totalR - envioRed),
        hr: horaStrRd, dia: diaPedRd,
        ef: efRd, tr: trRd, pef: pefRd, ptr: ptrRd,
        oDet: oDetStrR,
        retira: vendedorR,
        r: rr + _rOffR
      });
    }
  }
  canales.push(statsRed);

  } // fin if (!_cajaMode) — saltea el loop principal de pedidos en mode caja

  // ── pedidosLight: short-circuit para refresh rápido del tab Pedidos ──
  // Devuelve solo lo que el frontend necesita para re-renderizar el tab Pedidos
  // (pedidos[] y canales[] para los pills de filtro). Skipea: vendedoresDeuda,
  // stock, OC, caja, ingresos, egresos, movimientos, pagos red, cobros parciales.
  // Bajamos de ~12s a ~3s en el endpoint.
  if (_lightOnly) {
    var _outObj = { ts: Date.now(), pedidos: pedidos, light: true };
    if (_tail > 0) {
      _outObj.tail = _tail;
      // canales[] queda fuera en modo tail (stats no representativos). El frontend mantiene los anteriores.
    } else {
      _outObj.canales = canales;
      // Salud de Clientes Home (identidad canónica del CRM). Solo en modo no-tail:
      // necesita el historial completo de Home, que el tail no trae.
      _outObj.saludHome = _saludHomeSemana(argNow);
    }
    var _outLight = JSON.stringify(_outObj);
    if (!_tail) {
      try { CacheService.getScriptCache().put('admin_light_v1', _outLight, 5); } catch (e) {}
    }
    return ContentService.createTextOutput(_outLight).setMimeType(ContentService.MimeType.JSON);
  }

  // Defaults para mode='caja' — si no necesitamos vendedores/stock/OC, los
  // dejamos vacíos y saltamos los bloques pesados de abajo (vendedoresDeuda,
  // liquidaciones, totales, stock, OC). Las var-declaraciones internas son
  // function-scoped y no causan error si se redeclaran al volver a admin completo.
  var vendedoresArr = [];
  var totales = { pedidos: 0, entregados: 0, pendientes: 0, cancelados: 0, reservados: 0, facturado: 0, cobrado: 0, noCobrado: 0, ticket: 0 };
  var stock = [];
  var ocPend = 0, ocTotal = 0, ocLista = [];

  if (!_cajaMode) {

  // ── Vendedores Red: deuda pendiente a Maleu ──
  // Lee la hoja Red. Por cada pedido NO pagado a Maleu, acumula
  // deuda = Total × (1 - comision/100) según la comisión del vendedor.
  var vendedoresDeuda = {};
  var comisionPorVendedor = {};
  var waPorVendedor = {};
  var shVend = SS.getSheetByName('Vendedores');
  if (shVend && shVend.getLastRow() > 1) {
    var dV = shVend.getDataRange().getValues();
    for (var rv = 1; rv < dV.length; rv++) {
      var nombreV = String(dV[rv][0] || '').trim();
      if (!nombreV) continue;
      var waV = String(dV[rv][1] || '').trim();
      var comV = Number(dV[rv][9]) || 17;
      var estadoV = String(dV[rv][6] || '').trim();
      comisionPorVendedor[nombreV] = comV;
      waPorVendedor[nombreV] = waV;
      if (estadoV === 'Activo') {
        vendedoresDeuda[nombreV] = { nombre: nombreV, wa: waV, com: comV, estado: estadoV, deuda: 0, pedidos: [], deudaPorSemana: {}, liquidaciones: [] };
      }
    }
  }

  function _procesarRedPedidos(sh, origen) {
    if (!sh || sh.getLastRow() <= 1) return;
    var d = sh.getDataRange().getValues();
    var headersR = d[0];
    // Buscar índices por header (para ser robusto)
    var idxVendedor = -1, idxPedido = -1, idxFecha = -1, idxCliente = -1, idxTotal = -1;
    var idxEstadoPagoMaleu = -1, idxFechaPagoMaleu = -1, idxFormaPagoMaleu = -1;
    var idxSemana = -1, idxAnio = -1, idxAPagar = -1;
    for (var hx = 0; hx < headersR.length; hx++) {
      var nm = String(headersR[hx]).trim();
      if (nm === 'Vendedor') idxVendedor = hx;
      else if (nm === 'N° Pedido' || nm === 'N°') idxPedido = hx;
      else if (nm === 'Fecha') idxFecha = hx;
      else if (nm === 'Cliente') idxCliente = hx;
      else if (nm === 'Total ($)' || nm === 'Total') idxTotal = hx;
      else if (nm === 'Estado Pago a Maleu') idxEstadoPagoMaleu = hx;
      else if (nm === 'Fecha Pago a Maleu') idxFechaPagoMaleu = hx;
      else if (nm === 'Forma Pago a Maleu') idxFormaPagoMaleu = hx;
      else if (nm === 'Semana') idxSemana = hx;
      else if (nm === 'Año' || nm === 'Anio') idxAnio = hx;
      else if (nm === 'A Pagar') idxAPagar = hx;
    }
    if (idxVendedor < 0 || idxTotal < 0 || idxEstadoPagoMaleu < 0) return;

    for (var ri = 1; ri < d.length; ri++) {
      var row = d[ri];
      var vend = String(row[idxVendedor] || '').trim();
      if (!vend) continue;
      var estPagoM = String(row[idxEstadoPagoMaleu] || '').trim();
      if (estPagoM === 'Sí' || estPagoM === 'Si' || estPagoM === 'Pagado') continue;
      var totP = Number(row[idxTotal]) || 0;
      if (totP <= 0) continue;
      var comPct = comisionPorVendedor[vend] !== undefined ? comisionPorVendedor[vend] : 17;
      // Preferir "A Pagar" real por pedido — respeta comisiones por fila (caso
      // R-45/R-46 con comisión 0 a familia del vendedor). Fallback al cálculo
      // plano si la columna no existe o está vacía.
      var aPagarP = idxAPagar >= 0 ? (Number(row[idxAPagar]) || 0) : 0;
      var deudaP = aPagarP > 0 ? aPagarP : totP * (1 - comPct/100);
      if (!vendedoresDeuda[vend]) {
        vendedoresDeuda[vend] = { nombre: vend, wa: waPorVendedor[vend] || '', com: comPct, estado: 'Inactivo', deuda: 0, pedidos: [], deudaPorSemana: {}, liquidaciones: [] };
      }
      if (!vendedoresDeuda[vend].deudaPorSemana) vendedoresDeuda[vend].deudaPorSemana = {};
      vendedoresDeuda[vend].deuda += deudaP;
      var fR = row[idxFecha];
      var fechaStrR = '';
      if (fR instanceof Date) fechaStrR = Utilities.formatDate(fR, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
      else fechaStrR = String(fR || '');
      // Atribuir deuda a la semana del pedido (entrega ≈ misma semana en Red)
      var semP = idxSemana >= 0 ? (Number(row[idxSemana]) || 0) : 0;
      var anioP = idxAnio >= 0 ? (Number(row[idxAnio]) || 0) : 0;
      if (semP > 0) {
        var keySem = anioP > 0 ? (anioP + '-' + semP) : String(semP);
        vendedoresDeuda[vend].deudaPorSemana[keySem] = (vendedoresDeuda[vend].deudaPorSemana[keySem] || 0) + deudaP;
      }
      vendedoresDeuda[vend].pedidos.push({
        n: String(row[idxPedido] || '').trim(),
        f: fechaStrR,
        c: String(row[idxCliente] || '').trim(),
        total: totP,
        deuda: deudaP,
        sem: semP,
        anio: anioP,
        origen: origen,
        row: ri + 1
      });
    }
  }
  _procesarRedPedidos(SS.getSheetByName('Red'), 'Red');

  // ── Liquidaciones de "Pagos Red Liq" por vendedor (para mostrar "te acaba de pagar" en Inicio) ──
  var shLiqAdm = SS.getSheetByName('Pagos Red Liq');
  if (shLiqAdm && shLiqAdm.getLastRow() > 1) {
    var dLA = shLiqAdm.getDataRange().getValues();
    // Layout: A=Timestamp, B=Fecha, C=Vendedor, D=Año, E=Semana, F=Efectivo, G=Transferencia, H=Total
    var ahoraMs = Date.now();
    for (var rLA = 1; rLA < dLA.length; rLA++) {
      var tsLA = dLA[rLA][0];
      if (!tsLA) continue;
      var vendLA = String(dLA[rLA][2] || '').trim();
      if (!vendLA) continue;
      if (!vendedoresDeuda[vendLA]) {
        var comLA = comisionPorVendedor[vendLA] !== undefined ? comisionPorVendedor[vendLA] : 17;
        vendedoresDeuda[vendLA] = { nombre: vendLA, wa: waPorVendedor[vendLA] || '', com: comLA, estado: 'Inactivo', deuda: 0, pedidos: [], deudaPorSemana: {}, liquidaciones: [] };
      }
      if (!vendedoresDeuda[vendLA].liquidaciones) vendedoresDeuda[vendLA].liquidaciones = [];
      var dtLA = (tsLA instanceof Date) ? tsLA : new Date(tsLA);
      var tsMsLA = isNaN(dtLA.getTime()) ? 0 : dtLA.getTime();
      var efLA = Number(dLA[rLA][5]) || 0;
      var mpLA = Number(dLA[rLA][6]) || 0;
      var totLA = Number(dLA[rLA][7]) || (efLA + mpLA);
      vendedoresDeuda[vendLA].liquidaciones.push({
        ts: tsMsLA,
        f: tsMsLA ? Utilities.formatDate(dtLA, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm') : String(tsLA || ''),
        anio: Number(dLA[rLA][3]) || 0,
        sem: Number(dLA[rLA][4]) || 0,
        ef: efLA,
        mp: mpLA,
        tot: totLA,
        reciente: tsMsLA > 0 && (ahoraMs - tsMsLA) < (48 * 3600 * 1000)
      });
    }
    // Ordenar liquidaciones por ts desc (más reciente arriba)
    Object.keys(vendedoresDeuda).forEach(function(k){
      var arr = vendedoresDeuda[k].liquidaciones;
      if (arr && arr.length > 1) arr.sort(function(a, b){ return (b.ts || 0) - (a.ts || 0); });
    });
  }

  var vendedoresArr = Object.keys(vendedoresDeuda).map(function(k) {
    return vendedoresDeuda[k];
  }).sort(function(a, b) { return b.deuda - a.deuda; });

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
    var colIni = -1, colVen = -1, colComp = -1, colInv = -1;
    for (var h = 0; h < hdr.length; h++) {
      var hv = String(hdr[h]).replace(/\n/g,' ').trim().toLowerCase();
      if (hv === 'producto' || hv === 'nombre') colNombre = h;
      if (hv === 'abreviatura') colAbbr = h;
      if (hv === 'stock físico' || hv === 'stock fisico') colFisico = h;
      if (hv === 'reservado') colReservado = h;
      if (hv === 'stock disponible' || hv === 'disponible') colDisp = h;
      if (hv === 'stock inicial semana' || hv === 'inicial semana') colIni = h;
      if (hv === 'vendidos semana') colVen = h;
      if (hv === 'comprado semana') colComp = h;
      if (hv === 'invertido') colInv = h;
    }
    for (var rp = 1; rp < dataProd.length; rp++) {
      var nombre = colNombre >= 0 ? String(dataProd[rp][colNombre]).trim() : '';
      var abbr = colAbbr >= 0 ? String(dataProd[rp][colAbbr]).trim() : '';
      if (!nombre && !abbr) continue;
      var fisico = colFisico >= 0 ? (Number(dataProd[rp][colFisico]) || 0) : 0;
      var reservado = colReservado >= 0 ? (Number(dataProd[rp][colReservado]) || 0) : 0;
      var disponible = colDisp >= 0 ? (Number(dataProd[rp][colDisp]) || 0) : fisico - reservado;
      var inicialSem = colIni >= 0 ? (Number(dataProd[rp][colIni]) || 0) : 0;
      var vendidosSem = colVen >= 0 ? (Number(dataProd[rp][colVen]) || 0) : 0;
      var compradoSem = colComp >= 0 ? (Number(dataProd[rp][colComp]) || 0) : 0;
      // Precio y costo (cols I=8, J=9 0-based) — usados por el editor de pedidos para defaults
      var precio = parseFloat(String(dataProd[rp][8] || '').replace(/[$.]/g,'').replace(/,/g,'')) || 0;
      var costoU = parseFloat(String(dataProd[rp][9] || '').replace(/[$.]/g,'').replace(/,/g,'')) || 0;
      var invertido = colInv >= 0 ? (parseFloat(String(dataProd[rp][colInv] || '').replace(/[$.]/g,'').replace(/,/g,'')) || 0) : (fisico * costoU);
      stock.push({ n: nombre, a: abbr, f: fisico, r: reservado, d: disponible, p: precio, co: costoU, i: inicialSem, v: vendidosSem, c: compradoSem, iv: invertido });
    }
  }

  // ── OC pendientes + lista compacta ──
  // lista: usado por el Panel para mostrar el badge "✓ Listo" / "📦 N OC"
  // en cada pedido (saber si todas las OCs vinculadas ya están Recibidas).
  var ocPend = 0;
  var ocTotal = 0;
  var ocLista = [];
  var shOC = SS.getSheetByName('Orden de Compra');
  if (shOC && shOC.getLastRow() > 1) {
    var dataOC = shOC.getDataRange().getValues();
    for (var ro = 1; ro < dataOC.length; ro++) {
      var estOC = String(dataOC[ro][20]).trim();
      if (estOC === 'Pendiente' || estOC === 'Pedido') {
        ocPend++;
        ocTotal += Number(dataOC[ro][14]) || 0;
      }
      // Filas válidas: con canal y pedido vinculado a un canal real
      var canalOC = String(dataOC[ro][4] || '').trim();
      var pedidoOC = String(dataOC[ro][5] || '').trim();
      if (canalOC && pedidoOC) {
        ocLista.push({
          canal: canalOC,
          pedido: pedidoOC,
          cliente: String(dataOC[ro][6] || '').trim(), // col G: fallback matching cuando N° pedido queda desalineado
          abbr: String(dataOC[ro][11] || '').trim(),
          q: Number(dataOC[ro][12]) || 0,
          proveedor: String(dataOC[ro][9] || '').trim(),
          estado: estOC
        });
      }
    }
  }

  } // fin if (!_cajaMode)

  // ── Egresos (gastos) ──
  // gastos[] y totalGastos incluyen TODO (para listados históricos/P&L).
  // gastosEf/gastosMP incluyen SOLO los posteriores al último Saldo Base (para el saldo vivo).
  var gastos = [];
  var totalGastos = 0;
  var gastosEf = 0;
  var gastosMP = 0;
  var shEgr = SS.getSheetByName('Egresos');
  if (shEgr && shEgr.getLastRow() > 1) {
    var dataEgr = shEgr.getDataRange().getValues();
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
      if (_afterSaldo(fechaE)) {
        if (metodoE === 'Efectivo') gastosEf += montoE;
        else gastosMP += montoE;
      }
      // Timestamp completo (con hora si existe) para ordenar bien intra-día.
      var fDtE = _parseDateAny(fechaE);
      var tsE = fDtE ? fDtE.getTime() : 0;
      var fStrFullE = fDtE ? Utilities.formatDate(fDtE, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm') : fechaStrE;

      gastos.push({
        f: fechaStrE,
        fFull: fStrFullE,
        ts: tsE,
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

  // ── Cobrado por método de pago — SOLO registros posteriores al último Saldo Base ──
  // Usa los montos REALES en R (Efectivo) y S (Transferencia). Estas cols YA incluyen propina
  // (se setean con el monto TOTAL entrado por ese método). NO sumar T/U por separado: duplica.
  // Esto equivale a usar la col "Facturado" (V en Home/Pilar, W en Clubes) discriminada por método.
  var cobradoEf = 0;
  var cobradoMP = 0;
  // cobrosDetalle acumula TODOS los cobros (post-saldoBase en los que cuentan, + histórico
  // para el libro diario unificado en la tab Caja del Panel). Campos: f (fecha cobro string),
  // fTs (timestamp para sort), h (hoja), id, c (cliente), efTotal, mpTotal.
  var cobrosDetalle = [];

  // ── Cobros Parciales: cargar y sumar por separado ──
  // Cada parcial cuenta a caja viva inmediatamente (filtrado por saldoBase).
  // Cuando el pedido cierre como Cobrado, _sumarCobrado descuenta los parciales ya sumados
  // para no duplicar.
  var parcialesPorPed = {}; // { 'Hoja|id': { ef:0, mp:0, total:0 } }
  var shCP_caja = SS.getSheetByName('Cobros Parciales');
  if (shCP_caja && shCP_caja.getLastRow() > 1) {
    var dCP_caja = shCP_caja.getDataRange().getValues();
    for (var rcp_caja = 1; rcp_caja < dCP_caja.length; rcp_caja++) {
      var cpHoja = String(dCP_caja[rcp_caja][1] || '').trim();
      var cpId   = String(dCP_caja[rcp_caja][2] || '').trim();
      if (!cpHoja || !cpId) continue;
      var cpFp    = String(dCP_caja[rcp_caja][4] || '').trim();
      var cpMonto = Number(dCP_caja[rcp_caja][5]) || 0;
      var cpFecha = dCP_caja[rcp_caja][0];
      var cpCli   = String(dCP_caja[rcp_caja][3] || '').trim();
      if (cpMonto <= 0) continue;
      var keyCP = cpHoja + '|' + cpId;
      if (!parcialesPorPed[keyCP]) parcialesPorPed[keyCP] = { ef:0, mp:0, total:0 };
      if (cpFp === 'Efectivo') parcialesPorPed[keyCP].ef += cpMonto;
      else parcialesPorPed[keyCP].mp += cpMonto;
      parcialesPorPed[keyCP].total += cpMonto;

      var fDtCP = _parseDateAny(cpFecha);
      var fStrCP = fDtCP ? Utilities.formatDate(fDtCP, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm') : String(cpFecha || '');
      cobrosDetalle.push({
        f: fStrCP, fTs: fDtCP ? fDtCP.getTime() : 0,
        h: cpHoja, id: cpId + ' (parcial)', c: cpCli,
        efTotal: cpFp === 'Efectivo' ? cpMonto : 0,
        mpTotal: cpFp !== 'Efectivo' ? cpMonto : 0
      });

      // Sumar a caja viva sólo si es posterior al saldoBase
      if (saldoBase.fechaDate && !_afterSaldo(cpFecha)) continue;
      if (cpFp === 'Efectivo') cobradoEf += cpMonto;
      else cobradoMP += cpMonto;
    }
  }
  function _sumarCobrado(hojaName, colsC) {
    var shC = SS.getSheetByName(hojaName);
    if (!shC || shC.getLastRow() <= 1) return;
    var dC = shC.getDataRange().getValues();
    var hdrs = dC[0];
    var idxFc = -1;
    for (var hh = 0; hh < hdrs.length; hh++) {
      if (String(hdrs[hh]).trim() === 'Fecha de Cobro') { idxFc = hh; break; }
    }
    // Fecha entrega: col dinámica según hoja (fallback si fc vacío)
    var idxFeEnt = -1;
    for (var he = 0; he < hdrs.length; he++) {
      if (String(hdrs[he]).trim() === 'Fecha Entrega') { idxFeEnt = he; break; }
    }
    for (var rc2 = 1; rc2 < dC.length; rc2++) {
      if (String(dC[rc2][colsC.ep]).trim() !== 'Cobrado') continue;
      var ef = Number(dC[rc2][colsC.ef]) || 0;
      var tr = Number(dC[rc2][colsC.tr]) || 0;
      var facturado = Number(dC[rc2][colsC.fac]) || 0;
      var fpVal = String(dC[rc2][colsC.fp] || '').trim();
      // Fecha del cobro: fc → fe entrega → fecha pedido
      var fcRaw = idxFc >= 0 ? dC[rc2][idxFc] : '';
      var feEntRaw = idxFeEnt >= 0 ? dC[rc2][idxFeEnt] : '';
      var fPedRaw = dC[rc2][3];
      var fUsar = fcRaw || feEntRaw || fPedRaw;
      var fDt = _parseDateAny(fUsar);
      var fStr = fDt ? Utilities.formatDate(fDt, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm') : String(fUsar || '');
      var cliente = String(dC[rc2][colsC.cli] || '').trim();
      var pedidoId = String(dC[rc2][1] || '').trim();
      // Cobro = col Facturado (V en Home/Pilar, W en Clubes).
      // Se asigna al medio según Forma de Pago del pedido.
      // DESCONTAR parciales ya contabilizados para no duplicar.
      var parcAcum = parcialesPorPed[hojaName + '|' + pedidoId];
      var alreadyCounted = parcAcum ? (parcAcum.ef + parcAcum.mp) : 0;
      var netoCierre = facturado - alreadyCounted;
      var efTotal = 0, mpTotal = 0;
      if (netoCierre > 0) {
        if (fpVal === 'Efectivo') efTotal = netoCierre;
        else if (fpVal === 'Mixto') {
          // Bug Gonzalo 05/06: fp=Mixto mandaba TODO el facturado a MP en libro diario.
          // Usar la composición real de las cols Ef/Tr del pedido, escalando por si hubo parciales.
          if (facturado > 0 && (ef + tr) > 0) {
            efTotal = Math.round(netoCierre * (ef / (ef + tr)));
            mpTotal = netoCierre - efTotal;
          } else {
            mpTotal = netoCierre;
          }
        } else {
          mpTotal = netoCierre;
        }
      }
      if (efTotal > 0 || mpTotal > 0) {
        cobrosDetalle.push({
          f: fStr, fTs: fDt ? fDt.getTime() : 0,
          h: hojaName, id: pedidoId, c: cliente,
          efTotal: efTotal, mpTotal: mpTotal
        });
      }
      // Filtrar por saldoBase SOLO para sumar a caja viva.
      if (saldoBase.fechaDate) {
        if (idxFc < 0) continue;
        if (!_afterSaldo(fcRaw)) continue;
      }
      cobradoEf += efTotal;
      cobradoMP += mpTotal;
    }
  }
  // Home/Pilar v2: L=11=FP, M=12=EP, R=17=Efectivo, S=18=Transferencia, V=21=Facturado, H=7=Cliente
  _sumarCobrado('Home',   { ep: 12, fp: 11, ef: 17, tr: 18, fac: 21, cli: 7 });
  _sumarCobrado('Pilar',  { ep: 12, fp: 11, ef: 17, tr: 18, fac: 21, cli: 7 });
  // Clubes: O=14=FP, P=15=EP, S=18=Efectivo, T=19=Transferencia, W=22=Facturado, H=7=Cliente
  _sumarCobrado('Clubes', { ep: 15, fp: 14, ef: 18, tr: 19, fac: 22, cli: 7 });
  // Red: NO se suma desde la hoja Red (esa plata está en el bolsillo del vendedor, no en caja Maleu).
  // La caja de Maleu se alimenta de "Pagos Red Liq" cuando el vendedor liquida (ver _sumarPagosRedLiq abajo).
  _sumarPagosRedLiq();

  function _sumarPagosRedLiq() {
    var shL = SS.getSheetByName('Pagos Red Liq');
    if (!shL || shL.getLastRow() <= 1) return;
    var dL = shL.getDataRange().getValues();
    // Layout: A=Timestamp, B=Fecha, C=Vendedor, D=Año, E=Semana, F=Efectivo, G=Transferencia, H=Total, ...
    for (var rL = 1; rL < dL.length; rL++) {
      var tsL = dL[rL][0];
      if (!tsL) continue;
      var efL = Number(dL[rL][5]) || 0;
      var mpL = Number(dL[rL][6]) || 0;
      // Siempre agregar al detalle (histórico), filtrar por saldoBase solo para sumar a caja viva.
      var fDtL = _parseDateAny(tsL);
      if (efL > 0 || mpL > 0) {
        cobrosDetalle.push({
          f: fDtL ? Utilities.formatDate(fDtL, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm') : String(tsL || ''),
          fTs: fDtL ? fDtL.getTime() : 0,
          h: 'Red', id: 'Liq', c: String(dL[rL][2] || 'Liquidación Red').trim(),
          efTotal: efL, mpTotal: mpL
        });
      }
      if (saldoBase.fechaDate && !_afterSaldo(tsL)) continue;
      cobradoEf += efL;
      cobradoMP += mpL;
    }
  }

  // ── Ingresos manuales ──
  // ingresos[] incluye TODO. ingresosEf/MP solo los posteriores al Saldo Base (para el saldo vivo).
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
      if (_afterSaldo(fechaI)) {
        if (metodoI === 'Efectivo') ingresosEf += montoI;
        else ingresosMP += montoI;
      }
      var fDtI = _parseDateAny(fechaI);
      var tsI = fDtI ? fDtI.getTime() : 0;
      var fStrFullI = fDtI ? Utilities.formatDate(fDtI, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm') : fechaStrI;
      ingresos.push({
        f: fechaStrI, fFull: fStrFullI, ts: tsI,
        cat: String(dataIng[ri][3] || '').trim(),
        con: String(dataIng[ri][4] || '').trim(), met: metodoI,
        $: montoI, not: String(dataIng[ri][7] || '').trim()
      });
    }
  }

  // ── Vueltos desde la billetera (para el desglose por día del Resumen) ──
  // Permite explicar por qué el efectivo cobrado (neto) es menor a los billetes que
  // entraron en mano: EF neto + vuelto = recibido. Agrupado por día (dd/MM).
  var vueltos = [];
  var shCbV = SS.getSheetByName('Cambios Billetera');
  if (shCbV && shCbV.getLastRow() > 1) {
    var dCbV = shCbV.getDataRange().getValues();
    for (var rcb = 1; rcb < dCbV.length; rcb++) {
      var fCbRaw = dCbV[rcb][0];
      var mCb = Number(dCbV[rcb][4]) || 0;
      if (mCb <= 0) continue;
      var fCbDt = _parseDateAny(fCbRaw);
      vueltos.push({
        f: fCbDt ? Utilities.formatDate(fCbDt, 'America/Argentina/Buenos_Aires', 'dd/MM') : String(fCbRaw || ''),
        h: String(dCbV[rcb][1] || ''), c: String(dCbV[rcb][3] || ''), $: mCb
      });
    }
  }

  // ── Sobres activos (detalle: a qué proveedor va cada sobre) ──
  var sobresLista = [];
  var shSobL = SS.getSheetByName('Sobres');
  if (shSobL && shSobL.getLastRow() > 1) {
    var dSobL = shSobL.getDataRange().getValues();
    for (var rsl = 1; rsl < dSobL.length; rsl++) {
      if (String(dSobL[rsl][3] || '').trim() !== 'Activo') continue;
      var mSob = Number(dSobL[rsl][2]) || 0;
      if (mSob <= 0) continue;
      sobresLista.push({ r: rsl + 1, prov: String(dSobL[rsl][1] || '').trim(), monto: mSob });
    }
  }

  // ── Gastos por mes (para P&L) — fuente única: hoja Egresos ──
  var gastosHist = {};
  gastos.forEach(function(g) {
    var m = g.mes || '';
    if (m) { if (!gastosHist[m]) gastosHist[m] = 0; gastosHist[m] += g.$; }
  });

  // ── Libro diario unificado (vista derivada) ──
  // Combina: cobros de pedidos (Home/Pilar/Clubes + Liq Red) + ingresos + gastos.
  // El Panel → tab Caja → "Movimientos" lee esto. Sin duplicar datos en el Sheets.
  var movimientos = [];
  function _fechaStrToTs(f) {
    if (!f) return 0;
    var s = String(f).trim();
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
    if (!m) return 0;
    var dt = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]),
                      m[4] ? Number(m[4]) : 0, m[5] ? Number(m[5]) : 0);
    return dt.getTime();
  }
  cobrosDetalle.forEach(function(cd) {
    // 1 entry por método (si mixto, se parten en 2 para no confundir)
    if (cd.efTotal > 0) {
      movimientos.push({
        tipo: 'cobro', f: cd.f, ts: cd.fTs || _fechaStrToTs(cd.f),
        cat: 'COBRO ' + cd.h, con: cd.c + (cd.id ? ' #' + cd.id : ''),
        met: 'Efectivo', $: cd.efTotal, not: ''
      });
    }
    if (cd.mpTotal > 0) {
      movimientos.push({
        tipo: 'cobro', f: cd.f, ts: cd.fTs || _fechaStrToTs(cd.f),
        cat: 'COBRO ' + cd.h, con: cd.c + (cd.id ? ' #' + cd.id : ''),
        met: 'Mercado Pago', $: cd.mpTotal, not: ''
      });
    }
  });
  ingresos.forEach(function(i) {
    movimientos.push({
      tipo: 'ingreso', f: i.fFull || i.f, ts: i.ts || _fechaStrToTs(i.f),
      cat: i.cat || 'Ingreso', con: i.con, met: i.met, $: i.$, not: i.not || ''
    });
  });
  gastos.forEach(function(g) {
    movimientos.push({
      tipo: 'gasto', f: g.fFull || g.f, ts: g.ts || _fechaStrToTs(g.f),
      cat: g.cat || 'Gasto', con: g.con, met: g.met, $: g.$, not: g.not || ''
    });
  });
  // Orden: más reciente arriba.
  movimientos.sort(function(a, b) { return (b.ts || 0) - (a.ts || 0); });

  // En mode='caja' devolvemos solo lo que el frontend necesita para renderizar
  // tab Caja (saldo, movimientos, gastos, ingresos). Sin pedidos[], canales[],
  // stock, OC, vendedores, proveedores. Skipea también la lectura de Proveedores.
  if (_cajaMode) {
    var _outCaja = JSON.stringify({
      ts: Date.now(),
      caja: { cobradoEf: cobradoEf, cobradoMP: cobradoMP, gastosEf: gastosEf, gastosMP: gastosMP, totalGastos: totalGastos, ingresosEf: ingresosEf, ingresosMP: ingresosMP },
      saldoBase: saldoBase,
      gastos: gastos,
      ingresos: ingresos,
      gastosHist: gastosHist,
      movimientos: movimientos,
      sobres: sobresLista,
      cajaMode: true
    });
    try { CacheService.getScriptCache().put('admin_caja_v1', _outCaja, 5); } catch (e) {}
    return ContentService.createTextOutput(_outCaja).setMimeType(ContentService.MimeType.JSON);
  }

  // Lista de proveedores únicos (para override en CONFIRMAR ORIGEN del panel)
  var proveedoresList = [];
  var hProvList = SS.getSheetByName('Proveedores');
  if (hProvList && hProvList.getLastRow() > 1) {
    var pvData = hProvList.getDataRange().getValues();
    var seen = {};
    var lastProvN = '';
    for (var pv = 1; pv < pvData.length; pv++) {
      var nm = String(pvData[pv][2] || '').trim();
      if (nm) lastProvN = nm;
      if (lastProvN && !seen[lastProvN]) { seen[lastProvN] = true; proveedoresList.push(lastProvN); }
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      ts: Date.now(),
      canales: canales,
      totales: totales,
      pedidos: pedidos,
      stock: stock,
      proveedores: proveedoresList,
      oc: { pendientes: ocPend, costo: ocTotal, lista: ocLista },
      caja: { cobradoEf: cobradoEf, cobradoMP: cobradoMP, gastosEf: gastosEf, gastosMP: gastosMP, totalGastos: totalGastos, ingresosEf: ingresosEf, ingresosMP: ingresosMP },
      vendedores: vendedoresArr,
      saldoBase: saldoBase,
      gastos: gastos,
      ingresos: ingresos,
      gastosHist: gastosHist,
      movimientos: movimientos,
      vueltos: vueltos,
      sobres: sobresLista,
      config: _ajConfigMaleuArr(),
      saludHome: _saludHomeSemana(argNow)
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

  // Si la fecha pasada es hoy (o no viene), usar timestamp actual con hora
  // para que quede DESPUÉS del último Saldo Base y se refleje en caja.
  var fecha;
  if (!data.fecha) {
    fecha = argNow;
  } else {
    var hoyISO = argNow.getFullYear() + '-' + String(argNow.getMonth()+1).padStart(2,'0') + '-' + String(argNow.getDate()).padStart(2,'0');
    if (data.fecha === hoyISO) fecha = argNow;
    else fecha = new Date(data.fecha + 'T12:00:00');
  }
  var semana = _getWeekNumber(fecha);
  var meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var mes = meses[fecha.getMonth()];

  var cat = String(data.categoria || '').trim();
  var con = String(data.concepto || '').trim();
  var notas = String(data.notas || '').trim();

  // Soporta:
  //  a) legacy: { metodo, monto }                       → 1 fila
  //  b) batch:  { montoEf, montoMp }                    → 1 o 2 filas en un solo round-trip
  var rows = [];
  var legacyMonto = Number(data.monto) || 0;
  var legacyMet   = String(data.metodo || '').trim();
  var ef = Number(data.montoEf) || 0;
  var mp = Number(data.montoMp) || 0;
  if (legacyMonto > 0 && legacyMet) {
    rows.push([fecha, semana, mes, cat, con, legacyMet, legacyMonto, notas]);
  } else {
    if (ef > 0) rows.push([fecha, semana, mes, cat, con, 'Efectivo',     ef, notas]);
    if (mp > 0) rows.push([fecha, semana, mes, cat, con, 'Mercado Pago', mp, notas]);
  }
  if (rows.length === 0) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'monto vacio' })).setMimeType(ContentService.MimeType.JSON);
  }

  var startRow = shEgr.getLastRow() + 1;
  shEgr.getRange(startRow, 1, rows.length, 8).setValues(rows);
  shEgr.getRange(startRow, 1, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  // Sin flush(): el frontend usa UI optimista, no necesita lectura inmediata consistente.

  return ContentService.createTextOutput(JSON.stringify({ ok: true, rows: rows.length })).setMimeType(ContentService.MimeType.JSON);
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
  // Si la fecha pasada es hoy (o no viene), usar el timestamp actual (con hora real)
  // para que quede DESPUÉS del último Saldo Base y sume en caja.
  var fecha;
  if (!data.fecha) {
    fecha = argNow;
  } else {
    var hoyISO = argNow.getFullYear() + '-' + String(argNow.getMonth()+1).padStart(2,'0') + '-' + String(argNow.getDate()).padStart(2,'0');
    if (data.fecha === hoyISO) fecha = argNow;
    else fecha = new Date(data.fecha + 'T12:00:00');
  }
  var semana = _getWeekNumber(fecha);
  var meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var mes = meses[fecha.getMonth()];
  var cat = String(data.categoria||'').trim();
  var con = String(data.concepto||'').trim();
  var notas = String(data.notas||'').trim();

  // Mismo formato batch que _doPostGasto: legacy {metodo,monto} o {montoEf,montoMp}
  var rows = [];
  var legacyMonto = Number(data.monto) || 0;
  var legacyMet   = String(data.metodo || '').trim();
  var ef = Number(data.montoEf) || 0;
  var mp = Number(data.montoMp) || 0;
  if (legacyMonto > 0 && legacyMet) {
    rows.push([fecha, semana, mes, cat, con, legacyMet, legacyMonto, notas]);
  } else {
    if (ef > 0) rows.push([fecha, semana, mes, cat, con, 'Efectivo',     ef, notas]);
    if (mp > 0) rows.push([fecha, semana, mes, cat, con, 'Mercado Pago', mp, notas]);
  }
  if (rows.length === 0) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'monto vacio' })).setMimeType(ContentService.MimeType.JSON);
  }

  var startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, rows.length, 8).setValues(rows);
  sh.getRange(startRow, 1, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  // Sin flush(): UI optimista en frontend.

  return ContentService.createTextOutput(JSON.stringify({ ok: true, rows: rows.length })).setMimeType(ContentService.MimeType.JSON);
}

/** POST action=ajusteSaldo — guarda saldo base manual (EF y MP).
 *  { action:'ajusteSaldo', efectivo:number, mp:number } */
/** POST action=setOrigenProductos — define origen por producto dentro de un pedido.
 *  { action:'setOrigenProductos', hoja:'Home', id:'H-045', productos:[{a:'PPM',o:'D'},{a:'SCo',o:'OC'}] }
 *  o='D' (Deposito) / o='OC' (Orden de Compra) */
function _doPostSetOrigenProductos(data) {
  var hoja = String(data.hoja || '');
  var pedidoId = String(data.id || '');
  var prods = data.productos || [];

  var sh = SS.getSheetByName(hoja);
  if (!sh) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'hoja' })).setMimeType(ContentService.MimeType.JSON);

  // Buscar fila (último match para manejar N° reseteados semanalmente)
  var allData = sh.getDataRange().getValues();
  var row = Number(data.row) || -1;
  if (row < 2 || row > allData.length || String(allData[row - 1][1]).trim() !== pedidoId) {
    row = -1;
    for (var r = allData.length - 1; r >= 1; r--) {
      if (String(allData[r][1]).trim() === pedidoId) { row = r + 1; break; }
    }
  }
  if (row === -1) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'no encontrado' })).setMimeType(ContentService.MimeType.JSON);

  // Normalizar el formato de cada producto a {a, d, oc} (split por cantidad).
  // Acepta:
  //  Nuevo: {a:'PMu', d:5, oc:3}                    → split explícito
  //  Viejo: {a:'PMu', o:'D'} o {a:'PMu', o:'OC'}    → todo a un solo origen (back-compat)
  // Para el viejo necesitamos saber la cantidad total del producto en el pedido.
  // Determinar columnas de productos según hoja para resolver cantidades del pedido.
  var COL_MAP = (hoja === 'Pilar')  ? PILAR_COL_TO_ABBR
              : (hoja === 'Clubes') ? CLUBES_COL_TO_ABBR
              : (hoja === 'Red')    ? RED_COL_TO_ABBR
                                    : HOME_COL_TO_ABBR;
  var ABBR_TO_COL = {};
  Object.keys(COL_MAP).forEach(function(c){ ABBR_TO_COL[COL_MAP[c]] = Number(c); });

  var rowVals = allData[row - 1];
  var prodsNorm = prods.map(function(p) {
    var ab = p.a;
    var col = ABBR_TO_COL[ab];
    var totalQty = col ? (Number(rowVals[col - 1]) || 0) : 0;
    var prov = p.prov ? String(p.prov).trim() : '';
    if (p.d != null || p.oc != null) {
      var d = Math.max(0, Number(p.d) || 0);
      var oc = Math.max(0, Number(p.oc) || 0);
      return { a: ab, d: d, oc: oc, total: totalQty, prov: prov };
    }
    // formato viejo
    if (p.o === 'D')  return { a: ab, d: totalQty, oc: 0,        total: totalQty, prov: prov };
    if (p.o === 'OC') return { a: ab, d: 0,        oc: totalQty, total: totalQty, prov: prov };
    return { a: ab, d: totalQty, oc: 0, total: totalQty, prov: prov };
  });

  // Determinar resumen: Deposito / Orden de Compra / Mixto
  var anyD = false, anyOC = false;
  prodsNorm.forEach(function(p){
    if (p.d  > 0) anyD  = true;
    if (p.oc > 0) anyOC = true;
  });
  var summary = (anyD && !anyOC) ? 'Deposito'
              : (anyOC && !anyD) ? 'Orden de Compra'
              : (anyD && anyOC)  ? 'Mixto'
              : 'Pendiente';

  // Columna Origen según hoja
  var colOrigen = hoja === 'Clubes' ? 12 : hoja === 'Red' ? 10 : 9;

  // Detalle JSON (col "Origen Detalle"). Formato: { abbr: { d: N, oc: Y } }.
  var headers = allData[0];
  var colDetalle = -1;
  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h]).trim() === 'Origen Detalle') { colDetalle = h + 1; break; }
  }
  var detailObj = {};
  prodsNorm.forEach(function(p){ detailObj[p.a] = { d: p.d, oc: p.oc }; });

  // ── Reconciliar la hoja "Orden de Compra" con el estado deseado ──
  // Crea las OC que faltan, ajusta la cantidad de las que cambiaron y BORRA las que
  // ya no van (ej: un producto que pasó de OC a Depósito). Idempotente y reversible:
  // se puede ir y volver OC↔Depósito las veces que haga falta sin dejar OCs zombis.
  // Si una op destructiva (borrar/bajar) toca una OC comprometida —Recibida, o Pedida
  // ya pasado el cutoff del jueves 12hs— devuelve needsConfirm y NO toca NADA (ni la
  // hoja OC ni el Origen del pedido) hasta que Tadeo confirme mandando force:true.
  var clienteOrig = String(sh.getRange(row, hoja === 'Red' ? 9 : 8).getValue() || '').trim();
  var force = (data.force === true || data.force === 'true');
  var recon = _reconciliarOCPedido(hoja, pedidoId, row, prodsNorm, clienteOrig, force);
  if (recon && recon.needsConfirm) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, needsConfirm: true, warn: recon.warn })).setMimeType(ContentService.MimeType.JSON);
  }

  // Recién ahora que la hoja OC quedó consistente, persistimos Origen + detalle.
  sh.getRange(row, colOrigen).setValue(summary);
  if (colDetalle === -1) {
    colDetalle = headers.length + 1;
    sh.getRange(1, colDetalle).setValue('Origen Detalle');
  }
  sh.getRange(row, colDetalle).setValue(JSON.stringify(detailObj));

  return ContentService.createTextOutput(JSON.stringify({
    ok: true, origen: summary,
    ocCreadas: recon ? recon.created : 0,
    ocActualizadas: recon ? recon.updated : 0,
    ocBorradas: recon ? recon.deleted : 0
  })).setMimeType(ContentService.MimeType.JSON);
}

/** POST action=cambiarOrigen — cambia Origen de un pedido (Deposito / Orden de Compra).
 *  { action:'cambiarOrigen', hoja:'Home', id:'1', origen:'Deposito' } */
function _doPostCambiarOrigen(data) {
  var hoja = String(data.hoja || '');
  var pedidoId = String(data.id || '');
  var origen = String(data.origen || '').trim();

  var validos = ['Pendiente', 'Deposito', 'Orden de Compra'];
  if (validos.indexOf(origen) === -1) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'Origen no valido' })).setMimeType(ContentService.MimeType.JSON);
  }

  var sh = SS.getSheetByName(hoja);
  if (!sh) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'hoja no encontrada' })).setMimeType(ContentService.MimeType.JSON);

  var allData = sh.getDataRange().getValues();
  var row = Number(data.row) || -1;
  if (row < 2 || row > allData.length || String(allData[row - 1][1]).trim() !== pedidoId) {
    row = -1;
    for (var r = allData.length - 1; r >= 1; r--) {
      if (String(allData[r][1]).trim() === pedidoId) { row = r + 1; break; }
    }
  }
  if (row === -1) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'pedido no encontrado' })).setMimeType(ContentService.MimeType.JSON);

  var colOrigen = hoja === 'Clubes' ? 12 : hoja === 'Red' ? 10 : 9;
  sh.getRange(row, colOrigen).setValue(origen);

  // Mantener Origen Detalle (col "Origen Detalle") consistente con el origen plano:
  //   Deposito       → todos los productos del pedido como "D"
  //   Orden de Compra → todos como "OC"
  //   Pendiente      → vacía (no hay decisión todavía)
  // Sin esto, BB queda incompleta cuando se cambia origen sin pasar por el Panel
  // producto-por-producto (ej. pedidos creados desde Ruta + o cambios manuales).
  try {
    var headers = allData[0];
    var colDetalle = -1;
    for (var hh = 0; hh < headers.length; hh++) {
      if (String(headers[hh]).trim() === 'Origen Detalle') { colDetalle = hh + 1; break; }
    }
    if (colDetalle === -1) {
      colDetalle = headers.length + 1;
      sh.getRange(1, colDetalle).setValue('Origen Detalle');
    }
    if (origen === 'Pendiente') {
      sh.getRange(row, colDetalle).clearContent();
    } else {
      var COL_MAP = (hoja === 'Pilar') ? PILAR_COL_TO_ABBR
                  : (hoja === 'Clubes') ? CLUBES_COL_TO_ABBR
                  : (hoja === 'Red')    ? RED_COL_TO_ABBR
                                        : HOME_COL_TO_ABBR;
      var rowVals = allData[row - 1];
      var marca = (origen === 'Deposito') ? 'D' : 'OC';
      var detail = {};
      Object.keys(COL_MAP).forEach(function(colStr) {
        var ci = Number(colStr);
        var abbr = COL_MAP[ci];
        var qty = Number(rowVals[ci - 1]) || 0;
        if (qty > 0) detail[abbr] = marca;
      });
      sh.getRange(row, colDetalle).setValue(JSON.stringify(detail));
    }
  } catch (e) { /* no romper el cambio de origen si falla el detalle */ }

  // Reconciliar la hoja OC con el nuevo origen plano (evita OCs zombis):
  //   Deposito/Pendiente → borra las OC del pedido · Orden de Compra → crea/ajusta.
  // force:true porque este toggle no tiene UI de confirmación; es una vía secundaria.
  try {
    var COL_MAP2 = (hoja === 'Pilar') ? PILAR_COL_TO_ABBR
                 : (hoja === 'Clubes') ? CLUBES_COL_TO_ABBR
                 : (hoja === 'Red')    ? RED_COL_TO_ABBR
                                       : HOME_COL_TO_ABBR;
    var rowVals2 = allData[row - 1];
    var prodsNorm2 = [];
    Object.keys(COL_MAP2).forEach(function(colStr) {
      var ci = Number(colStr);
      var abbr = COL_MAP2[ci];
      var qty = Number(rowVals2[ci - 1]) || 0;
      if (qty > 0) {
        var oc = (origen === 'Orden de Compra') ? qty : 0;
        var d  = (origen === 'Deposito') ? qty : 0;
        prodsNorm2.push({ a: abbr, d: d, oc: oc, prov: '' });
      }
    });
    var clienteOrig2 = String(sh.getRange(row, hoja === 'Red' ? 9 : 8).getValue() || '').trim();
    _reconciliarOCPedido(hoja, pedidoId, row, prodsNorm2, clienteOrig2, true);
  } catch (e2) { /* no romper el cambio de origen si falla la reconciliación de OC */ }

  return ContentService.createTextOutput(JSON.stringify({ ok: true, origen: origen })).setMimeType(ContentService.MimeType.JSON);
}

/** POST action=editarPedido — modifica cantidades / agrega productos / override precio.
 *  Bloquea si el pedido está Entregado, Cancelado o Cobrado.
 *  Si tiene origen confirmado (no Pendiente), también deshace el origen y borra OCs.
 *  Recalcula Total y Costo según las nuevas cantidades y precios.
 *  Body: { hoja, id, row, lineas: [{a, q, precio?}], envio?, descuento? } */
function _doPostEditarPedido(data) {
  var hoja = String(data.hoja || '');
  var pedidoId = String(data.id || '');
  var lineas = data.lineas || [];

  var sh = SS.getSheetByName(hoja);
  if (!sh) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'hoja' })).setMimeType(ContentService.MimeType.JSON);

  var allData = sh.getDataRange().getValues();
  var row = Number(data.row) || -1;
  if (row < 2 || row > allData.length || String(allData[row - 1][1]).trim() !== pedidoId) {
    row = -1;
    for (var r = allData.length - 1; r >= 1; r--) {
      if (String(allData[r][1]).trim() === pedidoId) { row = r + 1; break; }
    }
  }
  if (row === -1) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'pedido no encontrado' })).setMimeType(ContentService.MimeType.JSON);

  // Validar estado: bloqueado si entregado/cancelado/cobrado
  var colEst, colOrigen, colPago, colTotal, colCosto, colMargen, COL_MAP, prodStartCol, prodEndCol, colEnvio, colDescuento, colFacturado;
  if (hoja === 'Clubes') {
    colEst = 14; colOrigen = 12; colPago = 16;
    colTotal = 17; colCosto = 32; colMargen = 33;
    COL_MAP = CLUBES_COL_TO_ABBR;
    prodStartCol = 24; prodEndCol = 31;
    colEnvio = 18; colFacturado = 23; // R envio (0 generalmente), W Facturado
  } else if (hoja === 'Red') {
    colEst = 12; colOrigen = 10; colPago = 14;
    colTotal = 15; colCosto = 45; colMargen = 46;
    COL_MAP = RED_COL_TO_ABBR;
    prodStartCol = 22; prodEndCol = 44;
    colEnvio = 16; colFacturado = 21;
  } else { // Home / Pilar
    var isPilar = (hoja === 'Pilar');
    // Layout v2 (abr/2026): N=14 Subtotal, O=15 Envío, P=16 Descuento, Q=17 Total
    // Estado/Pago: K=11 EstEntrega, L=12 FormaPago, M=13 EstPago
    colEst = 11; colOrigen = 9; colPago = 13;
    colTotal = 17;            // Q = Total a cobrar (antes apuntaba a 14 = Subtotal — bug)
    colSubtotal = 14;         // N = Subtotal Producto
    colCosto = isPilar ? 46 : 42;
    colMargen = isPilar ? 47 : 43;
    COL_MAP = isPilar ? PILAR_COL_TO_ABBR : HOME_COL_TO_ABBR;
    prodStartCol = 23;
    prodEndCol = isPilar ? 45 : 41;
    colEnvio = 15;            // O = Envío
    colFacturado = 22;
    colDescuento = 16;        // P = Descuento (antes apuntaba a 53 = Año Entrega — bug)
    colFormaPago = 12;        // L = Forma de Pago (para recalcular descuento por efectivo)
  }

  var est = String(sh.getRange(row, colEst).getValue()).trim();
  var pago = String(sh.getRange(row, colPago).getValue()).trim();
  if (est === 'Cancelado') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'No se puede editar: el pedido está Cancelado' })).setMimeType(ContentService.MimeType.JSON);
  }
  if (pago === 'Cobrado') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'No se puede editar: el pedido ya está Cobrado' })).setMimeType(ContentService.MimeType.JSON);
  }
  // ── Edición de un pedido YA ENTREGADO ──
  // Permitida solo para Home/Pilar no cobrados. El stock físico ya se descontó al
  // entregar, así que la edición lo reajusta por la diferencia (ver más abajo).
  // Se bloquea si tiene OC vinculada (la reasignación de OC no aplica a un entregado).
  var isEntregado = (est === 'Entregado');
  if (isEntregado && hoja !== 'Home' && hoja !== 'Pilar') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'Solo se pueden editar pedidos entregados de Home o Pilar' })).setMimeType(ContentService.MimeType.JSON);
  }

  var origenActual = String(sh.getRange(row, colOrigen).getValue()).trim();
  var hadOrigen = (origenActual && origenActual !== 'Pendiente');

  // Construir map abbr → {q, precio, costo} desde las líneas recibidas
  var lineasMap = {};
  lineas.forEach(function(l){
    if (!l || !l.a) return;
    var q = Math.max(0, Number(l.q) || 0);
    var precio = l.precio != null ? Number(l.precio) : null;
    var costoOverride = l.costo != null ? Number(l.costo) : null;
    lineasMap[l.a] = { q: q, precio: precio, costo: costoOverride };
  });

  // ── Cambio de forma de pago desde el editor (opcional) ──
  // Si el editor mandó `formaPago`, lo escribimos ANTES del cálculo de descuento
  // y del rebalance ef/tr — así el 10% auto (Efectivo/+100k) y el split R/S usan
  // el pago nuevo. Aplica a todas las hojas (col FP existe en las 4).
  if (data.formaPago) {
    var _colFPWrite = (hoja === 'Clubes') ? 15 : (hoja === 'Red') ? 13 : 12;
    var _fpNueva = String(data.formaPago).trim();
    if (_fpNueva) sh.getRange(row, _colFPWrite).setValue(_fpNueva);
  }

  // Leer Origen Detalle actual y cantidades viejas para decidir si podemos
  // PRESERVAR el origen (cuando solo se bajaron cantidades y/o se eliminaron
  // productos, sin agregar nuevos ni aumentar). Esto evita resetear a Pendiente
  // un pedido ya procesado solo por una corrección menor.
  var headers = allData[0];
  var colOrigenDetalle = 0;
  for (var hh = 0; hh < headers.length; hh++) {
    if (String(headers[hh]).trim() === 'Origen Detalle') { colOrigenDetalle = hh + 1; break; }
  }
  var origenDetalle = {};
  if (colOrigenDetalle) {
    var detStr = String(sh.getRange(row, colOrigenDetalle).getValue() || '').trim();
    try { if (detStr) origenDetalle = JSON.parse(detStr); } catch(e) {}
  }
  var qtyOld = {};
  Object.keys(COL_MAP).forEach(function(colStr){
    var colIdx = Number(colStr);
    var abbr = COL_MAP[colIdx];
    var oldQ = Number(allData[row - 1][colIdx - 1]) || 0;
    if (oldQ > 0) qtyOld[abbr] = oldQ;
  });

  // Localizar OCs vinculadas a este pedido
  var clienteOrig = String(sh.getRange(row, hoja === 'Red' ? 9 : 8).getValue() || '').trim();
  var shOC = SS.getSheetByName('Orden de Compra');
  var ocsDelPedido = [];
  if (shOC && shOC.getLastRow() > 1) {
    var ocAll = shOC.getRange(2, 1, shOC.getLastRow() - 1, shOC.getLastColumn()).getValues();
    for (var io = 0; io < ocAll.length; io++) {
      if (String(ocAll[io][4]).trim() === hoja &&
          String(ocAll[io][5]).trim() === pedidoId &&
          String(ocAll[io][6]).trim() === clienteOrig) {
        ocsDelPedido.push({
          rowOC: io + 2,
          abbr:  String(ocAll[io][11]).trim(),
          qty:   Number(ocAll[io][12]) || 0,
          estado: String(ocAll[io][20]).trim()
        });
      }
    }
  }

  // Entregado con OC vinculada: bloquear (la reasignación de OC no aplica a un entregado).
  if (isEntregado && ocsDelPedido.length > 0) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'No se puede editar un entregado con OC vinculada. Reasigná la OC a mano.' })).setMimeType(ContentService.MimeType.JSON);
  }

  // Reset/preserve de origen + OCs: SOLO para pedidos NO entregados. Un entregado
  // mantiene su origen/estado tal cual; el ajuste de stock se hace por diferencia abajo.
  var origenReseteado = false;
  if (!isEntregado) {
  // ¿Podemos preservar el origen?
  // Sí, si: hadOrigen + sin productos nuevos + sin aumentos + ninguna OC ya pasó
  // del estado Pendiente (si ya se pidió/recibió, hay que reasignar manual).
  var canPreserve = hadOrigen;
  if (canPreserve) {
    Object.keys(lineasMap).forEach(function(abbr){
      if (lineasMap[abbr].q > 0) {
        var oldQ = qtyOld[abbr] || 0;
        if (oldQ === 0) canPreserve = false;          // producto nuevo
        else if (lineasMap[abbr].q > oldQ) canPreserve = false; // aumento
      }
    });
  }
  if (canPreserve) {
    for (var oi = 0; oi < ocsDelPedido.length; oi++) {
      var stOC = ocsDelPedido[oi].estado;
      if (stOC && stOC !== 'Pendiente') { canPreserve = false; break; }
    }
  }

  if (hadOrigen && !canPreserve) {
    // RESET completo: borrar TODAS las OCs y poner Pendiente
    if (shOC) {
      var rowsToDelete = ocsDelPedido.map(function(o){return o.rowOC;}).sort(function(a,b){return b-a;});
      rowsToDelete.forEach(function(rr){ shOC.deleteRow(rr); });
    }
    sh.getRange(row, colOrigen).setValue('Pendiente');
    if (colOrigenDetalle) sh.getRange(row, colOrigenDetalle).setValue('');
    // Si estaba Reservado, bajar a Pendiente. Reservado solo tiene sentido con stock comprometido;
    // al perder el Origen confirmado, queda inconsistente y el SUMPRODUCT de Productos deja de
    // contar este pedido como reservado (riesgo: tienda muestra stock que en realidad está comprometido).
    if (est === 'Reservado') {
      sh.getRange(row, colEst).setValue('Pendiente');
      est = 'Pendiente';
    }
    origenReseteado = true;
  } else if (hadOrigen && canPreserve) {
    // PRESERVAR: ajustar OCs por línea (borrar las que quedaron en 0, actualizar qty
    // de las que bajaron) y recomputar Origen + Origen Detalle.
    if (shOC && ocsDelPedido.length) {
      var rowsToDelOC = [], rowsToUpdOC = [];
      ocsDelPedido.forEach(function(o){
        var newLine = lineasMap[o.abbr];
        var det = origenDetalle[o.abbr] || 'D';
        if (!newLine || newLine.q <= 0) {
          rowsToDelOC.push(o.rowOC);
        } else if (det === 'OC' && newLine.q !== o.qty) {
          rowsToUpdOC.push({ rowOC: o.rowOC, qty: newLine.q });
        }
      });
      rowsToDelOC.sort(function(a,b){return b-a;}).forEach(function(rr){ shOC.deleteRow(rr); });
      rowsToUpdOC.forEach(function(u){
        shOC.getRange(u.rowOC, 13).setValue(u.qty); // M = Cantidad
        // Costo Total (O=15) si no es fórmula
        var celdaCT = shOC.getRange(u.rowOC, 15);
        if (!celdaCT.getFormula()) {
          var cu = Number(shOC.getRange(u.rowOC, 14).getValue()) || 0;
          celdaCT.setValue(u.qty * cu);
        }
      });
    }
    // Limpiar Origen Detalle: solo abbrs que sigan con q > 0
    var newDetalle = {};
    Object.keys(lineasMap).forEach(function(abbr){
      if (lineasMap[abbr].q > 0 && origenDetalle[abbr]) newDetalle[abbr] = origenDetalle[abbr];
    });
    var hasD = false, hasOC = false;
    Object.keys(newDetalle).forEach(function(abbr){
      if (newDetalle[abbr] === 'OC') hasOC = true; else hasD = true;
    });
    var nuevoOrigen = (hasD && hasOC) ? 'Mixto' : (hasOC ? 'Orden de Compra' : 'Deposito');
    sh.getRange(row, colOrigen).setValue(nuevoOrigen);
    if (colOrigenDetalle) sh.getRange(row, colOrigenDetalle).setValue(JSON.stringify(newDetalle));
    // Si quedó 100% OC y el pedido estaba Reservado, bajar Estado: no hay stock comprometido.
    if (nuevoOrigen === 'Orden de Compra' && est === 'Reservado') {
      sh.getRange(row, colEst).setValue('Pendiente');
      est = 'Pendiente';
    }
  }
  } // fin if(!isEntregado) — reset/preserve de origen solo para no entregados

  // Catálogo de Productos: precio retail + costo por abbr
  var hProd = SS.getSheetByName('Productos');
  var precioRetail = {}, costoUnit = {};
  if (hProd) {
    var prodData = hProd.getDataRange().getValues();
    for (var rp = 1; rp < prodData.length; rp++) {
      var ab = String(prodData[rp][2]).trim();
      if (!ab) continue;
      precioRetail[ab] = parseFloat(String(prodData[rp][8]).replace(/[$.]/g,'').replace(/,/g,'')) || 0;
      costoUnit[ab]    = parseFloat(String(prodData[rp][9]).replace(/[$.]/g,'').replace(/,/g,'')) || 0;
    }
  }
  var CLUBES_PRECIOS = {'PMu':7500,'PMa':7500,'PJyQ':7500,'PCC':7500,'PJyM':7800,'PPM':12000,'PPJyQ':12000,'PPCyQ':12000,'ECaC':18400,'EJyQ':16000,'ECyQ':16000,'EV':16000};

  // Reescribir cantidades en TODAS las columnas de productos del bloque
  var nProds = prodEndCol - prodStartCol + 1;
  var newQty = new Array(nProds).fill(0);
  // Cols extra al final, no consecutivas con el bloque principal — se manejan aparte:
  // Tartas: Home BE-BH 57-60, Pilar BH-BK 60-63, Red BF-BI 58-61.
  // Empanadas Clubes: AL-AO 38-41 (28/05/2026).
  var tartaColsByHoja = {
    'Home':   [57, 58, 59, 60],
    'Pilar':  [60, 61, 62, 63],
    'Red':    [58, 59, 60, 61],
    'Clubes': [38, 39, 40, 41]
  };
  var TARTA_ABBRS = ['TP', 'TJyQ', 'TCa', 'TV'];
  var tartaCols = tartaColsByHoja[hoja] || [];
  var newTartaQty = [0, 0, 0, 0];
  // Wraps (RC, RP): en Home/Pilar van al final, fuera del bloque principal (igual que tartas).
  // Home BJ-BK 62-63, Pilar BM-BN 65-66. En Red/Clubes están dentro del bloque principal.
  var wrapColsByHoja = { 'Home': [62, 63], 'Pilar': [65, 66] };
  var wrapCols = wrapColsByHoja[hoja] || [];
  var newWrapQty = [0, 0];
  var subtotal = 0, costo = 0;
  var _recipeLines = [];   // [{abbr, qty, precio}] para recalcular la receta de descuento
  Object.keys(COL_MAP).forEach(function(colStr){
    var colIdx = Number(colStr);
    var abbr = COL_MAP[colIdx];
    var info = lineasMap[abbr];
    var qty = (info && info.q > 0) ? info.q : 0;
    // Determinar si esta col es tarta o wrap (cols separadas al final) y dónde escribirla
    var tartaIdx = tartaCols.indexOf(colIdx);
    var wrapIdx = wrapCols.indexOf(colIdx);
    if (tartaIdx >= 0) {
      newTartaQty[tartaIdx] = qty;
    } else if (wrapIdx >= 0) {
      newWrapQty[wrapIdx] = qty;
    } else if (colIdx >= prodStartCol && colIdx <= prodEndCol) {
      newQty[colIdx - prodStartCol] = qty;
    }
    if (qty > 0) {
      var precio = info.precio != null ? info.precio
                 : (hoja === 'Clubes') ? (CLUBES_PRECIOS[abbr] || 0)
                 : (precioRetail[abbr] || 0);
      var costoLinea = info.costo != null ? info.costo : (costoUnit[abbr] || 0);
      subtotal += qty * precio;
      costo    += qty * costoLinea;
      _recipeLines.push({ abbr: abbr, qty: qty, precio: precio });
    }
  });

  // ── Ajuste de stock físico para pedidos ENTREGADOS (Home/Pilar) ──
  // El físico ya se descontó al entregar. Para reflejar la edición lo reajustamos por
  // la DIFERENCIA: devolvemos las cantidades VIEJAS (+1, lee la fila antes de pisarla) y
  // luego descontamos las NUEVAS (-1, lee la fila ya actualizada). Neto: físico += viejo − nuevo.
  var _origenFis = isEntregado ? String(sh.getRange(row, colOrigen).getValue()).trim() : '';
  var _ajustaFis = isEntregado && hProd && (hoja === 'Home' || hoja === 'Pilar') && (_origenFis === 'Deposito' || _origenFis === 'Mixto');
  if (_ajustaFis) {
    if (_origenFis === 'Deposito') _homeStockFisico(sh, row, hProd, +1);
    else                          _homeStockFisicoMixto(sh, row, hProd, +1);
  }

  // Escribir cantidades en batch
  sh.getRange(row, prodStartCol, 1, nProds).setValues([newQty]);
  // Tartas: escribir las 4 cols (consecutivas entre sí pero en otro rango)
  if (tartaCols.length === 4) {
    sh.getRange(row, tartaCols[0], 1, 4).setValues([newTartaQty]);
  }
  // Wraps: escribir las 2 cols (Home 62-63 / Pilar 65-66, consecutivas)
  if (wrapCols.length === 2) {
    sh.getRange(row, wrapCols[0], 1, 2).setValues([newWrapQty]);
  }

  if (_ajustaFis) {
    SpreadsheetApp.flush(); // asegurar que el -1 lea las cantidades nuevas ya escritas
    if (_origenFis === 'Deposito') _homeStockFisico(sh, row, hProd, -1);
    else                          _homeStockFisicoMixto(sh, row, hProd, -1);
  }

  // ── Recalcular subtotal / descuento / total / facturado / margen ──
  // Reglas de descuento: Home/Pilar 10% si nuevo subtotal >= 100k O pago = Efectivo.
  // Clubes/Red: sin descuento (colDescuento undefined).
  // IMPORTANTE: Total/Facturado/Margen tienen FÓRMULA en el Sheets — se escriben con
  // setFormula para preservar la convención. Si se pisaran con setValue, una edición
  // posterior de envío/descuento/propinas no recalcularía.
  var envio = Number(sh.getRange(row, colEnvio).getValue()) || 0;
  var descuentoNuevo = 0;
  if (colDescuento && (hoja === 'Home' || hoja === 'Pilar')) {
    var formaPago = String(sh.getRange(row, colFormaPago).getValue()).trim();
    // ── Receta de descuento (tab "+") ──
    // Si el pedido tiene una receta guardada (o el editor mandó una nueva), col P se
    // recalcula APLICANDO LA RECETA al nuevo subtotal — así el descuento sobrevive el
    // cambio de productos. Si no hay receta, cae a la regla histórica del 10%.
    var colRecEd = 0;
    for (var hr = 0; hr < headers.length; hr++) {
      if (String(headers[hr]).trim() === 'Descuento Detalle') { colRecEd = hr + 1; break; }
    }
    // El editor del panel puede mandar una receta nueva (Tadeo ajustó el %) con el flag
    // descuentoRecetaEditor=true. Si el editor la mandó (aunque sea null = borró el
    // descuento), respetamos eso y NO revivimos la guardada. Si no la mandó (ej: edición
    // vieja sin descuento), usamos la receta guardada para que el descuento sobreviva.
    var editorSetReceta = (data.descuentoRecetaEditor === true || data.descuentoRecetaEditor === 'true');
    var recetaEd = _parseReceta(data.descuentoReceta);
    if (!recetaEd && !editorSetReceta && colRecEd) recetaEd = _parseReceta(sh.getRange(row, colRecEd).getValue());
    var descReceta = _recetaActiva(recetaEd) ? _calcDescuentoReceta(recetaEd, _recipeLines, formaPago) : null;
    if (descReceta != null) {
      descuentoNuevo = descReceta;
      // Persistir la receta (crea la col si falta) para futuras ediciones
      if (!colRecEd) { colRecEd = sh.getLastColumn() + 1; sh.getRange(1, colRecEd).setValue('Descuento Detalle'); }
      sh.getRange(row, colRecEd).setValue(JSON.stringify(recetaEd));
    } else {
      var aplicaBulk = subtotal >= 100000;
      var aplicaCash = (formaPago === 'Efectivo');
      descuentoNuevo = (aplicaBulk || aplicaCash) ? Math.round(subtotal * 0.10) : 0;
      // Si el editor borró el descuento, limpiar la receta guardada para que no reviva.
      if (editorSetReceta && colRecEd) sh.getRange(row, colRecEd).setValue('');
    }
    sh.getRange(row, colDescuento).setValue(descuentoNuevo);
  }
  var totalNuevo = subtotal + envio - descuentoNuevo;

  // Subtotal Producto (solo Home/Pilar tienen col separada N=14)
  if (typeof colSubtotal !== 'undefined' && colSubtotal) {
    sh.getRange(row, colSubtotal).setValue(subtotal);
  }
  // Total a cobrar: Home/Pilar tienen fórmula =N+O-P; Clubes/Red NO tienen fórmula.
  if (hoja === 'Home' || hoja === 'Pilar') {
    // colSubtotal=N=14, colEnvio=O=15, colDescuento=P=16
    sh.getRange(row, colTotal).setFormula('=N' + row + '+O' + row + '-P' + row);
  } else {
    sh.getRange(row, colTotal).setValue(totalNuevo);
  }
  sh.getRange(row, colCosto).setValue(costo);

  // ── Rebalancear Efectivo (R) / Transferencia (S) según forma de pago ──
  // Pre-cobro, R/S guardan el "monto teórico a cobrar" según fp puro. Si recalculamos
  // Total sin actualizar R/S, queda inconsistente: la caja sigue mostrando el monto
  // viejo cuando el pedido se editó. Solo aplica si fp es Efectivo o Transferencia
  // puros — pedidos con cobro Mixto los maneja la PWA Ruta al cobrar.
  var _colFpRS = (hoja === 'Clubes') ? 15 : (hoja === 'Red') ? 13 : 12;
  var _colEfRS = (hoja === 'Clubes') ? 19 : (hoja === 'Red') ? 17 : 18;
  var _colTrRS = (hoja === 'Clubes') ? 20 : (hoja === 'Red') ? 18 : 19;
  var _fpEdit = String(sh.getRange(row, _colFpRS).getValue()).trim();
  if (_fpEdit === 'Efectivo') {
    sh.getRange(row, _colEfRS).setValue(totalNuevo);
    sh.getRange(row, _colTrRS).setValue(0);
  } else if (_fpEdit === 'Transferencia') {
    sh.getRange(row, _colEfRS).setValue(0);
    sh.getRange(row, _colTrRS).setValue(totalNuevo);
  }

  // Facturado = Total + PropEf + PropTr. Cols por hoja:
  //   Home/Pilar V(22) = =Q+T+U  · Clubes W(23) = =Q+U+V
  //   Red U(21) = =O-P (productos SIN envío: el envío es 100% del vendedor, no
  //   entra en la comisión 17% ni en lo que rinde a Maleu 83%).
  // Margen Bruto = Facturado − Costo:
  //   Home AQ(43) = =V-AP · Pilar AU(47) = =V-AT · Clubes AG(33) = =W-AF · Red AT(46) = =U-AS
  if (colFacturado) {
    var fFacturado, fMargen;
    if (hoja === 'Clubes') {
      fFacturado = '=Q' + row + '+U' + row + '+V' + row;
      fMargen    = '=W' + row + '-AF' + row;
    } else if (hoja === 'Red') {
      fFacturado = '=O' + row + '-P' + row;   // productos sin envío (envío 100% del vendedor)
      fMargen    = '=U' + row + '-AS' + row;
    } else if (hoja === 'Pilar') {
      fFacturado = '=Q' + row + '+T' + row + '+U' + row;
      fMargen    = '=V' + row + '-AT' + row;
    } else { // Home
      fFacturado = '=Q' + row + '+T' + row + '+U' + row;
      fMargen    = '=V' + row + '-AP' + row;
    }
    sh.getRange(row, colFacturado).setFormula(fFacturado);
    sh.getRange(row, colMargen).setFormula(fMargen);
  } else {
    // Fallback defensivo: si por algún layout futuro no hay colFacturado, al menos margen como valor.
    sh.getRange(row, colMargen).setValue(totalNuevo - costo);
  }

  SpreadsheetApp.flush();

  // ── Auto-re-reserva post-edición ──
  // Si el edit reseteó el Origen (porque se agregó algo o subió cantidad), estamos
  // en ventana viernes 15hs–domingo 23hs y el pedido nuevo completo cabe en stock,
  // re-auto-reservamos para no obligar a Tadeo a re-decidir origen en horario pico.
  // El flush() previo aseguró que el SUMPRODUCT de Productos ya recalculó liberando
  // la reserva vieja, así el Stock Disponible que leemos es el real.
  var autoReReservado = false;
  if (origenReseteado && hoja === 'Home') {
    var _nowEd = new Date();
    var _argEd = new Date(_nowEd.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
    if (_inVentanaAutoReservaHome(_argEd)) {
      var _abbrMap = {};
      Object.keys(lineasMap).forEach(function(abbr){
        var q = Number(lineasMap[abbr].q) || 0;
        if (q > 0) _abbrMap[abbr] = q;
      });
      var _hProdRe = SS.getSheetByName('Productos');
      if (_hProdRe && _stockSuficienteParaAbbrs(_hProdRe, _abbrMap)) {
        sh.getRange(row, colOrigen).setValue('Deposito');
        sh.getRange(row, colEst).setValue('Reservado');
        if (colOrigenDetalle) {
          var oDNuevo = {};
          Object.keys(_abbrMap).forEach(function(a){ oDNuevo[a] = 'D'; });
          sh.getRange(row, colOrigenDetalle).setValue(JSON.stringify(oDNuevo));
        }
        origenReseteado = false;
        autoReReservado = true;
        SpreadsheetApp.flush();
      }
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    ok: true, total: totalNuevo, descuento: descuentoNuevo, subtotal: subtotal,
    costo: costo, margen: totalNuevo - costo,
    origenReseteado: origenReseteado,
    autoReReservado: autoReReservado
  })).setMimeType(ContentService.MimeType.JSON);
}

/** POST action=deshacerOrigen — vuelve un pedido a "Pendiente":
 *  - Resetea Origen a "Pendiente"
 *  - Borra "Origen Detalle"
 *  - Elimina las filas de Orden de Compra generadas para este pedido
 *  Bloquea si el pedido ya está Entregado, Cancelado o Cobrado.
 *  { action:'deshacerOrigen', hoja:'Home', id:'H-045', row:N } */
function _doPostDeshacerOrigen(data) {
  var hoja = String(data.hoja || '');
  var pedidoId = String(data.id || '');
  var sh = SS.getSheetByName(hoja);
  if (!sh) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'hoja' })).setMimeType(ContentService.MimeType.JSON);

  var allData = sh.getDataRange().getValues();
  var row = Number(data.row) || -1;
  if (row < 2 || row > allData.length || String(allData[row - 1][1]).trim() !== pedidoId) {
    row = -1;
    for (var r = allData.length - 1; r >= 1; r--) {
      if (String(allData[r][1]).trim() === pedidoId) { row = r + 1; break; }
    }
  }
  if (row === -1) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'no encontrado' })).setMimeType(ContentService.MimeType.JSON);

  // Columnas según hoja
  var colEst, colOrigen, colPago;
  if (hoja === 'Clubes')      { colEst = 14; colOrigen = 12; colPago = 16; }
  else if (hoja === 'Red')    { colEst = 12; colOrigen = 10; colPago = 14; }
  else                        { colEst = 11; colOrigen = 9;  colPago = 13; } // Home/Pilar

  var est  = String(sh.getRange(row, colEst).getValue()).trim();
  var pago = String(sh.getRange(row, colPago).getValue()).trim();
  if (est === 'Entregado' || est === 'Cancelado') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'No se puede deshacer: el pedido está ' + est })).setMimeType(ContentService.MimeType.JSON);
  }
  if (pago === 'Cobrado') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'No se puede deshacer: el pedido ya está Cobrado' })).setMimeType(ContentService.MimeType.JSON);
  }

  // Reset Origen → Pendiente
  sh.getRange(row, colOrigen).setValue('Pendiente');

  // Borrar "Origen Detalle"
  var headers = allData[0];
  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h]).trim() === 'Origen Detalle') {
      sh.getRange(row, h + 1).setValue('');
      break;
    }
  }

  // Borrar filas de OC asociadas (match por Canal + N° Pedido + Cliente)
  var clienteOrig = String(sh.getRange(row, 8).getValue() || '').trim();
  var shOC = SS.getSheetByName('Orden de Compra');
  var deletedOC = 0;
  if (shOC && shOC.getLastRow() > 1) {
    var ocData = shOC.getRange(2, 5, shOC.getLastRow() - 1, 3).getValues(); // E=Canal, F=N°, G=Cliente
    var rowsToDelete = [];
    for (var i = 0; i < ocData.length; i++) {
      if (String(ocData[i][0]).trim() === hoja &&
          String(ocData[i][1]).trim() === pedidoId &&
          String(ocData[i][2]).trim() === clienteOrig) {
        rowsToDelete.push(i + 2);
      }
    }
    rowsToDelete.sort(function(a, b){ return b - a; }); // de abajo hacia arriba
    rowsToDelete.forEach(function(rr){ shOC.deleteRow(rr); deletedOC++; });
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true, ocBorradas: deletedOC })).setMimeType(ContentService.MimeType.JSON);
}

/** POST action=updateDiaEntrega — cambia "Día de Entrega Elegido" del pedido.
 *  { action:'updateDiaEntrega', hoja:'Home', id:'H-001', fecha:'YYYY-MM-DD' }
 *  Escribe "dd/mm/yyyy" en la col correspondiente:
 *    Home/Pilar: J (10) · Clubes: M (13) · Red: K (11) */
function _doPostUpdateDiaEntrega(data) {
  var hoja = String(data.hoja || '');
  var pedidoId = String(data.id || '');
  var fechaISO = String(data.fecha || '').trim();

  var m = fechaISO.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'fecha YYYY-MM-DD requerida' })).setMimeType(ContentService.MimeType.JSON);
  var fechaStr = m[3] + '/' + m[2] + '/' + m[1];

  var sh = SS.getSheetByName(hoja);
  if (!sh) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'hoja' })).setMimeType(ContentService.MimeType.JSON);

  var allData = sh.getDataRange().getValues();
  var row = Number(data.row) || -1;
  if (row < 2 || row > allData.length || String(allData[row - 1][1]).trim() !== pedidoId) {
    row = -1;
    for (var r = allData.length - 1; r >= 1; r--) {
      if (String(allData[r][1]).trim() === pedidoId) { row = r + 1; break; }
    }
  }
  if (row === -1) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'pedido no encontrado' })).setMimeType(ContentService.MimeType.JSON);

  var colDia = (hoja === 'Clubes') ? 13 : (hoja === 'Red') ? 11 : 10;
  sh.getRange(row, colDia).setValue(fechaStr);

  return ContentService.createTextOutput(JSON.stringify({ ok: true, fecha: fechaStr })).setMimeType(ContentService.MimeType.JSON);
}

/** POST action=cambiarEstadoEntrega — cambia Estado de Entrega (Pendiente ↔ Reservado).
 *  { action:'cambiarEstadoEntrega', hoja:'Pilar', id:'P-001', estado:'Reservado' } */
function _doPostCambiarEstadoEntrega(data) {
  var hoja = String(data.hoja || '');
  var pedidoId = String(data.id || '');
  var nuevoEstado = String(data.estado || '').trim();

  var validos = ['Pendiente', 'Reservado'];
  if (validos.indexOf(nuevoEstado) === -1) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'Estado no valido' })).setMimeType(ContentService.MimeType.JSON);
  }

  var sh = SS.getSheetByName(hoja);
  if (!sh) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'hoja no encontrada' })).setMimeType(ContentService.MimeType.JSON);

  var allData = sh.getDataRange().getValues();
  var row = Number(data.row) || -1;
  if (row < 2 || row > allData.length || String(allData[row - 1][1]).trim() !== pedidoId) {
    row = -1;
    for (var r = allData.length - 1; r >= 1; r--) {
      if (String(allData[r][1]).trim() === pedidoId) { row = r + 1; break; }
    }
  }
  if (row === -1) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'pedido no encontrado' })).setMimeType(ContentService.MimeType.JSON);

  var colEstado = hoja === 'Clubes' ? 14 : hoja === 'Red' ? 12 : 11;

  // Pendiente/Reservado no disparan fecha ni stock — safe, pero igualmente ponemos skip-flag
  // para ser consistentes y evitar que el trigger se confunda con lógica futura.
  var propsCE = PropertiesService.getScriptProperties();
  var skipKeyCE = 'skip_onedit_' + hoja + '_' + row;
  propsCE.setProperty(skipKeyCE, String(Date.now()));
  try {
    sh.getRange(row, colEstado).setValue(nuevoEstado);
    SpreadsheetApp.flush();
  } finally {
    propsCE.deleteProperty(skipKeyCE);
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true, estado: nuevoEstado, row: row })).setMimeType(ContentService.MimeType.JSON);
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
    // Home v2: AR(43)=Barrio, AS(44)=Sub Barrio, AT(45)=Lote, AU(46)=Teléfono
    colCliente = 7; colPedido = 1; colTelefono = 46;
    direccion = [rowData[43], rowData[44], 'Lote ' + rowData[45]].filter(Boolean).join(' · ');
  } else if (canal === 'Pilar') {
    // Pilar v2: idx 47=Barrio/Dirección, 48=Domicilio/Lote, 49=Teléfono
    colCliente = 7; colPedido = 1; colTelefono = 49;
    direccion = [rowData[47], 'Lote ' + rowData[48]].filter(Boolean).join(' · ');
  } else if (canal === 'Clubes') {
    colCliente = 7; colPedido = 1; colTelefono = 33;
    direccion = [rowData[8], rowData[9], rowData[10]].filter(Boolean).join(' · ');
  } else if (canal === 'Red') {
    // Red v3 (55 cols): I(9)=Cliente → idx 8, B(2)=N° → idx 1, AZ(52)=Teléfono → idx 51
    // Dirección: AX(50)=Barrio → idx 49, AY(51)=Lote → idx 50
    colCliente = 8; colPedido = 1; colTelefono = 51;
    var _bR = String(rowData[49] || '').trim();
    var _lR = String(rowData[50] || '').trim();
    direccion = [_bR, _lR ? 'Lote ' + _lR : ''].filter(Boolean).join(' · ');
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

  var CLUBES_PRECIOS = {'PMu':7500,'PMa':7500,'PJyQ':7500,'PCC':7500,'PJyM':7800,'PPM':12000,'PPJyQ':12000,'PPCyQ':12000,'ECaC':18400,'EJyQ':16000,'ECyQ':16000,'EV':16000};
  var colToAbbrMap = (canal === 'Clubes')
    ? CLUBES_COL_TO_ABBR
    : (canal === 'Pilar') ? PILAR_COL_TO_ABBR
    : (canal === 'Red')   ? RED_COL_TO_ABBR
    : HOME_COL_TO_ABBR;

  var cliente = String(rowData[colCliente] || '').trim();
  var numPedido = String(rowData[colPedido] || '').trim();
  var telefono = String(rowData[colTelefono] || '').trim();

  Logger.log('OCSelectiva canal=' + canal + ' row=' + row + ' abbrs=' + JSON.stringify(abbrs) + ' cliente="' + cliente + '" nPed="' + numPedido + '" tel="' + telefono + '" dir="' + direccion + '"');
  if (!cliente || !numPedido) {
    SS.toast('OC abortada: fila ' + row + ' sin cliente/N° — revisá el pedido', 'OC error', 8);
    return;
  }

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

  // Soporta tres formatos en `abbrs`:
  //   ['PMu','PJyQ']                                 → usa cantidad TOTAL (legacy)
  //   [{a:'PMu',qty:30},{a:'PJyQ',qty:5}]            → split DEP/OC con qty parcial
  //   [{a:'PMu',qty:30,prov:'Bernardo Pisano'}, ...] → además override de proveedor
  var abbrInfo = {}; // abbr → {qty, prov}
  abbrs.forEach(function(x) {
    if (typeof x === 'string') abbrInfo[x] = { qty: null, prov: '' };
    else if (x && x.a) abbrInfo[x.a] = {
      qty: Number(x.qty) > 0 ? Number(x.qty) : null,
      prov: x.prov ? String(x.prov).trim() : ''
    };
  });

  var newRows = [];
  Object.keys(colToAbbrMap).forEach(function(colStr) {
    var colIdx = Number(colStr);
    var abbr = colToAbbrMap[colIdx];
    if (!(abbr in abbrInfo)) return;
    var info = abbrInfo[abbr];
    var totalQty = Number(rowData[colIdx - 1]) || 0;
    var qty = info.qty != null ? info.qty : totalQty;
    if (qty <= 0) return;

    var ocId = _nextId('OC-');
    var costoUnit = abbrToCostoMap[abbr] || 0;
    var costoTotal = costoUnit * qty;
    var precioVenta = (canal === 'Clubes') ? (CLUBES_PRECIOS[abbr] || 0) : (abbrToPrecioMap[abbr] || 0);
    // Override de proveedor si vino del panel; sino el del catálogo Proveedores
    var proveedorFinal = info.prov || abbrToProvMap[abbr] || '';

    newRows.push([
      ocId, fechaStr, semana, mesNombre, canal, numPedido, cliente, telefono, direccion,
      proveedorFinal, abbrToNameMap[abbr] || abbr, abbr, qty,
      costoUnit, costoTotal, precioVenta, 0, 0, 0,
      // Nace "Pendiente" (Momento 1: marcada, Tadeo todavía decidiendo, NO pasada a
      // proveedor). Pasa a "Pedido" cuando Tadeo la pasa al proveedor desde Abastecimiento
      // (botón "Marcar como Pedido"). Por eso V (Fecha Pedido Prov) queda vacía al nacer.
      'Orden de Compra', 'Pendiente', '', '', 'No', 'No'
    ]);
  });

  if (newRows.length > 0) {
    // Batch insert: 1 round-trip en lugar de N appendRow.
    var formulaStart = shOC.getLastRow() + 1;
    var nCols = newRows[0].length;
    shOC.getRange(formulaStart, 1, newRows.length, nCols).setValues(newRows);
    // Fórmulas Q/R/S por fila (P*M, Q-O, R/Q) en una sola pasada con setFormulas.
    var formulasQRS = [];
    for (var i = 0; i < newRows.length; i++) {
      var rr = formulaStart + i;
      formulasQRS.push(['=P'+rr+'*M'+rr, '=Q'+rr+'-O'+rr, '=R'+rr+'/Q'+rr]);
    }
    shOC.getRange(formulaStart, 17, newRows.length, 3).setFormulas(formulasQRS);
    shOC.getRange(formulaStart, 13, newRows.length, 1).setNumberFormat('0');
    shOC.getRange(formulaStart, 14, newRows.length, 2).setNumberFormat('$#,##0');
    shOC.getRange(formulaStart, 16, newRows.length, 3).setNumberFormat('$#,##0');
    shOC.getRange(formulaStart, 19, newRows.length, 1).setNumberFormat('0.0%');
  }
}

/** ¿Estamos pasado el cutoff semanal de pedidos a proveedores (Jueves 12:00 AR)?
 *  Maleu es comercializadora: hasta el Jue 12hs recibe pedidos y decide origen libre;
 *  a partir de ahí Tadeo consolida y le pasa el pedido a cada proveedor. Post-cutoff,
 *  retroceder una OC "capaz ya se la pasó al proveedor" → conviene avisar.
 *  Ventana post-cutoff = Jue≥12, Vie, Sáb. Dom–Mié y Jue<12 = pre-cutoff (libertad total).
 *  Es solo la señal del AVISO (confirmá y seguí), no un bloqueo. */
var CUTOFF_PROVEEDOR_HORA = 12; // Jueves 12:00 AR
function _esPostCutoffProveedor() {
  var arg = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var dow = arg.getDay(); // 0 Dom … 4 Jue … 6 Sáb
  if (dow === 4) return arg.getHours() >= CUTOFF_PROVEEDOR_HORA; // Jueves
  if (dow === 5 || dow === 6) return true;                        // Vie, Sáb
  return false;
}

/** Reconcilia la hoja "Orden de Compra" para que refleje EXACTAMENTE el estado deseado
 *  de un pedido (prodsNorm = [{a, d, oc, prov}]). Es el corazón del ida-y-vuelta reversible:
 *    - producto con oc>0 y sin OC existente  → CREA la OC (vía _generarOCSelectiva)
 *    - producto con oc>0 y OC de otra cantidad → AJUSTA la cantidad (M) + costo total (O)
 *    - producto con oc=0 y OC existente        → BORRA la fila de OC
 *  Guarda destructiva: si una op de borrar/bajar toca una OC comprometida (estado
 *  "Recibido", o "Pedido" ya pasado el cutoff del jueves) y force!==true → devuelve
 *  { needsConfirm:true, warn:[...] } SIN tocar nada. El caller re-manda con force:true
 *  después de que Tadeo confirme.
 *  Devuelve { ok:true, created, updated, deleted } cuando aplica. */
function _reconciliarOCPedido(hoja, pedidoId, row, prodsNorm, clienteOrig, force) {
  var shOC = SS.getSheetByName('Orden de Compra');
  if (!shOC) return { ok: false, err: 'sin hoja OC' };

  // Estado deseado de OC por abbr
  var desired = {}; // abbr → { oc, prov }
  prodsNorm.forEach(function(p) {
    desired[p.a] = { oc: Math.max(0, Number(p.oc) || 0), prov: p.prov || '' };
  });

  // OC existentes de este pedido (match Canal E + N° F + Cliente G)
  var existing = []; // { rowOC, abbr, qty, estado }
  if (shOC.getLastRow() > 1) {
    var vals = shOC.getRange(2, 1, shOC.getLastRow() - 1, 21).getValues(); // A..U
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][4]).trim() === hoja &&           // E Canal
          String(vals[i][5]).trim() === pedidoId &&        // F N° Pedido
          String(vals[i][6]).trim() === clienteOrig) {     // G Cliente
        existing.push({
          rowOC:  i + 2,
          abbr:   String(vals[i][11]).trim(),              // L Abreviatura
          qty:    Number(vals[i][12]) || 0,                // M Cantidad
          estado: String(vals[i][20]).trim()               // U Estado OC
        });
      }
    }
  }

  // Una OC es "sensible" (retroceder/bajar amerita aviso+confirmación) cuando ya está
  // comprometida con el proveedor: estado "Pedido" (ya se la pasé a ESE proveedor) o
  // "Recibido" (la mercadería ya entró). "Pendiente" = Momento 1, todavía decidiendo →
  // libre y silencioso. El discriminador es el estado REAL por proveedor, no el reloj:
  // cada proveedor se cierra en su momento (Claudia el lunes, otros el jueves 12hs).
  function esSensible(estado) {
    return estado === 'Pedido' || estado === 'Recibido';
  }

  var toUpdate = [], toDelete = [], sensitive = [], seenAbbr = {};
  existing.forEach(function(o) {
    seenAbbr[o.abbr] = true;
    var wantOC = desired[o.abbr] ? desired[o.abbr].oc : 0;
    if (wantOC <= 0) {
      toDelete.push(o);
      if (esSensible(o.estado)) sensitive.push({ abbr: o.abbr, estado: o.estado, accion: 'quitar' });
    } else if (wantOC !== o.qty) {
      toUpdate.push({ rowOC: o.rowOC, qty: wantOC });
      if (wantOC < o.qty && esSensible(o.estado)) sensitive.push({ abbr: o.abbr, estado: o.estado, accion: 'bajar' });
    }
  });

  var toCreate = [];
  Object.keys(desired).forEach(function(ab) {
    if (desired[ab].oc > 0 && !seenAbbr[ab]) toCreate.push({ a: ab, qty: desired[ab].oc, prov: desired[ab].prov });
  });

  // Guarda: si hay ops destructivas sobre OC comprometidas y no vino force → pedir confirmación.
  if (sensitive.length > 0 && !force) {
    return { needsConfirm: true, warn: sensitive };
  }

  // APLICAR. Orden importante: primero updates (no mueven filas), después deletes
  // (de abajo hacia arriba, así los índices no se corren), por último creates (append).
  toUpdate.forEach(function(u) {
    shOC.getRange(u.rowOC, 13).setValue(u.qty); // M Cantidad
    var celdaCT = shOC.getRange(u.rowOC, 15);   // O Costo Total (si no es fórmula)
    if (!celdaCT.getFormula()) {
      var cu = Number(shOC.getRange(u.rowOC, 14).getValue()) || 0; // N Costo Unitario
      celdaCT.setValue(u.qty * cu);
    }
  });
  toDelete.map(function(o){ return o.rowOC; })
          .sort(function(a, b){ return b - a; })
          .forEach(function(rr){ shOC.deleteRow(rr); });
  if (toCreate.length > 0) {
    _generarOCSelectiva(hoja, row, toCreate);
  }

  return { ok: true, created: toCreate.length, updated: toUpdate.length, deleted: toDelete.length };
}

/** POST action=marcarCobrado — marca un pedido como Cobrado desde el Panel.
 *  { action:'marcarCobrado', hoja:'Clubes', id:'C-005' } */
// Cobrar todos los pedidos de un Vendedor Red
/** POST action=cobrarVendedorRed — Tadeo cobra a un vendedor Red desde Ruta→COBROS.
 *  Antes tocaba la col equivocada (14, Estado de Pago al cliente). Ahora hace
 *  exactamente lo mismo que cuando el vendedor paga desde su portal (red.html):
 *  - Marca col BB (54) "Estado Pago a Maleu" = "Pagado".
 *  - Marca col BC (55) "Fecha Pago a Maleu" = ahora.
 *  - Crea filas en hoja "Pagos Red Liq" agrupadas por semana ISO.
 *  Criterios: vendedor=X, no cancelado, Estado de Pago = Cobrado (al cliente),
 *  Estado Pago a Maleu != Pagado.
 *  Body: { vendedor, ef?, tr? }
 *    - ef/tr opcional: monto total cobrado por método. Si no llegan, default
 *      100% transferencia (asumimos que Marcos siempre paga MP).
 *      Si llegan, se prorratean entre las semanas según el monto a pagar.
 *  Reportado por Tadeo (25/05/2026): cobrar desde Ruta no sincronizaba con
 *  el portal del vendedor (col equivocada). Fix replica el flujo de
 *  marcarSemanaPagadaRed (que es lo que ejecuta el portal del vendedor). */
function _doPostCobrarVendedorRed(data) {
  var vendedor = String(data.vendedor || '').trim();
  // Tolerar el sufijo "(Vendedor Red)" que el frontend de Cobros suele agregar
  // al nombre para mostrar. La col Vendedor del Sheets tiene solo "Nombre Apellido".
  // Sin este strip, el match strict equality fallaba silencioso (caso 30/06/2026).
  vendedor = vendedor.replace(/\s*\(Vendedor Red\)\s*$/i, '').trim();
  if (!vendedor) return ContentService.createTextOutput(JSON.stringify({ok:false,err:'vendedor vacio'})).setMimeType(ContentService.MimeType.JSON);
  var sh = SS.getSheetByName('Red');
  if (!sh || sh.getLastRow() <= 1) return ContentService.createTextOutput(JSON.stringify({ok:false,err:'hoja vacia'})).setMimeType(ContentService.MimeType.JSON);

  var argDate = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var pad = function(n){return String(n).padStart(2,'0');};
  var fechaStr = pad(argDate.getDate()) + '/' + pad(argDate.getMonth()+1) + '/' + argDate.getFullYear();

  // Lectura completa + índices por header (Red puede crecer en columnas).
  var dAll = sh.getDataRange().getValues();
  var hdr = dAll[0];
  var iEntVend = -1, iPagoMaleu = -1, iFpMaleu = -1, iFechaPagoMaleu = -1, iAPagar = -1;
  for (var hh = 0; hh < hdr.length; hh++) {
    var nm = String(hdr[hh]).trim();
    if (nm === 'Entregado a Vendedor') iEntVend = hh;
    else if (nm === 'Estado Pago a Maleu') iPagoMaleu = hh;
    else if (nm === 'Forma Pago a Maleu') iFpMaleu = hh;
    else if (nm === 'Fecha Pago a Maleu') iFechaPagoMaleu = hh;
    else if (nm === 'A Pagar') iAPagar = hh;
  }
  if (iAPagar < 0) iAPagar = 48; // fallback col AW
  var porSemana = {};
  var totalGlobal = 0, updatedTotal = 0;

  // Filtro ALINEADO con cobros pendientes (cobrosPendientes Red): cuenta pedidos
  // entregados al vendedor y NO pagados a Maleu, SIN importar si el cliente le pagó
  // al vendedor (Estado Pago al cliente). Esto permite que Tadeo cobre al vendedor
  // desde Ruta cuando el vendedor se olvidó de registrar (caso Marcos 23/06).
  for (var i = 1; i < dAll.length; i++) {
    var d = dAll[i];
    var vend = String(d[7] || '').trim();
    if (vend !== vendedor) continue;
    var estado = String(d[11] || '').trim();
    if (estado === 'Cancelado') continue;
    var entVend = iEntVend >= 0 ? String(d[iEntVend] || '').trim() : 'Entregado';
    if (entVend !== 'Entregado') continue;
    var pagoMaleu = iPagoMaleu >= 0 ? String(d[iPagoMaleu] || '').trim() : '';
    if (pagoMaleu === 'Pagado' || pagoMaleu === 'Sí' || pagoMaleu === 'Si') continue;
    var yr = Number(d[6]) || argDate.getFullYear();
    var sem = Number(d[5]) || 0;
    var aPagar = Number(d[iAPagar]) || 0;
    var pedId = String(d[1] || '');
    var skey = yr + '|' + sem;
    if (!porSemana[skey]) porSemana[skey] = { year: yr, sem: sem, filas: [], aPagar: 0, ids: [] };
    porSemana[skey].filas.push(i + 1);
    porSemana[skey].aPagar += aPagar;
    porSemana[skey].ids.push(pedId);
    totalGlobal += aPagar;
    updatedTotal++;
  }

  if (updatedTotal === 0) {
    // Distinguir "ya estaba todo pagado" (tranquilizador, card vieja) de "no coincide
    // el nombre / no existe" (problema real). Así el frontend no asusta cuando el
    // vendedor simplemente ya fue cobrado (caso Marcos 14/07: card cacheada).
    var totalDelVend = 0, yaPagados = 0, ultimaFechaPago = '';
    for (var iz = 1; iz < dAll.length; iz++) {
      if (String(dAll[iz][7] || '').trim() !== vendedor) continue;
      if (String(dAll[iz][11] || '').trim() === 'Cancelado') continue;
      totalDelVend++;
      var pmz = iPagoMaleu >= 0 ? String(dAll[iz][iPagoMaleu] || '').trim() : '';
      if (pmz === 'Pagado' || pmz === 'Sí' || pmz === 'Si') {
        yaPagados++;
        var fpg = iFechaPagoMaleu >= 0 ? dAll[iz][iFechaPagoMaleu] : '';
        if (fpg) ultimaFechaPago = (fpg instanceof Date) ? Utilities.formatDate(fpg, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : String(fpg).split(' ')[0];
      }
    }
    var reason = (totalDelVend === 0) ? 'sin_match' : (yaPagados > 0 ? 'ya_pagado' : 'sin_pendientes');
    return ContentService.createTextOutput(JSON.stringify({
      ok: true, updated: 0, reason: reason,
      yaPagados: yaPagados, totalDelVend: totalDelVend, ultimaFechaPago: ultimaFechaPago
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // Crear filas en hoja Pagos Red Liq (una por semana ISO).
  // Calcular ef/tr por semana ANTES para usar al setear la forma de pago en BA.
  var shLiq = SS.getSheetByName('Pagos Red Liq');
  if (!shLiq) {
    shLiq = SS.insertSheet('Pagos Red Liq');
    shLiq.appendRow(['Timestamp','Fecha','Vendedor','Año','Semana','Efectivo','Transferencia','Total','A Pagar Esperado','Diferencia','Pedidos','# Pedidos']);
    shLiq.setFrozenRows(1);
    shLiq.getRange('A1:L1').setFontWeight('bold').setBackground('#f3e5ab');
  }
  // Forma de pago elegida por Tadeo al cobrar al vendedor desde Ruta ('Efectivo' o
  // 'Transferencia'/MP). Fallback compat: si vienen montos ef/tr explícitos, ratio.
  var fpElegida = String(data.formaPago || '').trim();
  var totalEfIn = Number(data.ef) || 0;
  var totalTrIn = Number(data.tr) || 0;
  var totalIn = totalEfIn + totalTrIn;
  Object.keys(porSemana).forEach(function(skey) {
    var s = porSemana[skey];
    var ef, tr;
    if (fpElegida === 'Efectivo') { ef = s.aPagar; tr = 0; }
    else if (fpElegida === 'Transferencia' || fpElegida === 'Mercado Pago') { ef = 0; tr = s.aPagar; }
    else if (totalIn > 0 && totalGlobal > 0) {
      var ratio = s.aPagar / totalGlobal;
      ef = Math.round(totalEfIn * ratio);
      tr = Math.round(totalTrIn * ratio);
    } else { ef = 0; tr = s.aPagar; } // default 100% transferencia (MP)
    var formaPagoCol = (ef > 0 && tr > 0) ? 'Mixto' : (ef > 0 ? 'Efectivo' : 'Transferencia');
    s.filas.forEach(function(rowSheet) {
      if (iFpMaleu >= 0) sh.getRange(rowSheet, iFpMaleu + 1).setValue(formaPagoCol);
      if (iPagoMaleu >= 0) sh.getRange(rowSheet, iPagoMaleu + 1).setValue('Pagado');
      if (iFechaPagoMaleu >= 0) sh.getRange(rowSheet, iFechaPagoMaleu + 1).setValue(argDate);
    });
    shLiq.appendRow([
      argDate, fechaStr, vendedor, s.year, s.sem,
      ef, tr, ef + tr, s.aPagar, (ef + tr) - s.aPagar,
      s.ids.join(', '), s.ids.length
    ]);
  });

  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    updated: updatedTotal,
    total: totalGlobal,
    semanas: Object.keys(porSemana).length
  })).setMimeType(ContentService.MimeType.JSON);
}

// Cancelar pedido: Estado Entrega = Cancelado + revertir stock si corresponde
function _doPostCancelarPedido(data) {
  var hoja = String(data.hoja || '');
  var pedidoId = String(data.id || '');

  var sh = SS.getSheetByName(hoja);
  if (!sh) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'hoja' })).setMimeType(ContentService.MimeType.JSON);

  var allData = sh.getDataRange().getValues();
  var row = Number(data.row) || -1;
  if (row < 2 || row > allData.length || String(allData[row - 1][1]).trim() !== pedidoId) {
    row = -1;
    for (var r = allData.length - 1; r >= 1; r--) {
      if (String(allData[r][1]).trim() === pedidoId) { row = r + 1; break; }
    }
  }
  if (row === -1) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'no encontrado' })).setMimeType(ContentService.MimeType.JSON);

  // Columnas segun hoja
  var colEst, colOrigen;
  if (hoja === 'Clubes') { colEst = 14; colOrigen = 12; }
  else if (hoja === 'Red') { colEst = 12; colOrigen = 10; }
  else { colEst = 11; colOrigen = 9; } // Home/Pilar/CF

  var estadoAnterior = String(sh.getRange(row, colEst).getValue()).trim();
  var origen = String(sh.getRange(row, colOrigen).getValue()).trim();

  // Si ya estaba cancelado, no hacer nada
  if (estadoAnterior === 'Cancelado') {
    return ContentService.createTextOutput(JSON.stringify({ ok: true, ya: true })).setMimeType(ContentService.MimeType.JSON);
  }

  // Si estaba Entregado + Deposito/Mixto → revertir stock fisico (+1) en TODAS las hojas
  if (estadoAnterior === 'Entregado') {
    var hProd = SS.getSheetByName('Productos');
    if (hProd) {
      if (hoja === 'Clubes') {
        if (origen === 'Deposito') _clubesStockFisico(sh, row, hProd, +1);
        else if (origen === 'Mixto') _clubesStockFisicoMixto(sh, row, hProd, +1);
      } else if (hoja === 'Red') {
        if (origen === 'Deposito') _redStockFisico(sh, row, hProd, +1);
        else if (origen === 'Mixto') _redStockFisicoMixto(sh, row, hProd, +1);
      } else { // Home/Pilar
        if (origen === 'Deposito') _homeStockFisico(sh, row, hProd, +1);
        else if (origen === 'Mixto') _homeStockFisicoMixto(sh, row, hProd, +1);
      }
    }
  }

  // Marcar como Cancelado y poner N° = "-" (no cuenta en numeración semanal)
  sh.getRange(row, colEst).setValue('Cancelado');
  sh.getRange(row, 2).setValue('-');

  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  Log de movimiento de EFECTIVO EN MANO del repartidor (para el arqueo de fin de día).
//  Hoja "Cambios Billetera": Fecha | Hoja | Pedido | Cliente | Monto | Repartidor | Tipo
//  tipo:
//   - 'Billetera'  → vuelto en efectivo que salió de la billetera (el billete entró, el
//                    cambio salió de la billetera). El arqueo lo SUMA para reconstruir la
//                    billetera inicial del día (la col D ya fue bajada por este vuelto).
//   - 'CruzadoEf'  → el cliente pagó MP y le diste vuelto en EFECTIVO. Salió efectivo de
//                    tu mano sin que entre por el cobro. El arqueo lo RESTA.
//   - 'CambioMP'   → el cliente pagó efectivo de más y le devolviste por MP. Te quedó
//                    efectivo extra en mano. El arqueo lo SUMA.
//  Monto siempre positivo; el signo lo decide el Tipo al armar el arqueo.
// ════════════════════════════════════════════════════════════
function _appendCambioBilletera(hoja, pedidoId, cliente, monto, repartidor, tipo) {
  if (!(Number(monto) > 0)) return;
  var sh = SS.getSheetByName('Cambios Billetera');
  if (!sh) {
    sh = SS.insertSheet('Cambios Billetera');
    sh.getRange(1, 1, 1, 7).setValues([['Fecha', 'Hoja', 'Pedido', 'Cliente', 'Monto', 'Repartidor', 'Tipo']]);
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < 7) {
    // Extender el header viejo (5 cols) a 7 sin tocar las filas existentes.
    sh.getRange(1, 6, 1, 2).setValues([['Repartidor', 'Tipo']]);
  }
  var argNow = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  sh.appendRow([argNow, hoja, String(pedidoId), String(cliente || ''), Number(monto), String(repartidor || ''), tipo]);
  var lr = sh.getLastRow();
  sh.getRange(lr, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  sh.getRange(lr, 5).setNumberFormat('$#,##0');
}

function _doPostMarcarCobrado(data) {
  var hoja = String(data.hoja || '');
  var pedidoId = String(data.id || '');

  var sh = SS.getSheetByName(hoja);
  if (!sh) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'hoja' })).setMimeType(ContentService.MimeType.JSON);

  // Validar row hint con un solo getValue (rapido). Solo escanea la hoja completa si el hint falla.
  var row = Number(data.row) || 0;
  if (row > 1 && String(sh.getRange(row, 2).getValue()).trim() === pedidoId) {
    // row hint correcto — no hace falta leer toda la hoja
  } else {
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'hoja vacia' })).setMimeType(ContentService.MimeType.JSON);
    var ids = sh.getRange(2, 2, lastRow - 1, 1).getValues();
    row = -1;
    for (var i = ids.length - 1; i >= 0; i--) {
      if (String(ids[i][0]).trim() === pedidoId) { row = i + 2; break; }
    }
    if (row === -1) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'no encontrado' })).setMimeType(ContentService.MimeType.JSON);
  }

  // Columnas según hoja (1-based)
  // Home/Pilar v2 además: N(14)=Subtotal, O(15)=Envío, P(16)=Descuento, Q(17)=Total
  var cols;
  var isVD = (hoja === 'Home' || hoja === 'Pilar');
  if (hoja === 'Clubes') cols = {fp:15, estPago:16, ef:19, tr:20, propEf:21, propTr:22};
  else if (hoja === 'Red') cols = {fp:13, estPago:14, ef:17, tr:18, propEf:19, propTr:20};
  else cols = {fp:12, estPago:13, sub:14, env:15, desc:16, total:17, ef:18, tr:19, propEf:20, propTr:21}; // Home/Pilar v2

  // ── IDEMPOTENCIA (crítico) ──
  // Apps Script puede tardar >15s; el cliente aborta pero el server igual escribe.
  // Si el cliente reintenta, sin esta guarda se DUPLICARÍA todo: cobro, a-favor (col
  // BE/BH + hoja Saldos Clientes) y descuento de billetera. Si el pedido ya está
  // Cobrado, no re-escribimos nada y devolvemos los valores ya guardados.
  if (String(sh.getRange(row, cols.estPago).getValue()).trim() === 'Cobrado') {
    return ContentService.createTextOutput(JSON.stringify({
      ok: true, already: true,
      ef: Number(sh.getRange(row, cols.ef).getValue()) || 0,
      tr: Number(sh.getRange(row, cols.tr).getValue()) || 0
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // Marcar Cobrado + estampar fecha de cobro
  // IMPORTANTE: el flujo de Cobros del Panel/Ruta SOLO toca Estado de Pago + ef/tr.
  // NO debe tocar Estado de Entrega — eso es una acción aparte (Ruta → entregar).
  // Tadeo lo aclaró 27/05/2026.
  sh.getRange(row, cols.estPago).setValue('Cobrado');
  _stampFechaCobro(sh, row);

  // Cambiar Forma de Pago si viene
  var formaPago = String(data.formaPago || '').trim();
  if (formaPago) sh.getRange(row, cols.fp).setValue(formaPago);

  // ── Recalcular total si cambia la forma de pago (solo Home/Pilar) ──
  // Intención de Tadeo: el 10% aplicado al pedido representa "10% OFF Efectivo".
  // Al cambiar manualmente la fp post-entrega el descuento refleja el método real:
  //   fp=Efectivo → aplica 10%; fp=Transferencia → quita 10%.
  // (La regla de bulk >$100K del checkout no se replica acá: cuando Tadeo toca
  // el fp del pedido, lo hace por el método elegido, no por el bulk.)
  // Se recalcula solo si data.recalcular=true para no alterar pedidos sin cambio.
  var totalRecalculado = null;
  if (isVD && data.recalcular && formaPago && formaPago !== 'Mixto') {
    var subtotal = Number(sh.getRange(row, cols.sub).getValue()) || 0;
    var envio    = Number(sh.getRange(row, cols.env).getValue()) || 0;
    if (subtotal > 0) {
      // data.descuentoManual: Tadeo aplicó un descuento libre (% o $) al cobrar.
      // Gana sobre cualquier otra regla. Cobra subtotal+envío-monto. Si monto<=0,
      // se ignora.
      // data.noDescuento: Tadeo quitó manualmente el 10% en el cuadro de Cobros
      // (cobra precio normal aunque sea Efectivo o supere $100K).
      var nuevoDesc;
      if (data.descuentoManual && Number(data.descuentoManual.monto) > 0) {
        nuevoDesc = Math.min(Number(data.descuentoManual.monto), subtotal + envio);
      } else if (formaPago === 'Efectivo' && !data.noDescuento) {
        nuevoDesc = Math.round(subtotal * 0.10);
      } else {
        nuevoDesc = 0;
      }
      var nuevoTotal = subtotal + envio - nuevoDesc;
      sh.getRange(row, cols.desc).setValue(nuevoDesc);
      sh.getRange(row, cols.total).setValue(nuevoTotal);
      totalRecalculado = nuevoTotal;
    }
  }
  // Mixto + descuento manual: Mixto no entra al bloque de arriba pero igual hay que
  // bajar el Descuento y Total a cobrar de la hoja para que el Facturado quede correcto.
  if (isVD && formaPago === 'Mixto' && data.descuentoManual && Number(data.descuentoManual.monto) > 0) {
    var subM = Number(sh.getRange(row, cols.sub).getValue()) || 0;
    var envM = Number(sh.getRange(row, cols.env).getValue()) || 0;
    if (subM > 0) {
      var descM = Math.min(Number(data.descuentoManual.monto), subM + envM);
      sh.getRange(row, cols.desc).setValue(descM);
      sh.getRange(row, cols.total).setValue(subM + envM - descM);
    }
  }

  // Modelo: R (Efectivo) y S (Transferencia) = parte del TOTAL A COBRAR por método (sin propina).
  // T/U = propina por método (plus sobre el total). La col Facturado (V) = Q+T+U+BE/BH.
  // BE (Home col 57) / BH (Pilar col 60) = A Favor / Aplicado (positivo: cliente paga
  // más; negativo: se aplica saldo de crédito previo del cliente).
  // El cliente (PWA Ruta) envía ef/tr ya NETOS al total facturado + propina aparte.
  var propina = Number(data.propina) || 0;
  var propMet = String(data.propMet || '').trim();
  var ef = Number(data.ef) || 0;
  var tr = Number(data.tr) || 0;
  // aFavor: cliente paga más que el total. Suma a ef/tr según fp. Va a col BE/BH.
  // aplicacion: se descuenta del cobro porque el cliente tenía saldo a favor. Resta de ef/tr.
  // Ambos son siempre positivos en el payload; el signo neto se calcula acá.
  var aFavor = Number(data.aFavor) || 0;
  var aplicacion = Number(data.aplicacion) || 0;
  if (aFavor < 0) aFavor = 0;
  if (aplicacion < 0) aplicacion = 0;
  var deltaAFavor = aFavor - aplicacion;  // positivo = a favor; negativo = aplicación neta
  // Si hubo recálculo y no es Mixto, forzar ef/tr al nuevo total según fp
  if (totalRecalculado !== null && formaPago === 'Efectivo')      { ef = totalRecalculado; tr = 0; }
  else if (totalRecalculado !== null && formaPago === 'Transferencia') { tr = totalRecalculado; ef = 0; }
  // Ajustar ef/tr con el aFavor/aplicacion (lo que el cliente realmente entregó).
  if (aFavor > 0 || aplicacion > 0) {
    var ajuste = deltaAFavor;
    if (formaPago === 'Efectivo') ef = Math.max(0, ef + ajuste);
    else if (formaPago === 'Transferencia') tr = Math.max(0, tr + ajuste);
    else if (formaPago === 'Mixto') {
      // Si es mixto, ajustamos del que ya tiene mayor monto (heurística simple)
      if (tr >= ef) tr = Math.max(0, tr + ajuste);
      else ef = Math.max(0, ef + ajuste);
    }
  }
  // ── SUMAR COBROS PARCIALES PREVIOS ──
  // Si el pedido tuvo pagos parciales (ej. Cristina Plate: $28K MP el sábado + $30.700 Ef hoy),
  // hay que sumar la composición real (ef/tr) de cada parcial al cobro final. Sin esto el
  // backend escribía solo el último ef/tr y ocultaba que hubo otro método antes.
  // Bug Cristina Plate 01/06/2026: cierre mostraba Ef $53.430 / Tr $0 cuando en realidad
  // fueron $28.000 Tr (sáb) + $31.000 Ef (hoy + propina).
  var sumEfPrev = 0, sumTrPrev = 0;
  var propEfPrev = 0, propTrPrev = 0;
  var shCP_prev = SS.getSheetByName('Cobros Parciales');
  if (shCP_prev && shCP_prev.getLastRow() > 1) {
    var dataCPprev = shCP_prev.getDataRange().getValues();
    for (var rcpP = 1; rcpP < dataCPprev.length; rcpP++) {
      if (String(dataCPprev[rcpP][1]).trim() === hoja && String(dataCPprev[rcpP][2]).trim() === pedidoId) {
        var mP = Number(dataCPprev[rcpP][5]) || 0;
        var fpP = String(dataCPprev[rcpP][4]).trim();
        if (fpP === 'Efectivo') sumEfPrev += mP;
        else sumTrPrev += mP;
      }
    }
  }
  var efFinal = ef + sumEfPrev;
  var trFinal = tr + sumTrPrev;
  if (efFinal > 0 || trFinal > 0) {
    sh.getRange(row, cols.ef).setValue(efFinal);
    sh.getRange(row, cols.tr).setValue(trFinal);
  }
  // Si la composición real es Mixta (hubo ef Y tr en algún momento), forzar fp = Mixto.
  if (efFinal > 0 && trFinal > 0 && formaPago !== 'Mixto') {
    sh.getRange(row, cols.fp).setValue('Mixto');
  }
  if (propina !== 0 && propMet) {
    // Propina puede ser negativa (ajuste de redondeo).
    var colProp = propMet === 'Efectivo' ? cols.propEf : cols.propTr;
    sh.getRange(row, colProp).setValue(propina);
  }
  // Escribir A Favor / Aplicado en col BI (Home, 61) o BL (Pilar, 64). Solo aplica a VD.
  // BE/BH son cols de Tartas (TP), NO usar ahí.
  if (isVD && deltaAFavor !== 0) {
    var colAFav = (hoja === 'Pilar') ? 64 : 61;
    var actual = Number(sh.getRange(row, colAFav).getValue()) || 0;
    sh.getRange(row, colAFav).setValue(actual + deltaAFavor);
  }

  // ── RECALCULAR FACTURADO (V=22) tras el cobro ──
  // Facturado = Total a cobrar (Q) + Propina Ef (T) + Propina Tr (U) + A Favor/Aplicado (BI/BL).
  // Se re-setea como fórmula en cada cobro de VD para que SIEMPRE descuente el saldo
  // a favor aplicado. Bug De Bary (08/06/2026): el a favor se aplicaba al cobrar pero
  // el Facturado no lo restaba (quedó $69k en vez de $60k) cuando el pedido se había
  // creado sin saldo a favor (la fórmula original no tenía el término del a favor).
  if (isVD) {
    var afLetter = (hoja === 'Pilar') ? 'BL' : 'BI';
    sh.getRange(row, 22).setFormula('=Q' + row + '+T' + row + '+U' + row + '+' + afLetter + row);
  }

  // ── CAMBIO DESDE LA BILLETERA (mismo método: pagó efectivo, le di vuelto) ──
  // El billete recibido va a caja fuerte; el vuelto sale de la billetera (fondo de cambio).
  // Bajamos la col D del último Saldo Base EN SU LUGAR (no append) para NO resetear el
  // cutoff de los cobros vivos. El total de efectivo no cambia por esta resta (el cobro
  // ya registró el neto); solo se reasigna el sub-saldo billetera.
  var cambioBilletera = Number(data.cambioBilletera) || 0;
  if (cambioBilletera > 0) {
    var shSBcb = SS.getSheetByName('Saldo Base');
    if (shSBcb && shSBcb.getLastRow() > 1 && shSBcb.getLastColumn() >= 4) {
      var lrSBcb = shSBcb.getLastRow();
      var bilActCb = Number(shSBcb.getRange(lrSBcb, 4).getValue()) || 0;
      shSBcb.getRange(lrSBcb, 4).setValue(Math.max(0, bilActCb - cambioBilletera));
    }
    // Log del vuelto (fecha + pedido + repartidor) para que el arqueo reconstruya la
    // billetera inicial del día y el desglose explique por qué el efectivo cobrado (neto)
    // es menor a los billetes que recibiste en mano.
    _appendCambioBilletera(hoja, pedidoId, data.cliente, cambioBilletera, data.repartidor, 'Billetera');
  }

  // ── CAMBIO CRUZADO (Ef↔MP) ──
  // Caso: cliente paga con efectivo de más, Tadeo devuelve cambio por MP (porque
  // no tenía cambio físico). Físicamente: entra Ef extra y sale MP. Para que la
  // caja refleje la realidad, creamos un par balanceador: Ingreso Ef + Egreso MP
  // del mismo monto, categoría "Cambio cruzado". Ej: pedido $13.230 + propina
  // $770 = cobro $14.000 Ef. Cliente dio $20.000 Ef → Tadeo devolvió $6.000 MP.
  // → cambioMP=6000 → Ingreso $6.000 Ef + Egreso $6.000 MP.
  // Y al revés: cambioEf=N significa que el cliente pago MP de más y Tadeo le
  // devolvió cambio físico → Ingreso MP + Egreso Ef.
  var cambioMP = Number(data.cambioMP) || 0;
  var cambioEf = Number(data.cambioCruzadoEf) || 0;
  if (cambioMP > 0 || cambioEf > 0) {
    var shIng = SS.getSheetByName('Ingresos');
    var shEgr = SS.getSheetByName('Egresos');
    var ahoraCC = new Date();
    var argNowCC = new Date(ahoraCC.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
    var fechaCC = Utilities.formatDate(argNowCC, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');
    var semanaCC = _isoWeek(argNowCC);
    var MESES_CC = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
    var mesCC = MESES_CC[argNowCC.getMonth()];
    var clienteCC = String(data.cliente || sh.getRange(row, 7).getValue() || '').trim();
    var concepto = clienteCC + ' (' + hoja + ' #' + pedidoId + ')';
    if (cambioMP > 0) {
      // Cliente pagó Ef de más → Tadeo le devolvió MP. Entra Ef, sale MP.
      if (shIng) shIng.appendRow([fechaCC, semanaCC, mesCC, 'Cambio cruzado', 'Cambio Ef recibido extra · ' + concepto, 'Efectivo', cambioMP, '']);
      if (shEgr) shEgr.appendRow([fechaCC, semanaCC, mesCC, 'Cambio cruzado', 'Cambio MP devuelto · ' + concepto, 'Mercado Pago', cambioMP, '']);
      // Efectivo extra que te quedó en mano → el arqueo lo suma.
      _appendCambioBilletera(hoja, pedidoId, clienteCC, cambioMP, data.repartidor, 'CambioMP');
    }
    if (cambioEf > 0) {
      // Cliente pagó MP de más → Tadeo le devolvió Ef. Entra MP, sale Ef.
      if (shIng) shIng.appendRow([fechaCC, semanaCC, mesCC, 'Cambio cruzado', 'Cambio MP recibido extra · ' + concepto, 'Mercado Pago', cambioEf, '']);
      if (shEgr) shEgr.appendRow([fechaCC, semanaCC, mesCC, 'Cambio cruzado', 'Cambio Ef devuelto · ' + concepto, 'Efectivo', cambioEf, '']);
      // Efectivo que salió de tu mano (devuelto por un pago MP) → el arqueo lo resta.
      _appendCambioBilletera(hoja, pedidoId, clienteCC, cambioEf, data.repartidor, 'CruzadoEf');
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true, total: totalRecalculado, aFavor: aFavor, aplicacion: aplicacion, cambioMP: cambioMP, cambioEf: cambioEf })).setMimeType(ContentService.MimeType.JSON);
}

/** POST action=descobrar — DESHACE un cobro: vuelve el pedido a "No Cobrado" y limpia
 *  el cobro (ef/tr/propinas/Fecha de Cobro/A Favor). El pedido reaparece en COBROS para
 *  re-cobrarlo con la forma/monto correctos. Solo Home/Pilar/Clubes (Red se liquida aparte).
 *  Body: { hoja, id, row } */
function _doPostDescobrar(data) {
  var hoja = String(data.hoja || '');
  var pedidoId = String(data.id || '');
  if (hoja !== 'Home' && hoja !== 'Pilar' && hoja !== 'Clubes') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'solo Home/Pilar/Clubes' })).setMimeType(ContentService.MimeType.JSON);
  }
  var sh = SS.getSheetByName(hoja);
  if (!sh) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'hoja' })).setMimeType(ContentService.MimeType.JSON);
  var allData = sh.getDataRange().getValues();
  var row = Number(data.row) || -1;
  if (row < 2 || row > allData.length || String(allData[row - 1][1]).trim() !== pedidoId) {
    row = -1;
    for (var r = allData.length - 1; r >= 1; r--) {
      if (String(allData[r][1]).trim() === pedidoId) { row = r + 1; break; }
    }
  }
  if (row === -1) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'no encontrado' })).setMimeType(ContentService.MimeType.JSON);
  var isVD = (hoja === 'Home' || hoja === 'Pilar');
  // Cols 1-based por hoja (coherente con _doPostMarcarCobrado).
  var c = (hoja === 'Clubes')
    ? { estPago: 16, total: 17, ef: 19, tr: 20, propEf: 21, propTr: 22, fact: 23 }
    : { estPago: 13, total: 17, ef: 18, tr: 19, propEf: 20, propTr: 21, fact: 22 }; // Home/Pilar
  if (String(sh.getRange(row, c.estPago).getValue()).trim() !== 'Cobrado') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'el pedido no está cobrado' })).setMimeType(ContentService.MimeType.JSON);
  }
  sh.getRange(row, c.estPago).setValue('No Cobrado');
  sh.getRange(row, c.ef).setValue(0);
  sh.getRange(row, c.tr).setValue(0);
  sh.getRange(row, c.propEf).setValue(0);
  sh.getRange(row, c.propTr).setValue(0);
  // A Favor / Aplicado (solo VD): limpiar la aplicación de este pedido.
  if (isVD) {
    var colAF = (hoja === 'Pilar') ? 64 : 61;
    if (sh.getLastColumn() >= colAF) sh.getRange(row, colAF).setValue(0);
  }
  // Facturado = Total a cobrar (sin propinas/aFavor) — pedido vuelve a estado pendiente limpio.
  var totalCob = Number(sh.getRange(row, c.total).getValue()) || 0;
  sh.getRange(row, c.fact).setValue(totalCob);
  // Limpiar Fecha de Cobro (col por header).
  var hdrs = allData[0];
  for (var hh = 0; hh < hdrs.length; hh++) {
    if (String(hdrs[hh]).trim() === 'Fecha de Cobro') { sh.getRange(row, hh + 1).setValue(''); break; }
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

/** POST action=cobrarParcial — registra un cobro parcial de un pedido.
 *  Recibe: { action:'cobrarParcial', hoja, id, row, formaPago, monto, notas }
 *  - Inserta fila en hoja "Cobros Parciales".
 *  - Si la suma de parciales >= total del pedido → marca el pedido como Cobrado
 *    y distribuye los cobros en cols ef/tr según las formas de pago de los parciales.
 *  - Si no, deja el pedido como No Cobrado.
 *  Devuelve { ok, restante, totalCobrado, total }. */
function _doPostCobrarParcial(data) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, retry: true, err: 'lock' })).setMimeType(ContentService.MimeType.JSON);
  }
  try {
    var hoja = String(data.hoja || '');
    var pedidoId = String(data.id || '');
    var formaPago = String(data.formaPago || 'Transferencia').trim();
    var monto = Number(data.monto) || 0;
    var notas = String(data.notas || '').trim();

    if (!hoja || !pedidoId || monto <= 0) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'datos invalidos' })).setMimeType(ContentService.MimeType.JSON);
    }

    var sh = SS.getSheetByName(hoja);
    if (!sh) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'hoja' })).setMimeType(ContentService.MimeType.JSON);

    // Localizar fila
    var row = Number(data.row) || 0;
    if (row > 1 && String(sh.getRange(row, 2).getValue()).trim() === pedidoId) {
      // hint válido
    } else {
      var lastRow = sh.getLastRow();
      if (lastRow < 2) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'hoja vacia' })).setMimeType(ContentService.MimeType.JSON);
      var ids = sh.getRange(2, 2, lastRow - 1, 1).getValues();
      row = -1;
      for (var i = ids.length - 1; i >= 0; i--) {
        if (String(ids[i][0]).trim() === pedidoId) { row = i + 2; break; }
      }
      if (row === -1) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'no encontrado' })).setMimeType(ContentService.MimeType.JSON);
    }

    // Columnas según hoja (1-based). total = col que contiene el monto a cobrar
    // (mismo que usa _doGetCobrosPendientes para mostrar el "$").
    // Home/Pilar: V(22)=Facturado · Clubes: Q(17)=Total · Red: O(15)=Total
    var cols;
    var isVD = (hoja === 'Home' || hoja === 'Pilar');
    if (hoja === 'Clubes')   cols = {fp:15, estPago:16, total:17, ef:19, tr:20};
    else if (hoja === 'Red') cols = {fp:13, estPago:14, total:15, ef:17, tr:18};
    else                     cols = {fp:12, estPago:13, total:22, ef:18, tr:19}; // Home/Pilar v2

    // Total del pedido (col Total)
    var totalPedido = Number(sh.getRange(row, cols.total).getValue()) || 0;
    var cliente = String(sh.getRange(row, 8).getValue()).trim(); // col H (8) = Cliente

    // Insertar parcial en hoja "Cobros Parciales"
    var shCP = SS.getSheetByName('Cobros Parciales');
    if (!shCP) {
      shCP = SS.insertSheet('Cobros Parciales');
      shCP.getRange(1, 1, 1, 7).setValues([['Fecha/Hora','Hoja','N° Pedido','Cliente','Forma Pago','Monto','Notas']]);
      shCP.setFrozenRows(1);
    }
    var argNow = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
    shCP.appendRow([argNow, hoja, pedidoId, cliente, formaPago, monto, notas]);
    var lastCP = shCP.getLastRow();
    shCP.getRange(lastCP, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    shCP.getRange(lastCP, 6).setNumberFormat('$#,##0');

    // Sumar todos los parciales del pedido (incluido el recién insertado)
    var dataCP = shCP.getDataRange().getValues();
    var totalCobrado = 0;
    var sumEf = 0, sumTr = 0;
    for (var rcp = 1; rcp < dataCP.length; rcp++) {
      if (String(dataCP[rcp][1]).trim() === hoja && String(dataCP[rcp][2]).trim() === pedidoId) {
        var m = Number(dataCP[rcp][5]) || 0;
        var fp = String(dataCP[rcp][4]).trim();
        totalCobrado += m;
        if (fp === 'Efectivo') sumEf += m;
        else sumTr += m;
      }
    }

    var restante = Math.round(totalPedido - totalCobrado);
    var cerrado = (restante <= 0);

    if (cerrado) {
      // Cerrar el pedido: marcar Cobrado, escribir ef/tr distribuidos, estampar fecha
      sh.getRange(row, cols.estPago).setValue('Cobrado');
      _stampFechaCobro(sh, row);
      sh.getRange(row, cols.ef).setValue(sumEf);
      sh.getRange(row, cols.tr).setValue(sumTr);
      // Forma de pago: si hubo dos métodos, dejar Mixto; sino el único usado
      var fpFinal = (sumEf > 0 && sumTr > 0) ? 'Mixto' : (sumEf > 0 ? 'Efectivo' : 'Transferencia');
      sh.getRange(row, cols.fp).setValue(fpFinal);
    }

    SpreadsheetApp.flush();
    return ContentService.createTextOutput(JSON.stringify({
      ok: true,
      total: totalPedido,
      totalCobrado: totalCobrado,
      restante: cerrado ? 0 : restante,
      cerrado: cerrado
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

/** POST action=ajusteSaldo — guarda el saldo real como snapshot.
 *  El Panel calcula el saldo vivo como: SaldoBase + cobrado/ingreso/gasto DESPUÉS de la fecha del snapshot.
 *  Así, cobros históricos ya reflejados en el monto ajustado no cuentan dos veces.
 *  Acepta opcionalmente `billetera`: sub-saldo del Efectivo que está en la billetera
 *  (vs caja fuerte). Si no viene, hereda el último valor conocido. */
function _doPostAjusteSaldo(data) {
  var deseadoEf = Number(data.efectivo) || 0;
  var deseadoMP = Number(data.mp) || 0;

  var shSaldo = SS.getSheetByName('Saldo Base');
  if (!shSaldo) {
    shSaldo = SS.insertSheet('Saldo Base');
    shSaldo.getRange(1, 1, 1, 5).setValues([['Fecha', 'Efectivo', 'Mercado Pago', 'Billetera', 'Sobres']]);
    shSaldo.setFrozenRows(1);
    shSaldo.getRange(1, 1, 1, 5).setBackground(BROWN).setFontColor('#FFFFFF').setFontWeight('bold');
  }
  // Migración suave: agregar headers de cols E (Sobres) y F (Inversiones) si faltan.
  if (shSaldo.getLastColumn() < 5) shSaldo.getRange(1, 5).setValue('Sobres');
  if (shSaldo.getLastColumn() < 6) shSaldo.getRange(1, 6).setValue('Inversiones');
  var lastRow = shSaldo.getLastRow();
  // Billetera, Sobres, Inversiones: si el cliente los manda, los usa; sino hereda el último.
  function _heredar(col) {
    return (lastRow > 1 && shSaldo.getLastColumn() >= col) ? (Number(shSaldo.getRange(lastRow, col).getValue()) || 0) : 0;
  }
  var deseadoBil = (data.billetera !== undefined && data.billetera !== null) ? (Number(data.billetera) || 0) : _heredar(4);
  var deseadoSob = (data.sobres !== undefined && data.sobres !== null) ? (Number(data.sobres) || 0) : _heredar(5);
  var deseadoInv = (data.inversiones !== undefined && data.inversiones !== null) ? (Number(data.inversiones) || 0) : _heredar(6);
  if (deseadoBil < 0) deseadoBil = 0;
  if (deseadoSob < 0) deseadoSob = 0;
  if (deseadoInv < 0) deseadoInv = 0;
  // Caja fuerte = ef - bil - sob >= 0. Si la suma de sub-saldos excede el efectivo, capar.
  if (deseadoBil > deseadoEf) deseadoBil = deseadoEf;
  if (deseadoBil + deseadoSob > deseadoEf) deseadoSob = Math.max(0, deseadoEf - deseadoBil);

  var ahora = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  shSaldo.appendRow([argNow, deseadoEf, deseadoMP, deseadoBil, deseadoSob, deseadoInv]);
  shSaldo.getRange(shSaldo.getLastRow(), 1).setNumberFormat('dd/MM/yyyy HH:mm');
  // Sin flush(): UI optimista en frontend.

  return ContentService.createTextOutput(JSON.stringify({ ok: true, ef: deseadoEf, mp: deseadoMP, bil: deseadoBil, sob: deseadoSob, inv: deseadoInv })).setMimeType(ContentService.MimeType.JSON);
}

/** POST action=moverBilletera — transferir plata entre Caja Fuerte y Billetera.
 *  NO cambia el saldo Efectivo total — solo ajusta la sub-cuenta Billetera.
 *  Body: { monto: N, dir: 'aBilletera' | 'aCajaFuerte' }
 *  Append una fila a Saldo Base con efectivo y mp iguales al último snapshot
 *  pero billetera ajustada. Así la auditoría queda en el historial. */
function _doPostMoverBilletera(data) {
  var monto = Number(data.monto) || 0;
  var dir = String(data.dir || '').trim();
  if (monto <= 0) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'monto invalido' })).setMimeType(ContentService.MimeType.JSON);
  if (dir !== 'aBilletera' && dir !== 'aCajaFuerte') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'dir invalida' })).setMimeType(ContentService.MimeType.JSON);
  }

  var shSaldo = SS.getSheetByName('Saldo Base');
  if (!shSaldo) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'sin Saldo Base' })).setMimeType(ContentService.MimeType.JSON);
  var lastRow = shSaldo.getLastRow();
  if (lastRow < 2) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'sin snapshot previo — ajusta saldo primero' })).setMimeType(ContentService.MimeType.JSON);

  var bilActual = (shSaldo.getLastColumn() >= 4) ? (Number(shSaldo.getRange(lastRow, 4).getValue()) || 0) : 0;
  // Preservar Sobres (E) e Inversiones (F) al appendear — si no, se perderían esos sub-saldos.
  var sobActual = (shSaldo.getLastColumn() >= 5) ? (Number(shSaldo.getRange(lastRow, 5).getValue()) || 0) : 0;
  var invActual = (shSaldo.getLastColumn() >= 6) ? (Number(shSaldo.getRange(lastRow, 6).getValue()) || 0) : 0;

  // BUG FIX 18/05/2026: antes heredábamos ef/mp del snapshot anterior, lo que
  // descartaba todos los cobros/gastos posteriores (caja vivía caía al valor
  // del último snapshot). Ahora el frontend debe pasar saldoVivoEf y saldoVivoMP
  // calculados al momento del click (los que el Panel muestra en pantalla).
  // Si no llegan, fallback al snapshot anterior (compat con clientes viejos).
  var efAUsar, mpAUsar;
  if (data.saldoVivoEf !== undefined && data.saldoVivoEf !== null) {
    efAUsar = Math.round(Number(data.saldoVivoEf) || 0);
  } else {
    efAUsar = Number(shSaldo.getRange(lastRow, 2).getValue()) || 0;
  }
  if (data.saldoVivoMP !== undefined && data.saldoVivoMP !== null) {
    mpAUsar = Math.round(Number(data.saldoVivoMP) || 0);
  } else {
    mpAUsar = Number(shSaldo.getRange(lastRow, 3).getValue()) || 0;
  }

  // Cálculo nuevo billetera
  var bilNuevo;
  if (dir === 'aBilletera') {
    bilNuevo = bilActual + monto;
    if (bilNuevo > efAUsar) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'monto excede caja fuerte disponible' })).setMimeType(ContentService.MimeType.JSON);
  } else {
    bilNuevo = bilActual - monto;
    if (bilNuevo < 0) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'monto excede billetera' })).setMimeType(ContentService.MimeType.JSON);
  }

  var ahora = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  shSaldo.appendRow([argNow, efAUsar, mpAUsar, bilNuevo, sobActual, invActual]);
  shSaldo.getRange(shSaldo.getLastRow(), 1).setNumberFormat('dd/MM/yyyy HH:mm');

  return ContentService.createTextOutput(JSON.stringify({ ok: true, ef: efAUsar, mp: mpAUsar, bil: bilNuevo, sob: sobActual, inv: invActual })).setMimeType(ContentService.MimeType.JSON);
}

/** POST action=invertir — mueve plata entre MP LÍQUIDO (col C) e INVERSIONES (col F)
 *  SIN cambiar la Posición Neta (solo se reubica). Edita EN SU LUGAR el último Saldo
 *  Base (no appendea → no resetea el cutoff de cobros vivos).
 *  Body: { monto:N, dir:'invertir' | 'rescatar' }
 *    invertir = MP líquido → Inversiones (baja MP, sube Inversiones)
 *    rescatar = Inversiones → MP líquido (al revés) */
function _doPostInvertir(data) {
  var monto = Math.round(Number(data.monto) || 0);
  var dir = String(data.dir || 'invertir').trim();
  if (monto <= 0) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'monto invalido' })).setMimeType(ContentService.MimeType.JSON);
  if (dir !== 'invertir' && dir !== 'rescatar') return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'dir invalida' })).setMimeType(ContentService.MimeType.JSON);
  var shSaldo = SS.getSheetByName('Saldo Base');
  if (!shSaldo || shSaldo.getLastRow() < 2) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'sin snapshot previo — ajusta saldo primero' })).setMimeType(ContentService.MimeType.JSON);
  if (shSaldo.getLastColumn() < 6) shSaldo.getRange(1, 6).setValue('Inversiones');
  var lr = shSaldo.getLastRow();
  var mp  = Number(shSaldo.getRange(lr, 3).getValue()) || 0;   // C = MP líquido (base)
  var inv = Number(shSaldo.getRange(lr, 6).getValue()) || 0;   // F = Inversiones
  // OJO: mp acá es el saldo BASE, no el vivo. Invertir baja la BASE de MP; el saldo
  // vivo de MP (base + cobros − gastos) baja en la misma medida. La validación contra
  // el líquido real la hace el frontend (que conoce el MP vivo); acá solo evitamos negativos absurdos.
  if (dir === 'invertir') {
    mp = mp - monto;   // puede quedar la base baja pero el vivo lo cubre con cobros
    inv = inv + monto;
  } else {
    if (inv - monto < 0) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'no tenés tanto invertido' })).setMimeType(ContentService.MimeType.JSON);
    mp = mp + monto;
    inv = inv - monto;
  }
  shSaldo.getRange(lr, 3).setValue(mp);
  shSaldo.getRange(lr, 6).setValue(inv);
  return ContentService.createTextOutput(JSON.stringify({ ok: true, mp: mp, inv: inv })).setMimeType(ContentService.MimeType.JSON);
}

/** POST action=moverEfectivo — mover plata entre las 3 cuentas de efectivo
 *  (cajaFuerte / billetera / sobres) SIN cambiar el total. Edita los sub-saldos
 *  (col D billetera, col E sobres) del ÚLTIMO Saldo Base EN SU LUGAR — no appendea,
 *  así NO resetea el cutoff de cobros vivos. Caja fuerte se deriva (ef - bil - sob).
 *  Body: { monto:N, desde:'cajaFuerte'|'billetera'|'sobres', hacia:'...' } */
function _doPostMoverEfectivo(data) {
  var monto = Math.round(Number(data.monto) || 0);
  var desde = String(data.desde || '').trim();
  var hacia = String(data.hacia || '').trim();
  var validos = { cajaFuerte: 1, billetera: 1, sobres: 1 };
  if (monto <= 0) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'monto invalido' })).setMimeType(ContentService.MimeType.JSON);
  if (!validos[desde] || !validos[hacia] || desde === hacia) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'origen/destino invalido' })).setMimeType(ContentService.MimeType.JSON);

  var shSaldo = SS.getSheetByName('Saldo Base');
  if (!shSaldo || shSaldo.getLastRow() < 2) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'sin snapshot previo — ajusta saldo primero' })).setMimeType(ContentService.MimeType.JSON);
  if (shSaldo.getLastColumn() < 5) shSaldo.getRange(1, 5).setValue('Sobres');
  var lr = shSaldo.getLastRow();
  var ef  = Number(shSaldo.getRange(lr, 2).getValue()) || 0;
  var bil = Number(shSaldo.getRange(lr, 4).getValue()) || 0;
  var sob = Number(shSaldo.getRange(lr, 5).getValue()) || 0;
  var cf  = ef - bil - sob;
  var saldos = { cajaFuerte: cf, billetera: bil, sobres: sob };
  if (saldos[desde] < monto) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'no hay suficiente en ' + desde })).setMimeType(ContentService.MimeType.JSON);
  saldos[desde] -= monto;
  saldos[hacia] += monto;
  // CF se deriva: solo persistimos billetera (col D) y sobres (col E). El total (ef) no cambia.
  shSaldo.getRange(lr, 4).setValue(saldos.billetera);
  shSaldo.getRange(lr, 5).setValue(saldos.sobres);

  return ContentService.createTextOutput(JSON.stringify({ ok: true, cf: saldos.cajaFuerte, bil: saldos.billetera, sob: saldos.sobres })).setMimeType(ContentService.MimeType.JSON);
}

// Normaliza un nombre para comparar (minúsculas, sin tildes, sin espacios extra).
function _normNombre(s) {
  s = String(s || '').toLowerCase().trim().replace(/\s+/g, ' ');
  if (s.normalize) s = s.normalize('NFD').replace(/[̀-ͯ]/g, '');
  return s;
}

// Asegura la hoja "Sobres" (detalle de sobres preparados para proveedores).
function _ensureSobresSheet() {
  var sh = SS.getSheetByName('Sobres');
  if (!sh) {
    sh = SS.insertSheet('Sobres');
    sh.getRange(1, 1, 1, 6).setValues([['Fecha', 'Proveedor', 'Monto', 'Estado', 'Fecha Pago', 'Notas']]);
    sh.setFrozenRows(1);
    if (typeof BROWN !== 'undefined') sh.getRange(1, 1, 1, 6).setBackground(BROWN).setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sh;
}
// Suma el sub-saldo Sobres (col E del último Saldo Base) en su lugar.
function _ajustarColSobres(delta) {
  var shSB = SS.getSheetByName('Saldo Base');
  if (!shSB || shSB.getLastRow() < 2) return null;
  if (shSB.getLastColumn() < 5) shSB.getRange(1, 5).setValue('Sobres');
  var lr = shSB.getLastRow();
  var sob = Number(shSB.getRange(lr, 5).getValue()) || 0;
  var nuevo = Math.max(0, sob + delta);
  shSB.getRange(lr, 5).setValue(nuevo);
  return nuevo;
}

/** POST action=prepararSobre — aparta efectivo en un sobre para un proveedor.
 *  Mueve plata de Caja fuerte → Sobres (el total de efectivo NO cambia) y registra
 *  el detalle (a quién va) en la hoja "Sobres". Body: { proveedor, monto, notas? } */
function _doPostPrepararSobre(data) {
  var prov = String(data.proveedor || '').trim();
  var monto = Math.round(Number(data.monto) || 0);
  if (!prov) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'falta proveedor' })).setMimeType(ContentService.MimeType.JSON);
  if (monto <= 0) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'monto invalido' })).setMimeType(ContentService.MimeType.JSON);
  var sh = _ensureSobresSheet();
  var argNow = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  sh.appendRow([argNow, prov, monto, 'Activo', '', String(data.notas || '')]);
  var lr = sh.getLastRow();
  sh.getRange(lr, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  sh.getRange(lr, 3).setNumberFormat('$#,##0');
  var sobTotal = _ajustarColSobres(monto); // CF → Sobres
  return ContentService.createTextOutput(JSON.stringify({ ok: true, proveedor: prov, monto: monto, sobTotal: sobTotal, row: lr })).setMimeType(ContentService.MimeType.JSON);
}

/** POST action=eliminarSobre — cancela un sobre activo (la plata vuelve a Caja fuerte).
 *  Body: { row } (fila en la hoja Sobres) */
function _doPostEliminarSobre(data) {
  var row = Number(data.row) || 0;
  var sh = SS.getSheetByName('Sobres');
  if (!sh || row < 2 || row > sh.getLastRow()) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'sobre invalido' })).setMimeType(ContentService.MimeType.JSON);
  var estado = String(sh.getRange(row, 4).getValue()).trim();
  if (estado !== 'Activo') return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'el sobre ya no esta activo' })).setMimeType(ContentService.MimeType.JSON);
  var monto = Number(sh.getRange(row, 3).getValue()) || 0;
  sh.getRange(row, 4).setValue('Cancelado');
  var sobTotal = _ajustarColSobres(-monto); // Sobres → CF
  return ContentService.createTextOutput(JSON.stringify({ ok: true, sobTotal: sobTotal })).setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════════
//  VENTAS — Lee hojas operativas acumulativas (solo Entregados)
// ══════════════════════════════════════════════════════════════

/** GET ?action=ventas
 *  Devuelve TODAS las ventas Entregadas de las hojas acumulativas.
 *  Canales: Venta Directa (Home+Pilar+CF), Clubes, Red, B2B, Catering. */
function _doGetVentas() {
  var ventas = [];
  var MVAL = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

  // Normaliza cualquier fecha (Date o String) a "dd/MM/yyyy" en TZ AR.
  // - Si es Date → formatea sin hora.
  // - Si viene como "9/5/2026 19:13" → corta la hora y padea dia/mes (→ "09/05/2026").
  // Imprescindible para que el filtro de Fecha del Panel matchee strings exactos.
  function _fmtF(raw) {
    if (raw instanceof Date) return Utilities.formatDate(raw, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
    var s = String(raw || '').trim();
    if (!s) return '';
    s = s.split(' ')[0]; // sacar hora si la tiene ("09/05/2026 19:13" → "09/05/2026")
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return ('0' + m[1]).slice(-2) + '/' + ('0' + m[2]).slice(-2) + '/' + m[3];
    return s;
  }

  // Devuelve el nombre del mes (ej. "Mayo") desde una fecha Date o string "dd/MM/yyyy".
  // Es la base del criterio de reconocimiento: la venta cuenta en el mes en que se ENTREGÓ,
  // no en el mes en que se pidió. Aplica a todos los canales por consistencia.
  function _mesFromAny(raw) {
    var d = null;
    if (raw instanceof Date) d = raw;
    else if (raw) {
      var s = String(raw).trim().split(' ')[0];
      var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (m) d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    }
    if (!d || isNaN(d.getTime())) return '';
    return MVAL[d.getMonth()];
  }

  // ── Helper: parsear fila VD (Home, Pilar) — v2 abr/2026 ──
  // Home v2: Facturado V(21), Ef R(17), Tr S(18), Costo AP(41), Margen AQ(42), FechaEnt AX(49), MesEnt AY(50), SemEnt AZ(51)
  // Pilar v2: Facturado V(21), Ef R(17), Tr S(18), Costo AT(45), Margen AU(46), FechaEnt BA(52), MesEnt BB(53), SemEnt BC(54)
  function parseVD(data, r, zona) {
    var cliente = String(data[r][7] || '').trim();
    if (!cliente) return null;
    var isP = (zona === 'Pilar');
    // Home: fCob=54 (col BC). Pilar: fCob=57 (col BF).
    // Home: barrio=43 (AR), subBarrio=44 (AS). Pilar: barrio=47 (AV), no tiene sub.
    var IDX = isP
      ? { fact:21, totalAlt:16, ef:17, tr:18, costo:45, margen:46, fechaEnt:52, mesEnt:53, semEnt:54, fCob:57, barrio:47, subBarrio:-1 }
      : { fact:21, totalAlt:16, ef:17, tr:18, costo:41, margen:42, fechaEnt:49, mesEnt:50, semEnt:51, fCob:54, barrio:43, subBarrio:44 };
    var facturado = Number(data[r][IDX.fact]) || Number(data[r][IDX.totalAlt]) || 0;
    if (facturado === 0) return null;
    var fechaEnt = data[r][IDX.fechaEnt];
    var fechaRaw = (fechaEnt instanceof Date || (typeof fechaEnt === 'string' && fechaEnt.trim())) ? fechaEnt : data[r][3];
    var fechaStr = _fmtF(fechaRaw);
    var fCobStr = _fmtF(data[r][IDX.fCob]);
    var mesEnt = String(data[r][IDX.mesEnt] || '').trim();
    var semEnt = Number(data[r][IDX.semEnt]) || 0;
    var mesPed = String(data[r][4] || '').trim();
    var semPed = Number(data[r][5]) || 0;
    var mes = MVAL.indexOf(mesEnt) >= 0 ? mesEnt : mesPed;
    var sem = semEnt > 0 ? semEnt : semPed;
    var barrio = IDX.barrio >= 0 ? String(data[r][IDX.barrio] || '').trim() : '';
    var subBarrio = IDX.subBarrio >= 0 ? String(data[r][IDX.subBarrio] || '').trim() : '';
    // Propinas: Home/Pilar tienen Propina Ef en col T (idx 19) y Propina Tr en col U (idx 20).
    // El facturado (col V) las suma; ef/tr NO. Exponemos pEf/pTr para que el cliente
    // pueda reconciliar Cobrado = ef + tr + pEf + pTr = Facturado.
    return {
      canal: 'Venta Directa', zona: zona, fecha: fechaStr, fCob: fCobStr, mes: mes, sem: sem,
      cliente: cliente, estado: String(data[r][10] || '').trim(),
      fp: String(data[r][11] || '').trim(), ep: String(data[r][12] || '').trim(),
      $: facturado, ef: Number(data[r][IDX.ef]) || 0, tr: Number(data[r][IDX.tr]) || 0,
      pEf: Number(data[r][19]) || 0, pTr: Number(data[r][20]) || 0,
      costo: Number(data[r][IDX.costo]) || 0, margen: Number(data[r][IDX.margen]) || 0,
      barrio: barrio, subBarrio: subBarrio
    };
  }

  // ── Helper: parsear fila Clubes ──
  function parseClubes(data, r) {
    var cli = String(data[r][7] || '').trim();
    if (!cli) return null;
    var club = String(data[r][8] || '').trim();
    var fac = Number(data[r][22]) || Number(data[r][16]) || 0;
    if (fac === 0) return null;
    var fCS = _fmtF(data[r][3]);
    // Clubes: Propina Ef en col U (idx 20), Propina Tr en col V (idx 21). Fecha de Cobro col 35 (idx 34).
    var fCobCS = _fmtF(data[r][34]);
    // Mes de la venta = mes de ENTREGA (col 12 = M = Día de Entrega), no mes del pedido.
    // Criterio consistente con Home/Pilar. Fallback al mes del pedido si la fecha de entrega
    // no está cargada (típico de pedidos recién entrados sin agendar).
    var mesEntC = _mesFromAny(data[r][12]);
    var mesPedC = String(data[r][4] || '').trim();
    return {
      canal: 'Clubes', zona: club, fecha: fCS, fCob: fCobCS,
      mes: mesEntC || mesPedC, sem: Number(data[r][5]) || 0,
      cliente: cli + (club ? ' (' + club + ')' : ''),
      estado: String(data[r][13] || '').trim(),
      fp: String(data[r][14] || '').trim(), ep: String(data[r][15] || '').trim(),
      $: fac, ef: Number(data[r][18]) || 0, tr: Number(data[r][19]) || 0,
      pEf: Number(data[r][20]) || 0, pTr: Number(data[r][21]) || 0,
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

  // ── Venta Directa: operativa acumulativa, solo Entregados ──
  var vdCfg = [
    { zona: 'Home' },
    { zona: 'Pilar' }
  ];
  vdCfg.forEach(function(cfg) {
    readOperativa(cfg.zona, parseVD, cfg.zona, 10);
  });

  // ── Clubes: operativa acumulativa, solo Entregados ──
  readOperativa('Clubes', parseClubes, null, 13);

  // ── Red (55 cols). Facturación Red = "A Pagar" (AW, idx 48) = lo que le queda
  // a Maleu después de la comisión del vendedor. Costo = idx 44, Margen Neto = idx 47.
  var shRd = SS.getSheetByName('Red');
  if (shRd && shRd.getLastRow() > 1) {
    var dRd = shRd.getDataRange().getValues();
    for (var rr = 1; rr < dRd.length; rr++) {
      var estadoR = String(dRd[rr][11] || '').trim();
      if (estadoR !== 'Entregado') continue;
      var cliR = String(dRd[rr][8] || '').trim();
      if (!cliR) continue;
      var aPagarR = Number(dRd[rr][48]) || 0;
      if (aPagarR === 0) continue;
      var fRS = _fmtF(dRd[rr][3]);
      // Red: las propinas (S/T = idx 18/19) van 100% al vendedor, no a Maleu.
      // Dejamos pEf/pTr en 0 para no inflar el "Cobrado" de Red con plata que no es nuestra.
      // Fecha Pago a Maleu: idx 54 (col BC).
      var fCobRS = _fmtF(dRd[rr][54]);
      // Mes de la venta = mes de ENTREGA (col K idx 10 = Día de Entrega), no mes del pedido.
      // Criterio consistente con Home/Pilar/Clubes: la venta cuenta en el mes en que
      // efectivamente se entregó la mercadería al cliente final. Fallback al mes del pedido
      // si la fecha de entrega no está cargada.
      var mesEntR = _mesFromAny(dRd[rr][10]);
      var mesPedR = String(dRd[rr][4] || '').trim();
      ventas.push({
        canal: 'Red', zona: String(dRd[rr][7] || '').trim(), fecha: fRS, fCob: fCobRS,
        mes: mesEntR || mesPedR, sem: Number(dRd[rr][5]) || 0,
        cliente: cliR, estado: estadoR,
        fp: String(dRd[rr][12] || '').trim(),
        ep: String(dRd[rr][13] || '').trim(),
        $: aPagarR,
        ef: Number(dRd[rr][16]) || 0, tr: Number(dRd[rr][17]) || 0,
        pEf: 0, pTr: 0,
        costo: Number(dRd[rr][44]) || 0,
        margen: Number(dRd[rr][47]) || 0
      });
    }
  }

  // ── B2B (hoja operativa manual) ──
  var shB = SS.getSheetByName('B2B');
  if (shB && shB.getLastRow() > 1) {
    var dB = shB.getDataRange().getValues();
    for (var rb = 1; rb < dB.length; rb++) {
      var cliB = String(dB[rb][5] || '').trim();
      if (!cliB) continue;
      var facB = Number(dB[rb][8]) || 0;
      if (facB === 0) continue;
      var fBS = _fmtF(dB[rb][1]);
      ventas.push({
        canal: 'B2B', zona: '', fecha: fBS, fCob: fBS,
        mes: String(dB[rb][2] || '').trim(), sem: Number(dB[rb][3]) || 0,
        cliente: cliB, estado: 'Entregado',
        fp: Number(dB[rb][10]) > 0 ? 'Transferencia' : 'Efectivo', ep: 'Cobrado',
        $: facB, ef: Number(dB[rb][9]) || 0, tr: Number(dB[rb][10]) || 0,
        pEf: 0, pTr: 0,
        costo: 0, margen: 0
      });
    }
  }

  // ── Catering (hoja operativa manual) ──
  var shCt = SS.getSheetByName('Catering');
  if (shCt && shCt.getLastRow() > 1) {
    var dCt = shCt.getDataRange().getValues();
    for (var rt = 1; rt < dCt.length; rt++) {
      var cliT = String(dCt[rt][5] || '').trim();
      if (!cliT) continue;
      var facT = Number(dCt[rt][8]) || 0;
      if (facT === 0) continue;
      var fTS = _fmtF(dCt[rt][1]);
      ventas.push({
        canal: 'Catering', zona: String(dCt[rt][7] || '').trim(), fecha: fTS, fCob: fTS,
        mes: String(dCt[rt][2] || '').trim(), sem: Number(dCt[rt][3]) || 0,
        cliente: cliT, estado: 'Entregado',
        fp: Number(dCt[rt][11]) > 0 ? 'Transferencia' : 'Efectivo', ep: 'Cobrado',
        $: facT, ef: Number(dCt[rt][10]) || 0, tr: Number(dCt[rt][11]) || 0,
        pEf: 0, pTr: 0,
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

/** POST action=pagarVendedor — registra pago de un vendedor Red a Maleu.
 *  { action:'pagarVendedor', vendedor:'Marcos Bottcher', monto:230408, metodo:'Mercado Pago', fecha:'2026-04-21', notas:'' }
 *  - Crea 1 fila en Ingresos (categoría 'Liquidación Red')
 *  - Marca pedidos Red con Estado Pago a Maleu = 'Sí' en FIFO hasta cubrir el monto
 *  - Guarda Fecha Pago a Maleu y Forma Pago a Maleu en cada pedido marcado */
function _doPostPagarVendedor(data) {
  var vendedor = String(data.vendedor || '').trim();
  var monto = Number(data.monto) || 0;
  var metodo = String(data.metodo || 'Mercado Pago').trim();
  var fechaStr = String(data.fecha || '').trim();
  var notas = String(data.notas || '').trim();

  if (!vendedor || monto <= 0) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'Faltan datos' })).setMimeType(ContentService.MimeType.JSON);
  }

  // Leer comisión del vendedor
  var shVend = SS.getSheetByName('Vendedores');
  var comPct = 17;
  if (shVend && shVend.getLastRow() > 1) {
    var dV = shVend.getDataRange().getValues();
    for (var rv = 1; rv < dV.length; rv++) {
      if (String(dV[rv][0]).trim() === vendedor) { comPct = Number(dV[rv][9]) || 17; break; }
    }
  }

  // Fecha para registros
  var ahora = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var dd = String(argNow.getDate()).padStart(2, '0');
  var mm = String(argNow.getMonth() + 1).padStart(2, '0');
  var yyyy = argNow.getFullYear();
  var fechaHoy = dd + '/' + mm + '/' + yyyy;
  var MESES_V = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var mesNombre = MESES_V[argNow.getMonth()];
  var semana = _isoWeek(argNow);

  // Fecha elegida por Tadeo (opcional) — si viene en formato yyyy-MM-dd, convertir
  var fechaPagoStr = fechaHoy;
  if (fechaStr && fechaStr.indexOf('-') >= 0) {
    var fp = fechaStr.split('-');
    if (fp.length === 3) fechaPagoStr = fp[2] + '/' + fp[1] + '/' + fp[0];
  }

  // Recorrer Red y marcar FIFO
  var restante = monto;
  var pedidosMarcados = 0;

  function _marcarEnHoja(nombre) {
    if (restante <= 0) return;
    var sh = SS.getSheetByName(nombre);
    if (!sh || sh.getLastRow() <= 1) return;
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var idxVendedor=-1, idxTotal=-1, idxFecha=-1, idxEstPago=-1, idxFechaPago=-1, idxFormaPago=-1;
    for (var h = 0; h < headers.length; h++) {
      var nm = String(headers[h]).trim();
      if (nm === 'Vendedor') idxVendedor = h;
      else if (nm === 'Total ($)' || nm === 'Total') idxTotal = h;
      else if (nm === 'Fecha') idxFecha = h;
      else if (nm === 'Estado Pago a Maleu') idxEstPago = h;
      else if (nm === 'Fecha Pago a Maleu') idxFechaPago = h;
      else if (nm === 'Forma Pago a Maleu') idxFormaPago = h;
    }
    if (idxVendedor < 0 || idxEstPago < 0) return;
    // Ordenar filas candidatas por fecha ASC (FIFO)
    var candidatos = [];
    for (var r = 1; r < data.length; r++) {
      var v = String(data[r][idxVendedor] || '').trim();
      if (v !== vendedor) continue;
      var ep = String(data[r][idxEstPago] || '').trim();
      if (ep === 'Sí' || ep === 'Si') continue;
      var tot = Number(data[r][idxTotal]) || 0;
      if (tot <= 0) continue;
      var fRow = data[r][idxFecha];
      var dFila = null;
      if (fRow instanceof Date) dFila = fRow;
      else if (fRow) {
        var fp2 = String(fRow).split('/');
        if (fp2.length >= 3) {
          var yy = Number(fp2[2]); if (yy < 100) yy += 2000;
          dFila = new Date(yy, Number(fp2[1])-1, Number(fp2[0]));
        }
      }
      candidatos.push({ rowIdx: r + 1, total: tot, fecha: dFila || new Date(0) });
    }
    candidatos.sort(function(a, b) { return a.fecha - b.fecha; });
    candidatos.forEach(function(c) {
      if (restante <= 0.01) return;
      var deudaP = c.total * (1 - comPct/100);
      sh.getRange(c.rowIdx, idxEstPago + 1).setValue('Sí');
      if (idxFechaPago >= 0) sh.getRange(c.rowIdx, idxFechaPago + 1).setValue(fechaPagoStr);
      if (idxFormaPago >= 0) sh.getRange(c.rowIdx, idxFormaPago + 1).setValue(metodo);
      restante -= deudaP;
      pedidosMarcados++;
    });
  }

  _marcarEnHoja('Red');

  // Registrar Ingreso en caja
  var shIng = SS.getSheetByName('Ingresos');
  if (!shIng) {
    shIng = SS.insertSheet('Ingresos');
    shIng.getRange(1, 1, 1, 8).setValues([['Fecha','Mes','Semana','Categoría','Concepto / Detalle','Método','Monto','Nota']])
      .setBackground(BROWN).setFontColor('#FFFFFF').setFontWeight('bold');
    shIng.setFrozenRows(1);
  }
  var concepto = 'Liquidación ' + vendedor;
  // Headers reales de Ingresos: [Fecha, Semana, Mes, Categoría, Concepto, Método, Monto, Notas]
  shIng.appendRow([fechaPagoStr, semana, mesNombre, 'Liquidación Red', concepto, metodo, monto, notas]);

  return ContentService.createTextOutput(JSON.stringify({
    ok: true, pedidosMarcados: pedidosMarcados, sobrante: Math.max(0, restante),
    mensaje: pedidosMarcados + ' pedido(s) marcados como pagados. Sobrante: $' + Math.round(restante)
  })).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════════════════
//  RESUMEN SEMANAL — endpoint para el Panel y el cron de n8n (domingo 21hs)
//  Devuelve JSON con todo lo que rinde una semana de Maleu:
//  ventas por canal, finanzas (CMV, gastos, neto), clientes (con detección
//  de "caras nuevas" usando matching por teléfono/domicilio), top productos,
//  comparativa vs semana anterior y meta semanal.
// ════════════════════════════════════════════════════════════════════════════

function _doGetResumenSemanal(e) {
  var p = (e && e.parameter) || {};
  var rango = _resumenParseRango(p.desde, p.hasta, p.semanaOffset);
  var anterior = _resumenParseRango(
    Utilities.formatDate(_resumenAddDays(rango.desde, -7), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd'),
    Utilities.formatDate(_resumenAddDays(rango.hasta, -7), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd'),
    null
  );

  // 1. Leer hojas operativas (semana actual + semana anterior + año completo para "cara nueva")
  var canales = _resumenLeerCanales(rango, anterior);

  // 2. Detectar caras nuevas
  var clientesInfo = _resumenAnalizarClientes(canales, rango);

  // 3. Productos top
  var productosInfo = _resumenAnalizarProductos(canales.semana);

  // 4. Finanzas (devengado por fecha de pedido + caja por fecha de cobro)
  var finanzas = _resumenFinanzas(canales.semana, rango, canales.cobradosSemana);

  // 5. Comparativa
  var comparativa = _resumenComparativa(canales.semana, canales.semanaAnterior);

  // 6. Stock crítico
  var stockCritico = _resumenStockCritico();

  // 7. Meta
  var metaSemanal = _resumenMeta(canales, rango);

  var resp = {
    ok: true,
    generadoEn: Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm'),
    periodo: rango.label,
    desde: Utilities.formatDate(rango.desde, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy'),
    hasta: Utilities.formatDate(rango.hasta, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy'),
    semanaNum: rango.semana,
    ventas: _resumenVentas(canales.semana),
    finanzas: finanzas,
    clientes: clientesInfo,
    productos: productosInfo,
    stockCritico: stockCritico,
    comparativa: comparativa,
    meta: metaSemanal,
    cumpleHoy: (function(){ try { return _crmCumpleHoy(); } catch(_e){ return []; } })()
  };

  return ContentService.createTextOutput(JSON.stringify(resp))
    .setMimeType(ContentService.MimeType.JSON);
}

// ───────── helpers de fecha / rango ─────────

function _resumenParseRango(desdeStr, hastaStr, semanaOffsetStr) {
  var TZ = 'America/Argentina/Buenos_Aires';
  var hoy = new Date(new Date().toLocaleString('en-US', {timeZone: TZ}));
  var desde, hasta;
  if (desdeStr && hastaStr) {
    desde = _resumenFromYMD(desdeStr);
    hasta = _resumenFromYMD(hastaStr);
  } else {
    var offset = parseInt(semanaOffsetStr || '0', 10) || 0;
    var dow = hoy.getDay();
    var diasDesdeLunes = (dow === 0 ? 6 : dow - 1);
    var lunes = _resumenAddDays(hoy, -diasDesdeLunes + (offset * 7));
    desde = lunes;
    hasta = _resumenAddDays(lunes, 6);
  }
  desde.setHours(0,0,0,0); hasta.setHours(23,59,59,999);
  var sem = _resumenISOWeek(desde);
  var label = 'Semana ' + sem + ' · ' +
    Utilities.formatDate(desde, TZ, 'dd/MM') + ' al ' +
    Utilities.formatDate(hasta, TZ, 'dd/MM/yyyy');
  return { desde: desde, hasta: hasta, semana: sem, label: label };
}
function _resumenFromYMD(s) {
  var m = String(s).match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!m) return new Date();
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}
function _resumenAddDays(d, n) {
  var x = new Date(d.getTime());
  x.setDate(x.getDate() + n);
  return x;
}
function _resumenISOWeek(d) {
  var t = new Date(d.getTime());
  t.setHours(0,0,0,0);
  t.setDate(t.getDate() + 4 - (t.getDay()||7));
  var year = new Date(t.getFullYear(),0,1);
  return Math.ceil((((t-year)/86400000)+1)/7);
}
function _resumenStripTildes(s) {
  return String(s||'').toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g,'')
    .replace(/\s+/g,' ').trim();
}
function _resumenNormTel(t) {
  var s = String(t||'').replace(/\D+/g,'');
  if (!s) return '';
  if (s.length > 10) s = s.slice(-10);
  return s;
}
function _resumenClienteKey(canal, fila) {
  var tel = _resumenNormTel(fila.telefono);
  if (canal === 'Clubes') {
    var c = _resumenStripTildes(fila.club || '');
    var g = _resumenStripTildes(fila.grupo || '');
    if (c) return 'club::' + c + '|' + g;
  }
  if (tel) return 'tel::' + tel;
  var dom = _resumenStripTildes((fila.subBarrio || fila.barrio || '') + '|' + (fila.lote || ''));
  if (dom && dom !== '|') return 'dom::' + dom;
  return 'nom::' + _resumenStripTildes(fila.cliente || '');
}

// ───────── lectura de hojas operativas ─────────

function _resumenCfgCanales() {
  // 0-based para acceso al array (Home cFecha:3 = índice 3 = col D)
  // tartaStart / tartaEnd (0-based): cols al final donde están las tartas TP/TJyQ/TCa/TV.
  // Home BE-BH = idx 56-59. Pilar BH-BK = idx 59-62. Red BF-BI = idx 57-60.
  return [
    { canal:'Home', sheet:'Home',
      cFecha:3, cCliente:7, cOrigen:8, cEstEnt:10, cEstPago:12,
      cTotal:16, cEf:17, cTr:18, cPropEf:19, cPropTr:20, cFacturado:21,
      cCosto:41, cBarrio:43, cSubBarrio:44, cLote:45, cTel:46,
      cFechaCobro:54, prodStart:22, prodEnd:40, tartaStart:56, tartaEnd:59 },
    { canal:'Pilar', sheet:'Pilar',
      cFecha:3, cCliente:7, cOrigen:8, cEstEnt:10, cEstPago:12,
      cTotal:16, cEf:17, cTr:18, cPropEf:19, cPropTr:20, cFacturado:21,
      cCosto:45, cBarrio:47, cSubBarrio:47, cLote:48, cTel:49,
      cFechaCobro:57, prodStart:22, prodEnd:44, tartaStart:59, tartaEnd:62 },
    { canal:'Clubes', sheet:'Clubes',
      cFecha:3, cCliente:7, cClub:8, cDeporte:9, cGrupo:10, cOrigen:11, cEstEnt:13, cEstPago:15,
      cTotal:16, cEf:18, cTr:19, cPropEf:20, cPropTr:21, cFacturado:22,
      cCosto:31, cTel:33, cFechaCobro:34, prodStart:23, prodEnd:30, tartaStart:37, tartaEnd:40 },
    { canal:'Red', sheet:'Red',
      cFecha:3, cVendedor:7, cCliente:8, cOrigen:9, cEstEnt:11, cEstPago:13,
      cTotal:14, cEf:16, cTr:17, cPropEf:18, cPropTr:19, cFacturado:20,
      cCosto:44, cBarrio:49, cLote:50, cTel:51,
      cFechaCobro:54, prodStart:21, prodEnd:43, tartaStart:57, tartaEnd:60 }
  ];
}

function _resumenLeerCanales(rango, anterior) {
  var cfgs = _resumenCfgCanales();
  var sem = [], semAnt = [], anio = [], cobradosSem = [];
  cfgs.forEach(function(cfg){
    var sh = SS.getSheetByName(cfg.sheet);
    if (!sh || sh.getLastRow() < 2) return;
    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var f = row[cfg.cFecha];
      if (!(f instanceof Date)) {
        var m = String(f||'').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (m) f = new Date(Number(m[3]), Number(m[2])-1, Number(m[1]));
        else continue;
      }
      var fila = _resumenFilaToObj(row, cfg);
      fila.fecha = f;
      fila.canal = cfg.canal;
      fila.row = r + 1;

      if (f.getFullYear() === rango.desde.getFullYear()) anio.push(fila);
      if (f >= rango.desde && f <= rango.hasta) sem.push(fila);
      if (f >= anterior.desde && f <= anterior.hasta) semAnt.push(fila);

      // Cobrados de la semana (vista caja): fecha relevante = Fecha de Cobro
      // (fallback: fecha de pedido). Coincide con la lógica de la card de Inicio.
      if (fila.estadoEntrega !== 'Cancelado' && fila.estadoPago === 'Cobrado') {
        var fc = fila.fechaCobro, fcD = null;
        if (fc instanceof Date) fcD = fc;
        else if (fc) {
          var mm = String(fc).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
          if (mm) fcD = new Date(Number(mm[3]), Number(mm[2])-1, Number(mm[1]));
        }
        if (!fcD) fcD = f;
        if (fcD >= rango.desde && fcD <= rango.hasta) cobradosSem.push(fila);
      }
    }
  });
  return { semana: sem, semanaAnterior: semAnt, anio: anio, cobradosSemana: cobradosSem };
}

function _resumenFilaToObj(row, cfg) {
  function val(c, def) { return c != null ? (row[c] != null ? row[c] : (def == null ? '' : def)) : (def == null ? '' : def); }
  var productos = {};
  for (var c = cfg.prodStart; c <= cfg.prodEnd; c++) {
    var q = Number(row[c]) || 0;
    if (q > 0) productos[c] = q;
  }
  // Tartas (cols al final, no consecutivas con productos viejos)
  if (cfg.tartaStart != null && cfg.tartaEnd != null) {
    for (var ct = cfg.tartaStart; ct <= cfg.tartaEnd; ct++) {
      var qt = Number(row[ct]) || 0;
      if (qt > 0) productos[ct] = qt;
    }
  }
  return {
    cliente: String(val(cfg.cCliente,'')).trim(),
    vendedor: String(val(cfg.cVendedor,'')).trim(),
    club: String(val(cfg.cClub,'')).trim(),
    deporte: String(val(cfg.cDeporte,'')).trim(),
    grupo: String(val(cfg.cGrupo,'')).trim(),
    origen: String(val(cfg.cOrigen,'')).trim(),
    estadoEntrega: String(val(cfg.cEstEnt,'')).trim(),
    estadoPago: String(val(cfg.cEstPago,'')).trim(),
    total: Number(val(cfg.cTotal,0)) || 0,
    efectivo: Number(val(cfg.cEf,0)) || 0,
    transferencia: Number(val(cfg.cTr,0)) || 0,
    propEf: Number(val(cfg.cPropEf,0)) || 0,
    propTr: Number(val(cfg.cPropTr,0)) || 0,
    facturado: Number(val(cfg.cFacturado,0)) || 0,
    costo: Number(val(cfg.cCosto,0)) || 0,
    barrio: String(val(cfg.cBarrio,'')).trim(),
    subBarrio: String(val(cfg.cSubBarrio,'')).trim(),
    lote: String(val(cfg.cLote,'')).trim(),
    telefono: String(val(cfg.cTel,'')).trim(),
    fechaCobro: val(cfg.cFechaCobro, ''),
    productos: productos,
    cfg: cfg
  };
}

// ───────── ventas ─────────

function _resumenVentas(filas) {
  var canales = {};
  var porDia = {};
  var totalPed = 0, totalEnt = 0, totalPend = 0, totalCanc = 0, totalFact = 0, totalCob = 0;
  var pedVD = 0, entVD = 0, pendVD = 0, cancVD = 0, factVD = 0;

  filas.forEach(function(f){
    var fact = f.facturado || f.total || 0;
    var est = f.estadoEntrega;
    var isCanc = est === 'Cancelado';
    var isEnt  = est === 'Entregado';
    var isPend = !isCanc && !isEnt;
    var cob    = (!isCanc && f.estadoPago === 'Cobrado') ? fact : 0;
    var esVD   = (f.canal !== 'Red');

    if (!canales[f.canal]) canales[f.canal] = { canal:f.canal, pedidos:0, entregados:0, pendientes:0, cancelados:0, facturado:0, cobrado:0 };
    var c = canales[f.canal];
    c.pedidos++;
    if (isEnt)  c.entregados++;
    if (isPend) c.pendientes++;
    if (isCanc) c.cancelados++;

    if (isCanc) { totalCanc++; if (esVD) cancVD++; return; }

    c.facturado += fact;
    c.cobrado   += cob;

    var diaKey = Utilities.formatDate(f.fecha, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
    if (!porDia[diaKey]) porDia[diaKey] = { fecha: diaKey, pedidos:0, facturado:0 };
    porDia[diaKey].pedidos++;
    porDia[diaKey].facturado += fact;

    totalPed++; if (isEnt) totalEnt++; if (isPend) totalPend++;
    totalFact += fact; totalCob += cob;
    if (esVD) {
      pedVD++; if (isEnt) entVD++; if (isPend) pendVD++;
      factVD += fact;
    }
  });

  var canalesArr = Object.keys(canales).map(function(k){
    var c = canales[k];
    c.ticketProm = c.pedidos ? Math.round(c.facturado / c.pedidos) : 0;
    c.facturado = Math.round(c.facturado);
    c.cobrado = Math.round(c.cobrado);
    return c;
  });
  canalesArr.sort(function(a,b){ return b.facturado - a.facturado; });

  var diasArr = Object.keys(porDia).sort().map(function(k){
    porDia[k].facturado = Math.round(porDia[k].facturado);
    return porDia[k];
  });
  var mejorDia = diasArr.slice().sort(function(a,b){ return b.facturado - a.facturado; })[0] || null;
  if (mejorDia) {
    var DIAS = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
    var pp = mejorDia.fecha.split('-');
    var dd = new Date(Number(pp[0]), Number(pp[1]) - 1, Number(pp[2]));
    mejorDia.diaNombre = DIAS[dd.getDay()];
  }

  return {
    pedidos: totalPed,
    pedidosVD: pedVD,
    entregados: totalEnt,
    entregadosVD: entVD,
    pendientes: totalPend,
    pendientesVD: pendVD,
    cancelados: totalCanc,
    canceladosVD: cancVD,
    facturado: Math.round(totalFact),
    facturadoVD: Math.round(factVD),
    cobrado: Math.round(totalCob),
    sinCobrar: Math.round(totalFact - totalCob),
    ticketProm: totalPed ? Math.round(totalFact / totalPed) : 0,
    ticketPromVD: pedVD ? Math.round(factVD / pedVD) : 0,
    porCanal: canalesArr,
    porDia: diasArr,
    mejorDia: mejorDia
  };
}

// ───────── finanzas ─────────

function _resumenFinanzas(filas, rango, cobradosSem) {
  // ─── DEVENGADO (Económico): pedidos generados en la semana ───
  var fact = 0, cmv = 0, propinas = 0, cobradoDev = 0, efDev = 0, trDev = 0;
  var pendientesCobro = [];
  filas.forEach(function(f){
    if (f.estadoEntrega === 'Cancelado') return;
    var monto = f.facturado || f.total || 0;
    fact += monto;
    cmv  += (f.costo || 0);
    propinas += (f.propEf || 0) + (f.propTr || 0);
    if (f.estadoPago === 'Cobrado') {
      cobradoDev += monto;
      efDev += (f.efectivo || 0) + (f.propEf || 0);
      trDev += (f.transferencia || 0) + (f.propTr || 0);
    } else if (monto > 0) {
      pendientesCobro.push({
        cliente: (f.cliente || f.club || '(sin nombre)') + (f.canal === 'Red' && f.vendedor ? ' (Red: ' + f.vendedor + ')' : ''),
        canal: f.canal,
        fecha: Utilities.formatDate(f.fecha, 'America/Argentina/Buenos_Aires', 'dd/MM'),
        estado: f.estadoEntrega,
        monto: Math.round(monto)
      });
    }
  });
  pendientesCobro.sort(function(a,b){ return b.monto - a.monto; });

  var gastos = _resumenLeerEgresos(rango);
  var ingresos = _resumenLeerIngresos(rango);
  var margenBruto = fact - cmv;
  var resultadoNeto = margenBruto - gastos.total + ingresos.total;

  // ─── PERCIBIDO (Caja): cobros con fecha de cobro EN la semana ───
  var cobradoCaja = 0, efCaja = 0, trCaja = 0;
  (cobradosSem || []).forEach(function(f){
    var monto = f.facturado || f.total || 0;
    cobradoCaja += monto;
    efCaja += (f.efectivo || 0) + (f.propEf || 0);
    trCaja += (f.transferencia || 0) + (f.propTr || 0);
  });
  var resultadoCaja = cobradoCaja + ingresos.total - gastos.total;

  return {
    // Devengado (Económico)
    facturado: Math.round(fact),
    cmv: Math.round(cmv),
    margenBruto: Math.round(margenBruto),
    margenPct: fact ? Math.round((margenBruto / fact) * 100) : 0,
    gastos: gastos,
    ingresosNoVenta: ingresos,
    resultadoNeto: Math.round(resultadoNeto),
    // Caja (Percibido)
    caja: {
      ventasCobradas: (cobradosSem || []).length,
      cobrado: Math.round(cobradoCaja),
      cobradoEf: Math.round(efCaja),
      cobradoTr: Math.round(trCaja),
      gastosPagados: gastos.total,
      ingresosNoVenta: ingresos.total,
      resultadoCaja: Math.round(resultadoCaja)
    },
    cobranza: {
      cobrado: Math.round(cobradoDev),
      sinCobrar: Math.round(fact - cobradoDev),
      efectivo: Math.round(efDev),
      transferencia: Math.round(trDev),
      propinas: Math.round(propinas),
      pendientes: pendientesCobro
    }
  };
}

function _resumenLeerMovimientos(sheetName, rango) {
  var sh = SS.getSheetByName(sheetName);
  var porCat = {}, total = 0, detalle = [];
  if (sh && sh.getLastRow() > 1) {
    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var f = data[r][0];
      if (!(f instanceof Date)) {
        var m = String(f||'').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (m) f = new Date(Number(m[3]), Number(m[2])-1, Number(m[1]));
        else continue;
      }
      if (f < rango.desde || f > rango.hasta) continue;
      var cat = String(data[r][3] || 'Otro').trim() || 'Otro';
      var con = String(data[r][4] || '').trim();
      var met = String(data[r][5] || '').trim();
      var monto = Number(data[r][6]) || 0;
      var notas = String(data[r][7] || '').trim();
      if (!porCat[cat]) porCat[cat] = 0;
      porCat[cat] += monto;
      total += monto;
      detalle.push({
        fecha: Utilities.formatDate(f, 'America/Argentina/Buenos_Aires', 'dd/MM'),
        ts: f.getTime(),
        categoria: cat, concepto: con, metodo: met,
        monto: Math.round(monto), notas: notas
      });
    }
  }
  detalle.sort(function(a,b){ return a.ts - b.ts; });
  detalle.forEach(function(d){ delete d.ts; });
  var arr = Object.keys(porCat).map(function(k){ return { categoria:k, monto: Math.round(porCat[k]) }; });
  arr.sort(function(a,b){ return b.monto - a.monto; });
  return { total: Math.round(total), porCategoria: arr, detalle: detalle };
}
function _resumenLeerEgresos(rango)  { return _resumenLeerMovimientos('Egresos', rango); }
function _resumenLeerIngresos(rango) { return _resumenLeerMovimientos('Ingresos', rango); }

// ───────── clientes (con detección de "cara nueva") ─────────

function _resumenAnalizarClientes(canales, rango) {
  // 1. Primer pedido de cada cliente en TODO el año
  var primerPedido = {};
  canales.anio.forEach(function(f){
    if (f.estadoEntrega === 'Cancelado') return;
    var k = _resumenClienteKey(f.canal, f);
    if (!primerPedido[k] || f.fecha < primerPedido[k].fecha) {
      primerPedido[k] = { fecha: f.fecha, canal: f.canal, row: f.row };
    }
  });

  // 1b. Fecha del último pedido PREVIO a la semana (para separar recompra activa de reactivados)
  var ultimoPrevio = {};
  canales.anio.forEach(function(f){
    if (f.estadoEntrega === 'Cancelado') return;
    if (f.fecha >= rango.desde) return; // solo pedidos anteriores a la semana analizada
    var k = _resumenClienteKey(f.canal, f);
    if (!ultimoPrevio[k] || f.fecha > ultimoPrevio[k]) ultimoPrevio[k] = f.fecha;
  });
  var umbralActivo = _resumenAddDays(rango.desde, -28); // compró en las últimas 4 semanas → activo

  // 2. Recorrer la semana
  var unicos = {};
  var nuevos = [];
  var nuevosKeys = {};

  canales.semana.forEach(function(f){
    if (f.estadoEntrega === 'Cancelado') return;
    var k = _resumenClienteKey(f.canal, f);
    var fact = f.facturado || f.total || 0;
    if (!unicos[k]) unicos[k] = { key:k, nombre:f.cliente || f.club || '(sin nombre)', canal:f.canal, telefono:f.telefono, pedidos:0, monto:0 };
    unicos[k].pedidos++;
    unicos[k].monto += fact;

    var pp = primerPedido[k];
    if (pp && pp.fecha >= rango.desde && pp.fecha <= rango.hasta && !nuevosKeys[k]) {
      nuevosKeys[k] = true;
      var nombre = f.cliente || f.club || '(sin nombre)';
      var ubic = '';
      if (f.canal === 'Clubes') ubic = (f.club || '') + (f.grupo ? ' · ' + f.grupo : '');
      else ubic = (f.subBarrio || f.barrio || '') + (f.lote ? ' · ' + f.lote : '');
      nuevos.push({
        nombre: nombre,
        canal: f.canal,
        telefono: _resumenNormTel(f.telefono),
        ubicacion: ubic.trim(),
        fecha: Utilities.formatDate(f.fecha, 'America/Argentina/Buenos_Aires', 'dd/MM'),
        monto: Math.round(fact)
      });
    }
  });

  var unicosArr = Object.keys(unicos).map(function(k){
    var u = unicos[k];
    u.monto = Math.round(u.monto);
    return u;
  });
  // Top excluye Red: en Red el cliente final no es de Maleu sino del vendedor
  var top = unicosArr.filter(function(u){ return u.canal !== 'Red'; })
                     .sort(function(a,b){ return b.monto - a.monto; })
                     .slice(0, 5);

  // Separar a los recurrentes en "recompra activa" (compró en las últimas 4 semanas)
  // vs "reactivados" (volvió tras 5+ semanas dormido). nuevos + activa + react = unicos.
  // Los reactivados se listan con nombre/tel/cuánto hacía que no compraban → accionable.
  var recompraActiva = 0, reactivados = 0;
  var reactivadosList = [];
  unicosArr.forEach(function(u){
    if (nuevosKeys[u.key]) return; // ya contado como nuevo
    var ult = ultimoPrevio[u.key];
    if (ult && ult >= umbralActivo) { recompraActiva++; return; }
    reactivados++;
    var semanas = ult ? Math.round((rango.desde.getTime() - ult.getTime()) / (7 * 86400000)) : null;
    reactivadosList.push({
      nombre: u.nombre,
      canal: u.canal,
      telefono: _resumenNormTel(u.telefono),
      monto: u.monto,
      semanasDormido: semanas,
      ultima: ult ? Utilities.formatDate(ult, 'America/Argentina/Buenos_Aires', 'dd/MM') : ''
    });
  });
  reactivadosList.sort(function(a, b){ return b.monto - a.monto; });

  return {
    unicos: unicosArr.length,
    nuevosCount: nuevos.length,
    nuevos: nuevos,
    recurrentes: unicosArr.length - nuevos.length,
    recompraActiva: recompraActiva,
    reactivados: reactivados,
    reactivadosList: reactivadosList,
    top: top
  };
}

// ───────── productos ─────────

function _resumenAnalizarProductos(filas) {
  var map = {
    'Home':   typeof HOME_COL_TO_ABBR   !== 'undefined' ? HOME_COL_TO_ABBR   : {},
    'Pilar':  typeof PILAR_COL_TO_ABBR  !== 'undefined' ? PILAR_COL_TO_ABBR  : {},
    'Clubes': typeof CLUBES_COL_TO_ABBR !== 'undefined' ? CLUBES_COL_TO_ABBR : {},
    'Red':    typeof RED_COL_TO_ABBR    !== 'undefined' ? RED_COL_TO_ABBR    : {}
  };
  var catalogo = _resumenCatalogoProductos();

  var unidades = {}, facturadoProd = {};
  filas.forEach(function(f){
    if (f.estadoEntrega === 'Cancelado') return;
    var canalMap = map[f.canal] || {};
    Object.keys(f.productos).forEach(function(colStr){
      var col = Number(colStr);
      var abr = canalMap[col + 1];
      if (!abr) return;
      var q = f.productos[col];
      unidades[abr] = (unidades[abr] || 0) + q;
      var precio = (catalogo[abr] && catalogo[abr].precio) || 0;
      facturadoProd[abr] = (facturadoProd[abr] || 0) + (q * precio);
    });
  });

  function arrFromMap(m) {
    return Object.keys(m).map(function(abr){
      return {
        abr: abr,
        nombre: (catalogo[abr] && catalogo[abr].nombre) || abr,
        unidades: unidades[abr] || 0,
        facturado: Math.round(facturadoProd[abr] || 0)
      };
    });
  }
  var topUnidades  = arrFromMap(unidades).sort(function(a,b){ return b.unidades - a.unidades; }).slice(0, 5);
  var topFacturado = arrFromMap(facturadoProd).sort(function(a,b){ return b.facturado - a.facturado; }).slice(0, 5);
  return { topUnidades: topUnidades, topFacturado: topFacturado };
}

function _resumenCatalogoProductos() {
  var sh = SS.getSheetByName('Productos');
  var map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  var data = sh.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    var abr = String(data[r][2] || '').trim();
    if (!abr) continue;
    map[abr] = {
      nombre: String(data[r][1] || '').trim(),
      stock: Number(data[r][5]) || 0,
      precio: Number(data[r][8]) || 0,
      costo: Number(data[r][9]) || 0
    };
  }
  return map;
}

function _resumenStockCritico() {
  var cat = _resumenCatalogoProductos();
  var arr = [];
  Object.keys(cat).forEach(function(abr){
    var p = cat[abr];
    if (p.stock <= 3) arr.push({ abr:abr, nombre:p.nombre, stock:p.stock });
  });
  arr.sort(function(a,b){ return a.stock - b.stock; });
  return arr;
}

// ───────── comparativa ─────────

function _resumenComparativa(filasSem, filasAnt) {
  function totales(filas) {
    var p = 0, f = 0, c = 0;
    filas.forEach(function(x){
      if (x.estadoEntrega === 'Cancelado') return;
      p++;
      f += (x.facturado || x.total || 0);
      if (x.estadoPago === 'Cobrado') c += (x.facturado || x.total || 0);
    });
    return { pedidos:p, facturado:Math.round(f), cobrado:Math.round(c) };
  }
  var s = totales(filasSem), a = totales(filasAnt);
  function delta(now, prev) {
    if (prev === 0) return now > 0 ? 100 : 0;
    return Math.round(((now - prev) / prev) * 100);
  }
  return {
    semanaAnterior: a,
    deltas: {
      pedidos: delta(s.pedidos, a.pedidos),
      facturado: delta(s.facturado, a.facturado),
      cobrado: delta(s.cobrado, a.cobrado)
    }
  };
}

// ───────── meta semanal ─────────

// Lee un parámetro numérico de Config_Maleu (clave en col A, valor en col B).
// Robusto: si falta la hoja, la fila o el valor, devuelve el default. Acepta
// formatos "$1.500.000" / "0,08" / "50".
function _resumenCfgNum(nombre, def) {
  try {
    var sh = SS.getSheetByName('Config_Maleu');
    if (!sh || sh.getLastRow() < 2) return def;
    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]).trim() === nombre) {
        var v = String(data[r][1]).replace(/[^0-9.,\-]/g, '').replace(/\./g, '').replace(',', '.');
        var n = parseFloat(v);
        return isNaN(n) ? def : n;
      }
    }
  } catch (e) {}
  return def;
}

// Meta = pedidos entregados en Estancias del Pilar (el corazón del negocio).
// Objetivo: 50/semana sostenido todo el año (editable en Config_Maleu).
// Muestra el dato de la semana + el promedio semanal del mes en curso (mes a mes).
function _resumenMeta(canales, rango) {
  var objetivo = _resumenCfgNum('META_PED_SEM_ESTANCIAS', 50);
  function esEst(f) {
    return f.canal === 'Home'
      && f.estadoEntrega === 'Entregado'
      && String(f.barrio || '').trim().toLowerCase() === 'estancias del pilar';
  }
  // Esta semana
  var semana = 0;
  canales.semana.forEach(function(f){ if (esEst(f)) semana++; });

  // Mes del informe (mes del lunes de la semana) y promedio semanal hasta el cierre del informe
  var MS = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var mesNum = rango.desde.getMonth(), anio = rango.desde.getFullYear();
  var finMes = new Date(anio, mesNum + 1, 0);
  var tope = (rango.hasta < finMes) ? rango.hasta : finMes; // no pasar de fin de mes
  var semanasMes = Math.max(1, Math.round(tope.getDate() / 7));
  var pedMes = 0;
  canales.anio.forEach(function(f){
    if (!esEst(f)) return;
    if (f.fecha.getMonth() === mesNum && f.fecha.getFullYear() === anio && f.fecha <= tope) pedMes++;
  });
  var promedioMes = Math.round((pedMes / semanasMes) * 10) / 10;

  return {
    objetivo: objetivo,
    alcanzado: semana,
    pct: objetivo > 0 ? Math.round((semana / objetivo) * 100) : 0,
    detalle: 'Pedidos entregados en Estancias del Pilar',
    mes: MS[mesNum],
    pedidosMes: pedMes,
    semanasMes: semanasMes,
    promedioMes: promedioMes
  };
}

// ───────── helper para test desde el editor ─────────

function testResumenSemanal() {
  var fakeE = { parameter: { semanaOffset: '-1' } };
  var resp = _doGetResumenSemanal(fakeE);
  Logger.log(resp.getContent());
}

// ════════════════════════════════════════════════════════════
//  Sync WATI: actualiza tipo_de_contacto / zona / barrio_privado
//  cuando entra un pedido. Silencioso (no rompe el pedido).
// ════════════════════════════════════════════════════════════

const WATI_URL_   = 'https://live-mt-server.wati.io/1034656';
const WATI_TOKEN_ = 'wati_6cac1b8c-07cc-4946-b954-5f52df8ba948.iRUrSg_H28yY_zWU3jyMYFu96ErdgwhsnhNA-1_yHN5simg3-rUejn_ROEAGRhIOp2ulVLp4t-7g5VCyD2mMwXqqWGYn0_SahlRTLVoPczz3xwIH8bXV5NkyJob-dPKn';

function _normalizePhoneWATI_(phone) {
  var s = String(phone || '').replace(/\D/g, '');
  if (!s) return '';
  if (s.indexOf('549') === 0) return s;
  if (s.indexOf('54') === 0 && s.length === 12) return s.slice(0,2) + '9' + s.slice(2);
  if (s.indexOf('15') === 0) return '5491' + s.slice(2);
  if (s.length === 10) return '549' + s;
  if (s.length === 8)  return '54911' + s;
  return s;
}

function _zonaFromSubBarrio_(sub) {
  var sb = String(sub || '').trim().toLowerCase();
  if (!sb) return null;
  if (sb.indexOf('rio') >= 0 || sb.indexOf('río') >= 0) return 'Estancias del Rio';
  if (sb.indexOf('alcanfor') >= 0) return 'Los Alcanfores';
  return 'Estancias del Pilar';
}

function _barrioCanon_(sub) {
  var sb = String(sub || '').trim();
  if (!sb) return '';
  var key = sb.toLowerCase();
  var MAP = {
    'estancias del rio':'Estancias del Rio',
    'estancias del río':'Estancias del Rio',
    'los alcanfores':'Los Alcanfores',
    'alcanfores':'Los Alcanfores',
    'champagnat alto':'Champagnat Alto',
    'champagnat bajo':'Champagnat Bajo',
    'golf':'Golf',
    'la pionera':'La Pionera',
    'el recuerdo':'El Recuerdo',
    'la paz':'La Paz',
    'argentina 1':'Argentina 1',
    'argentina 2':'Argentina 2',
    'argentina 3':'Argentina 3',
    'argentina 4':'Argentina 4',
    'pilara':'Pilara'
  };
  return MAP[key] || sb;
}

/** Clave normalizada de barrio para comparar (minúsculas, sin acentos, espacios colapsados). */
function _barrioKey(s) {
  s = String(s || '').toLowerCase()
    .replace(/á/g, 'a').replace(/é/g, 'e').replace(/í/g, 'i')
    .replace(/ó/g, 'o').replace(/ú/g, 'u').replace(/ü/g, 'u').replace(/ñ/g, 'n')
    .replace(/\s*\(\d+\)\s*$/, '');   // tolera un "(444)" pegado por el front
  return s.replace(/\s+/g, ' ').trim();
}

/** ¿Es un BARRIO PRIVADO de Home (nivel superior)? Estancias del Pilar/Río + Los Alcanfores.
 *  Los sub-barrios (El Recuerdo, La Paz, Champagnat, Golf…) viven adentro de Estancias del Pilar. */
function _esBarrioPriv(canon) {
  var k = _barrioKey(canon);
  return k === 'estancias del pilar' || k === 'estancias del rio' || k === 'los alcanfores';
}

/** Acumula un barrio en el universo para poblar los dropdowns del filtro. */
function _accBarrio(u, canal, tipo, v) {
  if (!v || v === '-') return;
  var k = canal + '||' + tipo + '||' + v;
  if (!u[k]) u[k] = { canal: canal, tipo: tipo, v: v, n: 0 };
  u[k].n++;
}

/**
 * Sincroniza un contacto a WATI cuando entra un pedido.
 * Setea tipo_de_contacto y, cuando aplica, zona + barrio_privado.
 * Silencioso: si WATI falla, no rompe el pedido (loguea a Log Errores).
 *
 * @param {string} canal — 'Home' | 'Pilar' | 'Clubes' | 'Red'
 * @param {string} telefono — teléfono crudo del cliente
 * @param {string} subBarrio — Sub Barrio (Home) o Barrio Privado (Pilar/Red)
 * @param {string} barrio — Barrio (Home: 'Estancias del Pilar'/'Estancias del Río'/etc)
 */
function _syncWatiContact_(canal, telefono, subBarrio, barrio) {
  try {
    var phone = _normalizePhoneWATI_(telefono);
    if (!phone || phone.length < 11) return;

    var customParams = [];
    customParams.push({ name:'tipo_de_contacto', value: canal });

    if (canal === 'Home') {
      // Home: zona derivada de subBarrio O del barrio si subBarrio vacío
      var ref = subBarrio || barrio || '';
      var zona = _zonaFromSubBarrio_(ref);
      if (zona) customParams.push({ name:'zona', value: zona });
      var bcanon = _barrioCanon_(subBarrio || barrio);
      if (bcanon) customParams.push({ name:'barrio_privado', value: bcanon });
    }
    // Pilar / Clubes / Red: solo tipo_de_contacto

    var resp = UrlFetchApp.fetch(WATI_URL_ + '/api/v1/updateContactAttributes/' + phone, {
      method: 'post',
      contentType: 'application/json-patch+json',
      headers: { Authorization: 'Bearer ' + WATI_TOKEN_ },
      payload: JSON.stringify({ customParams: customParams }),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code !== 200) {
      _logError_('WATI sync ' + canal + ' ' + phone + ' code=' + code + ' body=' + resp.getContentText().slice(0,200));
    }
  } catch (e) {
    _logError_('WATI sync exception ' + canal + ': ' + e.message);
  }
}

function _logError_(msg) {
  try {
    var sh = SS.getSheetByName('Log Errores');
    if (sh) sh.appendRow([new Date(), msg]);
  } catch (e) { /* swallow */ }
}

// ════════════════════════════════════════════════════════════
//  CRM — Base de datos de clientes y productos
//  Endpoints: crmClientes, crmCliente, crmProductos, crmProducto
//  Posts:     crmUpdateClienteMeta, crmMergeClientes
// ════════════════════════════════════════════════════════════

// Normaliza un teléfono argentino para usarlo como clave del CRM.
// Estrategia: extraer dígitos, sacar prefijos (54, 549, 0, 15) y quedarse con
// los últimos 10 dígitos (cód. área + número).
function _crmNormTel(tel) {
  if (!tel) return '';
  var s = String(tel).replace(/\D/g, '');
  if (!s) return '';
  // Sacar prefijos comunes en Argentina
  if (s.length > 10 && s.indexOf('549') === 0) s = s.substring(3);
  else if (s.length > 10 && s.indexOf('54') === 0) s = s.substring(2);
  if (s.length > 10 && s.charAt(0) === '9') s = s.substring(1);
  if (s.length > 10 && s.charAt(0) === '0') s = s.substring(1);
  if (s.length > 10 && s.substring(0, 2) === '15') s = s.substring(2);
  // Tomar últimos 10
  if (s.length > 10) s = s.substring(s.length - 10);
  if (s.length < 8) return '';
  return s;
}

// Normaliza un nombre para fallback de match (lowercase, sin acentos, sin espacios extra).
function _crmNormNombre(nombre) {
  if (!nombre) return '';
  var s = String(nombre).toLowerCase();
  if (s.normalize) s = s.normalize('NFD').replace(/[̀-ͯ]/g, '');
  return s.replace(/\s+/g, ' ').trim();
}

// Convierte una celda de fecha (Date o string) a Date. Devuelve null si no parsea.
function _crmToDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// Configuración de columnas (0-indexed) para cada hoja operativa.
// tartaStart/tartaEnd: rango adicional al final con las 4 tartas (cols separadas).
function _crmHojasConfig() {
  return [
    {
      name: 'Home', cliente: 7, dia: 9, fecha: 3, est: 10, pago: 12, fp: 11,
      total: 21, costo: 41, barrio: 43, subBarrio: 44, domicilio: 45, tel: 46,
      prodStart: 22, prodEnd: 40, fechaCobro: 54,
      tartaStart: 56, tartaEnd: 59
    },
    {
      name: 'Pilar', cliente: 7, dia: 9, fecha: 3, est: 10, pago: 12, fp: 11,
      total: 21, costo: 45, barrio: 47, subBarrio: -1, domicilio: 48, tel: 49,
      prodStart: 22, prodEnd: 44, fechaCobro: 57,
      tartaStart: 59, tartaEnd: 62
    },
    {
      name: 'Clubes', cliente: 7, dia: 12, fecha: 3, est: 13, pago: 15, fp: 14,
      total: 22, costo: 31, barrio: -1, subBarrio: -1, domicilio: -1, tel: 33,
      prodStart: 23, prodEnd: 30, fechaCobro: 34, club: 8, deporte: 9, grupo: 10,
      tartaStart: 37, tartaEnd: 40
    },
    {
      name: 'Red', cliente: 8, dia: 10, fecha: 3, est: 11, pago: 13, fp: 12,
      total: 20, costo: 44, barrio: 49, subBarrio: -1, domicilio: 50, tel: 51,
      prodStart: 21, prodEnd: 43, vendedor: 7,
      tartaStart: 57, tartaEnd: 60
    }
  ];
}

// Lee headers de productos de cada hoja para mapear columna → abreviatura.
function _crmProductHeaders(sh, prodStart, prodEnd, tartaStart, tartaEnd) {
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var map = {};
  for (var i = prodStart; i <= prodEnd && i < headers.length; i++) {
    var abrev = String(headers[i] || '').trim();
    if (abrev) map[i] = abrev;
  }
  // Tartas (rango adicional)
  if (tartaStart != null && tartaEnd != null) {
    for (var j = tartaStart; j <= tartaEnd && j < headers.length; j++) {
      var ab = String(headers[j] || '').trim();
      if (ab) map[j] = ab;
    }
  }
  return map;
}

// Categoría de un producto por su abreviatura (las 6 del catálogo Maleu).
// Para segmentos cross-sell: "compra pizzas pero nunca probó sorrentinos".
function _crmCategoriaDe(abrev) {
  var MAP = {
    SCo:'Sorrentinos', SJyQ:'Sorrentinos', SCa:'Sorrentinos', SQB:'Sorrentinos', SL:'Sorrentinos', SPyP:'Sorrentinos', SE:'Sorrentinos',
    ECaC:'Empanadas', EJyQ:'Empanadas', ECyQ:'Empanadas', EV:'Empanadas',
    PMu:'Pizzas', PMa:'Pizzas', PJyQ:'Pizzas', PCC:'Pizzas', PJyM:'Pizzas', PPM:'Pizzas', PPJyQ:'Pizzas', PPCyQ:'Pizzas',
    TP:'Tartas', TJyQ:'Tartas', TCa:'Tartas', TV:'Tartas',
    RC:'Wraps', RP:'Wraps',
    TG:'Postres', TLC:'Postres', TC:'Postres', F:'Postres'
  };
  return MAP[abrev] || '';
}

// Convierte el valor de la celda Cumpleaños a texto "dd/MM/yyyy". Si Google Sheets
// auto-convirtió "13/05/2002" en una fecha real (Date), lo formatea; si es texto,
// lo devuelve tal cual. Evita el choclo "Mon May 13 2002 ... GMT-0300" y mantiene
// el formato DD/MM que espera la regex del filtro 🎂 Cumple del mes.
function _cumpleToStr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
  return String(v == null ? '' : v).trim();
}

// "DD/MM" de un cumple "13/05" o "13/05/2002". '' si no parsea.
function _cumpleDDMM(cumple) {
  var m = String(cumple || '').match(/^(\d{1,2})\/(\d{1,2})/);
  if (!m) return '';
  return ('0' + m[1]).slice(-2) + '/' + ('0' + m[2]).slice(-2);
}

// Clientes que cumplen años HOY. Para el cartel del Resumen del Panel.
// Barato cuando no hay cumples hoy (solo lee Clientes Meta); solo construye el
// índice de clientes —caro— los días que efectivamente hay un cumpleaños.
function _crmCumpleHoy() {
  var TZ = 'America/Argentina/Buenos_Aires';
  var hoy = Utilities.formatDate(new Date(), TZ, 'dd/MM');
  var meta = _crmGetClientesMeta();
  var telsHoy = Object.keys(meta).filter(function(tel) {
    return _cumpleDDMM(meta[tel].cumple) === hoy;
  });
  if (!telsHoy.length) return [];
  var index = _crmBuildClientesIndex();
  var byTel = {};
  Object.keys(index).forEach(function(k) {
    var c = index[k];
    if (c && c.tel) byTel[c.tel] = c;
  });
  return telsHoy.map(function(tel) {
    var c = byTel[tel];
    var nombre = meta[tel].nombreCanonico || (c ? _crmTopValue(c.nombres) : '') || 'Cliente';
    return { nombre: nombre, tel: tel, telDisplay: (c && c.tel) || tel, cumple: meta[tel].cumple };
  });
}

// Lee meta-data manual de clientes (cumple, notas, alias) desde hoja "Clientes Meta".
// La hoja se crea automáticamente si no existe.
function _crmGetClientesMeta() {
  var sh = SS.getSheetByName('Clientes Meta');
  if (!sh) {
    sh = SS.insertSheet('Clientes Meta');
    var hdr = ['Tel Normalizado', 'Nombre Canónico', 'Cumpleaños (DD/MM)', 'Alias MP', 'Notas', 'Tags', 'Updated', 'Nombres Ocultos', 'Sub Barrio Mapa', 'Lote Mapa', 'Sin Ubicacion', 'Barrio Mapa', 'Canal Mapa', 'Canales Extra', 'Visitante', 'Apodo', 'Residencia'];
    sh.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold').setBackground('#331C1C').setFontColor('#F2E8C7');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, hdr.length, 140);
    return {};
  }
  // Migración suave: agregar headers de columnas nuevas si faltan.
  if (sh.getLastColumn() < 14) sh.getRange(1, 14).setValue('Canales Extra');
  if (sh.getLastColumn() < 15) sh.getRange(1, 15).setValue('Visitante');
  if (sh.getLastColumn() < 16) sh.getRange(1, 16).setValue('Apodo');
  if (sh.getLastColumn() < 17) sh.getRange(1, 17).setValue('Residencia');
  if (sh.getLastRow() <= 1) return {};
  var data = sh.getDataRange().getValues();
  var map = {};
  for (var r = 1; r < data.length; r++) {
    var tel = String(data[r][0] || '').trim();
    if (!tel) continue;
    map[tel] = {
      tel: tel,
      nombreCanonico: String(data[r][1] || '').trim(),
      cumple: _cumpleToStr(data[r][2]),
      aliasMp: String(data[r][3] || '').trim(),
      notas: String(data[r][4] || '').trim(),
      tags: String(data[r][5] || '').trim(),
      updated: data[r][6] || '',
      nombresOcultos: String(data[r][7] || '').split('|').map(function(s){ return s.trim(); }).filter(Boolean),
      subBarrioMapa: String(data[r][8] || '').trim(),
      loteMapa: String(data[r][9] || '').trim(),
      sinUbicacion: String(data[r][10] || '').trim().toUpperCase() === 'TRUE',
      barrioMapa: String(data[r][11] || '').trim(),
      canalMapa: String(data[r][12] || '').trim(),
      canalesExtra: String(data[r][13] || '').split('|').map(function(s){ return s.trim(); }).filter(Boolean),
      apodo: String(data[r][15] || '').trim(),   // sobrenombre para campañas WATI (col Apodo del CSV)
      // Tipo de integrante (col 17). 'vive' = vive en Estancias · 'finde' = va los fines
      // de semana · 'visita' = pide a casa ajena, no reside. Compat: si col 17 vacía pero
      // col 15 (Visitante legacy) = TRUE → 'visita'. `visitante` se deriva de acá.
      residencia: (function(){ var r17=String(data[r][16]||'').trim().toLowerCase(); if(r17==='vive'||r17==='finde'||r17==='visita')return r17; return (String(data[r][14]||'').trim().toUpperCase()==='TRUE')?'visita':''; })(),
      visitante: (function(){ var r17=String(data[r][16]||'').trim().toLowerCase(); if(r17)return r17==='visita'; return String(data[r][14]||'').trim().toUpperCase()==='TRUE'; })(),
      _row: r + 1
    };
  }
  return map;
}

// Lee el mapa de fusiones manuales desde la hoja "Clientes Merge".
// Cada fila: [Alias Key, Canonical Key, Updated]. Devuelve {aliasKey: canonicalKey}.
// No destructivo: las filas operativas no se tocan; el merge se aplica al vuelo
// al construir el índice. Para revertir una fusión, borrar su fila en la hoja.
function _crmGetMergeMap() {
  var sh = SS.getSheetByName('Clientes Merge');
  if (!sh) {
    sh = SS.insertSheet('Clientes Merge');
    sh.getRange(1, 1, 1, 3).setValues([['Alias Key', 'Canonical Key', 'Updated']])
      .setFontWeight('bold').setBackground('#331C1C').setFontColor('#F2E8C7');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 3, 200);
    return {};
  }
  if (sh.getLastRow() <= 1) return {};
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  var map = {};
  for (var r = 0; r < data.length; r++) {
    var a = String(data[r][0] || '').trim();
    var b = String(data[r][1] || '').trim();
    if (a && b && a !== b) map[a] = b;
  }
  return map;
}

// ── Salud de Clientes Home para la semana ISO actual ──────────────────────
// Clasifica a los clientes Home con pedido en la semana actual en
// Nuevos / Recompra / Reactivados, usando la MISMA identidad que el CRM
// (teléfono normalizado + fusiones manuales de "Clientes Merge"). Así dos
// pedidos del mismo cliente cargados con nombres distintos ("Guada" vs
// "Guadalupe Roque Posse") cuentan como uno solo y no inflan "Nuevos".
//   Nuevo      = sin compra Home en semanas previas.
//   Recompra   = última compra Home en las últimas 4 semanas (>= W-4).
//   Reactivado = compró antes, pero hace 5+ semanas.
// Semana de cada pedido = fecha de entrega real (col 49) o, si falta, fecha de
// pedido (col 3). Mismo criterio que usaba el chip del Panel.
// Devuelve { sem, total, nuevos, recompra, react }.
function _saludHomeSemana(argNow) {
  var out = { sem: _isoWeek(argNow), total: 0, nuevos: 0, recompra: 0, react: 0 };
  var sh = SS.getSheetByName('Home');
  if (!sh || sh.getLastRow() <= 1) return out;
  var I = { cli: 7, fPed: 3, estado: 10, fEnt: 49, tel: 46 };
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var mergeMap = _crmGetMergeMap();
  function resolve(key) {
    var k = key, seen = {}, depth = 10;
    while (mergeMap[k] && !seen[k] && depth-- > 0) { seen[k] = true; k = mergeMap[k]; }
    return k;
  }
  var cliWeeks = {}; // canonKey -> { semanaISO: true }
  for (var r = 0; r < data.length; r++) {
    if (String(data[r][I.estado]).trim() === 'Cancelado') continue;
    var nombre = String(data[r][I.cli] || '').trim();
    if (!nombre) continue;
    var d = _crmToDate(data[r][I.fEnt]) || _crmToDate(data[r][I.fPed]);
    if (!d) continue;
    var telNorm = _crmNormTel(data[r][I.tel]);
    var key = resolve(telNorm || ('NOMBRE:' + _crmNormNombre(nombre)));
    var w = _isoWeek(d);
    if (!cliWeeks[key]) cliWeeks[key] = {};
    cliWeeks[key][w] = true;
  }
  var W = out.sem;
  Object.keys(cliWeeks).forEach(function(key) {
    if (!cliWeeks[key][W]) return;
    out.total++;
    var prior = Object.keys(cliWeeks[key]).map(Number).filter(function(x) { return x < W; });
    if (!prior.length) { out.nuevos++; return; }
    var last = Math.max.apply(null, prior);
    if (last >= W - 4) out.recompra++; else out.react++;
  });
  return out;
}

// Absorbe el bucket `src` dentro de `dst` (suma contadores, concatena pedidos,
// combina totales y fechas). Usado al aplicar fusiones manuales.
function _crmFoldBucket(dst, src) {
  function addCounts(d, s) { Object.keys(s || {}).forEach(function(k) { d[k] = (d[k] || 0) + s[k]; }); }
  addCounts(dst.nombres, src.nombres);
  addCounts(dst.telefonos, src.telefonos);
  addCounts(dst.barrios, src.barrios);
  addCounts(dst.subBarrios, src.subBarrios);
  addCounts(dst.domicilios, src.domicilios);
  addCounts(dst.clubes, src.clubes);
  addCounts(dst.canales, src.canales);
  Object.keys(src.productos || {}).forEach(function(ab) {
    if (!dst.productos[ab]) dst.productos[ab] = {cant: 0, monto: 0, ultimo: null};
    dst.productos[ab].cant += src.productos[ab].cant;
    dst.productos[ab].monto += src.productos[ab].monto;
    var su = src.productos[ab].ultimo;
    if (su && (!dst.productos[ab].ultimo || su > dst.productos[ab].ultimo)) dst.productos[ab].ultimo = su;
  });
  dst.pedidos = dst.pedidos.concat(src.pedidos || []);
  dst.totalFacturado += src.totalFacturado;
  dst.totalCobrado += src.totalCobrado;
  dst.deuda += src.deuda;
  dst.countPedidos += src.countPedidos;
  dst.countEntregados += src.countEntregados;
  if (src.firstFecha && (!dst.firstFecha || src.firstFecha < dst.firstFecha)) dst.firstFecha = src.firstFecha;
  if (src.lastFecha && (!dst.lastFecha || src.lastFecha > dst.lastFecha)) dst.lastFecha = src.lastFecha;
  dst.factHome += src.factHome || 0;
  dst.entHome += src.entHome || 0;
  dst.cobrHome += src.cobrHome || 0;
  dst.deudaHome += src.deudaHome || 0;
  if (src.firstHome && (!dst.firstHome || src.firstHome < dst.firstHome)) dst.firstHome = src.firstHome;
  if (src.lastHome && (!dst.lastHome || src.lastHome > dst.lastHome)) dst.lastHome = src.lastHome;
  dst.pendHome = (dst.pendHome || 0) + (src.pendHome || 0);
  if (src.proxHome && (!dst.proxHome || src.proxHome < dst.proxHome)) dst.proxHome = src.proxHome;
  if (!dst.tel && src.tel) { dst.tel = src.tel; dst.telRaw = src.telRaw; }
}

// Construye el agregado completo de clientes a partir de Home + Pilar + Clubes + Red.
// Devuelve un map {telNorm: clienteAgregado}.
function _crmBuildClientesIndex() {
  var hojas = _crmHojasConfig();
  var index = {}; // telNorm -> agregado
  var sinTel = {}; // nombreNorm -> agregado (clientes sin teléfono confiable)

  hojas.forEach(function(cfg) {
    var sh = SS.getSheetByName(cfg.name);
    if (!sh || sh.getLastRow() <= 1) return;
    var prodMap = _crmProductHeaders(sh, cfg.prodStart, cfg.prodEnd, cfg.tartaStart, cfg.tartaEnd);
    var data = sh.getDataRange().getValues();

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var nPed = row[1];
      if (!nPed || String(nPed).trim() === '-') continue;
      var estado = String(row[cfg.est] || '').trim();
      if (estado === 'Cancelado') continue; // No contar cancelados

      var nombre = String(row[cfg.cliente] || '').trim();
      if (!nombre) continue;

      var telRaw = cfg.tel >= 0 ? String(row[cfg.tel] || '').trim() : '';
      var telNorm = _crmNormTel(telRaw);
      var key = telNorm || ('NOMBRE:' + _crmNormNombre(nombre));
      var bucket = telNorm ? index : sinTel;

      if (!bucket[key]) {
        bucket[key] = {
          tel: telNorm,
          telRaw: telRaw,
          nombres: {},      // {nombre: count}
          telefonos: {},    // {telRaw: count}
          barrios: {},      // {barrio: count}
          subBarrios: {},   // {sub: count}
          domicilios: {},   // {dom: count}
          clubes: {},       // {club|deporte|grupo: count}
          canales: {},      // {canal: count}
          pedidos: [],      // [{canal, fecha, fechaSort, total, estado, pago, fp, productos:[{abrev, cant}], dia, club, deporte, grupo, barrio, subBarrio, domicilio, row}]
          productos: {},    // {abrev: {cant, monto}}
          totalFacturado: 0,
          totalCobrado: 0,
          deuda: 0,
          countPedidos: 0,
          countEntregados: 0,
          firstFecha: null,
          lastFecha: null,
          // Solo canal Home entregado en Estancias del Pilar (retail del cockpit de
          // Estancias). El facturado general cruza canales (Clubes, Red…); estos NO.
          factHome: 0, entHome: 0, cobrHome: 0, deudaHome: 0, firstHome: null, lastHome: null,
          // Pedidos Home/Estancias cargados pero AÚN NO entregados (ej. pedido para
          // mañana). No suman facturado ni cuentan como "entregado", pero el cliente
          // NO está "sin compras": tiene una compra en curso. proxHome = fecha de
          // entrega del pendiente más próximo (para mostrar "entrega mañana").
          pendHome: 0, proxHome: null
        };
      }
      var c = bucket[key];

      // Nombres / teléfonos / barrios — frecuencia
      c.nombres[nombre] = (c.nombres[nombre] || 0) + 1;
      if (telRaw) c.telefonos[telRaw] = (c.telefonos[telRaw] || 0) + 1;

      var barrio = cfg.barrio >= 0 ? String(row[cfg.barrio] || '').trim() : '';
      if (barrio) c.barrios[barrio] = (c.barrios[barrio] || 0) + 1;
      var subBarrio = cfg.subBarrio >= 0 ? String(row[cfg.subBarrio] || '').trim() : '';
      if (subBarrio) c.subBarrios[subBarrio] = (c.subBarrios[subBarrio] || 0) + 1;
      var domicilio = cfg.domicilio >= 0 ? String(row[cfg.domicilio] || '').trim() : '';
      if (domicilio) c.domicilios[domicilio] = (c.domicilios[domicilio] || 0) + 1;

      // Club info (Clubes)
      var club = '', deporte = '', grupo = '';
      if (cfg.club !== undefined) {
        club = String(row[cfg.club] || '').trim();
        deporte = String(row[cfg.deporte] || '').trim();
        grupo = String(row[cfg.grupo] || '').trim();
        if (club) {
          var ck = club + ' · ' + deporte + (grupo ? ' · ' + grupo : '');
          c.clubes[ck] = (c.clubes[ck] || 0) + 1;
        }
      }

      // Canal
      c.canales[cfg.name] = (c.canales[cfg.name] || 0) + 1;

      // Fecha
      var fecha = _crmToDate(row[cfg.fecha]);
      var fechaSort = fecha ? fecha.getTime() : 0;
      var fechaStr = fecha ? Utilities.formatDate(fecha, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : '';
      if (fecha) {
        if (!c.firstFecha || fecha < c.firstFecha) c.firstFecha = fecha;
        if (!c.lastFecha || fecha > c.lastFecha) c.lastFecha = fecha;
      }

      var total = Number(row[cfg.total]) || 0;
      var pago = String(row[cfg.pago] || '').trim();
      var fp = String(row[cfg.fp] || '').trim();
      var diaEnt = '';
      var diaRaw = row[cfg.dia];
      var diaDate = _crmToDate(diaRaw);
      if (diaDate) diaEnt = Utilities.formatDate(diaDate, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
      else diaEnt = String(diaRaw || '').trim();

      // Productos del pedido
      var prods = [];
      Object.keys(prodMap).forEach(function(idx) {
        var cant = Number(row[idx]) || 0;
        if (cant > 0) {
          var abrev = prodMap[idx];
          prods.push({a: abrev, q: cant});
          if (!c.productos[abrev]) c.productos[abrev] = {cant: 0, monto: 0, ultimo: null};
          c.productos[abrev].cant += cant;
          // Monto unitario aproximado: total/sumProds. Sin precio exacto por producto.
          if (fecha && (!c.productos[abrev].ultimo || fecha > c.productos[abrev].ultimo)) {
            c.productos[abrev].ultimo = fecha;
          }
        }
      });

      // Repartir total entre productos para estimar monto por abrev
      var totalProds = 0;
      prods.forEach(function(p) { totalProds += p.q; });
      if (totalProds > 0) {
        prods.forEach(function(p) {
          c.productos[p.a].monto += Math.round((p.q / totalProds) * total);
        });
      }

      c.pedidos.push({
        canal: cfg.name,
        nPed: String(nPed),
        row: r + 1,
        fecha: fechaStr,
        fechaSort: fechaSort,
        diaEntrega: diaEnt,
        total: total,
        estado: estado,
        pago: pago,
        fp: fp,
        productos: prods,
        club: club, deporte: deporte, grupo: grupo,
        barrio: barrio, subBarrio: subBarrio, domicilio: domicilio,
        nombre: nombre
      });

      c.countPedidos++;
      // ¿Pedido Home entregado en Estancias del Pilar? (retail del cockpit de Estancias)
      var _esHomeEst = (cfg.name === 'Home') && (String(barrio || '').trim().toLowerCase() === 'estancias del pilar');
      if (_esHomeEst && fecha && (!c.firstHome || fecha < c.firstHome)) c.firstHome = fecha;
      // Pendiente Home/Estancias (cargado, sin entregar). Los cancelados ya se
      // filtraron arriba, así que "no Entregado" == pendiente/en curso.
      if (_esHomeEst && estado !== 'Entregado') {
        c.pendHome++;
        var _pf = diaDate || fecha; // fecha de ENTREGA elegida; fallback: fecha del pedido
        if (_pf && (!c.proxHome || _pf < c.proxHome)) c.proxHome = _pf;
      }
      if (estado === 'Entregado') {
        c.countEntregados++;
        c.totalFacturado += total;
        if (pago === 'Cobrado') c.totalCobrado += total;
        else c.deuda += total;
        if (_esHomeEst) {
          c.entHome++; c.factHome += total;
          if (pago === 'Cobrado') c.cobrHome += total; else c.deudaHome += total;
          if (fecha && (!c.lastHome || fecha > c.lastHome)) c.lastHome = fecha;
        }
      }
    }
  });

  // Mergear sin-tel en index si hay match por nombre+barrio único
  Object.keys(sinTel).forEach(function(k) { index[k] = sinTel[k]; });

  // Aplicar fusiones manuales (no destructivo: viven en hoja "Clientes Merge").
  // Resuelve cadenas A→B→C y absorbe cada alias en su canónico final.
  try {
    var mergeMap = _crmGetMergeMap();
    var resolve = function(key) {
      var k = key, seen = {}, depth = 10;
      while (mergeMap[k] && !seen[k] && depth-- > 0) { seen[k] = true; k = mergeMap[k]; }
      return k;
    };
    Object.keys(mergeMap).forEach(function(aliasKey) {
      if (!index[aliasKey]) return;
      var canon = resolve(aliasKey);
      if (canon === aliasKey) return;
      if (!index[canon]) index[canon] = index[aliasKey];
      else _crmFoldBucket(index[canon], index[aliasKey]);
      delete index[aliasKey];
    });
  } catch (_e) { /* si el merge falla, devolvemos el índice sin fusionar */ }

  return index;
}

// Calcula KPIs derivados de un cliente agregado.
function _crmComputeKpis(c) {
  var hoy = new Date();
  var ticket = c.countEntregados > 0 ? Math.round(c.totalFacturado / c.countEntregados) : 0;
  // Frecuencia: días entre primer y último pedido / (n-1)
  var frecuencia = 0;
  if (c.firstFecha && c.lastFecha && c.countPedidos > 1) {
    var span = (c.lastFecha - c.firstFecha) / (1000 * 60 * 60 * 24);
    frecuencia = Math.round(span / (c.countPedidos - 1));
  }
  // Días desde última compra
  var diasUltima = c.lastFecha ? Math.round((hoy - c.lastFecha) / (1000 * 60 * 60 * 24)) : -1;
  // Estado
  var estado = 'Sin actividad';
  if (diasUltima >= 0) {
    if (c.firstFecha && (hoy - c.firstFecha) / (1000 * 60 * 60 * 24) <= 30 && c.countPedidos <= 2) estado = 'Nuevo';
    else if (diasUltima <= 30) estado = 'Activo';
    else if (diasUltima <= 90) estado = 'Dormido';
    else estado = 'Inactivo';
  }
  // VIP: >=10 pedidos en cualquier momento o >$500K facturado
  var vip = (c.countPedidos >= 10) || (c.totalFacturado >= 500000);
  return {
    ticket: ticket,
    frecuencia: frecuencia,
    diasUltima: diasUltima,
    estado: estado,
    vip: vip
  };
}

// Devuelve el "nombre representativo" (el más usado).
function _crmTopValue(map) {
  var top = '', max = 0;
  Object.keys(map || {}).forEach(function(k) {
    if (map[k] > max) { max = map[k]; top = k; }
  });
  return top;
}

// Endpoint: lista resumida de todos los clientes para la grilla.
// Cache server-side 120s (TTL).
// ?lite=1 → devuelve solo campos que necesita el buscador de la PWA (nombre,
// tel, barrio, subBarrio, lote, club, canalDom, pedidos, nombresAlt). El
// response baja de ~110KB a ~30KB y entra en CacheService (límite 100KB).
// Sin lite=1, devuelve el response completo para el panel CRM.
function _doGetCrmClientes(e) {
  var _lite = !!(e && e.parameter && e.parameter.lite);
  var _ck = _lite ? 'crm_clientes_lite_v1' : 'crm_clientes_v1';
  try {
    var _c = CacheService.getScriptCache();
    var _h = _c.get(_ck);
    if (_h) return ContentService.createTextOutput(_h).setMimeType(ContentService.MimeType.JSON);
  } catch (_e) { /* sin cache no rompe */ }
  var index = _crmBuildClientesIndex();
  var meta = _crmGetClientesMeta();
  var lista = [];
  Object.keys(index).forEach(function(key) {
    var c = index[key];
    var k = _crmComputeKpis(c);
    var m = c.tel ? meta[c.tel] : null;
    var nombreRep = (m && m.nombreCanonico) || _crmTopValue(c.nombres);
    var barrioRep = _crmTopValue(c.barrios);
    var subRep = _crmTopValue(c.subBarrios);
    var clubRep = _crmTopValue(c.clubes);
    // Lote/domicilio dominante — el cliente recurrente típicamente tiene el
    // mismo lote en todos los pedidos. Para invitados/familia varía, pero el
    // valor dominante es el correcto el 95% del tiempo.
    var loteRep = _crmTopValue(c.domicilios);
    var canalDom = _crmTopValue(c.canales);
    var _ocultos = (m && m.nombresOcultos) ? m.nombresOcultos : [];
    var nombresAlt = Object.keys(c.nombres).filter(function(n) { return n !== nombreRep && _ocultos.indexOf(n) < 0; });
    var telRep = c.tel || _crmTopValue(c.telefonos);
    // Corrección manual de ubicación (cuando el cliente cargó mal el dato en la tienda)
    var canalesFinal = Object.keys(c.canales);
    if (m && m.canalMapa) { canalDom = m.canalMapa; canalesFinal = [m.canalMapa]; }  // canal real (ej. pidió por Home pero es de Pilar)
    if (m && m.barrioMapa) barrioRep = m.barrioMapa;   // barrio privado real (ej. no es de Estancias)
    if (m && m.sinUbicacion) { subRep = ''; loteRep = ''; }   // dirección desconocida / no ubicar
    else {
      if (m && m.loteMapa) loteRep = m.loteMapa;
      if (m && m.subBarrioMapa) subRep = m.subBarrioMapa;
    }
    // Producto más comprado (por unidades) — para personalizar el mensaje del
    // Plan Semanal ("te dejo las milanesas que siempre llevás"). Solo el abrev.
    var prodTopAbrev = '';
    var _catSet = {};
    if (c.productos) {
      var _mxc = -1;
      Object.keys(c.productos).forEach(function(ab) {
        var ca = (c.productos[ab] && c.productos[ab].cant) || 0;
        if (ca > _mxc) { _mxc = ca; prodTopAbrev = ab; }
        var cat = _crmCategoriaDe(ab);
        if (cat) _catSet[cat] = true;
      });
    }
    var cats = Object.keys(_catSet);  // categorías compradas (para segmentos cross-sell)

    lista.push({
      key: key,
      tel: c.tel,
      telDisplay: telRep,
      nombre: nombreRep,
      nombresAlt: nombresAlt, // para detectar tipeos: Marcos / Marcps
      barrio: barrioRep,
      subBarrio: subRep,
      lote: loteRep,
      club: clubRep,
      canalDom: canalDom,
      canales: canalesFinal,
      pedidos: c.countPedidos,
      entregados: c.countEntregados,
      facturado: c.totalFacturado,
      cobrado: c.totalCobrado,
      deuda: c.deuda,
      ticket: k.ticket,
      frecuencia: k.frecuencia,
      diasUltima: k.diasUltima,
      ultimaFecha: c.lastFecha ? Utilities.formatDate(c.lastFecha, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : '',
      primeraFecha: c.firstFecha ? Utilities.formatDate(c.firstFecha, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : '',
      // Solo Home/Estancias (retail del cockpit de Estancias — sin Clubes ni otros canales)
      facturadoHome: c.factHome,
      entregadosHome: c.entHome,
      cobradoHome: c.cobrHome,
      deudaHome: c.deudaHome,
      ticketHome: c.entHome > 0 ? Math.round(c.factHome / c.entHome) : 0,
      ultimaFechaHome: c.lastHome ? Utilities.formatDate(c.lastHome, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : '',
      diasUltimaHome: c.lastHome ? Math.round((new Date() - c.lastHome) / 86400000) : -1,
      primeraFechaHome: c.firstHome ? Utilities.formatDate(c.firstHome, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : '',
      // Pedidos Home/Estancias en curso (cargados, sin entregar) + fecha de entrega
      // del más próximo. Para que un cliente con pedido para mañana no figure "Sin compras".
      pendientesHome: c.pendHome || 0,
      proxEntregaHome: c.proxHome ? Utilities.formatDate(c.proxHome, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : '',
      estado: k.estado,
      vip: k.vip,
      prodTop: prodTopAbrev,
      cats: cats,
      cumple: m ? m.cumple : '',
      notas: m ? m.notas : '',
      tags: m ? m.tags : '',
      aliasMp: m ? m.aliasMp : '',
      // Campos de meta para que la ficha (skeleton) tenga todo al instante y editar sea seguro
      nombreCanonico: (m && m.nombreCanonico) || '',
      nombresOcultos: (m && m.nombresOcultos) || [],
      subBarrioMapa: (m && m.subBarrioMapa) || '',
      loteMapa: (m && m.loteMapa) || '',
      sinUbicacion: !!(m && m.sinUbicacion),
      visitante: !!(m && m.visitante),   // no reside en Estancias (pidió a casa ajena)
      residencia: (m && m.residencia) || '',   // 'vive' | 'finde' | 'visita' | '' (sin clasificar)
      apodo: (m && m.apodo) || '',        // sobrenombre (para el CSV de WATI)
      barrioMapa: (m && m.barrioMapa) || '',
      canalMapa: (m && m.canalMapa) || '',
      canalesExtra: (m && m.canalesExtra) || []   // canales manuales adicionales (ej. Eduardo: Clubes compra + Home casa)
    });
  });
  // Ordenar por última compra desc
  lista.sort(function(a, b) {
    var fa = a.ultimaFecha ? a.ultimaFecha.split('/').reverse().join('') : '0';
    var fb = b.ultimaFecha ? b.ultimaFecha.split('/').reverse().join('') : '0';
    return fb.localeCompare(fa);
  });
  // En modo lite, dejamos solo lo que el buscador de la PWA usa para autocompletar.
  // Esto baja el response de ~110KB a ~30KB y entra en CacheService.
  var listaOut;
  if (_lite) {
    listaOut = lista.map(function(c){
      return {
        nombre: c.nombre,
        tel: c.tel,
        telDisplay: c.telDisplay,
        nombresAlt: c.nombresAlt,
        barrio: c.barrio,
        subBarrio: c.subBarrio,
        lote: c.lote,
        club: c.club,
        canalDom: c.canalDom,
        pedidos: c.pedidos
      };
    });
  } else {
    listaOut = lista;
  }
  var _resp = JSON.stringify({ts: Date.now(), clientes: listaOut, lite: _lite});
  // TTL: el buscador de la PWA (lite) tolera estar varias horas viejo — solo
  // autocompleta nombre/barrio/lote y un cliente nuevo se tipea a mano igual —
  // así que cacheamos 6h para que el rebuild pesado (~7-10s, escanea todas las
  // hojas) casi nunca corra. El CRM del Panel (no-lite) pide datos frescos → 120s.
  var _ttl = _lite ? 21600 : 120;
  try { CacheService.getScriptCache().put(_ck, _resp, _ttl); } catch (_e) { /* nada */ }
  return ContentService.createTextOutput(_resp).setMimeType(ContentService.MimeType.JSON);
}

// Endpoint: ficha completa de un cliente por key (telNorm o NOMBRE:xxx).
function _doGetCrmCliente(e) {
  var key = e && e.parameter && e.parameter.key;
  if (!key) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'falta key'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // Cache por cliente (120s): reabrir una ficha es instantáneo.
  var _ck = 'crm_ficha_' + Utilities.base64EncodeWebSafe(key).substring(0, 80);
  try { var _h = CacheService.getScriptCache().get(_ck); if (_h) return ContentService.createTextOutput(_h).setMimeType(ContentService.MimeType.JSON); } catch (_e) {}
  var index = _crmBuildClientesIndex();
  var meta = _crmGetClientesMeta();
  var c = index[key];
  if (!c) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'no encontrado'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var k = _crmComputeKpis(c);
  var m = c.tel ? meta[c.tel] : null;

  // Pedidos ordenados desc
  var pedidos = c.pedidos.slice().sort(function(a, b) { return b.fechaSort - a.fechaSort; });

  // Top productos (por monto)
  var topProds = Object.keys(c.productos).map(function(abrev) {
    var p = c.productos[abrev];
    return {
      abrev: abrev,
      cant: p.cant,
      monto: p.monto,
      ultimo: p.ultimo ? Utilities.formatDate(p.ultimo, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : ''
    };
  }).sort(function(a, b) { return b.monto - a.monto; });

  var _resp = JSON.stringify({
    ts: Date.now(),
    cliente: {
      key: key,
      tel: c.tel,
      nombre: (m && m.nombreCanonico) || _crmTopValue(c.nombres),
      nombresVariantes: c.nombres,         // todos los tipeos detectados con su frecuencia
      telefonos: c.telefonos,
      barrios: c.barrios,
      subBarrios: c.subBarrios,
      domicilios: c.domicilios,
      clubes: c.clubes,
      canales: c.canales,
      pedidos: pedidos,
      topProductos: topProds,
      totalFacturado: c.totalFacturado,
      totalCobrado: c.totalCobrado,
      deuda: c.deuda,
      countPedidos: c.countPedidos,
      countEntregados: c.countEntregados,
      kpis: k,
      meta: m || null,
      primeraFecha: c.firstFecha ? Utilities.formatDate(c.firstFecha, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : '',
      ultimaFecha: c.lastFecha ? Utilities.formatDate(c.lastFecha, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : ''
    }
  });
  try { CacheService.getScriptCache().put(_ck, _resp, 120); } catch (_e) {}
  return ContentService.createTextOutput(_resp).setMimeType(ContentService.MimeType.JSON);
}

// Endpoint: lista de productos enriquecida con ventas históricas.
function _doGetCrmProductos(e) {
  var shP = SS.getSheetByName('Productos');
  if (!shP) return ContentService.createTextOutput(JSON.stringify({productos: []})).setMimeType(ContentService.MimeType.JSON);

  // ─── Parámetros de período (Ola 2) ───
  // ?dias=N (default 30) o ?desde=YYYY-MM-DD&hasta=YYYY-MM-DD
  var prm = (e && e.parameter) || {};
  var hoy = new Date();
  var perDesde, perHasta, perDias;
  if (prm.desde && prm.hasta) {
    var pd = String(prm.desde).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    var ph = String(prm.hasta).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (pd && ph) {
      perDesde = new Date(+pd[1], +pd[2] - 1, +pd[3], 0, 0, 0);
      perHasta = new Date(+ph[1], +ph[2] - 1, +ph[3], 23, 59, 59);
      perDias = Math.max(1, Math.round((perHasta.getTime() - perDesde.getTime()) / 86400000));
    }
  }
  if (!perDesde || !perHasta) {
    perDias = Math.max(1, parseInt(prm.dias, 10) || 30);
    perHasta = hoy;
    perDesde = new Date(hoy.getTime() - perDias * 86400000);
  }
  var perPrevHasta = new Date(perDesde.getTime() - 1);
  var perPrevDesde = new Date(perDesde.getTime() - perDias * 86400000);

  // ─── Filtros: canal + barrio privado (sobre las VENTAS, no las compras) ───
  var fCanal = String(prm.canal || '').trim();
  if (!fCanal || fCanal.toLowerCase() === 'todos') fCanal = '';
  var fBarrio = String(prm.barrio || '').trim();
  if (!fBarrio || fBarrio.toLowerCase() === 'todos') fBarrio = '';
  var fBarrioKey = _barrioKey(fBarrio);
  var filtroActivo = !!(fCanal || fBarrio);
  var barrioUniverse = {};   // para poblar el dropdown (siempre completo)
  var canalSeen = {};

  var dataP = shP.getDataRange().getValues();
  var prods = [];
  for (var r = 1; r < dataP.length; r++) {
    var nombre = String(dataP[r][1] || '').trim();
    var abrev = String(dataP[r][2] || '').trim();
    if (!abrev) continue;
    prods.push({
      abrev: abrev,
      nombre: nombre,
      stockFisico: Number(dataP[r][5]) || 0,
      reservado: Number(dataP[r][6]) || 0,
      disponible: Number(dataP[r][7]) || 0,
      precio: Number(dataP[r][8]) || 0,
      costo: Number(dataP[r][9]) || 0,
      margenUnit: Number(dataP[r][10]) || 0,
      vendidosTotal: 0, vendidosMes: 0, vendidosSemana: 0,
      facturadoTotal: 0, facturadoMes: 0,
      pedidosCount: 0,
      clientesUnicos: 0,
      ultimaVenta: null,
      // Sell-through (Ola 1): compras a proveedor (OC Recibido, no Depósito)
      comprados7: 0, comprados30: 0, comprados90: 0,
      costoUltimo: 0, fechaCostoUltimo: null, provUltimo: '',
      // Ola 2: período configurable + comparativa
      vendidosPeriodo: 0, facturadoPeriodo: 0, compradosPeriodo: 0,
      vendidosPeriodoPrev: 0, compradosPeriodoPrev: 0,
      _clientes: {}
    });
  }
  var byAbrev = {};
  prods.forEach(function(p) { byAbrev[p.abrev] = p; });

  var hace30 = new Date(hoy.getTime() - 30 * 24 * 60 * 60 * 1000);
  var hace7 = new Date(hoy.getTime() - 7 * 24 * 60 * 60 * 1000);
  var hace90 = new Date(hoy.getTime() - 90 * 24 * 60 * 60 * 1000);

  // ─── Desglose por semana ISO: últimas 8 COMPLETAS (excluye la en curso) ───
  // Responde a "cuánto vendí de cada producto en la semana 24" (no solo el agregado).
  var tzW = 'America/Argentina/Buenos_Aires';
  var WK_N = 8;
  var curMon = _isoWeekMondayLocal(hoy);
  var wkTargets = [], wkIdx = {};
  for (var iw = WK_N; iw >= 1; iw--) {
    var wMon = new Date(curMon.getFullYear(), curMon.getMonth(), curMon.getDate() - iw * 7);
    var wSun = new Date(wMon.getFullYear(), wMon.getMonth(), wMon.getDate() + 6);
    wkIdx[wMon.getFullYear() + '-' + _isoWeek(wMon)] = wkTargets.length;
    wkTargets.push({
      w: _isoWeek(wMon),
      desde: Utilities.formatDate(wMon, tzW, 'dd/MM'),
      hasta: Utilities.formatDate(wSun, tzW, 'dd/MM')
    });
  }
  prods.forEach(function(p) { p.wk = []; for (var z = 0; z < WK_N; z++) p.wk.push(0); });

  var hojas = _crmHojasConfig();
  hojas.forEach(function(cfg) {
    var sh = SS.getSheetByName(cfg.name);
    if (!sh || sh.getLastRow() <= 1) return;
    var prodMap = _crmProductHeaders(sh, cfg.prodStart, cfg.prodEnd, cfg.tartaStart, cfg.tartaEnd);
    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var estado = String(row[cfg.est] || '').trim();
      if (estado !== 'Entregado') continue;
      canalSeen[cfg.name] = true;

      // Barrio de la fila: zona (col barrio) y privado (col subBarrio cuando existe)
      var bCanon = cfg.barrio >= 0 ? _barrioCanon_(String(row[cfg.barrio] || '').trim()) : '';
      var sCanon = (cfg.subBarrio != null && cfg.subBarrio >= 0) ? _barrioCanon_(String(row[cfg.subBarrio] || '').trim()) : '';
      // Universo para el dropdown (siempre, antes de filtrar): Home → zona+privado; resto → barrio
      _accBarrio(barrioUniverse, cfg.name, 'priv', bCanon);   // Barrio privado (col Barrio / Barrio Privado)
      if (sCanon && !_esBarrioPriv(sCanon)) _accBarrio(barrioUniverse, cfg.name, 'sub', sCanon); // Sub barrio real (no privado leakeado)

      // Aplicar filtros (sobre ventas)
      if (fCanal && cfg.name !== fCanal) continue;
      if (fBarrio && _barrioKey(bCanon) !== fBarrioKey && _barrioKey(sCanon) !== fBarrioKey) continue;

      var fecha = _crmToDate(row[cfg.fecha]);
      var wkSlot = -1;
      if (fecha) {
        var rk = _isoWeekMondayLocal(fecha).getFullYear() + '-' + _isoWeek(fecha);
        if (wkIdx[rk] !== undefined) wkSlot = wkIdx[rk];
      }
      var total = Number(row[cfg.total]) || 0;
      var nombreCli = String(row[cfg.cliente] || '').trim();
      var totalProds = 0;
      var prodsRow = [];
      Object.keys(prodMap).forEach(function(idx) {
        var cant = Number(row[idx]) || 0;
        if (cant > 0) {
          prodsRow.push({abrev: prodMap[idx], cant: cant});
          totalProds += cant;
        }
      });
      prodsRow.forEach(function(pr) {
        var p = byAbrev[pr.abrev];
        if (!p) return;
        p.vendidosTotal += pr.cant;
        if (wkSlot >= 0) p.wk[wkSlot] += pr.cant;
        if (fecha && fecha >= hace30) p.vendidosMes += pr.cant;
        if (fecha && fecha >= hace7) p.vendidosSemana += pr.cant;
        var monto = totalProds > 0 ? Math.round((pr.cant / totalProds) * total) : 0;
        p.facturadoTotal += monto;
        if (fecha && fecha >= hace30) p.facturadoMes += monto;
        // Ola 2: período configurable + comparativa
        if (fecha && fecha >= perDesde && fecha <= perHasta) {
          p.vendidosPeriodo += pr.cant;
          p.facturadoPeriodo += monto;
        } else if (fecha && fecha >= perPrevDesde && fecha <= perPrevHasta) {
          p.vendidosPeriodoPrev += pr.cant;
        }
        p.pedidosCount++;
        if (nombreCli) p._clientes[nombreCli] = (p._clientes[nombreCli] || 0) + pr.cant;
        if (fecha && (!p.ultimaVenta || fecha > p.ultimaVenta)) p.ultimaVenta = fecha;
      });
    }
  });

  // ─── Sell-through (Ola 1): leer hoja OC y sumar compras a proveedor por SKU ───
  var shOC = SS.getSheetByName('Orden de Compra');
  if (shOC && shOC.getLastRow() > 1) {
    var ocData = shOC.getRange(1, 1, shOC.getLastRow(), 25).getValues();
    for (var ro = 1; ro < ocData.length; ro++) {
      var rowO = ocData[ro];
      var estO = String(rowO[20] || '').trim();
      if (estO !== 'Recibido') continue;                // solo lo que entró
      var origenO = String(rowO[19] || '').trim();
      if (origenO.indexOf('Dep') === 0) continue;       // descartar Depósito (no es compra)
      var ab = String(rowO[11] || '').trim();
      if (!ab) continue;
      var p2 = byAbrev[ab];
      if (!p2) continue;
      var qty = Number(rowO[12]) || 0;
      if (qty <= 0) continue;
      var fechaO = _crmToDate(rowO[1]);                 // fecha creación OC
      if (fechaO) {
        if (fechaO >= hace7)  p2.comprados7  += qty;
        if (fechaO >= hace30) p2.comprados30 += qty;
        if (fechaO >= hace90) p2.comprados90 += qty;
        // Ola 2: período + período anterior
        if (fechaO >= perDesde && fechaO <= perHasta) p2.compradosPeriodo += qty;
        else if (fechaO >= perPrevDesde && fechaO <= perPrevHasta) p2.compradosPeriodoPrev += qty;
      }
      var costoU = Number(rowO[13]) || 0;
      if (costoU > 0 && fechaO && (!p2.fechaCostoUltimo || fechaO > p2.fechaCostoUltimo)) {
        p2.costoUltimo = costoU;
        p2.fechaCostoUltimo = fechaO;
        p2.provUltimo = String(rowO[9] || '').trim();
      }
    }
  }

  prods.forEach(function(p) {
    p.clientesUnicos = Object.keys(p._clientes).length;
    p.ultimaVenta = p.ultimaVenta ? Utilities.formatDate(p.ultimaVenta, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : '';
    p.fechaCostoUltimo = p.fechaCostoUltimo ? Utilities.formatDate(p.fechaCostoUltimo, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : '';
    p.sellThrough30 = p.comprados30 > 0 ? Math.round((p.vendidosMes / p.comprados30) * 100) : null;
    // Ola 2: sell-through del período + delta vs período anterior
    p.sellThroughPeriodo = p.compradosPeriodo > 0 ? Math.round((p.vendidosPeriodo / p.compradosPeriodo) * 100) : null;
    p.deltaVentas = p.vendidosPeriodoPrev > 0
      ? Math.round((p.vendidosPeriodo - p.vendidosPeriodoPrev) / p.vendidosPeriodoPrev * 100)
      : (p.vendidosPeriodo > 0 ? 999 : null);
    p.deltaCompras = p.compradosPeriodoPrev > 0
      ? Math.round((p.compradosPeriodo - p.compradosPeriodoPrev) / p.compradosPeriodoPrev * 100)
      : (p.compradosPeriodo > 0 ? 999 : null);
    delete p._clientes;
  });

  // Universo de barrios para el dropdown (orden por frecuencia desc)
  var barriosArr = Object.keys(barrioUniverse).map(function(k) { return barrioUniverse[k]; })
    .sort(function(a, b) { return b.n - a.n; });
  var canalesArr = ['Home', 'Pilar', 'Clubes', 'Red'].filter(function(c) { return canalSeen[c]; });

  var tz = 'America/Argentina/Buenos_Aires';
  return ContentService.createTextOutput(JSON.stringify({
    ts: Date.now(),
    productos: prods,
    semanas: wkTargets,
    filtros: {
      canales: canalesArr,
      barrios: barriosArr,           // [{canal, tipo:'zona'|'sub'|'barrio', v, n}]
      canalSel: fCanal || 'todos',
      barrioSel: fBarrio || 'todos',
      activo: filtroActivo
    },
    periodo: {
      dias: perDias,
      desde: Utilities.formatDate(perDesde, tz, 'yyyy-MM-dd'),
      hasta: Utilities.formatDate(perHasta, tz, 'yyyy-MM-dd'),
      desdeArg: Utilities.formatDate(perDesde, tz, 'dd/MM/yyyy'),
      hastaArg: Utilities.formatDate(perHasta, tz, 'dd/MM/yyyy'),
      prevDesde: Utilities.formatDate(perPrevDesde, tz, 'yyyy-MM-dd'),
      prevHasta: Utilities.formatDate(perPrevHasta, tz, 'yyyy-MM-dd'),
      prevDesdeArg: Utilities.formatDate(perPrevDesde, tz, 'dd/MM/yyyy'),
      prevHastaArg: Utilities.formatDate(perPrevHasta, tz, 'dd/MM/yyyy')
    }
  })).setMimeType(ContentService.MimeType.JSON);
}

// Endpoint: ficha producto detallada con top clientes y series.
function _doGetCrmProducto(e) {
  var abrev = e && e.parameter && e.parameter.abrev;
  if (!abrev) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'falta abrev'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var shP = SS.getSheetByName('Productos');
  var prodInfo = null;
  if (shP) {
    var dataP = shP.getDataRange().getValues();
    for (var r = 1; r < dataP.length; r++) {
      if (String(dataP[r][2] || '').trim() === abrev) {
        prodInfo = {
          abrev: abrev,
          nombre: String(dataP[r][1] || '').trim(),
          stockFisico: Number(dataP[r][5]) || 0,
          reservado: Number(dataP[r][6]) || 0,
          disponible: Number(dataP[r][7]) || 0,
          precio: Number(dataP[r][8]) || 0,
          costo: Number(dataP[r][9]) || 0,
          margenUnit: Number(dataP[r][10]) || 0
        };
        break;
      }
    }
  }
  if (!prodInfo) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'producto no encontrado'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Recolectar todas las ventas del producto
  var clientes = {}; // {nombre: {cant, monto, telNorm, ultimo}}
  var ventasPorMes = {}; // {YYYY-MM: cant}
  var ventasPorCanal = {};
  var pedidosRecientes = [];

  var hojas = _crmHojasConfig();
  hojas.forEach(function(cfg) {
    var sh = SS.getSheetByName(cfg.name);
    if (!sh || sh.getLastRow() <= 1) return;
    var prodMap = _crmProductHeaders(sh, cfg.prodStart, cfg.prodEnd, cfg.tartaStart, cfg.tartaEnd);
    var idxAbrev = -1;
    Object.keys(prodMap).forEach(function(idx) {
      if (prodMap[idx] === abrev) idxAbrev = parseInt(idx, 10);
    });
    if (idxAbrev < 0) return;
    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var estado = String(row[cfg.est] || '').trim();
      if (estado === 'Cancelado') continue;
      var cant = Number(row[idxAbrev]) || 0;
      if (cant <= 0) continue;
      var fecha = _crmToDate(row[cfg.fecha]);
      var nombreCli = String(row[cfg.cliente] || '').trim();
      var telRaw = cfg.tel >= 0 ? String(row[cfg.tel] || '').trim() : '';
      var telNorm = _crmNormTel(telRaw);
      var total = Number(row[cfg.total]) || 0;
      // Repartir
      var totalProds = 0;
      Object.keys(prodMap).forEach(function(idx) { totalProds += (Number(row[idx]) || 0); });
      var monto = totalProds > 0 ? Math.round((cant / totalProds) * total) : 0;

      var key = telNorm || ('NOMBRE:' + _crmNormNombre(nombreCli));
      if (!clientes[key]) clientes[key] = {nombre: nombreCli, telNorm: telNorm, cant: 0, monto: 0, ultimo: null};
      clientes[key].cant += cant;
      clientes[key].monto += monto;
      if (fecha && (!clientes[key].ultimo || fecha > clientes[key].ultimo)) {
        clientes[key].ultimo = fecha;
        if (nombreCli) clientes[key].nombre = nombreCli;
      }
      if (fecha) {
        var ym = fecha.getFullYear() + '-' + String(fecha.getMonth() + 1).padStart(2, '0');
        ventasPorMes[ym] = (ventasPorMes[ym] || 0) + cant;
      }
      ventasPorCanal[cfg.name] = (ventasPorCanal[cfg.name] || 0) + cant;
      if (estado === 'Entregado') {
        pedidosRecientes.push({
          canal: cfg.name,
          fecha: fecha ? Utilities.formatDate(fecha, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : '',
          fechaSort: fecha ? fecha.getTime() : 0,
          cliente: nombreCli,
          telNorm: telNorm,
          cant: cant,
          monto: monto,
          total: total
        });
      }
    }
  });

  var topClientes = Object.keys(clientes).map(function(k) {
    var v = clientes[k];
    return {
      key: k,
      nombre: v.nombre,
      telNorm: v.telNorm,
      cant: v.cant,
      monto: v.monto,
      ultimo: v.ultimo ? Utilities.formatDate(v.ultimo, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : ''
    };
  }).sort(function(a, b) { return b.cant - a.cant; }).slice(0, 20);

  pedidosRecientes.sort(function(a, b) { return b.fechaSort - a.fechaSort; });
  pedidosRecientes = pedidosRecientes.slice(0, 30);

  // ─── Compras (OC) del SKU: series mensuales + por proveedor + historial reciente ───
  var comprasPorMes = {};      // {YYYY-MM: cant}
  var comprasPorProv = {};     // {prov: {cant, monto, ultimaFecha, ultimoCosto}}
  var comprasRecientes = [];   // últimas 20 OC Recibidas
  var costoUltimo = 0, fechaCostoUltimoTs = 0, provCostoUltimo = '';
  var totalComprado = 0;

  var shOC2 = SS.getSheetByName('Orden de Compra');
  if (shOC2 && shOC2.getLastRow() > 1) {
    var ocData2 = shOC2.getRange(1, 1, shOC2.getLastRow(), 25).getValues();
    for (var ro2 = 1; ro2 < ocData2.length; ro2++) {
      var rO = ocData2[ro2];
      var ab2 = String(rO[11] || '').trim();
      if (ab2 !== abrev) continue;
      var est2 = String(rO[20] || '').trim();
      if (est2 !== 'Recibido') continue;
      var origen2 = String(rO[19] || '').trim();
      if (origen2.indexOf('Dep') === 0) continue;
      var qty2 = Number(rO[12]) || 0;
      if (qty2 <= 0) continue;
      var fO = _crmToDate(rO[1]);
      var costoU2 = Number(rO[13]) || 0;
      var prov2 = String(rO[9] || '').trim();
      var monto2 = costoU2 * qty2;
      totalComprado += qty2;
      if (fO) {
        var ymO = fO.getFullYear() + '-' + String(fO.getMonth() + 1).padStart(2, '0');
        comprasPorMes[ymO] = (comprasPorMes[ymO] || 0) + qty2;
        if (fO.getTime() > fechaCostoUltimoTs && costoU2 > 0) {
          costoUltimo = costoU2;
          fechaCostoUltimoTs = fO.getTime();
          provCostoUltimo = prov2;
        }
      }
      if (prov2) {
        if (!comprasPorProv[prov2]) comprasPorProv[prov2] = {cant: 0, monto: 0, ultimaFecha: null, ultimoCosto: 0};
        comprasPorProv[prov2].cant += qty2;
        comprasPorProv[prov2].monto += monto2;
        if (fO && (!comprasPorProv[prov2].ultimaFecha || fO > comprasPorProv[prov2].ultimaFecha)) {
          comprasPorProv[prov2].ultimaFecha = fO;
          if (costoU2 > 0) comprasPorProv[prov2].ultimoCosto = costoU2;
        }
      }
      comprasRecientes.push({
        oc: String(rO[0] || '').trim(),
        fecha: fO ? Utilities.formatDate(fO, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : '',
        fechaSort: fO ? fO.getTime() : 0,
        prov: prov2,
        cant: qty2,
        costoUnit: costoU2,
        monto: monto2
      });
    }
  }
  comprasRecientes.sort(function(a, b) { return b.fechaSort - a.fechaSort; });
  comprasRecientes = comprasRecientes.slice(0, 20);

  var proveedoresArr = Object.keys(comprasPorProv).map(function(pr) {
    var v = comprasPorProv[pr];
    return {
      prov: pr,
      cant: v.cant,
      monto: v.monto,
      ultimoCosto: v.ultimoCosto,
      ultimaFecha: v.ultimaFecha ? Utilities.formatDate(v.ultimaFecha, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : ''
    };
  }).sort(function(a, b) { return b.cant - a.cant; });

  return ContentService.createTextOutput(JSON.stringify({
    ts: Date.now(),
    producto: prodInfo,
    topClientes: topClientes,
    ventasPorMes: ventasPorMes,
    ventasPorCanal: ventasPorCanal,
    pedidosRecientes: pedidosRecientes,
    clientesUnicos: Object.keys(clientes).length,
    // Ola 1: compras
    comprasPorMes: comprasPorMes,
    proveedores: proveedoresArr,
    comprasRecientes: comprasRecientes,
    totalComprado: totalComprado,
    costoUltimo: costoUltimo,
    fechaCostoUltimo: fechaCostoUltimoTs ? Utilities.formatDate(new Date(fechaCostoUltimoTs), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : '',
    provCostoUltimo: provCostoUltimo
  })).setMimeType(ContentService.MimeType.JSON);
}

// POST: actualizar meta de un cliente (cumpleaños, notas, alias MP, nombre canónico)
function _doPostCrmUpdateClienteMeta(data) {
  var tel = String(data.tel || '').trim();
  if (!tel) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'falta tel'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var sh = SS.getSheetByName('Clientes Meta');
  if (!sh) {
    sh = SS.insertSheet('Clientes Meta');
    var hdr = ['Tel Normalizado', 'Nombre Canónico', 'Cumpleaños (DD/MM)', 'Alias MP', 'Notas', 'Tags', 'Updated'];
    sh.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold').setBackground('#331C1C').setFontColor('#F2E8C7');
    sh.setFrozenRows(1);
  }
  var lastRow = sh.getLastRow();
  var found = -1;
  if (lastRow > 1) {
    var tels = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < tels.length; i++) {
      if (String(tels[i][0]).trim() === tel) { found = i + 2; break; }
    }
  }
  var nombreCanonico = String(data.nombreCanonico || '').trim();
  var cumple = String(data.cumple || '').trim();
  var aliasMp = String(data.aliasMp || '').trim();
  var notas = String(data.notas || '').trim();
  var tags = String(data.tags || '').trim();
  // Nombres alternativos que Tadeo marcó para ocultar del "también:" (pipe-separated)
  var nombresOcultos = '';
  if (Array.isArray(data.nombresOcultos)) nombresOcultos = data.nombresOcultos.join('|');
  else nombresOcultos = String(data.nombresOcultos || '').trim();
  // Corrección de ubicación para el mapa (sub-barrio + lote) + "sin ubicación"
  var subBarrioMapa = String(data.subBarrioMapa || '').trim();
  var loteMapa = String(data.loteMapa || '').trim();
  var sinUbic = (data.sinUbicacion === true || data.sinUbicacion === 1 || data.sinUbicacion === '1' || String(data.sinUbicacion).toUpperCase() === 'TRUE') ? 'TRUE' : '';
  var barrioMapa = String(data.barrioMapa || '').trim();
  var canalMapa = String(data.canalMapa || '').trim();
  // Canales adicionales manuales (pipe-separated): el cliente aparece también en esos filtros.
  var canalesExtra = '';
  if (Array.isArray(data.canalesExtra)) canalesExtra = data.canalesExtra.filter(Boolean).join('|');
  else canalesExtra = String(data.canalesExtra || '').trim();
  var updated = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');

  if (found > 0) {
    sh.getRange(found, 1, 1, 14).setValues([[tel, nombreCanonico, cumple, aliasMp, notas, tags, updated, nombresOcultos, subBarrioMapa, loteMapa, sinUbic, barrioMapa, canalMapa, canalesExtra]]);
  } else {
    sh.appendRow([tel, nombreCanonico, cumple, aliasMp, notas, tags, updated, nombresOcultos, subBarrioMapa, loteMapa, sinUbic, barrioMapa, canalMapa, canalesExtra]);
    found = sh.getLastRow();
  }
  // Formato texto en la celda Cumpleaños para que Sheets no la convierta en fecha.
  if (cumple) sh.getRange(found, 3).setNumberFormat('@').setValue(cumple);
  // Apodo (col 16): se escribe SOLO si el caller lo mandó, para no pisarlo al editar
  // otros campos. El write principal (cols 1-14) no toca 15 (Visitante) ni 16 (Apodo).
  if (data.hasOwnProperty('apodo')) sh.getRange(found, 16).setValue(String(data.apodo || '').trim());
  // Invalidar caches para que la lista y la ficha reflejen el cambio al instante.
  _crmClearCache();
  try { CacheService.getScriptCache().remove('crm_ficha_' + Utilities.base64EncodeWebSafe(tel).substring(0, 80)); } catch (_e) {}
  return ContentService.createTextOutput(JSON.stringify({ok: true, tel: tel}))
    .setMimeType(ContentService.MimeType.JSON);
}

// Guarda el cumpleaños de un cliente en "Clientes Meta" SIN pisar el resto de la
// ficha (nombre canónico, alias, notas, tags). A diferencia de crmUpdateClienteMeta,
// este merge solo toca la columna Cumpleaños + Updated. Lo usa la tienda online:
// el cliente carga su cumple en el form y viaja con el pedido (data.cumple) o por
// la acción 'guardarCumpleCliente'. Formato esperado: "DD/MM" o "DD/MM/AAAA"
// (el filtro 🎂 del Panel lee solo DD/MM, así que el año es compatible).
function _guardarCumpleCliente(telRaw, cumpleRaw) {
  var tel = _crmNormTel(telRaw);
  var cumple = String(cumpleRaw || '').trim();
  if (!tel || !cumple) return false;
  var sh = SS.getSheetByName('Clientes Meta');
  if (!sh) {
    sh = SS.insertSheet('Clientes Meta');
    var hdr = ['Tel Normalizado', 'Nombre Canónico', 'Cumpleaños (DD/MM)', 'Alias MP', 'Notas', 'Tags', 'Updated'];
    sh.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold').setBackground('#331C1C').setFontColor('#F2E8C7');
    sh.setFrozenRows(1);
  }
  var lastRow = sh.getLastRow();
  var found = -1;
  if (lastRow > 1) {
    var tels = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < tels.length; i++) {
      if (String(tels[i][0]).trim() === tel) { found = i + 2; break; }
    }
  }
  var updated = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');
  if (found > 0) {
    // MERGE: solo Cumpleaños (col 3) + Updated (col 7). NO tocar el resto.
    // Formato texto ('@') para que Sheets NO convierta "13/05/2002" en fecha real.
    sh.getRange(found, 3).setNumberFormat('@').setValue(cumple);
    sh.getRange(found, 7).setValue(updated);
  } else {
    sh.appendRow([tel, '', cumple, '', '', '', updated]);
    sh.getRange(sh.getLastRow(), 3).setNumberFormat('@').setValue(cumple);
  }
  return true;
}

// POST: guardar solo el cumpleaños de un cliente (desde la tienda online).
function _doPostGuardarCumpleCliente(data) {
  try {
    var ok = _guardarCumpleCliente(data.tel, data.cumple);
    return ContentService.createTextOutput(JSON.stringify({ok: ok}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// POST: fusionar dos "clientes" (típicamente cuando son la misma persona con tipeo distinto)
// Estrategia: anota la fusión en Clientes Meta como "alias" del telNorm canónico.
// (El backend ya unifica por teléfono normalizado; esto cubre el caso de personas sin teléfono.)
function _doPostCrmMergeClientes(data) {
  var keyCanon = String(data.keyCanonical || '').trim();  // el cliente que queda
  var keyAlias = String(data.keyAlias || '').trim();       // el que se absorbe
  if (!keyCanon || !keyAlias || keyCanon === keyAlias) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'faltan keys o son iguales'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var sh = SS.getSheetByName('Clientes Merge');
  if (!sh) {
    sh = SS.insertSheet('Clientes Merge');
    sh.getRange(1, 1, 1, 3).setValues([['Alias Key', 'Canonical Key', 'Updated']])
      .setFontWeight('bold').setBackground('#331C1C').setFontColor('#F2E8C7');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 3, 200);
  }
  // Evitar duplicar la misma fila alias→canónico
  if (sh.getLastRow() > 1) {
    var ex = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    for (var r = 0; r < ex.length; r++) {
      if (String(ex[r][0] || '').trim() === keyAlias) {
        sh.getRange(r + 2, 2).setValue(keyCanon);  // re-apuntar alias existente
        sh.getRange(r + 2, 3).setValue(new Date());
        _crmClearCache();
        return ContentService.createTextOutput(JSON.stringify({ok: true, msg: 'fusión actualizada'}))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
  }
  sh.appendRow([keyAlias, keyCanon, new Date()]);
  _crmClearCache();
  return ContentService.createTextOutput(JSON.stringify({ok: true, msg: 'clientes fusionados'}))
    .setMimeType(ContentService.MimeType.JSON);
}

// Invalida el cache server-side del CRM para que la próxima lectura refleje cambios.
function _crmClearCache() {
  try {
    var c = CacheService.getScriptCache();
    c.remove('crm_clientes_v1');
    c.remove('crm_clientes_lite_v1');
  } catch (_e) { /* sin cache no rompe */ }
}

// ════════════════════════════════════════════════════════════
//  FICHA DE LOTE (Estancias) — capa editable manual por lote
//  Hoja "Lotes Estancias": [Sub Barrio, Lote, Estado Contacto, Familia,
//  Integrantes, Titular, Notas, Telefono, Updated]. Key = Sub Barrio + Lote.
//  Complementa la data de ventas (CRM) con el trabajo de territorio de Tadeo.
// ════════════════════════════════════════════════════════════
function _crmLotesSheet() {
  var sh = SS.getSheetByName('Lotes Estancias');
  if (!sh) {
    sh = SS.insertSheet('Lotes Estancias');
    sh.getRange(1, 1, 1, 13).setValues([['Sub Barrio', 'Lote', 'Estado Contacto', 'Familia', 'Integrantes', 'Titular', 'Notas', 'Telefono', 'Updated', 'Composicion JSON', 'Sobre Estado', 'Sobre Nota', 'Reside']])
      .setFontWeight('bold').setBackground('#331C1C').setFontColor('#F2E8C7');
    sh.setFrozenRows(1);
  }
  // Asegurar headers nuevos (cols J/K/L/M) en hojas creadas antes de esta versión
  var hdrs = {10:'Composicion JSON', 11:'Sobre Estado', 12:'Sobre Nota', 13:'Reside'};
  Object.keys(hdrs).forEach(function(col){
    col = Number(col);
    if (String(sh.getRange(1, col).getValue() || '').trim() !== hdrs[col]) {
      sh.getRange(1, col).setValue(hdrs[col]).setFontWeight('bold').setBackground('#331C1C').setFontColor('#F2E8C7');
    }
  });
  return sh;
}

function _doGetCrmLotes(e) {
  var sh = _crmLotesSheet();
  var out = [];
  if (sh.getLastRow() > 1) {
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 13).getValues();
    for (var r = 0; r < data.length; r++) {
      var sub = String(data[r][0] || '').trim(), lote = String(data[r][1] || '').trim();
      if (!sub && !lote) continue;
      var comp = null;
      var compRaw = String(data[r][9] || '').trim();
      if (compRaw) { try { comp = JSON.parse(compRaw); } catch (_e) {} }
      out.push({
        sub: sub, lote: lote,
        estadoContacto: String(data[r][2] || '').trim(),
        familia: String(data[r][3] || '').trim(),
        integrantes: String(data[r][4] || '').trim(),
        titular: String(data[r][5] || '').trim(),
        notas: String(data[r][6] || '').trim(),
        tel: String(data[r][7] || '').trim(),
        composicion: comp,
        sobreEstado: String(data[r][10] || '').trim(),
        sobreNota: String(data[r][11] || '').trim(),
        reside: String(data[r][12] || '').trim()
      });
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ok: true, lotes: out}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  FUSIÓN DE HOGARES (no destructiva, revertible)
//  Hoja "Hogares Merge": [From, To, Nota, Fecha]. From/To son las KEYS
//  normalizadas de hogar (subNorm|loteNorm) que usa el front (_hogBuild).
//  Fusionar = el hogar From se colapsa dentro del hogar To. Revertir = borrar fila.
//  La tab "Estancias del Pilar" aplica estas reglas sobre el resultado de _hogBuild.
// ════════════════════════════════════════════════════════════
function _hogMergeSheet() {
  var sh = SS.getSheetByName('Hogares Merge');
  if (!sh) {
    sh = SS.insertSheet('Hogares Merge');
    sh.getRange(1, 1, 1, 4).setValues([['From', 'To', 'Nota', 'Fecha']])
      .setFontWeight('bold').setBackground('#331C1C').setFontColor('#F2E8C7');
    sh.setFrozenRows(1);
  }
  return sh;
}
function _doGetHogaresMerge(e) {
  var sh = _hogMergeSheet();
  var out = [];
  if (sh.getLastRow() > 1) {
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
    for (var r = 0; r < data.length; r++) {
      var from = String(data[r][0] || '').trim();
      if (!from) continue;
      out.push({ from: from, to: String(data[r][1] || '').trim(), nota: String(data[r][2] || '').trim() });
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true, merges: out }))
    .setMimeType(ContentService.MimeType.JSON);
}
function _doPostHogaresMerge(data) {
  var from = String(data.from || '').trim();
  if (!from) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'falta from' })).setMimeType(ContentService.MimeType.JSON);
  var sh = _hogMergeSheet();
  var del = (data.del === 1 || data.del === '1' || data.del === true);
  var found = -1;
  if (sh.getLastRow() > 1) {
    var keys = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (var r = 0; r < keys.length; r++) {
      if (String(keys[r][0] || '').trim() === from) { found = r + 2; break; }
    }
  }
  if (del) {
    if (found > 0) sh.deleteRow(found);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, msg: 'separado' })).setMimeType(ContentService.MimeType.JSON);
  }
  var to = String(data.to || '').trim();
  if (!to) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'falta to' })).setMimeType(ContentService.MimeType.JSON);
  if (to === from) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'no se puede fusionar consigo mismo' })).setMimeType(ContentService.MimeType.JSON);
  var row = [from, to, String(data.nota || ''), new Date()];
  if (found > 0) sh.getRange(found, 1, 1, 4).setValues([row]);
  else sh.appendRow(row);
  return ContentService.createTextOutput(JSON.stringify({ ok: true, msg: 'fusionado' })).setMimeType(ContentService.MimeType.JSON);
}

function _doPostCrmLoteSave(data) {
  var sub = String(data.sub || '').trim(), lote = String(data.lote || '').trim();
  if (!sub || !lote) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'faltan sub/lote'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var sh = _crmLotesSheet();
  var compStr = data.composicion ? (typeof data.composicion === 'string' ? data.composicion : JSON.stringify(data.composicion)) : '';
  var row = [sub, lote, String(data.estadoContacto || ''), String(data.familia || ''),
            String(data.integrantes || ''), String(data.titular || ''), String(data.notas || ''),
            String(data.tel || ''), new Date(), compStr,
            String(data.sobreEstado || ''), String(data.sobreNota || ''), String(data.reside || '')];
  if (sh.getLastRow() > 1) {
    var keys = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    for (var r = 0; r < keys.length; r++) {
      if (String(keys[r][0] || '').trim() === sub && String(keys[r][1] || '').trim() === lote) {
        sh.getRange(r + 2, 1, 1, 13).setValues([row]);
        return ContentService.createTextOutput(JSON.stringify({ok: true, msg: 'lote actualizado'}))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
  }
  sh.appendRow(row);
  return ContentService.createTextOutput(JSON.stringify({ok: true, msg: 'lote creado'}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  VISITANTE — persona que pide a una casa de Estancias pero NO reside ahí
//  (novia/prima/amigo). Se la saca de Clientes y Hogares (no es residente), pero
//  el pedido/entrega queda intacto (fue real → sigue contando en el Pulso). Flag
//  de PERSONA en la meta (col 15), no-destructivo y reversible. Endpoint dedicado
//  que toca SOLO esa celda (no pisa nombre/cumple/ubicación como el editor de ficha).
// ════════════════════════════════════════════════════════════
function _doPostCrmSetVisitante(data) {
  var tel = String(data.tel || '').trim();
  if (!tel) return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'falta tel'})).setMimeType(ContentService.MimeType.JSON);
  var meta = _crmGetClientesMeta();  // asegura la hoja + header col 15 (Visitante)
  var sh = SS.getSheetByName('Clientes Meta');
  var on = (data.visitante === true || data.visitante === 1 || data.visitante === '1' || String(data.visitante).toUpperCase() === 'TRUE');
  var val = on ? 'TRUE' : '';
  var m = meta[tel];
  if (m && m._row) {
    sh.getRange(m._row, 15).setValue(val);
  } else {
    var updated = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');
    sh.appendRow([tel, '', '', '', '', '', updated, '', '', '', '', '', '', '', val]);
  }
  _crmClearCache();
  try { CacheService.getScriptCache().remove('crm_ficha_' + Utilities.base64EncodeWebSafe(tel).substring(0, 80)); } catch (_e) {}
  return ContentService.createTextOutput(JSON.stringify({ok: true, tel: tel, visitante: on})).setMimeType(ContentService.MimeType.JSON);
}

// Tipo de integrante (col 17): 'vive' | 'finde' | 'visita' | '' (sin clasificar).
// Generaliza al flag Visitante: escribe col 17 y mantiene col 15 (Visitante legacy)
// en sync (TRUE solo si 'visita'), para que _hogBuild y el resto sigan andando.
function _doPostCrmSetResidencia(data) {
  var tel = String(data.tel || '').trim();
  if (!tel) return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'falta tel'})).setMimeType(ContentService.MimeType.JSON);
  var r = String(data.residencia || '').trim().toLowerCase();
  if (r !== 'vive' && r !== 'finde' && r !== 'visita') r = '';   // '' = sin clasificar
  var visVal = (r === 'visita') ? 'TRUE' : '';
  var meta = _crmGetClientesMeta();  // asegura headers col 15/16/17
  var sh = SS.getSheetByName('Clientes Meta');
  var m = meta[tel];
  if (m && m._row) {
    sh.getRange(m._row, 17).setValue(r);
    sh.getRange(m._row, 15).setValue(visVal);   // sync legacy Visitante
  } else {
    var updated = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');
    sh.appendRow([tel, '', '', '', '', '', updated, '', '', '', '', '', '', '', visVal, '', r]);
  }
  _crmClearCache();
  try { CacheService.getScriptCache().remove('crm_ficha_' + Utilities.base64EncodeWebSafe(tel).substring(0, 80)); } catch (_e) {}
  return ContentService.createTextOutput(JSON.stringify({ok: true, tel: tel, residencia: r})).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  PUNTOS DE LOTE (correcciones de geometría sobre el GeoJSON base)
//  Hoja "Lotes Puntos": [ID, Lote, Lat, Lng, Borrado, Updated]. ID 'b<i>' = override
//  de un punto base (movido/renombrado); 'a<ts>' = punto agregado. Borrado=TRUE lo
//  oculta. El front carga el GeoJSON y aplica estas correcciones encima.
// ════════════════════════════════════════════════════════════
// Layout: [ID, Lote, Lat, Lng, Borrado, Updated, Tipo, Dueño]. Tipo (col 7) =
// 'casa' (default) | 'anexo' (parcela que es parte de otra casa) | 'baldio'
// (terreno vacío). Dueño (col 8) = nº de lote de la casa principal, solo para
// anexos (así el cliente registrado bajo cualquier parcela cae en la casa).
// Anexo y baldío se excluyen del denominador de penetración (no son unidades
// vendibles). Tipo/Dueño se agregaron DESPUÉS de Updated para no migrar datos.
function _crmPuntosSheet() {
  var sh = SS.getSheetByName('Lotes Puntos');
  if (!sh) {
    sh = SS.insertSheet('Lotes Puntos');
    sh.getRange(1, 1, 1, 8).setValues([['ID', 'Lote', 'Lat', 'Lng', 'Borrado', 'Updated', 'Tipo', 'Dueño']])
      .setFontWeight('bold').setBackground('#331C1C').setFontColor('#F2E8C7');
    sh.setFrozenRows(1);
  } else if (String(sh.getRange(1, 7).getValue() || '').trim() !== 'Tipo') {
    // Sheet viejo (6 cols): agrego los encabezados de Tipo/Dueño una sola vez.
    sh.getRange(1, 7, 1, 2).setValues([['Tipo', 'Dueño']])
      .setFontWeight('bold').setBackground('#331C1C').setFontColor('#F2E8C7');
  }
  return sh;
}

function _doGetCrmPuntos(e) {
  var sh = _crmPuntosSheet();
  var out = [];
  if (sh.getLastRow() > 1) {
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
    for (var r = 0; r < data.length; r++) {
      var id = String(data[r][0] || '').trim();
      if (!id) continue;
      out.push({
        id: id,
        lote: String(data[r][1] || '').trim(),
        lat: Number(data[r][2]),
        lng: Number(data[r][3]),
        borrado: String(data[r][4] || '').trim().toUpperCase() === 'TRUE',
        tipo: String(data[r][6] || '').trim().toLowerCase() || 'casa',
        dueno: String(data[r][7] || '').trim()
      });
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ok: true, puntos: out}))
    .setMimeType(ContentService.MimeType.JSON);
}

function _doPostCrmPuntoSave(data) {
  var id = String(data.id || '').trim();
  if (!id) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'falta id'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var sh = _crmPuntosSheet();
  var borrado = (data.borrado === 1 || data.borrado === '1' || data.borrado === true) ? 'TRUE' : '';
  var tipo = String(data.tipo || 'casa').trim().toLowerCase();
  if (tipo !== 'anexo' && tipo !== 'baldio') tipo = 'casa';
  var dueno = tipo === 'anexo' ? String(data.dueno || '').trim() : '';
  var row = [id, String(data.lote || ''), Number(data.lat) || '', Number(data.lng) || '', borrado, new Date(), tipo, dueno];
  if (sh.getLastRow() > 1) {
    var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (var r = 0; r < ids.length; r++) {
      if (String(ids[r][0] || '').trim() === id) {
        sh.getRange(r + 2, 1, 1, 8).setValues([row]);
        return ContentService.createTextOutput(JSON.stringify({ok: true, msg: 'punto actualizado'}))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
  }
  sh.appendRow(row);
  return ContentService.createTextOutput(JSON.stringify({ok: true, msg: 'punto guardado'}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  F3 · BITÁCORA DE INTERACCIONES (CRM relacional)
//  Hoja "Interacciones CRM": registro append-only de cada contacto con un
//  cliente que NO es un pedido (le escribí / no contestó / quedó en pedir...).
//  El CRM pasa de transaccional (solo veo lo que compró) a relacional (sé qué
//  hablamos y qué sigue). Cada fila es inmutable; el estado "vivo" de un cliente
//  (última vez contactado + próxima acción) se deriva de su fila más reciente.
//  Columnas: ID | Fecha | Cliente Key | Tel | Nombre | Resultado |
//            Próxima Acción | Nota | Origen
// ════════════════════════════════════════════════════════════
function _crmInteraccionesSheet() {
  var sh = SS.getSheetByName('Interacciones CRM');
  if (!sh) {
    sh = SS.insertSheet('Interacciones CRM');
    var hdr = ['ID', 'Fecha', 'Cliente Key', 'Tel', 'Nombre', 'Resultado', 'Próxima Acción', 'Nota', 'Origen'];
    sh.getRange(1, 1, 1, hdr.length).setValues([hdr])
      .setFontWeight('bold').setBackground('#331C1C').setFontColor('#F2E8C7');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 1, 150); sh.setColumnWidths(2, 1, 130); sh.setColumnWidths(3, 1, 160);
    sh.setColumnWidths(4, 1, 110); sh.setColumnWidths(5, 1, 150); sh.setColumnWidths(6, 1, 120);
    sh.setColumnWidths(7, 1, 110); sh.setColumnWidths(8, 1, 280); sh.setColumnWidths(9, 1, 80);
  }
  return sh;
}

// GET: devuelve TODAS las interacciones (compactas). El front las indexa por
// key y por tel para (a) silenciar contactados recientes en el cockpit y (b)
// mostrar la bitácora en cada ficha. Es un log liviano (Tadeo registra un puñado
// por día) → no paginamos por ahora.
function _doGetCrmInteracciones(e) {
  var sh = _crmInteraccionesSheet();
  var out = [];
  if (sh.getLastRow() > 1) {
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    for (var r = 0; r < data.length; r++) {
      var id = String(data[r][0] || '').trim();
      if (!id) continue;
      var fecha = data[r][1];
      var fechaStr = (fecha instanceof Date)
        ? Utilities.formatDate(fecha, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm')
        : String(fecha || '').trim();
      var prox = data[r][6];
      var proxStr = (prox instanceof Date)
        ? Utilities.formatDate(prox, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy')
        : String(prox || '').trim();
      out.push({
        id: id,
        fecha: fechaStr,
        key: String(data[r][2] || '').trim(),
        tel: String(data[r][3] || '').trim(),
        nombre: String(data[r][4] || '').trim(),
        resultado: String(data[r][5] || '').trim(),
        prox: proxStr,
        nota: String(data[r][7] || '').trim(),
        origen: String(data[r][8] || '').trim()
      });
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ts: Date.now(), interacciones: out}))
    .setMimeType(ContentService.MimeType.JSON);
}

// POST: registra una interacción (append). El ID es único (timestamp + random)
// para poder borrarla después sin depender del número de fila.
function _doPostCrmLogInteraccion(data) {
  var key = String(data.key || '').trim();
  var tel = String(data.tel || '').trim();
  var resultado = String(data.resultado || '').trim();
  if (!key && !tel) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'falta key o tel'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (!resultado) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'falta resultado'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var sh = _crmInteraccionesSheet();
  var now = new Date();
  var id = 'I' + now.getTime() + '-' + Math.floor(Math.random() * 1000);
  var fecha = Utilities.formatDate(now, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');
  var nombre = String(data.nombre || '').trim();
  var prox = String(data.prox || '').trim();   // dd/MM/yyyy o ''
  var nota = String(data.nota || '').trim();
  var origen = String(data.origen || '').trim();
  sh.appendRow([id, fecha, key, tel, nombre, resultado, prox, nota, origen]);
  // Fecha y Próxima Acción como texto para que Sheets no las re-interprete.
  var row = sh.getLastRow();
  sh.getRange(row, 2).setNumberFormat('@').setValue(fecha);
  if (prox) sh.getRange(row, 7).setNumberFormat('@').setValue(prox);
  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    interaccion: {id: id, fecha: fecha, key: key, tel: tel, nombre: nombre, resultado: resultado, prox: prox, nota: nota, origen: origen}
  })).setMimeType(ContentService.MimeType.JSON);
}

// POST: borra una interacción por ID (por si fue un error de tipeo).
function _doPostCrmDeleteInteraccion(data) {
  var id = String(data.id || '').trim();
  if (!id) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'falta id'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var sh = _crmInteraccionesSheet();
  if (sh.getLastRow() > 1) {
    var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === id) {
        sh.deleteRow(i + 2);
        return ContentService.createTextOutput(JSON.stringify({ok: true}))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'no encontrada'}))
    .setMimeType(ContentService.MimeType.JSON);
}

// POST: registra una CAMPAÑA enviada a una lista de clientes (difusión WATI).
// Escribe una fila por cliente en "Interacciones CRM" con origen='campaña' →
// el ERP usa eso para que un cliente NO entre en 2 difusiones el mismo día.
// data: {template, items:[{key,tel,nombre}]}
function _doPostCrmLogCampania(data) {
  var template = String(data.template || '').trim();
  var items = data.items;
  if (!template || !items || !items.length) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'falta template o items'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var sh = _crmInteraccionesSheet();
  var now = new Date();
  var fecha = Utilities.formatDate(now, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');
  var base = now.getTime();
  var rows = [], out = [];
  for (var i = 0; i < items.length; i++) {
    var it = items[i] || {};
    var key = String(it.key || '').trim();
    var tel = String(it.tel || '').trim();
    if (!key && !tel) continue;
    var id = 'C' + base + '-' + i;
    var nombre = String(it.nombre || '').trim();
    var res = '📣 Campaña';
    var nota = 'Campaña: ' + template;
    rows.push([id, fecha, key, tel, nombre, res, '', nota, 'campaña']);
    out.push({id: id, fecha: fecha, key: key, tel: tel, nombre: nombre, resultado: res, prox: '', nota: nota, origen: 'campaña'});
  }
  if (!rows.length) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'sin clientes válidos'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, rows.length, 9).setValues(rows);
  // Fecha como texto en todo el bloque para que Sheets no la re-interprete.
  sh.getRange(startRow, 2, rows.length, 1).setNumberFormat('@');
  return ContentService.createTextOutput(JSON.stringify({ok: true, n: rows.length, interacciones: out}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  ZONAS DE LOTES (sub-barrios de Estancias dibujados en el mapa)
//  Hoja "Lotes Zonas": [Sub Barrio, Poligono JSON ([[lat,lng],...]), Updated]
//  El sub-barrio de cada lote se deriva por punto-en-polígono en el front.
// ════════════════════════════════════════════════════════════
function _crmZonasSheet() {
  var sh = SS.getSheetByName('Lotes Zonas');
  if (!sh) {
    sh = SS.insertSheet('Lotes Zonas');
    sh.getRange(1, 1, 1, 4).setValues([['Sub Barrio', 'Poligono JSON', 'Updated', 'Tipo']])
      .setFontWeight('bold').setBackground('#331C1C').setFontColor('#F2E8C7');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 1, 160); sh.setColumnWidths(2, 1, 400);
  }
  // Asegurar header de Tipo (col D) en hojas creadas antes de esta versión
  if (String(sh.getRange(1, 4).getValue() || '').trim() !== 'Tipo') {
    sh.getRange(1, 4).setValue('Tipo').setFontWeight('bold').setBackground('#331C1C').setFontColor('#F2E8C7');
  }
  return sh;
}

function _doGetCrmZonas(e) {
  var sh = _crmZonasSheet();
  var out = [];
  if (sh.getLastRow() > 1) {
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    for (var r = 0; r < data.length; r++) {
      var nombre = String(data[r][0] || '').trim();
      var poly = String(data[r][1] || '').trim();
      var tipo = String(data[r][3] || '').trim() || 'zona';
      if (!nombre || !poly) continue;
      try { out.push({nombre: nombre, poly: JSON.parse(poly), tipo: tipo}); } catch (_e) {}
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ok: true, zonas: out}))
    .setMimeType(ContentService.MimeType.JSON);
}

function _doPostCrmZonaSave(data) {
  var nombre = String(data.nombre || '').trim();
  var poly = data.poly;
  var tipo = String(data.tipo || 'zona').trim();
  if (!nombre || !poly || !poly.length) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'faltan datos'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var sh = _crmZonasSheet();
  var polyStr = JSON.stringify(poly);
  if (sh.getLastRow() > 1) {
    var names = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (var r = 0; r < names.length; r++) {
      if (String(names[r][0] || '').trim() === nombre) {
        sh.getRange(r + 2, 2).setValue(polyStr);
        sh.getRange(r + 2, 3).setValue(new Date());
        sh.getRange(r + 2, 4).setValue(tipo);
        return ContentService.createTextOutput(JSON.stringify({ok: true, msg: 'zona actualizada'}))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
  }
  sh.appendRow([nombre, polyStr, new Date(), tipo]);
  return ContentService.createTextOutput(JSON.stringify({ok: true, msg: 'zona creada'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function _doPostCrmZonaDelete(data) {
  var nombre = String(data.nombre || '').trim();
  if (!nombre) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'falta nombre'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var sh = _crmZonasSheet();
  if (sh.getLastRow() > 1) {
    var names = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (var r = names.length - 1; r >= 0; r--) {
      if (String(names[r][0] || '').trim() === nombre) sh.deleteRow(r + 2);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ok: true}))
    .setMimeType(ContentService.MimeType.JSON);
}

/** POST action=marcarEntregadoAVendedor
 *  Marca filas de la hoja Red como "Entregado a Vendedor" = "Entregado".
 *  Esto saca esos pedidos del tab ARMADO de la PWA Ruta (Tadeo ya le paso la mercaderia
 *  al vendedor; el estado de entrega Marcos→cliente lo maneja Marcos aparte desde red.html).
 *  Body:
 *    { rows: [n1, n2, ...] }   numeros de fila explicitos (1-indexed)
 *    { dia: "yyyy-mm-dd" }     marca TODAS las filas Red con dia de entrega = ese dia
 *                              (incluye pedidos 100% OC que no aparecen en el bloque Armado).
 *    { vendedor: "Marcos Bottcher" }  opcional, filtra adicionalmente por vendedor.
 *  Las dos primeras se pueden combinar (rows y dia) — se unen los conjuntos. */
function _doPostMarcarEntregadoAVendedor(data) {
  var sh = SS.getSheetByName('Red');
  if (!sh) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'hoja Red no existe' })).setMimeType(ContentService.MimeType.JSON);
  // Asegurar que la columna BE (57) exista
  if (sh.getMaxColumns() < 57) sh.insertColumnsAfter(sh.getMaxColumns(), 57 - sh.getMaxColumns());
  if (String(sh.getRange(1, 57).getValue() || '').trim() !== 'Entregado a Vendedor') {
    sh.getRange(1, 57).setValue('Entregado a Vendedor');
  }

  var rowsSet = {};
  if (Array.isArray(data.rows)) {
    data.rows.forEach(function(n){ var r = Number(n)||0; if (r>=2) rowsSet[r] = true; });
  } else if (data.row) {
    var r0 = Number(data.row)||0; if (r0>=2) rowsSet[r0] = true;
  }

  // Si viene dia, agregar TODAS las filas Red con dia de entrega = ese dia (incluye 100% OC)
  var dia = String(data.dia || '').trim();
  var vendedorFilter = String(data.vendedor || '').trim();
  if (dia) {
    var lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      // Cols: H(8)=Vendedor, K(11)=Día de Entrega (fecha), L(12)=Estado de Entrega, BE(57)=Entregado a Vendedor
      var rng = sh.getRange(2, 1, lastRow - 1, 57).getValues();
      for (var i = 0; i < rng.length; i++) {
        var rowNum = i + 2;
        var diaCell = rng[i][10]; // K idx 10
        var feISO = _fechaAISO(diaCell);
        if (feISO !== dia) continue;
        var estL = String(rng[i][11] || '').trim();
        if (estL === 'Cancelado') continue;
        var entVend = String(rng[i][56] || '').trim();
        if (entVend === 'Entregado') continue;
        if (vendedorFilter) {
          var vendCell = String(rng[i][7] || '').trim();
          if (vendCell !== vendedorFilter) continue;
        }
        rowsSet[rowNum] = true;
      }
    }
  }

  var rows = Object.keys(rowsSet).map(Number);
  if (rows.length === 0) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'sin filas' })).setMimeType(ContentService.MimeType.JSON);

  // Setear "Entregado" en cada fila (col 57 = BE)
  rows.forEach(function(r){
    sh.getRange(r, 57).setValue('Entregado');
  });

  return ContentService.createTextOutput(JSON.stringify({ ok: true, n: rows.length })).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════════════════
//  GUARDADO EN STOCK — verificación de OCs Tadeo-Stock recibidas
//  Flujo: Tadeo recibe la mercadería del proveedor (Abastecimiento marca
//  Recibido + suma stock automáticamente). Después la trae al freezer.
//  Esta verificación es el último paso: confirmar que la mercadería
//  efectivamente quedó guardada en el depósito. Sirve como auditoría.
//  Cliente que califica: "Tadeo — Stock" (matchea por prefijo + 'stock').
//  Hoja OC col Y (25) = "Guardado en Stock" — vacío hasta que se confirma.
// ════════════════════════════════════════════════════════════════════════════

function _isClienteTadeoStock(cliente) {
  var s = String(cliente || '').toLowerCase().replace(/\s+/g, ' ').trim();
  return s.indexOf('tadeo') === 0 && s.indexOf('stock') !== -1;
}

/**
 * GET ?action=pendientesGuardarStock
 * Devuelve OCs Tadeo-Stock con Estado=Recibido y col Y (Guardado) vacía.
 * Agrega por producto (suma cantidades) para que la card en Ruta muestre
 * la suma a guardar. Incluye los row indices originales para que el POST
 * de confirmación los pueda marcar.
 *
 * Response shape:
 *   { ok:true,
 *     items:[{ abbr, prod, qty, rows:[r1,r2,...], proveedor }],
 *     totalRows: N }
 */
function _doGetPendientesGuardarStock() {
  var sh = SS.getSheetByName('Orden de Compra');
  if (!sh || sh.getLastRow() <= 1) {
    return ContentService.createTextOutput(JSON.stringify({ ok: true, items: [], totalRows: 0 })).setMimeType(ContentService.MimeType.JSON);
  }
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 25).getValues();
  // Cols: G(7)=Cliente, J(10)=Proveedor, K(11)=Producto, L(12)=Abreviatura,
  //       M(13)=Cantidad, U(21)=Estado OC, Y(25)=Guardado en Stock
  var byProd = {};
  var totalRows = 0;
  for (var i = 0; i < data.length; i++) {
    var cliente = String(data[i][6] || '');
    var estado = String(data[i][20] || '').trim();
    var guardado = String(data[i][24] || '').trim();
    if (!_isClienteTadeoStock(cliente)) continue;
    if (estado !== 'Recibido') continue;
    if (guardado) continue; // ya guardado
    var abbr = String(data[i][11] || '').trim();
    var prod = String(data[i][10] || '').trim();
    var qty = Number(data[i][12]) || 0;
    var prov = String(data[i][9] || '').trim();
    if (!abbr || qty <= 0) continue;
    var rowReal = i + 2;
    if (!byProd[abbr]) byProd[abbr] = { abbr: abbr, prod: prod, qty: 0, rows: [], proveedor: prov };
    byProd[abbr].qty += qty;
    byProd[abbr].rows.push(rowReal);
    totalRows++;
  }
  var items = Object.keys(byProd).map(function(k){ return byProd[k]; });
  items.sort(function(a,b){ return a.prod.localeCompare(b.prod); });
  return ContentService.createTextOutput(JSON.stringify({ ok: true, items: items, totalRows: totalRows })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * POST action=marcarGuardadoEnStock
 * Marca col Y (25) de cada row con timestamp ARG. Idempotente: si ya tiene
 * valor, lo preserva. NO toca stock físico (el stock ya se sumó al marcar
 * Recibido en Abastecimiento) — esta acción es solo verificación.
 * Body: { rows:[r1,r2,...], repartidor?:'nombre' }
 */
function _doPostMarcarGuardadoEnStock(data) {
  var sh = SS.getSheetByName('Orden de Compra');
  if (!sh) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'sheet' })).setMimeType(ContentService.MimeType.JSON);

  var rows = (data.rows || []).map(Number).filter(function(r){ return r >= 2; });
  if (rows.length === 0) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'sin filas' })).setMimeType(ContentService.MimeType.JSON);

  var ahora = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var marcados = 0;
  rows.forEach(function(r) {
    var actual = String(sh.getRange(r, 25).getValue() || '').trim();
    if (actual) return; // idempotente: ya guardado
    sh.getRange(r, 25).setValue(argNow);
    marcados++;
  });

  return ContentService.createTextOutput(JSON.stringify({ ok: true, n: marcados, total: rows.length })).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════════════════
//  SALDOS A FAVOR DE CLIENTES (Fase 1: registro)
//  Hoja "Saldos Clientes". Una fila = un movimiento (Crédito o Aplicación).
//  Saldo neto del cliente = Σ Crédito - Σ Aplicación.
//  Identificación: por teléfono normalizado (más único que nombre).
//  Caso de uso (17/05/2026): Lucia Fernandez Moores pagó \$40.150 por un pedido
//  de \$31.320 → \$8.830 quedan a favor para próxima compra.
// ════════════════════════════════════════════════════════════════════════════

function _normalizarTel(s) {
  return String(s || '').replace(/[^0-9]/g, '');
}

function _ensureSaldosClientesSheet() {
  var sh = SS.getSheetByName('Saldos Clientes');
  if (!sh) {
    sh = SS.insertSheet('Saldos Clientes');
    sh.getRange(1, 1, 1, 9).setValues([['Fecha','Cliente','Telefono','Tipo','Monto','Pedido Hoja','Pedido N','Notas','Repartidor']]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 9).setBackground('#331C1C').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sh;
}

/** POST action=crearSaldoCliente
 *  Body: { cliente, telefono, tipo:'Crédito'|'Aplicación', monto:N, pedidoHoja, pedidoId, notas, repartidor, metodo? }
 *  Crea una fila en hoja "Saldos Clientes". Monto siempre positivo; el signo lo da Tipo.
 *  Si llega `metodo` (Efectivo/Transferencia/Mixto), también registra el movimiento
 *  en caja para reflejar el flujo real de efectivo/MP:
 *    - Crédito + metodo → Ingreso (la plata extra entró a la caja).
 *    - Aplicación + metodo → Egreso (la caja "pierde" porque el cliente paga menos).
 */
function _doPostCrearSaldoCliente(data) {
  var cliente = String(data.cliente || '').trim();
  var telefono = _normalizarTel(data.telefono);
  var tipo = String(data.tipo || 'Crédito').trim();
  var monto = Number(data.monto) || 0;
  if (!cliente && !telefono) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'sin cliente ni telefono' })).setMimeType(ContentService.MimeType.JSON);
  if (monto <= 0) return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'monto inválido' })).setMimeType(ContentService.MimeType.JSON);
  if (tipo !== 'Crédito' && tipo !== 'Credito' && tipo !== 'Aplicación' && tipo !== 'Aplicacion') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'tipo inválido' })).setMimeType(ContentService.MimeType.JSON);
  }
  // Normalizar acentos
  if (tipo === 'Credito') tipo = 'Crédito';
  if (tipo === 'Aplicacion') tipo = 'Aplicación';

  var sh = _ensureSaldosClientesSheet();

  // ── IDEMPOTENCIA ── No duplicar el a-favor/aplicación si ya se registró para este
  // pedido (mismo tipo). Pasa cuando un cobro lento se reintenta. Col F=pedidoHoja,
  // G=pedidoId, D=tipo.
  var pedidoHojaIn = String(data.pedidoHoja || '').trim();
  var pedidoIdIn = String(data.pedidoId || '').trim();
  if (pedidoIdIn && sh.getLastRow() > 1) {
    var prev = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    for (var ip = 0; ip < prev.length; ip++) {
      if (String(prev[ip][6] || '').trim() === pedidoIdIn &&
          String(prev[ip][5] || '').trim() === pedidoHojaIn &&
          String(prev[ip][3] || '').trim() === tipo) {
        return ContentService.createTextOutput(JSON.stringify({ ok: true, already: true, saldo: _calcSaldoCliente(sh, telefono, cliente) })).setMimeType(ContentService.MimeType.JSON);
      }
    }
  }

  var ahora = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var pedidoRef = (String(data.pedidoHoja || '') + (data.pedidoId ? ' #' + data.pedidoId : '')).trim();
  sh.appendRow([
    argNow,
    cliente,
    telefono,
    tipo,
    monto,
    String(data.pedidoHoja || ''),
    String(data.pedidoId || ''),
    String(data.notas || ''),
    String(data.repartidor || '')
  ]);
  sh.getRange(sh.getLastRow(), 1).setNumberFormat('dd/MM/yyyy HH:mm');

  // Nota: el ajuste a caja efectivo/MP NO se hace acá. Se hace via col BE/BH
  // "A Favor / Aplicado" del pedido (que entra en la fórmula Facturado).
  // Esta hoja Saldos Clientes es solo historial/tracking del saldo del cliente.

  // Calcular saldo neto actual del cliente (por teléfono si hay, sino por nombre)
  var saldoNeto = _calcSaldoCliente(sh, telefono, cliente);
  return ContentService.createTextOutput(JSON.stringify({ ok: true, saldo: saldoNeto })).setMimeType(ContentService.MimeType.JSON);
}

/** GET ?action=saldoCliente&tel=...&cliente=...
 *  Devuelve el saldo neto a favor del cliente.
 *  Matcheo: si viene tel, prioriza match por tel normalizado; sino por nombre exacto.
 */
function _doGetSaldoCliente(e) {
  var tel = _normalizarTel(e && e.parameter && e.parameter.tel);
  var cliente = String((e && e.parameter && e.parameter.cliente) || '').trim();
  if (!tel && !cliente) return ContentService.createTextOutput(JSON.stringify({ ok: true, saldo: 0, movimientos: [] })).setMimeType(ContentService.MimeType.JSON);
  var sh = SS.getSheetByName('Saldos Clientes');
  if (!sh || sh.getLastRow() <= 1) return ContentService.createTextOutput(JSON.stringify({ ok: true, saldo: 0, movimientos: [] })).setMimeType(ContentService.MimeType.JSON);
  var saldo = _calcSaldoCliente(sh, tel, cliente);
  return ContentService.createTextOutput(JSON.stringify({ ok: true, saldo: saldo })).setMimeType(ContentService.MimeType.JSON);
}

function _calcSaldoCliente(sh, tel, cliente) {
  if (!sh || sh.getLastRow() <= 1) return 0;
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
  var nombreLower = String(cliente || '').toLowerCase().trim();
  var saldo = 0;
  for (var i = 0; i < data.length; i++) {
    var rTel = _normalizarTel(data[i][2]);
    var rCliente = String(data[i][1] || '').toLowerCase().trim();
    var match = (tel && rTel === tel) || (!tel && nombreLower && rCliente === nombreLower);
    if (!match) continue;
    var tipo = String(data[i][3] || '').trim();
    var monto = Number(data[i][4]) || 0;
    if (tipo === 'Crédito' || tipo === 'Credito') saldo += monto;
    else if (tipo === 'Aplicación' || tipo === 'Aplicacion') saldo -= monto;
  }
  return Math.max(0, Math.round(saldo));
}

// ══════════════════════════════════════════════════════════════
//  ANÁLISIS PROVEEDOR — dashboard para negociaciones
// ══════════════════════════════════════════════════════════════
//
// GET ?action=analisisProveedor
//   Devuelve TODAS las OCs (limpias) + catálogo de productos por proveedor.
//   Frontend filtra/agrupa por proveedor según necesite.
//
//   Notas de fidelidad de datos:
//   • Filtra OCs Canceladas.
//   • Las filas "OC-REC-*" históricas (cantidad agregada sin abr en col 11
//     pero con desglose textual en col 10) se EXPANDEN a filas virtuales
//     por producto. Sin esto se perdían semanas enteras de historia al
//     filtrar por categoría/producto.
//   • Cada OC expone su estado real (Pedido / Recibido) para que el
//     frontend pueda filtrar "Solo Recibido" si quiere la versión más
//     conservadora para la negociación.
function _doGetAnalisisProveedor(e) {
  var shOC = SS.getSheetByName('Orden de Compra');
  var shProv = SS.getSheetByName('Proveedores');

  function _money(v) {
    if (typeof v === 'number') return v;
    var s = String(v == null ? '' : v).replace(/[$\s]/g, '').replace(/\./g, '').replace(/,/g, '.');
    var n = Number(s);
    return isFinite(n) ? n : 0;
  }
  function _fmtF(raw) {
    if (raw instanceof Date) return Utilities.formatDate(raw, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
    var s = String(raw || '').trim().split(' ')[0];
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return ('0' + m[1]).slice(-2) + '/' + ('0' + m[2]).slice(-2) + '/' + m[3];
    return '';
  }
  function _parseAR(ddmmyyyy) {
    var m = String(ddmmyyyy || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])) : null;
  }

  // ── Catálogo de productos por proveedor ──
  // Cols hoja Proveedores: N°(0) Producto(1) Proveedor(2) Gustos(3) Abrev(4)
  //                       Info(5) Canales(6) Costo(7) VentaDirecta(8) Clubes(9) WhatsApp(10)
  var catalogo = [];
  var costoByAbr = {};   // para expandir RECs en filas con costo unitario coherente
  var whatsappByProv = {}; // proveedor → tel WhatsApp (primer no-vacío visto)
  if (shProv) {
    var pd = shProv.getDataRange().getValues();
    for (var i = 1; i < pd.length; i++) {
      var r = pd[i];
      if (!r[2]) continue;
      var abrC = String(r[4] || '');
      var provN = String(r[2] || '');
      var wa = String(r[10] || '').trim();
      var c = {
        producto: String(r[1] || ''),
        proveedor: provN,
        gustos: String(r[3] || ''),
        abr: abrC,
        canales: String(r[6] || ''),
        costo: _money(r[7]),
        precioRetail: _money(r[8]),
        precioClubes: _money(r[9])
      };
      catalogo.push(c);
      if (abrC) costoByAbr[abrC] = c.costo;
      if (wa && !whatsappByProv[provN]) whatsappByProv[provN] = wa;
    }
  }

  // ── Pagos a proveedores (hoja Pagos Proveedores) ──
  // Cols: Fecha(0) Proveedor(1) Efectivo(2) MercadoPago(3) Total(4) SemImputada(5) Notas(6) Bonificación(7)
  var pagos = [];
  var pagosByProv = {}; // proveedor → { total, ultimaFecha, count, totalEf, totalMP, ytd }
  var shPagos = SS.getSheetByName('Pagos Proveedores');
  if (shPagos && shPagos.getLastRow() > 1) {
    var dataPg = shPagos.getDataRange().getValues();
    for (var rp2 = 1; rp2 < dataPg.length; rp2++) {
      var row2 = dataPg[rp2];
      if (!row2[0] || !row2[1]) continue;
      var fechaPag = _fmtF(row2[0]);
      var provPag = String(row2[1] || '').trim();
      var ef = _money(row2[2]);
      var mp = _money(row2[3]);
      var tot = _money(row2[4]) || (ef + mp);
      var dtPag = _parseAR(fechaPag);
      var anioPag = dtPag ? dtPag.getFullYear() : 0;
      pagos.push({ fecha: fechaPag, proveedor: provPag, ef: ef, mp: mp, total: tot,
                   semImputada: String(row2[5] || ''), notas: String(row2[6] || ''),
                   bonif: _money(row2[7]) });
      if (!pagosByProv[provPag]) pagosByProv[provPag] = { total: 0, totalEf: 0, totalMP: 0, ytd: 0, count: 0, ultimaFecha: '' };
      pagosByProv[provPag].total += tot;
      pagosByProv[provPag].totalEf += ef;
      pagosByProv[provPag].totalMP += mp;
      if (anioPag === new Date().getFullYear()) pagosByProv[provPag].ytd += tot;
      pagosByProv[provPag].count += 1;
      if (!pagosByProv[provPag].ultimaFecha || _parseAR(pagosByProv[provPag].ultimaFecha) < dtPag) {
        pagosByProv[provPag].ultimaFecha = fechaPag;
      }
    }
  }

  // ── Parser de RECs históricas ──
  // Convierte "Packs (39 Mu + 20 JyQ + 11 CyQ) + Empanadas (7 CaC + 6 JyQ)"
  // en [{abr:'PPM',qty:39},{abr:'PPJyQ',qty:20},...].
  //
  // Mapeo (categoría texto → prefijo abr):
  //   "Packs" → "PP"        (PPM, PPJyQ, PPCyQ)
  //   "Empanadas" → "E"     (ECaC, EJyQ, ECyQ, EV)
  //   "Sorrentinos" → "S"   (SQB, SL, SCo, SPyP, SJyQ, SE, SCa)
  //   "Tortas" → "T"        (TG, TLC, TC)
  //   "Tartas" → "T" gourmet (TP, TJyQ, TCa, TV)  — distinto de Tortas
  //   "Franui" → "F"
  //   "Pizza Premium" → "P" (PMu, PMa, PJyQ, PCC, PJyM)
  //
  // Cada subnombre se normaliza y se matchea contra el catálogo del proveedor.
  function _expandREC(producto, totalQty, proveedor) {
    if (!producto) return null;
    // Match grupos "Cat (lista)" capturando categoría y contenido entre paréntesis.
    var groupRe = /([A-Za-zÁÉÍÓÚáéíóúñÑ ]+?)\s*\(([^)]+)\)/g;
    var matches = [];
    var m;
    while ((m = groupRe.exec(producto)) !== null) {
      matches.push({ cat: m[1].trim(), inner: m[2].trim() });
    }
    if (!matches.length) return null;

    // Catálogo del proveedor por categoría → array de productos
    var byCatProv = {};
    catalogo.forEach(function(c) {
      if (c.proveedor !== proveedor) return;
      // Normalizo Pack Pizza ↔ "Packs" (texto histórico)
      var catKey = c.producto.replace(/\s+/g, ' ').trim();
      if (!byCatProv[catKey]) byCatProv[catKey] = [];
      byCatProv[catKey].push(c);
    });
    // Alias categoría texto histórico → catálogo canonical (Sheets Proveedores col B)
    var catAliases = {
      'Packs': 'Pack Pizzas x2',
      'Pack': 'Pack Pizzas x2',
      'Pack Pizza': 'Pack Pizzas x2',
      'Pack Pizzas': 'Pack Pizzas x2',
      'Empanadas': 'Empanadas',
      'Sorrentinos': 'Sorrentinos',
      'Tortas': 'Tortas',
      'Tartas': 'Tartas',
      'Tartas Gourmet': 'Tartas',
      'Franuis': 'Franuis',
      'Franui': 'Franuis',
      'Pote de Franui': 'Franuis',
      'Pizza Premium': 'Pizzas Individuales',
      'Pizzas Premium': 'Pizzas Individuales',
      'Pizzas': 'Pizzas Individuales',
      'Pizzas Individuales': 'Pizzas Individuales',
      'Rolls': 'Wraps',
      'Rolls Gourmet': 'Wraps',
      'Wraps': 'Wraps'
    };

    var virtuales = [];
    matches.forEach(function(g) {
      var catName = catAliases[g.cat] || g.cat;
      var cands = byCatProv[catName] || [];
      // Items dentro: "39 Mu + 20 JyQ + 11 CyQ"
      var items = g.inner.split(/\s*\+\s*/);
      items.forEach(function(it) {
        var im = it.trim().match(/^(\d+)\s+(.+)$/);
        if (!im) return;
        var qty = Number(im[1]) || 0;
        var subname = im[2].trim();
        // Buscar el producto del catálogo cuyo "gustos" contenga el subname
        // (case-insensitive, robusto a abreviaturas "Mu" / "Muzzarella")
        var snLow = subname.toLowerCase();
        var found = null;
        for (var k = 0; k < cands.length; k++) {
          var c = cands[k];
          var gustosLow = (c.gustos || '').toLowerCase();
          var abrLow = (c.abr || '').toLowerCase();
          // Match fuerte: el subname aparece como token en gustos
          if (gustosLow.split(/[\s,]+/).some(function(t) { return t.startsWith(snLow) || snLow.startsWith(t); })) {
            found = c; break;
          }
          // Match alternativo: abreviatura
          if (abrLow.indexOf(snLow) >= 0) { found = c; break; }
        }
        if (found && qty > 0) {
          virtuales.push({ abr: found.abr, producto: found.producto, gustos: found.gustos, qty: qty, costo: found.costo });
        }
      });
    });
    // Sanity check: si la suma de virtuales no coincide con totalQty, igual devolvemos lo parseado pero marcamos.
    return virtuales.length ? virtuales : null;
  }

  // ── OCs (limpias, no canceladas) ──
  // Cols OC: N°(0) Fecha(1) Sem(2) Mes(3) Canal(4) NPed(5) Cliente(6) Tel(7) Dir(8)
  //          Prov(9) Producto(10) Abr(11) Cant(12) CostoU(13) CostoT(14)
  //          VentaU(15) IngresoT(16) Margen$(17) Margen%(18) Origen(19)
  //          Estado(20) FechaPedProv(21) FechaRec(22) Pagado(23) GuardadoStock(24)
  var ocs = [];
  if (shOC) {
    var od = shOC.getDataRange().getValues();
    for (var j = 1; j < od.length; j++) {
      var row = od[j];
      var estado = String(row[20] || '').trim();
      if (estado === 'Cancelado') continue;
      var fechaStr = _fmtF(row[1]);
      if (!fechaStr) continue;
      var dt = _parseAR(fechaStr);
      if (!dt) continue;
      var wk = _isoWeek(dt);
      var monthKey = dt.getFullYear() + '-' + (dt.getMonth() < 9 ? '0' + (dt.getMonth() + 1) : (dt.getMonth() + 1));
      var weekKey = dt.getFullYear() + '-W' + (wk < 10 ? '0' + wk : wk);
      var nOrden = String(row[0] || '');
      var proveedor = String(row[9] || '');
      var producto = String(row[10] || '');
      var abr = String(row[11] || '');
      var cant = Number(row[12]) || 0;
      var costoUnit = _money(row[13]);
      var costoTotal = _money(row[14]);

      // ── REC histórica: expandir en filas virtuales ──
      if (nOrden.indexOf('OC-REC') === 0 && !abr && producto.indexOf('(') >= 0) {
        var virtuales = _expandREC(producto, cant, proveedor);
        if (virtuales && virtuales.length) {
          virtuales.forEach(function(v) {
            ocs.push({
              nOrden: nOrden,
              fecha: fechaStr,
              weekKey: weekKey,
              monthKey: monthKey,
              canal: String(row[4] || ''),
              cliente: String(row[6] || ''),
              proveedor: proveedor,
              producto: v.producto,
              abr: v.abr,
              cantidad: v.qty,
              costoUnit: v.costo,
              costoTotal: v.qty * v.costo,
              precioUnit: 0,
              ingresoTotal: 0,
              margen: 0,
              estado: estado,
              esVirtual: true
            });
          });
          continue; // saltar la fila REC original
        }
        // Si no pudo expandir, sigue cargándola como estaba (sin abr).
      }

      ocs.push({
        nOrden: nOrden,
        fecha: fechaStr,
        weekKey: weekKey,
        monthKey: monthKey,
        canal: String(row[4] || ''),
        cliente: String(row[6] || ''),
        proveedor: proveedor,
        producto: producto,
        abr: abr,
        cantidad: cant,
        costoUnit: costoUnit,
        costoTotal: costoTotal,
        precioUnit: _money(row[15]),
        ingresoTotal: _money(row[16]),
        margen: _money(row[17]),
        estado: estado
      });
    }
  }

  // ── Resumen por proveedor (totales para selector) ──
  // Cuenta de OCs únicas (no de filas, porque las RECs expandidas comparten N°).
  var resumen = {};
  var seenOrden = {};
  var anioAct = new Date().getFullYear();
  var mesAct = new Date().getMonth() + 1;
  ocs.forEach(function(oc) {
    if (!oc.proveedor) return;
    if (!resumen[oc.proveedor]) resumen[oc.proveedor] = {
      qty: 0, costo: 0, ocs: 0,
      costoYtd: 0, qtyYtd: 0,
      costoMesAct: 0, qtyMesAct: 0,
      costoMesAnt: 0, qtyMesAnt: 0,
      ultimaFecha: ''
    };
    var R = resumen[oc.proveedor];
    R.qty += oc.cantidad;
    R.costo += oc.costoTotal;
    var key = oc.proveedor + '::' + oc.nOrden;
    if (!seenOrden[key]) {
      seenOrden[key] = true;
      R.ocs += 1;
    }
    var dtOc = _parseAR(oc.fecha);
    if (dtOc) {
      if (dtOc.getFullYear() === anioAct) { R.costoYtd += oc.costoTotal; R.qtyYtd += oc.cantidad; }
      if (dtOc.getFullYear() === anioAct && dtOc.getMonth() + 1 === mesAct) {
        R.costoMesAct += oc.costoTotal; R.qtyMesAct += oc.cantidad;
      }
      var mesAntNum = mesAct === 1 ? 12 : mesAct - 1;
      var anioMesAnt = mesAct === 1 ? anioAct - 1 : anioAct;
      if (dtOc.getFullYear() === anioMesAnt && dtOc.getMonth() + 1 === mesAntNum) {
        R.costoMesAnt += oc.costoTotal; R.qtyMesAnt += oc.cantidad;
      }
      if (!R.ultimaFecha || _parseAR(R.ultimaFecha) < dtOc) R.ultimaFecha = oc.fecha;
    }
  });

  // Deuda viva por proveedor = compras YTD recibidas + pendientes - pagos YTD
  // (aproximación buena para dashboard; FIFO exacto en endpoint busqueda)
  Object.keys(resumen).forEach(function(prov) {
    var pag = pagosByProv[prov] || { ytd: 0 };
    resumen[prov].pagadoYtd = pag.ytd || 0;
    resumen[prov].deudaCalc = Math.max(0, resumen[prov].costoYtd - (pag.ytd || 0));
    resumen[prov].ultimoPago = pag.ultimaFecha || '';
  });

  // Evolución de costo unitario por (proveedor, abr): primer y último costo del año + delta
  // Útil para detectar aumentos silenciosos.
  var costoEvol = {}; // key proveedor::abr → { primero, ultimo, qtyOcs, fechaPrim, fechaUlt }
  ocs.forEach(function(oc) {
    if (!oc.proveedor || !oc.abr || oc.costoUnit <= 0) return;
    var dtOc = _parseAR(oc.fecha);
    if (!dtOc || dtOc.getFullYear() !== anioAct) return;
    var k = oc.proveedor + '::' + oc.abr;
    if (!costoEvol[k]) costoEvol[k] = {
      proveedor: oc.proveedor, abr: oc.abr, producto: oc.producto, gustos: '',
      primero: oc.costoUnit, ultimo: oc.costoUnit, fechaPrim: oc.fecha, fechaUlt: oc.fecha, ocsCount: 1
    };
    var E = costoEvol[k];
    E.ocsCount++;
    var dtPrim = _parseAR(E.fechaPrim);
    var dtUlt = _parseAR(E.fechaUlt);
    if (dtPrim && dtOc < dtPrim) { E.primero = oc.costoUnit; E.fechaPrim = oc.fecha; }
    if (dtUlt && dtOc > dtUlt) { E.ultimo = oc.costoUnit; E.fechaUlt = oc.fecha; }
  });
  var evolArr = Object.keys(costoEvol).map(function(k) {
    var E = costoEvol[k];
    var delta = E.primero > 0 ? (E.ultimo - E.primero) / E.primero : 0;
    return { proveedor: E.proveedor, abr: E.abr, producto: E.producto,
             primero: E.primero, ultimo: E.ultimo, deltaPct: delta,
             fechaPrim: E.fechaPrim, fechaUlt: E.fechaUlt, ocsCount: E.ocsCount };
  });

  return ContentService
    .createTextOutput(JSON.stringify({
      ocs: ocs,
      catalogo: catalogo,
      resumen: resumen,
      pagos: pagos,
      whatsappByProv: whatsappByProv,
      costoEvolucion: evolArr,
      meta: { anio: anioAct, mes: mesAct, ts: Date.now() }
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  PRODUCTOS ANALYTICS — Subtab VENTAS → PRODUCTOS del Panel
//  GET ?action=productosAnalytics
//      &desde=YYYY-MM-DD&hasta=YYYY-MM-DD (default: ultimos 30 dias)
//      &canal=all|Home|Pilar|Clubes|Red (default: all)
// ════════════════════════════════════════════════════════════

// Mapeo abreviatura -> categoria (para analisis de mix)
function _prodCategoria(abrev) {
  if (!abrev) return 'Otro';
  var a = String(abrev).trim();
  if (a.indexOf('PP') === 0) return 'Pizzas';                 // packs de pizza → categoría Pizzas
  if (a.indexOf('P') === 0 && a !== 'PP') return 'Pizzas';    // pizzas individuales
  if (a.indexOf('S') === 0) return 'Sorrentinos';
  if (a.indexOf('E') === 0) return 'Empanadas';
  if (a === 'TG' || a === 'TLC' || a === 'TC') return 'Tortas';
  if (a === 'TP' || a === 'TJyQ' || a === 'TCa' || a === 'TV') return 'Tartas';
  if (a === 'F') return 'Postres';
  return 'Otro';
}

// Sub-categoría: dentro de Pizzas, distingue Pack Pizzas vs Pizzas Individuales.
function _prodSubcat(abrev) {
  var a = String(abrev || '').trim();
  if (a.indexOf('PP') === 0) return 'Pack Pizzas';
  if (a.charAt(0) === 'P') return 'Pizzas Individuales';
  return '';
}

// Asigna cuadrante BCG por volumen (unidades) y margen unit ($)
function _prodBCG(unidades, margenUnit, mediaUnid, mediaMargen) {
  var altoVol = unidades >= mediaUnid;
  var altoMargen = margenUnit >= mediaMargen;
  if (altoVol && altoMargen) return 'estrella';
  if (altoVol && !altoMargen) return 'vaca';
  if (!altoVol && altoMargen) return 'interrogante';
  return 'perro';
}

function _doGetProductosAnalytics(e) {
  var prm = (e && e.parameter) || {};
  var canal = String(prm.canal || 'all').trim();
  var hoy = new Date();

  var perDesde, perHasta, perDias;
  if (prm.desde && prm.hasta) {
    var pd = String(prm.desde).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    var ph = String(prm.hasta).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (pd && ph) {
      perDesde = new Date(+pd[1], +pd[2] - 1, +pd[3], 0, 0, 0);
      perHasta = new Date(+ph[1], +ph[2] - 1, +ph[3], 23, 59, 59);
      perDias = Math.max(1, Math.round((perHasta.getTime() - perDesde.getTime()) / 86400000));
    }
  }
  if (!perDesde || !perHasta) {
    perDias = Math.max(1, parseInt(prm.dias, 10) || 30);
    perHasta = hoy;
    perDesde = new Date(hoy.getTime() - perDias * 86400000);
  }
  var perPrevHasta = new Date(perDesde.getTime() - 1);
  var perPrevDesde = new Date(perDesde.getTime() - perDias * 86400000);
  // Período anterior EXPLÍCITO (ej. MoM calendario: junio 1-17 vs mayo 1-17). Pisa el rolling.
  if (prm.prevDesde && prm.prevHasta) {
    var ppd = String(prm.prevDesde).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    var pph = String(prm.prevHasta).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (ppd && pph) {
      perPrevDesde = new Date(+ppd[1], +ppd[2] - 1, +ppd[3], 0, 0, 0);
      perPrevHasta = new Date(+pph[1], +pph[2] - 1, +pph[3], 23, 59, 59);
    }
  }

  // Filtro por barrio privado (zona o sub-barrio). Reusa helpers del CRM.
  var fBarrio = String(prm.barrio || '').trim();
  if (!fBarrio || fBarrio.toLowerCase() === 'all' || fBarrio.toLowerCase() === 'todos') fBarrio = '';
  var fBarrioKey = _barrioKey(fBarrio);
  var barrioUniverse = {};

  var shP = SS.getSheetByName('Productos');
  if (!shP) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'no Productos sheet'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var dataP = shP.getDataRange().getValues();
  var prods = [];
  for (var rP = 1; rP < dataP.length; rP++) {
    var nombre = String(dataP[rP][1] || '').trim();
    var abrev = String(dataP[rP][2] || '').trim();
    if (!abrev) continue;
    var precio = Number(dataP[rP][8]) || 0;
    var costo = Number(dataP[rP][9]) || 0;
    prods.push({
      abrev: abrev,
      nombre: nombre,
      categoria: _prodCategoria(abrev),
      subcat: _prodSubcat(abrev),
      stockFisico: Number(dataP[rP][5]) || 0,
      reservado: Number(dataP[rP][6]) || 0,
      disponible: Number(dataP[rP][7]) || 0,
      precioRetail: precio,
      costo: costo,
      margenUnit: precio - costo,
      margenPct: precio > 0 ? (precio - costo) / precio : 0,
      unidades: 0,
      facturado: 0,
      pedidosCount: 0,
      clientesUnicos: 0,
      porCanal: { Home: 0, Pilar: 0, Clubes: 0, Red: 0 },
      _clientes: {},
      unidadesPrev: 0,
      facturadoPrev: 0,
      evol8sem: [0, 0, 0, 0, 0, 0, 0, 0],
      ultimaVenta: null,
      bcg: null
    });
  }
  var byAbrev = {};
  prods.forEach(function(p) { byAbrev[p.abrev] = p; });

  var hojas = _crmHojasConfig();
  var hojasUsadas = canal === 'all' ? hojas : hojas.filter(function(h) { return h.name === canal; });

  var combos = {};

  var lunesActual = new Date(hoy.getTime());
  var dow = lunesActual.getDay();
  var diasAtrasLun = dow === 0 ? 6 : dow - 1;
  lunesActual.setDate(lunesActual.getDate() - diasAtrasLun);
  lunesActual.setHours(0, 0, 0, 0);
  var inicio8sem = new Date(lunesActual.getTime() - 7 * 7 * 86400000);

  hojasUsadas.forEach(function(cfg) {
    var sh = SS.getSheetByName(cfg.name);
    if (!sh || sh.getLastRow() <= 1) return;
    var prodMap = _crmProductHeaders(sh, cfg.prodStart, cfg.prodEnd, cfg.tartaStart, cfg.tartaEnd);
    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var estado = String(row[cfg.est] || '').trim();
      if (estado !== 'Entregado') continue;
      var fecha = _crmToDate(row[cfg.fecha]);
      if (!fecha) continue;

      // Barrio: universo para el dropdown (siempre) + filtro sobre ventas
      var bCanon = cfg.barrio >= 0 ? _barrioCanon_(String(row[cfg.barrio] || '').trim()) : '';
      var sCanon = (cfg.subBarrio != null && cfg.subBarrio >= 0) ? _barrioCanon_(String(row[cfg.subBarrio] || '').trim()) : '';
      _accBarrio(barrioUniverse, cfg.name, 'priv', bCanon);   // Barrio privado (col Barrio / Barrio Privado)
      if (sCanon && !_esBarrioPriv(sCanon)) _accBarrio(barrioUniverse, cfg.name, 'sub', sCanon); // Sub barrio real (no privado leakeado)
      if (fBarrio && _barrioKey(bCanon) !== fBarrioKey && _barrioKey(sCanon) !== fBarrioKey) continue;

      var totalFila = Number(row[cfg.total]) || 0;
      var nombreCli = String(row[cfg.cliente] || '').trim();

      var totalCantPed = 0;
      var prodsRow = [];
      Object.keys(prodMap).forEach(function(idx) {
        var cant = Number(row[idx]) || 0;
        if (cant > 0) {
          prodsRow.push({ abrev: prodMap[idx], cant: cant });
          totalCantPed += cant;
        }
      });
      if (prodsRow.length === 0) continue;

      var enPeriodoActual = (fecha >= perDesde && fecha <= perHasta);
      if (enPeriodoActual && prodsRow.length >= 2) {
        for (var i = 0; i < prodsRow.length; i++) {
          for (var j = i + 1; j < prodsRow.length; j++) {
            var a1 = prodsRow[i].abrev, a2 = prodsRow[j].abrev;
            var key = (a1 < a2) ? a1 + '|' + a2 : a2 + '|' + a1;
            combos[key] = (combos[key] || 0) + 1;
          }
        }
      }

      prodsRow.forEach(function(pr) {
        var p = byAbrev[pr.abrev];
        if (!p) return;
        var monto = totalCantPed > 0 ? Math.round((pr.cant / totalCantPed) * totalFila) : 0;

        if (enPeriodoActual) {
          p.unidades += pr.cant;
          p.facturado += monto;
          p.pedidosCount++;
          p.porCanal[cfg.name] = (p.porCanal[cfg.name] || 0) + pr.cant;
          if (nombreCli) p._clientes[nombreCli] = (p._clientes[nombreCli] || 0) + pr.cant;
          if (!p.ultimaVenta || fecha > p.ultimaVenta) p.ultimaVenta = fecha;
        } else if (fecha >= perPrevDesde && fecha <= perPrevHasta) {
          p.unidadesPrev += pr.cant;
          p.facturadoPrev += monto;
        }
        if (fecha >= inicio8sem && fecha < new Date(lunesActual.getTime() + 7 * 86400000)) {
          var diff = Math.floor((fecha.getTime() - inicio8sem.getTime()) / (7 * 86400000));
          if (diff >= 0 && diff < 8) p.evol8sem[diff] += pr.cant;
        }
      });
    }
  });

  var prodsConVentas = prods.filter(function(p) { return p.unidades > 0; });
  var sumUnid = 0, sumMargen = 0;
  prodsConVentas.forEach(function(p) {
    sumUnid += p.unidades;
    sumMargen += p.margenUnit;
  });
  var mediaUnid = prodsConVentas.length > 0 ? sumUnid / prodsConVentas.length : 0;
  var mediaMargen = prodsConVentas.length > 0 ? sumMargen / prodsConVentas.length : 0;

  var totalFactPeriodo = 0;
  prods.forEach(function(p) { totalFactPeriodo += p.facturado; });

  prods.forEach(function(p) {
    p.clientesUnicos = Object.keys(p._clientes).length;
    delete p._clientes;

    p.porcMix = totalFactPeriodo > 0 ? p.facturado / totalFactPeriodo : 0;
    p.margenTotal = p.unidades * p.margenUnit;

    if (p.unidadesPrev > 0) {
      p.tendenciaPct = (p.unidades - p.unidadesPrev) / p.unidadesPrev;
    } else if (p.unidades > 0) {
      p.tendenciaPct = 999;
    } else {
      p.tendenciaPct = null;
    }

    p.velocidadDiaria = p.unidades / perDias;
    if (p.velocidadDiaria > 0) {
      p.diasCobertura = Math.round(p.disponible / p.velocidadDiaria);
    } else {
      p.diasCobertura = null;
    }

    if (p.unidades > 0) {
      p.bcg = _prodBCG(p.unidades, p.margenUnit, mediaUnid, mediaMargen);
    }

    var status = 'normal';
    if (p.unidades === 0 && p.unidadesPrev > 0) status = 'sin_movimiento';
    else if (p.diasCobertura !== null && p.diasCobertura <= 7 && p.unidades > 0) status = 'stock_critico';
    else if (p.tendenciaPct !== null && p.tendenciaPct !== 999 && p.tendenciaPct <= -0.3) status = 'caida_fuerte';
    else if (p.tendenciaPct !== null && p.tendenciaPct >= 0.5 && p.tendenciaPct !== 999) status = 'crecimiento';
    p.status = status;

    p.ultimaVenta = p.ultimaVenta
      ? Utilities.formatDate(p.ultimaVenta, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy')
      : '';
  });

  var alertas = {
    stockCritico: prods.filter(function(p) { return p.status === 'stock_critico'; })
      .sort(function(a, b) { return (a.diasCobertura || 0) - (b.diasCobertura || 0); })
      .map(function(p) { return { abrev: p.abrev, nombre: p.nombre, disponible: p.disponible, diasCobertura: p.diasCobertura, velocidad: Math.round(p.velocidadDiaria * 7 * 10) / 10 }; }),
    caidaFuerte: prods.filter(function(p) { return p.status === 'caida_fuerte'; })
      .sort(function(a, b) { return a.tendenciaPct - b.tendenciaPct; })
      .map(function(p) { return { abrev: p.abrev, nombre: p.nombre, unidades: p.unidades, unidadesPrev: p.unidadesPrev, tendenciaPct: p.tendenciaPct }; }),
    sinMovimiento: prods.filter(function(p) { return p.status === 'sin_movimiento'; })
      .map(function(p) { return { abrev: p.abrev, nombre: p.nombre, unidadesPrev: p.unidadesPrev, ultimaVenta: p.ultimaVenta }; }),
    crecimiento: prods.filter(function(p) { return p.status === 'crecimiento'; })
      .sort(function(a, b) { return b.tendenciaPct - a.tendenciaPct; })
      .map(function(p) { return { abrev: p.abrev, nombre: p.nombre, unidades: p.unidades, unidadesPrev: p.unidadesPrev, tendenciaPct: p.tendenciaPct }; })
  };

  var combosArr = Object.keys(combos).map(function(k) {
    var ps = k.split('|');
    return {
      a1: ps[0],
      n1: byAbrev[ps[0]] ? byAbrev[ps[0]].nombre : ps[0],
      a2: ps[1],
      n2: byAbrev[ps[1]] ? byAbrev[ps[1]].nombre : ps[1],
      pedidos: combos[k]
    };
  }).sort(function(a, b) { return b.pedidos - a.pedidos; }).slice(0, 10);

  var totales = {
    unidades: 0,
    facturado: 0,
    margenTotal: 0,
    pedidosUnicos: 0,
    unidadesPrev: 0
  };
  prods.forEach(function(p) {
    totales.unidades += p.unidades;
    totales.facturado += p.facturado;
    totales.margenTotal += p.margenTotal;
    totales.unidadesPrev += p.unidadesPrev;
  });
  totales.margenPct = totales.facturado > 0 ? totales.margenTotal / totales.facturado : 0;
  totales.tendenciaPct = totales.unidadesPrev > 0
    ? (totales.unidades - totales.unidadesPrev) / totales.unidadesPrev
    : null;

  var topProd = null;
  prods.forEach(function(p) {
    if (!topProd || p.unidades > topProd.unidades) topProd = p;
  });
  totales.topProducto = topProd && topProd.unidades > 0
    ? { abrev: topProd.abrev, nombre: topProd.nombre, unidades: topProd.unidades, facturado: topProd.facturado }
    : null;

  var mixCat = {};
  prods.forEach(function(p) {
    if (!mixCat[p.categoria]) mixCat[p.categoria] = { unidades: 0, facturado: 0 };
    mixCat[p.categoria].unidades += p.unidades;
    mixCat[p.categoria].facturado += p.facturado;
  });

  // Desglose de Pizzas: Pack Pizzas (×2) vs Individuales + total en pizzas físicas
  var pizzasDetalle = { packUnidades: 0, packPizzas: 0, packFact: 0, indivUnidades: 0, indivFact: 0 };
  prods.forEach(function(p) {
    if (p.categoria !== 'Pizzas') return;
    if (p.subcat === 'Pack Pizzas') { pizzasDetalle.packUnidades += p.unidades; pizzasDetalle.packPizzas += p.unidades * 2; pizzasDetalle.packFact += p.facturado; }
    else { pizzasDetalle.indivUnidades += p.unidades; pizzasDetalle.indivFact += p.facturado; }
  });
  pizzasDetalle.totalPizzas = pizzasDetalle.packPizzas + pizzasDetalle.indivUnidades;  // pizzas físicas

  var barriosArr = Object.keys(barrioUniverse).map(function(k) { return barrioUniverse[k]; })
    .sort(function(a, b) { return b.n - a.n; });

  var tz = 'America/Argentina/Buenos_Aires';
  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    ts: Date.now(),
    canal: canal,
    barrio: fBarrio || 'all',
    periodo: {
      dias: perDias,
      desde: Utilities.formatDate(perDesde, tz, 'yyyy-MM-dd'),
      hasta: Utilities.formatDate(perHasta, tz, 'yyyy-MM-dd'),
      desdeArg: Utilities.formatDate(perDesde, tz, 'dd/MM/yyyy'),
      hastaArg: Utilities.formatDate(perHasta, tz, 'dd/MM/yyyy')
    },
    periodoPrev: {
      desdeArg: Utilities.formatDate(perPrevDesde, tz, 'dd/MM/yyyy'),
      hastaArg: Utilities.formatDate(perPrevHasta, tz, 'dd/MM/yyyy')
    },
    filtros: { barrios: barriosArr },   // [{canal, tipo:'priv'|'sub', v, n}]
    bcgEjes: { mediaUnid: mediaUnid, mediaMargen: mediaMargen },  // umbrales de los cuadrantes (scatter)
    pizzasDetalle: pizzasDetalle,       // Pack Pizzas (×2) vs Individuales + total pizzas físicas
    productos: prods,
    totales: totales,
    alertas: alertas,
    combos: combosArr,
    mixCategorias: mixCat
  })).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  CIERRE MENSUAL — Subtab Inicio → CIERRE del Panel
//  GET ?action=cierreMensual&mes=YYYY-MM  (default: mes en curso)
//  Devuelve TODOS los datos del mes en un solo JSON:
//  resumen, eerr, tendencia6m, canales, productos, clientes, caja, insights
// ════════════════════════════════════════════════════════════

function _parseMesParam(s) {
  if (!s) {
    var d = new Date();
    return { y: d.getFullYear(), m: d.getMonth() };
  }
  var m = String(s).match(/^(\d{4})-(\d{1,2})$/);
  if (!m) {
    var d2 = new Date();
    return { y: d2.getFullYear(), m: d2.getMonth() };
  }
  return { y: +m[1], m: +m[2] - 1 };
}

function _rangoMes(y, m) {
  var ini = new Date(y, m, 1, 0, 0, 0);
  var fin = new Date(y, m + 1, 0, 23, 59, 59);
  return { ini: ini, fin: fin, dias: fin.getDate() };
}

function _labelMes(y, m) {
  var meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  return meses[m] + ' ' + y;
}

// Lee todos los pedidos entregados de Home/Pilar/Clubes/Red en el rango.
// Devuelve array de {hoja, fecha, cliente, telefono, total, costo, margen, productos, barrio}.
function _leerPedidosEntregados(ini, fin) {
  var out = [];
  var hojas = _crmHojasConfig();
  hojas.forEach(function(cfg) {
    var sh = SS.getSheetByName(cfg.name);
    if (!sh || sh.getLastRow() <= 1) return;
    var prodMap = _crmProductHeaders(sh, cfg.prodStart, cfg.prodEnd, cfg.tartaStart, cfg.tartaEnd);
    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var estado = String(row[cfg.est] || '').trim();
      if (estado !== 'Entregado') continue;
      var fecha = _crmToDate(row[cfg.fecha]);
      if (!fecha || fecha < ini || fecha > fin) continue;
      var total = Number(row[cfg.total]) || 0;
      var costo = Number(row[cfg.costo]) || 0;
      var nombreCli = String(row[cfg.cliente] || '').trim();
      var tel = cfg.tel >= 0 ? String(row[cfg.tel] || '').trim() : '';
      var barrio = cfg.barrio >= 0 ? String(row[cfg.barrio] || '').trim() : '';
      var estadoPago = String(row[cfg.pago] || '').trim();
      var prods = [];
      Object.keys(prodMap).forEach(function(idx) {
        var cant = Number(row[idx]) || 0;
        if (cant > 0) prods.push({ abrev: prodMap[idx], cant: cant });
      });
      out.push({
        hoja: cfg.name,
        fecha: fecha,
        cliente: nombreCli,
        tel: tel,
        total: total,
        costo: costo,
        margen: total - costo,
        productos: prods,
        barrio: barrio,
        estadoPago: estadoPago
      });
    }
  });
  return out;
}

// Lee TODOS los entregados de todas las hojas operativas UNA sola vez.
// Permite filtrar localmente para múltiples meses sin re-leer (Cierre Mensual).
function _leerTodosEntregados() {
  var out = [];
  var hojas = _crmHojasConfig();
  hojas.forEach(function(cfg) {
    var sh = SS.getSheetByName(cfg.name);
    if (!sh || sh.getLastRow() <= 1) return;
    var prodMap = _crmProductHeaders(sh, cfg.prodStart, cfg.prodEnd, cfg.tartaStart, cfg.tartaEnd);
    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var estado = String(row[cfg.est] || '').trim();
      if (estado !== 'Entregado') continue;
      var fecha = _crmToDate(row[cfg.fecha]);
      if (!fecha) continue;
      var total = Number(row[cfg.total]) || 0;
      var costo = Number(row[cfg.costo]) || 0;
      var nombreCli = String(row[cfg.cliente] || '').trim();
      var tel = cfg.tel >= 0 ? String(row[cfg.tel] || '').trim() : '';
      var barrio = cfg.barrio >= 0 ? String(row[cfg.barrio] || '').trim() : '';
      var estadoPago = String(row[cfg.pago] || '').trim();
      var prods = [];
      Object.keys(prodMap).forEach(function(idx) {
        var cant = Number(row[idx]) || 0;
        if (cant > 0) prods.push({ abrev: prodMap[idx], cant: cant });
      });
      out.push({
        hoja: cfg.name, fecha: fecha, fechaTs: fecha.getTime(),
        cliente: nombreCli, tel: tel,
        total: total, costo: costo, margen: total - costo,
        productos: prods, barrio: barrio, estadoPago: estadoPago
      });
    }
  });
  return out;
}

function _doGetCierreMensual(e) {
  var prm = (e && e.parameter) || {};
  var mesP = _parseMesParam(prm.mes);
  var rango = _rangoMes(mesP.y, mesP.m);
  var ini = rango.ini, fin = rango.fin;
  var iniTs = ini.getTime(), finTs = fin.getTime();

  // Mes anterior (comparar)
  var prevM = mesP.m - 1, prevY = mesP.y;
  if (prevM < 0) { prevM = 11; prevY--; }
  var rangoPrev = _rangoMes(prevY, prevM);
  var iniPrev = rangoPrev.ini, finPrev = rangoPrev.fin;
  var iniPrevTs = iniPrev.getTime(), finPrevTs = finPrev.getTime();

  // ── 1 SOLA LECTURA de todas las hojas operativas ──
  // (antes leía 4 hojas x 8 períodos = 32 lecturas. Ahora 4 lecturas totales.)
  var allEntregados = _leerTodosEntregados();
  var pedidos = allEntregados.filter(function(p) { return p.fechaTs >= iniTs && p.fechaTs <= finTs; });
  var pedidosPrev = allEntregados.filter(function(p) { return p.fechaTs >= iniPrevTs && p.fechaTs <= finPrevTs; });

  // ── RESUMEN ──
  var resumen = {
    facturado: 0, costo: 0, margen: 0,
    pedidos: pedidos.length,
    cobrado: 0, sin_cobrar: 0, sin_cobrar_cant: 0
  };
  pedidos.forEach(function(p) {
    resumen.facturado += p.total;
    resumen.costo += p.costo;
    resumen.margen += p.margen;
    if (p.estadoPago.toLowerCase() === 'cobrado') resumen.cobrado += p.total;
    else { resumen.sin_cobrar += p.total; resumen.sin_cobrar_cant++; }
  });
  resumen.margen_pct = resumen.facturado > 0 ? resumen.margen / resumen.facturado : 0;
  resumen.ticket = resumen.pedidos > 0 ? resumen.facturado / resumen.pedidos : 0;

  var resumenPrev = { facturado: 0, margen: 0, pedidos: pedidosPrev.length };
  pedidosPrev.forEach(function(p) {
    resumenPrev.facturado += p.total;
    resumenPrev.margen += p.margen;
  });
  resumen.pct_facturado = resumenPrev.facturado > 0 ? (resumen.facturado - resumenPrev.facturado) / resumenPrev.facturado : null;
  resumen.pct_margen = resumenPrev.margen > 0 ? (resumen.margen - resumenPrev.margen) / resumenPrev.margen : null;
  resumen.pct_pedidos = resumenPrev.pedidos > 0 ? (resumen.pedidos - resumenPrev.pedidos) / resumenPrev.pedidos : null;
  resumen.facturado_prev = resumenPrev.facturado;
  resumen.margen_prev = resumenPrev.margen;
  resumen.pedidos_prev = resumenPrev.pedidos;

  // ── TENDENCIA 6 MESES (filtra del array en memoria, NO re-lee Sheets) ──
  var tendencia6m = [];
  for (var i = 5; i >= 0; i--) {
    var ty = mesP.y, tm = mesP.m - i;
    while (tm < 0) { tm += 12; ty--; }
    var rt = _rangoMes(ty, tm);
    var rtIniTs = rt.ini.getTime(), rtFinTs = rt.fin.getTime();
    var f = 0, mar = 0, pedCount = 0;
    for (var ai = 0; ai < allEntregados.length; ai++) {
      var ap = allEntregados[ai];
      if (ap.fechaTs < rtIniTs || ap.fechaTs > rtFinTs) continue;
      f += ap.total; mar += ap.margen; pedCount++;
    }
    tendencia6m.push({
      mes: _labelMes(ty, tm),
      mes_corto: _labelMes(ty, tm).split(' ')[0].slice(0, 3),
      yyyymm: ty + '-' + String(tm + 1).padStart(2, '0'),
      facturado: f,
      margen: mar,
      pedidos: pedCount
    });
  }

  // ── CANALES ──
  var canales = { Home: { p: 0, f: 0, m: 0 }, Pilar: { p: 0, f: 0, m: 0 }, Clubes: { p: 0, f: 0, m: 0 }, Red: { p: 0, f: 0, m: 0 } };
  pedidos.forEach(function(p) {
    if (canales[p.hoja]) {
      canales[p.hoja].p++;
      canales[p.hoja].f += p.total;
      canales[p.hoja].m += p.margen;
    }
  });
  var canalesPrev = { Home: 0, Pilar: 0, Clubes: 0, Red: 0 };
  pedidosPrev.forEach(function(p) { if (p.hoja in canalesPrev) canalesPrev[p.hoja] += p.total; });
  var canalesArr = Object.keys(canales).map(function(k) {
    var c = canales[k];
    var prev = canalesPrev[k] || 0;
    return {
      canal: k,
      pedidos: c.p,
      facturado: c.f,
      margen: c.m,
      ticket: c.p > 0 ? c.f / c.p : 0,
      pct_vs_prev: prev > 0 ? (c.f - prev) / prev : (c.f > 0 ? null : 0),
      facturado_prev: prev
    };
  });

  // ── PRODUCTOS ──
  var prodMap = {}; // abrev → {unidades, facturado, ped_count}
  pedidos.forEach(function(p) {
    var totalCant = 0;
    p.productos.forEach(function(pr) { totalCant += pr.cant; });
    if (totalCant === 0) return;
    p.productos.forEach(function(pr) {
      if (!prodMap[pr.abrev]) prodMap[pr.abrev] = { abrev: pr.abrev, unidades: 0, facturado: 0, pedidos: 0 };
      prodMap[pr.abrev].unidades += pr.cant;
      prodMap[pr.abrev].facturado += Math.round((pr.cant / totalCant) * p.total);
      prodMap[pr.abrev].pedidos++;
    });
  });
  var prodPrevMap = {};
  pedidosPrev.forEach(function(p) {
    var tc = 0;
    p.productos.forEach(function(pr) { tc += pr.cant; });
    p.productos.forEach(function(pr) {
      if (!prodPrevMap[pr.abrev]) prodPrevMap[pr.abrev] = 0;
      prodPrevMap[pr.abrev] += pr.cant;
    });
  });
  // Cruzar con hoja Productos para nombre/categoria
  var shP = SS.getSheetByName('Productos');
  var nombreByAbr = {}, costoByAbr = {};
  if (shP) {
    var dpP = shP.getDataRange().getValues();
    for (var rp = 1; rp < dpP.length; rp++) {
      var ab = String(dpP[rp][2] || '').trim();
      if (!ab) continue;
      nombreByAbr[ab] = String(dpP[rp][1] || '').trim();
      costoByAbr[ab] = Number(dpP[rp][9]) || 0;
    }
  }
  var prodsArr = Object.keys(prodMap).map(function(ab) {
    var p = prodMap[ab];
    var prevU = prodPrevMap[ab] || 0;
    p.nombre = nombreByAbr[ab] || ab;
    p.categoria = _prodCategoria(ab);
    p.unidades_prev = prevU;
    p.tendenciaPct = prevU > 0 ? (p.unidades - prevU) / prevU : (p.unidades > 0 ? 999 : null);
    return p;
  });
  var topVendidos = prodsArr.slice().sort(function(a, b) { return b.unidades - a.unidades; }).slice(0, 10);
  var topCrecimiento = prodsArr.filter(function(p) { return p.tendenciaPct !== null && p.tendenciaPct !== 999 && p.unidades_prev >= 3; }).sort(function(a, b) { return b.tendenciaPct - a.tendenciaPct; }).slice(0, 5);
  var topCaida = prodsArr.filter(function(p) { return p.tendenciaPct !== null && p.tendenciaPct !== 999 && p.unidades_prev >= 3; }).sort(function(a, b) { return a.tendenciaPct - b.tendenciaPct; }).slice(0, 5);
  // Mix categorías
  var mixCat = {};
  prodsArr.forEach(function(p) {
    if (!mixCat[p.categoria]) mixCat[p.categoria] = { unidades: 0, facturado: 0 };
    mixCat[p.categoria].unidades += p.unidades;
    mixCat[p.categoria].facturado += p.facturado;
  });

  // ── CLIENTES ──
  // Únicos del mes (por nombre lowercased, agrupando Red por vendedor descartado: usamos cliente final)
  var clientesMes = {};
  pedidos.forEach(function(p) {
    var c = (p.cliente || '').toLowerCase();
    if (!c) return;
    if (!clientesMes[c]) clientesMes[c] = { nombre: p.cliente, total: 0, pedidos: 0 };
    clientesMes[c].total += p.total;
    clientesMes[c].pedidos++;
  });
  // Histórico previo al mes — reusa allEntregados (NO re-lee Sheets)
  var histClientes = {};
  for (var hi = 0; hi < allEntregados.length; hi++) {
    var hp = allEntregados[hi];
    if (hp.fechaTs >= iniTs) continue; // solo histórico previo al mes en curso
    var hc = (hp.cliente || '').toLowerCase();
    if (!hc) continue;
    if (!histClientes[hc]) histClientes[hc] = { ultima: hp.fecha, total: 0 };
    histClientes[hc].total += hp.total;
    if (hp.fecha > histClientes[hc].ultima) histClientes[hc].ultima = hp.fecha;
  }
  var nuevos = 0, recurrentes = 0;
  Object.keys(clientesMes).forEach(function(c) {
    if (histClientes[c]) recurrentes++; else nuevos++;
  });
  // Dormidos: compraron mes anterior, NO este mes
  var clientesPrev = {};
  pedidosPrev.forEach(function(p) {
    var c = (p.cliente || '').toLowerCase();
    if (c) clientesPrev[c] = true;
  });
  var dormidos = 0;
  Object.keys(clientesPrev).forEach(function(c) {
    if (!clientesMes[c]) dormidos++;
  });
  // Top VIP del mes (mayor facturación)
  var topVip = Object.keys(clientesMes).map(function(c) { return clientesMes[c]; })
    .sort(function(a, b) { return b.total - a.total; }).slice(0, 10);

  // ── CAJA: leer Ingresos y Egresos del mes ──
  var ingresos = 0, ingresosCount = 0;
  var egresosPorCat = {};
  var egresosTotal = 0, egresosCount = 0;
  var shIng = SS.getSheetByName('Ingresos');
  if (shIng && shIng.getLastRow() > 1) {
    var dIng = shIng.getDataRange().getValues();
    for (var ri = 1; ri < dIng.length; ri++) {
      var fI = _crmToDate(dIng[ri][0]);
      if (!fI || fI < ini || fI > fin) continue;
      var monto = Number(dIng[ri][5]) || 0;
      ingresos += monto;
      ingresosCount++;
    }
  }
  var shEgr = SS.getSheetByName('Egresos');
  if (shEgr && shEgr.getLastRow() > 1) {
    var dEgr = shEgr.getDataRange().getValues();
    for (var re = 1; re < dEgr.length; re++) {
      var fE = _crmToDate(dEgr[re][0]);
      if (!fE || fE < ini || fE > fin) continue;
      var cat = String(dEgr[re][2] || '').trim() || 'Otro';
      var mE = Number(dEgr[re][5]) || 0;
      if (!egresosPorCat[cat]) egresosPorCat[cat] = 0;
      egresosPorCat[cat] += mE;
      egresosTotal += mE;
      egresosCount++;
    }
  }
  var egresosArr = Object.keys(egresosPorCat).map(function(k) { return { categoria: k, monto: egresosPorCat[k] }; }).sort(function(a, b) { return b.monto - a.monto; });
  var ebitda = resumen.margen + ingresos - egresosTotal;
  var ebitdaPct = resumen.facturado > 0 ? ebitda / resumen.facturado : 0;

  // ── INSIGHTS AUTOMÁTICOS ──
  var insights = [];
  // 1. Tendencia general
  if (resumen.pct_facturado !== null) {
    if (resumen.pct_facturado >= 0.15) {
      insights.push({ tipo: 'positivo', icono: '📈', titulo: 'Mes en crecimiento', msg: 'Facturado subió ' + Math.round(resumen.pct_facturado * 100) + '% vs el mes anterior. Sostené lo que funcionó.' });
    } else if (resumen.pct_facturado <= -0.15) {
      insights.push({ tipo: 'critico', icono: '📉', titulo: 'Caída fuerte de facturación', msg: 'Bajó ' + Math.round(Math.abs(resumen.pct_facturado) * 100) + '% vs mes anterior. Revisar canales y reactivar base.' });
    }
  }
  // 2. Margen
  if (resumen.margen_pct < 0.30 && resumen.facturado > 0) {
    insights.push({ tipo: 'alerta', icono: '⚠️', titulo: 'Margen bajo', msg: 'Margen bruto del ' + Math.round(resumen.margen_pct * 100) + '%. Revisar precios y costos de proveedores.' });
  }
  // 3. Canales que cayeron fuerte
  canalesArr.forEach(function(c) {
    if (c.pct_vs_prev !== null && c.pct_vs_prev <= -0.30 && c.facturado_prev > 0) {
      insights.push({ tipo: 'alerta', icono: '🚨', titulo: 'Canal ' + c.canal + ' en caída', msg: 'Cayó ' + Math.round(Math.abs(c.pct_vs_prev) * 100) + '% vs mes anterior (' + Math.round(c.facturado / 1000) + 'k vs ' + Math.round(c.facturado_prev / 1000) + 'k).' });
    }
  });
  // 4. Producto en crecimiento
  topCrecimiento.slice(0, 2).forEach(function(p) {
    if (p.tendenciaPct >= 0.5) {
      insights.push({ tipo: 'positivo', icono: '🌱', titulo: 'Producto en crecimiento', msg: p.nombre + ' +' + Math.round(p.tendenciaPct * 100) + '% vs mes pasado. Empujar este producto en contenido.' });
    }
  });
  // 5. Producto en caída
  topCaida.slice(0, 2).forEach(function(p) {
    if (p.tendenciaPct <= -0.3) {
      insights.push({ tipo: 'alerta', icono: '⬇️', titulo: 'Producto en caída', msg: p.nombre + ' ' + Math.round(p.tendenciaPct * 100) + '% vs mes pasado. ¿Cambió algo? Revisar precio/stock.' });
    }
  });
  // 6. Clientes nuevos vs dormidos
  if (dormidos > nuevos && dormidos >= 5) {
    insights.push({ tipo: 'alerta', icono: '😴', titulo: 'Perdiendo más que ganando', msg: dormidos + ' clientes se enfriaron, ' + nuevos + ' nuevos. Priorizar reactivación.' });
  } else if (nuevos >= 5) {
    insights.push({ tipo: 'positivo', icono: '✨', titulo: 'Buena adquisición', msg: nuevos + ' clientes nuevos este mes. Cuidá la primera experiencia para que vuelvan.' });
  }
  // 7. Cobros pendientes
  if (resumen.sin_cobrar > 50000) {
    insights.push({ tipo: 'alerta', icono: '💰', titulo: 'Plata sin cobrar', msg: '$' + Math.round(resumen.sin_cobrar / 1000) + 'k pendientes (' + resumen.sin_cobrar_cant + ' pedidos). Pasá por la tab Cobros.' });
  }
  // 8. EBITDA
  if (ebitdaPct < 0.10 && resumen.facturado > 100000) {
    insights.push({ tipo: 'alerta', icono: '💸', titulo: 'EBITDA bajo', msg: 'EBITDA del ' + Math.round(ebitdaPct * 100) + '%. Revisar gastos operativos.' });
  } else if (ebitdaPct >= 0.20) {
    insights.push({ tipo: 'positivo', icono: '💎', titulo: 'EBITDA saludable', msg: Math.round(ebitdaPct * 100) + '% sobre facturado. Buen mes.' });
  }

  var tz = 'America/Argentina/Buenos_Aires';
  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    ts: Date.now(),
    mes: { yyyymm: mesP.y + '-' + String(mesP.m + 1).padStart(2, '0'), label: _labelMes(mesP.y, mesP.m), dias: rango.dias, desde: Utilities.formatDate(ini, tz, 'yyyy-MM-dd'), hasta: Utilities.formatDate(fin, tz, 'yyyy-MM-dd') },
    mes_prev: { yyyymm: prevY + '-' + String(prevM + 1).padStart(2, '0'), label: _labelMes(prevY, prevM) },
    resumen: resumen,
    tendencia6m: tendencia6m,
    canales: canalesArr,
    productos: {
      top_vendidos: topVendidos,
      top_crecimiento: topCrecimiento,
      top_caida: topCaida,
      mix_categorias: mixCat
    },
    clientes: {
      total_unicos: Object.keys(clientesMes).length,
      nuevos: nuevos,
      recurrentes: recurrentes,
      dormidos: dormidos,
      top_vip: topVip
    },
    caja: {
      cobrado: resumen.cobrado,
      sin_cobrar: resumen.sin_cobrar,
      sin_cobrar_cant: resumen.sin_cobrar_cant,
      ingresos_no_venta: ingresos,
      egresos_total: egresosTotal,
      egresos_por_cat: egresosArr,
      ebitda: ebitda,
      ebitda_pct: ebitdaPct
    },
    insights: insights
  })).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  PLAN DE VENTAS (rol Gerente de Ventas)
//  Hojas: "Plan Ventas" (metas por mes/canal/barrio)
//         "Plan Acciones" (acciones del Gerente de Ventas)
//  Endpoints GET: planMes
//  Endpoints POST: planMetaSet, planAccionAdd, planAccionUpdate, planAccionDelete
// ════════════════════════════════════════════════════════════

// Asegura que existan las dos hojas con headers. Crea si faltan.
function _planEnsureSheets() {
  var shM = SS.getSheetByName('Plan Ventas');
  if (!shM) {
    shM = SS.insertSheet('Plan Ventas');
    var hdrM = ['Mes','Canal','Barrio','Meta Facturacion','Meta Pedidos','Meta Ticket','Meta Clientes','Notas','Updated'];
    shM.getRange(1,1,1,hdrM.length).setValues([hdrM])
      .setFontWeight('bold').setBackground('#331C1C').setFontColor('#F2E8C7');
    shM.setFrozenRows(1);
    shM.setColumnWidths(1, hdrM.length, 130);
  }
  var shA = SS.getSheetByName('Plan Acciones');
  if (!shA) {
    shA = SS.insertSheet('Plan Acciones');
    var hdrA = ['ID','Mes','Canal','Barrio','Descripcion','Responsable','Fecha Objetivo','Estado','Impacto Estimado','Impacto Real','Notas','Created','Updated'];
    shA.getRange(1,1,1,hdrA.length).setValues([hdrA])
      .setFontWeight('bold').setBackground('#331C1C').setFontColor('#F2E8C7');
    shA.setFrozenRows(1);
    shA.setColumnWidths(1, hdrA.length, 130);
  }
  return { metas: shM, acciones: shA };
}

// Parsea "Junio 2026" -> {year:2026, month:5 (0-indexed)} y viceversa.
var _PLAN_MS = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
function _planParseMes(label) {
  if (!label) return null;
  var p = String(label).trim().split(/\s+/);
  if (p.length < 2) return null;
  var low = p[0].toLowerCase();
  var m = -1;
  for (var i = 0; i < _PLAN_MS.length; i++) { if (_PLAN_MS[i].toLowerCase() === low) { m = i; break; } }
  var y = parseInt(p[1], 10);
  if (m < 0 || !y) return null;
  return { y: y, m: m };
}
function _planLabelMes(y, m) { return _PLAN_MS[m] + ' ' + y; }
// Label canonico "Julio 2026" (o null). Tolera mayus/minus: "julio 2026" -> "Julio 2026".
// Clave para que read y write hablen el mismo idioma sin importar como se guardo el mes.
function _planCanonMes(label) {
  // OJO: Sheets auto-convierte "julio 2026" en una FECHA (01/07/2026). getValues() la devuelve como Date.
  if (label instanceof Date) return _planLabelMes(label.getFullYear(), label.getMonth());
  var pr = _planParseMes(label);
  return pr ? _planLabelMes(pr.y, pr.m) : null;
}

// "Barrio canónico" — igual normalizador que el filtro Ventas.
function _planBarrioCanon(b) {
  if (!b) return '';
  var s = String(b).trim();
  if (!s || s === '-') return '';
  var k = s.toLowerCase();
  if (k.indexOf('estancias del r') === 0) return 'Estancias del Río';
  if (k.indexOf('estancias del pilar') === 0) return 'Estancias del Pilar';
  if (k.indexOf('alcanfor') >= 0) return 'Los Alcanfores';
  return s;
}

// "Hoja" del registro de Ventas: para VD usa la zona (Home/Pilar). Para los demás canales, el propio canal.
function _planHojaOp(canal, zona) {
  if (canal === 'Venta Directa') return zona || '';
  return canal || '';
}

// Lee TODAS las metas de un mes. Devuelve un map indexado por "Canal|Barrio".
function _planReadMetas(label) {
  var sh = SS.getSheetByName('Plan Ventas');
  if (!sh || sh.getLastRow() <= 1) return {};
  var data = sh.getDataRange().getValues();
  var out = {};
  var canon = _planCanonMes(label);
  for (var r = 1; r < data.length; r++) {
    if (_planCanonMes(data[r][0]) !== canon) continue;
    var canal = String(data[r][1] || '').trim();
    var barrio = String(data[r][2] || '').trim();
    var key = canal + '|' + barrio;
    out[key] = {
      canal: canal, barrio: barrio,
      metaFact:    Number(data[r][3]) || 0,
      metaPedidos: Number(data[r][4]) || 0,
      metaTicket:  Number(data[r][5]) || 0,
      metaClientes:Number(data[r][6]) || 0,
      notas: String(data[r][7] || '').trim(),
      updated: data[r][8] || '',
      _row: r + 1
    };
  }
  return out;
}

// Lee TODAS las acciones de un mes (ordenadas por fecha objetivo).
function _planReadAcciones(label) {
  var sh = SS.getSheetByName('Plan Acciones');
  if (!sh || sh.getLastRow() <= 1) return [];
  var data = sh.getDataRange().getValues();
  var out = [];
  var canon = _planCanonMes(label);
  for (var r = 1; r < data.length; r++) {
    if (_planCanonMes(data[r][1]) !== canon) continue;
    var fechaObj = data[r][6];
    var fechaStr = (fechaObj instanceof Date)
      ? Utilities.formatDate(fechaObj, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy')
      : String(fechaObj || '').trim();
    out.push({
      id: String(data[r][0]).trim(),
      mes: label,
      canal: String(data[r][2] || '').trim(),
      barrio: String(data[r][3] || '').trim(),
      desc: String(data[r][4] || '').trim(),
      responsable: String(data[r][5] || '').trim() || 'Tadeo',
      fechaObjetivo: fechaStr,
      estado: String(data[r][7] || '').trim() || 'Pendiente',
      impactoEstimado: Number(data[r][8]) || 0,
      impactoReal: Number(data[r][9]) || 0,
      notas: String(data[r][10] || '').trim(),
      _row: r + 1
    });
  }
  // Orden: por estado (Hechas al final) y luego por fecha objetivo asc
  var ESTADO_ORD = { 'Pendiente': 0, 'En curso': 1, 'Hecho': 2, 'Cancelado': 3 };
  out.sort(function(a, b) {
    var ea = ESTADO_ORD[a.estado] !== undefined ? ESTADO_ORD[a.estado] : 9;
    var eb = ESTADO_ORD[b.estado] !== undefined ? ESTADO_ORD[b.estado] : 9;
    if (ea !== eb) return ea - eb;
    return (a.fechaObjetivo || '').split('/').reverse().join('').localeCompare((b.fechaObjetivo || '').split('/').reverse().join(''));
  });
  return out;
}

// Calcula el REAL del mes leyendo Home/Pilar/Clubes/Red. Agrupa por canal+barrio.
function _planRealMes(yyyy, m0) {
  var hojas = _crmHojasConfig(); // reutilizo el config del CRM
  var real = {}; // {canal|barrio: {fact, pedidos, clientes:Set}}
  function key(c, b) { return (c || '') + '|' + (b || ''); }
  function bump(c, b, fact, cli) {
    var k = key(c, b);
    if (!real[k]) real[k] = { fact: 0, pedidos: 0, clientes: {} };
    real[k].fact += fact;
    real[k].pedidos++;
    if (cli) real[k].clientes[cli.toLowerCase()] = 1;
  }
  hojas.forEach(function(cfg) {
    var sh = SS.getSheetByName(cfg.name);
    if (!sh || sh.getLastRow() <= 1) return;
    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var estado = String(row[cfg.est] || '').trim();
      if (estado === 'Cancelado') continue;
      var fecha = _crmToDate(row[cfg.fecha]);
      if (!fecha) continue;
      if (fecha.getFullYear() !== yyyy || fecha.getMonth() !== m0) continue;
      var fact = Number(row[cfg.total]) || 0;
      if (!fact) continue;
      var cli = String(row[cfg.cliente] || '').trim();
      // Canal del registro: Home/Pilar = "Venta Directa"; otros = el nombre de la hoja
      var canal;
      if (cfg.name === 'Home' || cfg.name === 'Pilar') canal = 'Venta Directa';
      else canal = cfg.name;
      var barrio = '';
      if (cfg.name === 'Home') barrio = _planBarrioCanon(cfg.barrio >= 0 ? String(row[cfg.barrio] || '') : '');
      else if (cfg.name === 'Pilar') barrio = ''; // Pilar agregamos sin desglose barrio por ahora
      // Para discriminar Home/Pilar también guardo en una clave por hoja operativa
      bump(canal, barrio, fact, cli);
      // Y una clave por "hoja operativa" (Home, Pilar) para que el frontend pueda mostrar dos cards diferentes en VD
      if (cfg.name === 'Home' || cfg.name === 'Pilar') {
        bump('__hoja:' + cfg.name, barrio, fact, cli);
      }
    }
  });
  // Convertir Sets de clientes a count
  Object.keys(real).forEach(function(k) {
    real[k].clientesUnicos = Object.keys(real[k].clientes).length;
    delete real[k].clientes;
  });
  return real;
}

// Endpoint: ?action=planMes&mes=Junio%202026
function _doGetPlanMes(e) {
  var label = (e && e.parameter && e.parameter.mes) || '';
  if (!label) {
    // Default: mes actual AR
    var now = new Date();
    var ar = new Date(now.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
    label = _planLabelMes(ar.getFullYear(), ar.getMonth());
  }
  _planEnsureSheets();
  var parsed = _planParseMes(label);
  if (!parsed) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Mes invalido: ' + label }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  label = _planLabelMes(parsed.y, parsed.m); // canonizar: la respuesta y las lecturas usan "Julio 2026"
  var metas = _planReadMetas(label);
  var acciones = _planReadAcciones(label);
  var real = _planRealMes(parsed.y, parsed.m);

  // Días del mes y días transcurridos (en AR)
  var diasTot = new Date(parsed.y, parsed.m + 1, 0).getDate();
  var ar = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var diasTrans;
  if (ar.getFullYear() < parsed.y || (ar.getFullYear() === parsed.y && ar.getMonth() < parsed.m)) diasTrans = 0;
  else if (ar.getFullYear() > parsed.y || (ar.getFullYear() === parsed.y && ar.getMonth() > parsed.m)) diasTrans = diasTot;
  else diasTrans = ar.getDate();

  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    mes: label,
    yyyy: parsed.y,
    mm: parsed.m + 1,
    diasMes: diasTot,
    diasTrans: diasTrans,
    metas: metas,
    acciones: acciones,
    real: real,
    barriosHome: ['Estancias del Pilar', 'Los Alcanfores', 'Estancias del Río', 'Pilara'],
    canalesPrincipales: ['Venta Directa', 'Clubes', 'Red', 'Catering', 'B2B']
  })).setMimeType(ContentService.MimeType.JSON);
}

// POST: setear meta de un canal+barrio para un mes (upsert).
function _doPostPlanMetaSet(data) {
  var sheets = _planEnsureSheets();
  var sh = sheets.metas;
  var label = _planCanonMes(data.mes) || String(data.mes || '').trim();
  var canal = String(data.canal || '').trim();
  var barrio = String(data.barrio || '').trim();
  if (!label || !canal) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Faltan mes y canal' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var metaFact     = Number(data.metaFact) || 0;
  var metaPedidos  = Number(data.metaPedidos) || 0;
  var metaTicket   = (metaPedidos > 0) ? Math.round(metaFact / metaPedidos) : (Number(data.metaTicket) || 0);
  var metaClientes = Number(data.metaClientes) || 0;
  var notas        = String(data.notas || '').trim();
  var updated      = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');

  var lastRow = sh.getLastRow();
  var found = -1;
  if (lastRow > 1) {
    var rows = sh.getRange(2, 1, lastRow - 1, 3).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (_planCanonMes(rows[i][0]) === label &&
          String(rows[i][1]).trim() === canal &&
          String(rows[i][2]).trim() === barrio) {
        found = i + 2;
        break;
      }
    }
  }
  var values = [[label, canal, barrio, metaFact, metaPedidos, metaTicket, metaClientes, notas, updated]];
  if (found > 0) sh.getRange(found, 1, 1, 9).setValues(values);
  else sh.appendRow(values[0]);
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

function _planNextActionId(sh) {
  if (sh.getLastRow() <= 1) return 'A-001';
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  var maxN = 0;
  for (var i = 0; i < ids.length; i++) {
    var n = parseInt(String(ids[i][0]).replace(/\D/g, ''), 10);
    if (n > maxN) maxN = n;
  }
  return 'A-' + String(maxN + 1).padStart(3, '0');
}

// POST: agregar acción
function _doPostPlanAccionAdd(data) {
  var sheets = _planEnsureSheets();
  var sh = sheets.acciones;
  var label = _planCanonMes(data.mes) || String(data.mes || '').trim();
  var desc = String(data.desc || '').trim();
  if (!label || !desc) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Faltan mes y descripcion' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var id = _planNextActionId(sh);
  var fechaObj = String(data.fechaObjetivo || '').trim();
  var now = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');
  sh.appendRow([
    id, label,
    String(data.canal || '').trim(),
    String(data.barrio || '').trim(),
    desc,
    String(data.responsable || 'Tadeo').trim(),
    fechaObj,
    String(data.estado || 'Pendiente').trim(),
    Number(data.impactoEstimado) || 0,
    Number(data.impactoReal) || 0,
    String(data.notas || '').trim(),
    now, now
  ]);
  return ContentService.createTextOutput(JSON.stringify({ ok: true, id: id })).setMimeType(ContentService.MimeType.JSON);
}

// POST: actualizar acción
function _doPostPlanAccionUpdate(data) {
  var sheets = _planEnsureSheets();
  var sh = sheets.acciones;
  var id = String(data.id || '').trim();
  if (!id) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Falta id' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (sh.getLastRow() <= 1) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'No existe' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var all = sh.getDataRange().getValues();
  for (var r = 1; r < all.length; r++) {
    if (String(all[r][0]).trim() !== id) continue;
    var row = all[r];
    if (data.desc !== undefined)            row[4] = String(data.desc).trim();
    if (data.responsable !== undefined)     row[5] = String(data.responsable).trim();
    if (data.fechaObjetivo !== undefined)   row[6] = String(data.fechaObjetivo).trim();
    if (data.estado !== undefined)          row[7] = String(data.estado).trim();
    if (data.impactoEstimado !== undefined) row[8] = Number(data.impactoEstimado) || 0;
    if (data.impactoReal !== undefined)     row[9] = Number(data.impactoReal) || 0;
    if (data.notas !== undefined)           row[10] = String(data.notas).trim();
    if (data.canal !== undefined)           row[2] = String(data.canal).trim();
    if (data.barrio !== undefined)          row[3] = String(data.barrio).trim();
    row[12] = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');
    sh.getRange(r + 1, 1, 1, row.length).setValues([row]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'ID no encontrado' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// POST: borrar acción
function _doPostPlanAccionDelete(data) {
  var sheets = _planEnsureSheets();
  var sh = sheets.acciones;
  var id = String(data.id || '').trim();
  if (!id) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Falta id' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (sh.getLastRow() <= 1) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'No existe' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === id) {
      sh.deleteRow(i + 2);
      return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'ID no encontrado' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  Lista de repartidores: combina los detectados en col Repartidor
//  de Home/Pilar/Clubes + los de hoja Usuarios con rol relevante.
//  Usado por el selector de Mi Reparto en modo admin.
// ════════════════════════════════════════════════════════════
function _doGetRepartidoresList() {
  var users = {}; // {nombre: cantEntregas}
  var cfg = [
    { name: 'Home',   col: 55 },  // col BD
    { name: 'Pilar',  col: 58 },  // col BG
    { name: 'Clubes', col: 36 }   // col AK
  ];
  cfg.forEach(function(c) {
    var sh = SS.getSheetByName(c.name);
    if (!sh || sh.getLastRow() <= 1) return;
    if (sh.getLastColumn() <= c.col) return;
    var data = sh.getRange(2, c.col + 1, sh.getLastRow() - 1, 1).getValues();
    data.forEach(function(r) {
      var u = String(r[0] || '').trim();
      if (u) users[u] = (users[u] || 0) + 1;
    });
  });
  // Sumar metadata de hoja Usuarios (si existe)
  var meta = {};
  var shU = SS.getSheetByName('Usuarios');
  if (shU && shU.getLastRow() > 1) {
    var dU = shU.getDataRange().getValues();
    for (var r = 1; r < dU.length; r++) {
      var nombre = String(dU[r][3] || '').trim() || String(dU[r][0] || '').trim();
      if (!nombre) continue;
      meta[nombre] = {
        rol: String(dU[r][2] || '').trim().toLowerCase(),
        activo: String(dU[r][4] || '').trim().toLowerCase() !== 'no'
      };
    }
  }
  var lista = Object.keys(users).map(function(nombre) {
    var m = meta[nombre] || {};
    return {
      nombre: nombre,
      entregas: users[nombre],
      rol: m.rol || '',
      activo: m.activo !== false
    };
  });
  lista.sort(function(a, b) {
    if (b.entregas !== a.entregas) return b.entregas - a.entregas;
    return a.nombre.localeCompare(b.nombre);
  });
  return ContentService.createTextOutput(JSON.stringify({ ok: true, repartidores: lista }))
    .setMimeType(ContentService.MimeType.JSON);
}


// ════════════════════════════════════════════════════════════════════════
//  AJUSTES — Configuración editable desde el Panel (tab "Ajustes", solo admin)
//  Reemplaza la edición manual del Sheets. Hojas cubiertas:
//    Usuarios · Permisos · Config_Maleu · Provisiones_Fijas · Config · Vendedores
// ════════════════════════════════════════════════════════════════════════

// Tabs del Panel que un rol puede tener permiso de ver (para la matriz Permisos).
var _AJ_TABS_META = [
  { id:'inicio', label:'Inicio' }, { id:'ventas', label:'Ventas' },
  { id:'planificacion', label:'Planificación' }, { id:'pedidos', label:'Pedidos' },
  { id:'caja', label:'Caja' }, { id:'egresos', label:'Pagos' },
  { id:'stock', label:'Stock' }, { id:'ruta', label:'Ruta' },
  { id:'busqueda', label:'Abastecimiento' }, { id:'bbdd', label:'BBDD' },
  { id:'estancias', label:'Estancias' }, { id:'proveedores', label:'Proveedores' },
  { id:'miportal', label:'Mi Portal' }, { id:'pedidoshome', label:'Pedidos Home' },
  { id:'mireparto', label:'Mi Reparto' }, { id:'ajustes', label:'Ajustes' }
];
var _AJ_ROLES = ['admin','empleado','repartidor','vendedor'];

function _ajJson(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function _ajSiNo(v){ var s=String(v==null?'':v).trim().toLowerCase(); return s==='sí'||s==='si'||s==='true'||s==='1'||s==='yes'; }
// Parsea "$1.500.000" / "0,08" / "50" -> número. Mismo criterio que _resumenCfgNum.
function _ajNum(v){
  var s=String(v==null?'':v).replace(/[^0-9.,\-]/g,'').replace(/\./g,'').replace(',','.');
  var n=parseFloat(s); return isNaN(n)?0:n;
}

// Config_Maleu como array [{param,valor(number),desde,notas}] — se inyecta en D.config
// para que el EERR del Panel corra con valores VIVOS (no el fallback hardcodeado).
function _ajConfigMaleuArr(){
  var out=[];
  try{
    var sh=SS.getSheetByName('Config_Maleu');
    if(!sh||sh.getLastRow()<2) return out;
    var d=sh.getDataRange().getValues();
    for(var r=1;r<d.length;r++){
      var p=String(d[r][0]||'').trim(); if(!p) continue;
      out.push({ param:p, valor:_ajNum(d[r][1]), desde:String(d[r][2]||'').trim(), notas:String(d[r][3]||'') });
    }
  }catch(e){}
  return out;
}

// ── GET ajustesData — foto completa de todas las hojas de configuración ──
function _doGetAjustesData(){
  var out={ ok:true, tabsMeta:_AJ_TABS_META, roles:_AJ_ROLES };

  // Usuarios: Usuario·PIN·Rol·Nombre·Activo·Notas
  var usuarios=[];
  var shU=SS.getSheetByName('Usuarios');
  if(shU && shU.getLastRow()>1){
    var dU=shU.getRange(2,1,shU.getLastRow()-1,6).getValues();
    for(var i=0;i<dU.length;i++){
      var u=String(dU[i][0]||'').trim(); if(!u) continue;
      usuarios.push({ usuario:u, pin:String(dU[i][1]||''), rol:String(dU[i][2]||'').trim().toLowerCase()||'empleado',
        nombre:String(dU[i][3]||''), activo:_ajSiNo(dU[i][4]===''?'sí':dU[i][4]), notas:String(dU[i][5]||'') });
    }
  }
  out.usuarios=usuarios;

  // Permisos: matriz rol -> [tabs con acceso Sí]
  out.permisos=_getPermisos();

  // Config_Maleu (valor legible tal cual está en la hoja)
  var confM=[];
  var shC=SS.getSheetByName('Config_Maleu');
  if(shC && shC.getLastRow()>1){
    var dC=shC.getDataRange().getValues();
    for(var c=1;c<dC.length;c++){
      var pm=String(dC[c][0]||'').trim(); if(!pm) continue;
      confM.push({ param:pm, valor:String(dC[c][1]||''), desde:String(dC[c][2]||'').trim(), notas:String(dC[c][3]||'') });
    }
  }
  out.configMaleu=confM;

  // Provisiones_Fijas: Concepto·Categoría·Monto·Desde·Hasta·Notas
  var prov=[];
  var shP=SS.getSheetByName('Provisiones_Fijas');
  if(shP && shP.getLastRow()>1){
    var dP=shP.getRange(2,1,shP.getLastRow()-1,6).getValues();
    for(var p2=0;p2<dP.length;p2++){
      var cc=String(dP[p2][0]||'').trim(); if(!cc) continue;
      prov.push({ concepto:cc, categoria:String(dP[p2][1]||''), monto:String(dP[p2][2]||''),
        desde:String(dP[p2][3]||'').trim(), hasta:String(dP[p2][4]||'').trim(), notas:String(dP[p2][5]||'') });
    }
  }
  out.provisiones=prov;

  // Config: separo parámetros de negocio (editables) de contadores (peligrosos)
  var negocio=[], contadores=[];
  var shCfg=SS.getSheetByName('Config');
  if(shCfg && shCfg.getLastRow()>0){
    var dCfg=shCfg.getDataRange().getValues();
    for(var g=0;g<dCfg.length;g++){
      var k=String(dCfg[g][0]||'').trim();
      if(!k || k==='Parámetro' || k==='Parámetro Negocio') continue;
      var val=String(dCfg[g][1]||'');
      if(k.indexOf('Último')===0) contadores.push({ param:k, valor:val });
      else negocio.push({ param:k, valor:val });
    }
  }
  out.configNegocio=negocio;
  out.contadores=contadores;

  // Vendedores Red: Nombre·WA·AliasMP·Barrios·Partido·Localidad·Estado·Usuario·PIN·Comisión·Notas
  var vend=[];
  var shV=SS.getSheetByName('Vendedores');
  if(shV && shV.getLastRow()>1){
    var dV=shV.getRange(2,1,shV.getLastRow()-1,11).getValues();
    for(var v=0;v<dV.length;v++){
      var nm=String(dV[v][0]||'').trim(); if(!nm) continue;
      vend.push({ nombre:nm, wa:String(dV[v][1]||''), aliasMp:String(dV[v][2]||''), barrios:String(dV[v][3]||''),
        partido:String(dV[v][4]||''), localidad:String(dV[v][5]||''), estado:String(dV[v][6]||''),
        usuario:String(dV[v][7]||''), pin:String(dV[v][8]||''), comision:Number(dV[v][9])||17, notas:String(dV[v][10]||'') });
    }
  }
  out.vendedores=vend;

  return _ajJson(out);
}

// ── Usuarios ──
function _doPostUsuarioSet(data){
  var usuario=String(data.usuario||'').trim().toLowerCase();
  if(!usuario) return _ajJson({ ok:false, err:'Falta usuario' });
  var pin=String(data.pin||'').trim();
  var rol=String(data.rol||'empleado').trim().toLowerCase();
  var nombre=String(data.nombre||'').trim();
  var activo=(data.activo===false||String(data.activo).toLowerCase()==='no')?'No':'Sí';
  var notas=String(data.notas||'').trim();
  var orig=String(data.origUsuario||'').trim().toLowerCase();
  var sh=SS.getSheetByName('Usuarios');
  if(!sh){
    sh=SS.insertSheet('Usuarios');
    sh.appendRow(['Usuario','PIN','Rol','Nombre','Activo','Notas']);
    sh.getRange('A1:F1').setFontWeight('bold').setBackground('#331C1C').setFontColor('#fff');
  }
  var key=orig||usuario;
  var row=-1;
  if(sh.getLastRow()>1){
    var d=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<d.length;i++){ if(String(d[i][0]||'').trim().toLowerCase()===key){ row=i+2; break; } }
  }
  // No permitir duplicar usuario (alta nueva o rename hacia uno ya existente)
  if((row===-1||key!==usuario) && sh.getLastRow()>1){
    var d2=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var j=0;j<d2.length;j++){
      if(String(d2[j][0]||'').trim().toLowerCase()===usuario && (j+2)!==row) return _ajJson({ ok:false, err:'Ese usuario ya existe' });
    }
  }
  var vals=[usuario,pin,rol,nombre,activo,notas];
  if(row===-1) sh.appendRow(vals);
  else sh.getRange(row,1,1,6).setValues([vals]);
  return _ajJson({ ok:true, usuario:usuario });
}
function _doPostUsuarioDelete(data){
  var usuario=String(data.usuario||'').trim().toLowerCase();
  if(!usuario) return _ajJson({ ok:false, err:'Falta usuario' });
  if(usuario==='tadeo') return _ajJson({ ok:false, err:'No se puede borrar al admin principal' });
  var sh=SS.getSheetByName('Usuarios');
  if(!sh||sh.getLastRow()<2) return _ajJson({ ok:false, err:'Sin usuarios' });
  var d=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  for(var i=0;i<d.length;i++){
    if(String(d[i][0]||'').trim().toLowerCase()===usuario){ sh.deleteRow(i+2); return _ajJson({ ok:true }); }
  }
  return _ajJson({ ok:false, err:'No encontrado' });
}

// ── Permisos (matriz rol×tab) ──
function _doPostPermisoSet(data){
  var rol=String(data.rol||'').trim().toLowerCase();
  var tab=String(data.tab||'').trim().toLowerCase();
  if(!rol||!tab) return _ajJson({ ok:false, err:'Falta rol o tab' });
  var acceso=_ajSiNo(data.acceso)?'Sí':'No';
  _getPermisos(); // asegura que la hoja exista con defaults + migración
  var sh=SS.getSheetByName('Permisos');
  var row=-1;
  if(sh.getLastRow()>1){
    var d=sh.getRange(2,1,sh.getLastRow()-1,2).getValues();
    for(var i=0;i<d.length;i++){
      if(String(d[i][0]||'').trim().toLowerCase()===rol && String(d[i][1]||'').trim().toLowerCase()===tab){ row=i+2; break; }
    }
  }
  if(row===-1) sh.appendRow([rol,tab,acceso]);
  else sh.getRange(row,3).setValue(acceso);
  return _ajJson({ ok:true });
}

// ── Config_Maleu (parámetros con trazabilidad "Activo desde") ──
function _doPostConfigMaleuSet(data){
  var param=String(data.param||'').trim();
  if(!param) return _ajJson({ ok:false, err:'Falta parámetro' });
  var valor=String(data.valor==null?'':data.valor).trim();
  var desde=String(data.desde||'').trim();
  var notas=String(data.notas||'').trim();
  var sh=SS.getSheetByName('Config_Maleu');
  if(!sh){
    sh=SS.insertSheet('Config_Maleu');
    sh.appendRow(['Parámetro','Valor','Activo desde','Notas']);
    sh.getRange('A1:D1').setFontWeight('bold').setBackground('#331C1C').setFontColor('#fff');
  }
  // Upsert por (param + desde). Misma vigencia -> update; vigencia nueva -> fila nueva (histórico).
  var row=-1;
  if(sh.getLastRow()>1){
    var d=sh.getRange(2,1,sh.getLastRow()-1,4).getValues();
    for(var i=0;i<d.length;i++){
      if(String(d[i][0]||'').trim()===param && String(d[i][2]||'').trim()===desde){ row=i+2; break; }
    }
  }
  if(row===-1) sh.appendRow([param,valor,desde,notas]);
  else { sh.getRange(row,2).setValue(valor); if(notas) sh.getRange(row,4).setValue(notas); }
  return _ajJson({ ok:true });
}

// ── Provisiones_Fijas ──
function _doPostProvisionSet(data){
  var concepto=String(data.concepto||'').trim();
  if(!concepto) return _ajJson({ ok:false, err:'Falta concepto' });
  var categoria=String(data.categoria||'').trim();
  var monto=String(data.monto==null?'':data.monto).trim();
  var desde=String(data.desde||'').trim();
  var hasta=String(data.hasta||'').trim();
  var notas=String(data.notas||'').trim();
  var orig=String(data.origConcepto||'').trim();
  var sh=SS.getSheetByName('Provisiones_Fijas');
  if(!sh){
    sh=SS.insertSheet('Provisiones_Fijas');
    sh.appendRow(['Concepto','Categoría EERR','Monto Mensual','Activa desde','Activa hasta','Notas']);
    sh.getRange('A1:F1').setFontWeight('bold').setBackground('#331C1C').setFontColor('#fff');
  }
  var key=orig||concepto;
  var row=-1;
  if(sh.getLastRow()>1){
    var d=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<d.length;i++){ if(String(d[i][0]||'').trim()===key){ row=i+2; break; } }
  }
  var vals=[concepto,categoria,monto,desde,hasta,notas];
  if(row===-1) sh.appendRow(vals);
  else sh.getRange(row,1,1,6).setValues([vals]);
  return _ajJson({ ok:true });
}
function _doPostProvisionDelete(data){
  var concepto=String(data.concepto||'').trim();
  if(!concepto) return _ajJson({ ok:false, err:'Falta concepto' });
  var sh=SS.getSheetByName('Provisiones_Fijas');
  if(!sh||sh.getLastRow()<2) return _ajJson({ ok:false, err:'Sin provisiones' });
  var d=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  for(var i=0;i<d.length;i++){
    if(String(d[i][0]||'').trim()===concepto){ sh.deleteRow(i+2); return _ajJson({ ok:true }); }
  }
  return _ajJson({ ok:false, err:'No encontrado' });
}

// ── Config (parámetros de negocio: envío, umbral stock, contadores) ──
function _doPostConfigOpSet(data){
  var param=String(data.param||'').trim();
  if(!param) return _ajJson({ ok:false, err:'Falta parámetro' });
  var valor=data.valor;
  var sh=SS.getSheetByName('Config');
  if(!sh) return _ajJson({ ok:false, err:'No existe hoja Config' });
  var d=sh.getDataRange().getValues();
  for(var i=0;i<d.length;i++){
    if(String(d[i][0]||'').trim()===param){
      var num=_ajNum(valor);
      sh.getRange(i+1,2).setValue(isNaN(num)?valor:num);
      return _ajJson({ ok:true });
    }
  }
  return _ajJson({ ok:false, err:'Parámetro no encontrado' });
}

// ── Vendedores Red ──
function _doPostVendedorSet(data){
  var usuario=String(data.usuario||'').trim().toLowerCase();
  var nombre=String(data.nombre||'').trim();
  if(!usuario||!nombre) return _ajJson({ ok:false, err:'Falta nombre o usuario' });
  var sh=SS.getSheetByName('Vendedores');
  if(!sh) return _ajJson({ ok:false, err:'No existe hoja Vendedores' });
  var orig=String(data.origUsuario||'').trim().toLowerCase();
  var key=orig||usuario;
  var row=-1;
  if(sh.getLastRow()>1){
    var d=sh.getRange(2,8,sh.getLastRow()-1,1).getValues(); // col H = Usuario
    for(var i=0;i<d.length;i++){ if(String(d[i][0]||'').trim().toLowerCase()===key){ row=i+2; break; } }
  }
  var vals=[nombre, String(data.wa||''), String(data.aliasMp||''), String(data.barrios||''),
    String(data.partido||''), String(data.localidad||''), String(data.estado||'Activo'),
    usuario, String(data.pin||''), (Number(data.comision)||17), String(data.notas||'')];
  if(row===-1) sh.appendRow(vals);
  else sh.getRange(row,1,1,11).setValues([vals]);
  return _ajJson({ ok:true });
}
