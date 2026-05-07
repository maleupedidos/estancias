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
  if (action === 'busqueda') return _doGetBusqueda();
  if (action === 'catalogo') return _doGetCatalogo();
  if (action === 'admin') return _doGetAdmin();
  if (action === 'ventas') return _doGetVentas();
  if (action === 'stock') return _doGetStock();
  if (action === 'precios') return _doGetPrecios();
  if (action === 'cobrosPendientes') return _doGetCobrosPendientes();
  if (action === 'resumenSemanal') return _doGetResumenSemanal(e);
  if (action === 'crmClientes') return _doGetCrmClientes();
  if (action === 'crmCliente') return _doGetCrmCliente(e);
  if (action === 'crmProductos') return _doGetCrmProductos();
  if (action === 'crmProducto') return _doGetCrmProducto(e);
  return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
}

// Cobros pendientes: pedidos Entregados + No Cobrados de TODAS las hojas
// Red: agrupado por Vendedor (aplica 17% de comision)
function _doGetCobrosPendientes() {
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
        desc: cfg.colDesc !== null ? (Number(data[r][cfg.colDesc]) || 0) : 0
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
    var idxEstPagoMaleu = -1, idxFpMaleu = -1;
    for (var hh = 0; hh < hdrRed.length; hh++) {
      var nm = String(hdrRed[hh]).trim();
      if (nm === 'Estado Pago a Maleu') idxEstPagoMaleu = hh;
      else if (nm === 'Forma Pago a Maleu') idxFpMaleu = hh;
    }
    var porVendedor = {};
    for (var r = 1; r < dRed.length; r++) {
      var estado = String(dRed[r][11] || '').trim(); // col 12 Estado Entrega
      // Antes solo entraban Entregados. Ahora cuenta también Pendiente/Reservado
      // — el monto total que el vendedor va a deber esta semana, así Tadeo lo ve
      // en Ruta → Cobros incluso antes de que el vendedor entregue.
      if (estado === 'Cancelado') continue;
      var estPagoMaleu = idxEstPagoMaleu >= 0 ? String(dRed[r][idxEstPagoMaleu] || '').trim() : '';
      if (estPagoMaleu === 'Pagado' || estPagoMaleu === 'Sí' || estPagoMaleu === 'Si') continue;
      var vendedor = String(dRed[r][7] || '').trim(); // col 8 Vendedor
      if (!vendedor) continue;
      var pedId = dRed[r][1];
      if (!pedId) continue;
      var total = Number(dRed[r][14]) || 0; // col 15 Total
      var fechaV = dRed[r][3];
      var fechaStr = fechaV instanceof Date
        ? Utilities.formatDate(fechaV, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy')
        : String(fechaV || '').trim();
      if (!porVendedor[vendedor]) {
        porVendedor[vendedor] = {
          vendedor: vendedor,
          ids: [],
          totalBruto: 0,
          fechas: [],
          tel: '',
          fps: {}
        };
      }
      var v = porVendedor[vendedor];
      v.ids.push(String(pedId));
      v.totalBruto += total;
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
      var comision = Math.round(v.totalBruto * (comPct / 100));
      var neto = v.totalBruto - comision;
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

  return ContentService.createTextOutput(JSON.stringify({ts: Date.now(), cobros: out}))
    .setMimeType(ContentService.MimeType.JSON);
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
        oc: ocItems
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
        oc: ocItemsC
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
      var estadoR = String(redData[rr][11]).trim();          // L = Estado entrega Marcos→cliente (informativo)
      if (estadoR === 'Cancelado') continue;
      var entVend = String(redData[rr][56] || '').trim();    // BE = Entregado a Vendedor (Tadeo→Marcos)
      if (entVend === 'Entregado') continue;                 // ya se lo entregamos a Marcos, no aparece en Armado

      var diaR = _fechaADiaSemana(redData[rr][10]) || ''; // K = Día de Entrega
      var feISOr = _fechaAISO(redData[rr][10]);
      if (dia && diaR !== dia) continue;

      // Origen Detalle Red idx 55 (col BD)
      var oDr = {};
      var rawODr = String(redData[rr][55] || '').trim();
      if (rawODr) { try { oDr = JSON.parse(rawODr) || {}; } catch(eR) { oDr = {}; } }

      // Solo items que salen del Depósito Maleu (origen = "D")
      var prodsR = [];
      for (var pr = 0; pr < ABBRS_RED.length; pr++) {
        var qR = Number(redData[rr][21 + pr]) || 0;       // V..AR = idx 21..43
        if (qR <= 0) continue;
        if (oDr[ABBRS_RED[pr]] !== 'D') continue;          // Solo D — el resto va por OC y se entrega directo
        prodsR.push({ a: ABBRS_RED[pr], q: qR });
      }
      if (prodsR.length === 0) continue;                   // Si no hay nada del depósito, no hace falta armar

      var vendedor = String(redData[rr][7] || '').trim();
      var clienteFinal = String(redData[rr][8] || '').trim();

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
        retira: vendedor                                  // Indicador: lo retira el vendedor, no lo entrega Tadeo
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

  // O (15) = Total, AS (45) = Costo
  sh.getRange(row, 15).setValue(totalNew);
  sh.getRange(row, 45).setValue(costoTotal);

  // Q (17) Efectivo / R (18) Transferencia segun forma de pago vigente
  var totalSinEnvio = totalNew - envio;
  if (formaPago === 'Efectivo') {
    sh.getRange(row, 17).setValue(totalSinEnvio);
    sh.getRange(row, 18).setValue(0);
  } else if (formaPago === 'Transferencia') {
    sh.getRange(row, 17).setValue(0);
    sh.getRange(row, 18).setValue(totalSinEnvio);
  }
  // Mixto: como solo se setea al cobrar y editar requiere no-cobrado, no debería caer acá.

  // Garantizar fórmulas vivas (por si la fila vino sin ellas)
  if (!sh.getRange(row, 21).getFormula()) sh.getRange(row, 21).setFormula('=O' + row + '+S' + row + '+T' + row);
  if (!sh.getRange(row, 46).getFormula()) sh.getRange(row, 46).setFormula('=U' + row + '-AS' + row);
  if (!sh.getRange(row, 47).getFormula()) sh.getRange(row, 47).setFormula('=U' + row + '*17/100');
  if (!sh.getRange(row, 48).getFormula()) sh.getRange(row, 48).setFormula('=AT' + row + '-AU' + row);
  if (!sh.getRange(row, 49).getFormula()) sh.getRange(row, 49).setFormula('=U' + row + '*83/100');

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
    var data = sh.getRange(2, 1, numRows, 55).getValues();
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
        totalFacturado += facturado; totalPedidos++;
        if (year === yearActual && semana === semanaActual) {
          semanaFacturado += facturado; semanaPedidos++;
          if (diaEntrega === 'Viernes') {
            viernesHoy++;
            if (estado === 'Entregado') viernesEntregados++;
          }
        }
        if (year === yearActual && mes === nombreMesActual) {
          mesFacturado += facturado; mesPedidos++;
        }
        // Acumulado por (año, mes) para desglose. Solo si tenemos fechaEntrega válida.
        if (fechaEntrega && mes) {
          var mNum = fechaEntrega.getMonth() + 1; // 1-12
          var mKey = year + '-' + (mNum < 10 ? '0' + mNum : mNum);
          if (!mesesMap[mKey]) mesesMap[mKey] = { n: mes, y: year, m: mNum, facturado: 0, pedidos: 0 };
          mesesMap[mKey].facturado += facturado;
          mesesMap[mKey].pedidos++;
        }
        // Acumulado por (año, semana ISO) para desglose semanal
        if (fechaEntrega && semana) {
          var wKey = year + '-' + (semana < 10 ? '0' + semana : semana);
          if (!semanasMap[wKey]) semanasMap[wKey] = { y: year, w: semana, facturado: 0, pedidos: 0 };
          semanasMap[wKey].facturado += facturado;
          semanasMap[wKey].pedidos++;
        }
        // Plata pendiente de cobrar al cliente (entregado pero no cobrado)
        if (estado === 'Entregado' && estadoPago !== 'Cobrado') {
          pendCobrar += facturado;
        }
        // Plata cobrada pero no liquidada a Maleu (cobrado y no pagado a Maleu)
        // Marcos se queda con comisión 17%, el resto (83%) es para Maleu
        // Acepta "Pagado" (nuevo) o "Sí" (formato viejo) como liquidado
        var liquidado = (estadoPagoMaleu === 'Pagado' || estadoPagoMaleu === 'Sí' || estadoPagoMaleu === 'Si');
        if (estadoPago === 'Cobrado' && !liquidado) {
          pendLiquidar += aPagarMaleu;
        }
      }

      // Productos del pedido (cols V-AR = 21-43)
      var prods = [];
      var ABBRS = ['PPM','PPJyQ','PPCyQ','SQB','SL','SCo','SPyP','SJyQ','SE','SCa',
                   'ECaC','EJyQ','ECyQ','EV','TG','TLC','TC','F','PMu','PMa','PJyQ','PCC','PJyM'];
      for (var p = 0; p < 23; p++) {
        var qty = Number(data[r][21 + p]) || 0;
        if (qty > 0) prods.push({ a: ABBRS[p], q: qty });
      }

      // Calcular si es "este viernes" (mismo año+semana y día=Viernes)
      var esViernesActual = (year === yearActual && semana === semanaActual && diaEntrega === 'Viernes');

      pedidos.push({
        n: nPedido, c: cliente, f: fechaStr, de: diaEntrega,
        $: facturado, com: comision, aPg: aPagarMaleu, es: estado,
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
      var bucket = semanasMap[key] || { facturado: 0, pedidos: 0 };
      res.push({
        n: wNum, y: wYear,
        lun: fmt(mon), dom: fmt(dom),
        facturado: bucket.facturado,
        pedidos: bucket.pedidos,
        comision: Math.round(bucket.facturado * 17 / 100)
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
                  comision: Math.round(semanaFacturado * 17 / 100) },
      mes:      { facturado: mesFacturado, pedidos: mesPedidos, nombre: nombreMesActual,
                  y: yearActual, m: mesActual,
                  comision: Math.round(mesFacturado * 17 / 100),
                  semanas: _semanasDelMes(yearActual, mesActual) },
      mesesAnt: mesesAnt,
      total:    { facturado: totalFacturado,  pedidos: totalPedidos },
      comision: Math.round(totalFacturado * 17 / 100),
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

  // Una sola lectura del rango usado (sin getDataRange que sobre-lee columnas vacías)
  var lastRow = shOC.getLastRow();
  var data = shOC.getRange(1, 1, lastRow, 25).getValues();

  function splitProducto(nombre) {
    var parts = nombre.split(' — ');
    return { cat: (parts[0] || nombre).trim(), var: (parts[1] || '').trim() };
  }

  // Semana ISO actual (lun-dom) en zona horaria AR — para filtrar OCs vivas
  var _nowArg = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var semanaActualBusq = _isoWeek(_nowArg);
  var anioActualBusq = _nowArg.getFullYear();

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

    var costoT = (Number(row[13]) || 0) * qty;

    // ── DEUDAS: Recibido + no-Depósito + Pagado != Sí ──
    if (esRecibido && !esDeposito) {
      var pagado = String(row[23]).trim();
      if (pagado !== 'Sí' && pagado !== 'Si' && costoT > 0) {
        var semD = String(row[2]).trim();
        if (!deudas[prov]) deudas[prov] = { total: 0, semanas: {}, rows: [] };
        deudas[prov].total += costoT;
        if (!deudas[prov].semanas[semD]) deudas[prov].semanas[semD] = 0;
        deudas[prov].semanas[semD] += costoT;
        deudas[prov].rows.push({ r: r + 1, prod: String(row[10]).trim(), q: qty, costo: costoT, sem: semD });
      }
    }

    // ── Filas activas para porProv/porCliente/ocRows: solo no-Depósito ──
    if (esDeposito) continue;

    // Filtro de semana actual: usar col C (Semana) que ya es número, en lugar de TZ-convert por fila
    var semCol = Number(row[2]) || 0;
    var fCre = row[1];
    var anioFila = (fCre instanceof Date) ? fCre.getFullYear() : anioActualBusq;
    if (semCol !== semanaActualBusq || anioFila !== anioActualBusq) continue;

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

    // Proveedor agrupado (solo Pendiente/Pedido)
    if (!esRecibido) {
      if (!porProv[prov]) porProv[prov] = { costo: 0, cats: {} };
      var sp = splitProducto(producto);
      var arr = porProv[prov].cats[sp.cat] || (porProv[prov].cats[sp.cat] = []);
      var found = false;
      for (var v = 0; v < arr.length; v++) {
        if (arr[v].v === sp.var) { arr[v].q += qty; arr[v].ct += costoT; found = true; break; }
      }
      if (!found) arr.push({ v: sp.var, q: qty, ct: costoT });
      porProv[prov].costo += costoT;
      totalGeneral += costoT;
    }

    // Cliente agrupado
    var cKey = cliente + '|' + nPedido;
    if (!porCliente[cKey]) porCliente[cKey] = { n: cliente, canal: canal, dir: dir, tel: tel, ped: nPedido, items: [], ing: 0 };
    porCliente[cKey].items.push({ prod: producto, q: qty, pv: ingresoT, est: estado });
    porCliente[cKey].ing += ingresoT;

    ocRows.push({ oc: ocNum, r: r + 1, prov: prov, prod: producto, abbr: abbr, q: qty, est: estado, canal: canal, cliente: cliente, nped: nPedido });
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

  // Armar provsArr
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
    provsArr.push({ n: prov, costo: d.costo, cats: catsArr, wa: waLines.join('\n'), fProg: fProgMap[prov] || '' });
  });

  var clientesArr = Object.keys(porCliente).map(function(k) { return porCliente[k]; });

  // ── Pagos Proveedores (ledger) ──
  var pagosImputados = {};
  var pagosLibres = {};
  var shPagos = SS.getSheetByName('Pagos Proveedores');
  if (shPagos && shPagos.getLastRow() > 1) {
    var pagosData = shPagos.getRange(1, 1, shPagos.getLastRow(), 7).getValues();
    for (var rp = 1; rp < pagosData.length; rp++) {
      var pp = String(pagosData[rp][1]).trim();
      if (!pp) continue;
      var ef = Number(pagosData[rp][2]) || 0;
      var mp = Number(pagosData[rp][3]) || 0;
      var tot = Number(pagosData[rp][4]) || (ef + mp);
      if (tot <= 0) continue;
      var semImp = String(pagosData[rp][5] || '').trim();
      var notas = String(pagosData[rp][6] || '').trim();
      var fechaRaw = pagosData[rp][0];
      var fechaStr = (fechaRaw instanceof Date)
        ? Utilities.formatDate(fechaRaw, 'America/Argentina/Buenos_Aires', 'dd/MM HH:mm')
        : String(fechaRaw || '').trim();
      var pagoObj = { fecha: fechaStr, ef: ef, mp: mp, tot: tot, notas: notas };
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
      semsArr.push({ sem: s, original: bruto, pagado: pagadoTotalSem, pendiente: pendiente, pagosImp: pagosSem, pagadoFifo: aplicarLibre });
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
  var notas = String(postData.notas || '').trim();
  var semanaImp = String(postData.semana || '').trim();
  var totalPago = montoEf + montoMp;

  if (!proveedor || totalPago <= 0) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'Sin monto' })).setMimeType(ContentService.MimeType.JSON);
  }

  var ahora = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var semana = _isoWeek(argNow);
  var mes = MESES[argNow.getMonth()];

  // ── Ledger: hoja "Pagos Proveedores" ──
  // Layout: A Fecha | B Proveedor | C Efectivo | D MP | E Total | F Semana Imputada | G Notas
  var shPag = SS.getSheetByName('Pagos Proveedores');
  if (!shPag) {
    shPag = SS.insertSheet('Pagos Proveedores');
    shPag.getRange(1, 1, 1, 7).setValues([['Fecha', 'Proveedor', 'Efectivo', 'Mercado Pago', 'Total', 'Semana Imputada', 'Notas']]);
    shPag.setFrozenRows(1);
  } else if (shPag.getLastColumn() < 7) {
    // Asegurar que existan las 7 columnas (migración suave si la hoja se creó con 6)
    shPag.getRange(1, 6, 1, 2).setValues([['Semana Imputada', 'Notas']]);
  }
  shPag.appendRow([argNow, proveedor, montoEf, montoMp, totalPago, semanaImp, notas]);
  var rPag = shPag.getLastRow();
  shPag.getRange(rPag, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  shPag.getRange(rPag, 3, 1, 3).setNumberFormat('$#,##0');

  // ── Egresos (impacto en caja) ──
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
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
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

  function isTadeoStock(cliente) {
    var s = String(cliente || '').toLowerCase().replace(/\s+/g, ' ').trim();
    return s.indexOf('tadeo') === 0 && s.indexOf('stock') !== -1;
  }

  function setRowQty(row, newQty) {
    var costoUnit = Number(shOC.getRange(row, 14).getValue()) || 0;
    var precioV   = Number(shOC.getRange(row, 16).getValue()) || 0;
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

  function sumarStock(abbr, qty, refOC, canal, origen) {
    if (!abbr || qty <= 0 || !hProd) return;
    if (canal.indexOf('Dep') !== 0) return;
    if (origen !== 'Orden de Compra') return;
    var prodData = hProd.getDataRange().getValues();
    for (var rp = 1; rp < prodData.length; rp++) {
      if (String(prodData[rp][2]).trim() === abbr) {
        var celdaFis = hProd.getRange(rp + 1, 6);
        var fisico = Number(celdaFis.getValue()) || 0;
        var nuevo = fisico + qty;
        celdaFis.setValue(nuevo);
        _logKardex(abbr, '+REC', qty, fisico, nuevo, 'OC', refOC);
        SpreadsheetApp.flush();
        return;
      }
    }
  }

  function cargarInfo(row) {
    return {
      row: row,
      canal:   String(shOC.getRange(row, 5).getValue()).trim(),
      cliente: String(shOC.getRange(row, 7).getValue()).trim(),
      abbr:    String(shOC.getRange(row, 12).getValue()).trim(),
      qty:     Number(shOC.getRange(row, 13).getValue()) || 0,
      origen:  String(shOC.getRange(row, 20).getValue()).trim(),
      estado:  String(shOC.getRange(row, 21).getValue()).trim(),
      refOC:   String(shOC.getRange(row, 1).getValue() || '')
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
    infos.forEach(function(i){ i.isStock = isTadeoStock(i.cliente); });
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
  return ContentService.createTextOutput(JSON.stringify({ ok: true, updated: updated, avisos: avisos })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
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
  var vendedor = String(postData.vendedor || 'Tadeo — Stock').trim();
  var items = postData.items || [];
  var esMarcos = vendedor === 'Marcos Bottcher';

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
  var canalLabel = esMarcos ? 'Red' : 'Depósito';
  var dirLabel = esMarcos ? 'Tortugas, Garín' : 'Depósito Maleu';

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

    // Escribir propina si existe
    var propina = Number(data.propina) || 0;
    if (propina > 0) {
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
    if (data.action === 'marcarCobrado')   return _doPostMarcarCobrado(data);
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
    if (data.action === 'crmMergeClientes')     return _doPostCrmMergeClientes(data);
    if (data.action === 'programarBusqueda')    return _doPostProgramarBusqueda(data);
    if (data.action === 'marcarEntregadoAVendedor') return _doPostMarcarEntregadoAVendedor(data);
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
    const canal = String(data.canal || 'Home');
    if (canal === 'Clubes')          _doPostClubes(data);
    else if (canal === 'Red')        _doPostRed(data);
    else if (canal === 'Pilar')      _doPostHome(data, 'Pilar', 'P');
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
// Home v2 abr/2026: productos empiezan en W (23) hasta AO (41)
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
};

// Mapeo col → abreviatura para Pilar (inverso de PILAR_PRODUCT_COLS vía PAGE_ID_TO_ABBR)
const PILAR_COL_TO_ABBR = {
  23:'PPM', 24:'PPJyQ', 25:'PPCyQ',
  26:'SQB', 27:'SL', 28:'SCo', 29:'SPyP', 30:'SJyQ', 31:'SE', 32:'SCa',
  33:'ECaC', 34:'EJyQ', 35:'ECyQ', 36:'EV',
  37:'TG', 38:'TLC', 39:'TC', 40:'F',
  41:'PMu', 42:'PMa', 43:'PJyQ', 44:'PCC', 45:'PJyM',
};

function _doPostHome(data, sheetName, prefix) {
  sheetName = sheetName || 'Home';
  prefix = prefix || 'H';
  const sh = SS.getSheetByName(sheetName);
  if (!sh) return;

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

  // ── SALVAGUARDA DESCUENTO (server-side) ──
  // El frontend a veces manda descuento=0 cuando debería haber aplicado el 10%
  // (ej: cache viejo, race condition al confirmar). El backend valida y corrige.
  // Reglas:
  //   • Home: 10% OFF si subtotalProd >= 100k  Ó  pago = Efectivo (siempre activo)
  //   • Pilar: idem (la tienda aplica solo en "Otro barrio" pero acá llegan justamente esos casos)
  //   • Clubes: sin descuento
  if (sheetName !== 'Clubes') {
    const aplicaBulk = subtotalProd >= 100000;
    const aplicaCash = (pago === 'Efectivo');
    const descEsperado = (aplicaBulk || aplicaCash) ? Math.round(subtotalProd * 0.10) : 0;
    if (descEsperado > 0 && descuento === 0) {
      // El cliente no aplicó el descuento que correspondía. Corregir.
      descuento = descEsperado;
      total = subtotalProd + envio - descuento;
    }
  }

  const efectivo        = pago === 'Efectivo'      ? total : 0;
  const transferencia   = pago === 'Transferencia' ? total : 0;

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
  row[17] = efectivo;                           // R  Efectivo
  row[18] = transferencia;                      // S  Transferencia
  // Propina: el frontend (PWA Ruta tab "+") puede mandar data.propina. Va a la col T si pago=Efectivo,
  // a la col U si pago=Transferencia. Si pago=Mixto u otro, no se asume nada.
  var propinaNueva = Number(data.propina) || 0;
  row[19] = (pago === 'Efectivo'      && propinaNueva > 0) ? propinaNueva : 0;  // T  Propina Efectivo
  row[20] = (pago === 'Transferencia' && propinaNueva > 0) ? propinaNueva : 0;  // U  Propina Transferencia
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

  // Fórmula Facturado (V=22): V = Q (Total a cobrar) + T (Propina Ef) + U (Propina Tr)
  // Refleja la venta realizada (no depende de si ya se cobró). La caja usa R/S aparte.
  sh.getRange(newRow, 22).setFormula('=Q' + newRow + '+T' + newRow + '+U' + newRow);
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

  // Sync WATI (silencioso, no bloquea)
  _syncWatiContact_(isPilar ? 'Pilar' : 'Home', data.telefono, isPilar ? '' : subBarrio, isPilar ? (data.barrio || data.direccion) : barrioPrivado);
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
  row[20] = (pago === 'Efectivo'      && propinaC > 0) ? propinaC : 0;  // U  Propina Efectivo
  row[21] = (pago === 'Transferencia' && propinaC > 0) ? propinaC : 0;  // V  Propina Transferencia
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

  // Sync WATI (silencioso)
  _syncWatiContact_('Clubes', data.telefono, '', '');
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
  sh.getRange(newRow, 21).setFormula('=O' + newRow + '+S' + newRow + '+T' + newRow);
  // Fórmula Margen Bruto en AT (col 46) = Facturado - Costo
  sh.getRange(newRow, 46).setFormula('=U' + newRow + '-AS' + newRow);
  // Fórmula Comisión 17% en AU (col 47) = Facturado * 17/100
  sh.getRange(newRow, 47).setFormula('=U' + newRow + '*17/100');
  // Fórmula Margen Neto en AV (col 48) = Margen Bruto - Comisión
  sh.getRange(newRow, 48).setFormula('=AT' + newRow + '-AU' + newRow);
  // Fórmula A Pagar en AW (col 49) = Facturado * 83/100 (lo que queda para Maleu)
  sh.getRange(newRow, 49).setFormula('=U' + newRow + '*83/100');
  // Forzar teléfono como texto (col AZ = 52)
  var telRed = String(data.telefono || '');
  if (telRed) sh.getRange(newRow, 52).setNumberFormat('@').setValue(telRed);

  // Sync WATI (silencioso)
  _syncWatiContact_('Red', data.telefono, '', '');
}

// Número de semana ISO (lunes = primer día de la semana)
function _isoWeek(date) {
  const d      = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
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

  // → Cancelado: N° pasa a "-"
  if (nuevo === 'Cancelado' && anterior !== 'Cancelado') {
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

  var cantidades = shRed.getRange(row, 22, 1, 23).getValues()[0];
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

    // → Cancelado: N° pasa a "-"
    if (nuevo === 'Cancelado' && anterior !== 'Cancelado') {
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

  var cantidades = shClubes.getRange(row, 24, 1, 8).getValues()[0];
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
    'PMu':7000, 'PMa':7000, 'PJyQ':7000, 'PCC':7000, 'PJyM':7800,
    'PPM':11000, 'PPJyQ':11000, 'PPCyQ':11000
  };

  const colToAbbrMap = (canal === 'Clubes')
    ? {24:'PMu', 25:'PMa', 26:'PJyQ', 27:'PCC', 28:'PJyM', 29:'PPM', 30:'PPJyQ', 31:'PPCyQ'}
    : (canal === 'Red') ? RED_COL_TO_ABBR
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

  // → Cancelado: N° pasa a "-" (no cuenta en la numeración semanal)
  if (nuevo === 'Cancelado' && anterior !== 'Cancelado') {
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
  const cantidades = shHome.getRange(row, 23, 1, prodCount).getValues()[0];
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
  var cantidades = shHome.getRange(row, 23, 1, prodCountMx).getValues()[0];
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
};

// Mapeo Red: abreviatura → letra de columna en Red (productos van de V(22) a AR(44))
const ABBR_TO_RED_COL = {
  'PPM':'V', 'PPJyQ':'W', 'PPCyQ':'X',
  'SQB':'Y', 'SL':'Z', 'SCo':'AA', 'SPyP':'AB', 'SJyQ':'AC', 'SE':'AD', 'SCa':'AE',
  'ECaC':'AF', 'EJyQ':'AG', 'ECyQ':'AH', 'EV':'AI',
  'TG':'AJ', 'TLC':'AK', 'TC':'AL', 'F':'AM',
  'PMu':'AN', 'PMa':'AO', 'PJyQ':'AP', 'PCC':'AQ', 'PJyM':'AR',
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

    // Col L (Check D-E) = Stock Inicial - Vendidos
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
    .addItem('Actualizar fórmulas Productos', 'setupProductosFormulas')
    .addItem('Reset stock semanal', 'resetStockSemanal')
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
      esMarcos ? 'Red' : 'Deposito',             // E  Canal
      '',                                         // F  N° Pedido Origen (no aplica)
      vendedor || 'Tadeo — Stock',                // G  Cliente
      '',                                         // H  Teléfono
      esMarcos ? 'Tortugas, Garín' : 'Deposito Maleu', // I  Dirección
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
function _doGetAdmin() {
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
  var saldoBase = { ef: 0, mp: 0, fecha: '', fechaDate: null };
  var shSaldoSnap = SS.getSheetByName('Saldo Base');
  if (shSaldoSnap && shSaldoSnap.getLastRow() > 1) {
    var lastRowSB = shSaldoSnap.getLastRow();
    saldoBase.ef = Number(shSaldoSnap.getRange(lastRowSB, 2).getValue()) || 0;
    saldoBase.mp = Number(shSaldoSnap.getRange(lastRowSB, 3).getValue()) || 0;
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

  // Home, Pilar
  ['Home', 'Pilar'].forEach(function(hoja) {
    var sh = SS.getSheetByName(hoja);
    var stats = { nombre: hoja, pedidos: 0, entregados: 0, pendientes: 0, cancelados: 0, reservados: 0, facturado: 0, cobrado: 0, noCobrado: 0 };
    if (!sh || sh.getLastRow() <= 1) { canales.push(stats); return; }
    var data = sh.getDataRange().getValues();
    var headersH = data[0];

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

      pedidos.push({
        n: nPedido, h: hoja, c: cliente, f: fechaStr, de: diaEntrega, dee: diaEntregaISO,
        es: estado, o: origen, ep: estadoPago, fp: formaPago,
        $: facturado, co: costoPed, mg: margenPed, br: subBarrio, p: prods,
        hr: horaStr, dia: diaPedido,
        ef: efR, tr: trR, pef: pefR, ptr: ptrR,
        fe: feStr, fed: feDiaStr,
        fc: fcStr, fcd: fcDiaStr,
        r: r + 1
      });
    }
    canales.push(stats);
  });

  // Clubes (34 cols, estructura diferente)
  var shClubes = SS.getSheetByName('Clubes');
  var statsClubes = { nombre: 'Clubes', pedidos: 0, entregados: 0, pendientes: 0, cancelados: 0, reservados: 0, facturado: 0, cobrado: 0, noCobrado: 0 };
  if (shClubes && shClubes.getLastRow() > 1) {
    var dataClubes = shClubes.getDataRange().getValues();
    var headersClub = dataClubes[0];
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

      pedidos.push({
        n: nPedidoC, h: 'Clubes',
        c: clienteC + (clubC ? ' (' + clubC + ')' : ''),
        f: fechaStrC, de: diaEntregaC, dee: diaEntregaISOC, es: estadoC, o: origenC,
        ep: estadoPagoC, fp: formaPagoC, $: facturadoC,
        co: costoCl, mg: margenCl, br: clubC, p: prodsC,
        hr: horaStrC, dia: diaPedC,
        ef: efC, tr: trC, pef: pefC, ptr: ptrC,
        fc: fcStrC, fcd: fcDiaStrC,
        r: rc + 1
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
    var dataRed = shRedDash.getDataRange().getValues();
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

      statsRed.facturado += aPagarR;
      if (estadoPagoR === 'Cobrado') statsRed.cobrado += aPagarR;
      else statsRed.noCobrado += aPagarR;

      var prodsR = [];
      for (var prr = 0; prr < 23; prr++) {
        var qtyR = Number(dataRed[rr][21 + prr]) || 0;
        if (qtyR > 0) prodsR.push({ a: ABBRS_RED_DASH[prr], q: qtyR });
      }

      var formaPagoR = String(dataRed[rr][12] || '').trim();
      var costoR = Number(dataRed[rr][44]) || 0;
      var margenR = Number(dataRed[rr][45]) || 0;

      var efRd  = Number(dataRed[rr][16]) || 0; // Q = Efectivo
      var trRd  = Number(dataRed[rr][17]) || 0; // R = Transferencia
      var pefRd = Number(dataRed[rr][18]) || 0; // S = Propina Ef
      var ptrRd = Number(dataRed[rr][19]) || 0; // T = Propina Trans

      var horaRd = dataRed[rr][0]; // A = Hora
      var horaStrRd = horaRd instanceof Date ? Utilities.formatDate(horaRd, 'America/Argentina/Buenos_Aires', 'HH:mm') : String(horaRd || '');
      var diaPedRd = String(dataRed[rr][2] || '').trim(); // C = Día

      pedidos.push({
        n: nPedidoR, h: 'Red',
        c: clienteR + (vendedorR ? ' (Red: ' + vendedorR + ')' : ''),
        f: fechaStrR, de: diaEntregaR, dee: diaEntregaISOR, es: estadoR, o: origenR,
        ep: estadoPagoR, fp: formaPagoR, $: aPagarR,
        co: costoR, mg: margenR, br: vendedorR, p: prodsR,
        hr: horaStrRd, dia: diaPedRd,
        ef: efRd, tr: trRd, pef: pefRd, ptr: ptrRd,
        r: rr + 1
      });
    }
  }
  canales.push(statsRed);

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
    var idxSemana = -1, idxAnio = -1;
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
      var deudaP = totP * (1 - comPct/100);
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
      // Precio y costo (cols I=8, J=9 0-based) — usados por el editor de pedidos para defaults
      var precio = parseFloat(String(dataProd[rp][8] || '').replace(/[$.]/g,'').replace(/,/g,'')) || 0;
      var costoU = parseFloat(String(dataProd[rp][9] || '').replace(/[$.]/g,'').replace(/,/g,'')) || 0;
      stock.push({ n: nombre, a: abbr, f: fisico, r: reservado, d: disponible, p: precio, co: costoU });
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
          abbr: String(dataOC[ro][11] || '').trim(),
          q: Number(dataOC[ro][12]) || 0,
          proveedor: String(dataOC[ro][9] || '').trim(),
          estado: estOC
        });
      }
    }
  }

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
        else mpTotal = netoCierre;
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
      movimientos: movimientos
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
              : (hoja === 'Clubes') ? { 24:'PMu',25:'PMa',26:'PJyQ',27:'PCC',28:'PJyM',29:'PPM',30:'PPJyQ',31:'PPCyQ' }
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
  sh.getRange(row, colOrigen).setValue(summary);

  // Guardar detalle JSON en columna "Origen Detalle".
  // Formato nuevo: { abbr: { d: N, oc: Y } }   (mantiene compat de lectura con regex de fórmulas)
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
  prodsNorm.forEach(function(p){ detailObj[p.a] = { d: p.d, oc: p.oc }; });
  sh.getRange(row, colDetalle).setValue(JSON.stringify(detailObj));

  // Generar OC para los productos con cantOC > 0 (con cantidad explícita)
  var ocProds = prodsNorm.filter(function(p){ return p.oc > 0; });
  if (ocProds.length > 0) {
    var clienteOrig = String(sh.getRange(row, hoja === 'Red' ? 9 : 8).getValue() || '').trim();
    var shOC = SS.getSheetByName('Orden de Compra');
    var yaExiste = false;
    if (shOC && shOC.getLastRow() > 1) {
      var existentes = shOC.getRange(2, 5, shOC.getLastRow() - 1, 3).getValues();
      for (var i = 0; i < existentes.length; i++) {
        if (String(existentes[i][0]).trim() === hoja &&
            String(existentes[i][1]).trim() === pedidoId &&
            String(existentes[i][2]).trim() === clienteOrig) { yaExiste = true; break; }
      }
    }
    if (!yaExiste) {
      _generarOCSelectiva(hoja, row, ocProds.map(function(p){ return { a: p.a, qty: p.oc, prov: p.prov }; }));
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true, origen: summary })).setMimeType(ContentService.MimeType.JSON);
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
    COL_MAP = { 24:'PMu',25:'PMa',26:'PJyQ',27:'PCC',28:'PJyM',29:'PPM',30:'PPJyQ',31:'PPCyQ' };
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
    colEst = 11; colOrigen = 9; colPago = 13;
    colTotal = 14;
    colCosto = isPilar ? 46 : 42;
    colMargen = isPilar ? 47 : 43;
    COL_MAP = isPilar ? PILAR_COL_TO_ABBR : HOME_COL_TO_ABBR;
    prodStartCol = 23;
    prodEndCol = isPilar ? 45 : 41;
    colEnvio = 15; // O = Envio
    colFacturado = 22;
    colDescuento = isPilar ? 53 : 53; // BA = Descuento (Home y Pilar)
  }

  var est = String(sh.getRange(row, colEst).getValue()).trim();
  var pago = String(sh.getRange(row, colPago).getValue()).trim();
  if (est === 'Entregado' || est === 'Cancelado') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'No se puede editar: el pedido está ' + est })).setMimeType(ContentService.MimeType.JSON);
  }
  if (pago === 'Cobrado') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: 'No se puede editar: el pedido ya está Cobrado' })).setMimeType(ContentService.MimeType.JSON);
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

  var origenReseteado = false;
  if (hadOrigen && !canPreserve) {
    // RESET completo: borrar TODAS las OCs y poner Pendiente
    if (shOC) {
      var rowsToDelete = ocsDelPedido.map(function(o){return o.rowOC;}).sort(function(a,b){return b-a;});
      rowsToDelete.forEach(function(rr){ shOC.deleteRow(rr); });
    }
    sh.getRange(row, colOrigen).setValue('Pendiente');
    if (colOrigenDetalle) sh.getRange(row, colOrigenDetalle).setValue('');
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
  }

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
  var CLUBES_PRECIOS = {'PMu':7000,'PMa':7000,'PJyQ':7000,'PCC':7000,'PJyM':7800,'PPM':11000,'PPJyQ':11000,'PPCyQ':11000};

  // Reescribir cantidades en TODAS las columnas de productos del bloque
  var nProds = prodEndCol - prodStartCol + 1;
  var newQty = new Array(nProds).fill(0);
  var subtotal = 0, costo = 0;
  Object.keys(COL_MAP).forEach(function(colStr){
    var colIdx = Number(colStr);
    var abbr = COL_MAP[colIdx];
    var info = lineasMap[abbr];
    if (!info || info.q <= 0) {
      newQty[colIdx - prodStartCol] = 0;
      return;
    }
    var precio = info.precio != null ? info.precio
               : (hoja === 'Clubes') ? (CLUBES_PRECIOS[abbr] || 0)
               : (precioRetail[abbr] || 0);
    var costoLinea = info.costo != null ? info.costo : (costoUnit[abbr] || 0);
    newQty[colIdx - prodStartCol] = info.q;
    subtotal += info.q * precio;
    costo    += info.q * costoLinea;
  });

  // Escribir cantidades en batch
  sh.getRange(row, prodStartCol, 1, nProds).setValues([newQty]);

  // Recalcular Total y Facturado
  var envio = Number(sh.getRange(row, colEnvio).getValue()) || 0;
  var descuento = colDescuento ? (Number(sh.getRange(row, colDescuento).getValue()) || 0) : 0;
  var totalNuevo = subtotal + envio - descuento;
  sh.getRange(row, colTotal).setValue(totalNuevo);
  sh.getRange(row, colCosto).setValue(costo);
  sh.getRange(row, colMargen).setValue(totalNuevo - costo);

  SpreadsheetApp.flush();
  return ContentService.createTextOutput(JSON.stringify({
    ok: true, total: totalNuevo, costo: costo, margen: totalNuevo - costo,
    origenReseteado: origenReseteado
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

  var CLUBES_PRECIOS = {'PMu':7000,'PMa':7000,'PJyQ':7000,'PCC':7000,'PJyM':7800,'PPM':11000,'PPJyQ':11000,'PPCyQ':11000};
  var colToAbbrMap = (canal === 'Clubes')
    ? {24:'PMu',25:'PMa',26:'PJyQ',27:'PCC',28:'PJyM',29:'PPM',30:'PPJyQ',31:'PPCyQ'}
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
      'Orden de Compra', 'Pedido', fechaStr, '', 'No', 'No'
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

/** POST action=marcarCobrado — marca un pedido como Cobrado desde el Panel.
 *  { action:'marcarCobrado', hoja:'Clubes', id:'C-005' } */
// Cobrar todos los pedidos de un Vendedor Red
function _doPostCobrarVendedorRed(data) {
  var vendedor = String(data.vendedor || '').trim();
  if (!vendedor) return ContentService.createTextOutput(JSON.stringify({ok:false,err:'vendedor vacio'})).setMimeType(ContentService.MimeType.JSON);
  var sh = SS.getSheetByName('Red');
  if (!sh) return ContentService.createTextOutput(JSON.stringify({ok:false,err:'hoja'})).setMimeType(ContentService.MimeType.JSON);
  var d = sh.getDataRange().getValues();
  var updated = 0;
  for (var r = 1; r < d.length; r++) {
    var vend = String(d[r][7] || '').trim();      // col 8 Vendedor
    var estado = String(d[r][11] || '').trim();   // col 12 Estado Entrega
    var pago = String(d[r][13] || '').trim();     // col 14 Estado Pago
    if (vend !== vendedor) continue;
    if (estado !== 'Entregado') continue;
    if (pago === 'Cobrado') continue;
    sh.getRange(r+1, 14).setValue('Cobrado');     // col 14 = N
    updated++;
  }
  return ContentService.createTextOutput(JSON.stringify({ok:true, updated:updated})).setMimeType(ContentService.MimeType.JSON);
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

  // Marcar Cobrado + estampar fecha de cobro
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
      var nuevoDesc  = (formaPago === 'Efectivo') ? Math.round(subtotal * 0.10) : 0;
      var nuevoTotal = subtotal + envio - nuevoDesc;
      sh.getRange(row, cols.desc).setValue(nuevoDesc);
      sh.getRange(row, cols.total).setValue(nuevoTotal);
      totalRecalculado = nuevoTotal;
    }
  }

  // Modelo: R (Efectivo) y S (Transferencia) = parte del TOTAL A COBRAR por método (sin propina).
  // T/U = propina por método (plus sobre el total). La col Facturado (V) = Q+T+U.
  // El cliente (PWA Ruta) envía ef/tr YA NETOS y propina aparte. Escribimos tal cual.
  var propina = Number(data.propina) || 0;
  var propMet = String(data.propMet || '').trim();
  var ef = Number(data.ef) || 0;
  var tr = Number(data.tr) || 0;
  // Si hubo recálculo y no es Mixto, forzar ef/tr al nuevo total según fp
  if (totalRecalculado !== null && formaPago === 'Efectivo')      { ef = totalRecalculado; tr = 0; }
  else if (totalRecalculado !== null && formaPago === 'Transferencia') { tr = totalRecalculado; ef = 0; }
  if (ef > 0 || tr > 0) {
    sh.getRange(row, cols.ef).setValue(ef);
    sh.getRange(row, cols.tr).setValue(tr);
  }
  if (propina > 0 && propMet) {
    var colProp = propMet === 'Efectivo' ? cols.propEf : cols.propTr;
    sh.getRange(row, colProp).setValue(propina);
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true, total: totalRecalculado })).setMimeType(ContentService.MimeType.JSON);
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
 *  Así, cobros históricos ya reflejados en el monto ajustado no cuentan dos veces. */
function _doPostAjusteSaldo(data) {
  var deseadoEf = Number(data.efectivo) || 0;
  var deseadoMP = Number(data.mp) || 0;

  var shSaldo = SS.getSheetByName('Saldo Base');
  if (!shSaldo) {
    shSaldo = SS.insertSheet('Saldo Base');
    shSaldo.getRange(1, 1, 1, 3).setValues([['Fecha', 'Efectivo', 'Mercado Pago']]);
    shSaldo.setFrozenRows(1);
    shSaldo.getRange(1, 1, 1, 3).setBackground(BROWN).setFontColor('#FFFFFF').setFontWeight('bold');
  }
  var ahora = new Date();
  var argNow = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  shSaldo.appendRow([argNow, deseadoEf, deseadoMP]);
  shSaldo.getRange(shSaldo.getLastRow(), 1).setNumberFormat('dd/MM/yyyy HH:mm');
  // Sin flush(): UI optimista en frontend.

  return ContentService.createTextOutput(JSON.stringify({ ok: true, ef: deseadoEf, mp: deseadoMP })).setMimeType(ContentService.MimeType.JSON);
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

  // ── Helper: parsear fila VD (Home, Pilar) — v2 abr/2026 ──
  // Home v2: Facturado V(21), Ef R(17), Tr S(18), Costo AP(41), Margen AQ(42), FechaEnt AX(49), MesEnt AY(50), SemEnt AZ(51)
  // Pilar v2: Facturado V(21), Ef R(17), Tr S(18), Costo AT(45), Margen AU(46), FechaEnt BA(52), MesEnt BB(53), SemEnt BC(54)
  function parseVD(data, r, zona) {
    var cliente = String(data[r][7] || '').trim();
    if (!cliente) return null;
    var isP = (zona === 'Pilar');
    var IDX = isP
      ? { fact:21, totalAlt:16, ef:17, tr:18, costo:45, margen:46, fechaEnt:52, mesEnt:53, semEnt:54 }
      : { fact:21, totalAlt:16, ef:17, tr:18, costo:41, margen:42, fechaEnt:49, mesEnt:50, semEnt:51 };
    var facturado = Number(data[r][IDX.fact]) || Number(data[r][IDX.totalAlt]) || 0;
    if (facturado === 0) return null;
    var fechaEnt = data[r][IDX.fechaEnt];
    var fecha = fechaEnt instanceof Date ? fechaEnt : data[r][3];
    var fechaStr = fecha instanceof Date ? Utilities.formatDate(fecha, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : String(fecha || '');
    var mesEnt = String(data[r][IDX.mesEnt] || '').trim();
    var semEnt = Number(data[r][IDX.semEnt]) || 0;
    var mesPed = String(data[r][4] || '').trim();
    var semPed = Number(data[r][5]) || 0;
    var mes = MVAL.indexOf(mesEnt) >= 0 ? mesEnt : mesPed;
    var sem = semEnt > 0 ? semEnt : semPed;
    return {
      canal: 'Venta Directa', zona: zona, fecha: fechaStr, mes: mes, sem: sem,
      cliente: cliente, estado: String(data[r][10] || '').trim(),
      fp: String(data[r][11] || '').trim(), ep: String(data[r][12] || '').trim(),
      $: facturado, ef: Number(data[r][IDX.ef]) || 0, tr: Number(data[r][IDX.tr]) || 0,
      costo: Number(data[r][IDX.costo]) || 0, margen: Number(data[r][IDX.margen]) || 0
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
      var fR = dRd[rr][3];
      var fRS = fR instanceof Date ? Utilities.formatDate(fR, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : String(fR || '');
      ventas.push({
        canal: 'Red', zona: String(dRd[rr][7] || '').trim(), fecha: fRS,
        mes: String(dRd[rr][4] || '').trim(), sem: Number(dRd[rr][5]) || 0,
        cliente: cliR, estado: estadoR,
        fp: String(dRd[rr][12] || '').trim(),
        ep: String(dRd[rr][13] || '').trim(),
        $: aPagarR,
        ef: Number(dRd[rr][16]) || 0, tr: Number(dRd[rr][17]) || 0,
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

  // ── Catering (hoja operativa manual) ──
  var shCt = SS.getSheetByName('Catering');
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
  var metaSemanal = _resumenMeta(canales.semana);

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
    meta: metaSemanal
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
  return [
    { canal:'Home', sheet:'Home',
      cFecha:3, cCliente:7, cOrigen:8, cEstEnt:10, cEstPago:12,
      cTotal:16, cEf:17, cTr:18, cPropEf:19, cPropTr:20, cFacturado:21,
      cCosto:41, cBarrio:43, cSubBarrio:44, cLote:45, cTel:46,
      cFechaCobro:54, prodStart:22, prodEnd:40 },
    { canal:'Pilar', sheet:'Pilar',
      cFecha:3, cCliente:7, cOrigen:8, cEstEnt:10, cEstPago:12,
      cTotal:16, cEf:17, cTr:18, cPropEf:19, cPropTr:20, cFacturado:21,
      cCosto:45, cBarrio:47, cSubBarrio:47, cLote:48, cTel:49,
      cFechaCobro:57, prodStart:22, prodEnd:44 },
    { canal:'Clubes', sheet:'Clubes',
      cFecha:3, cCliente:7, cClub:8, cDeporte:9, cGrupo:10, cOrigen:11, cEstEnt:13, cEstPago:15,
      cTotal:16, cEf:18, cTr:19, cPropEf:20, cPropTr:21, cFacturado:22,
      cCosto:31, cTel:33, cFechaCobro:34, prodStart:23, prodEnd:30 },
    { canal:'Red', sheet:'Red',
      cFecha:3, cVendedor:7, cCliente:8, cOrigen:9, cEstEnt:11, cEstPago:13,
      cTotal:14, cEf:16, cTr:17, cPropEf:18, cPropTr:19, cFacturado:20,
      cCosto:44, cBarrio:49, cLote:50, cTel:51,
      cFechaCobro:54, prodStart:21, prodEnd:43 }
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

  return {
    unicos: unicosArr.length,
    nuevosCount: nuevos.length,
    nuevos: nuevos,
    recurrentes: unicosArr.length - nuevos.length,
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

function _resumenMeta(filas) {
  var entregados = 0;
  filas.forEach(function(f){
    if (f.canal === 'Red') return;
    if (f.estadoEntrega === 'Entregado') entregados++;
  });
  return {
    objetivo: 50,
    alcanzado: entregados,
    pct: Math.round((entregados / 50) * 100),
    detalle: 'Pedidos entregados Home + Pilar + Clubes'
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
function _crmHojasConfig() {
  return [
    {
      name: 'Home', cliente: 7, dia: 9, fecha: 3, est: 10, pago: 12, fp: 11,
      total: 21, costo: 41, barrio: 43, subBarrio: 44, domicilio: 45, tel: 46,
      prodStart: 22, prodEnd: 40, fechaCobro: 54
    },
    {
      name: 'Pilar', cliente: 7, dia: 9, fecha: 3, est: 10, pago: 12, fp: 11,
      total: 21, costo: 45, barrio: 47, subBarrio: -1, domicilio: 48, tel: 49,
      prodStart: 22, prodEnd: 44, fechaCobro: 57
    },
    {
      name: 'Clubes', cliente: 7, dia: 12, fecha: 3, est: 13, pago: 15, fp: 14,
      total: 22, costo: 31, barrio: -1, subBarrio: -1, domicilio: -1, tel: 33,
      prodStart: 23, prodEnd: 30, fechaCobro: 34, club: 8, deporte: 9, grupo: 10
    },
    {
      name: 'Red', cliente: 8, dia: 10, fecha: 3, est: 11, pago: 13, fp: 12,
      total: 20, costo: 44, barrio: 49, subBarrio: -1, domicilio: 50, tel: 51,
      prodStart: 21, prodEnd: 43, vendedor: 7
    }
  ];
}

// Lee headers de productos de cada hoja para mapear columna → abreviatura.
function _crmProductHeaders(sh, prodStart, prodEnd) {
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var map = {};
  for (var i = prodStart; i <= prodEnd && i < headers.length; i++) {
    var abrev = String(headers[i] || '').trim();
    if (abrev) map[i] = abrev;
  }
  return map;
}

// Lee meta-data manual de clientes (cumple, notas, alias) desde hoja "Clientes Meta".
// La hoja se crea automáticamente si no existe.
function _crmGetClientesMeta() {
  var sh = SS.getSheetByName('Clientes Meta');
  if (!sh) {
    sh = SS.insertSheet('Clientes Meta');
    var hdr = ['Tel Normalizado', 'Nombre Canónico', 'Cumpleaños (DD/MM)', 'Alias MP', 'Notas', 'Tags', 'Updated'];
    sh.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold').setBackground('#331C1C').setFontColor('#F2E8C7');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, hdr.length, 140);
    return {};
  }
  if (sh.getLastRow() <= 1) return {};
  var data = sh.getDataRange().getValues();
  var map = {};
  for (var r = 1; r < data.length; r++) {
    var tel = String(data[r][0] || '').trim();
    if (!tel) continue;
    map[tel] = {
      tel: tel,
      nombreCanonico: String(data[r][1] || '').trim(),
      cumple: String(data[r][2] || '').trim(),
      aliasMp: String(data[r][3] || '').trim(),
      notas: String(data[r][4] || '').trim(),
      tags: String(data[r][5] || '').trim(),
      updated: data[r][6] || '',
      _row: r + 1
    };
  }
  return map;
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
    var prodMap = _crmProductHeaders(sh, cfg.prodStart, cfg.prodEnd);
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
          lastFecha: null
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
      if (estado === 'Entregado') {
        c.countEntregados++;
        c.totalFacturado += total;
        if (pago === 'Cobrado') c.totalCobrado += total;
        else c.deuda += total;
      }
    }
  });

  // Mergear sin-tel en index si hay match por nombre+barrio único
  Object.keys(sinTel).forEach(function(k) { index[k] = sinTel[k]; });

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
function _doGetCrmClientes() {
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
    var canalDom = _crmTopValue(c.canales);
    var nombresAlt = Object.keys(c.nombres).filter(function(n) { return n !== nombreRep; });
    var telRep = c.tel || _crmTopValue(c.telefonos);

    lista.push({
      key: key,
      tel: c.tel,
      telDisplay: telRep,
      nombre: nombreRep,
      nombresAlt: nombresAlt, // para detectar tipeos: Marcos / Marcps
      barrio: barrioRep,
      subBarrio: subRep,
      club: clubRep,
      canalDom: canalDom,
      canales: Object.keys(c.canales),
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
      estado: k.estado,
      vip: k.vip,
      cumple: m ? m.cumple : '',
      notas: m ? m.notas : '',
      tags: m ? m.tags : '',
      aliasMp: m ? m.aliasMp : ''
    });
  });
  // Ordenar por última compra desc
  lista.sort(function(a, b) {
    var fa = a.ultimaFecha ? a.ultimaFecha.split('/').reverse().join('') : '0';
    var fb = b.ultimaFecha ? b.ultimaFecha.split('/').reverse().join('') : '0';
    return fb.localeCompare(fa);
  });
  return ContentService.createTextOutput(JSON.stringify({ts: Date.now(), clientes: lista}))
    .setMimeType(ContentService.MimeType.JSON);
}

// Endpoint: ficha completa de un cliente por key (telNorm o NOMBRE:xxx).
function _doGetCrmCliente(e) {
  var key = e && e.parameter && e.parameter.key;
  if (!key) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'falta key'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
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

  return ContentService.createTextOutput(JSON.stringify({
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
  })).setMimeType(ContentService.MimeType.JSON);
}

// Endpoint: lista de productos enriquecida con ventas históricas.
function _doGetCrmProductos() {
  var shP = SS.getSheetByName('Productos');
  if (!shP) return ContentService.createTextOutput(JSON.stringify({productos: []})).setMimeType(ContentService.MimeType.JSON);

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
      _clientes: {}
    });
  }
  var byAbrev = {};
  prods.forEach(function(p) { byAbrev[p.abrev] = p; });

  var hoy = new Date();
  var hace30 = new Date(hoy.getTime() - 30 * 24 * 60 * 60 * 1000);
  var hace7 = new Date(hoy.getTime() - 7 * 24 * 60 * 60 * 1000);

  var hojas = _crmHojasConfig();
  hojas.forEach(function(cfg) {
    var sh = SS.getSheetByName(cfg.name);
    if (!sh || sh.getLastRow() <= 1) return;
    var prodMap = _crmProductHeaders(sh, cfg.prodStart, cfg.prodEnd);
    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var estado = String(row[cfg.est] || '').trim();
      if (estado !== 'Entregado') continue;
      var fecha = _crmToDate(row[cfg.fecha]);
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
        if (fecha && fecha >= hace30) p.vendidosMes += pr.cant;
        if (fecha && fecha >= hace7) p.vendidosSemana += pr.cant;
        var monto = totalProds > 0 ? Math.round((pr.cant / totalProds) * total) : 0;
        p.facturadoTotal += monto;
        if (fecha && fecha >= hace30) p.facturadoMes += monto;
        p.pedidosCount++;
        if (nombreCli) p._clientes[nombreCli] = (p._clientes[nombreCli] || 0) + pr.cant;
        if (fecha && (!p.ultimaVenta || fecha > p.ultimaVenta)) p.ultimaVenta = fecha;
      });
    }
  });

  prods.forEach(function(p) {
    p.clientesUnicos = Object.keys(p._clientes).length;
    p.ultimaVenta = p.ultimaVenta ? Utilities.formatDate(p.ultimaVenta, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : '';
    delete p._clientes;
  });

  return ContentService.createTextOutput(JSON.stringify({ts: Date.now(), productos: prods}))
    .setMimeType(ContentService.MimeType.JSON);
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
    var prodMap = _crmProductHeaders(sh, cfg.prodStart, cfg.prodEnd);
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

  return ContentService.createTextOutput(JSON.stringify({
    ts: Date.now(),
    producto: prodInfo,
    topClientes: topClientes,
    ventasPorMes: ventasPorMes,
    ventasPorCanal: ventasPorCanal,
    pedidosRecientes: pedidosRecientes,
    clientesUnicos: Object.keys(clientes).length
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
  var updated = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');

  if (found > 0) {
    sh.getRange(found, 1, 1, 7).setValues([[tel, nombreCanonico, cumple, aliasMp, notas, tags, updated]]);
  } else {
    sh.appendRow([tel, nombreCanonico, cumple, aliasMp, notas, tags, updated]);
  }
  return ContentService.createTextOutput(JSON.stringify({ok: true, tel: tel}))
    .setMimeType(ContentService.MimeType.JSON);
}

// POST: fusionar dos "clientes" (típicamente cuando son la misma persona con tipeo distinto)
// Estrategia: anota la fusión en Clientes Meta como "alias" del telNorm canónico.
// (El backend ya unifica por teléfono normalizado; esto cubre el caso de personas sin teléfono.)
function _doPostCrmMergeClientes(data) {
  var keyA = String(data.keyCanonical || '').trim();
  var keyB = String(data.keyAlias || '').trim();
  if (!keyA || !keyB) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'faltan keys'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // Por ahora solo registramos la intención en Clientes Meta como nota.
  // En futuras iteraciones se puede aplicar el merge en las hojas operativas.
  var sh = SS.getSheetByName('Clientes Meta');
  if (!sh) return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'sin meta'}))
    .setMimeType(ContentService.MimeType.JSON);
  return ContentService.createTextOutput(JSON.stringify({ok: true, msg: 'merge anotado (fase 2)'}))
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
