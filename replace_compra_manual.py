"""Replace _doPostCompraManual with optimized version."""
src_path = r'C:\Users\tadeu\estancias\apps-script.gs'

NEW_FN = r"""function _doPostCompraManual(postData) {
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
}"""

with open(src_path, 'r', encoding='utf-8') as f:
    src = f.read()

start_marker = 'function _doPostCompraManual(postData) {'
i = src.find(start_marker)
if i < 0:
    raise SystemExit('No encontré la función')
j = i + len(start_marker)
depth = 1
while j < len(src) and depth:
    c = src[j]
    if c == '{': depth += 1
    elif c == '}': depth -= 1
    j += 1

new_src = src[:i] + NEW_FN + src[j:]

with open(src_path, 'w', encoding='utf-8') as f:
    f.write(new_src)

print(f'OK. Reemplazada función ({j-i} chars -> {len(NEW_FN)} chars)')
