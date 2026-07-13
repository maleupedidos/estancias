"""Reescribe _doGetBusqueda con 1 sola pasada + sin TZ conversion por fila."""
src_path = r'C:\Users\tadeu\estancias\apps-script.gs'

NEW_FN = r"""function _doGetBusqueda() {
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
    provsArr.push({ n: prov, costo: d.costo, cats: catsArr, wa: waLines.join('\n') });
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
}"""

with open(src_path, 'r', encoding='utf-8') as f:
    src = f.read()

start_marker = 'function _doGetBusqueda() {'
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
