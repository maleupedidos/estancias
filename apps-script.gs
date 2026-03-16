/**
 * ═══════════════════════════════════════════════════════════════
 *  MALEU — Apps Script v2 (con gestión de stock)
 *  Instrucciones de deploy:
 *    1. Abrir Google Sheets → Extensiones → Apps Script
 *    2. Reemplazar el contenido con este archivo
 *    3. Guardar → Implementar → Nueva implementación
 *       · Tipo: Aplicación web
 *       · Ejecutar como: Yo
 *       · Quién tiene acceso: Cualquier usuario
 *    4. Copiar la URL de implementación → pegar en APPS_SCRIPT_URL del HTML
 *
 *  Estructura del Spreadsheet (3 hojas):
 *
 *  Hoja "Productos":
 *    A: id | B: nombre | C: stock_fisico | D: reservado | E: stock_disponible | F: precio | G: activo
 *    La columna E es una fórmula: =C2-D2
 *    ⚠ Para que la página lea el stock, publicar ESTA hoja como CSV:
 *       Archivo → Compartir → Publicar en la web → Hoja "Productos" → CSV
 *       Copiar la URL → pegar en STOCK_CSV_URL del HTML
 *
 *  Hoja "Pedidos":
 *    A: id_pedido | B: fecha | C: nombre | D: barrio | E: lote | F: telefono
 *    G: dia_entrega | H: horario | I: pago | J: total | K: estado
 *    Estado inicial: "pendiente". Cambiar a mano: confirmado / entregado / cancelado
 *
 *  Hoja "Detalle_Pedidos":
 *    A: id_pedido | B: id_producto | C: nombre_producto | D: cantidad | E: precio_unitario
 * ═══════════════════════════════════════════════════════════════
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();

// ─── doPost: recibe un pedido nuevo desde la página ───────────────────────────
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // espera hasta 10s si hay otro pedido simultáneo

  try {
    const data = JSON.parse(e.postData.contents);

    const idPedido  = 'P-' + Date.now();
    const fecha     = data.fecha     || new Date().toLocaleString('es-AR');
    const nombre    = data.nombre    || '';
    const barrio    = data.barrio    || '';
    const lote      = data.lote      || '';
    const telefono  = data.telefono  || '';
    const dia       = data.dia       || '';
    const horario   = data.horario   || '';
    const pago      = data.pago      || '';
    const total     = data.total     || 0;
    const items     = data.items     || []; // array de { id, nombre, qty, precio }

    // 1. Escribir fila en "Pedidos"
    const hPedidos = SS.getSheetByName('Pedidos');
    hPedidos.appendRow([
      idPedido, fecha, nombre, barrio, lote, telefono,
      dia, horario, pago, total, 'pendiente'
    ]);

    // 2. Escribir filas en "Detalle_Pedidos" y reservar stock
    const hDetalle  = SS.getSheetByName('Detalle_Pedidos');
    const hProductos = SS.getSheetByName('Productos');
    const prodData  = hProductos.getDataRange().getValues(); // [header, row, row, ...]

    items.forEach(item => {
      // Detalle
      hDetalle.appendRow([idPedido, item.id, item.nombre, item.qty, item.precio]);

      // Reservar stock: incrementar columna D (reservado) en la fila del producto
      for (let r = 1; r < prodData.length; r++) {
        if (prodData[r][0] === item.id) {
          const cell = hProductos.getRange(r + 1, 4); // columna D = reservado
          cell.setValue((cell.getValue() || 0) + item.qty);
          break;
        }
      }
    });

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, id: idPedido }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ─── onEdit: detecta cambios de estado y actualiza stock ──────────────────────
// IMPORTANTE: Instalar como trigger "On Edit" en el editor de Apps Script:
//   Activadores (reloj) → Agregar activador → onEdit → Desde hoja de cálculo → Al editar
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'Pedidos') return;

  const col = e.range.getColumn();
  const colEstado = 11; // columna K = estado
  if (col !== colEstado) return;

  const nuevoEstado  = e.value;
  const estadoAnterior = e.oldValue;

  // Solo actuar cuando el estado cambia a "entregado" o "cancelado"
  if (nuevoEstado !== 'entregado' && nuevoEstado !== 'cancelado') return;

  const idPedido  = sheet.getRange(e.range.getRow(), 1).getValue();
  const hDetalle  = SS.getSheetByName('Detalle_Pedidos');
  const hProductos = SS.getSheetByName('Productos');

  const detalleData = hDetalle.getDataRange().getValues();
  const prodData    = hProductos.getDataRange().getValues();

  // Buscar todas las líneas de este pedido
  detalleData.slice(1).forEach(row => {
    if (row[0] !== idPedido) return;
    const idProd = row[1];
    const qty    = row[3];

    for (let r = 1; r < prodData.length; r++) {
      if (prodData[r][0] === idProd) {
        const celdaReservado = hProductos.getRange(r + 1, 4); // col D = reservado
        const celdaFisico    = hProductos.getRange(r + 1, 3); // col C = stock_fisico

        const reservadoActual = celdaReservado.getValue() || 0;
        const fisicoActual    = celdaFisico.getValue() || 0;

        if (nuevoEstado === 'entregado') {
          // Descontar del stock físico y liberar la reserva
          celdaFisico.setValue(Math.max(0, fisicoActual - qty));
          celdaReservado.setValue(Math.max(0, reservadoActual - qty));
        } else if (nuevoEstado === 'cancelado') {
          // Solo liberar la reserva (el stock físico no se toca)
          celdaReservado.setValue(Math.max(0, reservadoActual - qty));
        }
        break;
      }
    }
  });
}

// ─── Compatibilidad: el doPost anterior enviaba el campo "productos" como string ─
// Esta función parsea ese formato si los items no vienen en el array estructurado.
function _parseProductosString(str) {
  // Formato: "Pizza Margarita ×2 | Sorrentinos ×1"
  // No tenemos ids en este formato, así que no podemos actualizar stock.
  // Migrar al nuevo formato enviando data.items desde la página.
  return [];
}
