// ============================================================
// 11_ImportarComercial.gs  —  Importar datos comerciales
//
// Estructura real del archivo _26_Comercial___mi_eelo.xlsx:
//
// Hoja "Sales":
//   Fila 1-2: meta/encabezados de sección
//   Fila 3:   Encabezados reales →
//             Mes | INVENT | Cliente | Invoice | FAC |
//             Numero de OP|OM | Ventas | Status de Producción |
//             Tipo de Venta | Importe ($) | Importe (Q) |
//             Shipping | Impuestos | Total | Moneda |
//             Qt | Unit Price | Código | Conceptos...
//
// Hoja "Compras":
//   Fila 1: "Reporte de Gastos Mi eelo 2026"
//   (revisar manualmente las columnas reales)
// ============================================================

// Importa ventas desde la hoja "Sales" (pégala en este Sheets primero).
// Copia la hoja "Sales " del archivo Comercial a este Spreadsheet con ese mismo nombre.
function importarVentas() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrig = ss.getSheetByName("Sales") || ss.getSheetByName("Sales ");
  var hojaDest = ss.getSheetByName(CONFIG.HOJAS.VENTAS);

  if (!hojaOrig) {
    SpreadsheetApp.getUi().alert(
      "❌ No existe la hoja 'Sales'.\n" +
      "Copia la hoja del archivo Comercial a este Spreadsheet."
    );
    return;
  }

  var datos = hojaOrig.getDataRange().getValues();
  var importadas = 0;

  // Los encabezados reales están en fila 3 (índice 2)
  // Los datos empiezan en fila 4 (índice 3)
  for (var i = 3; i < datos.length; i++) {
    var fila = datos[i];
    var mes    = fila[0];
    var cliente = fila[2];
    if (!cliente && !mes) continue;

    hojaDest.appendRow([
      mes,      // A: Mes
      fila[2],  // B: Cliente (col C en original)
      fila[3],  // C: Invoice
      fila[4],  // D: FAC
      fila[5],  // E: Numero_OP_OM
      fila[6],  // F: Ventas
      fila[7],  // G: Status_Produccion
      fila[8],  // H: Tipo_Venta
      fila[9],  // I: Importe_USD
      fila[10], // J: Importe_GTQ
      fila[11], // K: Shipping
      fila[12], // L: Impuestos
      fila[13], // M: Total
      fila[14], // N: Moneda
      fila[15], // O: Cantidad
      fila[16], // P: Precio_Unitario
      fila[17], // Q: Codigo
      fila[18], // R: Conceptos
    ]);
    importadas++;
  }

  SpreadsheetApp.getUi().alert("✅ " + importadas + " ventas importadas.");
}

// Importa compras desde la hoja "Compras".
function importarCompras() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrig = ss.getSheetByName("Compras");
  var hojaDest = ss.getSheetByName(CONFIG.HOJAS.COMPRAS);

  if (!hojaOrig) {
    SpreadsheetApp.getUi().alert("❌ No existe la hoja 'Compras'.");
    return;
  }

  var datos = hojaOrig.getDataRange().getValues();
  // Fila 1 es "Reporte de Gastos Mi eelo 2026" (título)
  // Detecta automáticamente la fila de encabezados (la que tenga más valores)
  var filaEncabezado = _detectarFilaEncabezado(datos);
  var importadas = 0;

  for (var i = filaEncabezado + 1; i < datos.length; i++) {
    var fila = datos[i];
    if (!fila[0] && !fila[1]) continue; // fila vacía

    hojaDest.appendRow([
      fila[0], // A: Fecha
      fila[1], // B: Descripcion
      fila[2], // C: Proveedor
      fila[3], // D: Monto
      fila[4], // E: Tipo_Compra
      fila[5], // F: Factura_DTE
      fila[6], // G: Pago_Realizado
      fila[7], // H: Notas
    ]);
    importadas++;
  }

  SpreadsheetApp.getUi().alert("✅ " + importadas + " compras importadas.");
}

// Genera un resumen de ventas por mes en el Dashboard.
function resumenVentasPorMes() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var hojaVent = ss.getSheetByName(CONFIG.HOJAS.VENTAS);
  if (!hojaVent) return {};

  var datos  = hojaVent.getDataRange().getValues();
  var totales = {};

  for (var i = 1; i < datos.length; i++) {
    var mes   = String(datos[i][0]).trim();
    var total = parseFloat(datos[i][9]) || 0; // Importe_GTQ
    if (!mes) continue;
    totales[mes] = (totales[mes] || 0) + total;
  }
  return totales;
}

function _detectarFilaEncabezado(datos) {
  var maxValores = 0, filaMax = 0;
  for (var i = 0; i < Math.min(datos.length, 5); i++) {
    var count = datos[i].filter(function(v){ return v !== null && v !== ""; }).length;
    if (count > maxValores) { maxValores = count; filaMax = i; }
  }
  return filaMax;
}
