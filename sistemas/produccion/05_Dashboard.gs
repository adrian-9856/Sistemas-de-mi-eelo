// ============================================================
// 05_Dashboard.gs — Métricas de producción en tiempo real
// ============================================================

function actualizarDashboard() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(CFG.HOJAS.DASHBOARD)
            || ss.insertSheet(CFG.HOJAS.DASHBOARD, 0);

  var ordenes = ss.getSheetByName(CFG.HOJAS.ORDENES);
  if (!ordenes) return;

  var datos = ordenes.getDataRange().getValues();
  var m = { pendiente: 0, enProceso: 0, completada: 0, cancelada: 0, total: 0 };
  var porCliente = {};

  for (var i = 1; i < datos.length; i++) {
    var estado  = String(datos[i][29] || "").toLowerCase().trim(); // col AD
    var cliente = String(datos[i][3]  || "").trim();              // col D
    m.total++;
    if (estado === "pendiente")   m.pendiente++;
    if (estado === "en proceso")  m.enProceso++;
    if (estado === "completada")  m.completada++;
    if (estado === "cancelada")   m.cancelada++;
    if (cliente) porCliente[cliente] = (porCliente[cliente] || 0) + 1;
  }

  var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  var filas = [
    ["MÉTRICA",            "VALOR"],
    ["Total órdenes",       m.total],
    ["Pendientes",          m.pendiente],
    ["En proceso",          m.enProceso],
    ["Completadas",         m.completada],
    ["Canceladas",          m.cancelada],
    ["", ""],
    ["CLIENTE",            "ÓRDENES"],
  ];

  Object.keys(porCliente).sort().forEach(function(c) {
    filas.push([c, porCliente[c]]);
  });

  filas.push(["", ""]);
  filas.push(["Actualizado", ahora]);

  hoja.clearContents();
  hoja.getRange(1, 1, filas.length, 2).setValues(filas);

  // Formato encabezados
  [1, 8].forEach(function(r) {
    var enc = hoja.getRange(r, 1, 1, 2);
    enc.setBackground("#1f54a8").setFontColor("#ffffff").setFontWeight("bold");
  });

  // Rojo si hay pendientes
  if (m.pendiente > 0) {
    hoja.getRange(3, 2).setBackground("#fce8e6").setFontColor("#c62828");
  }

  hoja.autoResizeColumns(1, 2);
}
