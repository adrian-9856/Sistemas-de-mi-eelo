// ============================================================
// 08_Dashboard.gs  —  Actualización automática del dashboard
// ============================================================

function actualizarDashboard() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hoja  = ss.getSheetByName(CONFIG.HOJAS.DASHBOARD);
  if (!hoja) hoja = _crearHojaDashboard(ss);

  var ahora    = new Date();
  var mes      = ahora.getMonth() + 1;
  var anio     = ahora.getFullYear();
  var nombreMes = CONFIG.MESES[mes - 1];

  var metricas = _calcularMetricas(ss, nombreMes, anio);

  // Escribe las métricas en la hoja
  var datos = [
    ["MÉTRICA",                       "VALOR",                        "PERÍODO"],
    ["Participantes activos",          metricas.participantesActivos,  "Total"],
    ["Horas trabajadas (mes actual)",  metricas.horasMes,             nombreMes + " " + anio],
    ["Total a pagar (mes actual)",     "Q " + metricas.montoMes,      nombreMes + " " + anio],
    ["Pagos completados",              metricas.pagados,               nombreMes + " " + anio],
    ["Pagos pendientes",               metricas.pendientes,            nombreMes + " " + anio],
    ["Órdenes en proceso",             metricas.ordenesEnProceso,      "Activas"],
    ["Órdenes pendientes",             metricas.ordenesPendientes,     "Activas"],
    ["Última actualización",           Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"), ""],
  ];

  hoja.clearContents();
  hoja.getRange(1, 1, datos.length, 3).setValues(datos);

  // Formato encabezados
  var encabezado = hoja.getRange(1, 1, 1, 3);
  encabezado.setBackground("#1f54a8");
  encabezado.setFontColor("#ffffff");
  encabezado.setFontWeight("bold");

  // Formato fila de actualización
  hoja.getRange(datos.length, 1, 1, 3).setFontStyle("italic").setFontColor("#888888");

  // Colorea pendientes en rojo si hay más de 0
  if (metricas.pendientes > 0) {
    hoja.getRange(6, 2).setBackground("#fce8e6").setFontColor("#c62828");
  }

  hoja.autoResizeColumns(1, 3);
}

function _calcularMetricas(ss, nombreMes, anio) {
  var hojaPart  = ss.getSheetByName(CONFIG.HOJAS.PARTICIPANTES);
  var hojaFact  = ss.getSheetByName(CONFIG.HOJAS.FACTURACION);
  var hojaOrd   = ss.getSheetByName(CONFIG.HOJAS.ORDENES);

  var m = {
    participantesActivos: 0,
    horasMes: 0, montoMes: 0,
    pagados: 0,  pendientes: 0,
    ordenesEnProceso: 0, ordenesPendientes: 0,
  };

  // Participantes activos
  if (hojaPart) {
    var datosPart = hojaPart.getDataRange().getValues();
    for (var i = 1; i < datosPart.length; i++) {
      if (String(datosPart[i][5]).toLowerCase() === "activo") m.participantesActivos++;
    }
  }

  // Facturación del mes
  if (hojaFact) {
    var datosFact = hojaFact.getDataRange().getValues();
    for (var i = 1; i < datosFact.length; i++) {
      var f = datosFact[i];
      if (f[2] !== nombreMes || f[3] !== anio) continue;
      m.horasMes  += parseFloat(f[5]) || 0;
      m.montoMes  += parseFloat(f[6]) || 0;
      if (f[10] === "Sí") { m.pagados++; } else { m.pendientes++; }
    }
    m.horasMes  = Math.round(m.horasMes  * 100) / 100;
    m.montoMes  = (Math.round(m.montoMes * 100) / 100).toFixed(2);
  }

  // Órdenes
  if (hojaOrd) {
    var datosOrd = hojaOrd.getDataRange().getValues();
    for (var i = 1; i < datosOrd.length; i++) {
      var estado = String(datosOrd[i][7]).toLowerCase();
      if (estado === "en proceso")  m.ordenesEnProceso++;
      if (estado === "pendiente")   m.ordenesPendientes++;
    }
  }

  return m;
}

function _crearHojaDashboard(ss) {
  var hoja = ss.insertSheet(CONFIG.HOJAS.DASHBOARD, 0); // primera pestaña
  return hoja;
}
