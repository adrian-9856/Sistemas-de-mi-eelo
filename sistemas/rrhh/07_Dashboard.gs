// ============================================================
// 07_Dashboard.gs — Métricas RRHH en tiempo real
// ============================================================

function actualizarDashboard() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(CFG.HOJAS.DASHBOARD)
            || ss.insertSheet(CFG.HOJAS.DASHBOARD, 0);

  var ahora    = new Date();
  var mes      = CFG.MESES[ahora.getMonth()];
  var anio     = ahora.getFullYear();

  var hojaP = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES);
  var hojaF = ss.getSheetByName(CFG.HOJAS.FACTURACION);

  var activos    = 0;
  var totalMonto = 0, pagados = 0, pendientes = 0;

  // Participantes activos
  if (hojaP) {
    var dp = hojaP.getDataRange().getValues();
    for (var i = 1; i < dp.length; i++) {
      if (String(dp[i][5]).toLowerCase() === "activo") activos++;
    }
  }

  // Facturación del mes
  if (hojaF) {
    var df = hojaF.getDataRange().getValues();
    for (var i = 1; i < df.length; i++) {
      if (df[i][2] !== mes || df[i][3] !== anio) continue;
      totalMonto += parseFloat(df[i][6]) || 0;
      if (df[i][10] === "Sí") { pagados++; } else { pendientes++; }
    }
  }

  var ts    = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  var filas = [
    ["MÉTRICA",                    "VALOR"],
    ["Participantes activos",       activos],
    ["Total a pagar (" + mes + ")", "Q " + totalMonto.toFixed(2)],
    ["Quincenas pagadas",           pagados],
    ["Quincenas pendientes",        pendientes],
    ["",                            ""],
    ["Actualizado",                 ts],
  ];

  hoja.clearContents();
  hoja.getRange(1, 1, filas.length, 2).setValues(filas);
  hoja.getRange(1, 1, 1, 2).setBackground("#639922").setFontColor("#fff").setFontWeight("bold");

  if (pendientes > 0) {
    hoja.getRange(5, 2).setBackground("#fce8e6").setFontColor("#c62828");
  }

  hoja.autoResizeColumns(1, 2);
}
