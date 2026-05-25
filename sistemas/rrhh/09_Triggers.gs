// ============================================================
// 09_Triggers.gs — Automatizaciones (ejecutar una sola vez)
// ============================================================

function configurarTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });

  // Kobo: importar cada hora
  ScriptApp.newTrigger("importarDesdeKobo")
    .timeBased().everyHours(1).create();

  // Facturación: recalcular diario a medianoche
  ScriptApp.newTrigger("calcularFacturacionMes")
    .timeBased().everyDays(1).atHour(0).create();

  // Dashboard: actualizar cada 30 minutos
  ScriptApp.newTrigger("actualizarDashboard")
    .timeBased().everyMinutes(30).create();

  // Resumen mensual: día 1 de cada mes a las 8am
  ScriptApp.newTrigger("enviarResumenMensual")
    .timeBased().onMonthDay(1).atHour(8).create();

  // Recordatorio pagos: cada viernes a las 9am
  ScriptApp.newTrigger("enviarRecordatorioPagos")
    .timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(9).create();

  SpreadsheetApp.getUi().alert(
    "✅ Automatizaciones activas:\n" +
    "• Importar Kobo: cada hora\n" +
    "• Facturación: diario 00:00\n" +
    "• Dashboard: cada 30 min\n" +
    "• Resumen mensual: día 1 / 08:00\n" +
    "• Recordatorio pagos: viernes 09:00"
  );
}
