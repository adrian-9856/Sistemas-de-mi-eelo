// ============================================================
// 09_Triggers.gs  —  Automatización: configurar triggers
// ============================================================

// Ejecuta una sola vez para instalar todos los triggers automáticos.
function configurarTriggers() {
  // Borra triggers anteriores para no duplicar
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });

  // 1. Importar desde Kobo cada hora
  ScriptApp.newTrigger("importarDesdeKobo")
    .timeBased()
    .everyHours(1)
    .create();

  // 2. Calcular facturación todos los días a medianoche
  ScriptApp.newTrigger("calcularFacturacionMes")
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();

  // 3. Actualizar dashboard cada 30 minutos
  ScriptApp.newTrigger("actualizarDashboard")
    .timeBased()
    .everyMinutes(30)
    .create();

  // 4. Enviar resumen mensual el día 1 de cada mes a las 8am
  ScriptApp.newTrigger("enviarResumenMensual")
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .create();

  // 5. Recordatorio de pagos pendientes cada viernes a las 9am
  ScriptApp.newTrigger("enviarRecordatorioPagos")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(9)
    .create();

  SpreadsheetApp.getUi().alert(
    "✅ Triggers configurados:\n" +
    "• Importar Kobo: cada hora\n" +
    "• Calcular facturación: diario (00:00)\n" +
    "• Actualizar dashboard: cada 30 min\n" +
    "• Resumen mensual: día 1 de cada mes (08:00)\n" +
    "• Recordatorio pagos: viernes (09:00)"
  );
}
