// ============================================================
// 06_Triggers.gs — Automatizaciones (ejecutar una sola vez)
// ============================================================

function configurarTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });

  // Dashboard se actualiza cada 30 minutos
  ScriptApp.newTrigger("actualizarDashboard")
    .timeBased().everyMinutes(30).create();

  SpreadsheetApp.getUi().alert(
    "✅ Triggers configurados:\n" +
    "• Dashboard: cada 30 minutos"
  );
}
