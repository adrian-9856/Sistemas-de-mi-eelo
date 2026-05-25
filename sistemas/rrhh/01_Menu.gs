function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("👥 RRHH")
    .addSubMenu(ui.createMenu("📂 Carga inicial (una sola vez)")
      .addItem("👥 Importar participantes desde Excel",     "importarParticipantes")
      .addItem("📋 Importar asistencia histórica (Planilla)","importarAsistenciaHistorica")
      .addItem("💰 Importar facturación histórica (14 meses)","importarFacturacionHistorica")
    )
    .addSeparator()
    .addItem("📥 Importar asistencia desde Kobo",           "importarDesdeKobo")
    .addItem("🔗 Emparejar entradas/salidas → horas",       "emparejarAsistencia")
    .addSeparator()
    .addItem("💰 Calcular facturación del mes actual",       "calcularFacturacionMes")
    .addItem("📄 Generar recibos de pago (Google Docs)",     "generarRecibosMes")
    .addItem("📊 Generar reporte mensual",                   "generarReporteMensual")
    .addSeparator()
    .addItem("🔔 Enviar recordatorio de pagos pendientes",   "enviarRecordatorioPagos")
    .addItem("📬 Enviar resumen mensual al admin",           "enviarResumenMensual")
    .addSeparator()
    .addItem("🔄 Actualizar Dashboard",                     "actualizarDashboard")
    .addItem("📁 Crear estructura en Drive",                "crearEstructuraDrive")
    .addItem("⚙️  Configurar automatizaciones",              "configurarTriggers")
    .addToUi();
}

function onEdit(e) {
  // Pago marcado en FACTURACION (col K = 11) → refresca dashboard
  if (e.range.getSheet().getName() === CFG.HOJAS.FACTURACION &&
      e.range.getColumn() === 11) {
    actualizarDashboard();
  }
}
