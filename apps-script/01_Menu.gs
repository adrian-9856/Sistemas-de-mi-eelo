// ============================================================
// 01_Menu.gs  —  Menú personalizado y punto de entrada
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚀 Mi eelo")
    .addItem("📥 Importar desde Kobo",          "importarDesdeKobo")
    .addItem("🔗 Normalizar IDs participantes",  "normalizarIDs")
    .addSeparator()
    .addItem("💰 Calcular facturación del mes",  "calcularFacturacionMes")
    .addItem("📄 Generar facturas en Google Docs","generarFacturasMes")
    .addSeparator()
    .addItem("📊 Generar reporte mensual (Doc)", "generarReporteMensual")
    .addItem("📦 Generar orden de producción",   "generarOrdenProduccion")
    .addSeparator()
    .addItem("🔔 Enviar recordatorio de pagos",  "enviarRecordatorioPagos")
    .addItem("📬 Resumen mensual por correo",    "enviarResumenMensual")
    .addSeparator()
    .addItem("📁 Crear estructura en Drive",     "crearEstructuraDrive")
    .addItem("🔄 Actualizar Dashboard",          "actualizarDashboard")
    .addSeparator()
    .addItem("⚙️  Configurar triggers automáticos","configurarTriggers")
    .addToUi();
}

// Ejecutado automáticamente cada vez que se abre el Spreadsheet.
function onEdit(e) {
  var hoja = e.range.getSheet().getName();
  if (hoja === CONFIG.HOJAS.FACTURACION) {
    // Si marcaron una factura como pagada, actualiza el dashboard.
    var col = e.range.getColumn();
    if (col === 11) actualizarDashboard(); // columna K = Pagado
  }
}
