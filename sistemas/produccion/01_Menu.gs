function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("📦 Producción")
    .addItem("➕ Nueva orden (OP o OM)",             "nuevaOrden")
    .addItem("📄 Generar Doc de la orden seleccionada", "generarDocOrden")
    .addSeparator()
    .addItem("🔍 Ver todas las órdenes de un cliente",  "filtrarPorCliente")
    .addItem("🧹 Quitar filtros",                       "limpiarFiltros")
    .addSeparator()
    .addItem("📁 Crear carpeta Drive del cliente",      "crearCarpetaCliente")
    .addItem("🔄 Actualizar Dashboard",                 "actualizarDashboard")
    .addSeparator()
    .addItem("⚙️  Configurar automatizaciones",          "configurarTriggers")
    .addToUi();
}

// Cuando cambia el Estado de una orden (col AD = 30), refresca el dashboard.
function onEdit(e) {
  if (e.range.getSheet().getName() === CFG.HOJAS.ORDENES &&
      e.range.getColumn() === 30) {
    actualizarDashboard();
  }
}
