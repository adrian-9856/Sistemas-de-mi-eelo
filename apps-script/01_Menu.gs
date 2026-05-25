// ============================================================
// 01_Menu.gs  —  Menú personalizado y punto de entrada
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚀 Mi eelo")

    // ── Importación inicial (una sola vez) ──
    .addSubMenu(SpreadsheetApp.getUi().createMenu("📂 Carga inicial de datos")
      .addItem("👥 Importar participantes desde Excel",  "importarParticipantesDesdeExcel")
      .addItem("📋 Importar asistencia desde Planilla",  "importarDesdePlanillaLocal")
      .addItem("📦 Importar OPs / OMs",                  "importarOrdenes")
      .addItem("💼 Importar ventas (Comercial)",          "importarVentas")
      .addItem("🛒 Importar compras (Comercial)",         "importarCompras")
    )
    .addSeparator()

    // ── Operaciones diarias ──
    .addItem("📥 Importar asistencia desde Kobo",        "importarDesdeKobo")
    .addItem("🔗 Emparejar entradas/salidas + horas",    "emparejarAsistencia")
    .addItem("✅ Verificar IDs de participantes",         "normalizarIDs")
    .addSeparator()

    // ── Facturación ──
    .addItem("💰 Calcular facturación del mes",          "calcularFacturacionMes")
    .addItem("📄 Generar facturas en Google Docs",        "generarFacturasMes")
    .addItem("📊 Generar reporte mensual (Doc)",          "generarReporteMensual")
    .addSeparator()

    // ── Órdenes ──
    .addItem("📦 Generar Doc de orden seleccionada",     "generarOrdenProduccion")
    .addSeparator()

    // ── Notificaciones ──
    .addItem("🔔 Enviar recordatorio de pagos",          "enviarRecordatorioPagos")
    .addItem("📬 Enviar resumen mensual al admin",        "enviarResumenMensual")
    .addSeparator()

    // ── Sistema ──
    .addItem("📁 Crear estructura en Drive",             "crearEstructuraDrive")
    .addItem("🔄 Actualizar Dashboard",                  "actualizarDashboard")
    .addItem("⚙️  Configurar triggers automáticos",       "configurarTriggers")

    .addToUi();
}

function onEdit(e) {
  var hoja = e.range.getSheet().getName();
  // Cuando se marca un pago como completado en FACTURACION_CONSOLIDADA (col K)
  if (hoja === CONFIG.HOJAS.FACTURACION && e.range.getColumn() === 11) {
    actualizarDashboard();
  }
  // Cuando se cambia el estado de una orden (col I de ORDENES_PRODUCCION)
  if (hoja === CONFIG.HOJAS.ORDENES && e.range.getColumn() === 9) {
    actualizarDashboard();
  }
}
