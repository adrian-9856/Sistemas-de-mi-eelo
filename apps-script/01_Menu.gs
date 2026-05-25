// ============================================================
// 01_Menu.gs  —  Menú personalizado y punto de entrada
// ============================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🚀 Mi eelo")

    // ── Carga inicial (una sola vez) ──────────────────────
    .addSubMenu(ui.createMenu("📂 Carga inicial de datos")
      .addItem("👥 Importar participantes desde Excel",   "importarParticipantesDesdeExcel")
      .addItem("📋 Importar asistencia (DatosKobo)",      "importarDesdePlanillaLocal")
      .addItem("💳 Importar datos de facturación (DPI/NIT/Banco)", "importarDatosFacturacion")
      .addItem("💰 Importar historial facturación (14 meses)",     "importarHistorialFacturacion")
      .addItem("📦 Importar OPs / OMs (hojas numéricas)", "importarOrdenes")
      .addItem("💼 Importar ventas (Comercial Sales)",    "importarVentas")
      .addItem("🛒 Importar compras (Comercial Compras)", "importarCompras")
    )
    .addSeparator()

    // ── Operaciones diarias ───────────────────────────────
    .addItem("📥 Importar asistencia desde Kobo",         "importarDesdeKobo")
    .addItem("🔗 Emparejar entradas/salidas → horas",     "emparejarAsistencia")
    .addItem("✅ Verificar IDs participantes",             "normalizarIDs")
    .addSeparator()

    // ── Facturación ──────────────────────────────────────
    .addItem("💰 Calcular facturación del mes",           "calcularFacturacionMes")
    .addItem("📄 Generar facturas en Google Docs",         "generarFacturasMes")
    .addItem("📊 Generar reporte mensual (Doc)",           "generarReporteMensual")
    .addSeparator()

    // ── Órdenes ──────────────────────────────────────────
    .addItem("📦 Generar Doc de la orden seleccionada",   "generarOrdenProduccion")
    .addSeparator()

    // ── Clientes / Proveedores ────────────────────────────
    .addSubMenu(ui.createMenu("👤 Clientes y Proveedores")
      .addItem("📁 Crear carpeta Drive del cliente seleccionado",    "crearCarpetaCliente")
      .addItem("📁 Crear carpetas para TODOS los clientes",          "crearCarpetasTodosClientes")
      .addItem("🔍 Ver órdenes del cliente seleccionado",            "verOrdenesCliente")
      .addItem("📁 Crear carpeta Drive del proveedor seleccionado",  "crearCarpetaProveedor")
      .addItem("🧹 Limpiar filtros en Órdenes",                      "limpiarFiltros")
    )
    .addSeparator()

    // ── Notificaciones ───────────────────────────────────
    .addItem("🔔 Enviar recordatorio de pagos",           "enviarRecordatorioPagos")
    .addItem("📬 Enviar resumen mensual al admin",         "enviarResumenMensual")
    .addSeparator()

    // ── Sistema ──────────────────────────────────────────
    .addItem("📁 Crear estructura en Drive",              "crearEstructuraDrive")
    .addItem("🔄 Actualizar Dashboard",                   "actualizarDashboard")
    .addItem("⚙️  Configurar triggers automáticos",        "configurarTriggers")

    .addToUi();
}

function onEdit(e) {
  var hoja = e.range.getSheet().getName();
  var col  = e.range.getColumn();

  // Pago marcado en FACTURACION_CONSOLIDADA (col K = 11)
  if (hoja === CONFIG.HOJAS.FACTURACION && col === 11) {
    actualizarDashboard();
  }

  // Estado de orden cambiado en ORDENES_PRODUCCION (col AD = 30)
  if (hoja === CONFIG.HOJAS.ORDENES && col === 30) {
    actualizarDashboard();
  }
}
