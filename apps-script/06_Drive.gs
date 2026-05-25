// ============================================================
// 06_Drive.gs  —  Gestión de carpetas en Google Drive
// ============================================================

// Crea (o reutiliza) toda la estructura de carpetas en Drive.
function crearEstructuraDrive() {
  var raiz = _obtenerOCrearCarpeta(null, CONFIG.NOMBRE_ORGANIZACION + " - Sistema Unificado");
  _guardarIDEnScript("CARPETA_RAIZ", raiz.getId());

  var facturas = _obtenerOCrearCarpeta(raiz, "Facturas");
  _guardarIDEnScript("CARPETA_FACTURAS", facturas.getId());

  var reportes = _obtenerOCrearCarpeta(raiz, "Reportes");
  _guardarIDEnScript("CARPETA_REPORTES", reportes.getId());

  var ordenes = _obtenerOCrearCarpeta(raiz, "Ordenes_Produccion");
  _obtenerOCrearCarpeta(ordenes, "OPs");
  _obtenerOCrearCarpeta(ordenes, "OMs");
  _guardarIDEnScript("CARPETA_ORDENES", ordenes.getId());

  _obtenerOCrearCarpeta(raiz, "Archivos_Historicos");

  SpreadsheetApp.getUi().alert(
    "✅ Estructura de Drive creada\n" +
    "Carpeta raíz: " + raiz.getUrl()
  );
}

// ---- Funciones internas usadas por Documentos.gs ----

function _obtenerCarpetaFacturas(anio, mes) {
  var raiz     = _carpetaRaiz();
  var facturas = _obtenerOCrearCarpeta(raiz, "Facturas");
  var carpetaAnio = _obtenerOCrearCarpeta(facturas, String(anio));
  return _obtenerOCrearCarpeta(carpetaAnio, mes);
}

function _obtenerCarpetaReportes(anio) {
  var raiz     = _carpetaRaiz();
  var reportes = _obtenerOCrearCarpeta(raiz, "Reportes");
  return _obtenerOCrearCarpeta(reportes, String(anio));
}

function _obtenerCarpetaOrdenes() {
  var raiz = _carpetaRaiz();
  return _obtenerOCrearCarpeta(raiz, "Ordenes_Produccion");
}

function _carpetaRaiz() {
  var id = CONFIG.DRIVE.CARPETA_RAIZ;
  if (id) {
    try { return DriveApp.getFolderById(id); } catch(e) {}
  }
  // Si no hay ID configurado, crea la estructura ahora
  crearEstructuraDrive();
  return DriveApp.getFolderById(
    PropertiesService.getScriptProperties().getProperty("CARPETA_RAIZ")
  );
}

function _obtenerOCrearCarpeta(padre, nombre) {
  var iter = padre
    ? padre.getFoldersByName(nombre)
    : DriveApp.getFoldersByName(nombre);

  if (iter.hasNext()) return iter.next();

  return padre ? padre.createFolder(nombre) : DriveApp.createFolder(nombre);
}

function _guardarIDEnScript(clave, valor) {
  PropertiesService.getScriptProperties().setProperty(clave, valor);
}
