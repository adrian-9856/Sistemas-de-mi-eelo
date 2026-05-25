// ============================================================
// 04_Drive.gs — Estructura de carpetas en Drive
//
// Mi eelo · Producción/
// └── Órdenes/
//     └── Clientes/
//         ├── Peace by Piece/       ← un folder por cliente
//         ├── Home Collection/
//         └── .../
// ============================================================

// Crea toda la estructura de Drive (ejecutar una sola vez).
function crearEstructuraDrive() {
  var raiz     = _getOCreate(null,      CFG.ORG + " · Producción");
  var ordenes  = _getOCreate(raiz,      "Órdenes");
  var clientes = _getOCreate(ordenes,   "Clientes");

  // Guarda los IDs en ScriptProperties para no buscar cada vez
  var props = PropertiesService.getScriptProperties();
  props.setProperty("DRIVE_RAIZ",     raiz.getId());
  props.setProperty("DRIVE_ORDENES",  ordenes.getId());
  props.setProperty("DRIVE_CLIENTES", clientes.getId());

  SpreadsheetApp.getUi().alert(
    "✅ Estructura Drive lista\n" + raiz.getUrl()
  );
}

// Crea (o abre) la carpeta del cliente dentro de Órdenes/Clientes/
function crearCarpetaCliente() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hoja  = ss.getSheetByName(CFG.HOJAS.CLIENTES);
  var fila  = hoja.getActiveRange().getRow();
  if (fila < 2) return;

  var nombre = hoja.getRange(fila, 2).getValue();
  if (!nombre) { SpreadsheetApp.getUi().alert("La fila no tiene nombre de cliente."); return; }

  var carpeta = _carpetaDeCliente(nombre);
  hoja.getRange(fila, 7).setValue(carpeta.getUrl()); // col G

  SpreadsheetApp.getUi().alert("✅ Carpeta: " + carpeta.getUrl());
}

// ---- Funciones internas usadas por otros módulos ----

function _carpetaDeCliente(nombreCliente) {
  var clientes = _carpetaClientes();
  return _getOCreate(clientes, String(nombreCliente).trim());
}

function _carpetaClientes() {
  var props = PropertiesService.getScriptProperties();
  var id    = props.getProperty("DRIVE_CLIENTES");
  if (id) {
    try { return DriveApp.getFolderById(id); } catch(e) {}
  }
  // Si no existe, crea la estructura completa
  crearEstructuraDrive();
  return DriveApp.getFolderById(
    PropertiesService.getScriptProperties().getProperty("DRIVE_CLIENTES")
  );
}

function _getOCreate(padre, nombre) {
  var iter = padre ? padre.getFoldersByName(nombre) : DriveApp.getFoldersByName(nombre);
  if (iter.hasNext()) return iter.next();
  return padre ? padre.createFolder(nombre) : DriveApp.createFolder(nombre);
}
