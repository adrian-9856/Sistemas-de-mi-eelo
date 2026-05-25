// ============================================================
// 05_Drive.gs — Estructura de carpetas en Drive para RRHH
//
// Mi eelo · RRHH/
// ├── Facturas/
// │   └── 2026/
// │       ├── Enero/
// │       └── .../
// └── Reportes/
//     └── 2026/
// ============================================================

function crearEstructuraDrive() {
  var raiz     = _getOCreate(null,   CFG.ORG + " · RRHH");
  var facturas = _getOCreate(raiz,   "Facturas");
  var reportes = _getOCreate(raiz,   "Reportes");

  var props = PropertiesService.getScriptProperties();
  props.setProperty("DRIVE_RAIZ",     raiz.getId());
  props.setProperty("DRIVE_FACTURAS", facturas.getId());
  props.setProperty("DRIVE_REPORTES", reportes.getId());

  SpreadsheetApp.getUi().alert("✅ Estructura Drive lista\n" + raiz.getUrl());
}

function _getOCreate(padre, nombre) {
  var iter = padre ? padre.getFoldersByName(nombre) : DriveApp.getFoldersByName(nombre);
  if (iter.hasNext()) return iter.next();
  return padre ? padre.createFolder(nombre) : DriveApp.createFolder(nombre);
}
