// ============================================================
// 13_Clientes.gs  —  Gestión de clientes y proveedores
//
// Hoja CLIENTES:
//   A: ID_Cliente  B: Nombre  C: Contacto  D: Email
//   E: Telefono    F: Pais    G: Total_Ordenes (fórmula)
//   H: Ultima_Orden (fórmula)  I: Total_Facturado (fórmula)
//   J: URL_Carpeta_Drive       K: Notas
//
// Hoja PROVEEDORES:
//   A: ID_Proveedor  B: Nombre  C: Contacto  D: Email
//   E: Telefono      F: Producto_Principal
//   G: Total_Compras (fórmula)  H: URL_Carpeta_Drive  I: Notas
// ============================================================

// Crea (o actualiza) la carpeta de Drive de un cliente
// y guarda el enlace en la hoja CLIENTES.
function crearCarpetaCliente() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hoja  = ss.getSheetByName("CLIENTES");
  if (!hoja) {
    SpreadsheetApp.getUi().alert("Crea primero la hoja CLIENTES.");
    return;
  }

  var fila = hoja.getActiveRange().getRow();
  if (fila < 2) return;

  var nombre = hoja.getRange(fila, 2).getValue();
  if (!nombre) {
    SpreadsheetApp.getUi().alert("La fila no tiene nombre de cliente.");
    return;
  }

  var carpeta = _carpetaCliente(nombre);
  hoja.getRange(fila, 10).setValue(carpeta.getUrl()); // col J

  SpreadsheetApp.getUi().alert(
    "✅ Carpeta creada para: " + nombre + "\n" + carpeta.getUrl()
  );
}

// Crear carpetas para TODOS los clientes en la hoja
function crearCarpetasTodosClientes() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hoja  = ss.getSheetByName("CLIENTES");
  if (!hoja) return;

  var datos = hoja.getDataRange().getValues();
  var creadas = 0;

  for (var i = 1; i < datos.length; i++) {
    var nombre = datos[i][1]; // col B
    if (!nombre) continue;
    if (datos[i][9]) continue; // col J ya tiene URL

    var carpeta = _carpetaCliente(nombre);
    hoja.getRange(i + 1, 10).setValue(carpeta.getUrl());
    creadas++;
  }
  SpreadsheetApp.getUi().alert("✅ " + creadas + " carpetas de clientes creadas en Drive.");
}

// Ver todas las órdenes de un cliente seleccionado
function verOrdenesCliente() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var hojaC  = ss.getSheetByName("CLIENTES");
  var fila   = hojaC.getActiveRange().getRow();
  if (fila < 2) return;

  var cliente = hojaC.getRange(fila, 2).getValue();
  if (!cliente) return;

  // Aplica filtro en ORDENES_PRODUCCION por el cliente
  var hojaO = ss.getSheetByName(CONFIG.HOJAS.ORDENES);
  ss.setActiveSheet(hojaO);

  var rango = hojaO.getDataRange();
  var filtro = rango.getFilter() || rango.createFilter();

  // Filtra columna D (col 4) por nombre de cliente
  var criteria = SpreadsheetApp.newFilterCriteria()
    .whenTextEqualTo(cliente)
    .build();
  filtro.setColumnFilterCriteria(4, criteria);

  SpreadsheetApp.getUi().alert(
    "Mostrando órdenes de: " + cliente +
    "\nUsa 🚀 Mi eelo → Limpiar filtros para volver a ver todas."
  );
}

// Limpia los filtros activos en ORDENES_PRODUCCION
function limpiarFiltros() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(CONFIG.HOJAS.ORDENES);
  var filtro = hoja.getDataRange().getFilter();
  if (filtro) filtro.remove();
}

// ---- Proveedores ----

function crearCarpetaProveedor() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hoja  = ss.getSheetByName("PROVEEDORES");
  if (!hoja) return;

  var fila   = hoja.getActiveRange().getRow();
  var nombre = hoja.getRange(fila, 2).getValue();
  if (!nombre) return;

  var carpeta = _carpetaProveedor(nombre);
  hoja.getRange(fila, 8).setValue(carpeta.getUrl()); // col H

  SpreadsheetApp.getUi().alert("✅ Carpeta creada: " + carpeta.getUrl());
}

// ---- Carpeta de cliente en Drive ----
// Esta función es usada también por 12_OrdenesDocs.gs

function _carpetaCliente(nombreCliente) {
  var raiz     = _carpetaRaiz();
  var ordenes  = _obtenerOCrearCarpeta(raiz, "Órdenes");
  var clientes = _obtenerOCrearCarpeta(ordenes, "Clientes");
  return _obtenerOCrearCarpeta(clientes, String(nombreCliente).trim());
}

function _carpetaProveedor(nombreProveedor) {
  var raiz      = _carpetaRaiz();
  var compras   = _obtenerOCrearCarpeta(raiz, "Compras");
  var proveedores = _obtenerOCrearCarpeta(compras, "Proveedores");
  return _obtenerOCrearCarpeta(proveedores, String(nombreProveedor).trim());
}
