// ============================================================
// 10_ImportarOPs.gs  —  Extrae datos de OPs y OMs en lote
//
// Estructura real de cada hoja (ej: "001"):
//   Fila 8:  [Cliente, _, _, _, Fecha, _, _, _, Numero_Orden]
//   Fila 9:  [valor_cliente, _, _, _, valor_fecha, _, _, _, valor_numero]
//   Fila 11: [Contacto, _, _, _, Proyecto, _, _, _, Fecha_Promesa]
//   Fila 12: [valor_contacto, _, _, _, valor_proyecto, _, _, _, valor_promesa]
//   Fila 14: [Descripción, _, _, _, _, _, Qty, Tela, _, Comentarios]
//   Fila 15: [valor_desc, _, _, _, _, _, valor_qty, valor_tela, _, valor_coment]
//
// IMPORTANTE: Para usar estas funciones, el archivo de OPs/OMs debe
// estar pegado en el mismo Google Sheets (como hojas ocultas) o
// importado manualmente hoja por hoja en ORDENES_PRODUCCION.
// ============================================================

// Extrae datos de todas las hojas que empiecen con un número (001, 002...)
// y los consolida en ORDENES_PRODUCCION.
// Ejecuta primero con OPs, luego cambia TIPO_ORDEN a "OM" para OMs.
function importarOrdenes() {
  var ui = SpreadsheetApp.getUi();
  var respuesta = ui.prompt(
    "Importar Órdenes",
    "¿Qué tipo de orden vas a importar?\nEscribe: OP  o  OM",
    ui.ButtonSet.OK_CANCEL
  );
  if (respuesta.getSelectedButton() !== ui.Button.OK) return;
  var tipo = respuesta.getResponseText().trim().toUpperCase();
  if (tipo !== "OP" && tipo !== "OM") {
    ui.alert("❌ Escribe OP o OM solamente.");
    return;
  }

  _procesarHojasDeOrdenes(tipo);
}

function _procesarHojasDeOrdenes(tipo) {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var hojas  = ss.getSheets();
  var destino = ss.getSheetByName(CONFIG.HOJAS.ORDENES);
  if (!destino) {
    SpreadsheetApp.getUi().alert("❌ No existe la hoja " + CONFIG.HOJAS.ORDENES);
    return;
  }

  // Detecta qué hojas son numéricas (001, 002...) — las de OPs/OMs
  var hojasOrden = hojas.filter(function(h) {
    return /^\d{3}$/.test(h.getName());
  });

  if (hojasOrden.length === 0) {
    SpreadsheetApp.getUi().alert(
      "❌ No se encontraron hojas con formato numérico (001, 002...).\n" +
      "Pega las hojas del archivo Excel en este Spreadsheet primero."
    );
    return;
  }

  var existentes = _numerosOrdenExistentes(destino);
  var importadas = 0;

  hojasOrden.forEach(function(hoja) {
    var datos = hoja.getDataRange().getValues();

    // Extrae valores de las filas clave
    var cliente      = _celda(datos, 9,  1);  // fila 9,  col A
    var fecha        = _celda(datos, 9,  5);  // fila 9,  col E
    var numeroOrden  = tipo + "26-" + hoja.getName(); // ej: OP26-001
    var contacto     = _celda(datos, 12, 1);
    var proyecto     = _celda(datos, 12, 5);
    var fechaPromesa = _celda(datos, 12, 9);
    var descripcion  = _celda(datos, 15, 1);
    var cantidad     = _celda(datos, 15, 7);
    var tela         = _celda(datos, 15, 8);
    var comentarios  = _celda(datos, 15, 10);

    if (!cliente && !descripcion) return; // hoja vacía

    if (existentes[numeroOrden]) return; // ya importada

    destino.appendRow([
      numeroOrden,   // A: Numero_Orden
      tipo,          // B: Tipo_Orden (OP | OM)
      fecha,         // C: Fecha_Creacion
      cliente,       // D: Cliente
      contacto,      // E: Contacto_Cliente
      descripcion,   // F: Descripcion
      cantidad,      // G: Cantidad
      tela,          // H: Tela/Material
      "Pendiente",   // I: Estado
      "",            // J: Participantes_Asignados
      "",            // K: Horas_Dedicadas
      fechaPromesa,  // L: Fecha_Entrega_Prometida
      proyecto,      // M: Proyecto
      comentarios,   // N: Comentarios
      "",            // O: Enlace_Documento_Drive
    ]);
    importadas++;
  });

  SpreadsheetApp.getUi().alert(
    "✅ " + importadas + " órdenes " + tipo + " importadas a ORDENES_PRODUCCION."
  );
}

function _celda(datos, fila, col) {
  try {
    var val = datos[fila - 1][col - 1];
    if (val instanceof Date) {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
    }
    return val || "";
  } catch(e) {
    return "";
  }
}

function _numerosOrdenExistentes(hoja) {
  var mapa  = {};
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var num = datos[i][0];
    if (num) mapa[num] = true;
  }
  return mapa;
}
