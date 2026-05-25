// ============================================================
// 03_Participantes.gs  —  Normalización de IDs y utilidades
// ============================================================

function normalizarIDs() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var hojaPart  = ss.getSheetByName(CONFIG.HOJAS.PARTICIPANTES);
  var hojaAsist = ss.getSheetByName(CONFIG.HOJAS.ASISTENCIA);

  if (!hojaPart || !hojaAsist) return;

  var mapeo = _construirMapeoNombreID(hojaPart);
  var datos  = hojaAsist.getDataRange().getValues();
  var cambios = 0;

  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0]) continue; // ya tiene ID
    var nombre = String(datos[i][1]).toLowerCase().trim();
    var id     = mapeo[nombre];
    if (id) {
      hojaAsist.getRange(i + 1, 1).setValue(id);
      cambios++;
    }
  }

  Logger.log("IDs normalizados: " + cambios);
}

// Genera IDs secuenciales P001, P002... en PARTICIPANTES_MASTER columna A
function generarIDsParticipantes() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(CONFIG.HOJAS.PARTICIPANTES);
  if (!hoja) return;

  var ultimaFila = hoja.getLastRow();
  for (var i = 2; i <= ultimaFila; i++) {
    var actual = hoja.getRange(i, 1).getValue();
    if (!actual) {
      var num = i - 1;
      hoja.getRange(i, 1).setValue("P" + String(num).padStart(3, "0"));
    }
  }
  SpreadsheetApp.getUi().alert("✅ IDs generados hasta la fila " + ultimaFila);
}

// Devuelve un objeto { "nombre en minúsculas": "P001" }
function _construirMapeoNombreID(hojaPart) {
  var datos = hojaPart.getDataRange().getValues();
  var mapa  = {};
  for (var i = 1; i < datos.length; i++) {
    var id     = datos[i][0]; // col A
    var nombre = datos[i][2]; // col C
    if (id && nombre) {
      mapa[String(nombre).toLowerCase().trim()] = id;
    }
  }
  return mapa;
}

// Busca y devuelve la fila completa de un participante por su ID
function _buscarParticipante(id) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(CONFIG.HOJAS.PARTICIPANTES);
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0] === id) return datos[i];
  }
  return null;
}
