// ============================================================
// 03_Participantes.gs  —  Gestión de participantes
//
// Estructura real de PARTICIPANTES_MASTER (desde Participantes___mi_eelo_26.xlsx):
//   Col A: Creamos ID  (ej: ANLA060686, o vacío en algunos)
//   Col B: Nombre
//   Col C: Proyecto    (Costura, Servicios, Olas...)
//   Col D: División    (Textil, Manufactura...)
//   Col E: Programa    (mi eelo)
//   Col F: Estado      (Activo, Inactivo...)
//   Col G: Etapa       (Inscritx...)
//   Col H: Educación   (TRUE/FALSE)
//   Col I: Apoyo Emocional (TRUE/FALSE)
//   Col J: Inclusión Laboral (TRUE/FALSE)
//   Col K: Categoría   (A, B, C, D...)
// ============================================================

// Normaliza los IDs en ASISTENCIA_KOBO usando el Creamos ID
// extraído directamente del campo "Participante (CREAMOS_ID)".
// Si la asistencia ya tiene el ID extraído, esta función verifica
// que exista en PARTICIPANTES_MASTER.
function normalizarIDs() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var hojaPart  = ss.getSheetByName(CONFIG.HOJAS.PARTICIPANTES);
  var hojaAsist = ss.getSheetByName(CONFIG.HOJAS.ASISTENCIA);
  if (!hojaPart || !hojaAsist) return;

  var idsValidos = _obtenerIDsValidos(hojaPart);
  var datos      = hojaAsist.getDataRange().getValues();
  var sinID = 0;

  for (var i = 1; i < datos.length; i++) {
    var id = String(datos[i][0]).trim();
    if (!id) {
      sinID++;
      continue;
    }
    // Si el ID no está en participantes, marca en rojo para revisión
    if (!idsValidos[id]) {
      hojaAsist.getRange(i + 1, 1).setBackground("#fce8e6");
    }
  }

  if (sinID > 0) {
    SpreadsheetApp.getUi().alert(
      "⚠️ " + sinID + " registros en ASISTENCIA_KOBO sin ID de participante.\n" +
      "Revisa que el campo 'Participante' en Kobo tenga el formato: Nombre (CREAMOS_ID)"
    );
  } else {
    SpreadsheetApp.getUi().alert("✅ Todos los IDs normalizados correctamente.");
  }
}

// Importa los participantes del archivo Excel a la hoja PARTICIPANTES_MASTER.
// Úsalo la primera vez para cargar los datos históricos.
// Pega el contenido del Excel en una hoja temporal llamada "IMPORT_PART" y ejecuta esto.
function importarParticipantesDesdeExcel() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var hojaTemp = ss.getSheetByName("IMPORT_PART");
  var hojaDest = ss.getSheetByName(CONFIG.HOJAS.PARTICIPANTES);

  if (!hojaTemp) {
    SpreadsheetApp.getUi().alert(
      "Crea una hoja llamada 'IMPORT_PART' y pega ahí el contenido de Participantes___mi_eelo_26.xlsx"
    );
    return;
  }

  var datos = hojaTemp.getDataRange().getValues();
  var importados = 0;
  var existentes = _obtenerIDsValidos(hojaDest);

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var nombre = fila[1];
    if (!nombre) continue;

    var id = fila[0] || ""; // Creamos ID (puede estar vacío)

    // Salta si ya existe
    if (id && existentes[id]) continue;

    hojaDest.appendRow([
      id,      // A: Creamos ID
      nombre,  // B: Nombre
      fila[2], // C: Proyecto
      fila[3], // D: División
      fila[4], // E: Programa
      fila[5], // F: Estado
      fila[6], // G: Etapa
      fila[7], // H: Educación
      fila[8], // I: Apoyo Emocional
      fila[9], // J: Inclusión Laboral
      fila[10],// K: Categoría
      "",      // L: DPI
      "",      // M: NIT
      "",      // N: Correo_Electronico
      "",      // O: Banco
      "",      // P: Numero_Cuenta
      "",      // Q: Forma_Pago
    ]);
    importados++;
  }

  SpreadsheetApp.getUi().alert("✅ " + importados + " participantes importados.");
}

// ---- Utilidades internas ----

function _obtenerIDsValidos(hojaPart) {
  var mapa  = {};
  var datos = hojaPart.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var id = String(datos[i][0]).trim();
    if (id) mapa[id] = true;
  }
  return mapa;
}

// Construye mapeo ID → fila completa de participante
function _construirMapeoIDParticipante() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hoja  = ss.getSheetByName(CONFIG.HOJAS.PARTICIPANTES);
  var mapa  = {};
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var id = String(datos[i][0]).trim();
    if (id) mapa[id] = datos[i];
  }
  return mapa;
}

// Busca la fila completa de un participante por Creamos ID
function _buscarParticipante(id) {
  var mapa = _construirMapeoIDParticipante();
  return mapa[String(id).trim()] || null;
}
