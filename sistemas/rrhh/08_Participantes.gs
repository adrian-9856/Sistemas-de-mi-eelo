// ============================================================
// 08_Participantes.gs — Importar participantes desde Excel
// ============================================================

// Pega la hoja "Listado Participantes" del Excel como "IMPORT_PART"
// y luego ejecuta esta función.
function importarParticipantes() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var hojaTemp = ss.getSheetByName("IMPORT_PART");
  var hojaP    = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES);

  if (!hojaTemp) {
    SpreadsheetApp.getUi().alert(
      "Crea una hoja 'IMPORT_PART' y pega ahí la hoja\n" +
      "'Listado Participantes' del archivo Participantes___mi_eelo_26.xlsx"
    );
    return;
  }

  var datos = hojaTemp.getDataRange().getValues();
  var existentes = _idsExistentes(hojaP);
  var importados = 0;

  // Fila 1 es encabezado: Creamos ID, Nombre, Proyecto, División, Programa, Estado...
  for (var i = 1; i < datos.length; i++) {
    var f = datos[i];
    if (!f[1]) continue; // sin nombre → salta

    var id = String(f[0] || "").trim();
    if (id && existentes[id]) continue; // ya existe

    hojaP.appendRow([
      id,      // A: Creamos ID
      f[1],    // B: Nombre
      f[2],    // C: Proyecto
      f[3],    // D: División
      f[4],    // E: Programa
      f[5],    // F: Estado
      f[6],    // G: Etapa
      f[7],    // H: Educación
      f[8],    // I: Apoyo_Emocional
      f[9],    // J: Inclusion_Laboral
      f[10],   // K: Categoria
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

function _idsExistentes(hoja) {
  var mapa = {}, datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var id = String(datos[i][0]).trim();
    if (id) mapa[id] = true;
  }
  return mapa;
}
