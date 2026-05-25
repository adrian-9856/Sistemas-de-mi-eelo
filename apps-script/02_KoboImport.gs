// ============================================================
// 02_KoboImport.gs  —  Importar asistencia desde KoboToolbox
// ============================================================

function importarDesdeKobo() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hoja  = ss.getSheetByName(CONFIG.HOJAS.ASISTENCIA);

  if (!hoja) {
    SpreadsheetApp.getUi().alert("❌ No se encontró la hoja " + CONFIG.HOJAS.ASISTENCIA);
    return;
  }

  var url = "https://kc.kobotoolbox.org/api/v2/assets/" +
            CONFIG.KOBO_ASSET_UID + "/data/?format=json&limit=5000";

  var opciones = {
    method: "GET",
    headers: { "Authorization": "Token " + CONFIG.KOBO_API_TOKEN },
    muteHttpExceptions: true,
  };

  var respuesta = UrlFetchApp.fetch(url, opciones);
  if (respuesta.getResponseCode() !== 200) {
    SpreadsheetApp.getUi().alert(
      "❌ Error al conectar con Kobo: " + respuesta.getResponseCode()
    );
    return;
  }

  var datos   = JSON.parse(respuesta.getContentText());
  var results = datos.results || [];

  // Leer UUIDs ya importados para no duplicar
  var existentes = _obtenerUUIDsExistentes(hoja);

  var nuevas = [];
  results.forEach(function(r) {
    var uuid = r[CONFIG.KOBO_CAMPO_UUID] || "";
    if (existentes[uuid]) return; // ya importado

    var nombre   = r[CONFIG.KOBO_CAMPO_NOMBRE]  || "";
    var entrada  = r[CONFIG.KOBO_CAMPO_ENTRADA]  || "";
    var salida   = r[CONFIG.KOBO_CAMPO_SALIDA]   || "";
    var tipo     = r[CONFIG.KOBO_CAMPO_TIPO]     || "Normal";

    var horasTrabajadas = _calcularHoras(entrada, salida);
    var esDiaEstudio    = tipo === "Estudio" ? "Sí" : "No";
    var esTerapia       = tipo === "Terapia" ? "Sí" : "No";
    var porcentaje      = esDiaEstudio === "Sí" ? 0 : 100;
    var horasAPagar     = horasTrabajadas * (porcentaje / 100);

    nuevas.push([
      "",              // A: ID_Participante (se normaliza después)
      nombre,          // B: Nombre
      entrada,         // C: Fecha_Hora_Entrada
      salida,          // D: Fecha_Hora_Salida
      tipo,            // E: Tipo_Registro
      horasTrabajadas, // F: Horas_Trabajadas
      esDiaEstudio,    // G: Es_Dia_Estudio
      esTerapia,       // H: Es_Terapia
      porcentaje,      // I: Porcentaje_Pago
      horasAPagar,     // J: Horas_A_Pagar
      uuid,            // K: UUID_Kobo
    ]);
  });

  if (nuevas.length > 0) {
    var ultimaFila = hoja.getLastRow() + 1;
    hoja.getRange(ultimaFila, 1, nuevas.length, 11).setValues(nuevas);
    normalizarIDs(); // asigna IDs automáticamente
  }

  SpreadsheetApp.getUi().alert(
    "✅ Importación completa\n" +
    nuevas.length + " registros nuevos añadidos."
  );
}

function _calcularHoras(entrada, salida) {
  if (!entrada || !salida) return 0;
  var tEntrada = new Date(entrada);
  var tSalida  = new Date(salida);
  var diff = (tSalida - tEntrada) / (1000 * 60 * 60);
  return Math.max(0, Math.round(diff * 100) / 100);
}

function _obtenerUUIDsExistentes(hoja) {
  var mapa = {};
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var uuid = datos[i][10]; // columna K
    if (uuid) mapa[uuid] = true;
  }
  return mapa;
}
