// ============================================================
// 02_KoboImport.gs  —  Importar y emparejar asistencia desde Kobo
//
// Estructura real en DatosKobo:
//   Col D ("Ingreso / Egreso"): "🟢 Entrada" | "🔴 Salida"
//   Col E ("Participante"):     "Nombre Completo (CREAMOS_ID)"
//   Col A ("start"):            timestamp del registro
//   Col L ("_uuid"):            ID único del envío
//
// Cada fila es UN evento. Se emparejan Entrada+Salida del mismo
// participante en el mismo día para calcular horas trabajadas.
// ============================================================

// Importa registros nuevos de Kobo (vía API) y los agrega a ASISTENCIA_KOBO.
function importarDesdeKobo() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(CONFIG.HOJAS.ASISTENCIA);
  if (!hoja) {
    SpreadsheetApp.getUi().alert("❌ No existe la hoja " + CONFIG.HOJAS.ASISTENCIA);
    return;
  }

  var url = "https://kc.kobotoolbox.org/api/v2/assets/" +
            CONFIG.KOBO_ASSET_UID + "/data/?format=json&limit=5000";

  var resp = UrlFetchApp.fetch(url, {
    method: "GET",
    headers: { "Authorization": "Token " + CONFIG.KOBO_API_TOKEN },
    muteHttpExceptions: true,
  });

  if (resp.getResponseCode() !== 200) {
    SpreadsheetApp.getUi().alert("❌ Error Kobo: " + resp.getResponseCode());
    return;
  }

  var results = JSON.parse(resp.getContentText()).results || [];
  var existentes = _uuidsExistentes(hoja);
  var nuevos = [];

  results.forEach(function(r) {
    var uuid = r["_uuid"] || "";
    if (existentes[uuid]) return;

    var tipo         = r["Ingreso / Egreso"] || "";
    var participante = r["Participante"]     || "";
    var timestamp    = new Date(r["start"]   || r["_submission_time"]);
    var creamos_id   = _extraerCreamos_ID(participante);
    var nombre       = _extraerNombre(participante);

    nuevos.push([
      creamos_id,   // A: ID_Participante (Creamos ID)
      nombre,       // B: Nombre_Participante
      timestamp,    // C: Fecha_Hora_Registro
      tipo,         // D: Tipo ("🟢 Entrada" | "🔴 Salida")
      "",           // E: Horas_Trabajadas (se calcula al emparejar)
      "",           // F: Es_Dia_Estudio
      "",           // G: Es_Terapia
      "",           // H: Porcentaje_Pago
      "",           // I: Horas_A_Pagar
      uuid,         // J: UUID_Kobo
    ]);
  });

  if (nuevos.length > 0) {
    hoja.getRange(hoja.getLastRow() + 1, 1, nuevos.length, 10).setValues(nuevos);
  }

  // Empareja todas las entradas/salidas y recalcula horas
  emparejarAsistencia();

  SpreadsheetApp.getUi().alert(
    "✅ Importación completa\n" + nuevos.length + " registros nuevos."
  );
}

// Opción alternativa: importar directamente desde la hoja DatosKobo
// del archivo Planilla (para la carga inicial histórica).
function importarDesdePlanillaLocal() {
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var hojaKobo   = ss.getSheetByName("DatosKobo");
  var hojaAsist  = ss.getSheetByName(CONFIG.HOJAS.ASISTENCIA);

  if (!hojaKobo) {
    SpreadsheetApp.getUi().alert("❌ No existe la hoja 'DatosKobo'. Pégala primero.");
    return;
  }

  var datos      = hojaKobo.getDataRange().getValues();
  var existentes = _uuidsExistentes(hojaAsist);
  var nuevos     = [];
  var C = CONFIG.KOBO_COL;

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var uuid         = fila[C.UUID - 1];
    if (!uuid || existentes[uuid]) continue;

    var tipo         = fila[C.TIPO - 1]        || "";
    var participante = fila[C.PARTICIPANTE - 1] || "";
    var timestamp    = fila[C.START - 1];

    if (!participante || !timestamp) continue;

    var creamos_id = _extraerCreamos_ID(participante);
    var nombre     = _extraerNombre(participante);

    nuevos.push([
      creamos_id, nombre, timestamp, tipo,
      "", "", "", "", "", uuid,
    ]);
  }

  if (nuevos.length > 0) {
    hojaAsist.getRange(hojaAsist.getLastRow() + 1, 1, nuevos.length, 10).setValues(nuevos);
  }

  emparejarAsistencia();

  SpreadsheetApp.getUi().alert(
    "✅ Importación desde Planilla local completa\n" +
    nuevos.length + " registros añadidos."
  );
}

// Empareja filas de Entrada con su Salida correspondiente
// y calcula Horas_Trabajadas + Horas_A_Pagar.
function emparejarAsistencia() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(CONFIG.HOJAS.ASISTENCIA);
  var datos = hoja.getDataRange().getValues();

  // Agrupa por participante + fecha (YYYY-MM-DD)
  var grupos = {};
  for (var i = 1; i < datos.length; i++) {
    var id        = datos[i][0];
    var timestamp = datos[i][2];
    var tipo      = datos[i][3];
    if (!id || !timestamp || !tipo) continue;

    var fecha = _fechaClave(new Date(timestamp));
    var clave = id + "|" + fecha;
    if (!grupos[clave]) grupos[clave] = { entradas: [], salidas: [], filas: [] };

    if (tipo === CONFIG.KOBO_ENTRADA) {
      grupos[clave].entradas.push({ ts: new Date(timestamp), fila: i + 1 });
    } else if (tipo === CONFIG.KOBO_SALIDA) {
      grupos[clave].salidas.push({ ts: new Date(timestamp), fila: i + 1 });
    }
    grupos[clave].filas.push(i + 1);
  }

  // Consulta dias de estudio y terapias
  var diasEstudio = _cargarDiasEstudio(ss);
  var terapias    = _cargarTerapias(ss);

  // Calcular horas por grupo y escribir en la hoja
  Object.keys(grupos).forEach(function(clave) {
    var g = grupos[clave];
    if (g.entradas.length === 0 || g.salidas.length === 0) return;

    var partes = clave.split("|");
    var id     = partes[0];
    var fecha  = partes[1];

    // Toma la primera entrada y la última salida del día
    g.entradas.sort(function(a,b){ return a.ts - b.ts; });
    g.salidas.sort(function(a,b){ return a.ts - b.ts; });
    var entrada = g.entradas[0].ts;
    var salida  = g.salidas[g.salidas.length - 1].ts;

    var horas       = Math.max(0, Math.round((salida - entrada) / 36000) / 100);
    var esEstudio   = diasEstudio[id]  ? "Sí" : "No";
    var esTerapia   = terapias[id]     ? "Sí" : "No";
    var porcentaje  = esEstudio === "Sí" ? 0 : 100;
    var horasAPagar = horas * (porcentaje / 100);

    // Escribe en TODAS las filas del grupo (entrada + salida)
    g.filas.forEach(function(numFila) {
      hoja.getRange(numFila, 5).setValue(horas);
      hoja.getRange(numFila, 6).setValue(esEstudio);
      hoja.getRange(numFila, 7).setValue(esTerapia);
      hoja.getRange(numFila, 8).setValue(porcentaje);
      hoja.getRange(numFila, 9).setValue(horasAPagar);
    });
  });
}

// ---- Utilidades internas ----

// Extrae el Creamos ID de un string "Nombre (CREAMOS_ID)"
function _extraerCreamos_ID(texto) {
  var match = String(texto).match(/\(([A-Z]{2,4}\d{6,})\)/);
  return match ? match[1] : "";
}

// Extrae solo el nombre limpio de "Nombre (CREAMOS_ID)"
function _extraerNombre(texto) {
  return String(texto).replace(/\s*\([A-Z]{2,4}\d{6,}\)\s*$/, "").trim();
}

function _fechaClave(fecha) {
  return fecha.getFullYear() + "-" +
    String(fecha.getMonth() + 1).padStart(2, "0") + "-" +
    String(fecha.getDate()).padStart(2, "0");
}

function _uuidsExistentes(hoja) {
  var mapa  = {};
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var uuid = datos[i][9]; // col J
    if (uuid) mapa[uuid] = true;
  }
  return mapa;
}

// Lee hoja DiasEstudio → { CREAMOS_ID: true }
function _cargarDiasEstudio(ss) {
  var hoja = ss.getSheetByName("DiasEstudio");
  var mapa = {};
  if (!hoja) return mapa;
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var nombre = datos[i][0];
    // Usa NombresCanonicos para obtener el ID
    if (nombre) mapa[nombre] = true;
  }
  return mapa;
}

// Lee hoja ListaTerapias → { nombre: true }
function _cargarTerapias(ss) {
  var hoja = ss.getSheetByName("ListaTerapias");
  var mapa = {};
  if (!hoja) return mapa;
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var nombre = datos[i][0];
    var recibe = datos[i][1]; // columna "Recibe Terapia (X)"
    if (nombre && recibe) mapa[nombre] = true;
  }
  return mapa;
}
