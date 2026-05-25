// ============================================================
// 02_Kobo.gs — Importar y emparejar asistencia desde KoboToolbox
//
// Cada registro Kobo es UN evento (entrada o salida).
// El campo "Participante" tiene formato: "Nombre (CREAMOS_ID)"
// Se emparejan por participante + día para calcular horas reales.
//
// Columnas de ASISTENCIA:
//  A CreAmosID  B Nombre  C Fecha_Registro  D Tipo
//  E Horas_Trabajadas  F Es_Dia_Estudio  G Es_Terapia
//  H Porcentaje_Pago   I Horas_A_Pagar   J UUID
// ============================================================

// Importa registros nuevos vía API de KoboToolbox.
function importarDesdeKobo() {
  var hoja = _hojaAsistencia();
  var url  = "https://kc.kobotoolbox.org/api/v2/assets/" +
             CFG.KOBO_ASSET_UID + "/data/?format=json&limit=5000";

  var resp = UrlFetchApp.fetch(url, {
    method: "GET",
    headers: { "Authorization": "Token " + CFG.KOBO_TOKEN },
    muteHttpExceptions: true,
  });

  if (resp.getResponseCode() !== 200) {
    SpreadsheetApp.getUi().alert("❌ Error Kobo: " + resp.getResponseCode());
    return;
  }

  var results    = JSON.parse(resp.getContentText()).results || [];
  var existentes = _uuidsExistentes(hoja);
  var nuevos     = [];

  results.forEach(function(r) {
    var uuid = r["_uuid"] || "";
    if (existentes[uuid]) return;

    var participante = r["Participante"] || "";
    var tipo         = r["Ingreso / Egreso"] || "";
    var timestamp    = new Date(r["start"] || r["_submission_time"]);

    if (!participante || isNaN(timestamp)) return;

    nuevos.push([
      _extraerID(participante),    // A
      _extraerNombre(participante),// B
      timestamp,                   // C
      tipo,                        // D
      "", "", "", "", "",          // E-I (se calculan al emparejar)
      uuid,                        // J
    ]);
  });

  if (nuevos.length > 0) {
    hoja.getRange(hoja.getLastRow() + 1, 1, nuevos.length, 10).setValues(nuevos);
  }

  emparejarAsistencia();
  SpreadsheetApp.getUi().alert("✅ " + nuevos.length + " registros nuevos importados.");
}

// Carga histórica: lee la hoja "DatosKobo" pegada en este Sheets.
function importarAsistenciaHistorica() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrig  = ss.getSheetByName("DatosKobo");
  var hojaDest  = _hojaAsistencia();

  if (!hojaOrig) {
    SpreadsheetApp.getUi().alert(
      "Crea una hoja llamada 'DatosKobo' y pega ahí el contenido\n" +
      "de la hoja DatosKobo del archivo Planilla___mi_eelo.xlsx"
    );
    return;
  }

  var datos      = hojaOrig.getDataRange().getValues();
  var existentes = _uuidsExistentes(hojaDest);
  var nuevos     = [];

  for (var i = 1; i < datos.length; i++) {
    var fila         = datos[i];
    var participante = String(fila[4] || ""); // col E
    var tipo         = String(fila[3] || ""); // col D
    var timestamp    = fila[0];               // col A (start)
    var uuid         = String(fila[11] || "");// col L

    if (!participante || !uuid || existentes[uuid]) continue;

    nuevos.push([
      _extraerID(participante),
      _extraerNombre(participante),
      timestamp,
      tipo,
      "", "", "", "", "",
      uuid,
    ]);
  }

  if (nuevos.length > 0) {
    hojaDest.getRange(hojaDest.getLastRow() + 1, 1, nuevos.length, 10).setValues(nuevos);
  }

  emparejarAsistencia();
  SpreadsheetApp.getUi().alert("✅ " + nuevos.length + " registros históricos importados.");
}

// Empareja cada entrada con su salida y calcula horas.
function emparejarAsistencia() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = _hojaAsistencia();
  var datos = hoja.getDataRange().getValues();

  var diasEstudio = _setDiasEstudio(ss);
  var terapias    = _setTerapias(ss);

  // Agrupa eventos por creAmosID + fecha
  var grupos = {};
  for (var i = 1; i < datos.length; i++) {
    var id        = String(datos[i][0]).trim();
    var tipo      = String(datos[i][3]).trim();
    var ts        = new Date(datos[i][2]);
    if (!id || isNaN(ts)) continue;

    var clave = id + "|" + _dClave(ts);
    if (!grupos[clave]) grupos[clave] = { entradas: [], salidas: [], filas: [] };
    if (tipo === CFG.KOBO_TIPO_ENTRADA) grupos[clave].entradas.push({ ts: ts, i: i + 1 });
    if (tipo === CFG.KOBO_TIPO_SALIDA)  grupos[clave].salidas.push( { ts: ts, i: i + 1 });
    grupos[clave].filas.push(i + 1);
  }

  Object.keys(grupos).forEach(function(clave) {
    var g = grupos[clave];
    if (!g.entradas.length || !g.salidas.length) return;

    g.entradas.sort(function(a,b){ return a.ts - b.ts; });
    g.salidas.sort( function(a,b){ return a.ts - b.ts; });

    var id    = clave.split("|")[0];
    var horas = Math.max(0,
      Math.round((g.salidas[g.salidas.length-1].ts - g.entradas[0].ts) / 36000) / 100
    );
    var esEstudio  = diasEstudio.has(id) ? "Sí" : "No";
    var esTerapia  = terapias.has(id)    ? "Sí" : "No";
    var pct        = esEstudio === "Sí"  ? 0    : 100;
    var horasAPagar = horas * (pct / 100);

    g.filas.forEach(function(r) {
      hoja.getRange(r, 5, 1, 5).setValues([[horas, esEstudio, esTerapia, pct, horasAPagar]]);
    });
  });
}

// ---- Utilidades ----

function _extraerID(txt) {
  var m = String(txt).match(/\(([A-Z]{2,4}\d{6,})\)/);
  return m ? m[1] : "";
}

function _extraerNombre(txt) {
  return String(txt).replace(/\s*\([A-Z]{2,4}\d{6,}\)\s*$/, "").trim();
}

function _dClave(d) {
  return d.getFullYear() + "-" +
    String(d.getMonth()+1).padStart(2,"0") + "-" +
    String(d.getDate()).padStart(2,"0");
}

function _uuidsExistentes(hoja) {
  var mapa = {}, datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) { if (datos[i][9]) mapa[datos[i][9]] = true; }
  return mapa;
}

function _setDiasEstudio(ss) {
  var set = new Set(), h = ss.getSheetByName("DiasEstudio");
  if (!h) return set;
  var d = h.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) { if (d[i][0]) set.add(String(d[i][0])); }
  return set;
}

function _setTerapias(ss) {
  var set = new Set(), h = ss.getSheetByName("ListaTerapias");
  if (!h) return set;
  var d = h.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) { if (d[i][1]) set.add(String(d[i][0])); }
  return set;
}

function _hojaAsistencia() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.HOJAS.ASISTENCIA);
}
