// ============================================================
// SISTEMA RRHH — Mi eelo
// Google Sheets: "Mi eelo · RRHH"
// ============================================================

const CFG = {
  ORG:          "Mi eelo",
  CORREO_ADMIN: "adrian@creamosguatemala.org",
  TARIFA_HORA:  25,
  HOJAS: {
    PARTICIPANTES: "PARTICIPANTES",
    ASISTENCIA:    "ASISTENCIA",
    FACTURACION:   "FACTURACION",
    DASHBOARD:     "DASHBOARD",
  },
  KOBO_TOKEN:        "TU_TOKEN_AQUI",
  KOBO_ASSET_UID:    "TU_ASSET_UID_AQUI",
  KOBO_TIPO_ENTRADA: "🟢 Entrada",
  KOBO_TIPO_SALIDA:  "🔴 Salida",
  MESES: ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
          "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"],
  // PARTICIPANTES: col 18 = URL_Doc_Proceso
  // FACTURACION:   col 14 = URL_Recibo
};

// ── Helpers ──────────────────────────────────────────────────

function _sh(nombre) {
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!h) throw new Error('Hoja "' + nombre + '" no encontrada. Ejecuta PASO 1 primero.');
  return h;
}
function _run(fn) {
  try { fn(); } catch(e) {
    try { _alert("❌ " + e.message); }
    catch(_) { Logger.log("❌ " + e.message); }
  }
}
function _alert(msg) {
  try { SpreadsheetApp.getUi().alert(msg); }
  catch(_) { Logger.log(msg); }
}
function autorizar() {
  Logger.log("✅ Autorizado: " + SpreadsheetApp.getActiveSpreadsheet().getName());
}

// ── Menú ─────────────────────────────────────────────────────

function onOpen() {
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch(_) { return; }
  ui.createMenu("👥 RRHH")
    .addItem("🏗️  PASO 1 — Crear hojas del sistema",    "crearHojas")
    .addItem("📁  PASO 2 — Crear estructura en Drive",  "crearEstructuraDrive")
    .addItem("⚙️  PASO 3 — Activar automatizaciones",   "configurarTriggers")
    .addSeparator()
    .addItem("➕ Nuevo participante",                   "nuevoParticipante")
    .addItem("📄 Generar/Actualizar DP (fila activa)",  "generarDpFilaActiva")
    .addItem("📄 Actualizar todos los DPs",             "actualizarTodosLosDps")
    .addSeparator()
    .addSubMenu(ui.createMenu("📂 Importar datos históricos")
      .addItem("👥 Importar participantes (IMPORT_PART)", "importarParticipantes")
      .addItem("📋 Importar asistencia (DatosKobo)",      "importarAsistenciaHistorica")
      .addItem("💰 Importar facturación (IMPORT_Mes)",    "importarFacturacionHistorica")
    )
    .addSeparator()
    .addItem("📥 Importar asistencia desde Kobo",       "importarDesdeKobo")
    .addItem("🔗 Emparejar entradas/salidas → horas",   "emparejarAsistencia")
    .addSeparator()
    .addItem("💰 Calcular facturación del mes",         "calcularFacturacionMes")
    .addItem("🧾 Generar recibos de pago",              "generarRecibosMes")
    .addItem("📊 Generar reporte mensual",              "generarReporteMensual")
    .addSeparator()
    .addItem("🔔 Recordatorio de pagos pendientes",     "enviarRecordatorioPagos")
    .addItem("📬 Resumen mensual al admin",             "enviarResumenMensual")
    .addSeparator()
    .addItem("🔄 Actualizar Dashboard",                 "actualizarDashboard")
    .addSeparator()
    .addItem("🗑️  Reinstalar sistema (borra TODO)",     "reinstalarSistema")
    .addToUi();
}

// Simple trigger — solo lee Sheets, no necesita autorización OAuth
function onEdit(e) {
  var nombre = e.range.getSheet().getName();
  if (nombre === CFG.HOJAS.FACTURACION && e.range.getColumn() === 11) {
    try { actualizarDashboard(); } catch(_) {}
  }
}

// ── Crear hojas ───────────────────────────────────────────────

function crearHojas() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // PARTICIPANTES — 18 columnas (A–R)
  // R = URL_Doc_Proceso (DP)
  var hP = ss.getSheetByName("PARTICIPANTES") || ss.insertSheet("PARTICIPANTES");
  if (hP.getLastRow() === 0) {
    hP.appendRow([
      "Creamos_ID","Nombre","Proyecto","Division","Programa","Estado","Etapa",
      "Educacion","Apoyo_Emocional","Inclusion_Laboral","Categoria",
      "DPI","NIT","Correo","Banco","Num_Cuenta","Forma_Pago","URL_Doc_Proceso"
    ]);
    _fmtEnc(hP, "#639922");
    hP.getRange("F2:F500").setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(["Activo","Inactivo","Egresado"], true).build());
    hP.setColumnWidth(18, 300);
  }

  // ASISTENCIA — 10 columnas (A–J)
  var hA = ss.getSheetByName("ASISTENCIA") || ss.insertSheet("ASISTENCIA");
  if (hA.getLastRow() === 0) {
    hA.appendRow([
      "Creamos_ID","Nombre","Fecha_Registro","Tipo",
      "Horas_Trabajadas","Es_Dia_Estudio","Es_Terapia","Porcentaje_Pago","Horas_A_Pagar","UUID_Kobo"
    ]);
    _fmtEnc(hA, "#1f54a8");
  }

  // FACTURACION — 14 columnas (A–N)
  // N = URL_Recibo
  var hF = ss.getSheetByName("FACTURACION") || ss.insertSheet("FACTURACION");
  if (hF.getLastRow() === 0) {
    hF.appendRow([
      "Creamos_ID","Nombre","Mes","Anio","Quincena",
      "Horas_Trabajadas","Monto_A_Pagar","Factura_Entregada","Numero_Factura",
      "Declaraguate","Pagado","Fecha_Pago","Comentarios","URL_Recibo"
    ]);
    _fmtEnc(hF, "#639922");
    hF.getRange("H2:H2000").setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(["Sí","No"], true).build());
    hF.getRange("J2:J2000").setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(["Sí","No"], true).build());
    hF.getRange("K2:K2000").setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(["Sí","No"], true).build());
    hF.setColumnWidth(14, 300);
  }

  // Borrar hoja vacía por defecto si existe
  ["Hoja 1","Sheet1"].forEach(function(n) {
    var h = ss.getSheetByName(n);
    if (h && ss.getSheets().length > 3) ss.deleteSheet(h);
  });

  ss.setActiveSheet(hP);
  _alert(
    "✅ Hojas creadas:\n• PARTICIPANTES (col R = URL Documento de Proceso)\n" +
    "• ASISTENCIA\n• FACTURACION (col N = URL Recibo)\n\n" +
    "Siguiente paso: PASO 2 — Crear estructura en Drive"
  );
}); }

function _fmtEnc(hoja, color) {
  var enc = hoja.getRange(1, 1, 1, hoja.getLastColumn());
  enc.setBackground(color).setFontColor("#ffffff").setFontWeight("bold");
  hoja.setFrozenRows(1);
}

// ── Nuevo participante ────────────────────────────────────────

function nuevoParticipante() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var r1 = ui.prompt("Nuevo participante", "Creamos_ID (ej. CR202401):", ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var id = r1.getResponseText().trim().toUpperCase();
  if (!id) { ui.alert("El ID no puede estar vacío."); return; }

  var r2 = ui.prompt("Nuevo participante", "Nombre completo:", ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var nombre = r2.getResponseText().trim();
  if (!nombre) { ui.alert("El nombre no puede estar vacío."); return; }

  var hP = _sh(CFG.HOJAS.PARTICIPANTES);
  var dp = hP.getDataRange().getValues();
  for (var i = 1; i < dp.length; i++) {
    if (String(dp[i][0]).trim() === id) { ui.alert("Ya existe un participante con ID " + id); return; }
  }

  // Crear DP primero para evitar race condition
  var carpeta = _carpetaDP();
  var doc = _abrirOCrearDocProceso(id, nombre, carpeta, "");
  var url = doc.getUrl();

  var part = {
    id: id, nombre: nombre, proyecto:"", division:"", programa:"", estado:"Activo",
    etapa:"", educacion:"", apoyoEmocional:"", inclusionLaboral:"", categoria:"",
    dpi:"", nit:"", correo:"", banco:"", numCuenta:"", formaPago:""
  };
  _escribirContenidoDP(doc, part, [], []);

  hP.appendRow([id, nombre, "", "", "", "Activo", "", "", "", "", "", "", "", "", "", "", "", url]);
  var fila = hP.getLastRow();
  hP.setActiveRange(hP.getRange(fila, 1));

  ui.alert("✅ Participante registrado: " + nombre + "\n\nDocumento de Proceso:\n" + url);
}); }

// ── Documentos de Proceso (DP) ────────────────────────────────

function generarDpFilaActiva() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getActiveSheet().getName() !== CFG.HOJAS.PARTICIPANTES) {
    _alert("Primero selecciona una fila en la hoja PARTICIPANTES."); return;
  }
  var fila = ss.getActiveRange().getRow();
  if (fila < 2) { _alert("Selecciona una fila de datos (no el encabezado)."); return; }
  var url = _sincronizarDP(_sh(CFG.HOJAS.PARTICIPANTES), fila);
  _alert("✅ DP actualizado.\n" + url);
}); }

function actualizarTodosLosDps() { _run(function() {
  var hP = _sh(CFG.HOJAS.PARTICIPANTES);
  var datos = hP.getDataRange().getValues();
  var n = 0, errores = [];
  for (var i = 1; i < datos.length; i++) {
    var id = String(datos[i][0]).trim();
    if (!id) continue;
    try { _sincronizarDP(hP, i + 1); n++; }
    catch(err) { errores.push(id + ": " + err.message); Logger.log("DP error " + id + ": " + err.message); }
  }
  var msg = "✅ " + n + " DPs actualizados.";
  if (errores.length) msg += "\n\n⚠️ Errores:\n" + errores.join("\n");
  _alert(msg);
}); }

function _sincronizarDP(hP, fila) {
  var datos = hP.getDataRange().getValues();
  var f = datos[fila - 1];
  var id     = String(f[0]).trim();
  var nombre = String(f[1]).trim();
  if (!id || !nombre) throw new Error("Fila " + fila + " no tiene ID o nombre válido.");

  var part = {
    id: id, nombre: nombre,
    proyecto:        String(f[2]  || ""), division:       String(f[3]  || ""),
    programa:        String(f[4]  || ""), estado:         String(f[5]  || ""),
    etapa:           String(f[6]  || ""), educacion:      String(f[7]  || ""),
    apoyoEmocional:  String(f[8]  || ""), inclusionLaboral:String(f[9] || ""),
    categoria:       String(f[10] || ""), dpi:            String(f[11] || ""),
    nit:             String(f[12] || ""), correo:         String(f[13] || ""),
    banco:           String(f[14] || ""), numCuenta:      String(f[15] || ""),
    formaPago:       String(f[16] || "")
  };
  var urlActual = String(f[17] || "");

  // Asistencia del participante (ordenada por fecha DESC)
  var hA = _sh(CFG.HOJAS.ASISTENCIA);
  var aRows = hA.getDataRange().getValues();
  var asistencia = [];
  for (var i = 1; i < aRows.length; i++) {
    if (String(aRows[i][0]).trim() !== id) continue;
    asistencia.push({
      fecha:       aRows[i][2], tipo:        aRows[i][3],
      horas:       aRows[i][4], esDiaEstudio:aRows[i][5],
      esTerapia:   aRows[i][6], pct:         aRows[i][7],
      horasAPagar: aRows[i][8]
    });
  }
  asistencia.sort(function(a,b){ return new Date(b.fecha) - new Date(a.fecha); });

  // Facturación del participante
  var hF = _sh(CFG.HOJAS.FACTURACION);
  var fRows = hF.getDataRange().getValues();
  var facturacion = [];
  for (var j = 1; j < fRows.length; j++) {
    if (String(fRows[j][0]).trim() !== id) continue;
    facturacion.push({
      mes: fRows[j][2], anio: fRows[j][3], quincena: fRows[j][4],
      horas: fRows[j][5], monto: fRows[j][6], facturaEntregada: fRows[j][7],
      numFactura: fRows[j][8], declaraguate: fRows[j][9],
      pagado: fRows[j][10], fechaPago: fRows[j][11]
    });
  }

  var carpeta = _carpetaDP();
  var doc = _abrirOCrearDocProceso(id, nombre, carpeta, urlActual);
  _escribirContenidoDP(doc, part, asistencia, facturacion);

  var urlNueva = doc.getUrl();
  if (urlNueva !== urlActual) hP.getRange(fila, 18).setValue(urlNueva);
  return urlNueva;
}

function _abrirOCrearDocProceso(id, nombre, carpeta, urlExistente) {
  var titulo = id + " — " + nombre + " — Documento de Proceso";
  var doc = null;
  if (urlExistente && urlExistente.startsWith("http")) {
    var m = urlExistente.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (m) {
      for (var i = 0; i < 3; i++) {
        try { doc = DocumentApp.openById(m[1]); break; }
        catch(_) { if (i < 2) Utilities.sleep(500); }
      }
    }
    if (!doc) throw new Error("No se pudo abrir el DP de " + nombre + ".\nURL: " + urlExistente);
    doc.setName(titulo);
  } else {
    doc = DocumentApp.create(titulo);
    DriveApp.getFileById(doc.getId()).moveTo(carpeta);
  }
  return doc;
}

function _escribirContenidoDP(doc, part, asistencia, facturacion) {
  var body = doc.getBody();
  body.clear();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);
  var tz  = Session.getScriptTimeZone();
  var hoy = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy");

  // ── Encabezado ──
  var pTit = body.appendParagraph("DOCUMENTO DE PROCESO");
  pTit.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  pTit.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  pTit.editAsText().setForegroundColor("#1a237e");

  var pNom = body.appendParagraph(part.nombre || "");
  pNom.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  pNom.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph("Mi eelo  ·  Actualizado: " + hoy)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .editAsText().setFontSize(9).setItalic(true).setForegroundColor("#888888");

  body.appendParagraph("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
      .editAsText().setFontSize(8).setForegroundColor("#cccccc");

  // ── 1. Datos del participante ──
  var s1 = body.appendParagraph("1.  DATOS DEL PARTICIPANTE");
  s1.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  s1.editAsText().setForegroundColor("#639922");

  var tDatos = body.appendTable([
    ["Creamos ID",      part.id          || "—",   "Estado",          part.estado          || "—"],
    ["Nombre completo", part.nombre      || "—",   "Etapa",           part.etapa           || "—"],
    ["Proyecto",        part.proyecto    || "—",   "Categoría",       part.categoria       || "—"],
    ["División",        part.division    || "—",   "Programa",        part.programa        || "—"],
    ["DPI",             part.dpi         || "—",   "NIT",             part.nit             || "—"],
    ["Correo",          part.correo      || "—",   "Forma de pago",   part.formaPago       || "—"],
    ["Banco",           part.banco       || "—",   "Núm. cuenta",     part.numCuenta       || "—"],
  ]);
  _estilTabla2col(tDatos, "#639922");

  // ── 2. Servicios de apoyo ──
  body.appendParagraph("");
  var s2 = body.appendParagraph("2.  SERVICIOS DE APOYO");
  s2.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  s2.editAsText().setForegroundColor("#1f54a8");

  var tApoyo = body.appendTable([
    ["Educación",         part.educacion        || "—"],
    ["Apoyo Emocional",   part.apoyoEmocional   || "—"],
    ["Inclusión Laboral", part.inclusionLaboral || "—"],
  ]);
  _estilTabla1col(tApoyo, "#1f54a8");

  // ── 3. Asistencia (últimos 90 días) ──
  body.appendParagraph("");
  var s3 = body.appendParagraph("3.  HISTORIAL DE ASISTENCIA  (últimos 90 días)");
  s3.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  s3.editAsText().setForegroundColor("#1f54a8");

  var hace90 = new Date(); hace90.setDate(hace90.getDate() - 90);
  var asistRec = asistencia.filter(function(a) {
    var d = new Date(a.fecha); return !isNaN(d) && d >= hace90 && a.tipo === CFG.KOBO_TIPO_ENTRADA;
  }).slice(0, 60);

  if (asistRec.length === 0) {
    body.appendParagraph("Sin registros en los últimos 90 días.")
        .editAsText().setItalic(true).setFontSize(9).setForegroundColor("#888888");
  } else {
    // Resumen de horas por mes
    var totalHrs = 0;
    var resumenMes = {};
    asistRec.forEach(function(a) {
      var d = new Date(a.fecha);
      var k = CFG.MESES[d.getMonth()] + " " + d.getFullYear();
      var h = parseFloat(a.horasAPagar) || 0;
      resumenMes[k] = (resumenMes[k] || 0) + h;
      totalHrs += h;
    });
    var filaAss = [["Fecha","Horas Trabajadas","Día Estudio","Terapia","Horas a Pagar"]];
    asistRec.forEach(function(a) {
      var fecha = a.fecha instanceof Date
        ? Utilities.formatDate(a.fecha, tz, "dd/MM/yyyy")
        : String(a.fecha || "—");
      filaAss.push([
        fecha,
        _n2(a.horas),
        String(a.esDiaEstudio || "—"),
        String(a.esTerapia    || "—"),
        _n2(a.horasAPagar)
      ]);
    });
    var tAss = body.appendTable(filaAss);
    _estilTablaEnc(tAss, "#1f54a8");
    body.appendParagraph("Total horas a pagar (período): " + _n2(totalHrs))
        .editAsText().setBold(true).setFontSize(10);
  }

  // ── 4. Historial de pagos ──
  body.appendParagraph("");
  var s4 = body.appendParagraph("4.  HISTORIAL DE PAGOS");
  s4.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  s4.editAsText().setForegroundColor("#639922");

  if (facturacion.length === 0) {
    body.appendParagraph("Sin registros de pago.")
        .editAsText().setItalic(true).setFontSize(9).setForegroundColor("#888888");
  } else {
    var filasPago = [["Mes","Año","Q","Horas","Monto","Factura","Declaraguate","Pagado","Fecha Pago"]];
    var totalPagado = 0, totalPend = 0;
    facturacion.slice(0, 30).forEach(function(p) {
      var fp = p.fechaPago instanceof Date
        ? Utilities.formatDate(p.fechaPago, tz, "dd/MM/yyyy")
        : String(p.fechaPago || "—");
      var m = parseFloat(p.monto) || 0;
      if (p.pagado === "Sí") totalPagado += m; else totalPend += m;
      filasPago.push([
        String(p.mes      || "—"),
        String(p.anio     || "—"),
        String(p.quincena || "—"),
        _n2(p.horas),
        "Q " + m.toFixed(2),
        String(p.facturaEntregada || "No"),
        String(p.declaraguate     || "No"),
        String(p.pagado           || "No"),
        fp
      ]);
    });
    var tPago = body.appendTable(filasPago);
    _estilTablaEnc(tPago, "#639922");

    body.appendParagraph("");
    body.appendTable([
      ["Total cobrado / pagado", "Q " + totalPagado.toFixed(2)],
      ["Pendiente de pago",      "Q " + totalPend.toFixed(2)],
    ]).editAsText().setFontSize(10);
  }

  // ── Pie de página ──
  body.appendParagraph("");
  body.appendParagraph("Generado por Sistema RRHH — " + CFG.ORG + "  ·  " + hoy)
      .setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
      .editAsText().setItalic(true).setFontSize(8).setForegroundColor("#aaaaaa");

  doc.saveAndClose();
}

function _n2(v) { var n = parseFloat(v); return isNaN(n) ? "0" : n.toFixed(2); }

function _estilTabla2col(tabla, color) {
  // Columnas 0 y 2 son encabezados (fondo color)
  for (var r = 0; r < tabla.getNumRows(); r++) {
    tabla.getRow(r).getCell(0).setBackgroundColor(color).editAsText().setForegroundColor("#ffffff").setBold(true);
    tabla.getRow(r).getCell(2).setBackgroundColor(color).editAsText().setForegroundColor("#ffffff").setBold(true);
  }
}
function _estilTabla1col(tabla, color) {
  for (var r = 0; r < tabla.getNumRows(); r++) {
    tabla.getRow(r).getCell(0).setBackgroundColor(color).editAsText().setForegroundColor("#ffffff").setBold(true);
  }
}
function _estilTablaEnc(tabla, color) {
  // Solo la primera fila es encabezado
  var enc = tabla.getRow(0);
  for (var c = 0; c < enc.getNumCells(); c++) {
    enc.getCell(c).setBackgroundColor(color).editAsText().setForegroundColor("#ffffff").setBold(true);
  }
}

// ── Kobo ─────────────────────────────────────────────────────
// Columnas ASISTENCIA: A CreAmosID  B Nombre  C Fecha  D Tipo
//   E Horas  F Dia_Estudio  G Terapia  H Pct  I HorasAPagar  J UUID

function importarDesdeKobo() { _run(function() {
  var hoja = _sh(CFG.HOJAS.ASISTENCIA);
  var url  = "https://kc.kobotoolbox.org/api/v2/assets/" +
             CFG.KOBO_ASSET_UID + "/data/?format=json&limit=5000";
  var resp = UrlFetchApp.fetch(url, {
    method: "GET",
    headers: { "Authorization": "Token " + CFG.KOBO_TOKEN },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error("Error Kobo HTTP " + resp.getResponseCode() + ": " + resp.getContentText().substring(0,200));
  }
  var results    = JSON.parse(resp.getContentText()).results || [];
  var existentes = _uuidsExistentes(hoja);
  var nuevos     = [];
  results.forEach(function(r) {
    var uuid = r["_uuid"] || "";
    if (existentes[uuid]) return;
    var p  = r["Participante"] || "";
    var t  = r["Ingreso / Egreso"] || r["tipo"] || "";
    var ts = new Date(r["start"] || r["_submission_time"]);
    if (!p || isNaN(ts)) return;
    nuevos.push([_extraerID(p), _extraerNombre(p), ts, t, "", "", "", "", "", uuid]);
  });
  if (nuevos.length > 0) {
    hoja.getRange(hoja.getLastRow() + 1, 1, nuevos.length, 10).setValues(nuevos);
  }
  emparejarAsistencia();
  _alert("✅ " + nuevos.length + " registros nuevos importados desde Kobo.");
}); }

// Pega el CSV/Excel de Kobo en hoja "DatosKobo" y ejecuta esto.
function importarAsistenciaHistorica() { _run(function() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrig = ss.getSheetByName("DatosKobo");
  if (!hojaOrig) throw new Error(
    'Crea una hoja llamada "DatosKobo" y pega ahí el contenido exportado de KoboToolbox.'
  );
  var hojaDest   = _sh(CFG.HOJAS.ASISTENCIA);
  var datos      = hojaOrig.getDataRange().getValues();
  var existentes = _uuidsExistentes(hojaDest);
  // Detecta columnas automáticamente buscando en la primera fila
  var enc = datos[0].map(function(v){ return String(v).toLowerCase(); });
  var colPart = _buscarCol(enc, ["participante","participant","nombre"]);
  var colTipo = _buscarCol(enc, ["ingreso","egreso","tipo","type"]);
  var colTs   = _buscarCol(enc, ["start","inicio","timestamp","fecha"]);
  var colUuid = _buscarCol(enc, ["_uuid","uuid","id"]);
  if (colPart < 0) throw new Error("No se encontró columna de participante en DatosKobo.");
  var nuevos = [];
  for (var i = 1; i < datos.length; i++) {
    var p    = String(datos[i][colPart]  || "");
    var tipo = String(datos[i][colTipo >= 0 ? colTipo : 3] || "");
    var ts   = datos[i][colTs >= 0 ? colTs : 0];
    var uuid = String(datos[i][colUuid >= 0 ? colUuid : enc.length - 1] || "");
    if (!p || !uuid || existentes[uuid]) continue;
    var d = ts instanceof Date ? ts : new Date(ts);
    if (isNaN(d)) continue;
    nuevos.push([_extraerID(p), _extraerNombre(p), d, tipo, "", "", "", "", "", uuid]);
  }
  if (nuevos.length > 0) {
    hojaDest.getRange(hojaDest.getLastRow() + 1, 1, nuevos.length, 10).setValues(nuevos);
  }
  emparejarAsistencia();
  _alert("✅ " + nuevos.length + " registros históricos importados.");
}); }

function _buscarCol(encabezados, palabras) {
  for (var i = 0; i < encabezados.length; i++) {
    for (var p = 0; p < palabras.length; p++) {
      if (encabezados[i].indexOf(palabras[p]) >= 0) return i;
    }
  }
  return -1;
}

// ── Emparejar entradas/salidas → calcular horas ───────────────

function emparejarAsistencia() { _run(function() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var hoja        = _sh(CFG.HOJAS.ASISTENCIA);
  var datos       = hoja.getDataRange().getValues();
  var diasEstudio = _setDiasEstudio(ss);
  var terapias    = _setTerapias(ss);

  // Agrupa por participante+día
  var grupos = {};
  for (var i = 1; i < datos.length; i++) {
    var id   = String(datos[i][0]).trim();
    var tipo = String(datos[i][3]).trim();
    var ts   = new Date(datos[i][2]);
    if (!id || isNaN(ts)) continue;
    var clave = id + "|" + _dClave(ts);
    if (!grupos[clave]) grupos[clave] = { ent:[], sal:[], filas:[] };
    if (tipo === CFG.KOBO_TIPO_ENTRADA) grupos[clave].ent.push({ ts: ts });
    if (tipo === CFG.KOBO_TIPO_SALIDA)  grupos[clave].sal.push({ ts: ts });
    grupos[clave].filas.push(i + 1);
  }

  Object.keys(grupos).forEach(function(clave) {
    var g = grupos[clave];
    if (!g.ent.length || !g.sal.length) return;
    g.ent.sort(function(a,b){ return a.ts - b.ts; });
    g.sal.sort(function(a,b){ return a.ts - b.ts; });
    var id   = clave.split("|")[0];
    // Horas = última salida - primera entrada, redondeado a 2 decimales
    var ms   = g.sal[g.sal.length - 1].ts - g.ent[0].ts;
    var hrs  = Math.max(0, Math.round(ms / 36000) / 100);
    var esE  = diasEstudio.has(id) ? "Sí" : "No";
    var esT  = terapias.has(id)    ? "Sí" : "No";
    // Días de estudio y terapias NO se pagan (porcentaje = 0)
    var pct  = (esE === "Sí" || esT === "Sí") ? 0 : 100;
    var hap  = Math.round(hrs * (pct / 100) * 100) / 100;
    g.filas.forEach(function(r) {
      hoja.getRange(r, 5, 1, 5).setValues([[hrs, esE, esT, pct, hap]]);
    });
  });
  _alert("✅ Asistencia emparejada. Horas calculadas correctamente.");
}); }

function _extraerID(txt) {
  var m = String(txt).match(/\(([A-Z]{2,4}\d{4,})\)/);
  return m ? m[1] : "";
}
function _extraerNombre(txt) {
  return String(txt).replace(/\s*\([A-Z]{2,4}\d{4,}\)\s*$/, "").trim();
}
function _dClave(d) {
  return d.getFullYear() + "-" +
    String(d.getMonth() + 1).padStart(2, "0") + "-" +
    String(d.getDate()).padStart(2, "0");
}
function _uuidsExistentes(hoja) {
  var m = {}, d = hoja.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) { if (d[i][9]) m[String(d[i][9])] = true; }
  return m;
}
function _setDiasEstudio(ss) {
  var s = new Set(), h = ss.getSheetByName("DiasEstudio");
  if (!h) return s;
  var d = h.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) { if (d[i][0]) s.add(String(d[i][0])); }
  return s;
}
function _setTerapias(ss) {
  var s = new Set(), h = ss.getSheetByName("ListaTerapias");
  if (!h) return s;
  var d = h.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) { if (d[i][0]) s.add(String(d[i][0])); }
  return s;
}

// ── Facturación ───────────────────────────────────────────────
// Columnas FACTURACION:
//  A CreAmosID  B Nombre  C Mes  D Anio  E Quincena
//  F Horas  G Monto  H Factura_Entregada  I Num_Factura
//  J Declaraguate  K Pagado  L Fecha_Pago  M Comentarios  N URL_Recibo

function calcularFacturacionMes() { _run(function() {
  var ahora = new Date();
  _calcular(ahora.getMonth() + 1, ahora.getFullYear());
}); }

function _calcular(mes, anio) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var hojaA     = _sh(CFG.HOJAS.ASISTENCIA);
  var hojaP     = _sh(CFG.HOJAS.PARTICIPANTES);
  var hojaF     = _sh(CFG.HOJAS.FACTURACION);
  var nombreMes = CFG.MESES[mes - 1];

  // Sumar horas_a_pagar por participante y quincena
  var horas = { "1": {}, "2": {} };
  var asist = hojaA.getDataRange().getValues();
  for (var i = 1; i < asist.length; i++) {
    var id   = String(asist[i][0]).trim();
    var tipo = String(asist[i][3]).trim();
    var ts   = new Date(asist[i][2]);
    var hap  = parseFloat(asist[i][8]) || 0;
    // Solo contar entradas para no duplicar
    if (!id || tipo !== CFG.KOBO_TIPO_ENTRADA) continue;
    if (isNaN(ts) || ts.getMonth() + 1 !== mes || ts.getFullYear() !== anio) continue;
    var q = ts.getDate() <= 15 ? "1" : "2";
    horas[q][id] = Math.round(((horas[q][id] || 0) + hap) * 100) / 100;
  }

  var part   = hojaP.getDataRange().getValues();
  var fact   = hojaF.getDataRange().getValues();
  var creados = 0, actualizados = 0;

  ["1", "2"].forEach(function(q) {
    part.slice(1).forEach(function(p) {
      var id     = String(p[0]).trim();
      var nombre = String(p[1]).trim();
      if (!id || !nombre) return;
      var h = Math.round((horas[q][id] || 0) * 100) / 100;
      var m = Math.round(h * CFG.TARIFA_HORA * 100) / 100;
      // Buscar fila existente
      var filaExistente = -1;
      for (var j = 1; j < fact.length; j++) {
        if (String(fact[j][0]) === id && fact[j][2] === nombreMes &&
            Number(fact[j][3]) === anio && String(fact[j][4]) === q) {
          filaExistente = j + 1; break;
        }
      }
      if (filaExistente > 0) {
        hojaF.getRange(filaExistente, 6, 1, 2).setValues([[h, m]]);
        actualizados++;
      } else {
        hojaF.appendRow([id, nombre, nombreMes, anio, q, h, m, "No", "", "No", "No", "", "", ""]);
        creados++;
      }
    });
  });

  actualizarDashboard();
  _alert(
    "✅ Facturación calculada — " + nombreMes + " " + anio + "\n" +
    "• " + creados + " registros nuevos\n• " + actualizados + " actualizados\n\n" +
    "Tarifa: Q " + CFG.TARIFA_HORA + "/hora"
  );
}

// Pega cada hoja del Excel como "IMPORT_Mayo", "IMPORT_Junio"... y ejecuta esto.
function importarFacturacionHistorica() { _run(function() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hojaF = _sh(CFG.HOJAS.FACTURACION);
  var lista = [
    {nombre:"Mayo",anio:2025},{nombre:"Junio",anio:2025},{nombre:"Julio",anio:2025},
    {nombre:"Agosto",anio:2025},{nombre:"Septiembre",anio:2025},{nombre:"Octubre",anio:2025},
    {nombre:"Noviembre",anio:2025},{nombre:"Diciembre",anio:2025},
    {nombre:"Enero",anio:2026},{nombre:"Febrero",anio:2026},
    {nombre:"Marzo",anio:2026},{nombre:"Abril",anio:2026},
  ];
  var total = 0;
  lista.forEach(function(m) {
    var ht = ss.getSheetByName("IMPORT_" + m.nombre);
    if (!ht) return;
    var datos  = ht.getDataRange().getValues();
    var inicio = _filaEnc(datos) + 1;
    for (var i = inicio; i < datos.length; i++) {
      var fila   = datos[i];
      var nombre = String(fila[0] || "").trim();
      if (!nombre) continue;
      var q1 = _primerNum(fila, 1, 6), q2 = _primerNum(fila, 5, 12);
      if (q1) { hojaF.appendRow(["", nombre, m.nombre, m.anio, "1", "", q1, "No", "", "No", "No", "", "", ""]); total++; }
      if (q2) { hojaF.appendRow(["", nombre, m.nombre, m.anio, "2", "", q2, "No", "", "No", "No", "", "", ""]); total++; }
    }
  });
  _alert("✅ " + total + " registros de facturación histórica importados.");
}); }

function _filaEnc(datos) {
  var max = 0, idx = 0;
  for (var i = 0; i < Math.min(datos.length, 5); i++) {
    var c = datos[i].filter(function(v){ return v !== null && v !== ""; }).length;
    if (c > max) { max = c; idx = i; }
  }
  return idx;
}
function _primerNum(fila, desde, hasta) {
  for (var i = desde; i < Math.min(fila.length, hasta); i++) {
    var v = parseFloat(fila[i]);
    if (!isNaN(v) && v > 0) return v;
  }
  return null;
}

// ── Recibos de pago ───────────────────────────────────────────

function generarRecibosMes() { _run(function() {
  var ahora     = new Date();
  var nombreMes = CFG.MESES[ahora.getMonth()];
  var anio      = ahora.getFullYear();
  var hojaF     = _sh(CFG.HOJAS.FACTURACION);
  var datos     = hojaF.getDataRange().getValues();
  var carpeta   = _carpetaRecibos(anio, nombreMes);
  var generados = 0;

  for (var i = 1; i < datos.length; i++) {
    var f = datos[i];
    if (f[2] !== nombreMes || Number(f[3]) !== anio) continue;
    if (!f[6] || parseFloat(f[6]) === 0) continue;
    var urlActual = String(f[13] || "");
    var doc = _crearOActualizarRecibo(f, carpeta, urlActual);
    var urlNueva = doc.getUrl();
    if (urlNueva !== urlActual) hojaF.getRange(i + 1, 14).setValue(urlNueva);
    generados++;
  }
  _alert("✅ " + generados + " recibos generados/actualizados en Drive.");
}); }

function _crearOActualizarRecibo(f, carpeta, urlExistente) {
  var id     = String(f[0] || "").trim();
  var nombre = String(f[1] || "").trim();
  var mes    = f[2], anio = f[3], q = f[4], monto = parseFloat(f[6]) || 0;
  var titulo = "Recibo_" + (id || nombre.replace(/\s/g,"_")) + "_Q" + q + "_" + mes + "_" + anio;
  var doc = null;

  if (urlExistente && urlExistente.startsWith("http")) {
    var m = urlExistente.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (m) { try { doc = DocumentApp.openById(m[1]); } catch(_) {} }
    if (doc) { doc.setName(titulo); } else { doc = null; }
  }
  if (!doc) {
    doc = DocumentApp.create(titulo);
    DriveApp.getFileById(doc.getId()).moveTo(carpeta);
  }

  var body = doc.getBody();
  body.clear();
  body.setMarginTop(54).setMarginBottom(54).setMarginLeft(72).setMarginRight(72);
  var tz  = Session.getScriptTimeZone();
  var hoy = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy");

  var pTit = body.appendParagraph(CFG.ORG + " — Recibo de Pago");
  pTit.setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  pTit.editAsText().setForegroundColor("#1a237e");

  body.appendParagraph("").setSpacingAfter(8);

  var t = body.appendTable([
    ["Participante",  nombre || "—"],
    ["Creamos ID",    id     || "—"],
    ["Período",       String(mes) + " " + String(anio) + " — Quincena " + String(q)],
    ["Horas trabajadas", _n2(f[5])],
    ["Monto a pagar", "Q " + monto.toFixed(2)],
    ["Factura entregada", String(f[7] || "No")],
    ["Número factura", String(f[8] || "—")],
    ["Declaraguate",  String(f[9] || "No")],
    ["Pagado",        String(f[10] || "No")],
    ["Fecha de pago", f[11] instanceof Date ? Utilities.formatDate(f[11], tz, "dd/MM/yyyy") : String(f[11] || "—")],
  ]);
  // Destacar fila de monto
  t.getRow(4).editAsText().setBold(true).setFontSize(13);
  t.getRow(4).getCell(0).setBackgroundColor("#1a237e").editAsText().setForegroundColor("#ffffff");
  t.getRow(4).getCell(1).setBackgroundColor("#e8eaf6");
  // Encabezados de la tabla
  for (var r2 = 0; r2 < t.getNumRows(); r2++) {
    t.getRow(r2).getCell(0).editAsText().setBold(true);
  }

  body.appendParagraph("").setSpacingAfter(12);
  body.appendParagraph("Emisión: " + hoy)
      .setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
      .editAsText().setItalic(true).setFontSize(9).setForegroundColor("#888888");

  doc.saveAndClose();
  return doc;
}

// ── Reporte mensual ───────────────────────────────────────────

function generarReporteMensual() { _run(function() {
  var ahora     = new Date();
  var nombreMes = CFG.MESES[ahora.getMonth()];
  var anio      = ahora.getFullYear();
  var hojaF     = _sh(CFG.HOJAS.FACTURACION);
  var datos     = hojaF.getDataRange().getValues();
  var carpeta   = _carpetaReportes(anio);
  var titulo    = "Reporte_" + nombreMes + "_" + anio;

  // Borrar versión anterior si existe
  var prev = carpeta.getFilesByName(titulo);
  while (prev.hasNext()) prev.next().setTrashed(true);

  var doc  = DocumentApp.create(titulo);
  var body = doc.getBody();
  var tz   = Session.getScriptTimeZone();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);

  var pTit = body.appendParagraph(CFG.ORG + " — Reporte RRHH " + nombreMes + " " + anio);
  pTit.setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  pTit.editAsText().setForegroundColor("#1a237e");
  body.appendParagraph("");

  // Agrupar por participante
  var agrup = {};
  datos.slice(1).forEach(function(f) {
    if (f[2] !== nombreMes || Number(f[3]) !== anio || !String(f[1]).trim()) return;
    var n = String(f[1]).trim();
    if (!agrup[n]) agrup[n] = { q1: 0, hrs1: 0, q2: 0, hrs2: 0, pagado: "No" };
    if (String(f[4]) === "1") { agrup[n].q1 = parseFloat(f[6]) || 0; agrup[n].hrs1 = parseFloat(f[5]) || 0; }
    if (String(f[4]) === "2") { agrup[n].q2 = parseFloat(f[6]) || 0; agrup[n].hrs2 = parseFloat(f[5]) || 0; }
    if (f[10] === "Sí") agrup[n].pagado = "Sí";
  });

  var filas = [["Participante", "Hrs Q1", "Q1 (Q)", "Hrs Q2", "Q2 (Q)", "Total (Q)", "Pagado"]];
  var totalM = 0, pendientes = 0;
  Object.keys(agrup).sort().forEach(function(n) {
    var a   = agrup[n];
    var tot = a.q1 + a.q2;
    totalM += tot;
    if (a.pagado !== "Sí") pendientes++;
    filas.push([
      n,
      _n2(a.hrs1), a.q1.toFixed(2),
      _n2(a.hrs2), a.q2.toFixed(2),
      tot.toFixed(2),
      a.pagado === "Sí" ? "✓ Pagado" : "Pendiente"
    ]);
  });

  var tabla = body.appendTable(filas);
  _estilTablaEnc(tabla, "#1f54a8");

  body.appendParagraph("");
  body.appendTable([
    ["Total a pagar",         "Q " + totalM.toFixed(2)],
    ["Pagos pendientes",      pendientes + " participante(s)"],
    ["Tarifa hora",           "Q " + CFG.TARIFA_HORA],
    ["Generado",              Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm")],
  ]);

  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(carpeta);
  _alert("✅ Reporte generado:\n" + doc.getUrl());
}); }

// ── Drive ─────────────────────────────────────────────────────

function crearEstructuraDrive() { _run(function() {
  var raiz     = _getOCreate(null,  CFG.ORG + " · RRHH");
  var docs     = _getOCreate(raiz,  "Docs_Proceso");
  var recibos  = _getOCreate(raiz,  "Recibos");
  var reportes = _getOCreate(raiz,  "Reportes");
  var p = PropertiesService.getScriptProperties();
  p.setProperty("RRHH_RAIZ",     raiz.getId());
  p.setProperty("RRHH_DOCS_DP",  docs.getId());
  p.setProperty("RRHH_RECIBOS",  recibos.getId());
  p.setProperty("RRHH_REPORTES", reportes.getId());
  _alert("✅ Estructura Drive creada:\n" + raiz.getUrl() +
         "\n\n• Docs_Proceso/ — un DP por participante" +
         "\n• Recibos/       — recibos por año/mes" +
         "\n• Reportes/      — reportes mensuales");
}); }

function _carpetaDP() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty("RRHH_DOCS_DP"), r;
  if (id) { try { r = DriveApp.getFolderById(id); } catch(_) {} }
  if (!r) {
    crearEstructuraDrive();
    r = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty("RRHH_DOCS_DP"));
  }
  return r;
}
function _carpetaRecibos(anio, mes) {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty("RRHH_RECIBOS"), r;
  if (id) { try { r = DriveApp.getFolderById(id); } catch(_) {} }
  if (!r) {
    crearEstructuraDrive();
    r = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty("RRHH_RECIBOS"));
  }
  return _getOCreate(_getOCreate(r, String(anio)), mes);
}
function _carpetaReportes(anio) {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty("RRHH_REPORTES"), r;
  if (id) { try { r = DriveApp.getFolderById(id); } catch(_) {} }
  if (!r) {
    crearEstructuraDrive();
    r = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty("RRHH_REPORTES"));
  }
  return _getOCreate(r, String(anio));
}
function _getOCreate(padre, nombre) {
  var it = padre ? padre.getFoldersByName(nombre) : DriveApp.getFoldersByName(nombre);
  if (it.hasNext()) return it.next();
  return padre ? padre.createFolder(nombre) : DriveApp.createFolder(nombre);
}

// ── Reinstalar sistema ────────────────────────────────────────

function reinstalarSistema() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    "⚠️  Reinstalar sistema RRHH — BORRADO COMPLETO",
    "Esto elimina TODO:\n" +
    "• Hojas PARTICIPANTES, ASISTENCIA, FACTURACION, DASHBOARD\n" +
    "• Carpeta «" + CFG.ORG + " · RRHH» y TODO su contenido en Drive\n" +
    "  (Documentos de Proceso, Recibos, Reportes)\n\n" +
    "Esta acción NO se puede deshacer.\n¿Continuar?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [CFG.HOJAS.PARTICIPANTES, CFG.HOJAS.ASISTENCIA, CFG.HOJAS.FACTURACION, CFG.HOJAS.DASHBOARD].forEach(function(n) {
    var h = ss.getSheetByName(n);
    if (h) ss.deleteSheet(h);
  });

  var p = PropertiesService.getScriptProperties();
  var idRaiz = p.getProperty("RRHH_RAIZ");
  if (idRaiz) {
    try { _borrarCarpetaRecursivo(DriveApp.getFolderById(idRaiz)); } catch(_) {}
  } else {
    var it = DriveApp.getFoldersByName(CFG.ORG + " · RRHH");
    while (it.hasNext()) _borrarCarpetaRecursivo(it.next());
  }

  // Eliminar triggers instalados
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });

  p.deleteAllProperties();
  Utilities.sleep(500);
  crearHojas();
  _alert("✅ Sistema reinstalado desde cero.\nSiguiente paso: PASO 2 — Crear estructura en Drive.");
}); }

function _borrarCarpetaRecursivo(carpeta) {
  var archivos = carpeta.getFiles();
  while (archivos.hasNext()) archivos.next().setTrashed(true);
  var subs = carpeta.getFolders();
  while (subs.hasNext()) _borrarCarpetaRecursivo(subs.next());
  carpeta.setTrashed(true);
}

// ── Notificaciones ────────────────────────────────────────────

function enviarRecordatorioPagos() { _run(function() {
  var hojaF = _sh(CFG.HOJAS.FACTURACION);
  var hojaP = _sh(CFG.HOJAS.PARTICIPANTES);
  var datos = hojaF.getDataRange().getValues();
  var mes   = CFG.MESES[new Date().getMonth()];
  var anio  = new Date().getFullYear();
  var corr  = _mapaCorreos(hojaP);
  var n     = 0;
  for (var i = 1; i < datos.length; i++) {
    var f = datos[i];
    if (f[2] !== mes || Number(f[3]) !== anio) continue;
    if (f[10] === "Sí" || !f[6] || parseFloat(f[6]) === 0) continue;
    var c = corr[String(f[0]).trim()];
    if (!c) continue;
    MailApp.sendEmail({
      to:      c,
      subject: "[" + CFG.ORG + "] Pago pendiente — " + mes + " " + anio,
      body:    "Hola " + f[1] + ",\n\nTienes pago pendiente:\n" +
               "Quincena " + f[4] + " — Q " + parseFloat(f[6]).toFixed(2) + "\n\n" +
               "Por favor entrega tu factura para procesar el pago.\n\nSaludos,\n" + CFG.ORG
    });
    n++;
  }
  _alert("✅ " + n + " recordatorios enviados.");
}); }

function enviarResumenMensual() { _run(function() {
  var hojaF = _sh(CFG.HOJAS.FACTURACION);
  var datos = hojaF.getDataRange().getValues();
  var mes   = CFG.MESES[new Date().getMonth()];
  var anio  = new Date().getFullYear();
  var total = 0, pag = 0, pend = 0, lineas = [];
  for (var i = 1; i < datos.length; i++) {
    var f = datos[i];
    if (f[2] !== mes || Number(f[3]) !== anio) continue;
    var m = parseFloat(f[6]) || 0;
    total += m;
    if (f[10] === "Sí") { pag++; } else { pend++; }
    lineas.push("  " + f[1] + "  Q" + f[4] + " — Q " + m.toFixed(2) + " — " + (f[10] === "Sí" ? "PAGADO" : "PENDIENTE"));
  }
  MailApp.sendEmail({
    to:      CFG.CORREO_ADMIN,
    subject: "[" + CFG.ORG + "] Resumen RRHH — " + mes + " " + anio,
    body:    "Resumen " + mes + " " + anio + "\n══════════════════\n\n" +
             lineas.join("\n") +
             "\n\n══════════════════\n" +
             "Total: Q " + total.toFixed(2) + "\n" +
             "Pagados: " + pag + "\n" +
             "Pendientes: " + pend
  });
  _alert("✅ Resumen enviado a " + CFG.CORREO_ADMIN);
}); }

function _mapaCorreos(hojaP) {
  var m = {}, d = hojaP.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    var id = String(d[i][0]).trim(), c = d[i][13];
    if (id && c) m[id] = c;
  }
  return m;
}

// ── Importar participantes ────────────────────────────────────

function importarParticipantes() { _run(function() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var hojaTemp = ss.getSheetByName("IMPORT_PART");
  if (!hojaTemp) throw new Error(
    'Crea una hoja llamada "IMPORT_PART" y pega ahí la lista de participantes.'
  );
  var hojaP  = _sh(CFG.HOJAS.PARTICIPANTES);
  var datos  = hojaTemp.getDataRange().getValues();
  var exis   = _idsExistentes(hojaP);
  var n      = 0;
  for (var i = 1; i < datos.length; i++) {
    var f = datos[i];
    if (!String(f[1] || "").trim()) continue;
    var id = String(f[0] || "").trim();
    if (id && exis[id]) continue;
    hojaP.appendRow([
      id, f[1], f[2], f[3], f[4], f[5] || "Activo", f[6],
      f[7], f[8], f[9], f[10],
      "", "", "", "", "", "", ""
    ]);
    n++;
  }
  _alert("✅ " + n + " participantes importados.");
}); }

function _idsExistentes(hoja) {
  var m = {}, d = hoja.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    var id = String(d[i][0]).trim();
    if (id) m[id] = true;
  }
  return m;
}

// ── Dashboard ─────────────────────────────────────────────────

function actualizarDashboard() { _run(function() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName(CFG.HOJAS.DASHBOARD) || ss.insertSheet(CFG.HOJAS.DASHBOARD, 0);
  var mes  = CFG.MESES[new Date().getMonth()];
  var anio = new Date().getFullYear();
  var hojaP = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES);
  var hojaF = ss.getSheetByName(CFG.HOJAS.FACTURACION);
  var activos = 0, totalM = 0, pag = 0, pend = 0;
  if (hojaP) {
    var dp = hojaP.getDataRange().getValues();
    for (var i = 1; i < dp.length; i++) {
      if (String(dp[i][5]).toLowerCase() === "activo") activos++;
    }
  }
  if (hojaF) {
    var df = hojaF.getDataRange().getValues();
    for (var j = 1; j < df.length; j++) {
      if (df[j][2] !== mes || Number(df[j][3]) !== anio) continue;
      totalM += parseFloat(df[j][6]) || 0;
      if (df[j][10] === "Sí") { pag++; } else { pend++; }
    }
  }
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  var filas = [
    ["MÉTRICA", "VALOR"],
    ["Participantes activos",              activos],
    ["Total a pagar (" + mes + " " + anio + ")", "Q " + totalM.toFixed(2)],
    ["Quincenas pagadas",                  pag],
    ["Quincenas pendientes",               pend],
    ["", ""],
    ["Actualizado", ts],
  ];
  dash.clearContents();
  dash.getRange(1, 1, filas.length, 2).setValues(filas);
  dash.getRange(1, 1, 1, 2).setBackground("#639922").setFontColor("#fff").setFontWeight("bold");
  if (pend > 0) {
    dash.getRange(5, 2).setBackground("#fce8e6").setFontColor("#c62828").setFontWeight("bold");
  }
  dash.autoResizeColumns(1, 2);
}); }

// ── Triggers ─────────────────────────────────────────────────

function configurarTriggers() { _run(function() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("importarDesdeKobo").timeBased().everyHours(1).create();
  ScriptApp.newTrigger("calcularFacturacionMes").timeBased().everyDays(1).atHour(0).create();
  ScriptApp.newTrigger("actualizarDashboard").timeBased().everyMinutes(30).create();
  ScriptApp.newTrigger("enviarResumenMensual").timeBased().onMonthDay(1).atHour(8).create();
  ScriptApp.newTrigger("enviarRecordatorioPagos").timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(9).create();
  _alert(
    "✅ Automatizaciones activas:\n" +
    "• Kobo → asistencia: cada hora\n" +
    "• Facturación: diario 00:00\n" +
    "• Dashboard: cada 30 min\n" +
    "• Resumen al admin: día 1 del mes a las 08:00\n" +
    "• Recordatorio pagos: viernes a las 09:00"
  );
}); }
