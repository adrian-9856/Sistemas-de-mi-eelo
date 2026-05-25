// ============================================================
// 04_Documentos.gs — Genera recibos de pago y reportes en Docs
// ============================================================

// Genera un recibo de pago en Google Docs para cada participante
// del mes actual con monto > 0.
function generarRecibosMes() {
  var ahora    = new Date();
  var mes      = ahora.getMonth() + 1;
  var anio     = ahora.getFullYear();
  var nombreMes = CFG.MESES[mes - 1];

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hojaF = ss.getSheetByName(CFG.HOJAS.FACTURACION);
  var datos = hojaF.getDataRange().getValues();

  var carpeta   = _carpetaFacturas(anio, nombreMes);
  var generados = 0;

  for (var i = 1; i < datos.length; i++) {
    var f = datos[i];
    if (f[2] !== nombreMes || f[3] !== anio) continue;
    if (!f[6] || f[6] === 0) continue;

    var doc = _crearRecibo(f, carpeta);
    hojaF.getRange(i + 1, 14).setValue(doc.getUrl()); // col N
    generados++;
  }

  SpreadsheetApp.getUi().alert(
    "✅ " + generados + " recibos generados\nCarpeta: " + carpeta.getName()
  );
}

// Genera un reporte mensual consolidado en Google Docs.
function generarReporteMensual() {
  var ahora    = new Date();
  var nombreMes = CFG.MESES[ahora.getMonth()];
  var anio     = ahora.getFullYear();

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hojaF = ss.getSheetByName(CFG.HOJAS.FACTURACION);
  var datos = hojaF.getDataRange().getValues();

  var carpeta = _carpetaReportes(anio);
  var titulo  = "Reporte_" + nombreMes + "_" + anio;

  var prev = carpeta.getFilesByName(titulo);
  while (prev.hasNext()) prev.next().setTrashed(true);

  var doc  = DocumentApp.create(titulo);
  var body = doc.getBody();
  var tz   = Session.getScriptTimeZone();

  var h1 = body.appendParagraph(CFG.ORG + " — Reporte " + nombreMes + " " + anio);
  h1.setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph("").setSpacingAfter(6);

  var totalH = 0, totalM = 0, pendientes = 0;
  var filas  = [["Participante","Q1 (Q)","Q2 (Q)","Total (Q)","Pagado"]];
  var agrup  = {};

  datos.slice(1).forEach(function(f) {
    if (f[2] !== nombreMes || f[3] !== anio || !f[1]) return;
    var n = f[1];
    if (!agrup[n]) agrup[n] = { q1: 0, q2: 0, pagado: f[10] };
    if (String(f[4]) === "1") agrup[n].q1 = f[6] || 0;
    if (String(f[4]) === "2") agrup[n].q2 = f[6] || 0;
  });

  Object.keys(agrup).sort().forEach(function(n) {
    var a = agrup[n];
    var tot = a.q1 + a.q2;
    totalM += tot;
    if (a.pagado !== "Sí") pendientes++;
    filas.push([n, a.q1.toFixed(2), a.q2.toFixed(2), tot.toFixed(2),
                a.pagado === "Sí" ? "✓" : "Pendiente"]);
  });

  var tabla = body.appendTable(filas);
  tabla.getRow(0).editAsText().setBold(true).setForegroundColor("#ffffff");
  for (var c = 0; c < 5; c++) {
    tabla.getRow(0).getCell(c).setBackgroundColor("#1f54a8");
  }

  body.appendParagraph("").setSpacingAfter(4);
  body.appendTable([
    ["Total a pagar",    "Q " + totalM.toFixed(2)],
    ["Pagos pendientes", pendientes + " participantes"],
    ["Generado",         Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm")],
  ]);

  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(carpeta);

  SpreadsheetApp.getUi().alert("✅ Reporte generado\n" + doc.getUrl());
}

// ---- Recibo individual ----

function _crearRecibo(f, carpeta) {
  var id     = f[0], nombre = f[1], mes = f[2], anio = f[3];
  var q      = f[4], monto  = f[6];
  var titulo = "Recibo_" + (id || nombre.replace(/\s/g,"_")) + "_Q" + q + "_" + mes + "_" + anio;

  var prev = carpeta.getFilesByName(titulo);
  while (prev.hasNext()) prev.next().setTrashed(true);

  var doc  = DocumentApp.create(titulo);
  var body = doc.getBody();

  var h1 = body.appendParagraph(CFG.ORG + " — Recibo de Pago");
  h1.setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .editAsText().setForegroundColor("#1a237e");

  body.appendParagraph("").setSpacingAfter(6);

  var tabla = body.appendTable([
    ["Participante",     nombre],
    ["Creamos ID",       id || "—"],
    ["Período",          mes + " " + anio + " — Quincena " + q],
    ["Monto a pagar",    "Q " + parseFloat(monto).toFixed(2)],
    ["Estado factura",   f[7] || "No entregada"],
    ["Declaraguate",     f[9] || "No"],
    ["Pagado",           f[10] || "No"],
  ]);

  // Resalta el monto
  tabla.getRow(3).editAsText().setBold(true);

  body.appendParagraph("").setSpacingAfter(8);
  body.appendParagraph(
    "Emisión: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy")
  ).setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setItalic(true).setFontSize(9);

  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(carpeta);
  return doc;
}

// ---- Carpetas Drive ----

function _carpetaFacturas(anio, mes) {
  var props = PropertiesService.getScriptProperties();
  var id    = props.getProperty("DRIVE_FACTURAS");
  var raiz;
  if (id) {
    try { raiz = DriveApp.getFolderById(id); } catch(e) { raiz = null; }
  }
  if (!raiz) { crearEstructuraDrive(); raiz = DriveApp.getFolderById(
    PropertiesService.getScriptProperties().getProperty("DRIVE_FACTURAS")); }

  return _getOCreate(_getOCreate(raiz, String(anio)), mes);
}

function _carpetaReportes(anio) {
  var props = PropertiesService.getScriptProperties();
  var id    = props.getProperty("DRIVE_REPORTES");
  var raiz;
  if (id) { try { raiz = DriveApp.getFolderById(id); } catch(e) { raiz = null; } }
  if (!raiz) { crearEstructuraDrive(); raiz = DriveApp.getFolderById(
    PropertiesService.getScriptProperties().getProperty("DRIVE_REPORTES")); }

  return _getOCreate(raiz, String(anio));
}
