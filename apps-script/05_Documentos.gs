// ============================================================
// 05_Documentos.gs  —  Generación de Google Docs automáticos
// ============================================================

// Genera un Doc de factura para cada participante del mes con monto > 0.
function generarFacturasMes() {
  var ahora    = new Date();
  var mes      = ahora.getMonth() + 1;
  var anio     = ahora.getFullYear();
  var nombreMes = CONFIG.MESES[mes - 1];

  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var hojaFact = ss.getSheetByName(CONFIG.HOJAS.FACTURACION);
  var datos    = hojaFact.getDataRange().getValues();

  var carpeta  = _obtenerCarpetaFacturas(anio, nombreMes);
  var generados = 0;

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    if (fila[2] !== nombreMes || fila[3] !== anio) continue;
    if (!fila[6] || fila[6] === 0) continue; // monto 0

    var doc = _crearFacturaDoc(fila, carpeta);
    // Guarda el enlace en columna N de la hoja (si no existe, añádela)
    hojaFact.getRange(i + 1, 14).setValue(doc.getUrl());
    generados++;
  }

  SpreadsheetApp.getUi().alert(
    "✅ " + generados + " facturas generadas\nCarpeta: " + carpeta.getName()
  );
}

function _crearFacturaDoc(fila, carpeta) {
  var id     = fila[0];
  var nombre = fila[1];
  var mes    = fila[2];
  var anio   = fila[3];
  var horas  = fila[5];
  var monto  = fila[6];

  var titulo = "Factura_" + id + "_" + mes + "_" + anio;

  // Si ya existe, bórralo y recrea (para actualizar)
  var archivos = carpeta.getFilesByName(titulo);
  while (archivos.hasNext()) archivos.next().setTrashed(true);

  var doc  = DocumentApp.create(titulo);
  var body = doc.getBody();

  // Encabezado
  var encabezado = body.appendParagraph(CONFIG.NOMBRE_ORGANIZACION + " — Comprobante de Pago");
  encabezado.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  encabezado.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph("").setSpacingAfter(4);

  // Tabla con datos
  var tabla = body.appendTable([
    ["Participante",    nombre],
    ["ID",             id],
    ["Período",        mes + " " + anio],
    ["Horas trabajadas", horas + " hrs"],
    ["Tarifa por hora",  "Q " + CONFIG.TARIFA_HORA + ".00"],
    ["TOTAL A PAGAR",  "Q " + monto.toFixed(2)],
  ]);
  tabla.getRow(5).editAsText().setBold(true);

  body.appendParagraph("").setSpacingAfter(8);

  var firma = body.appendParagraph(
    "Fecha de emisión: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy")
  );
  firma.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  doc.saveAndClose();

  // Mover a la carpeta correcta en Drive
  DriveApp.getFileById(doc.getId()).moveTo(carpeta);

  return doc;
}

// Genera un reporte mensual consolidado en Google Docs.
function generarReporteMensual() {
  var ahora    = new Date();
  var mes      = ahora.getMonth() + 1;
  var anio     = ahora.getFullYear();
  var nombreMes = CONFIG.MESES[mes - 1];

  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var hojaFact = ss.getSheetByName(CONFIG.HOJAS.FACTURACION);
  var datos    = hojaFact.getDataRange().getValues();

  var carpetaReportes = _obtenerCarpetaReportes(anio);
  var titulo = "Reporte_" + nombreMes + "_" + anio;

  // Borra versión anterior si existe
  var prev = carpetaReportes.getFilesByName(titulo);
  while (prev.hasNext()) prev.next().setTrashed(true);

  var doc  = DocumentApp.create(titulo);
  var body = doc.getBody();

  // Título
  var h1 = body.appendParagraph("Reporte Mensual — " + nombreMes + " " + anio);
  h1.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  h1.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph(CONFIG.NOMBRE_ORGANIZACION)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph("").setSpacingAfter(6);

  // Tabla resumen
  var totalHoras = 0, totalMonto = 0, pendientes = 0;
  var filas = [["Participante", "Horas", "Monto (Q)", "Pagado"]];

  for (var i = 1; i < datos.length; i++) {
    var f = datos[i];
    if (f[2] !== nombreMes || f[3] !== anio) continue;

    filas.push([
      f[1],
      f[5] || 0,
      (f[6] || 0).toFixed(2),
      f[10] === "Sí" ? "✓" : "Pendiente",
    ]);
    totalHoras += parseFloat(f[5]) || 0;
    totalMonto += parseFloat(f[6]) || 0;
    if (f[10] !== "Sí") pendientes++;
  }

  var tabla = body.appendTable(filas);
  tabla.getRow(0).editAsText().setBold(true);

  body.appendParagraph("").setSpacingAfter(4);

  // Totales
  var resumen = body.appendTable([
    ["Total horas trabajadas", totalHoras.toFixed(2) + " hrs"],
    ["Total a pagar",          "Q " + totalMonto.toFixed(2)],
    ["Pagos pendientes",       pendientes + " participantes"],
  ]);
  resumen.editAsText().setBold(true);

  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(carpetaReportes);

  SpreadsheetApp.getUi().alert(
    "✅ Reporte generado\n" + titulo + "\nEnlace: " + doc.getUrl()
  );
}

// Genera el documento de una Orden de Producción/Manufactura seleccionada.
function generarOrdenProduccion() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hoja  = ss.getSheetByName(CONFIG.HOJAS.ORDENES);
  var fila  = hoja.getActiveRange().getRow();

  if (fila < 2) {
    SpreadsheetApp.getUi().alert("Selecciona primero una fila de orden.");
    return;
  }

  var datos = hoja.getRange(fila, 1, 1, 12).getValues()[0];
  var carpeta = _obtenerCarpetaOrdenes();

  var doc  = DocumentApp.create("Orden_" + datos[0]);
  var body = doc.getBody();

  var h1 = body.appendParagraph(CONFIG.NOMBRE_ORGANIZACION + " — " + datos[1]);
  h1.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  h1.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph("").setSpacingAfter(4);

  body.appendTable([
    ["Número de orden",   datos[0]],
    ["Tipo",              datos[1]],
    ["Fecha creación",    datos[2]],
    ["Cliente",           datos[3]],
    ["Contacto",          datos[4]],
    ["Descripción",       datos[5]],
    ["Cantidad",          datos[6]],
    ["Estado",            datos[7]],
    ["Participantes",     datos[8]],
    ["Fecha entrega",     datos[10]],
    ["Notas",             datos[11]],
  ]);

  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(carpeta);

  // Guarda el enlace en la hoja (columna M)
  hoja.getRange(fila, 13).setValue(doc.getUrl());

  SpreadsheetApp.getUi().alert("✅ Orden generada: " + doc.getUrl());
}
