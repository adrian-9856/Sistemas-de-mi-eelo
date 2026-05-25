// ============================================================
// 03_Documentos.gs — Genera el Google Doc de una OP o OM
// Replica el formulario real: cliente, medidas, especificaciones,
// serigrafía, films, mockup. Guarda en carpeta del cliente en Drive.
// ============================================================

function generarDocOrden() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(CFG.HOJAS.ORDENES);
  var fila = hoja.getActiveRange().getRow();

  if (fila < 2) {
    SpreadsheetApp.getUi().alert("Selecciona primero una fila de ORDENES.");
    return;
  }

  var o = leerOrdenFila(hoja, fila);
  if (!o.numero) {
    SpreadsheetApp.getUi().alert("La fila seleccionada no tiene número de orden.");
    return;
  }

  var carpeta = _carpetaDeCliente(o.cliente || "Sin_Cliente");
  var doc     = _construirDoc(o, carpeta);

  // Guarda el enlace en col AH (34)
  hoja.getRange(fila, 34).setValue(doc.getUrl());
  hoja.getRange(fila, 34).setFormula('=HYPERLINK("' + doc.getUrl() + '","Ver Doc")');

  SpreadsheetApp.getUi().alert("✅ " + o.numero + " generado.\n" + doc.getUrl());
}

// ---- Construcción del Doc ----

function _construirDoc(o, carpeta) {
  var titulo = o.numero + (o.descripcion ? " — " + o.descripcion : "");

  // Elimina versión anterior si existe
  var prev = carpeta.getFilesByName(titulo);
  while (prev.hasNext()) prev.next().setTrashed(true);

  var doc  = DocumentApp.create(titulo);
  var body = doc.getBody();
  body.setMarginTop(36).setMarginBottom(36)
      .setMarginLeft(54).setMarginRight(54);

  // ── Encabezado ──────────────────────────────────────────
  _h1(body, CFG.ORG);
  _meta(body, CFG.DIRECCION);
  _meta(body, CFG.TELEFONO);
  _espacio(body);

  // ── Bloque 1: Cliente / Fecha / Número ──────────────────
  _tablaConEncabezado(body,
    ["Cliente / Client",  "Date / Fecha",  "Número # de Orden"],
    [o.cliente,           o.fecha,         o.numero]
  );
  _espacio(body);
  _tablaConEncabezado(body,
    ["Contact Name / Contacto",  "Project / Proyecto",  "Fecha de Promesa Cliente"],
    [o.contacto,                  o.proyecto,            o.fechaPromesa]
  );
  _espacio(body);

  // ── Bloque 2: Descripción ────────────────────────────────
  _tablaConEncabezado(body,
    ["Description / Descripción",  "Qty",            "Tela",    "Color",      "Comentarios de Confección"],
    [o.descripcion,                 String(o.cantidad || ""),  o.tela,    o.color,      o.comentarios]
  );
  _espacio(body);

  // ── Bloque 3: Medidas ────────────────────────────────────
  _seccion(body, "Medidas");
  _tablaConEncabezado(body,
    ["Ancho",          "Altura",         "Fuelle",         "Sis. Medición"],
    [String(o.ancho || ""), String(o.altura || ""), String(o.fuelle || ""), o.sisMedicion || "Inches"]
  );
  _espacio(body);

  // ── Bloque 4: Especificaciones ───────────────────────────
  _seccion(body, "Especificaciones Especiales");
  _tablaEspecificaciones(body, [
    ["Bolsillo Interno",  o.bolsInt],
    ["Bolsillo Externo",  o.bolsExt],
    ["Tirantes",          o.tirantes],
    ["Forros",            o.forros],
  ]);
  if (o.specsExtra) {
    body.appendParagraph("Specs adicionales: " + o.specsExtra).setItalic(true).setFontSize(9);
  }
  _espacio(body);

  // ── Bloque 5: Serigrafía ─────────────────────────────────
  _seccion(body, "Serigrafía");
  _tablaConEncabezado(body,
    ["¿Lleva Serigrafía?",  "N.° de Colores",   "Pantones"],
    [_siNo(o.serigrafia),   String(o.colores || ""),  o.pantones || "—"]
  );
  _espacio(body);
  _tablaConEncabezado(body,
    ["Localización de Impresión",  "Medidas de Impresión"],
    [o.locImp || "—",              o.medImp  || "—"]
  );
  _espacio(body);

  // ── Bloque 6: Films ──────────────────────────────────────
  _seccion(body, "Films");
  _tablaConEncabezado(body,
    ["N.° de Films Entregados",  "Códigos de Films"],
    [String(o.filmsNum || ""),   o.filmsCod || "—"]
  );
  _espacio(body);

  // ── Bloque 7: Mockup ─────────────────────────────────────
  _seccion(body, "Mockup / Muestra en Bodega");
  _tablaConEncabezado(body,
    ["Se adjunta Mockup",  "Existe Muestra en Bodega"],
    [_siNo(o.mockup),      _siNo(o.muestra)]
  );

  // Inserta imagen del mockup si hay URL de Drive
  if (o.urlMockup) {
    _espacio(body);
    try {
      var fileId = (o.urlMockup.match(/\/d\/([a-zA-Z0-9_-]+)/) || [])[1];
      if (fileId) {
        var img = body.appendImage(DriveApp.getFileById(fileId).getBlob());
        img.setWidth(280).setHeight(280);
      }
    } catch(e) {
      body.appendParagraph("Mockup: " + o.urlMockup).setItalic(true).setFontSize(9);
    }
  }

  // ── Pie ──────────────────────────────────────────────────
  _espacio(body);
  var pie = body.appendParagraph(
    "Emisión: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy") +
    "     Estado: " + (o.estado || "Pendiente")
  );
  pie.setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
     .setItalic(true).setFontSize(9);

  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(carpeta);
  return doc;
}

// ── Helpers de formato ───────────────────────────────────

function _h1(body, texto) {
  var p = body.appendParagraph(texto);
  p.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  p.editAsText().setForegroundColor("#1a237e").setFontSize(18);
}

function _meta(body, texto) {
  body.appendParagraph(texto)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setItalic(true).setFontSize(9);
}

function _seccion(body, texto) {
  var p = body.appendParagraph(texto);
  p.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  p.editAsText().setForegroundColor("#1f54a8").setFontSize(11);
}

function _espacio(body) {
  body.appendParagraph("").setSpacingAfter(2);
}

function _tablaConEncabezado(body, encabezados, valores) {
  var tabla = body.appendTable([encabezados, valores]);
  var enc   = tabla.getRow(0);
  enc.editAsText().setBold(true).setForegroundColor("#ffffff");
  for (var c = 0; c < enc.getNumCells(); c++) {
    enc.getCell(c).setBackgroundColor("#1f54a8");
  }
  tabla.setBorderColor("#cccccc");
}

function _tablaEspecificaciones(body, filas) {
  var datos = [["Especificación", "Sí", "No"]];
  filas.forEach(function(f) {
    var es = _esSi(f[1]);
    datos.push([f[0], es ? "✓" : "", es ? "" : "✓"]);
  });
  var tabla = body.appendTable(datos);
  tabla.getRow(0).editAsText().setBold(true).setForegroundColor("#ffffff");
  for (var c = 0; c < 3; c++) {
    tabla.getRow(0).getCell(c).setBackgroundColor("#1f54a8");
  }
}

function _siNo(v) { return _esSi(v) ? "✓ Sí" : "✗ No"; }
function _esSi(v) {
  var s = String(v || "").toLowerCase().trim();
  return s === "sí" || s === "si" || s === "true" || s === "yes" || s === "x" || s === "✓";
}
