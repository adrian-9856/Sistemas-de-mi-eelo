// ============================================================
// 12_OrdenesDocs.gs  —  Generar Google Doc de OP u OM
//
// Replica el formulario real de la empresa: cliente, fecha,
// contacto, descripción, medidas, especificaciones,
// serigrafía, films, mockup y comentarios adicionales.
//
// Columnas de ORDENES_PRODUCCION (A=0 en índice array):
//   A: Numero_Orden     B: Tipo            C: Fecha_Creacion
//   D: Cliente          E: Contacto        F: Proyecto
//   G: Fecha_Promesa    H: Descripcion     I: Cantidad
//   J: Tela_Material    K: Comentarios_Confeccion
//   L: Ancho            M: Altura          N: Fuelle
//   O: Sistema_Medicion
//   P: Bolsillo_Interno Q: Bolsillo_Externo R: Tirantes  S: Forros
//   T: Specs_Adicionales
//   U: Serigrafia       V: Colores         W: Pantones
//   X: Localizacion     Y: Medidas_Impresion
//   Z: Films_Entregados AA: Codigos_Films
//   AB: Mockup          AC: Muestra_Bodega
//   AD: Estado          AE: Participantes   AF: Horas
//   AG: URL_Mockup_Drive AH: URL_Doc_Drive  AI: Notas
// ============================================================

// Genera el Doc de la fila seleccionada en ORDENES_PRODUCCION.
function generarOrdenProduccion() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(CONFIG.HOJAS.ORDENES);
  var fila = hoja.getActiveRange().getRow();

  if (fila < 2) {
    SpreadsheetApp.getUi().alert("Selecciona primero una fila de la hoja ORDENES_PRODUCCION.");
    return;
  }

  var o = _leerOrden(hoja, fila);
  if (!o.numero) {
    SpreadsheetApp.getUi().alert("La fila seleccionada no tiene número de orden.");
    return;
  }

  var carpeta  = _carpetaCliente(o.cliente);
  var doc      = _crearDocOrden(o, carpeta);
  var url      = doc.getUrl();

  // Guarda el enlace en columna AH (col 34)
  hoja.getRange(fila, 34).setValue(url);

  SpreadsheetApp.getUi().alert(
    "✅ Documento generado\n" + o.numero + "\n" + url
  );
}

// ---- Lectura de datos ----

function _leerOrden(hoja, fila) {
  var r = hoja.getRange(fila, 1, 1, 35).getValues()[0];
  return {
    numero:       r[0],   // A
    tipo:         r[1],   // B  (OP | OM)
    fecha:        r[2],   // C
    cliente:      r[3],   // D
    contacto:     r[4],   // E
    proyecto:     r[5],   // F
    fechaPromesa: r[6],   // G
    descripcion:  r[7],   // H
    cantidad:     r[8],   // I
    tela:         r[9],   // J
    comentConf:   r[10],  // K
    ancho:        r[11],  // L
    altura:       r[12],  // M
    fuelle:       r[13],  // N
    sisMedicion:  r[14],  // O
    bolsInterior: r[15],  // P
    bolsExterior: r[16],  // Q
    tirantes:     r[17],  // R
    forros:       r[18],  // S
    specsAadic:   r[19],  // T
    serigrafia:   r[20],  // U
    colores:      r[21],  // V
    pantones:     r[22],  // W
    locImpresion: r[23],  // X
    medImpresion: r[24],  // Y
    filmsEntregados: r[25], // Z
    codigoFilms:  r[26],  // AA
    mockup:       r[27],  // AB
    muestraBodega:r[28],  // AC
    estado:       r[29],  // AD
    participantes:r[30],  // AE
    horas:        r[31],  // AF
    urlMockup:    r[32],  // AG
    urlDoc:       r[33],  // AH
  };
}

// ---- Construcción del Google Doc ----

function _crearDocOrden(o, carpeta) {
  var titulo = o.numero + " — " + (o.descripcion || o.cliente || "Orden");
  var tz     = Session.getScriptTimeZone();

  // Si ya existe, elimina la versión anterior
  var prev = carpeta.getFilesByName(titulo);
  while (prev.hasNext()) prev.next().setTrashed(true);

  var doc  = DocumentApp.create(titulo);
  var body = doc.getBody();

  // --- ENCABEZADO ---
  _titulo(body, CONFIG.NOMBRE_ORGANIZACION);
  _subtitulo(body, "Adress / Dirección: 13 calle 2-90 Zona 7, Landívar");
  _subtitulo(body, "Ciudad de Guatemala, Guatemala");
  _subtitulo(body, "Phone / Celular: (502) 3764 9769");
  body.appendParagraph("").setSpacingAfter(4);

  // --- SECCIÓN 1: Cliente / Fecha / Número de Orden ---
  var t1 = body.appendTable([
    ["Cliente / Client", "Date / Fecha", "Número # de Orden"],
    [
      o.cliente    || "",
      o.fecha instanceof Date ? Utilities.formatDate(o.fecha, tz, "dd/MM/yyyy") : (o.fecha || ""),
      o.numero     || "",
    ],
  ]);
  _estilizarEncabezadoTabla(t1);

  body.appendParagraph("").setSpacingAfter(2);

  var t2 = body.appendTable([
    ["Contact Name / Contacto", "Project / Proyecto", "Fecha de Promesa Cliente"],
    [
      o.contacto   || "",
      o.proyecto   || "",
      o.fechaPromesa instanceof Date ? Utilities.formatDate(o.fechaPromesa, tz, "dd/MM/yyyy") : (o.fechaPromesa || ""),
    ],
  ]);
  _estilizarEncabezadoTabla(t2);

  body.appendParagraph("").setSpacingAfter(4);

  // --- SECCIÓN 2: Descripción / Cantidad / Tela ---
  var t3 = body.appendTable([
    ["Description / Descripción", "Qty", "Tela", "Comentarios"],
    [o.descripcion || "", String(o.cantidad || ""), o.tela || "", o.comentConf || ""],
  ]);
  _estilizarEncabezadoTabla(t3);

  body.appendParagraph("").setSpacingAfter(4);

  // --- SECCIÓN 3: Medidas ---
  _seccion(body, "Medidas");
  var t4 = body.appendTable([
    ["Ancho", "Altura", "Fuelle", "Sis. Medición"],
    [String(o.ancho || ""), String(o.altura || ""), String(o.fuelle || ""), o.sisMedicion || "Inches"],
  ]);
  _estilizarEncabezadoTabla(t4);

  body.appendParagraph("").setSpacingAfter(4);

  // --- SECCIÓN 4: Especificaciones Especiales ---
  _seccion(body, "Especificaciones Especiales");
  var t5 = body.appendTable([
    ["Especificación", "Sí", "No", "Medidas (Alto × Ancho)"],
    ["Bolsillo Interno",  _si(o.bolsInterior), _no(o.bolsInterior), ""],
    ["Bolsillo Externo",  _si(o.bolsExterior), _no(o.bolsExterior), ""],
    ["Tirantes",          _si(o.tirantes),     _no(o.tirantes),     ""],
    ["Forros",            _si(o.forros),        _no(o.forros),       ""],
  ]);
  _estilizarEncabezadoTabla(t5);

  if (o.specsAadic) {
    body.appendParagraph("Especificaciones adicionales: " + o.specsAadic)
        .setItalic(true);
  }
  body.appendParagraph("").setSpacingAfter(4);

  // --- SECCIÓN 5: Comentarios de Confección ---
  if (o.comentConf) {
    _seccion(body, "Comentarios Adicionales de Confección");
    body.appendParagraph(o.comentConf).setIndentFirstLine(12);
    body.appendParagraph("").setSpacingAfter(4);
  }

  // --- SECCIÓN 6: Serigrafía ---
  _seccion(body, "Serigrafía");
  var t6 = body.appendTable([
    ["¿Lleva Serigrafía?", "Cantidad de Colores", "Pantones"],
    [_siNo(o.serigrafia), String(o.colores || ""), o.pantones || ""],
  ]);
  _estilizarEncabezadoTabla(t6);

  body.appendParagraph("").setSpacingAfter(2);

  var t7 = body.appendTable([
    ["Localización de Impresión", "Medidas de Impresión", "Comentarios Adicionales"],
    [o.locImpresion || "", o.medImpresion || "", ""],
  ]);
  _estilizarEncabezadoTabla(t7);
  body.appendParagraph("").setSpacingAfter(4);

  // --- SECCIÓN 7: Films ---
  _seccion(body, "Films");
  var t8 = body.appendTable([
    ["Número de Films Entregados", "Códigos de Films"],
    [String(o.filmsEntregados || ""), o.codigoFilms || ""],
  ]);
  _estilizarEncabezadoTabla(t8);
  body.appendParagraph("").setSpacingAfter(4);

  // --- SECCIÓN 8: Mockup y Muestra ---
  _seccion(body, "Mockup / Muestra");
  var t9 = body.appendTable([
    ["Se adjunta Mockup", "Existe Muestra en Bodega"],
    [_siNo(o.mockup), _siNo(o.muestraBodega)],
  ]);
  _estilizarEncabezadoTabla(t9);
  body.appendParagraph("").setSpacingAfter(4);

  // --- SECCIÓN 9: Comentarios adicionales ---
  if (o.notas) {
    _seccion(body, "Comentarios Adicionales");
    body.appendParagraph(o.notas).setIndentFirstLine(12);
    body.appendParagraph("").setSpacingAfter(4);
  }

  // --- MOCKUP desde Drive ---
  if (o.urlMockup) {
    _seccion(body, "Mockup / Imagen de referencia");
    try {
      var fileId = _extraerFileId(o.urlMockup);
      if (fileId) {
        var blob = DriveApp.getFileById(fileId).getBlob();
        var img  = body.appendImage(blob);
        img.setWidth(300).setHeight(300);
      }
    } catch(e) {
      body.appendParagraph("Ver mockup: " + o.urlMockup).setItalic(true);
    }
  }

  // --- PIE ---
  body.appendParagraph("").setSpacingAfter(6);
  var pie = body.appendParagraph(
    "Fecha de emisión: " + Utilities.formatDate(new Date(), tz, "dd/MM/yyyy") +
    "     Estado: " + (o.estado || "Pendiente")
  );
  pie.setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setItalic(true);

  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(carpeta);
  return doc;
}

// ---- Helpers de formato ----

function _titulo(body, texto) {
  var p = body.appendParagraph(texto);
  p.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  p.editAsText().setForegroundColor("#1a237e");
}

function _subtitulo(body, texto) {
  body.appendParagraph(texto)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .setFontSize(9).setItalic(true);
}

function _seccion(body, texto) {
  var p = body.appendParagraph(texto);
  p.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  p.editAsText().setForegroundColor("#1f54a8");
}

function _estilizarEncabezadoTabla(tabla) {
  var enc = tabla.getRow(0);
  enc.editAsText().setBold(true);
  enc.editAsText().setForegroundColor("#ffffff");
  for (var c = 0; c < enc.getNumCells(); c++) {
    enc.getCell(c).setBackgroundColor("#1f54a8");
  }
}

function _siNo(val) {
  if (!val) return "";
  var s = String(val).toLowerCase().trim();
  return (s === "sí" || s === "si" || s === "true" || s === "yes" || s === "x") ? "✓ Sí" : "✗ No";
}
function _si(val) { return _siNo(val) === "✓ Sí" ? "✓" : ""; }
function _no(val) { return _siNo(val) === "✗ No" ? "✓" : ""; }

function _extraerFileId(url) {
  var match = String(url).match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}
