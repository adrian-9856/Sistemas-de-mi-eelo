// ============================================================
// SISTEMA PRODUCCIÓN — Mi eelo
// Google Sheets: "Mi eelo · Producción"
// Maneja: OPs, OMs, Clientes, Drive, Docs
// ============================================================

const CFG = {
  ORG:           "Mi eelo",
  DIRECCION:     "13 calle 2-90 Zona 7, Landívar, Ciudad de Guatemala",
  TELEFONO:      "(502) 3764 9769",
  CORREO_ADMIN:  "adrian@creamosguatemala.org",

  HOJAS: {
    ORDENES:    "ORDENES",
    CLIENTES:   "CLIENTES",
    DASHBOARD:  "DASHBOARD",
  },

  // IDs de carpetas Drive (se llenan automáticamente la primera vez)
  DRIVE: {
    RAIZ:     "",   // "Mi eelo · Producción"
    ORDENES:  "",   // subcarpeta Órdenes/
    CLIENTES: "",   // subcarpeta Órdenes/Clientes/
  },
};
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("📦 Producción")
    .addItem("➕ Nueva orden (OP o OM)",             "nuevaOrden")
    .addItem("📄 Generar Doc de la orden seleccionada", "generarDocOrden")
    .addSeparator()
    .addItem("🔍 Ver todas las órdenes de un cliente",  "filtrarPorCliente")
    .addItem("🧹 Quitar filtros",                       "limpiarFiltros")
    .addSeparator()
    .addItem("📁 Crear carpeta Drive del cliente",      "crearCarpetaCliente")
    .addItem("🔄 Actualizar Dashboard",                 "actualizarDashboard")
    .addSeparator()
    .addItem("⚙️  Configurar automatizaciones",          "configurarTriggers")
    .addToUi();
}

// Cuando cambia el Estado de una orden (col AD = 30), refresca el dashboard.
function onEdit(e) {
  if (e.range.getSheet().getName() === CFG.HOJAS.ORDENES &&
      e.range.getColumn() === 30) {
    actualizarDashboard();
  }
}
// ============================================================
// 02_Ordenes.gs — CRUD y lógica de órdenes
//
// Columnas de la hoja ORDENES:
//  A  Numero_Orden      B  Tipo (OP|OM)      C  Fecha_Creacion
//  D  Cliente           E  Contacto          F  Proyecto
//  G  Fecha_Promesa     H  Descripcion       I  Cantidad
//  J  Tela_Material     K  Color_Tela
//  --- Medidas ---
//  L  Ancho             M  Altura            N  Fuelle
//  O  Sistema_Medicion
//  --- Especificaciones ---
//  P  Bolsillo_Interno  Q  Bolsillo_Externo
//  R  Tirantes          S  Forros
//  T  Specs_Adicionales
//  --- Serigrafía ---
//  U  Serigrafia        V  Num_Colores       W  Pantones
//  X  Loc_Impresion     Y  Medidas_Impresion
//  --- Films / Mockup ---
//  Z  Films_Entregados  AA Codigos_Films
//  AB Mockup            AC Muestra_Bodega
//  --- Control ---
//  AD Estado            AE Participantes_Asignados
//  AF Comentarios_Confeccion
//  AG URL_Mockup_Drive  AH URL_Doc_Drive
// ============================================================

// Abre un diálogo para crear una nueva orden guiada
function nuevaOrden() {
  var ui = SpreadsheetApp.getUi();

  var tipo = ui.prompt(
    "Nueva orden",
    "¿Es una OP (Orden de Producción) o una OM (Orden de Manufactura)?\nEscribe: OP  o  OM",
    ui.ButtonSet.OK_CANCEL
  );
  if (tipo.getSelectedButton() !== ui.Button.OK) return;
  var tipoVal = tipo.getResponseText().trim().toUpperCase();
  if (tipoVal !== "OP" && tipoVal !== "OM") {
    ui.alert("❌ Solo se permite OP o OM.");
    return;
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hoja  = ss.getSheetByName(CFG.HOJAS.ORDENES);
  var numero = _siguienteNumeroOrden(hoja, tipoVal);
  var fecha  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  // Inserta una fila vacía con el número y tipo pre-llenados
  hoja.appendRow([
    numero, tipoVal, fecha,
    "", "", "", "",   // Cliente, Contacto, Proyecto, Fecha_Promesa
    "", "", "", "",   // Descripcion, Cantidad, Tela, Color
    "", "", "", "Inches", // Medidas
    "No","No","No","No","", // Specs
    "No","","","","",  // Serigrafía
    "","",             // Films
    "No","No",         // Mockup
    "Pendiente","","", // Estado, Participantes, Comentarios
    "","",             // URLs
  ]);

  // Mueve el cursor a la nueva fila
  var ultimaFila = hoja.getLastRow();
  hoja.setActiveRange(hoja.getRange(ultimaFila, 1));

  ui.alert(
    "✅ Orden creada: " + numero + "\n" +
    "Completa los datos directamente en la fila."
  );
}

// Extrae todos los datos de la fila activa como objeto
function leerOrdenFila(hoja, fila) {
  var r = hoja.getRange(fila, 1, 1, 34).getValues()[0];
  var tz = Session.getScriptTimeZone();
  var _f = function(v) {
    return v instanceof Date
      ? Utilities.formatDate(v, tz, "dd/MM/yyyy")
      : (v || "");
  };
  return {
    numero:       r[0],   tipo:         r[1],
    fecha:        _f(r[2]), cliente:    r[3],
    contacto:     r[4],   proyecto:     r[5],
    fechaPromesa: _f(r[6]), descripcion: r[7],
    cantidad:     r[8],   tela:         r[9],
    color:        r[10],
    ancho:        r[11],  altura:       r[12],
    fuelle:       r[13],  sisMedicion:  r[14],
    bolsInt:      r[15],  bolsExt:      r[16],
    tirantes:     r[17],  forros:       r[18],
    specsExtra:   r[19],
    serigrafia:   r[20],  colores:      r[21],
    pantones:     r[22],  locImp:       r[23],
    medImp:       r[24],
    filmsNum:     r[25],  filmsCod:     r[26],
    mockup:       r[27],  muestra:      r[28],
    estado:       r[29],  participantes:r[30],
    comentarios:  r[31],
    urlMockup:    r[32],  urlDoc:       r[33],
  };
}

// Filtrar órdenes por cliente
function filtrarPorCliente() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt("Filtrar por cliente", "Escribe el nombre del cliente:", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var cliente = resp.getResponseText().trim();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hoja  = ss.getSheetByName(CFG.HOJAS.ORDENES);
  var rango = hoja.getDataRange();
  var filtro = rango.getFilter() || rango.createFilter();
  filtro.setColumnFilterCriteria(4,
    SpreadsheetApp.newFilterCriteria().whenTextContains(cliente).build()
  );
  ss.setActiveSheet(hoja);
}

function limpiarFiltros() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.HOJAS.ORDENES);
  var f = hoja.getDataRange().getFilter();
  if (f) f.remove();
}

// ---- Internas ----

function _siguienteNumeroOrden(hoja, tipo) {
  var datos = hoja.getDataRange().getValues();
  var anio  = new Date().getFullYear().toString().slice(-2);
  var prefijo = tipo + anio + "-";
  var max = 0;
  for (var i = 1; i < datos.length; i++) {
    var num = String(datos[i][0]);
    if (num.startsWith(prefijo)) {
      var n = parseInt(num.replace(prefijo, ""), 10);
      if (n > max) max = n;
    }
  }
  return prefijo + String(max + 1).padStart(3, "0");
}
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
// ============================================================
// 04_Drive.gs — Estructura de carpetas en Drive
//
// Mi eelo · Producción/
// └── Órdenes/
//     └── Clientes/
//         ├── Peace by Piece/       ← un folder por cliente
//         ├── Home Collection/
//         └── .../
// ============================================================

// Crea toda la estructura de Drive (ejecutar una sola vez).
function crearEstructuraDrive() {
  var raiz     = _getOCreate(null,      CFG.ORG + " · Producción");
  var ordenes  = _getOCreate(raiz,      "Órdenes");
  var clientes = _getOCreate(ordenes,   "Clientes");

  // Guarda los IDs en ScriptProperties para no buscar cada vez
  var props = PropertiesService.getScriptProperties();
  props.setProperty("DRIVE_RAIZ",     raiz.getId());
  props.setProperty("DRIVE_ORDENES",  ordenes.getId());
  props.setProperty("DRIVE_CLIENTES", clientes.getId());

  SpreadsheetApp.getUi().alert(
    "✅ Estructura Drive lista\n" + raiz.getUrl()
  );
}

// Crea (o abre) la carpeta del cliente dentro de Órdenes/Clientes/
function crearCarpetaCliente() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hoja  = ss.getSheetByName(CFG.HOJAS.CLIENTES);
  var fila  = hoja.getActiveRange().getRow();
  if (fila < 2) return;

  var nombre = hoja.getRange(fila, 2).getValue();
  if (!nombre) { SpreadsheetApp.getUi().alert("La fila no tiene nombre de cliente."); return; }

  var carpeta = _carpetaDeCliente(nombre);
  hoja.getRange(fila, 7).setValue(carpeta.getUrl()); // col G

  SpreadsheetApp.getUi().alert("✅ Carpeta: " + carpeta.getUrl());
}

// ---- Funciones internas usadas por otros módulos ----

function _carpetaDeCliente(nombreCliente) {
  var clientes = _carpetaClientes();
  return _getOCreate(clientes, String(nombreCliente).trim());
}

function _carpetaClientes() {
  var props = PropertiesService.getScriptProperties();
  var id    = props.getProperty("DRIVE_CLIENTES");
  if (id) {
    try { return DriveApp.getFolderById(id); } catch(e) {}
  }
  // Si no existe, crea la estructura completa
  crearEstructuraDrive();
  return DriveApp.getFolderById(
    PropertiesService.getScriptProperties().getProperty("DRIVE_CLIENTES")
  );
}

function _getOCreate(padre, nombre) {
  var iter = padre ? padre.getFoldersByName(nombre) : DriveApp.getFoldersByName(nombre);
  if (iter.hasNext()) return iter.next();
  return padre ? padre.createFolder(nombre) : DriveApp.createFolder(nombre);
}
// ============================================================
// 05_Dashboard.gs — Métricas de producción en tiempo real
// ============================================================

function actualizarDashboard() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(CFG.HOJAS.DASHBOARD)
            || ss.insertSheet(CFG.HOJAS.DASHBOARD, 0);

  var ordenes = ss.getSheetByName(CFG.HOJAS.ORDENES);
  if (!ordenes) return;

  var datos = ordenes.getDataRange().getValues();
  var m = { pendiente: 0, enProceso: 0, completada: 0, cancelada: 0, total: 0 };
  var porCliente = {};

  for (var i = 1; i < datos.length; i++) {
    var estado  = String(datos[i][29] || "").toLowerCase().trim(); // col AD
    var cliente = String(datos[i][3]  || "").trim();              // col D
    m.total++;
    if (estado === "pendiente")   m.pendiente++;
    if (estado === "en proceso")  m.enProceso++;
    if (estado === "completada")  m.completada++;
    if (estado === "cancelada")   m.cancelada++;
    if (cliente) porCliente[cliente] = (porCliente[cliente] || 0) + 1;
  }

  var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  var filas = [
    ["MÉTRICA",            "VALOR"],
    ["Total órdenes",       m.total],
    ["Pendientes",          m.pendiente],
    ["En proceso",          m.enProceso],
    ["Completadas",         m.completada],
    ["Canceladas",          m.cancelada],
    ["", ""],
    ["CLIENTE",            "ÓRDENES"],
  ];

  Object.keys(porCliente).sort().forEach(function(c) {
    filas.push([c, porCliente[c]]);
  });

  filas.push(["", ""]);
  filas.push(["Actualizado", ahora]);

  hoja.clearContents();
  hoja.getRange(1, 1, filas.length, 2).setValues(filas);

  // Formato encabezados
  [1, 8].forEach(function(r) {
    var enc = hoja.getRange(r, 1, 1, 2);
    enc.setBackground("#1f54a8").setFontColor("#ffffff").setFontWeight("bold");
  });

  // Rojo si hay pendientes
  if (m.pendiente > 0) {
    hoja.getRange(3, 2).setBackground("#fce8e6").setFontColor("#c62828");
  }

  hoja.autoResizeColumns(1, 2);
}
// ============================================================
// 06_Triggers.gs — Automatizaciones (ejecutar una sola vez)
// ============================================================

function configurarTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });

  // Dashboard se actualiza cada 30 minutos
  ScriptApp.newTrigger("actualizarDashboard")
    .timeBased().everyMinutes(30).create();

  SpreadsheetApp.getUi().alert(
    "✅ Triggers configurados:\n" +
    "• Dashboard: cada 30 minutos"
  );
}
