// ============================================================
// SISTEMA PRODUCCIÓN — Mi eelo
// Google Sheets: "Mi eelo · Producción"
// ============================================================

const CFG = {
  ORG:          "Mi eelo",
  DIRECCION:    "13 calle 2-90 Zona 7, Landívar, Ciudad de Guatemala",
  TELEFONO:     "(502) 3764 9769",
  CORREO_ADMIN: "adrian@creamosguatemala.org",
  HOJAS: {
    ORDENES:   "ORDENES",
    CLIENTES:  "CLIENTES",
    DASHBOARD: "DASHBOARD",
  },
};

// Devuelve la hoja o lanza un error amigable.
function _sh(nombre) {
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!h) throw new Error('Hoja "' + nombre + '" no encontrada.\nCréala primero en el Sheets.');
  return h;
}

// Ejecuta fn mostrando cualquier error como alerta (o en el log si no hay UI).
function _run(fn) {
  try { fn(); } catch(e) {
    try { _alert("❌ " + e.message); }
    catch(_) { Logger.log("❌ " + e.message); throw e; }
  }
}

// Muestra alerta si hay UI, o escribe en el log si se ejecuta desde el editor.
function _alert(msg) {
  try { SpreadsheetApp.getUi().alert(msg); }
  catch(_) { Logger.log(msg); }
}

// Corre ESTA función desde el editor para autorizar permisos (solo la primera vez).
function autorizar() {
  var nombre = SpreadsheetApp.getActiveSpreadsheet().getName();
  Logger.log("✅ Autorizado. Sheets: " + nombre + " — Ahora recarga el Sheets.");
}

// ── Menú ────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📦 Producción")
    .addItem("🏗️  PASO 1 — Crear hojas del sistema",    "crearHojas")
    .addItem("📁  PASO 2 — Crear estructura en Drive",  "crearEstructuraDrive")
    .addItem("⚙️  PASO 3 — Activar sincronización",      "configurarTriggers")
    .addSeparator()
    .addItem("➕ Nueva orden (OP o OM)",              "nuevaOrden")
    .addItem("🔄 Sincronizar Doc de esta fila",        "sincronizarDocFila")
    .addSeparator()
    .addItem("🔍 Filtrar órdenes por cliente",         "filtrarPorCliente")
    .addItem("🧹 Quitar filtros",                      "limpiarFiltros")
    .addSeparator()
    .addItem("📁 Crear carpeta Drive del cliente",     "crearCarpetaCliente")
    .addItem("🔄 Actualizar Dashboard",                "actualizarDashboard")
    .addSeparator()
    .addItem("🗑️  Reinstalar sistema (borra hojas)",    "reinstalarSistema")
    .addToUi();
}

// ── Crear hojas con encabezados ──────────────────────────────

function crearHojas() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── ORDENES ──────────────────────────────────────────────
  var hOrd = ss.getSheetByName("ORDENES") || ss.insertSheet("ORDENES");
  if (hOrd.getLastRow() === 0) {
    hOrd.appendRow([
      "Numero_Orden","Tipo","Fecha_Creacion","Cliente","Contacto","Proyecto",
      "Fecha_Promesa","Descripcion","Cantidad","Tela","Color",
      "Ancho","Altura","Fuelle","Sis_Medicion",
      "Bolsillo_Interno","Bolsillo_Externo","Tirantes","Forros","Specs_Adicionales",
      "Serigrafia","Num_Colores","Pantones","Loc_Impresion","Med_Impresion",
      "Films_Num","Films_Cod","Mockup","Muestra_Bodega",
      "Estado","Participantes_Asignados","Comentarios",
      "URL_Mockup_Drive","URL_Doc_Drive"
    ]);
    _formatearEncabezado(hOrd, "#1f54a8");
    // Validación de datos en columna Estado (col 30)
    var reglaEstado = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Pendiente","En Proceso","Completada","Cancelada"], true)
      .build();
    hOrd.getRange("AD2:AD1000").setDataValidation(reglaEstado);
    // Validación Tipo (col 2)
    var reglaTipo = SpreadsheetApp.newDataValidation()
      .requireValueInList(["OP","OM"], true).build();
    hOrd.getRange("B2:B1000").setDataValidation(reglaTipo);
  }

  // ── CLIENTES ─────────────────────────────────────────────
  var hCli = ss.getSheetByName("CLIENTES") || ss.insertSheet("CLIENTES");
  if (hCli.getLastRow() === 0) {
    hCli.appendRow([
      "ID_Cliente","Nombre","Contacto","Email","Telefono","Pais","URL_Carpeta_Drive","Notas"
    ]);
    _formatearEncabezado(hCli, "#1f54a8");
  }

  // Borra hoja "Hoja 1" vacía si existe
  var hDefault = ss.getSheetByName("Hoja 1") || ss.getSheetByName("Sheet1");
  if (hDefault && ss.getSheets().length > 2) ss.deleteSheet(hDefault);

  // Activa la hoja ORDENES
  ss.setActiveSheet(hOrd);

  _alert(
    "✅ Hojas creadas:\n• ORDENES (34 columnas con validaciones)\n• CLIENTES\n\n" +
    "Siguiente paso:\n📦 Producción → PASO 2 — Crear estructura en Drive"
  );
}); }

function _formatearEncabezado(hoja, color) {
  var enc = hoja.getRange(1, 1, 1, hoja.getLastColumn());
  enc.setBackground(color).setFontColor("#ffffff").setFontWeight("bold");
  hoja.setFrozenRows(1);
  hoja.getRange(1,1,1,hoja.getLastColumn()).setWrap(false);
}

// Simple trigger (no necesita autorización).
function onEdit(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== CFG.HOJAS.ORDENES) return;
  if (e.range.getRow() < 2) return;
  try { if (e.range.getColumn() === 30) actualizarDashboard(); } catch(_) {}
}

// Trigger instalable — actualiza el Doc al editar (requiere PASO 3).
function onEditInstalable(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== CFG.HOJAS.ORDENES) return;
  var fila = e.range.getRow();
  if (fila < 2) return;
  if (e.range.getColumn() === 34) return; // col AH = URL del Doc: ignorar para evitar loop
  try { _sincronizarFila(sheet, fila); } catch(err) { Logger.log("Sync error fila " + fila + ": " + err.message); }
}

// Sincroniza el Doc de la fila activa desde el menú (siempre funciona).
function sincronizarDocFila() { _run(function() {
  var hoja = _sh(CFG.HOJAS.ORDENES);
  var fila = hoja.getActiveRange().getRow();
  if (fila < 2) { _alert("Selecciona una fila con una orden primero."); return; }
  var url = _sincronizarFila(hoja, fila);
  _alert("✅ Doc actualizado.\n" + url);
}); }

// Núcleo: crea o actualiza el Doc de la fila. Nunca crea uno nuevo si ya existe.
function _sincronizarFila(hoja, fila) {
  var o        = _leerFila(hoja, fila);
  if (!o.numero) throw new Error("La fila " + fila + " no tiene número de orden.");

  // Lee la URL real de col 34 — funciona con valor plano O fórmula HYPERLINK en cualquier idioma.
  var celdaUrl = hoja.getRange(fila, 34);
  var urlActual = String(celdaUrl.getValue() || "");
  if (!urlActual.startsWith("http")) {
    var fm  = celdaUrl.getFormula();
    var mUrl = fm.match(/https?:\/\/[^\s"]+/); // extrae URL sin depender del nombre de la función
    if (mUrl) urlActual = mUrl[0];
  }
  o.urlDoc = urlActual; // garantiza que _construirDoc abra el Doc existente

  var carpeta  = _carpetaCliente(o.cliente || "Sin_Cliente");
  var doc      = _construirDoc(o, carpeta);
  var urlNueva = doc.getUrl();

  // Solo escribe en la celda si la URL cambió (evita disparar el trigger innecesariamente).
  if (urlNueva !== urlActual) celdaUrl.setValue(urlNueva);
  return urlNueva;
}

// ── Órdenes ─────────────────────────────────────────────────
// Columnas ORDENES:
//  A Numero  B Tipo  C Fecha  D Cliente  E Contacto  F Proyecto
//  G Fecha_Promesa  H Descripcion  I Cantidad  J Tela  K Color
//  L Ancho  M Altura  N Fuelle  O Sis_Medicion
//  P Bolsillo_Int  Q Bolsillo_Ext  R Tirantes  S Forros  T Specs_Extra
//  U Serigrafia  V Colores  W Pantones  X Loc_Impresion  Y Med_Impresion
//  Z Films_Num  AA Films_Cod  AB Mockup  AC Muestra_Bodega
//  AD Estado  AE Participantes  AF Comentarios  AG URL_Mockup  AH URL_Doc

function nuevaOrden() { _run(function() {
  var ui     = SpreadsheetApp.getUi();
  var resp = ui.prompt("Nueva orden", "¿OP o OM?", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var tipo = resp.getResponseText().trim().toUpperCase();
  if (tipo !== "OP" && tipo !== "OM") { ui.alert("Escribe OP o OM."); return; }

  var hoja   = _sh(CFG.HOJAS.ORDENES);
  var numero = _siguienteNumero(hoja, tipo);
  var fecha  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  hoja.appendRow([
    numero, tipo, fecha,
    "","","","",
    "","","","",
    "","","","Inches",
    "No","No","No","No","",
    "No","","","","",
    "","",
    "No","No",
    "Pendiente","","",
    "","",
  ]);

  var fila = hoja.getLastRow();
  hoja.setActiveRange(hoja.getRange(fila, 1));
  var url = _sincronizarFila(hoja, fila);
  ui.alert("✅ " + numero + " creada.\nDoc listo — completa los datos y usa 🔄 Sincronizar para actualizar.\n\n" + url);
}); }

function filtrarPorCliente() { _run(function() {
  var ui   = SpreadsheetApp.getUi();
  var resp = ui.prompt("Filtrar", "Nombre del cliente:", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var hoja  = _sh(CFG.HOJAS.ORDENES);
  var rango = hoja.getDataRange();
  var f = rango.getFilter() || rango.createFilter();
  f.setColumnFilterCriteria(4,
    SpreadsheetApp.newFilterCriteria().whenTextContains(resp.getResponseText().trim()).build()
  );
}); }

function limpiarFiltros() { _run(function() {
  var hoja = _sh(CFG.HOJAS.ORDENES);
  var f    = hoja.getDataRange().getFilter();
  if (f) f.remove();
}); }

function _siguienteNumero(hoja, tipo) {
  var datos   = hoja.getDataRange().getValues();
  var anio    = new Date().getFullYear().toString().slice(-2);
  var prefijo = tipo + anio + "-";
  var max     = 0;
  for (var i = 1; i < datos.length; i++) {
    var s = String(datos[i][0]);
    if (s.startsWith(prefijo)) {
      var n = parseInt(s.replace(prefijo, ""), 10);
      if (n > max) max = n;
    }
  }
  return prefijo + String(max + 1).padStart(3, "0");
}

function _leerFila(hoja, fila) {
  var r  = hoja.getRange(fila, 1, 1, 34).getValues()[0];
  var tz = Session.getScriptTimeZone();
  var f  = function(v) { return v instanceof Date ? Utilities.formatDate(v,tz,"dd/MM/yyyy") : (v||""); };
  return {
    numero:r[0], tipo:r[1], fecha:f(r[2]), cliente:r[3], contacto:r[4],
    proyecto:r[5], fechaPromesa:f(r[6]), descripcion:r[7], cantidad:r[8],
    tela:r[9], color:r[10], ancho:r[11], altura:r[12], fuelle:r[13],
    sisMed:r[14], bolsInt:r[15], bolsExt:r[16], tirantes:r[17], forros:r[18],
    specs:r[19], serigrafia:r[20], colores:r[21], pantones:r[22],
    locImp:r[23], medImp:r[24], filmsNum:r[25], filmsCod:r[26],
    mockup:r[27], muestra:r[28], estado:r[29], participantes:r[30],
    comentarios:r[31], urlMockup:r[32], urlDoc:r[33],
  };
}

// ── Documentos ──────────────────────────────────────────────

function _construirDoc(o, carpeta) {
  var titulo = o.numero + (o.descripcion ? " — " + o.descripcion : "");
  var doc = null;

  // Abre el Doc existente por URL — nunca borra nada.
  var urlExistente = String(o.urlDoc || "");
  if (urlExistente.startsWith("http")) {
    var m = urlExistente.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (m) { try { doc = DocumentApp.openById(m[1]); doc.setName(titulo); } catch(_) {} }
  }

  // Solo crea uno nuevo si no existe ninguno todavía.
  if (!doc) {
    doc = DocumentApp.create(titulo);
    DriveApp.getFileById(doc.getId()).moveTo(carpeta);
  }

  var body = doc.getBody();
  body.clear();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);
  _escribirContenidoDoc(body, o);
  doc.saveAndClose();
  return doc;
}

function _escribirContenidoDoc(body, o) {
  _h1(body, CFG.ORG);
  _meta(body, CFG.DIRECCION);
  _meta(body, CFG.TELEFONO);
  _sp(body);

  _tabla(body, ["Cliente","Fecha","Número de Orden"],        [o.cliente, o.fecha, o.numero]);
  _sp(body);
  _tabla(body, ["Contacto","Proyecto","Fecha de Promesa"],   [o.contacto, o.proyecto, o.fechaPromesa]);
  _sp(body);
  _tabla(body, ["Descripción","Qty","Tela","Color","Comentarios"],
               [o.descripcion, String(o.cantidad||""), o.tela, o.color, o.comentarios]);
  _sp(body);

  _sec(body, "Medidas");
  _tabla(body, ["Ancho","Altura","Fuelle","Sistema"],
               [String(o.ancho||""), String(o.altura||""), String(o.fuelle||""), o.sisMed||"Inches"]);
  _sp(body);

  _sec(body, "Especificaciones Especiales");
  _tablaSpecs(body, [
    ["Bolsillo Interno", o.bolsInt], ["Bolsillo Externo", o.bolsExt],
    ["Tirantes", o.tirantes],        ["Forros", o.forros],
  ]);
  if (o.specs) body.appendParagraph("Adicionales: " + o.specs).setItalic(true).setFontSize(9);
  _sp(body);

  _sec(body, "Serigrafía");
  _tabla(body, ["¿Lleva Serigrafía?","N.° Colores","Pantones"],
               [_siNo(o.serigrafia), String(o.colores||""), o.pantones||"—"]);
  _sp(body);
  _tabla(body, ["Localización de Impresión","Medidas de Impresión"],
               [o.locImp||"—", o.medImp||"—"]);
  _sp(body);

  _sec(body, "Films");
  _tabla(body, ["N.° Films Entregados","Códigos"], [String(o.filmsNum||""), o.filmsCod||"—"]);
  _sp(body);

  _sec(body, "Mockup / Muestra");
  _tabla(body, ["Se adjunta Mockup","Muestra en Bodega"], [_siNo(o.mockup), _siNo(o.muestra)]);

  if (o.urlMockup) {
    _sp(body);
    try {
      var fid = (o.urlMockup.match(/\/d\/([a-zA-Z0-9_-]+)/)||[])[1];
      if (fid) { var img = body.appendImage(DriveApp.getFileById(fid).getBlob()); img.setWidth(280).setHeight(280); }
    } catch(_) { body.appendParagraph("Mockup: " + o.urlMockup).setItalic(true).setFontSize(9); }
  }

  _sp(body);
  body.appendParagraph(
    "Actualizado: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") +
    "     Estado: " + (o.estado||"Pendiente")
  ).setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setItalic(true).setFontSize(9);
}

function _h1(body, t) {
  body.appendParagraph(t).setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .editAsText().setForegroundColor("#1a237e").setFontSize(18);
}
function _meta(body, t) {
  body.appendParagraph(t).setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .setItalic(true).setFontSize(9);
}
function _sec(body, t) {
  body.appendParagraph(t).setHeading(DocumentApp.ParagraphHeading.HEADING3)
      .editAsText().setForegroundColor("#1f54a8").setFontSize(11);
}
function _sp(body) { body.appendParagraph("").setSpacingAfter(2); }

function _tabla(body, enc, val) {
  var t = body.appendTable([enc, val]);
  var h = t.getRow(0);
  h.editAsText().setBold(true).setForegroundColor("#ffffff");
  for (var c = 0; c < h.getNumCells(); c++) h.getCell(c).setBackgroundColor("#1f54a8");
  t.setBorderColor("#cccccc");
}
function _tablaSpecs(body, filas) {
  var rows = [["Especificación","Sí","No"]];
  filas.forEach(function(f) { var s = _esSi(f[1]); rows.push([f[0], s?"✓":"", s?"":"✓"]); });
  var t = body.appendTable(rows);
  t.getRow(0).editAsText().setBold(true).setForegroundColor("#ffffff");
  for (var c = 0; c < 3; c++) t.getRow(0).getCell(c).setBackgroundColor("#1f54a8");
}
function _siNo(v) { return _esSi(v) ? "✓ Sí" : "✗ No"; }
function _esSi(v) {
  var s = String(v||"").toLowerCase().trim();
  return s==="sí"||s==="si"||s==="true"||s==="yes"||s==="x"||s==="✓";
}

// ── Drive ────────────────────────────────────────────────────

function crearEstructuraDrive() { _run(function() {
  var raiz = _getOCreate(null, CFG.ORG + " · Producción");
  var ord  = _getOCreate(raiz, "Órdenes");
  var cli  = _getOCreate(ord,  "Clientes");
  var p    = PropertiesService.getScriptProperties();
  p.setProperty("DRIVE_RAIZ",     raiz.getId());
  p.setProperty("DRIVE_CLIENTES", cli.getId());
  _alert("✅ Estructura Drive creada\n" + raiz.getUrl());
}); }

function crearCarpetaCliente() { _run(function() {
  var hoja   = _sh(CFG.HOJAS.CLIENTES);
  var fila   = hoja.getActiveRange().getRow();
  if (fila < 2) return;
  var nombre = hoja.getRange(fila, 2).getValue();
  if (!nombre) { _alert("La fila no tiene nombre de cliente."); return; }
  var carpeta = _carpetaCliente(nombre);
  hoja.getRange(fila, 7).setValue(carpeta.getUrl());
  _alert("✅ Carpeta: " + carpeta.getUrl());
}); }

function _carpetaCliente(nombre) {
  var p  = PropertiesService.getScriptProperties();
  var id = p.getProperty("DRIVE_CLIENTES");
  var cli;
  if (id) { try { cli = DriveApp.getFolderById(id); } catch(_) {} }
  if (!cli) {
    crearEstructuraDrive();
    cli = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty("DRIVE_CLIENTES"));
  }
  return _getOCreate(cli, String(nombre).trim());
}

function _getOCreate(padre, nombre) {
  var it = padre ? padre.getFoldersByName(nombre) : DriveApp.getFoldersByName(nombre);
  if (it.hasNext()) return it.next();
  return padre ? padre.createFolder(nombre) : DriveApp.createFolder(nombre);
}

// ── Dashboard ────────────────────────────────────────────────

function actualizarDashboard() { _run(function() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var dash    = ss.getSheetByName(CFG.HOJAS.DASHBOARD) || ss.insertSheet(CFG.HOJAS.DASHBOARD, 0);
  var ordenes = ss.getSheetByName(CFG.HOJAS.ORDENES);
  if (!ordenes) { dash.getRange(1,1).setValue("Crea la hoja ORDENES primero."); return; }

  var datos = ordenes.getDataRange().getValues();
  var m = { pendiente:0, enProceso:0, completada:0, cancelada:0, total:0 };
  var xCliente = {};
  for (var i = 1; i < datos.length; i++) {
    var est = String(datos[i][29]||"").toLowerCase().trim();
    var cli = String(datos[i][3]||"").trim();
    m.total++;
    if (est==="pendiente")  m.pendiente++;
    if (est==="en proceso") m.enProceso++;
    if (est==="completada") m.completada++;
    if (est==="cancelada")  m.cancelada++;
    if (cli) xCliente[cli] = (xCliente[cli]||0) + 1;
  }
  var ts    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  var filas = [
    ["MÉTRICA","VALOR"],
    ["Total órdenes", m.total],
    ["Pendientes",    m.pendiente],
    ["En proceso",    m.enProceso],
    ["Completadas",   m.completada],
    ["Canceladas",    m.cancelada],
    ["",""],
    ["CLIENTE","ÓRDENES"],
  ];
  Object.keys(xCliente).sort().forEach(function(c) { filas.push([c, xCliente[c]]); });
  filas.push(["",""]); filas.push(["Actualizado", ts]);

  dash.clearContents();
  dash.getRange(1,1,filas.length,2).setValues(filas);
  [1,8].forEach(function(r) {
    dash.getRange(r,1,1,2).setBackground("#1f54a8").setFontColor("#fff").setFontWeight("bold");
  });
  if (m.pendiente > 0) dash.getRange(3,2).setBackground("#fce8e6").setFontColor("#c62828");
  dash.autoResizeColumns(1,2);
}); }

// ── Triggers ─────────────────────────────────────────────────

function configurarTriggers() { _run(function() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger("onEditInstalable").forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger("actualizarDashboard").timeBased().everyMinutes(30).create();
  _alert("✅ Automatizaciones activas:\n• Doc se actualiza al editar cada fila\n• Dashboard cada 30 min");
}); }

function reinstalarSistema() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    "⚠️ Reinstalar sistema",
    "Esto BORRA las hojas ORDENES, CLIENTES y DASHBOARD y las recrea vacías.\n\nLos Docs y carpetas en Drive NO se tocan.\n\n¿Continuar?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ["ORDENES","CLIENTES","DASHBOARD"].forEach(function(nombre) {
    var h = ss.getSheetByName(nombre);
    if (h) ss.deleteSheet(h);
  });

  // Espera un momento para que Sheets procese los deletes.
  Utilities.sleep(500);

  // Recrea las hojas desde cero.
  crearHojas();
}); }
