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
};

// Devuelve la hoja o lanza error amigable.
function _sh(nombre) {
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!h) throw new Error('Hoja "' + nombre + '" no encontrada.\nCréala primero en el Sheets.');
  return h;
}

// Ejecuta fn mostrando cualquier error como alerta (o en el log si no hay UI).
function _run(fn) {
  try { fn(); } catch(e) {
    try { SpreadsheetApp.getUi().alert("❌ " + e.message); }
    catch(_) { Logger.log("❌ " + e.message); throw e; }
  }
}

// ── Menú ────────────────────────────────────────────────────

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("👥 RRHH")
    .addItem("🏗️  PASO 1 — Crear hojas del sistema",   "crearHojas")
    .addItem("📁  PASO 2 — Crear estructura en Drive", "crearEstructuraDrive")
    .addSeparator()
    .addSubMenu(ui.createMenu("📂 Carga inicial (una sola vez)")
      .addItem("👥 Importar participantes",           "importarParticipantes")
      .addItem("📋 Importar asistencia histórica",    "importarAsistenciaHistorica")
      .addItem("💰 Importar facturación histórica",   "importarFacturacionHistorica")
    )
    .addSeparator()
    .addItem("📥 Importar desde Kobo",                "importarDesdeKobo")
    .addItem("🔗 Emparejar entradas/salidas → horas", "emparejarAsistencia")
    .addSeparator()
    .addItem("💰 Calcular facturación del mes",       "calcularFacturacionMes")
    .addItem("📄 Generar recibos de pago (Docs)",     "generarRecibosMes")
    .addItem("📊 Generar reporte mensual",            "generarReporteMensual")
    .addSeparator()
    .addItem("🔔 Recordatorio de pagos pendientes",   "enviarRecordatorioPagos")
    .addItem("📬 Resumen mensual al admin",           "enviarResumenMensual")
    .addSeparator()
    .addItem("🔄 Actualizar Dashboard",               "actualizarDashboard")
    .addItem("📁 Crear estructura en Drive",          "crearEstructuraDrive")
    .addItem("⚙️  Configurar automatizaciones",        "configurarTriggers")
    .addToUi();
}

function onEdit(e) {
  if (e.range.getSheet().getName() === CFG.HOJAS.FACTURACION && e.range.getColumn() === 11) {
    try { actualizarDashboard(); } catch(_) {}
  }
}

// ── Crear hojas ─────────────────────────────────────────────

function crearHojas() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var hP = ss.getSheetByName("PARTICIPANTES") || ss.insertSheet("PARTICIPANTES");
  if (hP.getLastRow() === 0) {
    hP.appendRow(["Creamos_ID","Nombre","Proyecto","Division","Programa","Estado","Etapa",
      "Educacion","Apoyo_Emocional","Inclusion_Laboral","Categoria",
      "DPI","NIT","Correo_Electronico","Banco","Numero_Cuenta","Forma_Pago"]);
    _fmtEnc(hP, "#639922");
    hP.getRange("F2:F500").setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(["Activo","Inactivo","Egresado"], true).build());
  }

  var hA = ss.getSheetByName("ASISTENCIA") || ss.insertSheet("ASISTENCIA");
  if (hA.getLastRow() === 0) {
    hA.appendRow(["Creamos_ID","Nombre","Fecha_Registro","Tipo",
      "Horas_Trabajadas","Es_Dia_Estudio","Es_Terapia","Porcentaje_Pago","Horas_A_Pagar","UUID_Kobo"]);
    _fmtEnc(hA, "#1f54a8");
  }

  var hF = ss.getSheetByName("FACTURACION") || ss.insertSheet("FACTURACION");
  if (hF.getLastRow() === 0) {
    hF.appendRow(["Creamos_ID","Nombre","Mes","Anio","Quincena",
      "Horas_Trabajadas","Monto_A_Pagar","Factura_Entregada","Numero_Factura",
      "Declaraguate","Pagado","Fecha_Pago","Comentarios"]);
    _fmtEnc(hF, "#639922");
    hF.getRange("H2:H2000").setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(["Sí","No"], true).build());
    hF.getRange("J2:J2000").setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(["Sí","No"], true).build());
    hF.getRange("K2:K2000").setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(["Sí","No"], true).build());
  }

  var hDef = ss.getSheetByName("Hoja 1") || ss.getSheetByName("Sheet1");
  if (hDef && ss.getSheets().length > 3) ss.deleteSheet(hDef);

  ss.setActiveSheet(hP);
  SpreadsheetApp.getUi().alert("✅ Hojas creadas:\n• PARTICIPANTES\n• ASISTENCIA\n• FACTURACION\n\nSiguiente paso: PASO 2 — Crear estructura en Drive");
}); }

function _fmtEnc(hoja, color) {
  var enc = hoja.getRange(1, 1, 1, hoja.getLastColumn());
  enc.setBackground(color).setFontColor("#ffffff").setFontWeight("bold");
  hoja.setFrozenRows(1);
}

// ── Kobo ────────────────────────────────────────────────────
// Columnas ASISTENCIA:
//  A CreAmosID  B Nombre  C Fecha_Registro  D Tipo
//  E Horas  F Es_Dia_Estudio  G Es_Terapia  H Porcentaje  I Horas_A_Pagar  J UUID

function importarDesdeKobo() { _run(function() {
  var hoja = _sh(CFG.HOJAS.ASISTENCIA);
  var url  = "https://kc.kobotoolbox.org/api/v2/assets/" +
             CFG.KOBO_ASSET_UID + "/data/?format=json&limit=5000";
  var resp = UrlFetchApp.fetch(url, {
    method:"GET", headers:{"Authorization":"Token " + CFG.KOBO_TOKEN}, muteHttpExceptions:true
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error("Error Kobo HTTP " + resp.getResponseCode());
  }
  var results    = JSON.parse(resp.getContentText()).results || [];
  var existentes = _uuidsExistentes(hoja);
  var nuevos     = [];
  results.forEach(function(r) {
    var uuid = r["_uuid"] || "";
    if (existentes[uuid]) return;
    var p  = r["Participante"] || "";
    var t  = r["Ingreso / Egreso"] || "";
    var ts = new Date(r["start"] || r["_submission_time"]);
    if (!p || isNaN(ts)) return;
    nuevos.push([_extraerID(p), _extraerNombre(p), ts, t, "","","","","", uuid]);
  });
  if (nuevos.length > 0) {
    hoja.getRange(hoja.getLastRow() + 1, 1, nuevos.length, 10).setValues(nuevos);
  }
  emparejarAsistencia();
  SpreadsheetApp.getUi().alert("✅ " + nuevos.length + " registros nuevos importados.");
}); }

// Carga histórica desde hoja "DatosKobo" pegada en este Sheets.
function importarAsistenciaHistorica() { _run(function() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrig = ss.getSheetByName("DatosKobo");
  if (!hojaOrig) throw new Error(
    'Crea una hoja llamada "DatosKobo" y pega ahí el contenido\n' +
    'de la hoja DatosKobo del archivo Planilla___mi_eelo.xlsx'
  );
  var hojaDest   = _sh(CFG.HOJAS.ASISTENCIA);
  var datos      = hojaOrig.getDataRange().getValues();
  var existentes = _uuidsExistentes(hojaDest);
  var nuevos     = [];
  for (var i = 1; i < datos.length; i++) {
    var p    = String(datos[i][4]||"");   // col E: Participante
    var tipo = String(datos[i][3]||"");   // col D: Ingreso/Egreso
    var ts   = datos[i][0];               // col A: start
    var uuid = String(datos[i][11]||"");  // col L: _uuid
    if (!p || !uuid || existentes[uuid]) continue;
    nuevos.push([_extraerID(p), _extraerNombre(p), ts, tipo, "","","","","", uuid]);
  }
  if (nuevos.length > 0) {
    hojaDest.getRange(hojaDest.getLastRow() + 1, 1, nuevos.length, 10).setValues(nuevos);
  }
  emparejarAsistencia();
  SpreadsheetApp.getUi().alert("✅ " + nuevos.length + " registros históricos importados.");
}); }

function emparejarAsistencia() { _run(function() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = _sh(CFG.HOJAS.ASISTENCIA);
  var datos = hoja.getDataRange().getValues();
  var diasEstudio = _setDiasEstudio(ss);
  var terapias    = _setTerapias(ss);
  var grupos = {};
  for (var i = 1; i < datos.length; i++) {
    var id   = String(datos[i][0]).trim();
    var tipo = String(datos[i][3]).trim();
    var ts   = new Date(datos[i][2]);
    if (!id || isNaN(ts)) continue;
    var clave = id + "|" + _dClave(ts);
    if (!grupos[clave]) grupos[clave] = { ent:[], sal:[], filas:[] };
    if (tipo === CFG.KOBO_TIPO_ENTRADA) grupos[clave].ent.push({ts:ts});
    if (tipo === CFG.KOBO_TIPO_SALIDA)  grupos[clave].sal.push({ts:ts});
    grupos[clave].filas.push(i + 1);
  }
  Object.keys(grupos).forEach(function(clave) {
    var g = grupos[clave];
    if (!g.ent.length || !g.sal.length) return;
    g.ent.sort(function(a,b){return a.ts-b.ts;});
    g.sal.sort(function(a,b){return a.ts-b.ts;});
    var id   = clave.split("|")[0];
    var hrs  = Math.max(0, Math.round((g.sal[g.sal.length-1].ts - g.ent[0].ts)/36000)/100);
    var esE  = diasEstudio.has(id) ? "Sí" : "No";
    var esT  = terapias.has(id)    ? "Sí" : "No";
    var pct  = esE === "Sí" ? 0 : 100;
    var hap  = hrs * (pct/100);
    g.filas.forEach(function(r) {
      hoja.getRange(r, 5, 1, 5).setValues([[hrs, esE, esT, pct, hap]]);
    });
  });
}); }

function _extraerID(txt) {
  var m = String(txt).match(/\(([A-Z]{2,4}\d{6,})\)/);
  return m ? m[1] : "";
}
function _extraerNombre(txt) {
  return String(txt).replace(/\s*\([A-Z]{2,4}\d{6,}\)\s*$/, "").trim();
}
function _dClave(d) {
  return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");
}
function _uuidsExistentes(hoja) {
  var m = {}, d = hoja.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) { if (d[i][9]) m[d[i][9]] = true; }
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
  for (var i = 1; i < d.length; i++) { if (d[i][1]) s.add(String(d[i][0])); }
  return s;
}

// ── Facturación ─────────────────────────────────────────────
// Columnas FACTURACION:
//  A CreAmosID  B Nombre  C Mes  D Anio  E Quincena
//  F Horas  G Monto  H Factura_Entregada  I Num_Factura
//  J Declaraguate  K Pagado  L Fecha_Pago  M Comentarios

function calcularFacturacionMes() { _run(function() {
  var ahora = new Date();
  _calcular(ahora.getMonth()+1, ahora.getFullYear());
}); }

function _calcular(mes, anio) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var hojaA     = _sh(CFG.HOJAS.ASISTENCIA);
  var hojaP     = _sh(CFG.HOJAS.PARTICIPANTES);
  var hojaF     = _sh(CFG.HOJAS.FACTURACION);
  var nombreMes = CFG.MESES[mes-1];
  var horas     = {"1":{}, "2":{}};
  var asist     = hojaA.getDataRange().getValues();

  for (var i = 1; i < asist.length; i++) {
    var id   = String(asist[i][0]).trim();
    var tipo = String(asist[i][3]).trim();
    var ts   = new Date(asist[i][2]);
    var hap  = parseFloat(asist[i][8]) || 0;
    if (!id || tipo !== CFG.KOBO_TIPO_ENTRADA) continue;
    if (isNaN(ts) || ts.getMonth()+1 !== mes || ts.getFullYear() !== anio) continue;
    var q = ts.getDate() <= 15 ? "1" : "2";
    horas[q][id] = (horas[q][id]||0) + hap;
  }

  var part = hojaP.getDataRange().getValues();
  var fact = hojaF.getDataRange().getValues();

  ["1","2"].forEach(function(q) {
    part.slice(1).forEach(function(p) {
      var id = String(p[0]).trim(), nombre = p[1];
      if (!id || !nombre) return;
      var h = Math.round((horas[q][id]||0)*100)/100;
      var m = Math.round(h*CFG.TARIFA_HORA*100)/100;
      var filaE = -1;
      for (var j = 1; j < fact.length; j++) {
        if (String(fact[j][0])===id && fact[j][2]===nombreMes &&
            fact[j][3]===anio && String(fact[j][4])===q) { filaE=j+1; break; }
      }
      if (filaE > 0) {
        hojaF.getRange(filaE, 6, 1, 2).setValues([[h, m]]);
      } else {
        hojaF.appendRow([id, nombre, nombreMes, anio, q, h, m, "No","","No","No","",""]);
      }
    });
  });
  actualizarDashboard();
  SpreadsheetApp.getUi().alert("✅ Facturación calculada — " + nombreMes + " " + anio);
}

// Pega cada hoja del Excel como "IMPORT_Mayo", "IMPORT_Junio"... y ejecuta esto.
function importarFacturacionHistorica() { _run(function() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hojaF = _sh(CFG.HOJAS.FACTURACION);
  var total = 0;
  var meses = [
    {n:"Mayo",2025:true},{n:"Junio",2025:true},{n:"Julio",2025:true},
    {n:"Agosto",2025:true},{n:"Septiembre",2025:true},{n:"Octubre",2025:true},
    {n:"Noviembre",2025:true},{n:"Diciembre",2025:true},
    {n:"Enero",2026:true},{n:"Febrero",2026:true},{n:"Marzo",2026:true},{n:"Abril",2026:true},
  ];
  // Redefine como array de objetos {nombre, anio}
  var lista = [
    {nombre:"Mayo",anio:2025},{nombre:"Junio",anio:2025},{nombre:"Julio",anio:2025},
    {nombre:"Agosto",anio:2025},{nombre:"Septiembre",anio:2025},{nombre:"Octubre",anio:2025},
    {nombre:"Noviembre",anio:2025},{nombre:"Diciembre",anio:2025},
    {nombre:"Enero",anio:2026},{nombre:"Febrero",anio:2026},
    {nombre:"Marzo",anio:2026},{nombre:"Abril",anio:2026},
  ];
  lista.forEach(function(m) {
    var ht = ss.getSheetByName("IMPORT_" + m.nombre);
    if (!ht) return;
    var datos  = ht.getDataRange().getValues();
    var inicio = _filaEnc(datos) + 1;
    for (var i = inicio; i < datos.length; i++) {
      var fila = datos[i], nombre = fila[0];
      if (!nombre) continue;
      var q1 = _primerNum(fila,1,6), q2 = _primerNum(fila,5,12);
      if (q1) { hojaF.appendRow(["",nombre,m.nombre,m.anio,"1","",q1,"No","","No","No","",""]); total++; }
      if (q2) { hojaF.appendRow(["",nombre,m.nombre,m.anio,"2","",q2,"No","","No","No","",""]); total++; }
    }
  });
  SpreadsheetApp.getUi().alert("✅ " + total + " registros importados.");
}); }

function _filaEnc(datos) {
  var max=0, idx=0;
  for (var i=0; i<Math.min(datos.length,5); i++) {
    var c = datos[i].filter(function(v){return v!==null&&v!=="";}).length;
    if (c>max) { max=c; idx=i; }
  }
  return idx;
}
function _primerNum(fila, desde, hasta) {
  for (var i=desde; i<Math.min(fila.length,hasta); i++) {
    var v = parseFloat(fila[i]);
    if (!isNaN(v) && v>0) return v;
  }
  return null;
}

// ── Documentos ──────────────────────────────────────────────

function generarRecibosMes() { _run(function() {
  var ahora     = new Date();
  var nombreMes = CFG.MESES[ahora.getMonth()];
  var anio      = ahora.getFullYear();
  var hojaF     = _sh(CFG.HOJAS.FACTURACION);
  var datos     = hojaF.getDataRange().getValues();
  var carpeta   = _carpetaFacturas(anio, nombreMes);
  var generados = 0;
  for (var i = 1; i < datos.length; i++) {
    var f = datos[i];
    if (f[2]!==nombreMes || f[3]!==anio || !f[6] || f[6]===0) continue;
    var doc = _crearRecibo(f, carpeta);
    hojaF.getRange(i+1, 14).setValue(doc.getUrl());
    generados++;
  }
  SpreadsheetApp.getUi().alert("✅ " + generados + " recibos generados.");
}); }

function generarReporteMensual() { _run(function() {
  var ahora     = new Date();
  var nombreMes = CFG.MESES[ahora.getMonth()];
  var anio      = ahora.getFullYear();
  var hojaF     = _sh(CFG.HOJAS.FACTURACION);
  var datos     = hojaF.getDataRange().getValues();
  var carpeta   = _carpetaReportes(anio);
  var titulo    = "Reporte_" + nombreMes + "_" + anio;
  var prev = carpeta.getFilesByName(titulo);
  while (prev.hasNext()) prev.next().setTrashed(true);
  var doc  = DocumentApp.create(titulo);
  var body = doc.getBody();
  var tz   = Session.getScriptTimeZone();

  body.appendParagraph(CFG.ORG + " — Reporte " + nombreMes + " " + anio)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph("").setSpacingAfter(6);

  var agrup = {}, totalM = 0, pendientes = 0;
  datos.slice(1).forEach(function(f) {
    if (f[2]!==nombreMes || f[3]!==anio || !f[1]) return;
    var n = f[1];
    if (!agrup[n]) agrup[n] = {q1:0, q2:0, pagado:f[10]};
    if (String(f[4])==="1") agrup[n].q1 = f[6]||0;
    if (String(f[4])==="2") agrup[n].q2 = f[6]||0;
  });

  var filas = [["Participante","Q1 (Q)","Q2 (Q)","Total (Q)","Pagado"]];
  Object.keys(agrup).sort().forEach(function(n) {
    var a = agrup[n], tot = a.q1+a.q2;
    totalM += tot;
    if (a.pagado!=="Sí") pendientes++;
    filas.push([n, a.q1.toFixed(2), a.q2.toFixed(2), tot.toFixed(2), a.pagado==="Sí"?"✓":"Pendiente"]);
  });

  var tabla = body.appendTable(filas);
  tabla.getRow(0).editAsText().setBold(true).setForegroundColor("#ffffff");
  for (var c=0; c<5; c++) tabla.getRow(0).getCell(c).setBackgroundColor("#1f54a8");

  body.appendParagraph("").setSpacingAfter(4);
  body.appendTable([
    ["Total a pagar",    "Q " + totalM.toFixed(2)],
    ["Pagos pendientes", pendientes + " participantes"],
    ["Generado",         Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm")],
  ]);

  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(carpeta);
  SpreadsheetApp.getUi().alert("✅ Reporte: " + doc.getUrl());
}); }

function _crearRecibo(f, carpeta) {
  var id = f[0], nombre = f[1], mes = f[2], anio = f[3], q = f[4], monto = f[6];
  var titulo = "Recibo_" + (id||String(nombre).replace(/\s/g,"_")) + "_Q"+q+"_"+mes+"_"+anio;
  var prev = carpeta.getFilesByName(titulo);
  while (prev.hasNext()) prev.next().setTrashed(true);
  var doc = DocumentApp.create(titulo);
  var body = doc.getBody();
  body.appendParagraph(CFG.ORG + " — Recibo de Pago")
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .editAsText().setForegroundColor("#1a237e");
  body.appendParagraph("").setSpacingAfter(6);
  var t = body.appendTable([
    ["Participante", nombre],
    ["Creamos ID",   id||"—"],
    ["Período",      mes+" "+anio+" — Quincena "+q],
    ["Monto",        "Q "+parseFloat(monto).toFixed(2)],
    ["Factura",      f[7]||"No entregada"],
    ["Declaraguate", f[9]||"No"],
    ["Pagado",       f[10]||"No"],
  ]);
  t.getRow(3).editAsText().setBold(true);
  body.appendParagraph("").setSpacingAfter(8);
  body.appendParagraph("Emisión: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"))
      .setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setItalic(true).setFontSize(9);
  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(carpeta);
  return doc;
}

// ── Drive ────────────────────────────────────────────────────

function crearEstructuraDrive() { _run(function() {
  var raiz     = _getOCreate(null,  CFG.ORG + " · RRHH");
  var facturas = _getOCreate(raiz,  "Facturas");
  var reportes = _getOCreate(raiz,  "Reportes");
  var p = PropertiesService.getScriptProperties();
  p.setProperty("DRIVE_RAIZ",     raiz.getId());
  p.setProperty("DRIVE_FACTURAS", facturas.getId());
  p.setProperty("DRIVE_REPORTES", reportes.getId());
  SpreadsheetApp.getUi().alert("✅ Drive listo\n" + raiz.getUrl());
}); }

function _carpetaFacturas(anio, mes) {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty("DRIVE_FACTURAS"), r;
  if (id) { try { r = DriveApp.getFolderById(id); } catch(_) {} }
  if (!r) { crearEstructuraDrive(); r = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty("DRIVE_FACTURAS")); }
  return _getOCreate(_getOCreate(r, String(anio)), mes);
}
function _carpetaReportes(anio) {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty("DRIVE_REPORTES"), r;
  if (id) { try { r = DriveApp.getFolderById(id); } catch(_) {} }
  if (!r) { crearEstructuraDrive(); r = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty("DRIVE_REPORTES")); }
  return _getOCreate(r, String(anio));
}
function _getOCreate(padre, nombre) {
  var it = padre ? padre.getFoldersByName(nombre) : DriveApp.getFoldersByName(nombre);
  if (it.hasNext()) return it.next();
  return padre ? padre.createFolder(nombre) : DriveApp.createFolder(nombre);
}

// ── Notificaciones ───────────────────────────────────────────

function enviarRecordatorioPagos() { _run(function() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hojaF = _sh(CFG.HOJAS.FACTURACION);
  var hojaP = _sh(CFG.HOJAS.PARTICIPANTES);
  var datos = hojaF.getDataRange().getValues();
  var mes   = CFG.MESES[new Date().getMonth()];
  var anio  = new Date().getFullYear();
  var corr  = _mapaCorreos(hojaP);
  var n     = 0;
  for (var i=1; i<datos.length; i++) {
    var f = datos[i];
    if (f[2]!==mes || f[3]!==anio || f[10]==="Sí" || !f[6] || f[6]===0) continue;
    var c = corr[String(f[0]).trim()];
    if (!c) continue;
    MailApp.sendEmail({
      to:c, subject:"["+CFG.ORG+"] Pago pendiente — "+mes+" "+anio,
      body:"Hola "+f[1]+",\n\nTienes pago pendiente:\nQ"+f[4]+" — Q "+parseFloat(f[6]).toFixed(2)+
           "\n\nEntrega tu factura para procesarlo.\n\nSaludos,\n"+CFG.ORG
    });
    n++;
  }
  SpreadsheetApp.getUi().alert("✅ "+n+" recordatorios enviados.");
}); }

function enviarResumenMensual() { _run(function() {
  var hojaF = _sh(CFG.HOJAS.FACTURACION);
  var datos = hojaF.getDataRange().getValues();
  var mes   = CFG.MESES[new Date().getMonth()];
  var anio  = new Date().getFullYear();
  var total = 0, pag = 0, pend = 0, lineas = [];
  for (var i=1; i<datos.length; i++) {
    var f = datos[i];
    if (f[2]!==mes || f[3]!==anio) continue;
    var m = parseFloat(f[6])||0;
    total += m;
    if (f[10]==="Sí") { pag++; } else { pend++; }
    lineas.push("  "+f[1]+" Q"+f[4]+" — Q "+m.toFixed(2)+" — "+(f[10]==="Sí"?"PAGADO":"PENDIENTE"));
  }
  MailApp.sendEmail({
    to:CFG.CORREO_ADMIN,
    subject:"["+CFG.ORG+"] Resumen RRHH — "+mes+" "+anio,
    body:"Resumen "+mes+" "+anio+"\n══════════════════\n\n"+lineas.join("\n")+
         "\n\n══════════════════\nTotal: Q "+total.toFixed(2)+"\nPagados: "+pag+"\nPendientes: "+pend
  });
  SpreadsheetApp.getUi().alert("✅ Resumen enviado.");
}); }

function _mapaCorreos(hojaP) {
  var m = {}, d = hojaP.getDataRange().getValues();
  for (var i=1; i<d.length; i++) {
    var id = String(d[i][0]).trim(), c = d[i][13];
    if (id && c) m[id] = c;
  }
  return m;
}

// ── Dashboard ────────────────────────────────────────────────

function actualizarDashboard() { _run(function() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName(CFG.HOJAS.DASHBOARD) || ss.insertSheet(CFG.HOJAS.DASHBOARD, 0);
  var mes  = CFG.MESES[new Date().getMonth()];
  var anio = new Date().getFullYear();
  var hojaP = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES);
  var hojaF = ss.getSheetByName(CFG.HOJAS.FACTURACION);
  var activos=0, totalM=0, pag=0, pend=0;
  if (hojaP) {
    var dp = hojaP.getDataRange().getValues();
    for (var i=1; i<dp.length; i++) { if (String(dp[i][5]).toLowerCase()==="activo") activos++; }
  }
  if (hojaF) {
    var df = hojaF.getDataRange().getValues();
    for (var i=1; i<df.length; i++) {
      if (df[i][2]!==mes || df[i][3]!==anio) continue;
      totalM += parseFloat(df[i][6])||0;
      if (df[i][10]==="Sí") { pag++; } else { pend++; }
    }
  }
  var ts    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  var filas = [
    ["MÉTRICA","VALOR"],
    ["Participantes activos", activos],
    ["Total a pagar ("+mes+")", "Q "+totalM.toFixed(2)],
    ["Quincenas pagadas", pag],
    ["Quincenas pendientes", pend],
    ["",""],
    ["Actualizado", ts],
  ];
  dash.clearContents();
  dash.getRange(1,1,filas.length,2).setValues(filas);
  dash.getRange(1,1,1,2).setBackground("#639922").setFontColor("#fff").setFontWeight("bold");
  if (pend > 0) dash.getRange(5,2).setBackground("#fce8e6").setFontColor("#c62828");
  dash.autoResizeColumns(1,2);
}); }

// ── Participantes ────────────────────────────────────────────

// Pega "Listado Participantes" del Excel en una hoja "IMPORT_PART" y ejecuta esto.
function importarParticipantes() { _run(function() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var hojaTemp = ss.getSheetByName("IMPORT_PART");
  if (!hojaTemp) throw new Error(
    'Crea una hoja "IMPORT_PART" y pega ahí la hoja\n"Listado Participantes" del Excel.'
  );
  var hojaP  = _sh(CFG.HOJAS.PARTICIPANTES);
  var datos  = hojaTemp.getDataRange().getValues();
  var exis   = _idsExistentes(hojaP);
  var n      = 0;
  for (var i=1; i<datos.length; i++) {
    var f = datos[i];
    if (!f[1]) continue;
    var id = String(f[0]||"").trim();
    if (id && exis[id]) continue;
    hojaP.appendRow([id, f[1], f[2], f[3], f[4], f[5], f[6], f[7], f[8], f[9], f[10],
                     "", "", "", "", "", ""]);
    n++;
  }
  SpreadsheetApp.getUi().alert("✅ "+n+" participantes importados.");
}); }

function _idsExistentes(hoja) {
  var m = {}, d = hoja.getDataRange().getValues();
  for (var i=1; i<d.length; i++) { var id=String(d[i][0]).trim(); if (id) m[id]=true; }
  return m;
}

// ── Triggers ─────────────────────────────────────────────────

function configurarTriggers() { _run(function() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("importarDesdeKobo").timeBased().everyHours(1).create();
  ScriptApp.newTrigger("calcularFacturacionMes").timeBased().everyDays(1).atHour(0).create();
  ScriptApp.newTrigger("actualizarDashboard").timeBased().everyMinutes(30).create();
  ScriptApp.newTrigger("enviarResumenMensual").timeBased().onMonthDay(1).atHour(8).create();
  ScriptApp.newTrigger("enviarRecordatorioPagos").timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(9).create();
  SpreadsheetApp.getUi().alert(
    "✅ Automatizaciones activas:\n"+
    "• Kobo: cada hora\n• Facturación: diario 00:00\n"+
    "• Dashboard: cada 30 min\n• Resumen: día 1 / 08:00\n• Recordatorios: viernes 09:00"
  );
}); }
