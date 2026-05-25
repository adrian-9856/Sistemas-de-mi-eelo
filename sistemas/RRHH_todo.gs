// ============================================================
// SISTEMA RRHH — Mi eelo
// Google Sheets: "Mi eelo · RRHH"
// Maneja: Participantes, Asistencia Kobo, Facturación
// ============================================================

const CFG = {
  ORG:          "Mi eelo",
  CORREO_ADMIN: "adrian@creamosguatemala.org",
  TARIFA_HORA:  25,   // Q por hora trabajada

  HOJAS: {
    PARTICIPANTES: "PARTICIPANTES",
    ASISTENCIA:    "ASISTENCIA",
    FACTURACION:   "FACTURACION",
    DASHBOARD:     "DASHBOARD",
  },

  // KoboToolbox
  KOBO_TOKEN:     "TU_TOKEN_AQUI",
  KOBO_ASSET_UID: "TU_ASSET_UID_AQUI",

  // Campo "Participante" en Kobo tiene formato: "Nombre (CREAMOS_ID)"
  // Ej: "Juana del Rosario Vicente Choy (JUVI281187)"
  KOBO_TIPO_ENTRADA: "🟢 Entrada",
  KOBO_TIPO_SALIDA:  "🔴 Salida",

  MESES: ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
          "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"],

  DRIVE: {
    RAIZ:      "",  // "Mi eelo · RRHH"
    FACTURAS:  "",  // subcarpeta Facturas/
  },
};
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("👥 RRHH")
    .addSubMenu(ui.createMenu("📂 Carga inicial (una sola vez)")
      .addItem("👥 Importar participantes desde Excel",     "importarParticipantes")
      .addItem("📋 Importar asistencia histórica (Planilla)","importarAsistenciaHistorica")
      .addItem("💰 Importar facturación histórica (14 meses)","importarFacturacionHistorica")
    )
    .addSeparator()
    .addItem("📥 Importar asistencia desde Kobo",           "importarDesdeKobo")
    .addItem("🔗 Emparejar entradas/salidas → horas",       "emparejarAsistencia")
    .addSeparator()
    .addItem("💰 Calcular facturación del mes actual",       "calcularFacturacionMes")
    .addItem("📄 Generar recibos de pago (Google Docs)",     "generarRecibosMes")
    .addItem("📊 Generar reporte mensual",                   "generarReporteMensual")
    .addSeparator()
    .addItem("🔔 Enviar recordatorio de pagos pendientes",   "enviarRecordatorioPagos")
    .addItem("📬 Enviar resumen mensual al admin",           "enviarResumenMensual")
    .addSeparator()
    .addItem("🔄 Actualizar Dashboard",                     "actualizarDashboard")
    .addItem("📁 Crear estructura en Drive",                "crearEstructuraDrive")
    .addItem("⚙️  Configurar automatizaciones",              "configurarTriggers")
    .addToUi();
}

function onEdit(e) {
  // Pago marcado en FACTURACION (col K = 11) → refresca dashboard
  if (e.range.getSheet().getName() === CFG.HOJAS.FACTURACION &&
      e.range.getColumn() === 11) {
    actualizarDashboard();
  }
}
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
// ============================================================
// 03_Facturacion.gs — Cálculo y migración de facturación
//
// Columnas de FACTURACION:
//  A CreAmosID  B Nombre  C Mes  D Anio  E Quincena
//  F Horas      G Monto   H Factura_Entregada  I Num_Factura
//  J Declaraguate  K Pagado  L Fecha_Pago  M Comentarios
// ============================================================

// Calcula horas y montos del mes actual desde ASISTENCIA.
function calcularFacturacionMes() {
  var ahora = new Date();
  _calcular(ahora.getMonth() + 1, ahora.getFullYear());
}

function _calcular(mes, anio) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var hojaA    = ss.getSheetByName(CFG.HOJAS.ASISTENCIA);
  var hojaP    = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES);
  var hojaF    = ss.getSheetByName(CFG.HOJAS.FACTURACION);
  var nombreMes = CFG.MESES[mes - 1];

  // Acumula horas a pagar por ID + quincena
  var horas = { "1": {}, "2": {} };
  var asist = hojaA.getDataRange().getValues();

  for (var i = 1; i < asist.length; i++) {
    var id   = String(asist[i][0]).trim();
    var tipo = String(asist[i][3]).trim();
    var ts   = new Date(asist[i][2]);
    var hap  = parseFloat(asist[i][8]) || 0;

    if (!id || tipo !== CFG.KOBO_TIPO_ENTRADA) continue;
    if (isNaN(ts) || ts.getMonth()+1 !== mes || ts.getFullYear() !== anio) continue;

    var q = ts.getDate() <= 15 ? "1" : "2";
    horas[q][id] = (horas[q][id] || 0) + hap;
  }

  // Actualiza o crea filas en FACTURACION
  var part  = hojaP.getDataRange().getValues();
  var fact  = hojaF.getDataRange().getValues();

  ["1","2"].forEach(function(q) {
    part.slice(1).forEach(function(p) {
      var id     = String(p[0]).trim();
      var nombre = p[1];
      if (!id || !nombre) return;

      var h = Math.round((horas[q][id] || 0) * 100) / 100;
      var m = Math.round(h * CFG.TARIFA_HORA * 100) / 100;

      var filaExist = -1;
      for (var j = 1; j < fact.length; j++) {
        if (String(fact[j][0]) === id && fact[j][2] === nombreMes &&
            fact[j][3] === anio && String(fact[j][4]) === q) {
          filaExist = j + 1; break;
        }
      }

      if (filaExist > 0) {
        hojaF.getRange(filaExist, 6, 1, 2).setValues([[h, m]]);
      } else {
        hojaF.appendRow([id, nombre, nombreMes, anio, q, h, m, "No","","No","No","",""]);
      }
    });
  });

  actualizarDashboard();
  SpreadsheetApp.getUi().alert("✅ Facturación calculada — " + nombreMes + " " + anio);
}

// ── Importación histórica ──────────────────────────────────

// Pega cada hoja del Excel (Mayo, Junio...) como "IMPORT_Mayo", "IMPORT_Junio"...
function importarFacturacionHistorica() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hojaF = ss.getSheetByName(CFG.HOJAS.FACTURACION);
  var total = 0;

  var meses = [
    {nombre:"Mayo",      anio:2025},
    {nombre:"Junio",     anio:2025},
    {nombre:"Julio",     anio:2025},
    {nombre:"Agosto",    anio:2025},
    {nombre:"Septiembre",anio:2025},
    {nombre:"Octubre",   anio:2025},
    {nombre:"Noviembre", anio:2025},
    {nombre:"Diciembre", anio:2025},
    {nombre:"Enero",     anio:2026},
    {nombre:"Febrero",   anio:2026},
    {nombre:"Marzo",     anio:2026},
    {nombre:"Abril",     anio:2026},
  ];

  meses.forEach(function(m) {
    var hTemp = ss.getSheetByName("IMPORT_" + m.nombre);
    if (!hTemp) return;

    var datos  = hTemp.getDataRange().getValues();
    var inicio = _filaEncabezado(datos) + 1;

    for (var i = inicio; i < datos.length; i++) {
      var fila   = datos[i];
      var nombre = fila[0];
      if (!nombre) continue;

      var q1 = _primerNumero(fila, 1, 6);
      var q2 = _primerNumero(fila, 5, 12);

      if (q1) {
        hojaF.appendRow(["", nombre, m.nombre, m.anio, "1", "", q1, "No","","No","No","",""]);
        total++;
      }
      if (q2) {
        hojaF.appendRow(["", nombre, m.nombre, m.anio, "2", "", q2, "No","","No","No","",""]);
        total++;
      }
    }
  });

  SpreadsheetApp.getUi().alert("✅ " + total + " registros históricos importados.");
}

function _filaEncabezado(datos) {
  var max = 0, idx = 0;
  for (var i = 0; i < Math.min(datos.length, 5); i++) {
    var cnt = datos[i].filter(function(v){ return v !== null && v !== ""; }).length;
    if (cnt > max) { max = cnt; idx = i; }
  }
  return idx;
}

function _primerNumero(fila, desde, hasta) {
  for (var i = desde; i < Math.min(fila.length, hasta); i++) {
    var v = parseFloat(fila[i]);
    if (!isNaN(v) && v > 0) return v;
  }
  return null;
}
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
// ============================================================
// 05_Drive.gs — Estructura de carpetas en Drive para RRHH
//
// Mi eelo · RRHH/
// ├── Facturas/
// │   └── 2026/
// │       ├── Enero/
// │       └── .../
// └── Reportes/
//     └── 2026/
// ============================================================

function crearEstructuraDrive() {
  var raiz     = _getOCreate(null,   CFG.ORG + " · RRHH");
  var facturas = _getOCreate(raiz,   "Facturas");
  var reportes = _getOCreate(raiz,   "Reportes");

  var props = PropertiesService.getScriptProperties();
  props.setProperty("DRIVE_RAIZ",     raiz.getId());
  props.setProperty("DRIVE_FACTURAS", facturas.getId());
  props.setProperty("DRIVE_REPORTES", reportes.getId());

  SpreadsheetApp.getUi().alert("✅ Estructura Drive lista\n" + raiz.getUrl());
}

function _getOCreate(padre, nombre) {
  var iter = padre ? padre.getFoldersByName(nombre) : DriveApp.getFoldersByName(nombre);
  if (iter.hasNext()) return iter.next();
  return padre ? padre.createFolder(nombre) : DriveApp.createFolder(nombre);
}
// ============================================================
// 06_Notificaciones.gs — Correos automáticos
// ============================================================

// Envía recordatorio a participantes con pagos pendientes del mes actual.
function enviarRecordatorioPagos() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hojaF = ss.getSheetByName(CFG.HOJAS.FACTURACION);
  var hojaP = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES);
  var datos = hojaF.getDataRange().getValues();
  var ahora = new Date();
  var mes   = CFG.MESES[ahora.getMonth()];
  var anio  = ahora.getFullYear();

  var correos  = _mapaCorreos(hojaP);
  var enviados = 0;

  for (var i = 1; i < datos.length; i++) {
    var f = datos[i];
    if (f[2] !== mes || f[3] !== anio) continue;
    if (f[10] === "Sí") continue;     // ya pagado
    if (!f[6] || f[6] === 0) continue;

    var correo = correos[String(f[0]).trim()];
    if (!correo) continue;

    MailApp.sendEmail({
      to:      correo,
      subject: "[" + CFG.ORG + "] Recordatorio de pago — " + mes + " " + anio,
      body:
        "Hola " + f[1] + ",\n\n" +
        "Te recordamos que tienes un pago pendiente:\n\n" +
        "  Período: " + mes + " " + anio + " — Quincena " + f[4] + "\n" +
        "  Monto:   Q " + parseFloat(f[6]).toFixed(2) + "\n\n" +
        "Recuerda entregar tu factura para procesar el pago.\n\n" +
        "Saludos,\n" + CFG.ORG,
    });
    enviados++;
  }

  SpreadsheetApp.getUi().alert("✅ " + enviados + " recordatorios enviados.");
}

// Envía resumen mensual al administrador.
function enviarResumenMensual() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hojaF = ss.getSheetByName(CFG.HOJAS.FACTURACION);
  var datos = hojaF.getDataRange().getValues();
  var ahora = new Date();
  var mes   = CFG.MESES[ahora.getMonth()];
  var anio  = ahora.getFullYear();

  var totalM = 0, pagados = 0, pendientes = 0;
  var lineas = [];

  for (var i = 1; i < datos.length; i++) {
    var f = datos[i];
    if (f[2] !== mes || f[3] !== anio) continue;
    var m = parseFloat(f[6]) || 0;
    totalM += m;
    if (f[10] === "Sí") { pagados++; } else { pendientes++; }
    lineas.push("  " + f[1] + " Q" + f[4] + " — Q " + m.toFixed(2) +
                " — " + (f[10] === "Sí" ? "PAGADO" : "PENDIENTE"));
  }

  MailApp.sendEmail({
    to:      CFG.CORREO_ADMIN,
    subject: "[" + CFG.ORG + "] Resumen RRHH — " + mes + " " + anio,
    body:
      "Resumen de facturación " + mes + " " + anio + "\n" +
      "═══════════════════════════════\n\n" +
      lineas.join("\n") + "\n\n" +
      "═══════════════════════════════\n" +
      "Total a pagar: Q " + totalM.toFixed(2) + "\n" +
      "Pagados:       " + pagados + "\n" +
      "Pendientes:    " + pendientes,
  });

  SpreadsheetApp.getUi().alert("✅ Resumen enviado a " + CFG.CORREO_ADMIN);
}

// Mapeo CreAmosID → correo desde PARTICIPANTES (col A → col N)
function _mapaCorreos(hojaP) {
  var mapa = {}, datos = hojaP.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var id     = String(datos[i][0]).trim();
    var correo = datos[i][13]; // col N = Correo_Electronico
    if (id && correo) mapa[id] = correo;
  }
  return mapa;
}
// ============================================================
// 07_Dashboard.gs — Métricas RRHH en tiempo real
// ============================================================

function actualizarDashboard() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(CFG.HOJAS.DASHBOARD)
            || ss.insertSheet(CFG.HOJAS.DASHBOARD, 0);

  var ahora    = new Date();
  var mes      = CFG.MESES[ahora.getMonth()];
  var anio     = ahora.getFullYear();

  var hojaP = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES);
  var hojaF = ss.getSheetByName(CFG.HOJAS.FACTURACION);

  var activos    = 0;
  var totalMonto = 0, pagados = 0, pendientes = 0;

  // Participantes activos
  if (hojaP) {
    var dp = hojaP.getDataRange().getValues();
    for (var i = 1; i < dp.length; i++) {
      if (String(dp[i][5]).toLowerCase() === "activo") activos++;
    }
  }

  // Facturación del mes
  if (hojaF) {
    var df = hojaF.getDataRange().getValues();
    for (var i = 1; i < df.length; i++) {
      if (df[i][2] !== mes || df[i][3] !== anio) continue;
      totalMonto += parseFloat(df[i][6]) || 0;
      if (df[i][10] === "Sí") { pagados++; } else { pendientes++; }
    }
  }

  var ts    = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  var filas = [
    ["MÉTRICA",                    "VALOR"],
    ["Participantes activos",       activos],
    ["Total a pagar (" + mes + ")", "Q " + totalMonto.toFixed(2)],
    ["Quincenas pagadas",           pagados],
    ["Quincenas pendientes",        pendientes],
    ["",                            ""],
    ["Actualizado",                 ts],
  ];

  hoja.clearContents();
  hoja.getRange(1, 1, filas.length, 2).setValues(filas);
  hoja.getRange(1, 1, 1, 2).setBackground("#639922").setFontColor("#fff").setFontWeight("bold");

  if (pendientes > 0) {
    hoja.getRange(5, 2).setBackground("#fce8e6").setFontColor("#c62828");
  }

  hoja.autoResizeColumns(1, 2);
}
// ============================================================
// 08_Participantes.gs — Importar participantes desde Excel
// ============================================================

// Pega la hoja "Listado Participantes" del Excel como "IMPORT_PART"
// y luego ejecuta esta función.
function importarParticipantes() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var hojaTemp = ss.getSheetByName("IMPORT_PART");
  var hojaP    = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES);

  if (!hojaTemp) {
    SpreadsheetApp.getUi().alert(
      "Crea una hoja 'IMPORT_PART' y pega ahí la hoja\n" +
      "'Listado Participantes' del archivo Participantes___mi_eelo_26.xlsx"
    );
    return;
  }

  var datos = hojaTemp.getDataRange().getValues();
  var existentes = _idsExistentes(hojaP);
  var importados = 0;

  // Fila 1 es encabezado: Creamos ID, Nombre, Proyecto, División, Programa, Estado...
  for (var i = 1; i < datos.length; i++) {
    var f = datos[i];
    if (!f[1]) continue; // sin nombre → salta

    var id = String(f[0] || "").trim();
    if (id && existentes[id]) continue; // ya existe

    hojaP.appendRow([
      id,      // A: Creamos ID
      f[1],    // B: Nombre
      f[2],    // C: Proyecto
      f[3],    // D: División
      f[4],    // E: Programa
      f[5],    // F: Estado
      f[6],    // G: Etapa
      f[7],    // H: Educación
      f[8],    // I: Apoyo_Emocional
      f[9],    // J: Inclusion_Laboral
      f[10],   // K: Categoria
      "",      // L: DPI
      "",      // M: NIT
      "",      // N: Correo_Electronico
      "",      // O: Banco
      "",      // P: Numero_Cuenta
      "",      // Q: Forma_Pago
    ]);
    importados++;
  }

  SpreadsheetApp.getUi().alert("✅ " + importados + " participantes importados.");
}

function _idsExistentes(hoja) {
  var mapa = {}, datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var id = String(datos[i][0]).trim();
    if (id) mapa[id] = true;
  }
  return mapa;
}
// ============================================================
// 09_Triggers.gs — Automatizaciones (ejecutar una sola vez)
// ============================================================

function configurarTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });

  // Kobo: importar cada hora
  ScriptApp.newTrigger("importarDesdeKobo")
    .timeBased().everyHours(1).create();

  // Facturación: recalcular diario a medianoche
  ScriptApp.newTrigger("calcularFacturacionMes")
    .timeBased().everyDays(1).atHour(0).create();

  // Dashboard: actualizar cada 30 minutos
  ScriptApp.newTrigger("actualizarDashboard")
    .timeBased().everyMinutes(30).create();

  // Resumen mensual: día 1 de cada mes a las 8am
  ScriptApp.newTrigger("enviarResumenMensual")
    .timeBased().onMonthDay(1).atHour(8).create();

  // Recordatorio pagos: cada viernes a las 9am
  ScriptApp.newTrigger("enviarRecordatorioPagos")
    .timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(9).create();

  SpreadsheetApp.getUi().alert(
    "✅ Automatizaciones activas:\n" +
    "• Importar Kobo: cada hora\n" +
    "• Facturación: diario 00:00\n" +
    "• Dashboard: cada 30 min\n" +
    "• Resumen mensual: día 1 / 08:00\n" +
    "• Recordatorio pagos: viernes 09:00"
  );
}
