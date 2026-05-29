// ============================================================
// SISTEMA RRHH — Mi eelo
// ============================================================

const CFG = {
  ORG:          "Mi eelo",
  CORREO_ADMIN: "adrian@creamosguatemala.org",
  // Tarifas por categoría (Q por hora) — A=Q16.50 B=Q15.75 C=Q15.00 D=Q14.00
  CATEGORIAS:   { A: 16.50, B: 15.75, C: 15.00, D: 14.00 },
  IVA_PCT:      0.05,  // 5% Pequeño Contribuyente Guatemala (solo quien tiene factura)
  HORAS_JORNADA_NORMAL: 7,
  KOBO_URL_CSV: "https://kf.kobotoolbox.org/api/v2/assets/agi395bJj6ojXJzPPDT9n6/export-settings/es4oUjEmPvovgLd6Y5yrQ4K/data.csv",
  KOBO_TIPO_ENTRADA: "🟢 Entrada",
  KOBO_TIPO_SALIDA:  "🔴 Salida",
  HOJAS: {
    PARTICIPANTES: "PARTICIPANTES",
    ASISTENCIA:    "ASISTENCIA",
    FACTURACION:   "FACTURACION",
    DASHBOARD:     "DASHBOARD",
    DATOS_KOBO:    "DatosKobo",
    CLASIFICACION: "CLASIFICACION",
  },
  MESES: ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
          "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"],
  DIAS_SEMANA: ["Domingo","Lunes","Martes","Miércoles","Jueves","Viernes","Sábado"],
};

/*
 PARTICIPANTES — 20 columnas (A–T)
 A  Creamos_ID (0)          K  Categoria (10)       — A/B/C/D
 B  Nombre (1)               L  Tarifa_Hora (11)     — Q/hr (auto desde Categoria)
 C  Proyecto (2)             M  Tiene_Factura (12)   — Sí/No (aplica IVA)
 D  Division (3)             N  DPI (13)
 E  Programa (4)             O  NIT (14)
 F  Estado (5)               P  Correo (15)
 G  Etapa (6)                Q  Banco (16)
 H  Educacion (7)            R  Num_Cuenta (17)
 I  Apoyo_Emocional (8)      S  Forma_Pago (18)
 J  Inclusion_Laboral (9)    T  URL_Doc_Proceso (19)

 FACTURACION — 21 columnas (A–U)
 A  Creamos_ID (0)           L  IVA_5pct (11)        — Base × 5% (solo si Tiene_IVA=Sí)
 B  Nombre (1)               M  Total_Factura (12)   — lo que PAGA la org (Base + IVA)
 C  Mes (2)                  N  Monto_Neto (13)      — lo que QUEDA al participante (= Base)
 D  Anio (3)                 O  Factura_Entregada (14)
 E  Quincena (4)             P  Numero_Factura (15)
 F  Horas_Trabajadas (5)     Q  Declaraguate (16)
 G  Horas_A_Reponer (6)      R  Pagado (17)
 H  Horas_A_Pagar (7)  F+G   S  Fecha_Pago (18)
 I  Tarifa_Hora (8)          T  Comentarios (19)
 J  Monto_Base (9)  H×I      U  URL_Recibo (20)
 K  Tiene_IVA (10) Sí/No
*/

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
    .addItem("📥 Importar asistencia desde Kobo",       "importarDesdeKobo")
    .addItem("🔗 Emparejar entradas/salidas → horas",   "emparejarAsistencia")
    .addItem("✨ Normalizar nombres y datos Kobo",       "normalizarTodo")
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

// onEdit: col R (18) = Pagado en FACTURACION | col K (11) = Categoria en PARTICIPANTES
function onEdit(e) {
  var sheet = e.range.getSheet();
  var nombre = sheet.getName();
  var col    = e.range.getColumn();
  var fila   = e.range.getRow();

  // FACTURACION — Pagado cambia → actualizar Dashboard
  if (nombre === CFG.HOJAS.FACTURACION && col === 18) {
    try { actualizarDashboard(); } catch(_) {}
  }

  // PARTICIPANTES — Categoria cambia → auto-llenar Tarifa_Hora
  if (nombre === CFG.HOJAS.PARTICIPANTES && col === 11 && fila >= 2) {
    var cat = String(e.range.getValue()).trim().toUpperCase();
    var tarifa = CFG.CATEGORIAS[cat];
    if (tarifa) sheet.getRange(fila, 12).setValue(tarifa);
  }
}

// ── Crear hojas ───────────────────────────────────────────────

function crearHojas() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // CLASIFICACION — referencia de categorías y tarifas
  var hCl = ss.getSheetByName(CFG.HOJAS.CLASIFICACION) || ss.insertSheet(CFG.HOJAS.CLASIFICACION);
  if (hCl.getLastRow() === 0) {
    hCl.getRange(1,1,1,3).setValues([["Categoría","Tarifa_Hora (Q/hr)","Descripción"]]);
    hCl.getRange(2,1,4,3).setValues([
      ["A", 16.50, "Categoría A — Mayor antigüedad/responsabilidad"],
      ["B", 15.75, "Categoría B"],
      ["C", 15.00, "Categoría C"],
      ["D", 14.00, "Categoría D — Ingreso reciente"],
    ]);
    _fmtEnc(hCl, "#ff6f00");
    hCl.getRange("B2:B5").setNumberFormat("Q#,##0.00");
  }

  // PARTICIPANTES — 20 cols (A–T)
  var hP = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES) || ss.insertSheet(CFG.HOJAS.PARTICIPANTES);
  if (hP.getLastRow() === 0) {
    hP.appendRow([
      "Creamos_ID","Nombre","Proyecto","Division","Programa","Estado","Etapa",
      "Educacion","Apoyo_Emocional","Inclusion_Laboral",
      "Categoria","Tarifa_Hora","Tiene_Factura",
      "DPI","NIT","Correo","Banco","Num_Cuenta","Forma_Pago","URL_Doc_Proceso"
    ]);
    _fmtEnc(hP, "#639922");
    var vEstado = SpreadsheetApp.newDataValidation().requireValueInList(["Activo","Inactivo","Egresado"],true).build();
    var vCat    = SpreadsheetApp.newDataValidation().requireValueInList(["A","B","C","D"],true).build();
    var vSiNo   = SpreadsheetApp.newDataValidation().requireValueInList(["Sí","No"],true).build();
    hP.getRange("F2:F500").setDataValidation(vEstado);
    hP.getRange("K2:K500").setDataValidation(vCat);
    hP.getRange("M2:M500").setDataValidation(vSiNo);
    hP.getRange("L2:L500").setNumberFormat("Q#,##0.00");
    hP.setColumnWidth(20, 300); // URL_Doc_Proceso
  }

  // ASISTENCIA — 10 cols
  var hA = ss.getSheetByName(CFG.HOJAS.ASISTENCIA) || ss.insertSheet(CFG.HOJAS.ASISTENCIA);
  if (hA.getLastRow() === 0) {
    hA.appendRow([
      "Nombre","Creamos_ID","Fecha","Tipo",
      "Horas_Trabajadas","Es_Dia_Estudio","Es_Terapia","Porcentaje_Pago","Horas_A_Pagar","Clave_Dia"
    ]);
    _fmtEnc(hA, "#1f54a8");
  }

  // FACTURACION — 21 cols (A–U)
  var hF = ss.getSheetByName(CFG.HOJAS.FACTURACION) || ss.insertSheet(CFG.HOJAS.FACTURACION);
  if (hF.getLastRow() === 0) {
    hF.appendRow([
      "Creamos_ID","Nombre","Mes","Anio","Quincena",
      "Horas_Trabajadas","Horas_A_Reponer","Horas_A_Pagar",
      "Tarifa_Hora","Monto_Base","Tiene_IVA","IVA_5pct","Total_Factura","Monto_Neto",
      "Factura_Entregada","Numero_Factura","Declaraguate","Pagado",
      "Fecha_Pago","Comentarios","URL_Recibo"
    ]);
    _fmtEnc(hF, "#639922");
    var v2 = SpreadsheetApp.newDataValidation().requireValueInList(["Sí","No"],true).build();
    hF.getRange("K2:K2000").setDataValidation(v2); // Tiene_IVA
    hF.getRange("O2:O2000").setDataValidation(v2); // Factura_Entregada
    hF.getRange("Q2:Q2000").setDataValidation(v2); // Declaraguate
    hF.getRange("R2:R2000").setDataValidation(v2); // Pagado
    hF.getRange("I2:N2000").setNumberFormat("Q#,##0.00"); // tarifa + montos
    hF.setColumnWidth(21, 300); // URL_Recibo
  }

  ["Hoja 1","Sheet1"].forEach(function(n) {
    var h = ss.getSheetByName(n);
    if (h && ss.getSheets().length > 3) ss.deleteSheet(h);
  });

  ss.setActiveSheet(hP);
  _alert(
    "✅ Hojas creadas:\n" +
    "• CLASIFICACION (categorías A/B/C/D)\n" +
    "• PARTICIPANTES (20 cols — incl. Categoría, Tarifa, Tiene_Factura)\n" +
    "• ASISTENCIA\n" +
    "• FACTURACION (21 cols — incl. Horas_A_Reponer, Tarifa por participante)\n\n" +
    "Tarifas: A=Q16.50 | B=Q15.75 | C=Q15.00 | D=Q14.00\n" +
    "IVA 5%: solo participantes con Tiene_Factura=Sí\n\n" +
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
    if (String(dp[i][0]).trim() === id) { ui.alert("Ya existe el ID " + id); return; }
  }

  var carpeta = _carpetaDP();
  var doc = _abrirOCrearDocProceso(id, nombre, carpeta, "");
  var url = doc.getUrl();
  var part = { id:id, nombre:nombre, proyecto:"", division:"", programa:"", estado:"Activo",
               etapa:"", educacion:"", apoyoEmocional:"", inclusionLaboral:"",
               categoria:"", tarifa:"", tieneFactura:"No",
               dpi:"", nit:"", correo:"", banco:"", numCuenta:"", formaPago:"" };
  _escribirContenidoDP(doc, part, [], []);

  // 20 cols: ID | Nombre | Proyecto | Division | Programa | Estado | Etapa |
  //           Educacion | Apoyo | Inclusion | Categoria | Tarifa | Tiene_Factura |
  //           DPI | NIT | Correo | Banco | Num_Cuenta | Forma_Pago | URL_DP
  hP.appendRow([id, nombre, "", "", "", "Activo", "", "", "", "", "", "", "No", "", "", "", "", "", "", url]);
  hP.setActiveRange(hP.getRange(hP.getLastRow(), 1));
  ui.alert("✅ Participante registrado: " + nombre + "\nDocumento de Proceso:\n" + url);
}); }

// ── Documentos de Proceso (DP) ────────────────────────────────

function generarDpFilaActiva() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getActiveSheet().getName() !== CFG.HOJAS.PARTICIPANTES) {
    _alert("Selecciona primero una fila en PARTICIPANTES."); return;
  }
  var fila = ss.getActiveRange().getRow();
  if (fila < 2) { _alert("Selecciona una fila de datos (no el encabezado)."); return; }
  var url = _sincronizarDP(_sh(CFG.HOJAS.PARTICIPANTES), fila);
  _alert("✅ DP actualizado.\n" + url);
}); }

function actualizarTodosLosDps() { _run(function() {
  var hP = _sh(CFG.HOJAS.PARTICIPANTES);
  var datos = hP.getDataRange().getValues();
  var n = 0, errs = [];
  for (var i = 1; i < datos.length; i++) {
    if (!String(datos[i][0]).trim()) continue;
    try { _sincronizarDP(hP, i + 1); n++; }
    catch(err) { errs.push(String(datos[i][0]) + ": " + err.message); }
  }
  var msg = "✅ " + n + " DPs actualizados.";
  if (errs.length) msg += "\n\n⚠️ Errores:\n" + errs.join("\n");
  _alert(msg);
}); }

function _sincronizarDP(hP, fila) {
  var datos = hP.getDataRange().getValues();
  var f     = datos[fila - 1];
  var id     = String(f[0]).trim();
  var nombre = String(f[1]).trim();
  if (!id || !nombre) throw new Error("Fila " + fila + " sin ID o nombre.");

  var part = {
    id: id, nombre: nombre,
    proyecto:         String(f[2]  || ""), division:         String(f[3]  || ""),
    programa:         String(f[4]  || ""), estado:           String(f[5]  || ""),
    etapa:            String(f[6]  || ""), educacion:        String(f[7]  || ""),
    apoyoEmocional:   String(f[8]  || ""), inclusionLaboral: String(f[9]  || ""),
    categoria:        String(f[10] || ""), tarifa:           String(f[11] || ""),
    tieneFactura:     String(f[12] || ""),
    dpi:              String(f[13] || ""), nit:              String(f[14] || ""),
    correo:           String(f[15] || ""), banco:            String(f[16] || ""),
    numCuenta:        String(f[17] || ""), formaPago:        String(f[18] || "")
  };
  var urlActual = String(f[19] || "");

  // Asistencia del participante
  var hA    = _sh(CFG.HOJAS.ASISTENCIA);
  var aRows = hA.getDataRange().getValues();
  var asistencia = [];
  for (var i = 1; i < aRows.length; i++) {
    if (String(aRows[i][0]).trim() !== nombre && String(aRows[i][1]).trim() !== id) continue;
    asistencia.push({
      fecha:       aRows[i][2], tipo:         aRows[i][3],
      horas:       aRows[i][4], esDiaEstudio: aRows[i][5],
      esTerapia:   aRows[i][6], pct:          aRows[i][7], horasAPagar: aRows[i][8]
    });
  }
  asistencia.sort(function(a,b){ return new Date(b.fecha) - new Date(a.fecha); });

  // Facturación del participante
  var hF    = _sh(CFG.HOJAS.FACTURACION);
  var fRows = hF.getDataRange().getValues();
  var facturacion = [];
  for (var j = 1; j < fRows.length; j++) {
    if (String(fRows[j][0]).trim() !== id && String(fRows[j][1]).trim() !== nombre) continue;
    facturacion.push({
      mes:          fRows[j][2],  anio:         fRows[j][3],
      quincena:     fRows[j][4],  horasTrab:    fRows[j][5],
      horasReponer: fRows[j][6],  horasPagar:   fRows[j][7],
      tarifa:       fRows[j][8],  montoBase:    fRows[j][9],
      tieneIVA:     fRows[j][10], iva:          fRows[j][11],
      totalFactura: fRows[j][12], montoNeto:    fRows[j][13],
      facturaEnt:   fRows[j][14], numFactura:   fRows[j][15],
      declaraguate: fRows[j][16], pagado:       fRows[j][17],
      fechaPago:    fRows[j][18]
    });
  }

  var carpeta = _carpetaDP();
  var doc = _abrirOCrearDocProceso(id, nombre, carpeta, urlActual);
  _escribirContenidoDP(doc, part, asistencia, facturacion);

  var urlNueva = doc.getUrl();
  if (urlNueva !== urlActual) hP.getRange(fila, 20).setValue(urlNueva);
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
    if (!doc) throw new Error("No se pudo abrir el DP de " + nombre + ": " + urlExistente);
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

  var tarifahora = part.tarifa ? "Q "+part.tarifa+"/hr" : "según categoría "+part.categoria;

  // Título
  body.appendParagraph("DOCUMENTO DE PROCESO")
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .editAsText().setForegroundColor("#1a237e");
  body.appendParagraph(part.nombre || "")
      .setHeading(DocumentApp.ParagraphHeading.HEADING2)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph("Mi eelo  ·  Actualizado: " + hoy)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .editAsText().setFontSize(9).setItalic(true).setForegroundColor("#888888");
  body.appendParagraph("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
      .editAsText().setFontSize(8).setForegroundColor("#cccccc");

  // 1. Datos del participante
  var s1 = body.appendParagraph("1.  DATOS DEL PARTICIPANTE");
  s1.setHeading(DocumentApp.ParagraphHeading.HEADING3).editAsText().setForegroundColor("#639922");
  var tD = body.appendTable([
    ["Creamos ID",     part.id          ||"—", "Estado",        part.estado      ||"—"],
    ["Nombre",         part.nombre      ||"—", "Etapa",         part.etapa       ||"—"],
    ["Proyecto",       part.proyecto    ||"—", "División",      part.division    ||"—"],
    ["Programa",       part.programa    ||"—", "Categoría",     part.categoria   ||"—"],
    ["Tarifa/hora",    tarifahora,              "Tiene Factura", part.tieneFactura||"—"],
    ["DPI",            part.dpi         ||"—", "NIT",           part.nit         ||"—"],
    ["Correo",         part.correo      ||"—", "Forma de pago", part.formaPago   ||"—"],
    ["Banco",          part.banco       ||"—", "Núm. cuenta",   part.numCuenta   ||"—"],
  ]);
  for (var r=0; r<tD.getNumRows(); r++) {
    tD.getRow(r).getCell(0).setBackgroundColor("#639922").editAsText().setForegroundColor("#fff").setBold(true);
    tD.getRow(r).getCell(2).setBackgroundColor("#639922").editAsText().setForegroundColor("#fff").setBold(true);
  }

  // 2. Servicios de apoyo
  body.appendParagraph("");
  body.appendParagraph("2.  SERVICIOS DE APOYO")
      .setHeading(DocumentApp.ParagraphHeading.HEADING3)
      .editAsText().setForegroundColor("#1f54a8");
  var tA = body.appendTable([
    ["Educación",          part.educacion        ||"—"],
    ["Apoyo Emocional",    part.apoyoEmocional   ||"—"],
    ["Inclusión Laboral",  part.inclusionLaboral ||"—"],
  ]);
  for (var r=0; r<tA.getNumRows(); r++) {
    tA.getRow(r).getCell(0).setBackgroundColor("#1f54a8").editAsText().setForegroundColor("#fff").setBold(true);
  }

  // 3. Asistencia (últimos 90 días)
  body.appendParagraph("");
  body.appendParagraph("3.  HISTORIAL DE ASISTENCIA  (últimos 90 días)")
      .setHeading(DocumentApp.ParagraphHeading.HEADING3)
      .editAsText().setForegroundColor("#1f54a8");
  var hace90 = new Date(); hace90.setDate(hace90.getDate() - 90);
  var asistRec = asistencia.filter(function(a) {
    var d = new Date(a.fecha); return !isNaN(d) && d >= hace90;
  }).slice(0, 60);
  if (asistRec.length === 0) {
    body.appendParagraph("Sin registros en los últimos 90 días.")
        .editAsText().setItalic(true).setFontSize(9).setForegroundColor("#888888");
  } else {
    var totalHrsDP = 0;
    var filaAss = [["Fecha","Tipo","Horas Trab.","Día Estudio","Terapia","Horas a Pagar"]];
    asistRec.forEach(function(a) {
      var fecha = a.fecha instanceof Date
        ? Utilities.formatDate(a.fecha, tz, "dd/MM/yyyy") : String(a.fecha||"—");
      totalHrsDP += parseFloat(a.horasAPagar) || 0;
      filaAss.push([fecha, String(a.tipo||"—"), _n2(a.horas),
                    String(a.esDiaEstudio||"—"), String(a.esTerapia||"—"), _n2(a.horasAPagar)]);
    });
    _estilTablaEnc(body.appendTable(filaAss), "#1f54a8");
    body.appendParagraph("Total horas a pagar (período): " + _n2(totalHrsDP))
        .editAsText().setBold(true).setFontSize(10);
  }

  // 4. Historial de pagos
  body.appendParagraph("");
  body.appendParagraph("4.  HISTORIAL DE PAGOS  (" + tarifahora +
      (part.tieneFactura==="Sí" ? " · IVA 5% Pcv." : " · sin IVA") + ")")
      .setHeading(DocumentApp.ParagraphHeading.HEADING3)
      .editAsText().setForegroundColor("#639922");
  if (facturacion.length === 0) {
    body.appendParagraph("Sin registros de pago.")
        .editAsText().setItalic(true).setFontSize(9).setForegroundColor("#888888");
  } else {
    var totPagado = 0, totPend = 0, totIVA = 0;
    var filasPago = [["Mes","Q","Hrs Trab","Hrs Pagar","Tarifa","Base (Q)","IVA 5%","Total Fact.","Neto Part.","Pagado","Fecha"]];
    facturacion.slice(0, 30).forEach(function(p) {
      var fp = p.fechaPago instanceof Date
        ? Utilities.formatDate(p.fechaPago, tz, "dd/MM/yyyy") : String(p.fechaPago||"—");
      var tot  = parseFloat(p.totalFactura) || 0;
      var iva  = parseFloat(p.iva)          || 0;
      var base = parseFloat(p.montoBase)    || 0;
      var neto = parseFloat(p.montoNeto)    || 0;
      if (p.pagado === "Sí") totPagado += tot; else totPend += tot;
      totIVA += iva;
      filasPago.push([
        String(p.mes||"—") + " " + String(p.anio||""),
        String(p.quincena||"—"),
        _n2(p.horasTrab),
        _n2(p.horasPagar),
        "Q " + (parseFloat(p.tarifa)||0).toFixed(2),
        "Q " + base.toFixed(2),
        iva > 0 ? "Q " + iva.toFixed(2) : "—",
        "Q " + tot.toFixed(2),
        "Q " + neto.toFixed(2),
        String(p.pagado||"No"), fp
      ]);
    });
    _estilTablaEnc(body.appendTable(filasPago), "#639922");
    body.appendParagraph("");
    body.appendTable([
      ["Total pagado (factura)",  "Q " + totPagado.toFixed(2)],
      ["Pendiente",               "Q " + totPend.toFixed(2)],
      ["IVA total declarado",     "Q " + totIVA.toFixed(2)],
    ]).editAsText().setFontSize(10);
  }

  body.appendParagraph("");
  body.appendParagraph("Generado por Sistema RRHH — " + CFG.ORG + "  ·  " + hoy)
      .setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
      .editAsText().setItalic(true).setFontSize(8).setForegroundColor("#aaaaaa");
  doc.saveAndClose();
}

function _n2(v) { var n=parseFloat(v); return isNaN(n)?"0":n.toFixed(2); }
function _estilTablaEnc(tabla, color) {
  var enc = tabla.getRow(0);
  for (var c=0; c<enc.getNumCells(); c++) {
    enc.getCell(c).setBackgroundColor(color).editAsText().setForegroundColor("#fff").setBold(true);
  }
}

// ── Importar desde Kobo (CSV — sin token) ────────────────────

function importarDesdeKobo() { _run(function() {
  var resp = UrlFetchApp.fetch(CFG.KOBO_URL_CSV, { muteHttpExceptions: true });
  var code = resp.getResponseCode();
  if (code === 503) { _alert("⏳ Kobo ocupado (503). Espera 2 min e intenta de nuevo."); return; }
  if (code !== 200) throw new Error("Error Kobo HTTP " + code + ": " + resp.getContentText().substring(0,200));

  var datosNuevos = Utilities.parseCsv(resp.getContentText(), ";");
  if (datosNuevos.length < 2) { _alert("Kobo no devolvió registros."); return; }

  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);

  if (!hoja) {
    hoja = ss.insertSheet(CFG.HOJAS.DATOS_KOBO);
    hoja.getRange(1,1,datosNuevos.length,datosNuevos[0].length).setValues(datosNuevos);
    hoja.getRange(1,1,1,datosNuevos[0].length)
        .setFontWeight("bold").setBackground("#4a86e8").setFontColor("#fff");
    hoja.setFrozenRows(1);
    _limpiarColumnasKobo(hoja, datosNuevos[0]);
    _normalizarAccionSilencioso(hoja);
    emparejarAsistencia();
    _alert("✅ Importación inicial: " + (datosNuevos.length-1) + " registros.\nAsistencia emparejada.");
    return;
  }

  // Incremental — solo UUIDs nuevos
  var encNuevos = datosNuevos[0];
  var uuidColN  = _buscarIndice(encNuevos, "_uuid");
  var datosEx   = hoja.getDataRange().getValues();
  var uuidColE  = _buscarIndice(datosEx[0], "_uuid");
  var uuidsExist = {};
  if (uuidColE >= 0) {
    for (var i=1; i<datosEx.length; i++) {
      var u = String(datosEx[i][uuidColE]||"").trim(); if (u) uuidsExist[u] = true;
    }
  }
  var filasNuevas = [];
  for (var j=1; j<datosNuevos.length; j++) {
    var uid = uuidColN >= 0 ? String(datosNuevos[j][uuidColN]||"").trim() : "";
    if (!uid || !uuidsExist[uid]) filasNuevas.push(datosNuevos[j]);
  }
  if (filasNuevas.length === 0) { _alert("✅ Ya está al día. Sin registros nuevos."); return; }

  hoja.getRange(hoja.getLastRow()+1,1,filasNuevas.length,filasNuevas[0].length).setValues(filasNuevas);
  _normalizarAccionSilencioso(hoja);
  emparejarAsistencia();
  _alert("✅ " + filasNuevas.length + " registros nuevos importados.\nAsistencia emparejada.");
}); }

function _buscarIndice(enc, clave) {
  var c = clave.toLowerCase();
  for (var i=0; i<enc.length; i++) { if (String(enc[i]).trim().toLowerCase()===c) return i; }
  return -1;
}

function _limpiarColumnasKobo(hoja, enc) {
  var imp = ["start","end","ingreso","egreso","entrada","salida",
             "participante","nombre","seleccione","c_id","_uuid","uuid","accion","acción","terapia","permiso","comput"];
  for (var i=0; i<enc.length; i++) {
    var h = String(enc[i]).trim().toLowerCase();
    var esImp = h && imp.some(function(p){ return h.indexOf(p)!==-1; });
    try { if (!esImp) hoja.hideColumns(i+1); else hoja.showColumns(i+1); } catch(_) {}
  }
}

function _normalizarAccionSilencioso(hoja) {
  try {
    var datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return;
    var cols = detectarColumnas(datos[0], datos.slice(1));
    if (cols.accionUnificada === undefined) return;
    var colSub = cols.subtipoEgreso;
    if (colSub === undefined) {
      var nc = datos[0].length;
      hoja.getRange(1,nc+1).setValue("subtipo_egreso")
          .setFontWeight("bold").setBackground("#e6b8a2").setFontColor("#000");
      colSub = nc; cols.subtipoEgreso = colSub;
    }
    var SUBS = [{k:"terapia",s:"Terapia"},{k:"permiso",s:"Permiso"},{k:"comput",s:"Computacion"}];
    for (var f=1; f<datos.length; f++) {
      var valRaw = String(datos[f][cols.accionUnificada]||"").trim();
      if (!valRaw) continue;
      var tipo = obtenerTipoRegistro(datos[f], cols);
      var correcto = null, subtipo = "";
      var valLow = valRaw.toLowerCase();
      for (var s=0; s<SUBS.length; s++) { if (valLow.indexOf(SUBS[s].k)!==-1){subtipo=SUBS[s].s;break;} }
      if (tipo.esIngreso) { correcto=CFG.KOBO_TIPO_ENTRADA; subtipo=""; }
      else if (tipo.esEgreso) { correcto=CFG.KOBO_TIPO_SALIDA; }
      if (correcto) {
        var celda = hoja.getRange(f+1, cols.accionUnificada+1);
        if (valRaw!==correcto) celda.setValue(correcto);
        celda.setBackground(tipo.esIngreso?"#b7e1cd":"#f4cccc")
             .setFontColor(tipo.esIngreso?"#0b5c30":"#7f0000").setFontWeight("bold");
      }
      if (subtipo && !String(datos[f][colSub]||"").trim())
        hoja.getRange(f+1, colSub+1).setValue(subtipo);
    }
    ["start","end"].forEach(function(n){
      var c=_buscarIndice(datos[0],n); if(c<0)return;
      hoja.getRange(2,c+1,datos.length-1,1).setNumberFormat("dd/MM/yyyy HH:mm");
      for(var f=1;f<datos.length;f++){
        var v=datos[f][c];
        if(typeof v==="string"&&v.indexOf("T")!==-1){var d=new Date(v);if(!isNaN(d))hoja.getRange(f+1,c+1).setValue(d);}
      }
    });
  } catch(_) {}
}

// ── Emparejar entradas/salidas → ASISTENCIA ───────────────────

function emparejarAsistencia() { _run(function() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var hojaKobo = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (!hojaKobo) throw new Error("No existe DatosKobo. Importa primero desde Kobo.");
  var datos = hojaKobo.getDataRange().getValues();
  if (datos.length < 2) throw new Error("DatosKobo está vacío.");

  var cols           = detectarColumnas(datos[0], datos.slice(1));
  var diasEstudioMap = obtenerDiasEstudio();
  var listaTerapias  = obtenerListaTerapias();
  var mapeoNombres   = cargarMapeoNombres();

  if (cols.start === undefined || cols.participante === undefined)
    throw new Error("No se detectaron columnas start/participante en DatosKobo.\nEjecuta 'Normalizar nombres y datos Kobo' primero.");

  var grupos = {}, uuidVistos = {};
  for (var i=1; i<datos.length; i++) {
    var fila = datos[i];
    var uuid = cols.uuid !== undefined ? String(fila[cols.uuid]||"").trim() : "";
    if (uuid && uuidVistos[uuid]) continue;
    if (uuid) uuidVistos[uuid] = true;

    var nombreRaw = obtenerParticipanteFila(fila, cols);
    if (!nombreRaw) continue;
    var nombre = normalizarNombre(nombreRaw, mapeoNombres);

    var tipo = obtenerTipoRegistro(fila, cols);
    if (!tipo.esIngreso && !tipo.esEgreso) continue;

    var tsRaw = fila[cols.start];
    var ts    = tsRaw instanceof Date ? tsRaw : new Date(tsRaw);
    if (isNaN(ts)) continue;

    if (tipo.esEgreso && cols.end !== undefined && fila[cols.end]) {
      var tsEnd = fila[cols.end] instanceof Date ? fila[cols.end] : new Date(fila[cols.end]);
      if (!isNaN(tsEnd)) { var d2=(tsEnd-ts)/3600000; if(d2>0&&d2<24) ts=tsEnd; }
    }

    var clave = nombre + "|" + _dClave(ts);
    if (!grupos[clave]) grupos[clave] = { nombre:nombre, fecha:ts, ent:[], sal:[], esTerapia:false };
    if (tipo.esIngreso) grupos[clave].ent.push(ts);
    if (tipo.esEgreso)  grupos[clave].sal.push(ts);
    if (tipo.esTerapia) grupos[clave].esTerapia = true;
  }

  var filasAsist = [];
  Object.keys(grupos).forEach(function(clave) {
    var g = grupos[clave];
    if (!g.ent.length || !g.sal.length) return;
    g.ent.sort(function(a,b){return a-b;});
    g.sal.sort(function(a,b){return a-b;});
    var horas = Math.max(0, Math.round((g.sal[g.sal.length-1]-g.ent[0])/36000)/100);
    var id    = extraerCodigo(g.nombre) || "";
    var esDiaEst = esDiaDeEstudio(g.nombre, g.fecha, diasEstudioMap) ? "Sí" : "No";
    var esTer    = (listaTerapias[g.nombre] || g.esTerapia) ? "Sí" : "No";
    var pct      = (esDiaEst==="Sí" || esTer==="Sí") ? 0 : 100;
    var hap      = Math.round(horas*(pct/100)*100)/100;
    var tipo2    = esDiaEst==="Sí"?"Día de Estudio":(esTer==="Sí"?"Terapia":"Normal");
    filasAsist.push([g.nombre, id, g.ent[0], tipo2, horas, esDiaEst, esTer, pct, hap, clave]);
  });

  filasAsist.sort(function(a,b){ return new Date(b[2])-new Date(a[2]); });
  var hA = _sh(CFG.HOJAS.ASISTENCIA);
  if (hA.getLastRow() > 1) hA.deleteRows(2, hA.getLastRow()-1);
  if (filasAsist.length > 0) hA.getRange(2,1,filasAsist.length,10).setValues(filasAsist);
  _alert("✅ " + filasAsist.length + " pares entrada/salida procesados en ASISTENCIA.");
}); }

// ── Normalizar nombres y datos Kobo ──────────────────────────

function normalizarTodo() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaKobo = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (!hojaKobo) { ui.alert("Primero importa datos desde Kobo."); return; }
  var resp = ui.alert("✨ Normalizar Todo",
    "Fases:\n1️⃣  Estandarizar 🟢 Entrada / 🔴 Salida\n" +
    "2️⃣  Crear/actualizar NombresCanonicos\n3️⃣  Actualizar DiasEstudio y ListaTerapias\n\n¿Continuar?",
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  var log = [];
  _normalizarAccionSilencioso(hojaKobo);
  log.push("1️⃣  Entrada/Salida normalizados en DatosKobo.");
  var datos = hojaKobo.getDataRange().getValues();
  var colsP = _buscarColsParticipante(datos[0]);
  var nombresUnicos = {};
  for (var f=1; f<datos.length; f++) { var n=_nombreDeFila(datos[f],colsP); if(n) nombresUnicos[n]=true; }
  var mapeo = _generarMapeoNombres(Object.keys(nombresUnicos), ss);
  _escribirNombresCanonicos(mapeo, ss);
  var dup = Object.keys(mapeo).filter(function(k){return mapeo[k]!==k;}).length;
  log.push("2️⃣  NombresCanonicos: "+Object.keys(mapeo).length+" nombres, "+dup+" variantes unificadas.");
  var act = _normalizarNombresEnHojas(mapeo, ss);
  log.push("3️⃣  Hojas actualizadas: "+(act||"ninguna pendiente"));
  ui.alert("✅ NORMALIZACIÓN COMPLETA\n\n"+log.join("\n")+"\n\nYa puedes emparejar y calcular.");
}); }

function _generarMapeoNombres(todos, ss) {
  var pC={}, sC=[];
  todos.forEach(function(n){
    var cod=extraerCodigo(n);
    if(cod){if(!pC[cod])pC[cod]=[];pC[cod].push(n);}else sC.push(n);
  });
  sC.forEach(function(n,i){
    var lim=textoParaComparar(limpiarNombre(n)), enc=false;
    Object.keys(pC).forEach(function(k){
      if(enc)return;
      pC[k].forEach(function(g){ if(!enc&&nombresCoinciden(lim,textoParaComparar(limpiarNombre(g)))){pC[k].push(n);enc=true;} });
    });
    if(!enc) pC["_SIN_"+i]=[n];
  });
  var mapeo={};
  Object.keys(pC).forEach(function(k){
    var grupo=pC[k], codigoReal=k.indexOf("_SIN_")===0?"":k;
    var mejor="", maxL=0;
    grupo.forEach(function(n){var l=limpiarNombre(n);if(l.length>maxL){maxL=l.length;mejor=l;}});
    var can=codigoReal?mejor+" ("+codigoReal+")":mejor;
    grupo.forEach(function(n){mapeo[n]=can;});
  });
  var hNC=ss.getSheetByName("NombresCanonicos");
  if(hNC){
    var d=hNC.getDataRange().getValues();
    for(var f=1;f<d.length;f++){var o=String(d[f][0]||"").trim(),c=String(d[f][1]||"").trim();if(o&&c&&mapeo[o]!==undefined)mapeo[o]=c;}
  }
  return mapeo;
}
function _escribirNombresCanonicos(mapeo, ss) {
  var h=ss.getSheetByName("NombresCanonicos")||ss.insertSheet("NombresCanonicos");
  h.clearContents();
  h.getRange(1,1,1,3).setValues([["Nombre Original (Kobo)","Nombre Canónico","Código"]]);
  h.getRange(1,1,1,3).setFontWeight("bold").setBackground("#ff6f00").setFontColor("#fff").setHorizontalAlignment("center");
  h.setFrozenRows(1);
  var filas=Object.keys(mapeo).sort().map(function(k){return[k,mapeo[k],extraerCodigo(k)||""];});
  if(filas.length>0){
    h.getRange(2,1,filas.length,3).setValues(filas);
    filas.forEach(function(f,i){if(f[0]!==f[1])h.getRange(i+2,1,1,3).setBackground("#fff3e0");});
  }
  h.setColumnWidth(1,350);h.setColumnWidth(2,350);h.setColumnWidth(3,130);
}
function _normalizarNombresEnHojas(mapeo, ss) {
  var act=[];
  ["DiasEstudio","ListaTerapias"].forEach(function(nm){
    var h=ss.getSheetByName(nm);if(!h)return;
    var d=h.getDataRange().getValues(),cam=0;
    for(var f=1;f<d.length;f++){var n=String(d[f][0]||"").trim();if(n&&mapeo[n]&&mapeo[n]!==n){h.getRange(f+1,1).setValue(mapeo[n]);cam++;}}
    if(cam>0) act.push(nm+"("+cam+")");
  });
  return act.join(", ");
}

// ── Detección de columnas Kobo ────────────────────────────────

function detectarColumnas(encabezados, datosEjemplo) {
  var cols={};
  for(var i=0;i<encabezados.length;i++){
    var h=String(encabezados[i]).trim(), hLow=h.toLowerCase();
    if(hLow==="start"){cols.start=i;continue;}
    if(hLow==="end"){cols.end=i;continue;}
    if(hLow==="_uuid"){cols.uuid=i;continue;}
    if(hLow.indexOf("uuid")!==-1&&cols.uuid===undefined){cols.uuid=i;continue;}
    if(hLow.indexOf("participante")!==-1||hLow.indexOf("nombre")!==-1||hLow.indexOf("seleccione")!==-1){
      if(cols.participante===undefined)cols.participante=i; else if(cols.participante2===undefined)cols.participante2=i; continue;
    }
    if((hLow.indexOf("ingreso")!==-1||hLow.indexOf("entrada")!==-1)&&(hLow.indexOf("egreso")!==-1||hLow.indexOf("salida")!==-1)){cols.accionUnificada=i;continue;}
    if(hLow.indexOf("accion")!==-1||hLow.indexOf("acción")!==-1||hLow==="type"||hLow.indexOf("marcar")!==-1){cols.accionUnificada=i;continue;}
    if(hLow==="subtipo_egreso"){cols.subtipoEgreso=i;continue;}
    if(hLow.indexOf("/ingreso")!==-1||hLow.indexOf("/entrada")!==-1){cols.ingreso=i;continue;}
    if(hLow.indexOf("/egreso")!==-1||hLow.indexOf("/salida")!==-1){cols.egreso=i;continue;}
    if(hLow.indexOf("/terapia")!==-1){cols.terapia=i;continue;}
    if(hLow.indexOf("/permiso")!==-1){cols.permiso=i;continue;}
    if(hLow.indexOf("/comput")!==-1){cols.computacion=i;continue;}
  }
  if(cols.accionUnificada===undefined&&cols.ingreso===undefined&&datosEjemplo){
    var PALS=["ingreso","egreso","entrada","salida","terapia","permiso","comput"];
    var max=0,best=-1;
    for(var c=0;c<encabezados.length;c++){
      var score=0,lim=Math.min(datosEjemplo.length,50);
      for(var f=0;f<lim;f++){
        var v=String(datosEjemplo[f][c]||"").toLowerCase();
        for(var p=0;p<PALS.length;p++){if(v.indexOf(PALS[p])!==-1){score++;break;}}
      }
      if(score>max){max=score;best=c;}
    }
    if(best>=0&&max>=Math.max(1,datosEjemplo.length*0.3))cols.accionUnificada=best;
  }
  return cols;
}

function obtenerTipoRegistro(fila, cols) {
  var r={tipo:"",esIngreso:false,esEgreso:false,esTerapia:false,esPermiso:false,esComputacion:false};
  if(cols.accionUnificada!==undefined){
    var vR=String(fila[cols.accionUnificada]||"").trim(), v=vR.toLowerCase();
    if(v.indexOf("entrada")!==-1||v.indexOf("ingreso")!==-1||vR.indexOf("🟢")!==-1) r.esIngreso=true;
    if(v.indexOf("salida") !==-1||v.indexOf("egreso") !==-1||vR.indexOf("🔴")!==-1) r.esEgreso=true;
    if(!r.esIngreso&&!r.esEgreso){
      if(v.indexOf("terapia")!==-1){r.esTerapia=true;r.esEgreso=true;}
      if(v.indexOf("permiso")!==-1){r.esPermiso=true;r.esEgreso=true;}
      if(v.indexOf("comput") !==-1){r.esComputacion=true;r.esEgreso=true;}
    }else{
      if(v.indexOf("terapia")!==-1) r.esTerapia=true;
      if(v.indexOf("permiso")!==-1) r.esPermiso=true;
      if(v.indexOf("comput") !==-1) r.esComputacion=true;
    }
    if(r.esEgreso&&cols.subtipoEgreso!==undefined&&fila[cols.subtipoEgreso]){
      var sub=String(fila[cols.subtipoEgreso]).trim().toLowerCase();
      if(sub.indexOf("terapia")!==-1)r.esTerapia=true;
      if(sub.indexOf("permiso")!==-1)r.esPermiso=true;
      if(sub.indexOf("comput") !==-1)r.esComputacion=true;
    }
  }else{
    var _c=function(c){var v=String(fila[c]||"").trim().toLowerCase();return v==="true"||v==="1"||v==="x"||v==="yes";};
    if(cols.ingreso!==undefined&&_c(cols.ingreso)) r.esIngreso=true;
    if(cols.egreso!==undefined&&_c(cols.egreso))   r.esEgreso=true;
    if(cols.terapia!==undefined&&_c(cols.terapia)){r.esTerapia=true;r.esEgreso=true;}
    if(cols.permiso!==undefined&&_c(cols.permiso)){r.esPermiso=true;r.esEgreso=true;}
    if(cols.computacion!==undefined&&_c(cols.computacion)){r.esComputacion=true;r.esEgreso=true;}
  }
  return r;
}

function obtenerParticipanteFila(fila, cols) {
  var n1=cols.participante!==undefined?String(fila[cols.participante]||"").trim():"";
  var n2=cols.participante2!==undefined?String(fila[cols.participante2]||"").trim():"";
  if(!n1)return n2; if(n2&&n2.length>n1.length)return n2; return n1;
}
function _buscarColsParticipante(enc) {
  var r={col1:-1,col2:-1};
  for(var i=0;i<enc.length;i++){
    var h=String(enc[i]).trim().toLowerCase();
    if(h.indexOf("participante")!==-1||h.indexOf("nombre")!==-1||h.indexOf("seleccione")!==-1){
      if(r.col1===-1)r.col1=i; else if(r.col2===-1)r.col2=i;
    }
  }
  return r;
}
function _nombreDeFila(fila, c) {
  var n1=c.col1>=0?String(fila[c.col1]||"").trim():"";
  var n2=c.col2>=0?String(fila[c.col2]||"").trim():"";
  if(!n1)return n2; if(n2&&n2.length>n1.length)return n2; return n1;
}

// ── Normalización de nombres ──────────────────────────────────

function extraerCodigo(nombre) {
  var m=String(nombre).match(/([A-ZÁÉÍÓÚÑÜ]{4}\d{6})/i); return m?m[1].toUpperCase():null;
}
function limpiarNombre(nombre) {
  var s=String(nombre).replace(/^[A-ZÁÉÍÓÚÑÜ]{4}\d{6}\s*/i,"");
  s=s.replace(/\s*\([A-ZÁÉÍÓÚÑÜ]{4}\d{6}\)\s*/i,"");
  return s.replace(/^[•\s]+/,"").trim();
}
function textoParaComparar(texto) {
  return String(texto).toLowerCase()
    .replace(/[áàä]/g,"a").replace(/[éèë]/g,"e").replace(/[íìï]/g,"i")
    .replace(/[óòö]/g,"o").replace(/[úùü]/g,"u").replace(/ñ/g,"n")
    .replace(/\s+/g," ").trim();
}
function nombresCoinciden(n1, n2) {
  if(n1===n2)return true;
  if(n1.indexOf(n2)!==-1||n2.indexOf(n1)!==-1)return true;
  var c=n1.length<=n2.length?n1:n2, l=n1.length<=n2.length?n2:n1;
  var pals=c.split(" "), coin=0, sig=0;
  pals.forEach(function(p){if(p.length<3)return;sig++;if(l.indexOf(p)!==-1)coin++;});
  return sig>0&&(coin===sig||(coin>=2&&coin/sig>=0.6));
}
function cargarMapeoNombres() {
  var m={},h=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NombresCanonicos");
  if(!h)return m;
  var d=h.getDataRange().getValues();
  for(var i=1;i<d.length;i++){var o=String(d[i][0]||"").trim(),c=String(d[i][1]||"").trim();if(o&&c)m[o]=c;}
  return m;
}
function normalizarNombre(nombre, mapeo) {
  if(!mapeo||!Object.keys(mapeo).length)return nombre; return mapeo[nombre]||nombre;
}

// ── Días de estudio y terapias ────────────────────────────────

function obtenerDiasEstudio() {
  var mapa={},h=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DiasEstudio");
  if(!h)return mapa;
  var d=h.getDataRange().getValues();
  for(var f=1;f<d.length;f++){
    var p=String(d[f][0]||"").trim();if(!p)continue;
    mapa[p]={
      dias:{1:d[f][1]==="X"||d[f][1]==="x",2:d[f][2]==="X"||d[f][2]==="x",
             3:d[f][3]==="X"||d[f][3]==="x",4:d[f][4]==="X"||d[f][4]==="x",
             5:d[f][5]==="X"||d[f][5]==="x",6:d[f][6]==="X"||d[f][6]==="x",
             0:d[f][7]==="X"||d[f][7]==="x"},
      fechaInicio:d[f][8]?new Date(d[f][8]):null,
      fechaFin:d[f][9]?new Date(d[f][9]):null
    };
    if(mapa[p].fechaFin) mapa[p].fechaFin.setHours(23,59,59,999);
  }
  return mapa;
}
function esDiaDeEstudio(nombre, fecha, mapa) {
  if(!mapa[nombre])return false;
  var c=mapa[nombre];
  if(c.fechaInicio&&fecha<c.fechaInicio)return false;
  if(c.fechaFin&&fecha>c.fechaFin)return false;
  return c.dias[fecha.getDay()]===true;
}
function obtenerListaTerapias() {
  var lista={},h=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ListaTerapias");
  if(!h)return lista;
  var d=h.getDataRange().getValues();
  for(var f=1;f<d.length;f++){
    var p=String(d[f][0]||"").trim();
    if(p&&String(d[f][1]||"").trim().toUpperCase()==="X")lista[p]=true;
  }
  return lista;
}
function _dClave(d) {
  return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");
}

// ── Facturación ───────────────────────────────────────────────
// Tarifa: por categoría (A=Q16.50 B=Q15.75 C=Q15.00 D=Q14.00)
// IVA 5%: solo participantes con Tiene_Factura=Sí (Pequeño Contribuyente)
// Horas_A_Reponer: ajuste manual — se preserva en la fila existente

function calcularFacturacionMes() { _run(function() {
  var ahora = new Date();
  _calcular(ahora.getMonth()+1, ahora.getFullYear());
}); }

function _calcular(mes, anio) {
  var hojaA    = _sh(CFG.HOJAS.ASISTENCIA);
  var hojaP    = _sh(CFG.HOJAS.PARTICIPANTES);
  var hojaF    = _sh(CFG.HOJAS.FACTURACION);
  var nombreMes = CFG.MESES[mes-1];

  // Mapa de participantes: id → { tarifa, tieneFactura }
  var partRows = hojaP.getDataRange().getValues();
  var partMap  = {};
  for (var pi=1; pi<partRows.length; pi++) {
    var pid = String(partRows[pi][0]).trim();
    if (!pid) continue;
    var tarifa = parseFloat(partRows[pi][11]); // col L = Tarifa_Hora
    if (isNaN(tarifa) || tarifa <= 0) {
      var cat = String(partRows[pi][10]).trim().toUpperCase(); // col K = Categoria
      tarifa = CFG.CATEGORIAS[cat] || CFG.CATEGORIAS.C;
    }
    var t = String(partRows[pi][12]).trim().toLowerCase(); // col M = Tiene_Factura
    partMap[pid] = {
      nombre:       String(partRows[pi][1]).trim(),
      tarifa:       tarifa,
      tieneFactura: t === "sí" || t === "si"
    };
  }

  // Sumar horas a pagar por participante y quincena
  var horas = {"1":{}, "2":{}};
  var asist = hojaA.getDataRange().getValues();
  for (var ai=1; ai<asist.length; ai++) {
    var nombre = String(asist[ai][0]).trim();
    var id     = String(asist[ai][1]).trim();
    var tipo   = String(asist[ai][3]).trim();
    var ts     = new Date(asist[ai][2]);
    var hap    = parseFloat(asist[ai][8]) || 0;
    if (!nombre) continue;
    if (tipo === "Día de Estudio" || tipo === "Terapia") continue;
    if (isNaN(ts) || ts.getMonth()+1 !== mes || ts.getFullYear() !== anio) continue;
    var q = ts.getDate() <= 15 ? "1" : "2";
    var k = id || nombre;
    horas[q][k] = Math.round(((horas[q][k]||0) + hap)*100)/100;
  }

  var fact = hojaF.getDataRange().getValues();
  var creados=0, actualizados=0;

  ["1","2"].forEach(function(q) {
    partRows.slice(1).forEach(function(p) {
      var id     = String(p[0]).trim();
      var nombre = String(p[1]).trim();
      if (!id || !nombre) return;
      var k     = id || nombre;
      var hTrab = Math.round((horas[q][k]||0)*100)/100;

      var info       = partMap[id] || {};
      var tarifa     = info.tarifa || CFG.CATEGORIAS.C;
      var ivaStr     = info.tieneFactura ? "Sí" : "No";

      // Buscar fila existente — preservar Horas_A_Reponer (col G = idx 6)
      var filaE=-1, hReponer=0;
      for (var j=1; j<fact.length; j++) {
        if (String(fact[j][0])===id && fact[j][2]===nombreMes &&
            Number(fact[j][3])===anio && String(fact[j][4])===q) {
          filaE    = j+1;
          hReponer = parseFloat(fact[j][6]) || 0;
          break;
        }
      }

      var hPagar = Math.round((hTrab + hReponer)*100)/100;
      var base   = Math.round(hPagar * tarifa * 100)/100;
      var iva    = info.tieneFactura ? Math.round(base * CFG.IVA_PCT * 100)/100 : 0;
      var total  = Math.round((base + iva)*100)/100;
      var neto   = base;

      if (filaE > 0) {
        // Actualizar cols F–N (1-indexed 6–14): hTrab|hReponer|hPagar|tarifa|base|tieneIVA|iva|total|neto
        hojaF.getRange(filaE, 6, 1, 9).setValues([[
          hTrab, hReponer, hPagar, tarifa, base, ivaStr, iva, total, neto
        ]]);
        actualizados++;
      } else {
        hojaF.appendRow([
          id, nombre, nombreMes, anio, q,
          hTrab, 0, hTrab, tarifa, base, ivaStr, iva, total, neto,
          "No","","No","No","","",""
        ]);
        creados++;
      }
    });
  });

  actualizarDashboard();
  _alert(
    "✅ Facturación calculada — " + nombreMes + " " + anio + "\n" +
    "• " + creados + " registros nuevos\n• " + actualizados + " actualizados\n\n" +
    "Tarifas: A=Q16.50 | B=Q15.75 | C=Q15.00 | D=Q14.00\n" +
    "IVA 5%: solo quien tiene Tiene_Factura=Sí en PARTICIPANTES\n" +
    "Horas_A_Reponer: edita col G en FACTURACION y recalcula."
  );
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
  for (var i=1; i<datos.length; i++) {
    var f = datos[i];
    if (f[2]!==nombreMes || Number(f[3])!==anio) continue;
    if (!f[12] || parseFloat(f[12])===0) continue; // f[12] = Total_Factura (idx 12)
    var urlActual = String(f[20]||"");              // f[20] = URL_Recibo (idx 20)
    var doc = _crearOActualizarRecibo(f, carpeta, urlActual);
    var urlNueva = doc.getUrl();
    if (urlNueva !== urlActual) hojaF.getRange(i+1, 21).setValue(urlNueva);
    generados++;
  }
  _alert("✅ " + generados + " recibos generados/actualizados en Drive.");
}); }

function _crearOActualizarRecibo(f, carpeta, urlExistente) {
  // Índices FACTURACION:
  // 0=ID 1=Nombre 2=Mes 3=Anio 4=Quincena
  // 5=HorasTrab 6=HorasReponer 7=HorasPagar
  // 8=Tarifa 9=Base 10=TieneIVA 11=IVA 12=Total 13=Neto
  // 14=FactEnt 15=NumFact 16=Declaraguate 17=Pagado 18=FechaPago
  var id     = String(f[0]||"").trim();
  var nombre = String(f[1]||"").trim();
  var mes = f[2], anio = f[3], q = f[4];
  var hTrab   = parseFloat(f[5])||0;
  var hReponer= parseFloat(f[6])||0;
  var hPagar  = parseFloat(f[7])||0;
  var tarifa  = parseFloat(f[8])||0;
  var base    = parseFloat(f[9])||0;
  var tieneIVA= String(f[10]||"No");
  var iva     = parseFloat(f[11])||0;
  var total   = parseFloat(f[12])||0;
  var neto    = parseFloat(f[13])||0;

  var titulo = "Recibo_"+(id||nombre.replace(/\s/g,"_"))+"_Q"+q+"_"+mes+"_"+anio;
  var doc=null;
  if (urlExistente && urlExistente.startsWith("http")) {
    var m=urlExistente.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if(m){try{doc=DocumentApp.openById(m[1]);}catch(_){}}
    if(doc) doc.setName(titulo); else doc=null;
  }
  if (!doc) { doc=DocumentApp.create(titulo); DriveApp.getFileById(doc.getId()).moveTo(carpeta); }

  var body=doc.getBody();
  body.clear();
  body.setMarginTop(54).setMarginBottom(54).setMarginLeft(72).setMarginRight(72);
  var tz=Session.getScriptTimeZone();
  var hoy=Utilities.formatDate(new Date(),tz,"dd/MM/yyyy");

  body.appendParagraph(CFG.ORG + " — Recibo de Pago")
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .editAsText().setForegroundColor("#1a237e");
  body.appendParagraph("Período: "+String(mes)+" "+String(anio)+" — Quincena "+String(q))
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .editAsText().setFontSize(11).setForegroundColor("#555555");
  body.appendParagraph("");

  var t1=body.appendTable([
    ["Participante", nombre||"—"],
    ["Creamos ID",   id||"—"],
  ]);
  t1.getRow(0).getCell(0).setBackgroundColor("#1a237e").editAsText().setForegroundColor("#fff").setBold(true);
  t1.getRow(1).getCell(0).setBackgroundColor("#1a237e").editAsText().setForegroundColor("#fff").setBold(true);

  body.appendParagraph("");

  // Desglose de pago — condicional según si tiene IVA
  var filasDesglose = [
    ["Concepto","Detalle","Monto (Q)"],
    ["Horas trabajadas (Kobo)",  _n2(hTrab)+" hrs",                   ""],
  ];
  if (hReponer > 0) {
    filasDesglose.push(["Horas a reponer", _n2(hReponer)+" hrs",        ""]);
    filasDesglose.push(["Total horas a pagar", _n2(hPagar)+" hrs",      ""]);
  }
  filasDesglose.push(["Tarifa por hora",         "Q "+tarifa.toFixed(2)+"/hr",    ""]);
  filasDesglose.push(["Monto base",               "Horas × tarifa",               base.toFixed(2)]);
  if (tieneIVA === "Sí") {
    filasDesglose.push(["IVA 5% (Pcv.)",          "Se declara en Declaraguate",   iva.toFixed(2)]);
    filasDesglose.push(["Total Factura",           "Lo que paga la organización",  total.toFixed(2)]);
    filasDesglose.push(["Monto neto participante", "Total − IVA (queda con usted)",neto.toFixed(2)]);
  } else {
    filasDesglose.push(["Total a pagar",           "Sin IVA (no factura)",         total.toFixed(2)]);
  }
  var t2=body.appendTable(filasDesglose);
  _estilTablaEnc(t2, "#1a237e");
  // Resaltar total
  var filaTotal = tieneIVA==="Sí" ? filasDesglose.length-2 : filasDesglose.length-1;
  t2.getRow(filaTotal).editAsText().setBold(true).setFontSize(12);
  t2.getRow(filaTotal).getCell(2).setBackgroundColor("#e8eaf6");

  body.appendParagraph("");

  var fp = f[18] instanceof Date ? Utilities.formatDate(f[18],tz,"dd/MM/yyyy") : String(f[18]||"—");
  var t3=body.appendTable([
    ["Factura entregada",  String(f[14]||"No")],
    ["Número de factura",  String(f[15]||"—")],
    ["Declaraguate",       String(f[16]||"No")],
    ["Pagado",             String(f[17]||"No")],
    ["Fecha de pago",      fp],
  ]);
  for(var r=0;r<t3.getNumRows();r++) t3.getRow(r).getCell(0).editAsText().setBold(true);

  body.appendParagraph("");
  body.appendParagraph("Emisión: "+hoy)
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
  var titulo    = "Reporte_"+nombreMes+"_"+anio;
  var prev=carpeta.getFilesByName(titulo); while(prev.hasNext()) prev.next().setTrashed(true);

  var doc=DocumentApp.create(titulo);
  var body=doc.getBody();
  var tz=Session.getScriptTimeZone();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);

  body.appendParagraph(CFG.ORG+" — Reporte RRHH "+nombreMes+" "+anio)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .editAsText().setForegroundColor("#1a237e");
  body.appendParagraph("Tarifas A=Q16.50 | B=Q15.75 | C=Q15.00 | D=Q14.00  ·  IVA 5% solo Pequeño Contribuyente")
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .editAsText().setFontSize(10).setItalic(true).setForegroundColor("#555555");
  body.appendParagraph("");

  // Agrupar por participante (índices nuevos)
  var agrup={};
  datos.slice(1).forEach(function(f){
    if(f[2]!==nombreMes || Number(f[3])!==anio || !String(f[1]).trim()) return;
    var n=String(f[1]).trim();
    if(!agrup[n]) agrup[n]={q1b:0,q1i:0,q1t:0,q1h:0,q2b:0,q2i:0,q2t:0,q2h:0,tarifa:0,pagado:"No"};
    if(String(f[4])==="1"){
      agrup[n].q1b=parseFloat(f[9])||0; agrup[n].q1i=parseFloat(f[11])||0;
      agrup[n].q1t=parseFloat(f[12])||0; agrup[n].q1h=parseFloat(f[7])||0;
      agrup[n].tarifa=parseFloat(f[8])||0;
    }
    if(String(f[4])==="2"){
      agrup[n].q2b=parseFloat(f[9])||0; agrup[n].q2i=parseFloat(f[11])||0;
      agrup[n].q2t=parseFloat(f[12])||0; agrup[n].q2h=parseFloat(f[7])||0;
      agrup[n].tarifa=agrup[n].tarifa||parseFloat(f[8])||0;
    }
    if(f[17]==="Sí") agrup[n].pagado="Sí"; // col R = Pagado (idx 17)
  });

  var filas=[["Participante","Tarifa","Hrs Q1","Base Q1","IVA Q1","Total Q1","Hrs Q2","Base Q2","IVA Q2","Total Q2","TOTAL","Pagado"]];
  var sumBase=0,sumIVA=0,sumTot=0,pend=0;
  Object.keys(agrup).sort().forEach(function(n){
    var a=agrup[n];
    var totPart=a.q1t+a.q2t;
    sumBase+=a.q1b+a.q2b; sumIVA+=a.q1i+a.q2i; sumTot+=totPart;
    if(a.pagado!=="Sí") pend++;
    filas.push([n,
      "Q"+(a.tarifa||0).toFixed(2),
      _n2(a.q1h), "Q"+a.q1b.toFixed(2), a.q1i>0?"Q"+a.q1i.toFixed(2):"—", "Q"+a.q1t.toFixed(2),
      _n2(a.q2h), "Q"+a.q2b.toFixed(2), a.q2i>0?"Q"+a.q2i.toFixed(2):"—", "Q"+a.q2t.toFixed(2),
      "Q"+totPart.toFixed(2), a.pagado==="Sí"?"✓ Pagado":"Pendiente"
    ]);
  });
  _estilTablaEnc(body.appendTable(filas), "#1f54a8");

  body.appendParagraph("");
  body.appendTable([
    ["Total Monto Base",         "Q "+sumBase.toFixed(2)],
    ["Total IVA 5%",             "Q "+sumIVA.toFixed(2)],
    ["TOTAL A PAGAR (facturas)", "Q "+sumTot.toFixed(2)],
    ["Pagos pendientes",         pend+" participante(s)"],
    ["Generado",                 Utilities.formatDate(new Date(),tz,"dd/MM/yyyy HH:mm")],
  ]).editAsText().setFontSize(10);

  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(carpeta);
  _alert("✅ Reporte generado:\n"+doc.getUrl());
}); }

// ── Drive ─────────────────────────────────────────────────────

function crearEstructuraDrive() { _run(function() {
  var raiz     = _getOCreate(null, CFG.ORG+" · RRHH");
  var docsPD   = _getOCreate(raiz, "Docs_Proceso");
  var recibos  = _getOCreate(raiz, "Recibos");
  var reportes = _getOCreate(raiz, "Reportes");
  var p=PropertiesService.getScriptProperties();
  p.setProperty("RRHH_RAIZ",    raiz.getId());
  p.setProperty("RRHH_DOCS_DP", docsPD.getId());
  p.setProperty("RRHH_RECIBOS", recibos.getId());
  p.setProperty("RRHH_REPORTES",reportes.getId());
  _alert("✅ Drive listo:\n"+raiz.getUrl()+"\n\n• Docs_Proceso/\n• Recibos/\n• Reportes/");
}); }

function _carpetaDP() {
  var p=PropertiesService.getScriptProperties(),id=p.getProperty("RRHH_DOCS_DP"),r;
  if(id){try{r=DriveApp.getFolderById(id);}catch(_){}}
  if(!r){crearEstructuraDrive();r=DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty("RRHH_DOCS_DP"));}
  return r;
}
function _carpetaRecibos(anio, mes) {
  var p=PropertiesService.getScriptProperties(),id=p.getProperty("RRHH_RECIBOS"),r;
  if(id){try{r=DriveApp.getFolderById(id);}catch(_){}}
  if(!r){crearEstructuraDrive();r=DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty("RRHH_RECIBOS"));}
  return _getOCreate(_getOCreate(r,String(anio)),mes);
}
function _carpetaReportes(anio) {
  var p=PropertiesService.getScriptProperties(),id=p.getProperty("RRHH_REPORTES"),r;
  if(id){try{r=DriveApp.getFolderById(id);}catch(_){}}
  if(!r){crearEstructuraDrive();r=DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty("RRHH_REPORTES"));}
  return _getOCreate(r,String(anio));
}
function _getOCreate(padre, nombre) {
  var it=padre?padre.getFoldersByName(nombre):DriveApp.getFoldersByName(nombre);
  if(it.hasNext())return it.next();
  return padre?padre.createFolder(nombre):DriveApp.createFolder(nombre);
}

// ── Triggers ──────────────────────────────────────────────────

function configurarTriggers() { _run(function() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function(t){ ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("onEdit").forSpreadsheet(ss).onEdit().create();
  _alert("✅ Trigger onEdit activado.\n\n" +
    "Automatizaciones activas:\n" +
    "• Cambiar Categoría (col K) → auto-llena Tarifa_Hora (col L)\n" +
    "• Marcar Pagado (col R) → actualiza Dashboard");
}); }

// ── Reinstalar ────────────────────────────────────────────────

function reinstalarSistema() { _run(function() {
  var ui=SpreadsheetApp.getUi();
  var resp=ui.alert("⚠️  Reinstalar sistema RRHH — BORRADO COMPLETO",
    "Elimina TODO:\n• Hojas PARTICIPANTES, CLASIFICACION, ASISTENCIA, FACTURACION,\n" +
    "  DASHBOARD, DatosKobo, NombresCanonicos, DiasEstudio, ListaTerapias\n" +
    "• Carpeta «"+CFG.ORG+" · RRHH» con Docs de Proceso, Recibos y Reportes\n\n" +
    "Esta acción NO se puede deshacer.\n¿Continuar?", ui.ButtonSet.YES_NO);
  if(resp!==ui.Button.YES)return;
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  [CFG.HOJAS.PARTICIPANTES, CFG.HOJAS.CLASIFICACION, CFG.HOJAS.ASISTENCIA,
   CFG.HOJAS.FACTURACION, CFG.HOJAS.DASHBOARD, CFG.HOJAS.DATOS_KOBO,
   "NombresCanonicos","DiasEstudio","ListaTerapias"].forEach(function(n){
    var h=ss.getSheetByName(n);if(h)ss.deleteSheet(h);
  });
  var p=PropertiesService.getScriptProperties();
  var idRaiz=p.getProperty("RRHH_RAIZ");
  if(idRaiz){try{_borrarCarpetaRecursivo(DriveApp.getFolderById(idRaiz));}catch(_){}}
  else{var it=DriveApp.getFoldersByName(CFG.ORG+" · RRHH");while(it.hasNext())_borrarCarpetaRecursivo(it.next());}
  ScriptApp.getProjectTriggers().forEach(function(t){ScriptApp.deleteTrigger(t);});
  p.deleteAllProperties();
  Utilities.sleep(500);
  crearHojas();
  _alert("✅ Sistema reinstalado.\nSiguiente paso: PASO 2 — Crear estructura en Drive.");
}); }

function _borrarCarpetaRecursivo(carpeta) {
  var a=carpeta.getFiles();while(a.hasNext())a.next().setTrashed(true);
  var s=carpeta.getFolders();while(s.hasNext())_borrarCarpetaRecursivo(s.next());
  carpeta.setTrashed(true);
}

// ── Notificaciones ────────────────────────────────────────────

function enviarRecordatorioPagos() { _run(function() {
  var hojaF=_sh(CFG.HOJAS.FACTURACION), hojaP=_sh(CFG.HOJAS.PARTICIPANTES);
  var datos=hojaF.getDataRange().getValues();
  var mes=CFG.MESES[new Date().getMonth()], anio=new Date().getFullYear();
  var corr=_mapaCorreos(hojaP), n=0;
  for(var i=1; i<datos.length; i++){
    var f=datos[i];
    // idx 17=Pagado  idx 12=Total_Factura  idx 9=Base  idx 11=IVA  idx 10=TieneIVA
    if(f[2]!==mes || Number(f[3])!==anio || f[17]==="Sí" || !f[12] || parseFloat(f[12])===0) continue;
    var c=corr[String(f[0]).trim()]; if(!c) continue;
    var ivaLinea = f[10]==="Sí"
      ? " (base Q "+parseFloat(f[9]).toFixed(2)+" + IVA Q "+parseFloat(f[11]).toFixed(2)+")"
      : "";
    MailApp.sendEmail({to:c,
      subject:"["+CFG.ORG+"] Pago pendiente — "+mes+" "+anio,
      body:"Hola "+f[1]+",\n\nTienes pago pendiente:\n"+
           "Quincena "+f[4]+" — Total: Q "+parseFloat(f[12]).toFixed(2)+ivaLinea+"\n\n"+
           "Tarifa: Q "+parseFloat(f[8]).toFixed(2)+"/hr"+
           (f[10]==="Sí"?"\nEntrega tu factura para procesar el pago.\n":"\n")+
           "\nSaludos,\n"+CFG.ORG
    });
    n++;
  }
  _alert("✅ "+n+" recordatorios enviados.");
}); }

function enviarResumenMensual() { _run(function() {
  var hojaF=_sh(CFG.HOJAS.FACTURACION), datos=hojaF.getDataRange().getValues();
  var mes=CFG.MESES[new Date().getMonth()], anio=new Date().getFullYear();
  var sumBase=0,sumIVA=0,sumTot=0,pag=0,pend=0,lineas=[];
  for(var i=1; i<datos.length; i++){
    var f=datos[i]; if(f[2]!==mes||Number(f[3])!==anio) continue;
    var base=parseFloat(f[9])||0, iva=parseFloat(f[11])||0, tot=parseFloat(f[12])||0;
    sumBase+=base; sumIVA+=iva; sumTot+=tot;
    if(f[17]==="Sí") pag++; else pend++;
    var ivaStr = f[10]==="Sí" ? " IVA:Q"+iva.toFixed(2) : "";
    lineas.push("  "+f[1]+" Q"+f[4]+" — Base:Q"+base.toFixed(2)+ivaStr+" Total:Q"+tot.toFixed(2)+
                " — "+(f[17]==="Sí"?"PAGADO":"PENDIENTE"));
  }
  MailApp.sendEmail({to:CFG.CORREO_ADMIN,
    subject:"["+CFG.ORG+"] Resumen RRHH — "+mes+" "+anio,
    body:"Resumen "+mes+" "+anio+"\n══════════════════\n\n"+lineas.join("\n")+
         "\n\n══════════════════\n"+
         "Monto base total:  Q "+sumBase.toFixed(2)+"\n"+
         "IVA total (5%):    Q "+sumIVA.toFixed(2)+"\n"+
         "Total a pagar:     Q "+sumTot.toFixed(2)+"\n"+
         "Pagados: "+pag+"  |  Pendientes: "+pend
  });
  _alert("✅ Resumen enviado a "+CFG.CORREO_ADMIN);
}); }

function _mapaCorreos(hojaP) {
  var m={}, d=hojaP.getDataRange().getValues();
  for(var i=1; i<d.length; i++){
    var id=String(d[i][0]).trim(), c=d[i][15]; // col P = Correo (idx 15)
    if(id&&c) m[id]=c;
  }
  return m;
}

// ── Dashboard ─────────────────────────────────────────────────

function actualizarDashboard() { _run(function() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var dash=ss.getSheetByName(CFG.HOJAS.DASHBOARD)||ss.insertSheet(CFG.HOJAS.DASHBOARD,0);
  var mes=CFG.MESES[new Date().getMonth()], anio=new Date().getFullYear();
  var hojaP=ss.getSheetByName(CFG.HOJAS.PARTICIPANTES);
  var hojaF=ss.getSheetByName(CFG.HOJAS.FACTURACION);
  var activos=0, sumBase=0, sumIVA=0, sumTot=0, pag=0, pend=0;
  if(hojaP){
    var dp=hojaP.getDataRange().getValues();
    for(var i=1;i<dp.length;i++){
      if(String(dp[i][5]).toLowerCase()==="activo") activos++; // col F = Estado (idx 5)
    }
  }
  if(hojaF){
    var df=hojaF.getDataRange().getValues();
    for(var j=1;j<df.length;j++){
      if(df[j][2]!==mes||Number(df[j][3])!==anio) continue;
      sumBase += parseFloat(df[j][9])||0;   // col J = Monto_Base (idx 9)
      sumIVA  += parseFloat(df[j][11])||0;  // col L = IVA (idx 11)
      sumTot  += parseFloat(df[j][12])||0;  // col M = Total_Factura (idx 12)
      if(df[j][17]==="Sí") pag++; else pend++; // col R = Pagado (idx 17)
    }
  }
  var ts=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"dd/MM/yyyy HH:mm");
  var filas=[
    ["MÉTRICA","VALOR"],
    ["Participantes activos",               activos],
    ["Monto base ("+mes+" "+anio+")",       "Q "+sumBase.toFixed(2)],
    ["IVA 5% del mes",                      "Q "+sumIVA.toFixed(2)],
    ["TOTAL A PAGAR (facturas)",            "Q "+sumTot.toFixed(2)],
    ["Quincenas pagadas",                   pag],
    ["Quincenas pendientes",                pend],
    ["",""],
    ["Tarifas vigentes",  "A=Q16.50 | B=Q15.75 | C=Q15.00 | D=Q14.00"],
    ["Actualizado",       ts],
  ];
  dash.clearContents();
  dash.getRange(1,1,filas.length,2).setValues(filas);
  dash.getRange(1,1,1,2).setBackground("#639922").setFontColor("#fff").setFontWeight("bold");
  dash.getRange(5,1,1,2).setBackground("#e8eaf6").setFontWeight("bold");
  if(pend>0) dash.getRange(7,2).setBackground("#fce8e6").setFontColor("#c62828").setFontWeight("bold");
  dash.autoResizeColumns(1,2);
}); }
