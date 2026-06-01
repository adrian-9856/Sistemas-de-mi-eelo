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
    PERIODOS:      "PERIODOS",
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

  // ── Submenú: Más reportes ─────────────────────────────────────
  var menuReportes = ui.createMenu("📊 Más reportes")
    .addItem("📊 Reporte por día",                     "generarReportePorDia")
    .addItem("📊 Reporte por semana",                  "generarReportePorSemana")
    .addItem("📊 Reporte por mes",                     "generarReportePorMes")
    .addItem("📊 Reporte completo (todos los datos)",  "generarReporteTodo");

  // ── Submenú: Facturación ──────────────────────────────────────
  var menuFact = ui.createMenu("💰 Facturación y recibos")
    .addItem("📅 Proceso mensual completo",             "procesarMesCompleto")
    .addSeparator()
    .addItem("💰 Calcular facturación del mes",         "calcularFacturacionMes")
    .addItem("📑 Resumen facturación en hoja",          "generarResumenFacturacionEnHoja")
    .addItem("🧾 Generar recibos de pago",              "generarRecibosMes")
    .addSeparator()
    .addItem("🔔 Recordatorio de pagos pendientes",     "enviarRecordatorioPagos")
    .addItem("📬 Resumen mensual al admin",             "enviarResumenMensual");

  // ── Submenú: Configuración avanzada ──────────────────────────
  var menuConfig = ui.createMenu("⚙️ Configuración avanzada")
    .addItem("📚 Configurar Días de Estudio",           "crearHojaDiasEstudio")
    .addItem("🧘 Configurar Lista de Terapias",         "crearHojaListaTerapias")
    .addSeparator()
    .addItem("👥 Directorio de participantes",          "generarDirectorioParticipantes")
    .addItem("📄 Generar DP (fila activa)",             "generarDpFilaActiva")
    .addItem("📄 Actualizar todos los DPs",             "actualizarTodosLosDps")
    .addSeparator()
    .addItem("✏️  Cambiar nombre de participante",      "cambiarNombreParticipante")
    .addItem("✨ Normalizar nombres y datos Kobo",       "normalizarTodo")
    .addItem("🔗 Emparejar entradas/salidas (manual)",  "emparejarAsistencia")
    .addItem("🔄 Reimportar todo desde Kobo",           "reimportarTodoDesdeKobo")
    .addSeparator()
    .addItem("🔍 Diagnosticar Datos Kobo",              "diagnosticarDatosKobo")
    .addItem("🔧 Reparar Datos Kobo",                   "repararDatosKobo")
    .addSeparator()
    .addItem("⚡ Activar automatizaciones",             "configurarTriggers")
    .addSeparator()
    .addItem("🗑️  Reinstalar sistema (borra TODO)",     "reinstalarSistema");

  ui.createMenu("👥 RRHH")
    // ── Configuración inicial ──
    .addItem("🚀 Instalación completa",                 "instalarTodo")
    .addSeparator()
    // ── Participantes ──
    .addItem("➕ Nuevo participante",                    "nuevoParticipante")
    .addItem("📋 Cargar lista oficial",                  "cargarListaParticipantes")
    .addSeparator()
    // ── Kobo ──
    .addItem("📥 Importar asistencia desde Kobo",        "importarDesdeKobo")
    .addSeparator()
    // ── QUINCENAS (flujo principal) ──
    .addItem("📅 Ver quincena actual",                   "verQuincenaActual")
    .addItem("✅ Cerrar quincena y crear siguiente",      "cerrarQuincenaYCrearSiguiente")
    .addItem("🗓️ Configurar nueva quincena",             "configurarNuevaQuincena")
    .addItem("🔍 Reporte de quincena pasada",            "generarReporteQuincenaPasada")
    .addSeparator()
    // ── Reportes ──
    .addItem("📊 Reporte por rango de fechas",           "generarReportePorRango")
    .addSubMenu(menuReportes)
    .addSeparator()
    // ── Finanzas ──
    .addSubMenu(menuFact)
    .addSeparator()
    // ── Dashboard ──
    .addItem("🔄 Actualizar Dashboard",                  "actualizarDashboard")
    .addSeparator()
    .addSubMenu(menuConfig)
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
      "Horas_Trabajadas","Es_Dia_Estudio","Es_Terapia","Porcentaje_Pago","Horas_A_Pagar","Clave_Dia",
      "Hora_Entrada","Hora_Salida"
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

  // PERIODOS — hoja de control de quincenas
  crearHojaPeriodos();

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
    "• FACTURACION (21 cols — incl. Horas_A_Reponer, Tarifa por participante)\n" +
    "• PERIODOS (control de quincenas)\n\n" +
    "Tarifas: A=Q16.50 | B=Q15.75 | C=Q15.00 | D=Q14.00\n" +
    "IVA 5%: solo participantes con Tiene_Factura=Sí"
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

// ── Lista oficial de participantes ────────────────────────────────

// Lista maestra: [Nombre, Categoria].  Editar aquí para actualizar.
var LISTA_OFICIAL = [
  ["ANGELICA VELIZ",               "A"],
  ["EMILY ZACARÍAS",               "B"],
  ["JUANA VICENTE",                "B"],
  ["KARIN BALCARCEL",              "B"],
  ["SARA RAYMUNDO",                "B"],
  ["SINDY LAZARO",                 "B"],
  ["SINDY VELIZ",                  "B"],
  ["ANAID MATEO",                  "C"],
  ["ERICKA VASQUEZ",               "C"],
  ["HEIDY LÁZARO",                 "C"],
  ["HELEN RODAS",                  "C"],
  ["LAURA GONZALEZ",               "C"],
  ["LETICIA SUMALÉ",               "C"],
  ["MARIA AUDELIA VELASQUEZ",      "C"],
  ["MARIA DEL CARMEN BORRAYO",     "C"],
  ["MAYRA LORENA CIFUENTES",       "C"],
  ["JEANNETTE SAQUIC",             "C"],
  ["SARAÍ PIVARAL",                "C"],
  ["SANDRA ARACELY VICENTE",       "C"],
  ["VILMA ELIZABETH LOPEZ",        "C"],
  ["ANGÉLICA MARIBEL CUXE",        "D"],
  ["BRENDA AZUCENA DEL CID URREA", "D"],
  ["ELENDY NICOLE PEDROZA",        "D"],
  ["ALICIA LÓPEZ REYNOSO",         "D"],
  ["ANA REBECA LARIOS PEREZ",      "D"],
  ["JEIMY SUCELI BARRIENTOS",      "D"],
  ["MARÍA AIDÉ ALVARADO CORTÉZ",   "D"],
  ["MIRNA LETICIA RODRIGUEZ",      "D"],
  ["OTILIA TURUY PAZ",             "D"]
];

/*
 * Carga la lista oficial en PARTICIPANTES.
 * Solo agrega filas nuevas; no sobreescribe existentes.
 */
function cargarListaParticipantes() { _run(function() {
  var hP = _sh(CFG.HOJAS.PARTICIPANTES);

  // índice de nombres ya existentes (normalizados)
  var existentes = {};
  var lastRow = hP.getLastRow();
  if (lastRow > 1) {
    hP.getRange(2, 2, lastRow - 1, 1).getValues().forEach(function(r) {
      if (r[0]) existentes[limpiarNombre(String(r[0]))] = true;
    });
  }

  var agregados = 0, omitidos = 0;
  LISTA_OFICIAL.forEach(function(item) {
    var nombre = item[0], cat = item[1];
    if (existentes[limpiarNombre(nombre)]) { omitidos++; return; }
    var tarifa = CFG.CATEGORIAS[cat] || 0;
    hP.appendRow([
      "",       // Creamos_ID
      nombre,   // Nombre
      CFG.ORG,  // Proyecto
      "","",    // Division, Programa
      "Activo", // Estado
      "","","","", // Etapa, Educacion, Apoyo, Inclusion
      cat,      // Categoria
      tarifa,   // Tarifa_Hora
      "No",     // Tiene_Factura
      "","","","","","",""  // DPI…URL_Doc_Proceso
    ]);
    existentes[limpiarNombre(nombre)] = true;
    agregados++;
  });

  _alert(
    "✅ Lista cargada en PARTICIPANTES.\n" +
    "  Nuevos: " + agregados + "\n" +
    "  Ya existían: " + omitidos
  );
}); }

/*
 * Genera (o actualiza) una hoja "Directorio" con todos los participantes
 * agrupados por categoría, con tarifa y estado.
 */
function generarDirectorioParticipantes() { _run(function() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var hP  = _sh(CFG.HOJAS.PARTICIPANTES);
  var lastRow = hP.getLastRow();
  if (lastRow < 2) { _alert("No hay participantes en PARTICIPANTES."); return; }

  var datos = hP.getRange(2, 1, lastRow - 1, 13).getValues();

  // Agrupar por categoría
  var porCat = { A: [], B: [], C: [], D: [] };
  datos.forEach(function(r) {
    var nombre  = String(r[1]).trim();
    var cat     = String(r[10]).trim().toUpperCase();
    var tarifa  = parseFloat(r[11]) || CFG.CATEGORIAS[cat] || 0;
    var estado  = String(r[5]).trim() || "Activo";
    if (!nombre || !porCat[cat]) return;
    porCat[cat].push({ nombre: nombre, tarifa: tarifa, estado: estado });
  });

  // Recrear hoja Directorio
  var HOJA = "Directorio";
  var hD = ss.getSheetByName(HOJA);
  if (hD) ss.deleteSheet(hD);
  hD = ss.insertSheet(HOJA);

  var filas = [], tipos = [];

  function push(fila, tipo) { filas.push(fila); tipos.push(tipo); }

  var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  push(["DIRECTORIO DE PARTICIPANTES — " + CFG.ORG, "", "", "", ""], "titulo");
  push(["Actualizado: " + ahora,                      "", "", "", ""], "sub");
  push(["", "", "", "", ""],                                           "vacio");

  // encabezado de columnas
  push(["#", "NOMBRE", "CATEGORÍA", "TARIFA Q/HR", "ESTADO"], "enc");

  var num = 1;
  var COLORES_CAT = { A: "#e6f4ea", B: "#e8f0fe", C: "#fef7e0", D: "#fce8e6" };
  var TITULO_CAT  = { A: "#34a853", B: "#4285f4", C: "#fbbc04", D: "#ea4335" };

  ["A","B","C","D"].forEach(function(cat) {
    var lista = (porCat[cat] || []).slice().sort(function(a,b){ return a.nombre.localeCompare(b.nombre,"es"); });
    if (!lista.length) return;
    var tarifa = CFG.CATEGORIAS[cat];

    push(["▶  CATEGORÍA " + cat, "Q" + tarifa.toFixed(2) + " / hora",
          lista.length + " participante" + (lista.length > 1 ? "s" : ""),
          "", ""], "cat_" + cat);

    lista.forEach(function(p) {
      push([num++, p.nombre, cat, "Q" + tarifa.toFixed(2), p.estado],
           p.estado === "Activo" ? "activo" : "inactivo");
    });

    push(["", "Subtotal categoría " + cat + ": " + lista.length, "", "", ""], "subtotal");
    push(["", "", "", "", ""], "vacio");
  });

  // Total general
  var totalActivos = Object.keys(porCat).reduce(function(s, c) {
    return s + porCat[c].filter(function(p){ return p.estado === "Activo"; }).length;
  }, 0);
  push(["", "TOTAL ACTIVOS: " + totalActivos, "", "", ""], "total");

  hD.getRange(1, 1, filas.length, 5).setValues(filas);

  // ── Formato ──────────────────────────────────────────────────
  tipos.forEach(function(tipo, idx) {
    var r = hD.getRange(idx + 1, 1, 1, 5);
    r.setFontFamily("Arial").setFontSize(10);

    if (tipo === "titulo") {
      r.merge().setFontSize(14).setFontWeight("bold")
       .setBackground("#1a73e8").setFontColor("#ffffff")
       .setHorizontalAlignment("center");

    } else if (tipo === "sub") {
      r.merge().setFontColor("#5f6368").setHorizontalAlignment("center")
       .setBackground("#f8f9fa");

    } else if (tipo === "enc") {
      r.setFontWeight("bold").setBackground("#202124").setFontColor("#ffffff")
       .setHorizontalAlignment("center");

    } else if (tipo.indexOf("cat_") === 0) {
      var c = tipo.split("_")[1];
      r.setFontWeight("bold").setFontSize(11)
       .setBackground(TITULO_CAT[c]).setFontColor("#ffffff");

    } else if (tipo === "activo") {
      r.setBackground("#ffffff");
      hD.getRange(idx + 1, 5, 1, 1).setFontColor("#137333").setFontWeight("bold");

    } else if (tipo === "inactivo") {
      r.setBackground("#f1f3f4").setFontColor("#9aa0a6");

    } else if (tipo === "subtotal") {
      hD.getRange(idx + 1, 2, 1, 1).setFontStyle("italic").setFontColor("#5f6368");

    } else if (tipo === "total") {
      r.setFontWeight("bold").setBackground("#e8f0fe");
    }
  });

  // bordes en filas de datos y encabezado
  tipos.forEach(function(tipo, idx) {
    if (tipo === "activo" || tipo === "inactivo" || tipo === "enc") {
      hD.getRange(idx + 1, 1, 1, 5)
        .setBorder(null, null, true, null, null, null, "#dadce0", SpreadsheetApp.BorderStyle.SOLID);
    }
  });

  // anchos de columna
  hD.setColumnWidth(1, 45);
  hD.setColumnWidth(2, 260);
  hD.setColumnWidth(3, 100);
  hD.setColumnWidth(4, 110);
  hD.setColumnWidth(5, 90);
  hD.setFrozenRows(4);

  // Activar la hoja
  ss.setActiveSheet(hD);
  _alert("✅ Directorio generado con " + (num - 1) + " participantes.");
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

// Llamado por el trigger instalable onOpen (tiene permisos completos)
function importarAlAbrir() { try { importarDesdeKobo(); } catch(_) {} }

// Reimportación completa — borra DatosKobo y lo reconstruye desde cero
function reimportarTodoDesdeKobo() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert("⚠️ REIMPORTACIÓN COMPLETA",
    "Esto BORRARÁ la hoja DatosKobo y la reimportará desde cero desde Kobo.\n\n" +
    "Usa esto solo si los datos están muy desordenados.\n\n¿Continuar?",
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  var resKobo = UrlFetchApp.fetch(CFG.KOBO_URL_CSV, { muteHttpExceptions: true });
  var code = resKobo.getResponseCode();
  if (code === 503) { _alert("⏳ Kobo ocupado (503). Espera 2 min e intenta de nuevo."); return; }
  if (code !== 200) throw new Error("Error Kobo HTTP " + code);

  var datos = Utilities.parseCsv(resKobo.getContentText(), ";");
  if (datos.length < 2) { _alert("Kobo no devolvió registros."); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hOld = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (hOld) ss.deleteSheet(hOld);

  var hoja = ss.insertSheet(CFG.HOJAS.DATOS_KOBO);
  hoja.getRange(1,1,datos.length,datos[0].length).setValues(datos);
  hoja.getRange(1,1,1,datos[0].length).setFontWeight("bold").setBackground("#4a86e8").setFontColor("#fff");
  hoja.setFrozenRows(1);
  _limpiarColumnasKobo(hoja, datos[0]);
  _normalizarAccionSilencioso(hoja);
  _alert("✅ Reimportación completa: " + (datos.length-1) + " registros.\n\nRecuerda ejecutar 'Emparejar entradas/salidas' para recalcular ASISTENCIA.");
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
    filasAsist.push([g.nombre, id, g.ent[0], tipo2, horas, esDiaEst, esTer, pct, hap, clave, g.ent[0], g.sal[g.sal.length-1]]);
  });

  filasAsist.sort(function(a,b){ return new Date(b[2])-new Date(a[2]); });
  var hA = _sh(CFG.HOJAS.ASISTENCIA);
  if (hA.getLastRow() > 1) hA.deleteRows(2, hA.getLastRow()-1);
  if (filasAsist.length > 0) {
    hA.getRange(2,1,filasAsist.length,12).setValues(filasAsist);
    hA.getRange("C2:C"+(filasAsist.length+1)).setNumberFormat("dd/MM/yyyy");
    hA.getRange("K2:L"+(filasAsist.length+1)).setNumberFormat("HH:mm");
  }
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

/*
 * Escribe el resumen de facturación del mes en una hoja del Spreadsheet.
 * Agrupa por categoría, muestra Q1+Q2, subtotales, total general.
 * Nombre de hoja: "Fact_Enero_2025" etc.
 */
function generarResumenFacturacionEnHoja(mes, anio) { _run(function() {
  var ahora     = new Date();
  mes  = mes  || ahora.getMonth()+1;
  anio = anio || ahora.getFullYear();
  var nombreMes = CFG.MESES[mes-1];

  var hojaF = _sh(CFG.HOJAS.FACTURACION);
  var hojaP = _sh(CFG.HOJAS.PARTICIPANTES);
  var datos  = hojaF.getDataRange().getValues();

  // Mapa categoría por nombre
  var catMap = {};
  hojaP.getDataRange().getValues().slice(1).forEach(function(r){
    var n = String(r[1]).trim();
    if (n) catMap[n] = String(r[10]).trim().toUpperCase() || "?";
  });

  // Agrupar por nombre → {cat, tarifa, q1{h,base,iva,total,pagado}, q2{...}}
  var personas = {};
  datos.slice(1).forEach(function(f) {
    if (f[2] !== nombreMes || Number(f[3]) !== anio) return;
    var nombre = String(f[1]).trim();
    if (!nombre) return;
    if (!personas[nombre]) {
      personas[nombre] = {
        cat: catMap[nombre] || "?",
        tarifa: parseFloat(f[8]) || 0,
        q1: { h:0, base:0, iva:0, total:0, pagado:"No", urlRecibo:"" },
        q2: { h:0, base:0, iva:0, total:0, pagado:"No", urlRecibo:"" }
      };
    }
    var q = String(f[4]);
    if (q === "1" || q === "2") {
      var qd = personas[nombre]["q"+q];
      qd.h       = parseFloat(f[7])  || 0;
      qd.base    = parseFloat(f[9])  || 0;
      qd.iva     = parseFloat(f[11]) || 0;
      qd.total   = parseFloat(f[12]) || 0;
      qd.pagado  = String(f[17]).trim() || "No";
      qd.urlRecibo = String(f[20] || "");
    }
  });

  if (!Object.keys(personas).length) {
    _alert("⚠️ No hay datos de facturación para " + nombreMes + " " + anio + ".\nEjecuta primero 'Calcular facturación del mes'."); return;
  }

  // Crear/reemplazar hoja
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var titulo = "Fact_" + nombreMes + "_" + anio;
  var hR = ss.getSheetByName(titulo);
  if (hR) ss.deleteSheet(hR);
  hR = ss.insertSheet(titulo);

  var filas = [], tipos = [];
  function push(f, t) { filas.push(f); tipos.push(t); }

  var cols = ["#","NOMBRE","CAT","TARIFA","HRS Q1","BASE Q1","IVA Q1","TOTAL Q1","PAG Q1",
              "HRS Q2","BASE Q2","IVA Q2","TOTAL Q2","PAG Q2","TOTAL MES","ESTADO"];

  push([CFG.ORG + " — Facturación " + nombreMes + " " + anio,
        "","","","","","","","","","","","","","",""], "titulo");
  push(["Generado: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
        "","","","","","","","","","","","","","",""], "sub");
  push(cols, "enc");

  var COLORES = { A:"#e6f4ea", B:"#e8f0fe", C:"#fef7e0", D:"#fce8e6", "?":"#f1f3f4" };
  var TITULO_CAT = { A:"#34a853", B:"#4285f4", C:"#fbbc04", D:"#ea4335", "?":"#9aa0a6" };
  var num = 1;
  var totalBase=0, totalIVA=0, totalMes=0, pendientes=0;

  ["A","B","C","D","?"].forEach(function(cat) {
    var grupo = Object.keys(personas).filter(function(n){ return personas[n].cat === cat; });
    if (!grupo.length) return;
    grupo.sort(function(a,b){ return a.localeCompare(b,"es"); });

    var tarifa = CFG.CATEGORIAS[cat] || 0;
    push(["","▶  CATEGORÍA "+cat+(tarifa?" — Q"+tarifa.toFixed(2)+"/hr":""),
          "","","","","","","","","","","","","",""], "cat_"+cat);

    var catBase=0, catIVA=0, catTotal=0;
    grupo.forEach(function(nombre) {
      var p  = personas[nombre];
      var q1 = p.q1, q2 = p.q2;
      var totMes = q1.total + q2.total;
      var pagado = (q1.pagado==="Sí" && q2.pagado==="Sí") ? "✓ Pagado" :
                   (q1.pagado==="Sí" || q2.pagado==="Sí") ? "Parcial"  : "Pendiente";
      catBase  += q1.base  + q2.base;
      catIVA   += q1.iva   + q2.iva;
      catTotal += totMes;
      if (pagado !== "✓ Pagado") pendientes++;

      push([num++, nombre, cat, "Q"+(p.tarifa||tarifa).toFixed(2),
            q1.h||"", q1.base?"Q"+q1.base.toFixed(2):"", q1.iva?"Q"+q1.iva.toFixed(2):"—",
            q1.total?"Q"+q1.total.toFixed(2):"", q1.pagado,
            q2.h||"", q2.base?"Q"+q2.base.toFixed(2):"", q2.iva?"Q"+q2.iva.toFixed(2):"—",
            q2.total?"Q"+q2.total.toFixed(2):"", q2.pagado,
            "Q"+totMes.toFixed(2), pagado],
           pagado==="✓ Pagado" ? "pagado" : pagado==="Parcial" ? "parcial" : "pendiente");
    });

    totalBase+=catBase; totalIVA+=catIVA; totalMes+=catTotal;
    push(["","Subtotal Categoría "+cat,"","",
          "","Q"+catBase.toFixed(2),"Q"+catIVA.toFixed(2),"","",
          "","","","","","Q"+catTotal.toFixed(2),""], "subtotal_"+cat);
    push(Array(16).fill(""), "vacio");
  });

  // Total general
  push(["","TOTAL GENERAL","","",
        "","Q"+totalBase.toFixed(2),"Q"+totalIVA.toFixed(2),"","",
        "","","","","","Q"+totalMes.toFixed(2),
        pendientes+" pendiente"+(pendientes!==1?"s":"")], "total");

  hR.getRange(1, 1, filas.length, 16).setValues(filas);

  // ── Formato ─────────────────────────────────────────────────
  tipos.forEach(function(tipo, idx) {
    var r = hR.getRange(idx+1, 1, 1, 16);
    r.setFontFamily("Arial").setFontSize(10);
    if (tipo === "titulo") {
      r.merge().setFontSize(13).setFontWeight("bold")
       .setBackground("#1a73e8").setFontColor("#ffffff").setHorizontalAlignment("center");
    } else if (tipo === "sub") {
      r.merge().setFontColor("#5f6368").setBackground("#f8f9fa").setHorizontalAlignment("center");
    } else if (tipo === "enc") {
      r.setFontWeight("bold").setBackground("#202124").setFontColor("#ffffff").setHorizontalAlignment("center");
    } else if (tipo.indexOf("cat_") === 0) {
      var c = tipo.split("_")[1];
      r.setFontWeight("bold").setBackground(TITULO_CAT[c]||"#9aa0a6").setFontColor("#ffffff");
    } else if (tipo === "pagado") {
      r.setBackground("#e6f4ea");
      hR.getRange(idx+1,16,1,1).setFontColor("#137333").setFontWeight("bold");
    } else if (tipo === "parcial") {
      r.setBackground("#fef7e0");
      hR.getRange(idx+1,16,1,1).setFontColor("#b06000").setFontWeight("bold");
    } else if (tipo === "pendiente") {
      r.setBackground("#fff8f7");
      hR.getRange(idx+1,16,1,1).setFontColor("#c5221f").setFontWeight("bold");
    } else if (tipo.indexOf("subtotal") === 0) {
      var c2 = tipo.split("_")[1];
      r.setFontWeight("bold").setBackground(COLORES[c2]||"#f1f3f4").setFontStyle("italic");
    } else if (tipo === "total") {
      r.setFontWeight("bold").setFontSize(11).setBackground("#e8f0fe");
    }
  });
  tipos.forEach(function(tipo, idx) {
    if (tipo==="pagado"||tipo==="parcial"||tipo==="pendiente"||tipo==="enc") {
      hR.getRange(idx+1,1,1,16).setBorder(null,null,true,null,null,null,"#dadce0",SpreadsheetApp.BorderStyle.SOLID);
    }
  });

  // Anchos
  [40,200,55,80,65,85,75,85,75,65,85,75,85,75,95,90].forEach(function(w,i){ hR.setColumnWidth(i+1,w); });
  hR.setFrozenRows(3);
  ss.setActiveSheet(hR);

  _alert("✅ Resumen de facturación generado en hoja '" + titulo + "'\n\nTotal del mes: Q" + totalMes.toFixed(2) + "\nPagos pendientes: " + pendientes);
}); }

/*
 * Proceso mensual completo en un solo clic:
 *   1. Importar Kobo
 *   2. Emparejar entradas/salidas
 *   3. Calcular facturación del mes
 *   4. Generar recibos de pago
 *   5. Generar resumen en hoja
 */
function procesarMesCompleto() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    "Proceso mensual completo",
    "Esto ejecutará en orden:\n\n" +
    "1️⃣  Importar datos desde Kobo\n" +
    "2️⃣  Emparejar entradas/salidas\n" +
    "3️⃣  Calcular facturación del mes actual\n" +
    "4️⃣  Generar recibos de pago\n" +
    "5️⃣  Generar resumen de facturación en hoja\n\n" +
    "¿Continuar?",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;

  var log = [];
  var ahora = new Date();
  var mes  = ahora.getMonth()+1;
  var anio = ahora.getFullYear();
  var nombreMes = CFG.MESES[mes-1];

  // Paso 1: Importar Kobo
  try { importarDesdeKobo(); log.push("✅ 1️⃣  Kobo importado"); }
  catch(e) { log.push("⚠️ 1️⃣  Kobo: " + e.message); }

  // Paso 2: Emparejar
  try { emparejarAsistencia(); log.push("✅ 2️⃣  Asistencia emparejada"); }
  catch(e) { log.push("⚠️ 2️⃣  Emparejar: " + e.message); }

  // Paso 3: Calcular facturación
  try { _calcular(mes, anio); log.push("✅ 3️⃣  Facturación calculada — " + nombreMes + " " + anio); }
  catch(e) { log.push("⚠️ 3️⃣  Facturación: " + e.message); }

  // Paso 4: Recibos
  try {
    var hojaF = _sh(CFG.HOJAS.FACTURACION);
    var datos = hojaF.getDataRange().getValues();
    var carpeta = _carpetaRecibos(anio, nombreMes);
    var generados = 0;
    for (var i=1; i<datos.length; i++) {
      var f = datos[i];
      if (f[2] !== nombreMes || Number(f[3]) !== anio) continue;
      if (!f[12] || parseFloat(f[12]) === 0) continue;
      var urlActual = String(f[20]||"");
      var doc = _crearOActualizarRecibo(f, carpeta, urlActual);
      var urlNueva = doc.getUrl();
      if (urlNueva !== urlActual) hojaF.getRange(i+1, 21).setValue(urlNueva);
      generados++;
    }
    log.push("✅ 4️⃣  " + generados + " recibos generados");
  } catch(e) { log.push("⚠️ 4️⃣  Recibos: " + e.message); }

  // Paso 5: Resumen en hoja
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var titulo = "Fact_" + nombreMes + "_" + anio;
    // Llamar directamente la lógica (sin _run para no anidar)
    generarResumenFacturacionEnHoja(mes, anio);
    log.push("✅ 5️⃣  Resumen generado en hoja '" + titulo + "'");
  } catch(e) { log.push("⚠️ 5️⃣  Resumen: " + e.message); }

  _alert("Proceso mensual — " + nombreMes + " " + anio + "\n\n" + log.join("\n"));
}); }

// ══════════════════════════════════════════════════════════════════
// SISTEMA DE QUINCENAS
// ──────────────────────────────────────────────────────────────────
// Flujo:
//   1. configurarNuevaQuincena() → el usuario define fecha inicio/fin
//   2. El reporte se auto-actualiza diariamente (trigger)
//   3. cerrarQuincenaYCrearSiguiente() → cierra la actual, abre la próxima
//
// Hoja PERIODOS (cols A–H):
//   A: ID | B: Label | C: Fecha_Inicio | D: Fecha_Fin
//   E: Estado (Activo/Cerrado) | F: Total_Q | G: Tab_Reporte | H: Fecha_Cierre
// ══════════════════════════════════════════════════════════════════

function crearHojaPeriodos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var h  = ss.getSheetByName(CFG.HOJAS.PERIODOS);
  if (h) return h;
  h = ss.insertSheet(CFG.HOJAS.PERIODOS);
  h.appendRow(["ID","Período","Fecha_Inicio","Fecha_Fin","Estado","Total_Q","Tab_Reporte","Fecha_Cierre"]);
  _fmtEnc(h, "#37474f");
  h.setColumnWidth(2, 200);
  h.setColumnWidth(3, 120);
  h.setColumnWidth(4, 120);
  h.setColumnWidth(7, 200);
  return h;
}

/*
 * Pide al usuario las fechas de inicio y fin de la nueva quincena,
 * registra en PERIODOS y genera el reporte inicial.
 */
function configurarNuevaQuincena() { _run(function() {
  var ui   = SpreadsheetApp.getUi();
  var tz   = Session.getScriptTimeZone();
  var ahora = new Date();

  // Sugerir fecha inicio = hoy
  var sugeridaIni = Utilities.formatDate(ahora, tz, "dd/MM/yyyy");
  // Sugerir fecha fin = hoy + 14 días
  var sugeridaFin = Utilities.formatDate(new Date(ahora.getTime() + 14*24*3600*1000), tz, "dd/MM/yyyy");

  var r1 = ui.prompt("🗓️ Nueva quincena — Fecha inicio",
    "Formato dd/mm/yyyy\n(sugerido: "+sugeridaIni+")", ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var fi = _parseFecha(r1.getResponseText().trim() || sugeridaIni);
  if (!fi) { _alert("Fecha de inicio inválida."); return; }

  var r2 = ui.prompt("🗓️ Nueva quincena — Fecha fin",
    "Formato dd/mm/yyyy\n(sugerido: "+sugeridaFin+")", ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var ff = _parseFecha(r2.getResponseText().trim() || sugeridaFin);
  if (!ff || ff < fi) { _alert("Fecha de fin inválida o anterior al inicio."); return; }

  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var hP  = crearHojaPeriodos();

  // Marcar cualquier período activo anterior como "En pausa"
  var datos = hP.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][4]) === "Activo") {
      hP.getRange(i+1, 5).setValue("En pausa");
    }
  }

  var label = _labelPeriodo(fi, ff);
  var id    = datos.length; // simple incremental ID
  var tab   = _tabNombreQuincena(fi, ff);

  hP.appendRow([id, label, fi, ff, "Activo", 0, tab, ""]);

  // Formato de fechas en columnas C y D
  var lastR = hP.getLastRow();
  hP.getRange(lastR, 3, 1, 2).setNumberFormat("dd/MM/yyyy");
  hP.getRange(lastR, 5).setBackground("#e6f4ea").setFontColor("#137333").setFontWeight("bold");

  // Generar el reporte de inmediato
  var total = _generarReporteQuincena(fi, ff, label, tab, {});
  hP.getRange(lastR, 6).setValue(total);

  ss.setActiveSheet(ss.getSheetByName(tab) || ss.getActiveSheet());
  _alert("✅ Quincena configurada: " + label + "\n\nTotal actual: Q" + total.toFixed(2) +
         "\n\nEl reporte se actualizará automáticamente cada día.\n" +
         "Usa 'Cerrar quincena y crear siguiente' cuando pagues.");
}); }

/*
 * Muestra (o crea) el reporte de la quincena activa.
 * Si no hay ninguna activa, ofrece crear una.
 */
function verQuincenaActual() { _run(function() {
  var periodo = _periodoActivo();
  if (!periodo) {
    var ui = SpreadsheetApp.getUi();
    var r  = ui.alert("No hay quincena activa",
      "No hay ninguna quincena configurada como Activa.\n¿Quieres configurar una ahora?",
      ui.ButtonSet.YES_NO);
    if (r === ui.Button.YES) { configurarNuevaQuincena(); }
    return;
  }
  var fi  = new Date(periodo.fi);
  var ff  = new Date(periodo.ff);
  var tab = periodo.tab;
  var label = periodo.label;

  // Re-generar (actualizar) el reporte
  var hP = _sh(CFG.HOJAS.PERIODOS);
  var horasReponer = _leerHorasReponerExistentes(tab);
  var total = _generarReporteQuincena(fi, ff, label, tab, horasReponer);

  // Actualizar total en PERIODOS
  hP.getRange(periodo.fila, 6).setValue(total);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hTab = ss.getSheetByName(tab);
  if (hTab) ss.setActiveSheet(hTab);

  _alert("✅ Quincena actualizada: " + label + "\nTotal: Q" + total.toFixed(2));
}); }

/*
 * Función llamada por el trigger diario — actualiza silenciosamente.
 */
function actualizarQuincenaActual() {
  try {
    var periodo = _periodoActivo();
    if (!periodo) return;
    var fi  = new Date(periodo.fi);
    var ff  = new Date(periodo.ff);
    var horasReponer = _leerHorasReponerExistentes(periodo.tab);
    var total = _generarReporteQuincena(fi, ff, periodo.label, periodo.tab, horasReponer);
    var hP = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.HOJAS.PERIODOS);
    if (hP) hP.getRange(periodo.fila, 6).setValue(total);
  } catch(e) {
    Logger.log("actualizarQuincenaActual error: " + e.message);
  }
}

/*
 * Genera el reporte de quincena para cualquier período pasado.
 * Pide fecha inicio y fin, calcula horas desde DatosKobo y muestra
 * el resultado en el mismo formato de planilla.
 */
function generarReporteQuincenaPasada() { _run(function() {
  var ui  = SpreadsheetApp.getUi();
  var tz  = Session.getScriptTimeZone();

  var r1 = ui.prompt("📅 Reporte de quincena pasada",
    "Fecha INICIO del período (dd/mm/yyyy):", ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var fi = _parseFecha(r1.getResponseText().trim());
  if (!fi) { _alert("Fecha de inicio inválida. Usa dd/mm/yyyy (ej. 11/04/2026)"); return; }

  var r2 = ui.prompt("📅 Reporte de quincena pasada",
    "Fecha FIN del período (dd/mm/yyyy):", ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var ff = _parseFecha(r2.getResponseText().trim());
  if (!ff || ff < fi) { _alert("Fecha de fin inválida o anterior al inicio."); return; }

  var label  = _labelPeriodo(fi, ff);
  var tab    = "Q_" + Utilities.formatDate(fi, tz, "dd_MM") + "_" +
               Utilities.formatDate(ff, tz, "dd_MM_yyyy") + "_hist";

  // Leer horas a reponer previas si ya existía esta hoja
  var prevReponer = _leerHorasReponerExistentes(tab);

  var total  = _generarReporteQuincena(fi, ff, label, tab, prevReponer);
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var hTab   = ss.getSheetByName(tab);
  if (hTab) ss.setActiveSheet(hTab);

  _alert("✅ Reporte generado: " + label + "\nTotal: Q" + total.toFixed(2) +
         "\n\nHoja: '" + tab + "'\n\n" +
         (total === 0
           ? "⚠️ Total en cero — verifica que DatosKobo tenga registros para ese período."
           : "Puedes editar la columna 'Horas a reponer' y el reporte se actualiza al guardar."));
}); }

/*
 * Cierra la quincena activa y abre la siguiente.
 */
function cerrarQuincenaYCrearSiguiente() { _run(function() {
  var periodo = _periodoActivo();
  if (!periodo) { _alert("No hay quincena activa. Usa 'Configurar nueva quincena'."); return; }

  var ui  = SpreadsheetApp.getUi();
  var tz  = Session.getScriptTimeZone();
  var hP  = _sh(CFG.HOJAS.PERIODOS);

  var r = ui.alert("Cerrar quincena",
    "¿Cerrar el período " + periodo.label + "?\n\n" +
    "Total: Q" + (periodo.total || "—") + "\n\n" +
    "Se marcará como Cerrado y se configurará la siguiente quincena.",
    ui.ButtonSet.OK_CANCEL);
  if (r !== ui.Button.OK) return;

  // Cerrar período actual
  var hoy = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy");
  hP.getRange(periodo.fila, 5).setValue("Cerrado").setBackground("#fce8e6").setFontColor("#c5221f");
  hP.getRange(periodo.fila, 8).setValue(hoy);

  // Sugerir siguientes fechas: inicio = fin anterior + 1 día
  var ffAnterior = new Date(periodo.ff);
  var siguienteIni = new Date(ffAnterior.getTime() + 24*3600*1000);
  var siguienteFin = new Date(siguienteIni.getTime() + 14*24*3600*1000);
  var fmtIni = Utilities.formatDate(siguienteIni, tz, "dd/MM/yyyy");
  var fmtFin = Utilities.formatDate(siguienteFin, tz, "dd/MM/yyyy");

  var r2 = ui.prompt("🗓️ Siguiente quincena — Fecha inicio",
    "Sugerido: "+fmtIni, ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) { _alert("Quincena cerrada. Configura la siguiente cuando quieras."); return; }
  var fi2 = _parseFecha(r2.getResponseText().trim() || fmtIni);

  var r3 = ui.prompt("🗓️ Siguiente quincena — Fecha fin",
    "Sugerido: "+fmtFin, ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() !== ui.Button.OK) { _alert("Quincena cerrada. Configura la siguiente cuando quieras."); return; }
  var ff2 = _parseFecha(r3.getResponseText().trim() || fmtFin);

  if (!fi2 || !ff2 || ff2 < fi2) { _alert("Fechas inválidas. Quincena cerrada. Configura manualmente."); return; }

  var label2 = _labelPeriodo(fi2, ff2);
  var tab2   = _tabNombreQuincena(fi2, ff2);
  var datos  = hP.getDataRange().getValues();
  hP.appendRow([datos.length, label2, fi2, ff2, "Activo", 0, tab2, ""]);
  var lastR = hP.getLastRow();
  hP.getRange(lastR, 3, 1, 2).setNumberFormat("dd/MM/yyyy");
  hP.getRange(lastR, 5).setBackground("#e6f4ea").setFontColor("#137333").setFontWeight("bold");

  var total2 = _generarReporteQuincena(fi2, ff2, label2, tab2, {});
  hP.getRange(lastR, 6).setValue(total2);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hTab = ss.getSheetByName(tab2);
  if (hTab) ss.setActiveSheet(hTab);

  _alert("✅ Quincena anterior cerrada.\n\nNueva quincena abierta: " + label2 +
         "\nTotal actual: Q" + total2.toFixed(2));
}); }

// ── Helpers de quincena ─────────────────────────────────────────

function _periodoActivo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hP = ss.getSheetByName(CFG.HOJAS.PERIODOS);
  if (!hP || hP.getLastRow() < 2) return null;
  var datos = hP.getDataRange().getValues();
  for (var i = datos.length - 1; i >= 1; i--) {
    if (String(datos[i][4]) === "Activo") {
      return {
        fila:  i+1,
        id:    datos[i][0],
        label: String(datos[i][1]),
        fi:    datos[i][2],
        ff:    datos[i][3],
        estado:String(datos[i][4]),
        total: parseFloat(datos[i][5]) || 0,
        tab:   String(datos[i][6])
      };
    }
  }
  return null;
}

function _labelPeriodo(fi, ff) {
  var tz  = Session.getScriptTimeZone();
  return Utilities.formatDate(fi, tz, "dd/MM/yyyy") + " al " +
         Utilities.formatDate(ff, tz, "dd/MM/yyyy");
}

function _tabNombreQuincena(fi, ff) {
  var tz = Session.getScriptTimeZone();
  return "Q_" + Utilities.formatDate(fi, tz, "dd_MM") + "_" +
                Utilities.formatDate(ff, tz, "dd_MM_yyyy");
}

/*
 * Lee las "Horas a reponer" que el usuario ya editó en el reporte anterior,
 * para no perderlas al actualizar.
 */
function _leerHorasReponerExistentes(tabNombre) {
  var mapa = {};
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var h    = ss.getSheetByName(tabNombre);
  if (!h || h.getLastRow() < 4) return mapa;
  var datos = h.getDataRange().getValues();
  // Col B = nombre, Col G (idx 6) = Horas a reponer (fila 4 en adelante)
  for (var i = 3; i < datos.length; i++) {
    var nombre = String(datos[i][1]).trim().replace(/^SUBTOTAL\s*/i, "").trim();
    var hrs    = parseFloat(datos[i][6]);
    if (nombre && !isNaN(hrs) && hrs > 0) mapa[nombre] = hrs;
  }
  return mapa;
}

/*
 * Calcula el resumen de horas por participante para un período dado.
 * Lee desde DatosKobo emparejando entradas y salidas.
 * Retorna: { "Nombre": { horas, tarifa, tieneFactura, codigo } }
 */
function _calcularResumenPeriodo(fi, ff) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hK   = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (!hK || hK.getLastRow() < 2) return {};

  var enc  = hK.getRange(1, 1, 1, hK.getLastColumn()).getValues()[0];
  var raw  = hK.getRange(2, 1, hK.getLastRow()-1, hK.getLastColumn()).getValues();
  var cols = detectarColumnas(enc, raw.slice(0, 50));

  var mapeoNombres = cargarMapeoNombres();
  var diasEstudio  = obtenerDiasEstudio();
  var listaTerapias= obtenerListaTerapias();

  var mapa = _construirMapaTarifas();

  // Agrupar registros por participante, ordenados por timestamp
  var porPart = {};

  raw.forEach(function(fila) {
    var ts  = new Date(fila[cols.start !== undefined ? cols.start : 0]);
    if (isNaN(ts)) return;
    // Filtrar por rango de fechas (comparar solo fecha, no hora)
    var dia = new Date(ts.getFullYear(), ts.getMonth(), ts.getDate());
    var dIni = new Date(fi.getFullYear(), fi.getMonth(), fi.getDate());
    var dFin = new Date(ff.getFullYear(), ff.getMonth(), ff.getDate());
    if (dia < dIni || dia > dFin) return;

    var nombreRaw = obtenerParticipanteFila(fila, cols);
    if (!nombreRaw) return;
    var nombre = normalizarNombre(nombreRaw, mapeoNombres) || limpiarNombre(nombreRaw);
    if (!nombre) return;

    var tipo = obtenerTipoRegistro(fila, cols);
    if (!tipo.esIngreso && !tipo.esEgreso) return;

    if (!porPart[nombre]) porPart[nombre] = [];
    porPart[nombre].push({ ts: ts, tipo: tipo, fila: fila });
  });

  // Calcular horas por participante
  var resultado = {};

  Object.keys(porPart).forEach(function(nombre) {
    var registros = porPart[nombre].sort(function(a,b){ return a.ts - b.ts; });
    var totalHoras = 0;
    var pendienteEntrada = null;

    registros.forEach(function(reg) {
      if (reg.tipo.esIngreso && !pendienteEntrada) {
        pendienteEntrada = reg.ts;
      } else if (reg.tipo.esEgreso && pendienteEntrada) {
        var diffH = (reg.ts - pendienteEntrada) / 3600000;
        if (diffH > 0 && diffH <= 16) {
          // Verificar si es día de estudio (0% pago — pero contamos horas trabajadas)
          var fechaDia = new Date(pendienteEntrada.getFullYear(),
                                  pendienteEntrada.getMonth(),
                                  pendienteEntrada.getDate());
          totalHoras += diffH;
        }
        pendienteEntrada = null;
      }
    });

    // Entrada sin salida al final → estimar salida con jornada normal
    if (pendienteEntrada) {
      var estimada = new Date(pendienteEntrada.getTime() + CFG.HORAS_JORNADA_NORMAL * 3600000);
      var diffH = Math.min((estimada - pendienteEntrada) / 3600000, CFG.HORAS_JORNADA_NORMAL);
      if (diffH > 0) totalHoras += diffH;
    }

    if (totalHoras <= 0) return;

    var info = mapa[limpiarNombre(nombre)] || mapa[nombre] || {};
    var codigo = extraerCodigo(nombre) || info.codigo || "";

    resultado[nombre] = {
      horas:         Math.round(totalHoras * 100) / 100,
      tarifa:        info.tarifa || CFG.CATEGORIAS.C,
      tieneFactura:  info.tieneFactura || false,
      codigo:        codigo
    };
  });

  return resultado;
}

/*
 * Genera (o actualiza) la hoja del reporte de quincena.
 * Formato: coincide con la captura de pantalla del usuario.
 * Preserva "Horas a reponer" que el usuario haya editado.
 * Retorna: total Q del período.
 */
function _generarReporteQuincena(fi, ff, label, tabNombre, prevHorasReponer) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var resumen = _calcularResumenPeriodo(fi, ff);
  var mapa    = _construirMapaTarifas();
  var tz      = Session.getScriptTimeZone();
  var ts      = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm");

  // Recrear o limpiar hoja
  var h = ss.getSheetByName(tabNombre);
  if (h) {
    h.clearContents();
    h.clearFormats();
  } else {
    h = ss.insertSheet(tabNombre);
  }

  // Asegurar suficientes columnas
  if (h.getMaxColumns() < 12) h.insertColumnsAfter(h.getMaxColumns(), 12 - h.getMaxColumns());

  // ── Fila 1: referencia de tarifas ──────────────────────────────
  h.getRange(1, 1, 1, 12).setValues([["Q16.50","Q15.75","Q15.00","Q14.00","","","","","","","",""]]);
  ["#639922","#4285f4","#fbbc04","#ea4335"].forEach(function(c, i) {
    h.getRange(1, i+1).setBackground(c).setFontColor("#ffffff").setFontWeight("bold")
     .setHorizontalAlignment("center");
  });

  // ── Fila 2: título del período ─────────────────────────────────
  h.getRange(2, 1, 1, 12).merge()
   .setValue("Período: " + label)
   .setBackground("#f8f9fa").setFontWeight("bold").setFontSize(11)
   .setHorizontalAlignment("center")
   .setBorder(true,true,true,true,null,null,"#dadce0",SpreadsheetApp.BorderStyle.SOLID);

  // ── Fila 3: encabezados de columnas ───────────────────────────
  var encabezados = ["#","Participantes","","","",
    "Total de horas","Horas a reponer","Total a pagar",
    "Monto (Q)","IVA – 5%","Pago + IVA","Redondeo"];
  h.getRange(3, 1, 1, 12).setValues([encabezados])
   .setBackground("#546e7a").setFontColor("#ffffff").setFontWeight("bold")
   .setHorizontalAlignment("center")
   .setBorder(true,true,true,true,null,null,"#37474f",SpreadsheetApp.BorderStyle.SOLID);

  // ── Filas de datos: una por participante ───────────────────────
  var nombres = Object.keys(resumen).sort(function(a,b){ return a.localeCompare(b,"es"); });
  var filaActual = 4;
  var totalGeneral = 0;
  var num = 1;

  nombres.forEach(function(nombre) {
    var d = resumen[nombre];
    // Recuperar horas a reponer previas (editadas manualmente)
    var hReponer = prevHorasReponer[nombre] || 0;
    var hTotal   = Math.round((d.horas + hReponer) * 100) / 100;
    var monto    = Math.round(hTotal * d.tarifa * 100) / 100;
    var iva      = d.tieneFactura ? Math.round(monto * CFG.IVA_PCT * 100) / 100 : 0;
    var conIVA   = Math.round((monto + iva) * 100) / 100;
    var redond   = Math.round(conIVA);
    totalGeneral += redond;

    var label = nombre + (d.codigo ? " (" + d.codigo + ")" : "");

    h.getRange(filaActual, 1, 1, 12).setValues([[
      num++,
      label,
      "","","",
      d.horas,
      hReponer > 0 ? hReponer : "",
      hTotal,
      monto,
      iva > 0 ? iva : "",
      iva > 0 ? conIVA : monto,
      redond
    ]]);

    // Color de fila alternado
    var bg = (num % 2 === 0) ? "#f8f9fa" : "#ffffff";
    h.getRange(filaActual, 1, 1, 12).setBackground(bg);
    h.getRange(filaActual, 1).setHorizontalAlignment("center");
    h.getRange(filaActual, 6, 1, 7).setHorizontalAlignment("right");
    // Formato moneda en cols 9-12
    h.getRange(filaActual, 9, 1, 4).setNumberFormat('"Q"#,##0.00');

    filaActual++;
  });

  // ── Fila total ─────────────────────────────────────────────────
  filaActual++;
  h.getRange(filaActual, 1, 1, 12).setValues([["","","","","","","","","","","",totalGeneral]]);
  h.getRange(filaActual, 12)
   .setBackground("#00c853").setFontColor("#ffffff").setFontWeight("bold")
   .setFontSize(12).setHorizontalAlignment("center")
   .setNumberFormat('"Q"#,##0.00');

  // ── Fila timestamp ─────────────────────────────────────────────
  filaActual++;
  h.getRange(filaActual, 1, 1, 12).merge()
   .setValue("Actualizado: " + ts + "  |  " + nombres.length + " participantes")
   .setFontSize(8).setFontColor("#9aa0a6").setHorizontalAlignment("right");

  // ── Ancho de columnas ──────────────────────────────────────────
  [35, 230, 30, 30, 30, 90, 100, 90, 90, 80, 90, 85].forEach(function(w, i) {
    h.setColumnWidth(i+1, w);
  });
  h.setRowHeight(2, 28);
  h.setRowHeight(3, 24);
  h.setFrozenRows(3);

  return totalGeneral;
}

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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Eliminar triggers manejados por este sistema
  ScriptApp.getProjectTriggers().forEach(function(t){
    var h = t.getHandlerFunction();
    if (h === "importarDesdeKobo" || h === "importarAlAbrir" ||
        h === "actualizarQuincenaActual") {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Importar Kobo cada hora
  ScriptApp.newTrigger("importarDesdeKobo").timeBased().everyHours(1).create();
  // Importar Kobo al abrir el Spreadsheet
  ScriptApp.newTrigger("importarAlAbrir").forSpreadsheet(ss).onOpen().create();
  // Actualizar reporte de quincena activa cada día a las 7am
  ScriptApp.newTrigger("actualizarQuincenaActual")
    .timeBased().everyDays(1).atHour(7).create();

  _alert("✅ Automatizaciones activadas:\n\n" +
    "• ⏰ Kobo: importa datos cada hora\n" +
    "• 🔄 Kobo: importa al abrir la hoja\n" +
    "• 📅 Quincena activa: se actualiza cada día a las 7am\n\n" +
    "onEdit (automático):\n" +
    "• Categoría → auto-llena Tarifa\n" +
    "• Pagado → actualiza Dashboard");
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
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName(CFG.HOJAS.DASHBOARD) || ss.insertSheet(CFG.HOJAS.DASHBOARD, 0);
  var kpi  = _calcularKpis(ss);
  var tz   = Session.getScriptTimeZone();
  var ts   = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm");
  var mes  = kpi.mes + " " + kpi.anio;

  // ── Construir filas ───────────────────────────────────────────
  // Formato: [col A (etiqueta), col B (valor), col C (descripción), col D (período)]
  var filas = [], tipos = [];
  function push(f, t) { filas.push(f); tipos.push(t); }

  push(["MI EELO — INDICADORES CLAVE DE DESEMPEÑO", "", "", ""], "banner");
  push(["Actualizado: " + ts + "  |  Período mensual: " + mes, "", "", ""], "sub");
  push(["", "", "", ""], "vacio");

  // ── Sección KPIs ─────────────────────────────────────────────
  push(["#", "INDICADOR", "VALOR", "PERÍODO / FUENTE"], "enc");

  // KPI 1: Ciclos de Vida completados
  push(["KPI 1",
        "Ciclos de Vida completados (año en curso)",
        kpi.ciclosCompletados,
        "Año " + kpi.anio + " · RRHH"],
       kpi.ciclosCompletados > 0 ? "kpi_ok" : "kpi_cero");

  // KPI 2: Ingresos brutos promedio por participante/mes
  push(["KPI 2",
        "Ingresos brutos promedio por participante",
        kpi.ingresoPromedio > 0 ? "Q " + kpi.ingresoPromedio.toFixed(2) : "—",
        mes + " · RRHH"],
       kpi.ingresoPromedio > 0 ? "kpi_ok" : "kpi_nd");

  // KPI 3: Pedidos/encargos totales
  push(["KPI 3",
        "Total pedidos / encargos formalizados",
        kpi.pedidosTotales !== null ? String(kpi.pedidosTotales) : "Ver PRODUCCIÓN",
        mes + " · PRODUCCIÓN"],
       kpi.pedidosTotales !== null ? "kpi_ok" : "kpi_prod");

  // KPI 4: Clientes únicos activos
  push(["KPI 4",
        "Clientes únicos activos del programa",
        kpi.clientesUnicos !== null ? String(kpi.clientesUnicos) : "Ver PRODUCCIÓN",
        mes + " · PRODUCCIÓN"],
       kpi.clientesUnicos !== null ? "kpi_ok" : "kpi_prod");

  // KPI 5: Horas de formación
  push(["KPI 5",
        "Horas de formación (nuevas y existentes participantes)",
        kpi.horasFormacion > 0 ? kpi.horasFormacion.toFixed(1) + " hrs" : "0 hrs",
        mes + " · RRHH"],
       kpi.horasFormacion > 0 ? "kpi_ok" : "kpi_cero");

  // KPI 6: Promedio horas laborales mensuales
  push(["KPI 6",
        "Promedio de horas laborales mensuales por participante",
        kpi.promedioHorasLaborales > 0 ? kpi.promedioHorasLaborales.toFixed(1) + " hrs" : "—",
        mes + " · RRHH"],
       kpi.promedioHorasLaborales > 0 ? "kpi_ok" : "kpi_nd");

  push(["", "", "", ""], "vacio");

  // ── Sección resumen del mes ───────────────────────────────────
  push(["RESUMEN MENSUAL — " + mes, "", "", ""], "sec");
  push(["Participantes activas",   String(kpi.activos),             "Estado = Activo en PARTICIPANTES", ""], "dato");
  push(["Total horas trabajadas",  kpi.totalHorasMes.toFixed(1)+" hrs", "Suma Q1+Q2 del mes", ""], "dato");
  push(["Monto base total",        "Q "+kpi.sumBase.toFixed(2),    "Sin IVA, todas las participantes", ""], "dato");
  push(["IVA 5% (Pequeño Contrib.)","Q "+kpi.sumIVA.toFixed(2),   "Solo quienes tienen factura", ""], "dato");
  push(["TOTAL A PAGAR (facturas)","Q "+kpi.sumTot.toFixed(2),     "Lo que paga la organización", ""], "total");
  push(["Quincenas pagadas",       String(kpi.pagadas),             "", ""], kpi.pagadas>0?"dato_pag":"dato");
  push(["Quincenas pendientes",    String(kpi.pendientes),          "", ""],
       kpi.pendientes>0 ? "dato_pend" : "dato");

  push(["", "", "", ""], "vacio");

  // ── Referencia tarifas ────────────────────────────────────────
  push(["TARIFAS VIGENTES", "A = Q16.50/hr", "B = Q15.75/hr", "C = Q15.00/hr  |  D = Q14.00/hr"], "ref");

  // ── Notas sobre KPIs que necesitan seguimiento manual ─────────
  push(["", "", "", ""], "vacio");
  push(["NOTAS", "", "", ""], "sec");
  push(["KPI 1 — Ciclo de Vida",
        "Marca Estado = 'Egresado' en PARTICIPANTES cuando alguien completa el programa.",
        "", ""], "nota");
  push(["KPI 3 y 4 — Producción",
        "Estos datos vienen del sistema de Producción (PRODUCCION_todo.gs). Actualizar manualmente o integrar.",
        "", ""], "nota");
  push(["KPI 5 — Formación",
        "Se cuentan horas en días de estudio y terapias desde la hoja ASISTENCIA del mes.",
        "", ""], "nota");

  // ── Escribir en hoja ──────────────────────────────────────────
  dash.clearContents();
  dash.clearFormats();
  if (dash.getMaxColumns() < 4) dash.insertColumnsAfter(dash.getMaxColumns(), 4 - dash.getMaxColumns());

  dash.getRange(1, 1, filas.length, 4).setValues(filas);

  var VERDE  = "#34a853", AZUL = "#1a73e8", ROJO = "#c5221f";
  var AMARILLO = "#f9ab00", GRIS = "#9aa0a6";

  tipos.forEach(function(tipo, idx) {
    var r = dash.getRange(idx+1, 1, 1, 4);
    r.setFontFamily("Arial").setFontSize(10).setVerticalAlignment("middle");

    switch (tipo) {
      case "banner":
        r.merge().setFontSize(14).setFontWeight("bold")
         .setBackground("#1a237e").setFontColor("#ffffff")
         .setHorizontalAlignment("center").setRowHeight && dash.setRowHeight(idx+1, 36);
        break;
      case "sub":
        r.merge().setFontSize(9).setBackground("#e8eaf6")
         .setFontColor("#5f6368").setHorizontalAlignment("center");
        break;
      case "enc":
        r.setFontWeight("bold").setBackground("#202124").setFontColor("#ffffff")
         .setHorizontalAlignment("center");
        break;
      case "kpi_ok":
        dash.getRange(idx+1,1,1,1).setBackground("#e6f4ea").setFontWeight("bold").setHorizontalAlignment("center");
        dash.getRange(idx+1,2,1,1).setBackground("#f8f9fa");
        dash.getRange(idx+1,3,1,1).setBackground("#e6f4ea").setFontWeight("bold").setFontSize(13).setFontColor(VERDE).setHorizontalAlignment("center");
        dash.getRange(idx+1,4,1,1).setBackground("#f8f9fa").setFontColor(GRIS);
        break;
      case "kpi_cero":
        dash.getRange(idx+1,1,1,1).setBackground("#fef7e0").setFontWeight("bold").setHorizontalAlignment("center");
        dash.getRange(idx+1,2,1,1).setBackground("#fefefe");
        dash.getRange(idx+1,3,1,1).setBackground("#fef7e0").setFontWeight("bold").setFontSize(13).setFontColor(AMARILLO).setHorizontalAlignment("center");
        dash.getRange(idx+1,4,1,1).setFontColor(GRIS);
        break;
      case "kpi_nd":
        dash.getRange(idx+1,1,1,1).setBackground("#f1f3f4").setFontWeight("bold").setHorizontalAlignment("center");
        dash.getRange(idx+1,2,1,1).setBackground("#fefefe");
        dash.getRange(idx+1,3,1,1).setBackground("#f1f3f4").setFontColor(GRIS).setFontSize(12).setHorizontalAlignment("center");
        dash.getRange(idx+1,4,1,1).setFontColor(GRIS);
        break;
      case "kpi_prod":
        dash.getRange(idx+1,1,1,1).setBackground("#e8f0fe").setFontWeight("bold").setHorizontalAlignment("center");
        dash.getRange(idx+1,2,1,1).setBackground("#fefefe");
        dash.getRange(idx+1,3,1,1).setBackground("#e8f0fe").setFontColor(AZUL).setFontStyle("italic").setFontSize(11).setHorizontalAlignment("center");
        dash.getRange(idx+1,4,1,1).setFontColor(GRIS);
        break;
      case "sec":
        r.merge().setFontWeight("bold").setFontSize(11)
         .setBackground("#37474f").setFontColor("#ffffff");
        break;
      case "total":
        r.setFontWeight("bold").setBackground("#e8eaf6").setFontSize(11);
        dash.getRange(idx+1,2,1,1).setFontColor(AZUL).setFontSize(12).setFontWeight("bold");
        break;
      case "dato_pag":
        dash.getRange(idx+1,2,1,1).setFontColor(VERDE).setFontWeight("bold");
        break;
      case "dato_pend":
        dash.getRange(idx+1,2,1,1).setFontColor(ROJO).setFontWeight("bold");
        r.setBackground("#fff8f7");
        break;
      case "ref":
        r.setBackground("#f0f4c3").setFontWeight("bold").setFontSize(9);
        break;
      case "nota":
        dash.getRange(idx+1,1,1,1).setFontWeight("bold").setFontColor(GRIS);
        dash.getRange(idx+1,2,1,3).merge().setFontColor(GRIS).setFontStyle("italic").setFontSize(9);
        break;
    }
  });

  // Bordes horizontales en KPIs
  for (var i=0; i<tipos.length; i++) {
    if (tipos[i].indexOf("kpi_") === 0) {
      dash.getRange(i+1,1,1,4)
        .setBorder(null,null,true,null,null,null,"#dadce0",SpreadsheetApp.BorderStyle.SOLID);
    }
  }

  // Anchos de columna
  dash.setColumnWidth(1, 70);
  dash.setColumnWidth(2, 310);
  dash.setColumnWidth(3, 160);
  dash.setColumnWidth(4, 200);
  dash.setFrozenRows(2);

}); }

/*
 * Calcula todos los KPIs del mes en curso.
 * Retorna objeto con valores para el Dashboard.
 */
function _calcularKpis(ss) {
  var ahora    = new Date();
  var mesNum   = ahora.getMonth()+1;
  var anio     = ahora.getFullYear();
  var nombreMes = CFG.MESES[mesNum-1];
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();

  var kpi = {
    mes: nombreMes, anio: anio,
    // KPI 1
    ciclosCompletados: 0,
    // KPI 2
    ingresoPromedio: 0,
    // KPI 3 & 4 — vienen de Producción
    pedidosTotales: null,
    clientesUnicos: null,
    // KPI 5
    horasFormacion: 0,
    // KPI 6
    promedioHorasLaborales: 0,
    // Resumen
    activos: 0, sumBase: 0, sumIVA: 0, sumTot: 0,
    pagadas: 0, pendientes: 0, totalHorasMes: 0
  };

  // ── PARTICIPANTES ─────────────────────────────────────────────
  var hojaP = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES);
  if (hojaP && hojaP.getLastRow() > 1) {
    hojaP.getDataRange().getValues().slice(1).forEach(function(r) {
      var estado = String(r[5]).trim().toLowerCase();
      if (estado === "activo") kpi.activos++;
      if (estado === "egresado") kpi.ciclosCompletados++;
    });
  }

  // ── FACTURACION ───────────────────────────────────────────────
  var hojaF = ss.getSheetByName(CFG.HOJAS.FACTURACION);
  if (hojaF && hojaF.getLastRow() > 1) {
    var participantesMes = {};
    hojaF.getDataRange().getValues().slice(1).forEach(function(f) {
      if (f[2] !== nombreMes || Number(f[3]) !== anio) return;
      kpi.sumBase += parseFloat(f[9])  || 0;
      kpi.sumIVA  += parseFloat(f[11]) || 0;
      kpi.sumTot  += parseFloat(f[12]) || 0;
      if (String(f[17]) === "Sí") kpi.pagadas++; else kpi.pendientes++;

      // Acumular horas por participante para KPI 6
      var nombre = String(f[1]).trim();
      if (!participantesMes[nombre]) participantesMes[nombre] = 0;
      participantesMes[nombre] += parseFloat(f[5]) || 0; // Horas_Trabajadas (idx 5)
      kpi.totalHorasMes += parseFloat(f[5]) || 0;
    });

    // KPI 2: promedio ingresos brutos por participante del mes
    var numPart = Object.keys(participantesMes).length;
    if (numPart > 0) {
      // Suma total por participante (Q1+Q2), luego promedio
      var totalPorPart = {};
      hojaF.getDataRange().getValues().slice(1).forEach(function(f) {
        if (f[2] !== nombreMes || Number(f[3]) !== anio) return;
        var n = String(f[1]).trim();
        totalPorPart[n] = (totalPorPart[n] || 0) + (parseFloat(f[12]) || 0);
      });
      var vals = Object.keys(totalPorPart).map(function(n){ return totalPorPart[n]; });
      kpi.ingresoPromedio = vals.reduce(function(s,v){ return s+v; },0) / vals.length;

      // KPI 6: promedio horas laborales por participante
      var hVals = Object.keys(participantesMes).map(function(n){ return participantesMes[n]; });
      kpi.promedioHorasLaborales = hVals.reduce(function(s,v){ return s+v; },0) / hVals.length;
    }
  }

  // ── KPI 5: Horas de formación ─────────────────────────────────
  // Cuenta desde ASISTENCIA: filas del mes donde Es_Dia_Estudio="Sí" o Es_Terapia="Sí"
  var hojaA = ss.getSheetByName(CFG.HOJAS.ASISTENCIA);
  if (hojaA && hojaA.getLastRow() > 1) {
    hojaA.getDataRange().getValues().slice(1).forEach(function(a) {
      var fecha = new Date(a[2]);
      if (isNaN(fecha) || fecha.getMonth()+1 !== mesNum || fecha.getFullYear() !== anio) return;
      var esEstudio  = String(a[5]).toLowerCase() === "sí" || String(a[5]) === "TRUE";
      var esTerapia  = String(a[6]).toLowerCase() === "sí" || String(a[6]) === "TRUE";
      if (esEstudio || esTerapia) {
        kpi.horasFormacion += parseFloat(a[4]) || 0; // Horas_Trabajadas (idx 4)
      }
    });
  }

  // ── KPI 3 & 4: intentar leer desde hoja de Producción ─────────
  // Si existe la hoja "ORDENES" en el mismo Spreadsheet, contamos
  var hOrd = ss.getSheetByName("ORDENES");
  if (hOrd && hOrd.getLastRow() > 1) {
    var ordenes = hOrd.getDataRange().getValues().slice(1);
    // Pedidos del mes (columna de fecha — asumimos col C=idx2)
    var clientesSet = {};
    ordenes.forEach(function(o) {
      var fOrd = new Date(o[2]);
      if (isNaN(fOrd)) return;
      if (fOrd.getMonth()+1 === mesNum && fOrd.getFullYear() === anio) {
        kpi.pedidosTotales = (kpi.pedidosTotales || 0) + 1;
        var cli = String(o[4]||"").trim(); // cliente — columna E (idx 4), ajustar si difiere
        if (cli) clientesSet[cli] = true;
      }
    });
    kpi.clientesUnicos = Object.keys(clientesSet).length || null;
  }

  return kpi;
}

// ══════════════════════════════════════════════════════════════════
// REPORTES — escritos como pestañas del Spreadsheet
// Igual que codigo.gs: por día / semana / mes / rango / completo
// Formato: Fecha | Día | 🟢 Entrada | 🔴 Salida | Tipo |
//          Horas Trabajadas | Porcentaje | Horas a Pagar | Tarifa | Monto (Q)
// Agrupado por participante con SUBTOTAL y resumen general al final.
// Lee desde DatosKobo (como codigo.gs): empareja en tiempo real,
// salidas estimadas, filtro por participante, opción nueva hoja.
// ══════════════════════════════════════════════════════════════════

function generarReportePorDia() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt("📅 Reporte por Día", "Fecha (dd/mm/yyyy):", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var fecha = _parseFecha(r.getResponseText().trim());
  if (!fecha) { ui.alert("Fecha inválida. Usa formato dd/mm/yyyy (ej. 15/04/2025)"); return; }
  var params = solicitarParametrosReporte();
  if (!params) return;
  generarReporte("dia", fecha, null, params.filtro, params.nuevaHoja);
}); }

function generarReportePorSemana() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt("📆 Reporte por Semana",
    "Fecha inicio de semana (dd/mm/yyyy):", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var fecha = _parseFecha(r.getResponseText().trim());
  if (!fecha) { ui.alert("Fecha inválida. Usa formato dd/mm/yyyy"); return; }
  var params = solicitarParametrosReporte();
  if (!params) return;
  generarReporte("semana", fecha, null, params.filtro, params.nuevaHoja);
}); }

function generarReportePorMes() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var ahora = new Date();
  var r = ui.prompt("🗓️ Reporte por Mes",
    "Mes/Año (mm/yyyy):\n(ej. "+String(ahora.getMonth()+1).padStart(2,"0")+"/"+ahora.getFullYear()+")", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var parts = r.getResponseText().trim().split("/");
  if (parts.length < 2) { ui.alert("Formato inválido. Usa mm/yyyy"); return; }
  var m = parseInt(parts[0]), y = parseInt(parts[1]);
  if (isNaN(m)||isNaN(y)) { ui.alert("Mes/año inválidos."); return; }
  var params = solicitarParametrosReporte();
  if (!params) return;
  generarReporte("mes", new Date(y, m-1, 1, 0, 0, 0, 0), null, params.filtro, params.nuevaHoja);
}); }

function generarReportePorRango() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var r1 = ui.prompt("📊 Reporte por Rango (1/2)", "Fecha inicio (dd/mm/yyyy):", ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var inicio = _parseFecha(r1.getResponseText().trim());
  if (!inicio) { ui.alert("Fecha inválida."); return; }
  var r2 = ui.prompt("📊 Reporte por Rango (2/2)", "Fecha fin (dd/mm/yyyy):", ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var fin = _parseFecha(r2.getResponseText().trim());
  if (!fin) { ui.alert("Fecha inválida."); return; }
  fin.setHours(23,59,59,999);
  var params = solicitarParametrosReporte();
  if (!params) return;
  generarReporte("rango", inicio, fin, params.filtro, params.nuevaHoja);
}); }

function generarReporteTodo() { _run(function() {
  var params = solicitarParametrosReporte();
  if (!params) return;
  generarReporte("todo", null, null, params.filtro, params.nuevaHoja);
}); }

// ── Solicitar filtro + opción de hoja ────────────────────────

function solicitarParametrosReporte() {
  var ui = SpreadsheetApp.getUi();
  var lista = _obtenerListaParticipantesKobo();
  var txt = lista.slice(0,15).join("\n");
  if (lista.length > 15) txt += "\n... (" + (lista.length-15) + " más)";

  var rP = ui.prompt("👤 Filtro por participante",
    "Escribe el nombre EXACTO del participante para filtrar.\nDeja VACÍO para incluir TODOS.\n\n" +
    "Participantes disponibles:\n" + (txt || "(ninguno — importa Kobo primero)"),
    ui.ButtonSet.OK_CANCEL);
  if (rP.getSelectedButton() !== ui.Button.OK) return null;
  var filtro = rP.getResponseText().trim() || null;

  var rH = ui.alert("📄 ¿Crear nueva pestaña?",
    "SÍ → nueva pestaña (Rep_XXX_fecha)\nNO → sobreescribir pestaña 'Reporte_Fijo'",
    ui.ButtonSet.YES_NO);

  return { filtro: filtro, nuevaHoja: rH === ui.Button.YES };
}

function _obtenerListaParticipantesKobo() {
  var hojaK = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (!hojaK) return [];
  var datos = hojaK.getDataRange().getValues();
  var colsP = _buscarColsParticipante(datos[0]);
  if (colsP.col1 === -1) return [];
  var mapeo = cargarMapeoNombres();
  var names = {};
  for (var f = 1; f < datos.length; f++) {
    var n = _nombreDeFila(datos[f], colsP);
    if (n) names[normalizarNombre(n, mapeo)] = true;
  }
  return Object.keys(names).sort();
}

// ── Motor de reportes — lee desde DatosKobo ──────────────────

function generarReporte(tipo, fechaInicio, fechaFin, filtroParticipante, nuevaHoja) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaKobo = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (!hojaKobo) {
    _alert("No existe DatosKobo.\nEjecuta primero '📥 Importar asistencia desde Kobo'.");
    return;
  }

  var datos = hojaKobo.getDataRange().getValues();
  var cols = detectarColumnas(datos[0], datos.slice(1));

  if (cols.start === undefined || cols.participante === undefined) {
    _alert("ERROR: No se detectaron columnas start/participante.\nUsa '🔍 Diagnosticar Datos Kobo' para más información.");
    return;
  }
  if (cols.accionUnificada === undefined && cols.ingreso === undefined) {
    _alert("ERROR: No se detectó columna de Entrada/Salida.\nUsa '🔍 Diagnosticar Datos Kobo'.");
    return;
  }

  var diasEstudioMapa = obtenerDiasEstudio();
  var listaTerapias   = obtenerListaTerapias();
  var mapeoNombres    = cargarMapeoNombres();
  var partMap         = _construirMapaTarifas();
  var tz              = Session.getScriptTimeZone();

  // Agrupar registros por participante (dedup UUID + clave)
  var uuidVistos = {}, regVistos = {}, regPorEmp = {};
  for (var f = 1; f < datos.length; f++) {
    var fila = datos[f];

    if (cols.uuid !== undefined) {
      var uid = String(fila[cols.uuid]||"").trim();
      if (uid) { if (uuidVistos[uid]) continue; uuidVistos[uid] = true; }
    }

    var nombreRaw = obtenerParticipanteFila(fila, cols);
    if (!nombreRaw) continue;
    var emp = normalizarNombre(nombreRaw, mapeoNombres);
    if (filtroParticipante && emp.toLowerCase() !== filtroParticipante.toLowerCase()) continue;

    var tipoReg = obtenerTipoRegistro(fila, cols);
    if (!tipoReg.esIngreso && !tipoReg.esEgreso) continue;

    var tsRaw = fila[cols.start];
    if (!tsRaw) continue;
    var ts = tsRaw instanceof Date ? tsRaw : new Date(tsRaw);
    if (isNaN(ts)) continue;

    // Para salidas: usar 'end' si disponible y razonable (< 24 h)
    if (tipoReg.esEgreso && cols.end !== undefined && fila[cols.end]) {
      var tsEnd = fila[cols.end] instanceof Date ? fila[cols.end] : new Date(fila[cols.end]);
      if (!isNaN(tsEnd) && tsEnd > ts && (tsEnd - ts) < 86400000) ts = tsEnd;
    }

    var claveReg = emp + "|" + ts.getTime() + "|" + (tipoReg.esIngreso ? "E" : "S");
    if (regVistos[claveReg]) continue; regVistos[claveReg] = true;

    if (!_validarEnRango(tipo, ts, fechaInicio, fechaFin)) continue;

    if (!regPorEmp[emp]) regPorEmp[emp] = [];
    regPorEmp[emp].push({
      fecha: ts,
      esIngreso:    tipoReg.esIngreso,
      esEgreso:     tipoReg.esEgreso,
      esTerapia:    tipoReg.esTerapia,
      esPermiso:    tipoReg.esPermiso,
      esComputacion:tipoReg.esComputacion
    });
  }

  var listaEmps = Object.keys(regPorEmp).sort();
  if (listaEmps.length === 0) {
    _alert("No hay datos para el período seleccionado.\n\n" +
      "Verifica que:\n• Hayas importado datos de Kobo\n• El período tenga registros\n• El nombre del filtro sea exacto");
    return;
  }

  // Crear / limpiar pestaña
  var hoja;
  if (nuevaHoja === false) {
    var nh = ss.getSheetByName("Reporte_Fijo");
    hoja = nh || ss.insertSheet("Reporte_Fijo");
    hoja.clearContents(); hoja.clearFormats();
  } else {
    hoja = ss.insertSheet("Rep_" + tipo.substring(0,3).toUpperCase() + "_" +
                          Utilities.formatDate(new Date(), tz, "ddMM_HHmm"));
  }

  var NCOLS = 10;
  var HDRS = ["Fecha","Día","🟢 Entrada","🔴 Salida","Tipo","Horas Trabajadas","Porcentaje","Horas a Pagar","Tarifa Q/hr","Monto (Q)"];
  hoja.getRange(1,1,1,NCOLS).setValues([HDRS])
    .setFontWeight("bold").setBackground("#1f54a8").setFontColor("#fff").setHorizontalAlignment("center");
  hoja.setFrozenRows(1);

  var filaActual = 2;
  var textoPeriodo = _obtenerTextoPeriodo(tipo, fechaInicio, fechaFin);
  hoja.getRange(filaActual,1,1,NCOLS).merge()
    .setValue("Período: " + textoPeriodo + (filtroParticipante ? " | " + filtroParticipante : " | Todos los participantes"))
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center").setBackground("#e8eaf6");
  filaActual += 2;

  var totGen = { horas: 0, pagar: 0, terapia: 0, monto: 0 };

  listaEmps.forEach(function(empId) {
    var registros = regPorEmp[empId];
    registros.sort(function(a,b){ return a.fecha - b.fecha; });

    var tarifa = partMap[empId] ? partMap[empId].tarifa    : CFG.CATEGORIAS.C;
    var cat    = partMap[empId] ? partMap[empId].categoria : "";

    filaActual++;
    hoja.getRange(filaActual,1,1,NCOLS).merge()
      .setValue("👤  " + empId + (cat ? "   [Cat. " + cat + " · Q" + tarifa.toFixed(2) + "/hr]" : ""))
      .setFontWeight("bold").setBackground("#e3f2fd");
    filaActual++;

    var statsEmp = { horas: 0, pagar: 0, terapia: 0, monto: 0 };

    function _addFila(fRef, fStart, fEnd, tipoLbl, hrsV, porcV) {
      var hPagar = hrsV * porcV / 100;
      var monto  = Math.round(hPagar * tarifa * 100) / 100;
      _escribirFilaReporte(hoja, filaActual++, fRef, fStart, fEnd, tipoLbl, hrsV, porcV, tarifa, monto);
      statsEmp.horas += hrsV;
      statsEmp.pagar += hPagar;
      statsEmp.monto += monto;
      if (tipoLbl.indexOf("Terapia") !== -1) statsEmp.terapia += hrsV;
    }

    var currentIngreso = null, lastEgreso = null, fechaAnt = null;

    for (var i = 0; i < registros.length; i++) {
      var reg = registros[i];

      // Resetear al cambiar de día
      if (fechaAnt && !_esMismaFecha(fechaAnt, reg.fecha)) {
        if (currentIngreso && currentIngreso.fecha.getHours() < 17) {
          var sEst = new Date(currentIngreso.fecha.getTime() + CFG.HORAS_JORNADA_NORMAL * 3600000);
          var isDiaEst = esDiaDeEstudio(empId, currentIngreso.fecha, diasEstudioMapa);
          _addFila(currentIngreso.fecha, currentIngreso.fecha, sEst,
            (isDiaEst ? "Día de Estudio (Est.)" : "Normal (Estimado)") + "*",
            CFG.HORAS_JORNADA_NORMAL, isDiaEst ? 0 : 100);
          currentIngreso = null;
        }
        lastEgreso = null;
      }
      fechaAnt = reg.fecha;

      if (reg.esIngreso) {
        if (currentIngreso === null && lastEgreso === null && reg.fecha.getHours() >= 17) continue;
        if (currentIngreso && _esMismaFecha(currentIngreso.fecha, reg.fecha)) continue;
        currentIngreso = reg; lastEgreso = null;
      } else if (reg.esEgreso && currentIngreso) {
        var horas  = (reg.fecha - currentIngreso.fecha) / 3600000;
        var isDiaEst2 = esDiaDeEstudio(empId, currentIngreso.fecha, diasEstudioMapa);
        var tipoLbl, porc;
        if      (isDiaEst2)      { tipoLbl = "Día de Estudio"; porc = 0; }
        else if (reg.esPermiso)  { tipoLbl = "Permiso"; porc = 0; }
        else if (reg.esTerapia || listaTerapias[empId]) { tipoLbl = "Terapia"; porc = 100; }
        else if (reg.esComputacion) { tipoLbl = "Computación"; porc = 50; }
        else                     { tipoLbl = "Normal"; porc = 100; }
        _addFila(currentIngreso.fecha, currentIngreso.fecha, reg.fecha, tipoLbl, horas, porc);
        currentIngreso = null; lastEgreso = reg;
      }
    }

    // Ingreso sin salida al final del set
    if (currentIngreso && currentIngreso.fecha.getHours() < 17) {
      var sEst2 = new Date(currentIngreso.fecha.getTime() + CFG.HORAS_JORNADA_NORMAL * 3600000);
      var isDiaEstF = esDiaDeEstudio(empId, currentIngreso.fecha, diasEstudioMapa);
      _addFila(currentIngreso.fecha, currentIngreso.fecha, sEst2,
        (isDiaEstF ? "Día de Estudio (Est.)" : "Normal (Estimado)") + "*",
        CFG.HORAS_JORNADA_NORMAL, isDiaEstF ? 0 : 100);
    }

    // Subtotal del participante
    hoja.getRange(filaActual,1,1,5).merge()
      .setValue("SUBTOTAL  " + empId).setFontWeight("bold").setHorizontalAlignment("right").setBackground("#f5f5f5");
    hoja.getRange(filaActual,6).setValue(Math.round(statsEmp.horas*100)/100).setFontWeight("bold").setBackground("#f5f5f5").setHorizontalAlignment("center");
    hoja.getRange(filaActual,8).setValue(Math.round(statsEmp.pagar*100)/100).setFontWeight("bold").setBackground("#f5f5f5").setHorizontalAlignment("center");
    hoja.getRange(filaActual,9).setValue(tarifa).setFontWeight("bold").setBackground("#f5f5f5").setHorizontalAlignment("center");
    hoja.getRange(filaActual,10).setValue(Math.round(statsEmp.monto*100)/100).setFontWeight("bold").setBackground("#fff9c4").setHorizontalAlignment("center");
    filaActual++;

    if (statsEmp.terapia > 0) {
      hoja.getRange(filaActual,1,1,NCOLS).merge()
        .setValue("🧘 Total Terapia: " + statsEmp.terapia.toFixed(2) + " hrs")
        .setFontStyle("italic").setFontSize(9).setFontColor("#00796b");
      filaActual++;
    }
    filaActual++;

    totGen.horas  += statsEmp.horas;
    totGen.pagar  += statsEmp.pagar;
    totGen.terapia += statsEmp.terapia;
    totGen.monto  += statsEmp.monto;
  });

  // Resumen general
  filaActual++;
  hoja.getRange(filaActual,1,1,NCOLS).merge()
    .setValue("RESUMEN GENERAL")
    .setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center").setBackground("#cfd8dc");
  filaActual++;

  var tablaRes = [
    ["Total Personas:", listaEmps.length, "Total Horas Lab.:", Math.round(totGen.horas*100)/100],
    ["Total Terapias:", Math.round(totGen.terapia*100)/100, "Total Horas a Pagar:", Math.round(totGen.pagar*100)/100],
    ["", "", "TOTAL MONTO (Q):", Math.round(totGen.monto*100)/100]
  ];
  hoja.getRange(filaActual,1,3,4).setValues(tablaRes);
  hoja.getRange(filaActual,3,3,1).setFontWeight("bold");
  hoja.getRange(filaActual+2,3,1,2).setBackground("#fff9c4").setFontWeight("bold").setFontSize(12);
  filaActual += 5;

  hoja.getRange(filaActual,1,1,NCOLS).merge()
    .setValue("⚠️ * = Salida estimada (" + CFG.HORAS_JORNADA_NORMAL + " hrs) | " +
              "🟢 Verde = Terapia | 🟡 Amarillo = Computación (50%) | 🔴 Rojo = Permiso/Estudio (0%) | 🟣 Morado = Día de Estudio")
    .setFontSize(8).setFontStyle("italic").setFontColor("#d32f2f");

  // Anchos de columna
  hoja.setColumnWidth(1,85); hoja.setColumnWidth(2,85);
  hoja.setColumnWidth(3,70); hoja.setColumnWidth(4,70);
  hoja.setColumnWidth(5,140); hoja.setColumnWidth(6,80);
  hoja.setColumnWidth(7,80); hoja.setColumnWidth(8,80);
  hoja.setColumnWidth(9,80); hoja.setColumnWidth(10,90);

  ss.setActiveSheet(hoja);
  _alert("✅ Reporte generado: " + hoja.getName() +
    "\nParticipantes: " + listaEmps.length +
    "\nHoras a pagar: " + Math.round(totGen.pagar*100)/100 +
    "\nMonto total: Q " + Math.round(totGen.monto*100)/100);
}

// ── Auxiliar: escribir una fila en el reporte ─────────────────

function _escribirFilaReporte(hoja, fila, fechaRef, start, end, tipo, horas, porc, tarifa, monto) {
  var tz = Session.getScriptTimeZone();
  hoja.getRange(fila,1,1,10).setValues([[
    Utilities.formatDate(fechaRef, tz, "dd/MM/yyyy"),
    CFG.DIAS_SEMANA[fechaRef.getDay()],
    Utilities.formatDate(start, tz, "HH:mm"),
    Utilities.formatDate(end, tz, "HH:mm"),
    tipo,
    horas.toFixed(2),
    porc + "%",
    (horas * porc / 100).toFixed(2),
    tarifa.toFixed(2),
    monto.toFixed(2)
  ]]).setHorizontalAlignment("center").setFontSize(9);

  var color = "#ffffff";
  if      (tipo.indexOf("Terapia")  !== -1)                  color = "#e0f2f1";
  else if (tipo.indexOf("Comput")   !== -1)                  color = "#fff9c4";
  else if (tipo.indexOf("Permiso")  !== -1 || porc === 0)    color = "#ffebee";
  else if (tipo.indexOf("Estudio")  !== -1)                  color = "#f3e5f5";
  if (color !== "#ffffff") hoja.getRange(fila,1,1,10).setBackground(color);
}

function _esMismaFecha(d1, d2) {
  return d1.getFullYear()===d2.getFullYear() && d1.getMonth()===d2.getMonth() && d1.getDate()===d2.getDate();
}

// ── Helpers para reportes ─────────────────────────────────────

function _parseFecha(str) {
  var p = str.replace(/\s/g,"").split("/");
  if (p.length < 3) return null;
  var d = parseInt(p[0]), m = parseInt(p[1])-1, y = parseInt(p[2]);
  if (isNaN(d)||isNaN(m)||isNaN(y)) return null;
  var f = new Date(y, m, d, 0, 0, 0, 0);
  return isNaN(f) ? null : f;
}

function _validarEnRango(tipo, fecha, fi, ff) {
  if (tipo === "dia") {
    return fi && fecha.getFullYear()===fi.getFullYear() &&
           fecha.getMonth()===fi.getMonth() && fecha.getDate()===fi.getDate();
  }
  if (tipo === "semana") {
    var finSem = new Date(fi); finSem.setDate(finSem.getDate()+6); finSem.setHours(23,59,59,999);
    return fecha >= fi && fecha <= finSem;
  }
  if (tipo === "mes") {
    return fecha.getFullYear()===fi.getFullYear() && fecha.getMonth()===fi.getMonth();
  }
  if (tipo === "rango") {
    var fc = new Date(fecha); fc.setHours(0,0,0,0);
    return fc >= fi && fc <= ff;
  }
  return true; // todo
}

function _obtenerTextoPeriodo(tipo, fi, ff) {
  var tz = Session.getScriptTimeZone();
  if (tipo === "dia") return Utilities.formatDate(fi, tz, "dd/MM/yyyy");
  if (tipo === "semana") {
    var fin = new Date(fi); fin.setDate(fin.getDate()+6);
    return Utilities.formatDate(fi,tz,"dd/MM/yyyy") + " al " + Utilities.formatDate(fin,tz,"dd/MM/yyyy");
  }
  if (tipo === "mes") return Utilities.formatDate(fi, tz, "MMMM yyyy").toUpperCase();
  if (tipo === "rango") return Utilities.formatDate(fi,tz,"dd/MM/yyyy") + " al " + Utilities.formatDate(ff,tz,"dd/MM/yyyy");
  return "TODOS LOS REGISTROS";
}

function _construirMapaTarifas() {
  var map = {};
  var hojaP = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.HOJAS.PARTICIPANTES);
  if (!hojaP) return map;
  var datos = hojaP.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var nombre = String(datos[i][1]||"").trim();
    if (!nombre) continue;
    var tarifa = parseFloat(datos[i][11]); // col L = Tarifa_Hora
    if (isNaN(tarifa) || tarifa <= 0) {
      var cat = String(datos[i][10]).trim().toUpperCase(); // col K = Categoria
      tarifa = CFG.CATEGORIAS[cat] || CFG.CATEGORIAS.C;
    }
    var t = String(datos[i][12]).trim().toLowerCase(); // col M = Tiene_Factura
    map[nombre] = {
      tarifa:       tarifa,
      categoria:    String(datos[i][10]).trim().toUpperCase(),
      tieneFactura: t === "sí" || t === "si"
    };
  }
  return map;
}

// ══════════════════════════════════════════════════════════════════
// INSTALACIÓN COMPLETA — wizard de 3 pasos
// ══════════════════════════════════════════════════════════════════

function instalarTodo() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert("🚀 INSTALACIÓN COMPLETA — RRHH " + CFG.ORG,
    "Se ejecutarán 6 pasos automáticamente:\n\n" +
    "1 — Crear hojas del sistema (PARTICIPANTES, ASISTENCIA, FACTURACION…)\n" +
    "2 — Importar datos desde Kobo\n" +
    "3 — Crear estructura en Drive (Docs_Proceso, Recibos, Reportes)\n" +
    "4 — Crear hoja Días de Estudio\n" +
    "5 — Crear hoja Lista de Terapias\n" +
    "6 — Activar automatizaciones (Kobo cada hora + al abrir)\n\n" +
    "Los datos existentes NO se borran.\n\n¿Continuar?",
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [], errores = [];

  // PASO 1
  try {
    ss.toast("Paso 1/6: Creando hojas del sistema...", "🚀", -1);
    crearHojas();
    log.push("✅ Paso 1: Hojas del sistema creadas/verificadas");
  } catch(e) { errores.push("❌ Paso 1: " + e.message); }
  Utilities.sleep(500);

  // PASO 2: Importar Kobo
  try {
    ss.toast("Paso 2/6: Importando datos de Kobo...", "🚀", -1);
    var res = UrlFetchApp.fetch(CFG.KOBO_URL_CSV, { muteHttpExceptions: true });
    var code = res.getResponseCode();
    if (code === 200) {
      var datos = Utilities.parseCsv(res.getContentText(), ";");
      if (datos.length > 1) {
        var hK = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO) || ss.insertSheet(CFG.HOJAS.DATOS_KOBO);
        hK.clearContents();
        hK.getRange(1,1,datos.length,datos[0].length).setValues(datos);
        hK.getRange(1,1,1,datos[0].length).setFontWeight("bold").setBackground("#4a86e8").setFontColor("#fff");
        hK.setFrozenRows(1);
        _limpiarColumnasKobo(hK, datos[0]);
        _normalizarAccionSilencioso(hK);
        log.push("✅ Paso 2: " + (datos.length-1) + " registros importados desde Kobo");
      } else { log.push("⚠️ Paso 2: Kobo sin registros"); }
    } else if (code === 503) {
      log.push("⏳ Paso 2: Kobo ocupado (503) — ejecuta manualmente después");
    } else {
      errores.push("❌ Paso 2: Error Kobo HTTP " + code);
    }
  } catch(e) { errores.push("❌ Paso 2: " + e.message); }
  Utilities.sleep(500);

  // PASO 3: Drive
  try {
    ss.toast("Paso 3/6: Creando estructura en Drive...", "🚀", -1);
    crearEstructuraDrive();
    log.push("✅ Paso 3: Estructura Drive creada/verificada");
  } catch(e) { errores.push("❌ Paso 3: " + e.message); }
  Utilities.sleep(500);

  // PASO 4: DiasEstudio
  try {
    ss.toast("Paso 4/6: Configurando Días de Estudio...", "🚀", -1);
    if (!ss.getSheetByName("DiasEstudio")) { crearHojaDiasEstudio(); log.push("✅ Paso 4: Hoja DiasEstudio creada"); }
    else { log.push("ℹ️ Paso 4: Hoja DiasEstudio ya existe"); }
  } catch(e) { errores.push("❌ Paso 4: " + e.message); }
  Utilities.sleep(300);

  // PASO 5: ListaTerapias
  try {
    ss.toast("Paso 5/6: Configurando Lista de Terapias...", "🚀", -1);
    if (!ss.getSheetByName("ListaTerapias")) { crearHojaListaTerapias(); log.push("✅ Paso 5: Hoja ListaTerapias creada"); }
    else { log.push("ℹ️ Paso 5: Hoja ListaTerapias ya existe"); }
  } catch(e) { errores.push("❌ Paso 5: " + e.message); }
  Utilities.sleep(300);

  // PASO 6: Triggers
  try {
    ss.toast("Paso 6/6: Activando automatizaciones...", "🚀", -1);
    configurarTriggers();
    log.push("✅ Paso 6: Triggers activados (cada hora + al abrir)");
  } catch(e) { errores.push("❌ Paso 6: " + e.message); }

  ss.toast("", "", 1);
  var resumen = "🚀 INSTALACIÓN COMPLETA — " + CFG.ORG + "\n\n";
  resumen += log.join("\n");
  if (errores.length) resumen += "\n\n--- PROBLEMAS ---\n" + errores.join("\n");
  resumen += "\n\n--- HOJAS DEL SISTEMA ---\n";
  resumen += "• PARTICIPANTES — agrega personas aquí (Categoría A/B/C/D)\n";
  resumen += "• ASISTENCIA — horas calculadas al emparejar\n";
  resumen += "• FACTURACION — pagos por quincena\n";
  resumen += "• DatosKobo — datos crudos de Kobo\n";
  resumen += "• DiasEstudio — marca días que no pagan\n";
  resumen += "• ListaTerapias — marca quién va a terapia\n";
  resumen += "\nTarifas: A=Q16.50 | B=Q15.75 | C=Q15.00 | D=Q14.00\n";
  resumen += "IVA 5%: solo participantes con Tiene_Factura=Sí\n";
  resumen += "\nYa puedes generar reportes desde el menú 📊";
  _alert(resumen);
}); }

// ══════════════════════════════════════════════════════════════════
// HOJAS DE APOYO — DiasEstudio y ListaTerapias
// ══════════════════════════════════════════════════════════════════

function crearHojaDiasEstudio() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("DiasEstudio");
  if (!hoja) {
    hoja = ss.insertSheet("DiasEstudio");
    var enc = ["Participante","Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo","Fecha_Inicio","Fecha_Fin"];
    hoja.getRange(1,1,1,enc.length).setValues([enc])
      .setFontWeight("bold").setBackground("#7b1fa2").setFontColor("#fff").setHorizontalAlignment("center");
    hoja.setFrozenRows(1);

    // Pre-llenar con participantes desde DatosKobo
    var hojaK = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
    if (hojaK) {
      var datos = hojaK.getDataRange().getValues();
      var colsP = _buscarColsParticipante(datos[0]);
      if (colsP.col1 !== -1) {
        var mapeo = cargarMapeoNombres();
        var partics = {};
        for (var f=1; f<datos.length; f++) {
          var n = _nombreDeFila(datos[f], colsP);
          if (n) partics[normalizarNombre(n, mapeo)] = true;
        }
        var lista = Object.keys(partics).sort();
        for (var p=0; p<lista.length; p++) hoja.getRange(p+2,1).setValue(lista[p]);
      }
    }

    var vXO = SpreadsheetApp.newDataValidation().requireValueInList(["X",""],true).build();
    hoja.getRange(2,2,200,7).setDataValidation(vXO).setHorizontalAlignment("center");
    hoja.setColumnWidth(1,220); hoja.setColumnWidth(9,100); hoja.setColumnWidth(10,100);
    for (var c=2;c<=8;c++) hoja.setColumnWidth(c,80);
  }
  hoja.activate();
  _alert("📚 HOJA DÍAS DE ESTUDIO\n\nMarca con X los días que cada participante estudia (0% de pago).\n\n" +
    "Columnas I y J: Fecha_Inicio y Fecha_Fin (opcionales) para limitar el período de vigencia.");
}); }

function crearHojaListaTerapias() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("ListaTerapias");
  if (!hoja) {
    hoja = ss.insertSheet("ListaTerapias");
    hoja.getRange(1,1,1,3).setValues([["Participante","Recibe Terapia (X)","Notas"]])
      .setFontWeight("bold").setBackground("#00897b").setFontColor("#fff").setHorizontalAlignment("center");
    hoja.setFrozenRows(1);

    // Pre-llenar con participantes desde DatosKobo
    var hojaK = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
    if (hojaK) {
      var datos = hojaK.getDataRange().getValues();
      var colsP = _buscarColsParticipante(datos[0]);
      if (colsP.col1 !== -1) {
        var mapeo = cargarMapeoNombres();
        var partics = {};
        for (var f=1; f<datos.length; f++) {
          var n = _nombreDeFila(datos[f], colsP);
          if (n) partics[normalizarNombre(n, mapeo)] = true;
        }
        var lista = Object.keys(partics).sort();
        for (var p=0; p<lista.length; p++) hoja.getRange(p+2,1).setValue(lista[p]);
      }
    }

    var vX = SpreadsheetApp.newDataValidation().requireValueInList(["X",""],true).build();
    hoja.getRange(2,2,200,1).setDataValidation(vX).setHorizontalAlignment("center");
    hoja.setColumnWidth(1,250); hoja.setColumnWidth(2,140); hoja.setColumnWidth(3,300);
  }
  hoja.activate();
  _alert("🧘 HOJA LISTA DE TERAPIAS\n\nMarca con X a las personas que reciben terapia.\nLos registros de tipo 'Terapia' se contabilizan al 100% del pago.");
}); }

// ══════════════════════════════════════════════════════════════════
// DIAGNÓSTICO, REPARACIÓN Y CAMBIO DE NOMBRE
// ══════════════════════════════════════════════════════════════════

function diagnosticarDatosKobo() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaK = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (!hojaK) { _alert("No existe DatosKobo. Importa primero desde Kobo."); return; }

  var datos = hojaK.getDataRange().getValues();
  var enc = datos[0];
  var cols = detectarColumnas(enc, datos.slice(1));

  var listaEnc = "";
  for (var i=0; i<enc.length; i++) listaEnc += "Col "+i+": \""+enc[i]+"\"\n";

  var det = "\n--- COLUMNAS DETECTADAS ---\n";
  det += "start: "         + (cols.start          !== undefined ? "Col "+cols.start : "❌ NO ENCONTRADA") + "\n";
  det += "end: "           + (cols.end             !== undefined ? "Col "+cols.end   : "no (opcional)")  + "\n";
  det += "participante: "  + (cols.participante    !== undefined ? "Col "+cols.participante+" (\""+enc[cols.participante]+"\")" : "❌ NO ENCONTRADA") + "\n";
  if (cols.participante2 !== undefined) det += "participante2: Col "+cols.participante2+" (\""+enc[cols.participante2]+"\")\n";
  det += "acción: "        + (cols.accionUnificada !== undefined ? "Col "+cols.accionUnificada+" (\""+enc[cols.accionUnificada]+"\")": "❌ NO ENCONTRADA") + "\n";
  det += "uuid: "          + (cols.uuid            !== undefined ? "Col "+cols.uuid : "no") + "\n";

  var conteo = { entrada:0, salida:0, sinTipo:0, terapia:0, computacion:0, permiso:0 };
  for (var f=1; f<datos.length; f++) {
    var t = obtenerTipoRegistro(datos[f], cols);
    if (t.esIngreso) conteo.entrada++;
    else if (t.esEgreso) conteo.salida++;
    else conteo.sinTipo++;
    if (t.esTerapia) conteo.terapia++;
    if (t.esComputacion) conteo.computacion++;
    if (t.esPermiso) conteo.permiso++;
  }

  var resumen = "\n--- RESUMEN ---\n";
  resumen += "Total filas: "+(datos.length-1)+"\n🟢 Entradas: "+conteo.entrada+"\n🔴 Salidas: "+conteo.salida;
  resumen += "\n🧘 Terapias: "+conteo.terapia+"\n💻 Computación: "+conteo.computacion+"\n📝 Permisos: "+conteo.permiso;
  if (conteo.sinTipo>0) resumen += "\n⚠️ Sin tipo: "+conteo.sinTipo;
  resumen += (conteo.entrada===0 && conteo.salida===0)
    ? "\n\n❌ PROBLEMA: No se detectaron entradas ni salidas.\nRevisa la columna de acción en DatosKobo."
    : "\n\n✅ Datos listos para generar reportes.";

  SpreadsheetApp.getUi().alert(listaEnc + det + resumen);
}); }

function repararDatosKobo() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaK = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (!hojaK) { _alert("No existe DatosKobo."); return; }

  var datos = hojaK.getDataRange().getValues();
  var cols = detectarColumnas(datos[0], datos.slice(1));
  var mapeoN = cargarMapeoNombres();

  // Fase 1: detectar problemas
  var desconocidos = {}, nombresANorm = 0;
  if (cols.accionUnificada !== undefined) {
    for (var f=1; f<datos.length; f++) {
      var t = obtenerTipoRegistro(datos[f], cols);
      if (!t.esIngreso && !t.esEgreso) {
        var v = String(datos[f][cols.accionUnificada]||"").trim();
        if (v) desconocidos[v] = (desconocidos[v]||0)+1;
      }
    }
  }
  if (Object.keys(mapeoN).length && cols.participante !== undefined) {
    for (var f=1; f<datos.length; f++) {
      var n = String(datos[f][cols.participante]||"").trim();
      if (n && mapeoN[n] && mapeoN[n]!==n) nombresANorm++;
    }
  }

  var CONOCIDOS = [
    {k:"ingreso",c:"🟢 Entrada"},{k:"entrada",c:"🟢 Entrada"},
    {k:"egreso",c:"🔴 Salida"},{k:"salida",c:"🔴 Salida"},
    {k:"terapia",c:"🔴 Salida"},{k:"permiso",c:"🔴 Salida"},{k:"comput",c:"🔴 Salida"}
  ];
  var correcciones = {};
  Object.keys(desconocidos).forEach(function(v) {
    var vL = v.toLowerCase();
    for (var k=0; k<CONOCIDOS.length; k++) {
      if (vL.indexOf(CONOCIDOS[k].k.substring(0,4)) !== -1) { correcciones[v]=CONOCIDOS[k].c; break; }
    }
  });

  var msg = "🔧 REPARAR DATOS KOBO\n\n";
  if (!Object.keys(desconocidos).length) msg += "✅ Sin valores desconocidos en Entrada/Salida.\n";
  else {
    msg += "⚠️ Valores no reconocidos:\n";
    Object.keys(desconocidos).forEach(function(v){ msg += "• \""+v+"\" ("+desconocidos[v]+"x)"+( correcciones[v]?" → "+correcciones[v]:" → sin corrección")+"\n"; });
  }
  msg += nombresANorm>0 ? "\n📋 Nombres a normalizar en DatosKobo: "+nombresANorm+" celdas\n" : "\n✅ Nombres ya normalizados.\n";

  if (!Object.keys(desconocidos).length && !nombresANorm) { ui.alert(msg+"\nNo hay nada que reparar."); return; }

  if (ui.alert("🔧 Reparar Datos", msg+"\n¿Aplicar correcciones?", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  var cam1=0, cam2=0;
  if (cols.accionUnificada !== undefined) {
    for (var f=1; f<datos.length; f++) {
      var v = String(datos[f][cols.accionUnificada]||"").trim();
      if (v && correcciones[v]) { hojaK.getRange(f+1,cols.accionUnificada+1).setValue(correcciones[v]); cam1++; }
    }
  }
  datos = hojaK.getDataRange().getValues();
  if (Object.keys(mapeoN).length && cols.participante !== undefined) {
    for (var f=1; f<datos.length; f++) {
      var n = String(datos[f][cols.participante]||"").trim();
      if (n && mapeoN[n] && mapeoN[n]!==n) { hojaK.getRange(f+1,cols.participante+1).setValue(mapeoN[n]); cam2++; }
    }
  }
  ui.alert("✅ REPARACIÓN COMPLETADA\n\nEntrada/Salida corregidos: "+cam1+"\nNombres normalizados: "+cam2);
}); }

function cambiarNombreParticipante() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hNC = ss.getSheetByName("NombresCanonicos");
  if (!hNC) { ui.alert("No existe NombresCanonicos.\nEjecuta 'Normalizar nombres' primero."); return; }

  var datosNC = hNC.getDataRange().getValues();
  var canonicos = {};
  for (var f=1; f<datosNC.length; f++) {
    var c = String(datosNC[f][1]||"").trim(); if (c) canonicos[c]=true;
  }
  var lista = Object.keys(canonicos).sort();

  var r1 = ui.prompt("✏️ Cambiar Nombre (1/2)",
    "Participantes actuales:\n"+lista.slice(0,20).join("\n")+(lista.length>20?"\n... ("+(lista.length-20)+" más)":"")+
    "\n\nEscribe el nombre ACTUAL (exacto):", ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var nombreActual = r1.getResponseText().trim();
  if (!canonicos[nombreActual]) {
    var sug = lista.filter(function(n){ return n.toLowerCase().indexOf(nombreActual.toLowerCase())!==-1; });
    ui.alert("No se encontró \""+nombreActual+"\"."+(sug.length?" Sugerencias:\n"+sug.slice(0,5).join("\n"):""));
    return;
  }

  var r2 = ui.prompt("✏️ Cambiar Nombre (2/2)",
    "Nombre actual:\n\""+nombreActual+"\"\n\n¿Nuevo nombre? (formato: Nombre Completo (CÓDIGO))", ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var nombreNuevo = r2.getResponseText().trim();
  if (!nombreNuevo || nombreNuevo===nombreActual) return;

  var cam1=0, cam2=0, cam3=0;
  for (var f=1; f<datosNC.length; f++) {
    if (String(datosNC[f][1]||"").trim()===nombreActual) { hNC.getRange(f+1,2).setValue(nombreNuevo); cam1++; }
  }
  ["DiasEstudio","ListaTerapias"].forEach(function(nm) {
    var h=ss.getSheetByName(nm); if(!h)return;
    var d=h.getDataRange().getValues();
    for (var f=1; f<d.length; f++) {
      if (String(d[f][0]||"").trim()===nombreActual) {
        h.getRange(f+1,1).setValue(nombreNuevo);
        nm==="DiasEstudio"?cam2++:cam3++;
      }
    }
  });

  ui.alert("✅ NOMBRE CAMBIADO\n\""+nombreActual+"\"\n→ \""+nombreNuevo+"\"\n\n"+
    "NombresCanonicos: "+cam1+" fila(s)\n"+
    (cam2?"DiasEstudio: "+cam2+" fila(s)\n":"")+
    (cam3?"ListaTerapias: "+cam3+" fila(s)\n":""));
}); }
