// ============================================================
// SISTEMA RRHH — mi eelo
// ============================================================

const CFG = {
  ORG:          "mi eelo",   // programa
  PROYECTO:     "Textil",    // proyecto padre
  CORREO_ADMIN: "adrian@creamosguatemala.org",
  // Tarifas por categoría (Q por hora) — A=Q16.50 B=Q15.75 C=Q15.00 D=Q14.00
  CATEGORIAS:   { A: 16.50, B: 15.75, C: 15.00, D: 14.00 },
  // Colores oficiales por categoría — usados en TODOS los reportes y hojas
  COLORES_CAT: {
    A: { bg: "#1e7e34", fg: "#ffffff", bgClaro: "#d4edda" }, // verde oscuro / verde claro
    B: { bg: "#1565c0", fg: "#ffffff", bgClaro: "#bbdefb" }, // azul oscuro / azul claro
    C: { bg: "#e65100", fg: "#ffffff", bgClaro: "#ffe0b2" }, // naranja oscuro / naranja claro
    D: { bg: "#6a1a6a", fg: "#ffffff", bgClaro: "#f3e5f5" }  // morado oscuro / morado claro
  },
  IVA_PCT:      0.05,  // 5% Pequeño Contribuyente Guatemala (solo quien tiene factura)
  HORAS_JORNADA_NORMAL: 7,
  KOBO_URL_CSV: "https://kf.kobotoolbox.org/api/v2/assets/agi395bJj6ojXJzPPDT9n6/export-settings/es4oUjEmPvovgLd6Y5yrQ4K/data.csv",
  KOBO_TIPO_ENTRADA: "🟢 Entrada",   // valor normalizado interno (no el label de Kobo)
  KOBO_TIPO_SALIDA:  "🔴 Salida",    // Kobo exporta "Entrada"/"Salida" — detectarColumnas detecta ambos
  HOJAS: {
    PARTICIPANTES: "PARTICIPANTES",
    ASISTENCIA:    "ASISTENCIA",
    FACTURACION:   "FACTURACION",
    DASHBOARD:     "DASHBOARD",
    DATOS_KOBO:    "DatosKobo",
    CLASIFICACION: "CLASIFICACION",  // mantenida solo para borrado en reinstalar
    PERIODOS:      "PERIODOS",
    CREAMOS_DB:    "Copy of CREAMOS ID nuevo",  // BASE DE DATOS OFICIAL — SOLO LECTURA
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

  // ══════════════════════════════════════════════════════════
  // BLOQUE 1: ASISTENCIA
  // ══════════════════════════════════════════════════════════
  var menuAsistencia = ui.createMenu("📥 Asistencia")
    .addItem("📥 Importar desde Kobo",                   "importarDesdeKobo")
    .addSeparator()
    .addItem("🔍 Diagnosticar registros Kobo",           "diagnosticarDatosKobo")
    .addItem("🔧 Reparar datos Kobo",                    "repararDatosKobo");

  // ══════════════════════════════════════════════════════════
  // BLOQUE 2: QUINCENA
  // ══════════════════════════════════════════════════════════
  var menuQuincena = ui.createMenu("📅 Quincena")
    .addItem("📊 Ver / actualizar quincena actual",      "verQuincenaActual")
    .addItem("✅ Cerrar y abrir siguiente",               "cerrarQuincenaYCrearSiguiente")
    .addSeparator()
    .addItem("🗓️ Nueva quincena (manual)",               "configurarNuevaQuincena")
    .addItem("🔍 Reporte de quincena pasada",            "generarReporteQuincenaPasada")
    .addItem("📊 Reporte por rango de fechas",           "generarReportePorRango");

  // ══════════════════════════════════════════════════════════
  // BLOQUE 3: ADMIN (uso ocasional)
  // ══════════════════════════════════════════════════════════
  var menuAdmin = ui.createMenu("⚙️ Admin")
    .addItem("📋 Cargar lista oficial (33 participantes)","cargarListaParticipantes")
    .addItem("🔄 Sincronizar desde Creamos DB",          "sincronizarDesdeCreamos")
    .addItem("➕ Nuevo participante",                     "nuevoParticipante")
    .addItem("👥 Directorio de participantes",           "generarDirectorioParticipantes")
    .addSeparator()
    .addItem("📚 Días de estudio",                       "crearHojaDiasEstudio")
    .addItem("🧘 Lista de terapias",                     "crearHojaListaTerapias")
    .addItem("🔗 Sincronizar participación (→ PARTICIPANTES)", "sincronizarParticipacion")
    .addSeparator()
    .addItem("🧾 Configurar IVA (quién tiene factura)",  "configurarFacturacion")
    .addItem("✅ Guardar cambios de facturación",         "aplicarCambiosFacturacion")
    .addSeparator()
    .addItem("✏️ Cambiar nombre de participante",        "cambiarNombreParticipante")
    .addItem("📄 Generar DP (fila activa)",              "generarDpFilaActiva")
    .addItem("📄 Actualizar todos los DPs",              "actualizarTodosLosDps")
    .addSeparator()
    .addItem("⚡ Activar automatizaciones",              "configurarTriggers")
    .addItem("🗑️ Reinstalar sistema (borra TODO)",       "reinstalarSistema");

  // ══════════════════════════════════════════════════════════
  // MENÚ PRINCIPAL
  // ══════════════════════════════════════════════════════════
  ui.createMenu("👥 RRHH")
    .addItem("🚀 Instalación completa",                  "instalarTodo")
    .addSeparator()
    .addSubMenu(menuAsistencia)
    .addSubMenu(menuQuincena)
    .addSubMenu(menuAdmin)
    .addToUi();
}

// onEdit: col K = Categoria → auto-llenar Tarifa_Hora en PARTICIPANTES
//         DiasEstudio cols B-H → Educacion en PARTICIPANTES
//         ListaTerapias col C  → Apoyo_Emocional en PARTICIPANTES
function onEdit(e) {
  var sheet = e.range.getSheet();
  var nombre = sheet.getName();
  var col    = e.range.getColumn();
  var fila   = e.range.getRow();

  // PARTICIPANTES — Categoria cambia → auto-llenar Tarifa_Hora
  if (nombre === CFG.HOJAS.PARTICIPANTES && col === 11 && fila >= 2) {
    var cat = String(e.range.getValue()).trim().toUpperCase();
    var tarifa = CFG.CATEGORIAS[cat];
    if (tarifa) sheet.getRange(fila, 12).setValue(tarifa);
  }

  // DIASESTUDIO — cualquier día (cols B-H = 2-8) → Educacion en PARTICIPANTES
  if (nombre === "DiasEstudio" && col >= 2 && col <= 8 && fila >= 2) {
    try { _syncEducacionFila(sheet, fila); } catch(_) {}
  }

  // LISTATERAPIAS — col C (Recibe Terapia) → Apoyo_Emocional en PARTICIPANTES
  if (nombre === "ListaTerapias" && col === 3 && fila >= 2) {
    try { _syncApoyoFila(sheet, fila); } catch(_) {}
  }
}

/** Sincroniza la fila fila de DiasEstudio hacia col H (Educacion) de PARTICIPANTES */
function _syncEducacionFila(hDE, fila) {
  var fila_ = hDE.getRange(fila, 1, 1, 8).getValues()[0];
  var participante = String(fila_[0] || "").trim();
  if (!participante) return;

  var DIAS_NOM = ["Lun","Mar","Mié","Jue","Vie","Sáb","Dom"];
  var activos = [];
  for (var d = 0; d < 7; d++) {
    if (String(fila_[d + 1] || "").trim().toUpperCase() === "X") activos.push(DIAS_NOM[d]);
  }
  var texto = activos.length > 0 ? "Sí — " + activos.join(", ") : "No";
  _actualizarColParticipante(participante, 8, texto); // col H = index 8
}

/** Sincroniza la fila fila de ListaTerapias hacia col I (Apoyo_Emocional) de PARTICIPANTES */
function _syncApoyoFila(hLT, fila) {
  var fila_ = hLT.getRange(fila, 1, 1, 3).getValues()[0];
  var participante = String(fila_[1] || "").trim(); // col B = Participante
  if (!participante) return;
  var texto = String(fila_[2] || "").trim().toUpperCase() === "X" ? "Sí" : "No";
  _actualizarColParticipante(participante, 9, texto); // col I = index 9
}

/** Actualiza la celda de colNum (1-based) para el participante con ese nombre en PARTICIPANTES */
function _actualizarColParticipante(nombreBuscar, colNum, valor) {
  var hP = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.HOJAS.PARTICIPANTES);
  if (!hP || hP.getLastRow() < 2) return;
  var nombres = hP.getRange(2, 2, hP.getLastRow() - 1, 1).getValues();
  var normBuscar = textoParaComparar(nombreBuscar);
  for (var i = 0; i < nombres.length; i++) {
    if (textoParaComparar(String(nombres[i][0])) === normBuscar) {
      hP.getRange(i + 2, colNum).setValue(valor);
      return;
    }
  }
}

/**
 * Sincroniza TODA la participación (DiasEstudio + ListaTerapias → PARTICIPANTES).
 * Llámalo desde el menú después de llenar las hojas.
 */
function sincronizarParticipacion() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hDE = ss.getSheetByName("DiasEstudio");
  var hLT = ss.getSheetByName("ListaTerapias");
  var hP  = _sh(CFG.HOJAS.PARTICIPANTES);
  if (!hP || hP.getLastRow() < 2) { _alert("No hay participantes cargados."); return; }

  var DIAS_NOM = ["Lun","Mar","Mié","Jue","Vie","Sáb","Dom"];
  var actDE = 0, actLT = 0;

  // DiasEstudio → Educacion (col H = 8)
  if (hDE && hDE.getLastRow() >= 2) {
    var datDE = hDE.getRange(2, 1, hDE.getLastRow() - 1, 8).getValues();
    datDE.forEach(function(f) {
      var participante = String(f[0] || "").trim();
      if (!participante) return;
      var activos = [];
      for (var d = 0; d < 7; d++) {
        if (String(f[d + 1] || "").trim().toUpperCase() === "X") activos.push(DIAS_NOM[d]);
      }
      var texto = activos.length > 0 ? "Sí — " + activos.join(", ") : "No";
      _actualizarColParticipante(participante, 8, texto);
      if (activos.length) actDE++;
    });
  }

  // ListaTerapias → Apoyo_Emocional (col I = 9)
  if (hLT && hLT.getLastRow() >= 2) {
    var datLT = hLT.getRange(2, 1, hLT.getLastRow() - 1, 3).getValues();
    datLT.forEach(function(f) {
      var participante = String(f[1] || "").trim(); // col B
      if (!participante) return;
      var texto = String(f[2] || "").trim().toUpperCase() === "X" ? "Sí" : "No";
      _actualizarColParticipante(participante, 9, texto);
      if (texto === "Sí") actLT++;
    });
  }

  _alert(
    "✅ Participación sincronizada\n\n" +
    "📚 Educación (días de estudio): " + actDE + " con asistencia marcada\n" +
    "🧘 Apoyo Emocional (terapias): " + actLT + " con terapia marcada"
  );
}); }

// ── Crear hojas ───────────────────────────────────────────────

function crearHojas() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // (CLASIFICACION eliminada — la info de categorías está en el Directorio)

  // PARTICIPANTES — 21 cols (A–U)
  var hP = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES) || ss.insertSheet(CFG.HOJAS.PARTICIPANTES);
  var esNuevaP = hP.getLastRow() === 0;
  if (esNuevaP) {
    hP.appendRow([
      "Creamos_ID","Nombre","Proyecto","Division","Programa","Estado","Etapa",
      "Educacion","Apoyo_Emocional","Inclusion_Laboral",
      "Categoria","Tarifa_Hora","Tiene_Factura",
      "DPI","NIT","Correo",
      "Banco","Tipo_Cuenta","Num_Cuenta","Forma_Pago","URL_Doc_Proceso"
    ]);
  }
  _fmtEnc(hP, "#639922");
  var vEstado  = SpreadsheetApp.newDataValidation().requireValueInList(["Activo","Inactivo","Egresado"],true).build();
  var vCat     = SpreadsheetApp.newDataValidation().requireValueInList(["A","B","C","D"],true).build();
  var vSiNo    = SpreadsheetApp.newDataValidation().requireValueInList(["Sí","No"],true).build();
  var vBanco   = SpreadsheetApp.newDataValidation().requireValueInList([
    "Banrural","G&T Continental","BAC Credomatic","Industrial","Agromercantil",
    "Occidente","Promerica","Vivibanco","Bantrab","CHN","Otro"
  ],true).build();
  var vTipoCta = SpreadsheetApp.newDataValidation().requireValueInList(["Monetaria","Ahorro",""],true).build();
  var vPago    = SpreadsheetApp.newDataValidation().requireValueInList(["Transferencia","Efectivo","Cheque"],true).build();
  hP.getRange("F2:F500").setDataValidation(vEstado);
  hP.getRange("K2:K500").setDataValidation(vCat);
  hP.getRange("M2:M500").setDataValidation(vSiNo);
  hP.getRange("Q2:Q500").setDataValidation(vBanco);
  hP.getRange("R2:R500").setDataValidation(vTipoCta);
  hP.getRange("T2:T500").setDataValidation(vPago);
  hP.getRange("L2:L500").setNumberFormat("Q#,##0.00");
  if (esNuevaP) {
    hP.setColumnWidth(2, 220);  // Nombre
    hP.setColumnWidth(17, 120); // Banco
    hP.setColumnWidth(18, 100); // Tipo_Cuenta
    hP.setColumnWidth(19, 130); // Num_Cuenta
    hP.setColumnWidth(20, 120); // Forma_Pago
    hP.setColumnWidth(21, 300); // URL_Doc_Proceso
  }

  // PERIODOS — hoja de control de quincenas
  crearHojaPeriodos();

  ["Hoja 1","Sheet1"].forEach(function(n) {
    var h = ss.getSheetByName(n);
    if (h && ss.getSheets().length > 3) ss.deleteSheet(h);
  });

  // Proteger y ocultar la base de datos oficial de Creamos (si existe)
  _protegerHojaCreamos_DB();

  ss.setActiveSheet(hP);
  _alert(
    "✅ Hojas creadas:\n" +
    "• PARTICIPANTES (21 cols — incluye Banco, Tipo_Cuenta, Num_Cuenta, Forma_Pago)\n" +
    "• DatosKobo (importación automática cada hora desde Kobo)\n" +
    "• PERIODOS (control de quincenas)\n\n" +
    "Tarifas: A=Q16.50 | B=Q15.75 | C=Q15.00 | D=Q14.00\n" +
    "IVA 5%: solo participantes con Tiene_Factura=Sí (Admin → Configurar IVA)"
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

  var r2 = ui.prompt("➕ Nuevo participante", "Nombre completo (tal como aparece en el sistema):", ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var nombre = r2.getResponseText().trim();
  if (!nombre) { ui.alert("El nombre no puede estar vacío."); return; }

  var hP = _sh(CFG.HOJAS.PARTICIPANTES);
  var dp = hP.getDataRange().getValues();

  // Verificar duplicado por nombre
  for (var i = 1; i < dp.length; i++) {
    if (limpiarNombre(String(dp[i][1])).toLowerCase() === limpiarNombre(nombre).toLowerCase()) {
      ui.alert("Ya existe un participante con ese nombre:\n" + dp[i][1]); return;
    }
  }

  // Buscar en base de datos oficial de Creamos
  SpreadsheetApp.getActiveSpreadsheet().toast("Buscando en base de datos Creamos...", "⏳", -1);
  var db  = _buscarEnCreamos_DB(nombre);
  var id  = db.id  || "";
  var dpi = db.dpi || "";

  if (!id) {
    var resp2 = ui.alert(
      "⚠️ No encontrado en base Creamos",
      "No se encontró '" + nombre + "' en 'Copy of CREAMOS ID nuevo'.\n\n" +
      "¿Deseas continuar sin Creamos ID?\n" +
      "(El ID quedará vacío — podrás llenarlo después en col A)",
      ui.ButtonSet.YES_NO
    );
    if (resp2 !== ui.Button.YES) return;
  }

  var carpeta = _carpetaDP();
  var doc = _abrirOCrearDocProceso(id || nombre, nombre, carpeta, "");
  var url = doc.getUrl();
  var part = { id:id, nombre:nombre, proyecto:CFG.PROYECTO, division:"", programa:CFG.ORG, estado:"Activo",
               etapa:"", educacion:"", apoyoEmocional:"", inclusionLaboral:"",
               categoria:"", tarifa:"", tieneFactura:"No",
               dpi:dpi, nit:"", correo:"", banco:"", numCuenta:"", formaPago:"" };
  _escribirContenidoDP(doc, part, [], []);

  hP.appendRow([id, nombre, CFG.PROYECTO, "", CFG.ORG, "Activo", "", "", "", "", "", "", "No", dpi, "", "", "", "", "", "", url]);
  hP.setActiveRange(hP.getRange(hP.getLastRow(), 1));
  ui.alert(
    "✅ Participante registrado: " + nombre + "\n\n" +
    "Creamos ID: " + (id || "— (llenar manualmente)") + "\n" +
    "DPI: "        + (dpi || "— (no encontrado)") + "\n\n" +
    "Documento de Proceso:\n" + url
  );
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
// Formato: [nombre_oficial, categoria, creamos_id, kobo_slug]
// Orden oficial del taller (1-33). creamos_id="" = se auto-genera.
// kobo_slug: valor exacto que usa KoboToolbox en el campo Participante.
var LISTA_OFICIAL = [
  //  #   Nombre oficial                           Cat  Creamos_ID    Kobo slug
  [ 1, "Sindy Paola Lazaro Diaz",            "B", "SILA161192", "sindy_paola_lazaro_diaz"],
  [ 2, "Rosaura Jeannette Saquic Lopez",     "C", "ROSA301178", "rosaura_jeannette_saquic_lopez"],
  [ 3, "Maria del Carmen Borrayo Vásquez",   "C", "MABO280871", "maria_del_carmen_borrayo_vasquez"],
  [ 4, "Maria Audelia Velasquez Cabrera",    "C", "MAVE130766", "maria_audelia_velasquez_cabrera"],
  [ 5, "María Ricarda Suret Chamalé",        "C", "MASU150784", "maria_ricarda_suret_chamale"],
  [ 6, "Mirna Leticia Rodriguez Paniagua",   "D", "MIRO080964", "mirna_leticia_rodriguez_paniagua"],
  [ 7, "Sindy Sucely Veliz Vian",            "B", "SIVE150196", "sindy_sucely_veliz_vian"],
  [ 8, "Yoselin Melissa Zurdo Tocay",        "C", "YOZU230190", "yoselin_melissa_zurdo_tocay"],
  [ 9, "Erika Vásquez Tocay de López",       "C", "ERVA031282", "ericka_vasquez_tocay"],
  [10, "Leticia Sumalé Arredondo",           "C", "LESU210172", "leticia_sumale_arredondo"],
  [11, "Alicia Lopez Reynoso",               "D", "",           "alicia_lopez_reynoso"],
  [12, "Otilia Turuy Paz",                   "D", "",           "otilia_turuy_paz"],
  [13, "Angelica Casandra Veliz Vián",       "A", "ANVE241097", "angelica_casandra_veliz_vian"],
  [14, "Karin Nineth Balcarcel Santizo",     "B", "KABA221273", "karin_nineth_balcarcel_santizo"],
  [15, "Lorena del Rosario Urrea",           "C", "",           "lorena_del_rosario_urrea"],
  [16, "Brenda Azucena del Cid Urrea",       "D", "",           "brenda_azucena_del_cid_urrea"],
  [17, "Vilma Elizabeth Lopez Vasquez",      "C", "VILO040971", "vilma_elizabeth_lopez_vasquez"],
  [18, "Angélica Maribel Cuxe Pérez",        "D", "ANCU300380", "angelica_maribel_cuxe_perez"],
  [19, "Sara Evilia Raymundo Rivera",        "B", "SARA010779", "sara_evilia_raymundo_rivera"],
  [20, "Helen Melany Rodas López",           "C", "HERO171201", "helen_melany_rodas_lopez"],
  [21, "Elendi Nicol Pedroza Cuxé",          "D", "",           "elendi_nicol_pedroza_cuxe"],
  [22, "Emily Cristina Zacarías Morales",    "B", "EMZA021099", "emily_cristina_zacarias_morales"],
  [23, "Juana del Rosario Vicente Choy",     "B", "",           "juana_del_rosario_vicente_choy"],
  [24, "Mayra Lorena Cifuentes García",      "C", "MACI030373", "mayra_lorena_cifuentes_garcia"],
  [25, "Ana Rebeca Larios Perez",            "D", "ANLA060686", "ana_rebeca_larios_perez"],
  [26, "Jeimy Suceli Barrientos",            "D", "JEBA011090", "jeimy_suceli_barrientos"],
  [27, "María Aidé Alvarado Cortéz",         "D", "MAAL070490", "maria_aide_alvarado_cortez"],
  [28, "Yocelin Yajaira Celada Rodriguez",   "C", "YOCE041291", "yocelin_yajaira_celada_rodriguez"],
  [29, "Ruth Saraí Pivaral Sequen",          "C", "RUPI170992", "ruth_sarai_pivaral_sequen"],
  [30, "Laura Elizabeth Gonzalez Figueroa",  "C", "LAGO091289", "laura_elizabeth_gonzalez_figueroa"],
  [31, "Sandra Aracely Vicente Cortéz",      "C", "",           "sandra_aracely_vicente_cortez"],
  [32, "Heidy Yessenía Morales Lázaro",      "C", "",           "heidy_yessenia_morales_lazaro"],
  [33, "Anaid Lluleydi Mateo Morales",       "C", "ANMA100605", "anaid_lluleydi_mateo_morales"]
];

/*
 * Carga la lista oficial en PARTICIPANTES.
 * Limpia duplicados y genera Creamos_ID automático (formato XXXX001).
 */
function cargarListaParticipantes() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    "📋 Cargar lista oficial",
    "Se borrarán todas las filas actuales de PARTICIPANTES y se cargará la lista\noficial limpia (29 participantes) con IDs generados automáticamente.\n\n¿Continuar?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var hP = _sh(CFG.HOJAS.PARTICIPANTES);

  // Borrar datos anteriores (preservar encabezado)
  var lastRow = hP.getLastRow();
  if (lastRow > 1) hP.getRange(2, 1, lastRow - 1, 20).clearContent();

  // Construir filas buscando cada participante en la DB de Creamos
  SpreadsheetApp.getActiveSpreadsheet().toast("Buscando en base de datos Creamos...", "⏳", -1);
  var filas      = [];
  var encontrados = 0, sinEncontrar = [];

  LISTA_OFICIAL.forEach(function(item) {
    var nombre = item[1], cat = item[2];
    var db     = _buscarEnCreamos_DB(nombre);  // busca ID, DPI, género, fecha nac...
    var id     = db.id || "";                  // ID viene SOLO de la DB, nunca se genera
    var dpi    = db.dpi || "";
    var tarifa = CFG.CATEGORIAS[cat] || 0;

    if (id) encontrados++;
    else    sinEncontrar.push(item[0] + ". " + nombre);

    filas.push([
      id,            // A: Creamos_ID   ← de la DB oficial
      nombre,        // B: Nombre
      CFG.PROYECTO,  // C: Proyecto = "Textil"
      "", CFG.ORG,   // D: Division, E: Programa = "mi eelo"
      "Activo", // F: Estado
      "", "", "", "", // G–J: Etapa, Educacion, Apoyo_Emocional, Inclusion_Laboral
      cat,      // K: Categoria
      tarifa,   // L: Tarifa_Hora
      "No",     // M: Tiene_Factura
      dpi,      // N: DPI          ← de la DB oficial
      "", "", "", "", "", "", ""  // O–U: NIT, Correo, Banco, Tipo_Cuenta, Num_Cuenta, Forma_Pago, URL
    ]);
  });

  if (filas.length > 0) {
    hP.getRange(2, 1, filas.length, 21).setValues(filas);
    hP.getRange(2, 12, filas.length, 1).setNumberFormat("Q#,##0.00");
    _colorearParticipantes(hP, filas.length);
    // Marcar celdas sin Creamos ID
    filas.forEach(function(f, i) {
      if (!f[0]) {
        var celdaID = hP.getRange(i + 2, 1);
        celdaID.setValue("⚠️ Sin perfil Salesforce")
               .setBackground("#fce8e6").setFontColor("#c0392b").setFontStyle("italic");
      }
    });
  }

  _alert(
    "✅ Lista oficial cargada — " + filas.length + " participantes.\n\n" +
    "• Encontrados en base Creamos: " + encontrados + "\n" +
    (sinEncontrar.length
      ? "• Sin perfil en Creamos (" + sinEncontrar.length + "):\n  " + sinEncontrar.join("\n  ") +
        "\n\n⚠️ Marcadas en rojo — cuando creen el perfil en Salesforce\n" +
        "usa Administración → Sincronizar desde Creamos DB."
      : "• Todos tienen Creamos ID ✅")
  );
}); }

/**
 * Genera ID en formato: 2 letras del primer nombre + 2 letras del primer apellido + seq 3 dígitos
 * Ejemplo: ANGELICA VELIZ → ANVE001
 */
/**
 * Busca un participante en la base de datos oficial "Copy of CREAMOS ID nuevo".
 * Estrategias: exacto → normalizado → primeras 2 palabras.
 * Retorna: { id, dpi, edad, genero, fechaNac, anioEntrada } o {} si no encontrado.
 */
function _buscarEnCreamos_DB(nombre) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hDB  = ss.getSheetByName(CFG.HOJAS.CREAMOS_DB);
  if (!hDB || hDB.getLastRow() < 2) return {};

  var datos = hDB.getDataRange().getValues();
  var enc   = datos[0];

  // Detectar índices de columnas por encabezado
  var iNombre = -1, iID = -1, iAnio = -1, iEdad = -1, iGenero = -1, iFechaNac = -1, iDPI = -1;
  enc.forEach(function(h, i) {
    var hl = String(h).toLowerCase().trim();
    if (hl.indexOf("nombre") !== -1)               iNombre   = i;
    if ((hl.indexOf("creamos") !== -1 && hl.indexOf("id") !== -1) || hl === "id") iID = i;
    if (hl.indexOf("año") !== -1 || hl.indexOf("anio") !== -1 || hl.indexOf("entró") !== -1) iAnio = i;
    if (hl === "age" || hl === "edad")             iEdad     = i;
    if (hl === "gender" || hl.indexOf("género") !== -1 || hl.indexOf("genero") !== -1) iGenero = i;
    if (hl.indexOf("nacimiento") !== -1 || hl.indexOf("fecha") !== -1) iFechaNac = i;
    if (hl.indexOf("dpi") !== -1)                  iDPI      = i;
  });
  if (iNombre === -1) return {};

  var normBuscar = textoParaComparar(nombre);
  var palabras   = normBuscar.split(/\s+/).slice(0, 2).join(" ");

  function extraer(fila) {
    return {
      id:        iID      >= 0 ? String(fila[iID]      || "").trim() : "",
      dpi:       iDPI     >= 0 ? String(fila[iDPI]     || "").trim() : "",
      edad:      iEdad    >= 0 ? String(fila[iEdad]    || "").trim() : "",
      genero:    iGenero  >= 0 ? String(fila[iGenero]  || "").trim() : "",
      fechaNac:  iFechaNac>= 0 ? fila[iFechaNac]                     : "",
      anioEntrada: iAnio  >= 0 ? String(fila[iAnio]   || "").trim() : ""
    };
  }

  // 1) Exacto
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][iNombre]).trim() === nombre) return extraer(datos[i]);
  }
  // 2) Normalizado
  for (var i = 1; i < datos.length; i++) {
    if (textoParaComparar(String(datos[i][iNombre])) === normBuscar) return extraer(datos[i]);
  }
  // 3) Primeras 2 palabras
  for (var i = 1; i < datos.length; i++) {
    var normDB = textoParaComparar(String(datos[i][iNombre]));
    if (normDB.indexOf(palabras) === 0 || palabras.indexOf(normDB.split(/\s+/).slice(0,2).join(" ")) === 0) {
      return extraer(datos[i]);
    }
  }
  return { noEncontrado: true };
}

/** Devuelve true si el valor de la celda Creamos_ID es real (no el marcador de "sin perfil") */
function _esCreamos_ID_real(valor) {
  var v = String(valor || "").trim();
  return v.length > 0 && v.indexOf("⚠️") === -1;
}

/**
 * Protege y oculta la hoja de la base de datos Creamos.
 * Solo el admin puede verla/editarla.
 */
function _protegerHojaCreamos_DB() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var hDB = ss.getSheetByName(CFG.HOJAS.CREAMOS_DB);
  if (!hDB) return;
  // Ocultar la hoja
  hDB.hideSheet();
  // Proteger contra edición
  var protecciones = hDB.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (protecciones.length === 0) {
    var prot = hDB.protect();
    prot.setDescription("Base de datos oficial Creamos — NO EDITAR");
    prot.setWarningOnly(false);
    // Solo el propietario/admin puede editar
    var yo = Session.getEffectiveUser();
    prot.addEditor(yo);
    prot.removeEditors(prot.getEditors().filter(function(e){ return e.getEmail() !== yo.getEmail(); }));
  }
}

/** Colorea filas de PARTICIPANTES según categoría */
function _colorearParticipantes(hP, total) {
  for (var i = 0; i < total; i++) {
    var cat = String(hP.getRange(i + 2, 11).getValue()).trim().toUpperCase();
    var color = (CFG.COLORES_CAT[cat] || {}).bgClaro || "#ffffff";
    hP.getRange(i + 2, 1, 1, 20).setBackground(color);
  }
}

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
    var id      = String(r[0]).trim();
    var nombre  = String(r[1]).trim();
    var cat     = String(r[10]).trim().toUpperCase();
    var tarifa  = parseFloat(r[11]) || CFG.CATEGORIAS[cat] || 0;
    var estado  = String(r[5]).trim() || "Activo";
    if (!nombre || !porCat[cat]) return;
    porCat[cat].push({ id: id, nombre: nombre, tarifa: tarifa, estado: estado });
  });

  // Recrear hoja Directorio
  var HOJA = "Directorio";
  var hD = ss.getSheetByName(HOJA);
  if (hD) ss.deleteSheet(hD);
  hD = ss.insertSheet(HOJA);

  var filas = [], tipos = [];

  function push(fila, tipo) { filas.push(fila); tipos.push(tipo); }

  var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  push(["DIRECTORIO DE PARTICIPANTES — " + CFG.ORG, "", "", "", "", ""], "titulo");
  push(["Actualizado: " + ahora,                      "", "", "", "", ""], "sub");
  push(["", "", "", "", "", ""],                                           "vacio");

  // encabezado de columnas
  push(["#", "CREAMOS ID", "NOMBRE", "CATEGORÍA", "TARIFA Q/HR", "ESTADO"], "enc");

  var num = 1;
  var COLORES_CAT = { A: CFG.COLORES_CAT.A.bgClaro, B: CFG.COLORES_CAT.B.bgClaro, C: CFG.COLORES_CAT.C.bgClaro, D: CFG.COLORES_CAT.D.bgClaro };
  var TITULO_CAT  = { A: CFG.COLORES_CAT.A.bg, B: CFG.COLORES_CAT.B.bg, C: CFG.COLORES_CAT.C.bg, D: CFG.COLORES_CAT.D.bg };

  ["A","B","C","D"].forEach(function(cat) {
    var lista = (porCat[cat] || []).slice().sort(function(a,b){ return a.nombre.localeCompare(b.nombre,"es"); });
    if (!lista.length) return;
    var tarifa = CFG.CATEGORIAS[cat];

    push(["▶  CATEGORÍA " + cat, "", "Q" + tarifa.toFixed(2) + " / hora",
          lista.length + " participante" + (lista.length > 1 ? "s" : ""),
          "", ""], "cat_" + cat);

    lista.forEach(function(p) {
      push([num++, p.id, p.nombre, cat, "Q" + tarifa.toFixed(2), p.estado],
           p.estado === "Activo" ? "activo" : "inactivo");
    });

    push(["", "", "Subtotal categoría " + cat + ": " + lista.length, "", "", ""], "subtotal");
    push(["", "", "", "", "", ""], "vacio");
  });

  // Total general
  var totalActivos = Object.keys(porCat).reduce(function(s, c) {
    return s + porCat[c].filter(function(p){ return p.estado === "Activo"; }).length;
  }, 0);
  push(["", "", "TOTAL ACTIVOS: " + totalActivos, "", "", ""], "total");

  hD.getRange(1, 1, filas.length, 6).setValues(filas);

  // ── Formato ──────────────────────────────────────────────────
  tipos.forEach(function(tipo, idx) {
    var r = hD.getRange(idx + 1, 1, 1, 6);
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
      hD.getRange(idx + 1, 1, 1, 6)
        .setBorder(null, null, true, null, null, null, "#dadce0", SpreadsheetApp.BorderStyle.SOLID);
    }
  });

  // anchos de columna
  hD.setColumnWidth(1, 45);
  hD.setColumnWidth(2, 100);
  hD.setColumnWidth(3, 260);
  hD.setColumnWidth(4, 100);
  hD.setColumnWidth(5, 110);
  hD.setColumnWidth(6, 90);
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
    tipoCuenta:       String(f[17] || ""), numCuenta:        String(f[18] || ""),
    formaPago:        String(f[19] || "")
  };
  var urlActual = String(f[20] || "");

  var carpeta = _carpetaDP();
  var doc = _abrirOCrearDocProceso(id, nombre, carpeta, urlActual);
  _escribirContenidoDP(doc, part);

  var urlNueva = doc.getUrl();
  if (urlNueva !== urlActual) hP.getRange(fila, 21).setValue(urlNueva);
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

function _escribirContenidoDP(doc, part) {
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
    ["Banco",          part.banco       ||"—", "Tipo cuenta",   part.tipoCuenta  ||"—"],
    ["Núm. cuenta",    part.numCuenta   ||"—", "",              ""],
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

  // 3. Información de pago
  body.appendParagraph("");
  body.appendParagraph("3.  INFORMACIÓN DE PAGO")
      .setHeading(DocumentApp.ParagraphHeading.HEADING3)
      .editAsText().setForegroundColor("#639922");
  body.appendTable([
    ["Tarifa",        tarifahora],
    ["IVA 5%",        part.tieneFactura === "Sí" ? "Sí — Pequeño Contribuyente" : "No aplica"],
    ["Banco",         part.banco || "—"],
    ["Tipo cuenta",   part.tipoCuenta || "—"],
    ["Núm. cuenta",   part.numCuenta || "—"],
    ["Forma de pago", part.formaPago || "—"],
  ]);
  body.appendParagraph("");
  body.appendParagraph("Ver reportes de quincena en el Spreadsheet para historial de horas y montos.")
      .editAsText().setItalic(true).setFontSize(9).setForegroundColor("#888888");

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

/**
 * ⚡ IMPORTAR + EMPAREJAR TODO (1 clic)
 * Hace en secuencia: importar Kobo → normalizar nombres → emparejar entradas/salidas
 */
function importarYEmparejar() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [], errores = [];

  ss.toast("Paso 1/3: Importando desde Kobo...", "⚡", -1);
  try {
    var resp = UrlFetchApp.fetch(CFG.KOBO_URL_CSV, { muteHttpExceptions: true });
    var code = resp.getResponseCode();
    if (code === 503) { _alert("⏳ Kobo ocupado (503). Espera 2 min e intenta."); return; }
    if (code !== 200) throw new Error("HTTP " + code);
    var datos = Utilities.parseCsv(resp.getContentText(), ";");
    if (datos.length < 2) throw new Error("Kobo sin registros");
    var hK = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO) || ss.insertSheet(CFG.HOJAS.DATOS_KOBO);
    hK.clearContents();
    hK.getRange(1,1,datos.length,datos[0].length).setValues(datos);
    hK.getRange(1,1,1,datos[0].length).setFontWeight("bold").setBackground("#4a86e8").setFontColor("#fff");
    hK.setFrozenRows(1);
    _limpiarColumnasKobo(hK, datos[0]);   // ocultar cols irrelevantes + normalizar nombres
    _normalizarAccionSilencioso(hK);
    log.push("✅ Paso 1: " + (datos.length-1) + " registros importados y normalizados");
  } catch(e) { errores.push("❌ Paso 1: " + e.message); }

  ss.toast("Paso 2/3: Emparejando entradas y salidas...", "⚡", -1);
  try {
    emparejarAsistencia();  // llama directamente sin _run (ya estamos dentro de uno)
    log.push("✅ Paso 2: Emparejamiento completado → hoja ASISTENCIA actualizada");
  } catch(e) { errores.push("❌ Paso 2: " + e.message); }

  ss.toast("Paso 3/3: Actualizando quincena actual...", "⚡", -1);
  try {
    var periodo = _periodoActivo();
    if (periodo) {
      var fi = new Date(periodo.fi), ff = new Date(periodo.ff);
      var horasReponer = _leerHorasReponerExistentes(periodo.tab);
      _generarReporteQuincena(fi, ff, periodo.label, periodo.tab, horasReponer);
      log.push("✅ Paso 3: Quincena '" + periodo.label + "' actualizada");
    } else {
      log.push("ℹ️ Paso 3: Sin quincena activa — ve a 📅 Quincena → Nueva quincena");
    }
  } catch(e) { errores.push("❌ Paso 3: " + e.message); }

  _alert(
    "⚡ IMPORTAR + EMPAREJAR TODO\n\n" +
    log.join("\n") +
    (errores.length ? "\n\n" + errores.join("\n") : "")
  );
}); }

/**
 * ⚡ CHECKLIST + RECIBOS (1 clic)
 * Genera el checklist de pago y los recibos del período activo en secuencia.
 */
function procesarPagoCompleto() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var periodo = _periodoActivo();
  var label   = periodo ? periodo.label : "período actual";

  var resp = ui.alert(
    "⚡ Procesar pago completo",
    "Se generará:\n" +
    "1. Checklist de pago (Checklist_...)\n" +
    "2. Recibos individuales en Drive\n\n" +
    "Período: " + label + "\n\n¿Continuar?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var log = [], errores = [];

  SpreadsheetApp.getActiveSpreadsheet().toast("Paso 1/2: Generando checklist...", "💰", -1);
  try {
    generarChecklistPago();
    log.push("✅ Paso 1: Checklist generado");
  } catch(e) { errores.push("❌ Paso 1 checklist: " + e.message); }

  SpreadsheetApp.getActiveSpreadsheet().toast("Paso 2/2: Generando recibos...", "💰", -1);
  try {
    generarRecibosMes();
    log.push("✅ Paso 2: Recibos generados en Drive");
  } catch(e) { errores.push("❌ Paso 2 recibos: " + e.message); }

  _alert(
    "💰 PROCESO DE PAGO COMPLETO\n\n" +
    log.join("\n") +
    (errores.length ? "\n\n" + errores.join("\n") : "")
  );
}); }

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
  _alert("✅ Reimportación completa: " + (datos.length-1) + " registros importados desde Kobo.");
}); }

function _buscarIndice(enc, clave) {
  var c = clave.toLowerCase();
  for (var i=0; i<enc.length; i++) { if (String(enc[i]).trim().toLowerCase()===c) return i; }
  return -1;
}

function _limpiarColumnasKobo(hoja, enc) {
  // Ocultar columnas irrelevantes
  var imp = ["start","end","ingreso","egreso","entrada","salida",
             "participante","nombre","seleccione","c_id","_uuid","uuid",
             "accion","acción","terapia","permiso","comput","ingreso_egreso"];
  for (var i = 0; i < enc.length; i++) {
    var h = String(enc[i]).trim().toLowerCase();
    var esImp = h && imp.some(function(p){ return h.indexOf(p) !== -1; });
    try { if (!esImp) hoja.hideColumns(i+1); else hoja.showColumns(i+1); } catch(_) {}
  }

  // Normalizar columna Participante: slug → nombre oficial
  var cols = detectarColumnas(enc, []);
  if (cols.participante === undefined) return;
  var mapeo = cargarMapeoNombres();
  var lastRow = hoja.getLastRow();
  if (lastRow < 2) return;

  var colP = cols.participante + 1; // 1-indexed
  var valores = hoja.getRange(2, colP, lastRow - 1, 1).getValues();
  var cambiados = 0;
  var nuevos = valores.map(function(r) {
    var raw = String(r[0] || "").trim();
    if (!raw) return [raw];
    var normalizado = normalizarNombre(raw, mapeo);
    if (normalizado !== raw) cambiados++;
    return [normalizado];
  });
  if (cambiados > 0) hoja.getRange(2, colP, lastRow - 1, 1).setValues(nuevos);
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
  var estimados  = 0;  // entradas sin salida → jornada estimada

  Object.keys(grupos).forEach(function(clave) {
    var g = grupos[clave];
    g.ent.sort(function(a,b){return a-b;});
    g.sal.sort(function(a,b){return a-b;});

    var horas, salida;
    if (g.ent.length > 0 && g.sal.length > 0) {
      // Par completo
      horas  = Math.max(0, Math.round((g.sal[g.sal.length-1] - g.ent[0]) / 36000) / 100);
      salida = g.sal[g.sal.length-1];
    } else if (g.ent.length > 0 && g.sal.length === 0) {
      // Entrada sin salida → estimar jornada normal
      horas  = CFG.HORAS_JORNADA_NORMAL;
      salida = new Date(g.ent[0].getTime() + horas * 3600000);
      estimados++;
    } else {
      return; // Salida sin entrada: ignorar (caso raro)
    }

    var id       = extraerCodigo(g.nombre) || "";
    var esDiaEst = esDiaDeEstudio(g.nombre, g.fecha, diasEstudioMap) ? "Sí" : "No";
    var esTer    = (listaTerapias[g.nombre] || g.esTerapia) ? "Sí" : "No";
    var pct      = (esDiaEst==="Sí" || esTer==="Sí") ? 0 : 100;
    var hap      = Math.round(horas * (pct/100) * 100) / 100;
    var tipo2    = esDiaEst==="Sí" ? "Día de Estudio" : (esTer==="Sí" ? "Terapia" : "Normal");
    filasAsist.push([g.nombre, id, g.ent[0], tipo2, horas, esDiaEst, esTer, pct, hap, clave, g.ent[0], salida]);
  });

  filasAsist.sort(function(a,b){ return new Date(b[2])-new Date(a[2]); });
  var hA = _sh(CFG.HOJAS.ASISTENCIA);
  if (hA.getLastRow() > 1) hA.deleteRows(2, hA.getLastRow()-1);
  if (filasAsist.length > 0) {
    hA.getRange(2,1,filasAsist.length,12).setValues(filasAsist);
    hA.getRange("C2:C"+(filasAsist.length+1)).setNumberFormat("dd/MM/yyyy");
    hA.getRange("K2:L"+(filasAsist.length+1)).setNumberFormat("HH:mm");
    // Colorear filas estimadas en amarillo para identificarlas
    if (estimados > 0) {
      for (var ei=0; ei<filasAsist.length; ei++) {
        var salEstimada = filasAsist[ei][11];
        var entRaw      = filasAsist[ei][10];
        var diffMs = (salEstimada instanceof Date && entRaw instanceof Date)
                     ? salEstimada - entRaw : 0;
        var diffH = diffMs / 3600000;
        if (Math.abs(diffH - CFG.HORAS_JORNADA_NORMAL) < 0.01) {
          hA.getRange(ei+2, 1, 1, 12).setBackground("#fff9c4"); // amarillo suave
        }
      }
    }
  }
  _alert(
    "✅ Emparejamiento completado en ASISTENCIA\n\n" +
    "• Pares completos (entrada+salida): " + (filasAsist.length - estimados) + "\n" +
    (estimados > 0
      ? "• Estimados (solo entrada):          " + estimados + " ⚠️\n" +
        "  Se usó jornada de "+CFG.HORAS_JORNADA_NORMAL+"h. Filas en amarillo.\n" +
        "  Ejecuta 'Diagnosticar Datos Kobo' para ver cuáles."
      : "• Sin entradas sin par ✅")
  );
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
  // Formato largo: ANVE300380 (4 letras + 6 dígitos)
  var m = String(nombre).match(/([A-ZÁÉÍÓÚÑÜ]{4}\d{6})/i);
  if (m) return m[1].toUpperCase();
  // Formato corto: ANVE001 (4 letras + 3 dígitos)
  var m2 = String(nombre).match(/([A-ZÁÉÍÓÚÑÜ]{4}\d{3})/i);
  return m2 ? m2[1].toUpperCase() : null;
}
function limpiarNombre(nombre) {
  var s = String(nombre)
    .replace(/^[A-ZÁÉÍÓÚÑÜ]{4}\d{3,6}\s*/i, "")       // ID al inicio
    .replace(/\s*\([A-ZÁÉÍÓÚÑÜ]{4}\d{3,6}\)\s*/i, ""); // ID en paréntesis
  return s.replace(/^[•\s]+/, "").trim();
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
/**
 * Construye mapa de normalización de nombres:
 *   slug_kobo         → nombre_oficial  (ej. "sindy_paola_lazaro_diaz" → "Sindy Paola Lazaro Diaz")
 *   nombre_sin_tildes → nombre_oficial  (búsqueda tolerante)
 * Fuente primaria: LISTA_OFICIAL (índice 3 = kobo_slug, índice 0 = nombre oficial)
 * Fuente secundaria: hoja NombresCanonicos (para aliases manuales extra)
 */
function cargarMapeoNombres() {
  var m = {};

  // 1. Desde LISTA_OFICIAL: slug kobo → nombre oficial
  LISTA_OFICIAL.forEach(function(item) {
    var nombreOficial = item[1];
    var slug          = item[4] || "";
    if (slug) {
      m[slug] = nombreOficial;
      // También mapear slug con espacios (kobo a veces usa espacios en lugar de _)
      m[slug.replace(/_/g, " ")] = nombreOficial;
    }
    // Mapear versión sin tildes del nombre oficial
    var sinTildes = textoParaComparar(item[1]);
    if (sinTildes !== nombreOficial.toLowerCase()) {
      m[sinTildes] = nombreOficial;
    }
  });

  // 2. Desde hoja NombresCanonicos (aliases manuales extras)
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NombresCanonicos");
  if (h && h.getLastRow() > 1) {
    var d = h.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      var alias   = String(d[i][0] || "").trim();
      var oficial = String(d[i][1] || "").trim();
      if (alias && oficial) m[alias] = oficial;
    }
  }

  return m;
}
function normalizarNombre(nombre, mapeo) {
  if (!mapeo || !Object.keys(mapeo).length) return nombre;
  // Intento exacto
  if (mapeo[nombre]) return mapeo[nombre];
  // Intento sin tildes / minúsculas
  var norm = textoParaComparar(nombre);
  if (mapeo[norm]) return mapeo[norm];
  // Intento slug (reemplazar espacios por _)
  var slug = norm.replace(/\s+/g, "_");
  if (mapeo[slug]) return mapeo[slug];
  return nombre;
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
    var nombre=String(d[f][1]||"").trim(); // col B = Participante
    if(nombre && String(d[f][2]||"").trim().toUpperCase()==="X") lista[nombre]=true; // col C = Recibe Terapia
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

  var COLORES = { A: CFG.COLORES_CAT.A.bgClaro, B: CFG.COLORES_CAT.B.bgClaro, C: CFG.COLORES_CAT.C.bgClaro, D: CFG.COLORES_CAT.D.bgClaro, "?":"#f1f3f4" };
  var TITULO_CAT = { A: CFG.COLORES_CAT.A.bg, B: CFG.COLORES_CAT.B.bg, C: CFG.COLORES_CAT.C.bg, D: CFG.COLORES_CAT.D.bg, "?":"#9aa0a6" };
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
 *
 * Calcula horas por participante para el período.
 * Parte de TODOS los participantes de PARTICIPANTES (horas=0),
 * luego suma horas reales desde ASISTENCIA (si existe) o DatosKobo.
 *
 * Retorna: { "Nombre": { horas, tarifa, tieneFactura, categoria, codigo } }
 */
function _calcularResumenPeriodo(fi, ff) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var mapa = _construirMapaTarifas(); // {nombre: {id, tarifa, categoria, tieneFactura}}
  var dIni = new Date(fi.getFullYear(), fi.getMonth(), fi.getDate());
  var dFin = new Date(ff.getFullYear(), ff.getMonth(), ff.getDate()); // inclusive

  // Siempre inicializar con TODOS los participantes de PARTICIPANTES (horas = 0)
  // Así aparecen los 33 aunque no hayan registrado asistencia en el período.
  var resultado = {};
  Object.keys(mapa).forEach(function(nombre) {
    var info = mapa[nombre];
    resultado[nombre] = {
      horas:        0,
      tarifa:       info.tarifa,
      tieneFactura: info.tieneFactura,
      categoria:    info.categoria,
      id:           _esCreamos_ID_real(info.id) ? info.id : "",
      codigo:       _esCreamos_ID_real(info.id) ? info.id : ""
    };
  });

  // Sumar horas reales desde DatosKobo (si no existe ASISTENCIA pre-procesada)
  var hojaA = ss.getSheetByName(CFG.HOJAS.ASISTENCIA);
  if (hojaA && hojaA.getLastRow() > 1) {
    var asistRows = hojaA.getDataRange().getValues();
    for (var ai = 1; ai < asistRows.length; ai++) {
      var r     = asistRows[ai];
      var fecha = new Date(r[2]);
      if (isNaN(fecha)) continue;
      var dia   = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate());
      if (dia < dIni || dia > dFin) continue;
      var nombre = String(r[0]).trim();
      if (!nombre) continue;
      var horas = parseFloat(r[8]) || 0; // col I = Horas_A_Pagar
      if (!resultado[nombre]) {
        // Persona en DatosKobo pero no en PARTICIPANTES → agregarla igual
        var info2 = _buscarInfoParticipante(mapa, nombre);
        resultado[nombre] = { horas: 0, tarifa: info2.tarifa || CFG.CATEGORIAS.C,
          tieneFactura: info2.tieneFactura || false, categoria: info2.categoria || "?",
          id: info2.id || "", codigo: info2.id || "" };
      }
      resultado[nombre].horas = Math.round((resultado[nombre].horas + horas) * 100) / 100;
    }
    return resultado;
  }

  // ── Estrategia 2: DatosKobo (fallback) ───────────────────────
  var hK = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (!hK || hK.getLastRow() < 2) return {};

  var enc  = hK.getRange(1, 1, 1, hK.getLastColumn()).getValues()[0];
  var raw  = hK.getRange(2, 1, hK.getLastRow()-1, hK.getLastColumn()).getValues();
  var cols = detectarColumnas(enc, raw.slice(0, 50));
  var mapeoNombres = cargarMapeoNombres();

  // Índice de timestamp: preferir cols.start; si no existe usar col 0
  var iTS = (cols.start !== undefined) ? cols.start : 0;
  // Índice de timestamp de fin de sesión (para salidas estimadas)
  var iEnd = (cols.end !== undefined) ? cols.end : -1;

  var porPart = {};

  raw.forEach(function(fila) {
    var ts = new Date(fila[iTS]);
    if (isNaN(ts)) return;

    var dia = new Date(ts.getFullYear(), ts.getMonth(), ts.getDate());
    if (dia < dIni || dia > dFin) return;

    var nombreRaw = obtenerParticipanteFila(fila, cols);
    if (!nombreRaw) return;
    var nombre = normalizarNombre(nombreRaw, mapeoNombres) || limpiarNombre(nombreRaw);
    if (!nombre) return;

    var tipo = obtenerTipoRegistro(fila, cols);
    if (!tipo.esIngreso && !tipo.esEgreso) return;

    if (!porPart[nombre]) porPart[nombre] = [];
    var tsEnd = (iEnd >= 0) ? new Date(fila[iEnd]) : null;
    porPart[nombre].push({ ts: ts, tsEnd: tsEnd, tipo: tipo });
  });

  // Partir de todos los participantes (horas=0), luego sumar horas reales de Kobo
  var resultado2 = {};
  Object.keys(mapa).forEach(function(nombre) {
    var info = mapa[nombre];
    resultado2[nombre] = { horas: 0, tarifa: info.tarifa, tieneFactura: info.tieneFactura,
      categoria: info.categoria, id: _esCreamos_ID_real(info.id) ? info.id : "",
      codigo: _esCreamos_ID_real(info.id) ? info.id : "" };
  });

  Object.keys(porPart).forEach(function(nombre) {
    var regs = porPart[nombre].sort(function(a,b){ return a.ts - b.ts; });
    var totalH = 0;
    var entrada = null;

    regs.forEach(function(reg) {
      if (reg.tipo.esIngreso && !entrada) {
        entrada = reg.ts;
      } else if (reg.tipo.esEgreso && entrada) {
        // Usar tsEnd del registro de salida si es confiable (< 16h de diferencia)
        var tSalida = (reg.tsEnd && !isNaN(reg.tsEnd) &&
                       (reg.tsEnd - entrada)/3600000 > 0 &&
                       (reg.tsEnd - entrada)/3600000 < 16)
                      ? reg.tsEnd : reg.ts;
        var diffH = (tSalida - entrada) / 3600000;
        if (diffH > 0 && diffH <= 16) totalH += diffH;
        entrada = null;
      }
    });

    if (entrada) {
      // Entrada sin salida → estimar jornada normal
      totalH += CFG.HORAS_JORNADA_NORMAL;
    }

    if (!resultado2[nombre]) {
      var info2b = _buscarInfoParticipante(mapa, nombre);
      resultado2[nombre] = { horas: 0, tarifa: info2b.tarifa || CFG.CATEGORIAS.C,
        tieneFactura: info2b.tieneFactura || false, categoria: info2b.categoria || "?",
        id: info2b.id || "", codigo: info2b.id || "" };
    }
    resultado2[nombre].horas = Math.round(totalH * 100) / 100;
  });

  return resultado2;
}

/*
 * Busca la información de tarifa/categoría/factura de un participante
 * en el mapa, usando múltiples estrategias de normalización de nombre.
 */
function _buscarInfoParticipante(mapa, nombre) {
  // 1. Exacto
  if (mapa[nombre]) return mapa[nombre];
  // 2. Normalizado (sin acentos, minúsculas)
  var norm = limpiarNombre(nombre);
  for (var k in mapa) {
    if (limpiarNombre(k) === norm) return mapa[k];
  }
  // 3. Primeras dos palabras (apellido + nombre parcial)
  var palabras = norm.split(/\s+/);
  if (palabras.length >= 2) {
    var inicio = palabras.slice(0, 2).join(" ");
    for (var k2 in mapa) {
      if (limpiarNombre(k2).indexOf(inicio) === 0) return mapa[k2];
    }
  }
  return {};
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

  // Asegurar suficientes columnas (ahora usamos 13)
  if (h.getMaxColumns() < 13) h.insertColumnsAfter(h.getMaxColumns(), 13 - h.getMaxColumns());

  // ── Fila 1: referencia de tarifas ──────────────────────────────
  h.getRange(1, 1, 1, 13).setValues([["Q16.50","Q15.75","Q15.00","Q14.00","","","","","","","","",""]]);
  ["#639922","#4285f4","#fbbc04","#ea4335"].forEach(function(c, i) {
    h.getRange(1, i+1).setBackground(c).setFontColor("#ffffff").setFontWeight("bold")
     .setHorizontalAlignment("center");
  });

  // ── Fila 2: título del período ─────────────────────────────────
  h.getRange(2, 1, 1, 13).merge()
   .setValue("Período: " + label)
   .setBackground("#f8f9fa").setFontWeight("bold").setFontSize(11)
   .setHorizontalAlignment("center")
   .setBorder(true,true,true,true,null,null,"#dadce0",SpreadsheetApp.BorderStyle.SOLID);

  // ── Fila 3: encabezados de columnas ───────────────────────────
  // A=#  B=Participante  C=Fact.  D-E=spacers
  // F=Total hrs  G=Hrs reponer  H=Total a pagar
  // I=Monto Base  J=IVA 5%  K=Total org paga  L=Redondeo  M=Neto part.
  var encabezados = ["#","Participante","Fact.","ID","",
    "Total hrs","Hrs reponer","Total a pagar",
    "Monto Base","IVA 5%","Total org paga","Redondeo","Neto part."];
  h.getRange(3, 1, 1, 13).setValues([encabezados])
   .setBackground("#546e7a").setFontColor("#ffffff").setFontWeight("bold")
   .setHorizontalAlignment("center")
   .setBorder(true,true,true,true,null,null,"#37474f",SpreadsheetApp.BorderStyle.SOLID);

  // Colores de fondo por categoría (claro = lectura fácil)
  var BG_CAT = { A: CFG.COLORES_CAT.A.bgClaro, B: CFG.COLORES_CAT.B.bgClaro, C: CFG.COLORES_CAT.C.bgClaro, D: CFG.COLORES_CAT.D.bgClaro, "?": "#f3f3f3" };

  // ── Filas de datos: una por participante ───────────────────────
  var nombres = Object.keys(resumen).sort(function(a,b){ return a.localeCompare(b,"es"); });
  var filaActual = 4;
  var sumBase = 0, sumIVA = 0, sumTotal = 0, totalGeneral = 0;
  var num = 1;
  var conFactura = 0, sinFactura = 0;

  // Acumular todo en un array para un solo setValues masivo (más rápido)
  var bloqueValores = [];
  var bloqueFilas   = [];   // {fila, cat, tieneFactura}

  nombres.forEach(function(nombre) {
    var d = resumen[nombre];
    var hReponer = prevHorasReponer[nombre] || 0;

    // ── Cálculo exacto (2 decimales en cada paso) ─────────────────
    var horas   = Math.round(d.horas   * 100) / 100;
    var hTotal  = Math.round((horas + hReponer) * 100) / 100;
    var base    = Math.round(hTotal * d.tarifa * 100) / 100;
    var iva     = d.tieneFactura ? Math.round(base * CFG.IVA_PCT * 100) / 100 : 0;
    var orgPaga = Math.round((base + iva) * 100) / 100;
    var redond  = Math.round(orgPaga);
    var neto    = base;  // participante retiene base; IVA va a SAT vía Declaraguate

    sumBase  += base;
    sumIVA   += iva;
    sumTotal += redond;
    totalGeneral += redond;
    if (d.tieneFactura) conFactura++; else sinFactura++;

    var factInd = d.tieneFactura ? "★ Sí" : "—";

    bloqueValores.push([
      num++, nombre, factInd, d.codigo || "", "",
      horas || "",
      hReponer > 0 ? hReponer : "",
      hTotal || "",
      base || "",
      iva > 0 ? iva : "",
      orgPaga || "",
      redond || "",
      neto || ""
    ]);
    bloqueFilas.push({ fila: filaActual, cat: d.categoria || "?", tieneFactura: d.tieneFactura, sinHoras: horas === 0 });
    filaActual++;
  });

  // Escribir todos los valores de una vez
  if (bloqueValores.length > 0) {
    h.getRange(4, 1, bloqueValores.length, 13).setValues(bloqueValores);
  }

  // Aplicar formato fila a fila (colores por categoría + números)
  bloqueFilas.forEach(function(bf) {
    // Filas sin horas = gris claro + itálica para distinguirlas
    var bg = bf.sinHoras ? "#f5f5f5" : (BG_CAT[bf.cat] || BG_CAT["?"]);
    h.getRange(bf.fila, 1, 1, 13)
     .setBackground(bg).setFontFamily("Arial").setFontSize(10).setVerticalAlignment("middle")
     .setFontStyle(bf.sinHoras ? "italic" : "normal")
     .setFontColor(bf.sinHoras ? "#9e9e9e" : "#000000");
    h.getRange(bf.fila, 1).setHorizontalAlignment("center").setFontWeight("bold");
    h.getRange(bf.fila, 2, 1, 3).setHorizontalAlignment("left");
    h.getRange(bf.fila, 6, 1, 8).setHorizontalAlignment("right");
    // Col C (Fact.) — color especial si tiene factura
    if (bf.tieneFactura) {
      h.getRange(bf.fila, 3)
       .setBackground("#fce8e6").setFontColor("#c5221f").setFontWeight("bold")
       .setHorizontalAlignment("center");
    } else {
      h.getRange(bf.fila, 3).setFontColor("#9aa0a6").setHorizontalAlignment("center");
    }
    // Formato moneda cols I–M (9–13)
    h.getRange(bf.fila, 9, 1, 5).setNumberFormat('"Q"#,##0.00');
    h.getRange(bf.fila, 1, 1, 13)
     .setBorder(null,null,true,null,null,null,"#cccccc",SpreadsheetApp.BorderStyle.SOLID);
  });

  // ── Bloque de totales desglosados ──────────────────────────────
  filaActual++;

  // Fila: subtotal monto base
  h.getRange(filaActual, 1, 1, 13).setValues([
    ["","Monto base (sin IVA)","","","","","","",
     Math.round(sumBase*100)/100,"","","",""]
  ]);
  h.getRange(filaActual,2).setFontStyle("italic").setFontColor("#555555");
  h.getRange(filaActual,9).setNumberFormat('"Q"#,##0.00').setBackground("#f8f9fa");
  filaActual++;

  // Fila: total IVA (Declaraguate)
  h.getRange(filaActual, 1, 1, 13).setValues([
    ["","IVA 5% total a declarar (Declaraguate)","","","","","","",
     "","",Math.round(sumIVA*100)/100,"",""]
  ]);
  h.getRange(filaActual,2).setFontWeight("bold").setFontColor("#c5221f");
  h.getRange(filaActual,11)
   .setNumberFormat('"Q"#,##0.00').setBackground("#fce8e6").setFontColor("#c5221f").setFontWeight("bold");
  filaActual++;

  // Fila: GRAN TOTAL (lo que paga la organización)
  h.getRange(filaActual, 1, 1, 13).setValues([
    ["","TOTAL QUE PAGA LA ORGANIZACIÓN","","","","","","",
     "","","",totalGeneral,""]
  ]);
  h.getRange(filaActual,2).setFontWeight("bold").setFontSize(11);
  h.getRange(filaActual,12)
   .setBackground("#00c853").setFontColor("#ffffff").setFontWeight("bold")
   .setFontSize(12).setHorizontalAlignment("center").setNumberFormat('"Q"#,##0.00');
  filaActual++;

  // ── Fila leyenda + resumen ─────────────────────────────────────
  filaActual++;
  h.getRange(filaActual, 1, 1, 13).setValues([[
    "A=Q16.50","B=Q15.75","C=Q15.00","D=Q14.00","",
    "★ = emite factura","","",
    "Con factura: "+conFactura,"Sin factura: "+sinFactura,"","",""
  ]]);
  [BG_CAT.A,BG_CAT.B,BG_CAT.C,BG_CAT.D].forEach(function(c,i){
    h.getRange(filaActual,i+1).setBackground(c).setFontSize(9)
     .setHorizontalAlignment("center").setFontWeight("bold");
  });
  h.getRange(filaActual,6,1,4).setFontSize(9).setFontColor("#555555");

  // ── Fila timestamp ─────────────────────────────────────────────
  filaActual++;
  h.getRange(filaActual, 1, 1, 12).merge()
   .setValue("Actualizado: " + ts + "  |  " + nombres.length + " participantes  |  " +
             nombres.filter(function(n){ return resumen[n].horas > 0; }).length + " con horas registradas")
   .setFontSize(8).setFontColor("#9aa0a6").setHorizontalAlignment("right");

  // ── Ancho de columnas (13 cols) ────────────────────────────────
  [35, 220, 60, 20, 20, 90, 100, 95, 95, 80, 95, 90, 90]
    .forEach(function(w, i) { h.setColumnWidth(i+1, w); });
  h.setRowHeight(2, 28);
  h.setRowHeight(3, 24);
  h.setFrozenRows(3);

  return totalGeneral;
}

// ══════════════════════════════════════════════════════════════════
// CHECKLIST DE PAGO
// Flujo post-quincena: qué hacer después de calcular el reporte
// ══════════════════════════════════════════════════════════════════

/*
 * Genera una hoja "Checklist_[período]" con el flujo completo de pago:
 *  SECCIÓN 1 — Con factura: Factura recibida → Declaraguate → Pago
 *  SECCIÓN 2 — Sin factura: Pago directo
 *  Resumen: total IVA a declarar, total pagado, pendientes
 *
 * Cada participante tiene columnas editables para marcar el avance.
 */
/**
 * Genera (o recarga) la hoja "Config_IVA" donde el admin marca quién emite factura.
 * Al guardar cambios en esa hoja, se propagan automáticamente a PARTICIPANTES col M.
 * También permite marcar TODAS con Sí / TODAS con No de un clic.
 */
function configurarFacturacion() { _run(function() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var hP  = _sh(CFG.HOJAS.PARTICIPANTES);
  var lastRow = hP.getLastRow();
  if (lastRow < 2) { _alert("Primero carga la lista de participantes."); return; }

  // Leer PARTICIPANTES
  var datos = hP.getRange(2, 1, lastRow - 1, 13).getValues(); // A→M

  // Crear/limpiar hoja Config_IVA
  var tabName = "Config_IVA";
  var hC = ss.getSheetByName(tabName);
  if (hC) { hC.clearContents(); hC.clearFormats(); }
  else    { hC = ss.insertSheet(tabName); }
  if (hC.getMaxColumns() < 5) hC.insertColumnsAfter(hC.getMaxColumns(), 5 - hC.getMaxColumns());

  // ── Encabezado ─────────────────────────────────────────────────
  hC.getRange(1,1,1,5).merge()
    .setValue("⚠️  CONFIGURACIÓN IVA — ¿Quién emite factura? (Pequeño Contribuyente 5%)")
    .setBackground("#e65100").setFontColor("#fff").setFontWeight("bold")
    .setFontSize(11).setHorizontalAlignment("center");

  hC.getRange(2,1,1,5).merge()
    .setValue("Marca 'Sí' en la col D para quienes emiten factura. Luego usa el botón 'Aplicar cambios a PARTICIPANTES'.")
    .setBackground("#fff3e0").setFontColor("#bf360c").setFontSize(9)
    .setHorizontalAlignment("center");

  hC.getRange(3,1,1,5).setValues([["#","Creamos ID","Participante","¿Tiene Factura?","Categoría"]])
    .setBackground("#37474f").setFontColor("#fff").setFontWeight("bold")
    .setHorizontalAlignment("center");

  // Validación desplegable Sí/No
  var vSN = SpreadsheetApp.newDataValidation().requireValueInList(["Sí","No"],true).build();

  var COLORES = { A: CFG.COLORES_CAT.A.bgClaro, B: CFG.COLORES_CAT.B.bgClaro, C: CFG.COLORES_CAT.C.bgClaro, D: CFG.COLORES_CAT.D.bgClaro };
  var filas   = [];
  datos.forEach(function(r, i) {
    var id     = String(r[0] || "").trim();
    var nombre = String(r[1] || "").trim();
    var cat    = String(r[10]|| "").trim().toUpperCase();
    var tieneFact = String(r[12]||"").trim();
    if (!nombre) return;
    filas.push([i+1, id, nombre, tieneFact||"No", cat]);
  });

  if (filas.length > 0) {
    var rango = hC.getRange(4, 1, filas.length, 5);
    rango.setValues(filas);
    // Color por categoría
    filas.forEach(function(f, i) {
      hC.getRange(i+4, 1, 1, 5).setBackground(COLORES[f[4]] || "#ffffff");
    });
    // Validación en col D (Tiene Factura)
    hC.getRange(4, 4, filas.length, 1).setDataValidation(vSN).setHorizontalAlignment("center");
    // Negrita y color especial para quienes ya tienen Sí
    filas.forEach(function(f, i) {
      if (f[3] === "Sí") {
        hC.getRange(i+4, 4).setBackground("#f9cb9c").setFontWeight("bold");
      }
    });
  }

  // Fila de totales
  var filaTot = 4 + filas.length + 1;
  hC.getRange(filaTot, 1, 1, 5).merge()
    .setValue("▶  Cuando termines de marcar, ve al menú: 💰 Facturación → 🧾 Configurar quién emite factura → para aplicar cambios")
    .setBackground("#e8f5e9").setFontColor("#1b5e20").setFontStyle("italic").setFontSize(9);

  // Ancho columnas
  hC.setColumnWidth(1,40); hC.setColumnWidth(2,130); hC.setColumnWidth(3,260);
  hC.setColumnWidth(4,130); hC.setColumnWidth(5,80);
  hC.setFrozenRows(3);
  hC.activate();

  _alert(
    "📋 HOJA 'Config_IVA' lista.\n\n" +
    "Instrucciones:\n" +
    "1. Cambia 'Sí' o 'No' en la columna D para cada participante\n" +
    "2. Vuelve al menú: 💰 Facturación → 🧾 Configurar quién emite factura\n" +
    "3. Elige 'Aplicar cambios'\n\n" +
    "¿Qué hace el IVA 5%?\n" +
    "• Si tiene factura: org paga Base + 5% (IVA lo declara en Declaraguate)\n" +
    "• Si NO tiene factura: org paga solo el Monto Base\n\n" +
    "¿Todas tienen factura? → abre PARTICIPANTES, col M, cambia todas a Sí."
  );
}); }

/**
 * Lee la hoja Config_IVA y aplica los cambios de Tiene_Factura a PARTICIPANTES.
 */
function aplicarCambiosFacturacion() { _run(function() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var hC  = ss.getSheetByName("Config_IVA");
  if (!hC) { _alert("Primero ejecuta 'Configurar quién emite factura' para crear la hoja."); return; }

  var hP  = _sh(CFG.HOJAS.PARTICIPANTES);
  var datosC = hC.getDataRange().getValues();
  var datosP = hP.getDataRange().getValues();

  // Construir índice nombre→fila en PARTICIPANTES
  var idxP = {};
  for (var i=1; i<datosP.length; i++) {
    var n = String(datosP[i][1]||"").trim();
    if (n) idxP[textoParaComparar(n)] = i+1; // fila real (1-indexed)
  }

  var actualizados = 0;
  for (var r=3; r<datosC.length; r++) { // desde fila 4 (datos)
    var nombre = String(datosC[r][2]||"").trim();
    var valor  = String(datosC[r][3]||"").trim();
    if (!nombre || (valor !== "Sí" && valor !== "No")) continue;
    var filaP = idxP[textoParaComparar(nombre)];
    if (!filaP) continue;
    var actual = String(hP.getRange(filaP, 13).getValue()).trim();
    if (actual !== valor) {
      hP.getRange(filaP, 13).setValue(valor);
      actualizados++;
    }
  }

  // Contar totales
  var total = hP.getLastRow()-1;
  var conFact = 0;
  if (total > 0) {
    var vals = hP.getRange(2,13,total,1).getValues();
    vals.forEach(function(v){ if(String(v[0]).trim()==="Sí") conFact++; });
  }

  _alert(
    "✅ Cambios aplicados a PARTICIPANTES\n\n" +
    "• Filas actualizadas: " + actualizados + "\n" +
    "• Con factura (IVA 5%): " + conFact + " participantes\n" +
    "• Sin factura:          " + (total - conFact) + " participantes\n\n" +
    "Los reportes de quincena ahora calcularán IVA solo para las " + conFact + " marcadas con Sí."
  );
}); }

function generarChecklistPago() { _run(function() {
  var ui  = SpreadsheetApp.getUi();
  var tz  = Session.getScriptTimeZone();

  // ¿Usar período activo o pedir fechas?
  var periodo = _periodoActivo();
  var fi, ff, label;

  if (periodo) {
    var r = ui.alert("Checklist de pago",
      "¿Generar el checklist para la quincena activa?\n" + periodo.label,
      ui.ButtonSet.YES_NO);
    if (r === ui.Button.YES) {
      fi = new Date(periodo.fi);
      ff = new Date(periodo.ff);
      label = periodo.label;
    }
  }

  if (!fi) {
    var r1 = ui.prompt("📋 Checklist de pago — Fecha inicio",
      "Formato dd/mm/yyyy:", ui.ButtonSet.OK_CANCEL);
    if (r1.getSelectedButton() !== ui.Button.OK) return;
    fi = _parseFecha(r1.getResponseText().trim());
    if (!fi) { _alert("Fecha inválida."); return; }

    var r2 = ui.prompt("📋 Checklist de pago — Fecha fin",
      "Formato dd/mm/yyyy:", ui.ButtonSet.OK_CANCEL);
    if (r2.getSelectedButton() !== ui.Button.OK) return;
    ff = _parseFecha(r2.getResponseText().trim());
    if (!ff) { _alert("Fecha inválida."); return; }
    label = _labelPeriodo(fi, ff);
  }

  var resumen = _calcularResumenPeriodo(fi, ff);
  if (!Object.keys(resumen).length) {
    _alert("No hay datos para ese período. Ejecuta 'Emparejar entradas/salidas' primero."); return;
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var tab   = "Checklist_" + Utilities.formatDate(fi, tz, "dd_MM") + "_" +
              Utilities.formatDate(ff, tz, "dd_MM_yyyy");
  var h = ss.getSheetByName(tab);
  if (h) ss.deleteSheet(h);
  h = ss.insertSheet(tab);

  if (h.getMaxColumns() < 12) h.insertColumnsAfter(h.getMaxColumns(), 12 - h.getMaxColumns());

  var ts = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm");

  // ── Fila 1: título ─────────────────────────────────────────────
  h.getRange(1,1,1,12).merge()
   .setValue("CHECKLIST DE PAGO — Período: " + label)
   .setBackground("#1a237e").setFontColor("#ffffff")
   .setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");

  // ── Fila 2: instrucciones ──────────────────────────────────────
  h.getRange(2,1,1,12).merge()
   .setValue("Marca cada paso al completarlo. Las columnas H, I, J, K son editables.")
   .setBackground("#e8eaf6").setFontColor("#3949ab").setFontSize(9)
   .setHorizontalAlignment("center");

  // ── Encabezados ────────────────────────────────────────────────
  var enc = ["#","Participante","ID","Cat.","Tarifa","Monto Base","IVA 5%",
             "Total org paga","★ Factura recibida","★ Declaraguate","★ Pagado","Notas"];
  h.getRange(3,1,1,12).setValues([enc])
   .setBackground("#37474f").setFontColor("#ffffff").setFontWeight("bold")
   .setHorizontalAlignment("center");

  var nombres = Object.keys(resumen).sort(function(a,b){ return a.localeCompare(b,"es"); });

  // Separar con factura / sin factura
  var conFact = nombres.filter(function(n){ return resumen[n].tieneFactura; });
  var sinFact = nombres.filter(function(n){ return !resumen[n].tieneFactura; });

  var filaActual = 4;
  var totalBase = 0, totalIVA = 0, totalOrg = 0;
  var num = 1;

  function seccion(titulo, bg, lista) {
    h.getRange(filaActual,1,1,12).merge()
     .setValue(titulo).setBackground(bg).setFontColor("#ffffff")
     .setFontWeight("bold").setFontSize(10).setHorizontalAlignment("left");
    filaActual++;

    lista.forEach(function(nombre) {
      var d    = resumen[nombre];
      var base = Math.round(d.horas * d.tarifa * 100) / 100;
      var iva  = d.tieneFactura ? Math.round(base * CFG.IVA_PCT * 100) / 100 : 0;
      var org  = Math.round((base + iva) * 100) / 100;
      totalBase += base; totalIVA += iva; totalOrg += org;

      h.getRange(filaActual,1,1,12).setValues([[
        num++,
        nombre,
        d.codigo || "",
        d.categoria || "?",
        "Q" + (d.tarifa||0).toFixed(2),
        base, iva > 0 ? iva : "—", org,
        "Pendiente","Pendiente","No",""
      ]]);

      // Formato
      var bgF = d.tieneFactura ? "#fff8e1" : "#f1f8e9";
      h.getRange(filaActual,1,1,12).setBackground(bgF).setFontSize(10);
      h.getRange(filaActual,1).setHorizontalAlignment("center");
      h.getRange(filaActual,6,1,3).setNumberFormat('"Q"#,##0.00').setHorizontalAlignment("right");
      // Cols editables: I, J, K, L (9-12) — fondo blanco destacado
      h.getRange(filaActual,9,1,4).setBackground("#ffffff")
       .setBorder(true,true,true,true,null,null,"#aaaaaa",SpreadsheetApp.BorderStyle.SOLID);
      // Borde fila
      h.getRange(filaActual,1,1,12)
       .setBorder(null,null,true,null,null,null,"#cccccc",SpreadsheetApp.BorderStyle.SOLID);

      filaActual++;
    });

    // Subtotal de sección
    h.getRange(filaActual,1,1,12).setValues([
      ["","Subtotal","","","",
       Math.round(totalBase*100)/100,
       Math.round(totalIVA*100)/100,
       Math.round(totalOrg*100)/100,"","","",""]
    ]);
    h.getRange(filaActual,1,1,12).setBackground("#eceff1").setFontStyle("italic");
    h.getRange(filaActual,6,1,3).setNumberFormat('"Q"#,##0.00').setHorizontalAlignment("right");
    filaActual += 2;
    // Reset para la siguiente sección
    totalBase = 0; totalIVA = 0; totalOrg = 0;
  }

  if (conFact.length) {
    seccion("  ★  CON FACTURA — Pequeño Contribuyente (IVA 5% → declarar en Declaraguate)",
            "#c62828", conFact);
  }
  if (sinFact.length) {
    seccion("  ✓  SIN FACTURA — Pago directo (sin IVA)",
            "#2e7d32", sinFact);
  }

  // ── Resumen total ──────────────────────────────────────────────
  // Recalcular totales generales
  var gBase = 0, gIVA = 0, gOrg = 0;
  nombres.forEach(function(n) {
    var d   = resumen[n];
    var base= Math.round(d.horas * d.tarifa * 100) / 100;
    var iva = d.tieneFactura ? Math.round(base * CFG.IVA_PCT * 100) / 100 : 0;
    gBase += base; gIVA += iva; gOrg += Math.round((base+iva)*100)/100;
  });

  h.getRange(filaActual,1,1,12).setValues([
    ["","TOTALES GENERALES","","","",
     Math.round(gBase*100)/100, Math.round(gIVA*100)/100, Math.round(gOrg*100)/100,
     "","","",""]
  ]);
  h.getRange(filaActual,1,1,12).setBackground("#1a237e").setFontColor("#ffffff").setFontWeight("bold");
  h.getRange(filaActual,6,1,3).setNumberFormat('"Q"#,##0.00').setHorizontalAlignment("right");
  filaActual += 2;

  // ── Guía de pasos ──────────────────────────────────────────────
  var pasos = [
    ["FLUJO DE PAGO — PASOS A SEGUIR", "", "", "", "", "", "", "", "", "", "", ""],
    ["PASO","Quién","Acción","Cuándo","","","","","","","",""],
    ["1","Participante (con factura)","Emite factura al " + CFG.ORG + " por el monto total (base + IVA 5%)","Antes del pago","","","","","","","",""],
    ["2","Organización",             "Recibe la factura → marca col H = 'Recibida'",                         "Al recibir factura","","","","","","","",""],
    ["3","Organización",             "Declara el IVA en Declaraguate a nombre del participante",              "Dentro del mes","","","","","","","",""],
    ["4","Organización",             "Realiza el pago al participante (monto total)",                         "Fecha acordada","","","","","","","",""],
    ["5","Organización",             "Marca col J = 'Sí' y col K = fecha en este checklist",                  "Al pagar","","","","","","","",""],
    ["","","","","","","","","","","",""],
    ["Sin factura:","Organización",  "Paga directamente el monto base (sin IVA ni declaraguate)",             "Fecha acordada","","","","","","","",""],
    ["","","","","","","","","","","",""],
    ["IVA 5%:","Pequeño Contribuyente","El participante retiene el IVA que recibió y lo paga a SAT via Declaraguate","Mensualmente","","","","","","","",""],
  ];

  h.getRange(filaActual, 1, pasos.length, 12).setValues(pasos);
  h.getRange(filaActual, 1, 1, 12).merge()
   .setBackground("#e8eaf6").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
  h.getRange(filaActual+1, 1, 1, 12)
   .setBackground("#c5cae9").setFontWeight("bold");
  // Formato filas de pasos
  for (var p=2; p<pasos.length; p++) {
    var bgP = p % 2 === 0 ? "#f5f5f5" : "#fafafa";
    h.getRange(filaActual+p, 1, 1, 12).setBackground(bgP).setFontSize(9);
    h.getRange(filaActual+p, 1).setFontWeight("bold").setHorizontalAlignment("center");
  }

  // ── Anchos y timestamp ─────────────────────────────────────────
  [35, 220, 50, 80, 90, 80, 90, 100, 90, 70, 90, 120]
    .forEach(function(w,i){ h.setColumnWidth(i+1,w); });
  h.setFrozenRows(3);

  // Validaciones de datos en cols editables
  var vFact = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pendiente","Recibida","No aplica"],true).build();
  var vDecl = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pendiente","Declarado","No aplica"],true).build();
  var vPag  = SpreadsheetApp.newDataValidation()
    .requireValueInList(["No","Sí"],true).build();
  h.getRange("H4:H200").setDataValidation(vFact);
  h.getRange("I4:I200").setDataValidation(vDecl);
  h.getRange("J4:J200").setDataValidation(vPag);

  // Timestamp al final
  h.getRange(h.getLastRow()+2, 1, 1, 12).merge()
   .setValue("Generado: " + ts + "  ·  " + nombres.length + " participantes")
   .setFontSize(8).setFontColor("#9aa0a6").setHorizontalAlignment("right");

  ss.setActiveSheet(h);
  _alert("✅ Checklist generado: '" + tab + "'\n\n" +
         "Con factura: " + conFact.length + " participantes\n" +
         "Sin factura: " + sinFact.length + " participantes\n\n" +
         "Las columnas ★ son editables para marcar el avance.");
}); }

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

// ── Sincronizar desde Creamos DB ──────────────────────────────

/**
 * Recorre PARTICIPANTES y para cada fila sin Creamos ID real,
 * vuelve a buscar en la DB. Si ahora lo encuentra, actualiza ID y DPI.
 * Útil cuando se crean perfiles nuevos en Salesforce y se sincronizan a la DB.
 */
function sincronizarDesdeCreamos() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hP = _sh(CFG.HOJAS.PARTICIPANTES);
  if (!hP || hP.getLastRow() < 2) { _alert("No hay participantes cargados."); return; }

  var datos = hP.getRange(2, 1, hP.getLastRow() - 1, 20).getValues();
  var actualizados = [], sinEncontrar = [];

  datos.forEach(function(fila, i) {
    var idActual = String(fila[0] || "").trim();
    if (_esCreamos_ID_real(idActual)) return; // ya tiene ID real, saltar

    var nombre = String(fila[1] || "").trim();
    if (!nombre) return;

    ss.toast("Buscando: " + nombre, "🔍", -1);
    var db = _buscarEnCreamos_DB(nombre);
    if (db.noEncontrado || !db.id) {
      sinEncontrar.push(nombre);
      return;
    }

    // Actualizar ID y DPI en la hoja
    hP.getRange(i + 2, 1).setValue(db.id).setBackground(null).setFontColor(null).setFontStyle("normal");
    if (db.dpi) hP.getRange(i + 2, 14).setValue(db.dpi);
    actualizados.push(nombre + " → " + db.id);
  });

  ss.toast("", "", 1);
  _colorearParticipantes(hP, datos.length);

  _alert(
    "🔄 Sincronización completada\n\n" +
    (actualizados.length
      ? "✅ Actualizados (" + actualizados.length + "):\n  " + actualizados.join("\n  ") + "\n\n"
      : "") +
    (sinEncontrar.length
      ? "⚠️ Aún sin perfil (" + sinEncontrar.length + "):\n  " + sinEncontrar.join("\n  ")
      : "Todos los participantes tienen Creamos ID ✅")
  );
}); }

// ── Reinstalar ────────────────────────────────────────────────

function reinstalarSistema() { _run(function() {
  var ui  = SpreadsheetApp.getUi();
  var resp = ui.alert("⚠️  Reinstalar sistema RRHH — BORRADO COMPLETO",
    "Elimina TODO:\n• TODAS las hojas del Spreadsheet\n" +
    "• Carpeta «" + CFG.ORG + " · RRHH» en Drive (Docs, Recibos, Reportes)\n" +
    "• Todos los triggers automáticos\n\n" +
    "Esta acción NO se puede deshacer.\n¿Continuar?", ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var hojas  = ss.getSheets();

  // Google Sheets requiere al menos 1 hoja: preservar la DB de Creamos, borrar el resto
  var temp = ss.insertSheet("_temp_reinstal_");
  hojas.forEach(function(h) {
    if (h.getName() === CFG.HOJAS.CREAMOS_DB) return; // NUNCA borrar la base de datos oficial
    try { ss.deleteSheet(h); } catch(_) {}
  });

  // Borrar carpeta en Drive
  var p = PropertiesService.getScriptProperties();
  var idRaiz = p.getProperty("RRHH_RAIZ");
  if (idRaiz) {
    try { _borrarCarpetaRecursivo(DriveApp.getFolderById(idRaiz)); } catch(_) {}
  } else {
    var it = DriveApp.getFoldersByName(CFG.ORG + " · RRHH");
    while (it.hasNext()) _borrarCarpetaRecursivo(it.next());
  }

  // Borrar triggers y propiedades
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  p.deleteAllProperties();

  Utilities.sleep(500);

  // Recrear hojas del sistema (borra la hoja temporal internamente)
  crearHojas();

  // Borrar la hoja temporal si quedó
  var t2 = ss.getSheetByName("_temp_reinstal_");
  if (t2) try { ss.deleteSheet(t2); } catch(_) {}

  _alert("✅ Sistema reinstalado limpiamente.\n\nTodas las hojas antiguas fueron eliminadas.\nSiguiente paso: ejecuta 'Instalación completa'.");
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
  // Cuenta participantes activas con Educacion o Apoyo_Emocional marcados en PARTICIPANTES
  var hojaP2 = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES);
  if (hojaP2 && hojaP2.getLastRow() > 1) {
    hojaP2.getDataRange().getValues().slice(1).forEach(function(p) {
      var edu   = String(p[7] || "").trim(); // col H = Educacion
      var apoyo = String(p[8] || "").trim(); // col I = Apoyo_Emocional
      if (edu.indexOf("Sí") === 0 || apoyo === "Sí") kpi.horasFormacion++;
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
      id:           String(datos[i][0]||"").trim(),  // col A = Creamos_ID
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
  var resp = ui.alert("🚀 INSTALACIÓN COMPLETA — " + CFG.PROYECTO + " / " + CFG.ORG,
    "Se ejecutarán 6 pasos automáticamente:\n\n" +
    "1 — Crear hojas: PARTICIPANTES, DatosKobo, PERIODOS\n" +
    "2 — Importar datos desde Kobo\n" +
    "3 — Crear estructura en Drive (Docs_Proceso, Reportes)\n" +
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
    ss.toast("Paso 1/6: Creando hojas (PARTICIPANTES, DatosKobo, PERIODOS)...", "🚀", -1);
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
  resumen += "• PARTICIPANTES   — lista maestra con banco, cuenta, categoría\n";
  resumen += "• DatosKobo       — datos de Kobo (se actualiza automático cada hora)\n";
  resumen += "• PERIODOS        — control de quincenas\n";
  resumen += "• DiasEstudio     — qué días estudia cada participante\n";
  resumen += "• ListaTerapias   — quién recibe terapia\n";
  resumen += "• Q_[fecha]       — reporte generado por quincena\n";
  resumen += "\nTarifas: A=Q16.50 | B=Q15.75 | C=Q15.00 | D=Q14.00\n";
  resumen += "IVA 5%: Admin → Configurar IVA (quién tiene factura)\n";
  resumen += "\n✅ Siguiente paso: Admin → Cargar lista oficial (33 participantes)";
  _alert(resumen);
}); }

// ══════════════════════════════════════════════════════════════════
// HOJAS DE APOYO — DiasEstudio y ListaTerapias
// ══════════════════════════════════════════════════════════════════

function crearHojaDiasEstudio() { _run(function() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("DiasEstudio");
  var esNueva = !hoja;
  if (esNueva) hoja = ss.insertSheet("DiasEstudio");

  // Encabezado siempre
  var enc = ["Participante","Lun","Mar","Mié","Jue","Vie","Sáb","Dom","Fecha_Inicio","Fecha_Fin"];
  hoja.getRange(1,1,1,enc.length).setValues([enc])
    .setFontWeight("bold").setBackground("#7b1fa2").setFontColor("#fff").setHorizontalAlignment("center");
  hoja.setFrozenRows(1);

  if (esNueva) {
    // Pre-llenar desde LISTA_OFICIAL (siempre limpio)
    // Orden oficial del taller (por número 1-33)
    var filas = LISTA_OFICIAL.map(function(it) {
      return [it[1], "","","","","","","","",""];
    });

    hoja.getRange(2, 1, filas.length, 10).setValues(filas);

    // Validación X / vacío en cols días (B-H)
    var vX = SpreadsheetApp.newDataValidation().requireValueInList(["X",""],true).build();
    hoja.getRange(2,2,filas.length,7).setDataValidation(vX).setHorizontalAlignment("center");

    // Colores por categoría
    var CAT_BG = { A: CFG.COLORES_CAT.A.bgClaro, B: CFG.COLORES_CAT.B.bgClaro, C: CFG.COLORES_CAT.C.bgClaro, D: CFG.COLORES_CAT.D.bgClaro };
    LISTA_OFICIAL.forEach(function(it, i) {
      hoja.getRange(i+2, 1, 1, 10).setBackground(CAT_BG[it[2]] || "#ffffff");
    });
  }

  hoja.setColumnWidth(1, 260);
  for (var c=2;c<=8;c++) hoja.setColumnWidth(c,55);
  hoja.setColumnWidth(9,110); hoja.setColumnWidth(10,110);
  hoja.activate();

  _alert(
    "📚 DÍAS DE ESTUDIO — Para qué sirve:\n\n" +
    "Marca con X los días que cada participante asiste a estudiar.\n" +
    "Al calcular horas, el sistema marcará esos días como 'Es_Dia_Estudio=Sí'\n" +
    "y aplicará el porcentaje de pago correspondiente.\n\n" +
    "Fecha_Inicio / Fecha_Fin: opcionales, para limitar vigencia (dd/mm/yyyy).\n\n" +
    "Ejemplo: si Sindy estudia Lunes y Miércoles → marca X en Lun y Mié."
  );
}); }

function crearHojaListaTerapias() { _run(function() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("ListaTerapias");
  var esNueva = !hoja;
  if (esNueva) hoja = ss.insertSheet("ListaTerapias");

  // Encabezado siempre
  hoja.getRange(1,1,1,4).setValues([["Creamos_ID","Participante","Recibe Terapia (X)","Notas"]])
    .setFontWeight("bold").setBackground("#00897b").setFontColor("#fff").setHorizontalAlignment("center");
  hoja.setFrozenRows(1);

  if (esNueva) {
    // Pre-llenar desde LISTA_OFICIAL con ID y nombre
    // Orden oficial del taller (por número 1-33)
    var filas = LISTA_OFICIAL.map(function(it) {
      return [it[3] || "", it[1], "", ""];
    });

    hoja.getRange(2, 1, filas.length, 4).setValues(filas);

    var vX = SpreadsheetApp.newDataValidation().requireValueInList(["X",""],true).build();
    hoja.getRange(2,3,filas.length,1).setDataValidation(vX).setHorizontalAlignment("center");
  }

  hoja.setColumnWidth(1,120); hoja.setColumnWidth(2,260);
  hoja.setColumnWidth(3,150); hoja.setColumnWidth(4,300);
  hoja.activate();

  _alert(
    "🧘 LISTA DE TERAPIAS — Para qué sirve:\n\n" +
    "Marca con X a las personas que reciben sesiones de terapia.\n" +
    "Cuando llegue un registro de Kobo con tipo 'Terapia' o 'Salida Terapia',\n" +
    "el sistema lo contabiliza como horas trabajadas (Es_Terapia=Sí).\n\n" +
    "Esto afecta el cálculo de horas del período y los reportes."
  );
}); }

// ══════════════════════════════════════════════════════════════════
// DIAGNÓSTICO, REPARACIÓN Y CAMBIO DE NOMBRE
// ══════════════════════════════════════════════════════════════════

/**
 * Diagnóstico completo de DatosKobo.
 * Genera hoja "Diagnóstico_Kobo" con:
 *  - Entradas sin salida (horas perdidas)
 *  - Salidas sin entrada
 *  - Nombres no reconocidos (slugs sin normalizar)
 *  - Registros sin tipo válido
 *  - Resumen por participante
 */
function diagnosticarDatosKobo() { _run(function() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hojaK = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (!hojaK) { _alert("No existe DatosKobo. Importa primero desde Kobo."); return; }

  var datos  = hojaK.getDataRange().getValues();
  var cols   = detectarColumnas(datos[0], datos.slice(1));
  var mapeo  = cargarMapeoNombres();
  var tz     = Session.getScriptTimeZone();

  // ── 1. Recorrer todos los registros ─────────────────────────────
  var grupos   = {};  // clave "nombre|fecha" → {ent:[], sal:[], nombre, fecha}
  var sinTipo  = [];
  var sinNombre= [];
  var slugs    = [];  // nombres que parecen slugs (con _)

  for (var i=1; i<datos.length; i++) {
    var fila = datos[i];
    var nombreRaw = cols.participante !== undefined ? String(fila[cols.participante]||"").trim() : "";
    if (!nombreRaw) { sinNombre.push(i+1); continue; }

    // Detectar slug sin normalizar
    if (nombreRaw.indexOf("_") !== -1 && nombreRaw.indexOf(" ") === -1) {
      slugs.push({ fila:i+1, valor:nombreRaw });
    }

    var nombre = normalizarNombre(nombreRaw, mapeo);
    var tipo   = obtenerTipoRegistro(fila, cols);

    if (!tipo.esIngreso && !tipo.esEgreso) {
      sinTipo.push({ fila:i+1, nombre:nombre, valor: cols.accionUnificada!==undefined?String(fila[cols.accionUnificada]||""):"?" });
      continue;
    }

    var tsRaw = cols.start !== undefined ? fila[cols.start] : null;
    var ts    = tsRaw instanceof Date ? tsRaw : new Date(tsRaw);
    if (isNaN(ts)) continue;
    var clave = nombre + "|" + _dClave(ts);
    if (!grupos[clave]) grupos[clave] = { nombre:nombre, fecha:ts, ent:[], sal:[] };
    if (tipo.esIngreso) grupos[clave].ent.push(ts);
    if (tipo.esEgreso)  grupos[clave].sal.push(ts);
  }

  // ── 2. Detectar pares rotos ──────────────────────────────────────
  var sinSalida  = [];  // tiene Entrada pero no Salida
  var sinEntrada = [];  // tiene Salida pero no Entrada
  var pareados   = 0;
  var resumenPart= {}; // nombre → {dias:0, diasSinSalida:0, diasSinEntrada:0}

  Object.keys(grupos).forEach(function(k) {
    var g = grupos[k];
    var fechaStr = Utilities.formatDate(g.fecha, tz, "dd/MM/yyyy");
    if (!resumenPart[g.nombre]) resumenPart[g.nombre] = { dias:0, sinSalida:0, sinEntrada:0 };

    if (g.ent.length > 0 && g.sal.length > 0) {
      pareados++;
      resumenPart[g.nombre].dias++;
    } else if (g.ent.length > 0 && g.sal.length === 0) {
      sinSalida.push({ nombre:g.nombre, fecha:fechaStr, horaEntrada: Utilities.formatDate(g.ent[0],tz,"HH:mm") });
      resumenPart[g.nombre].sinSalida++;
    } else if (g.ent.length === 0 && g.sal.length > 0) {
      sinEntrada.push({ nombre:g.nombre, fecha:fechaStr, horaSalida: Utilities.formatDate(g.sal[0],tz,"HH:mm") });
      resumenPart[g.nombre].sinEntrada++;
    }
  });

  // ── 3. Crear hoja de diagnóstico ────────────────────────────────
  var tabName = "Diagnóstico_Kobo";
  var hD = ss.getSheetByName(tabName);
  if (hD) ss.deleteSheet(hD);
  hD = ss.insertSheet(tabName);
  if (hD.getMaxColumns() < 6) hD.insertColumnsAfter(hD.getMaxColumns(), 6-hD.getMaxColumns());

  var filas=[], tipos=[];
  function p(f,t){filas.push(f);tipos.push(t);}
  var ts2 = Utilities.formatDate(new Date(),tz,"dd/MM/yyyy HH:mm");

  p(["DIAGNÓSTICO DE DATOS KOBO — "+CFG.ORG,"","","","",""],"titulo");
  p(["Generado: "+ts2+" | Total registros: "+(datos.length-1)+" | Pares OK: "+pareados,"","","","",""],"sub");
  p(["","","","","",""],"vacio");

  // ── Sección 1: Entradas sin salida (⚠️ horas perdidas) ──────────
  p(["⚠️  ENTRADAS SIN SALIDA — horas no contabilizadas","","","","",""],"sec_warn");
  if (sinSalida.length === 0) {
    p(["✅ Ninguna — todas tienen salida registrada","","","","",""],"ok");
  } else {
    p(["#","Participante","Fecha","Hora entrada","Problema","Acción sugerida"],"enc");
    sinSalida.forEach(function(r,i){
      p([i+1, r.nombre, r.fecha, r.horaEntrada, "Sin salida registrada", "Marcar salida manualmente o usar 'Reparar'"],"warn");
    });
  }
  p(["","","","","",""],"vacio");

  // ── Sección 2: Salidas sin entrada ───────────────────────────────
  p(["⚠️  SALIDAS SIN ENTRADA — posible doble salida o error","","","","",""],"sec_warn");
  if (sinEntrada.length === 0) {
    p(["✅ Ninguna — todas las salidas tienen entrada","","","","",""],"ok");
  } else {
    p(["#","Participante","Fecha","Hora salida","Problema","Acción sugerida"],"enc");
    sinEntrada.forEach(function(r,i){
      p([i+1, r.nombre, r.fecha, r.horaSalida, "Sin entrada previa","Verificar si marcó entrada ese día"],"warn");
    });
  }
  p(["","","","","",""],"vacio");

  // ── Sección 3: Slugs sin normalizar ──────────────────────────────
  p(["⚠️  NOMBRES SIN NORMALIZAR (slug de Kobo)","","","","",""],"sec_warn");
  if (slugs.length === 0) {
    p(["✅ Todos los nombres están normalizados","","","","",""],"ok");
  } else {
    p(["Fila","Valor en DatosKobo","","Acción","",""],"enc");
    slugs.forEach(function(r){
      var normalizado = normalizarNombre(r.valor, mapeo);
      p([r.fila, r.valor, "→", normalizado===r.valor?"❌ No encontrado":normalizado, "",""],"warn");
    });
    p(["","Ejecuta: Configuración → Reparar Datos Kobo","","","",""],"nota");
  }
  p(["","","","","",""],"vacio");

  // ── Sección 4: Sin tipo reconocido ───────────────────────────────
  if (sinTipo.length > 0) {
    p(["⚠️  REGISTROS SIN TIPO RECONOCIDO (ni Entrada ni Salida)","","","","",""],"sec_warn");
    p(["Fila","Participante","Valor detectado","","",""],"enc");
    sinTipo.forEach(function(r){
      p([r.fila, r.nombre, '"'+r.valor+'"','','',''],"warn");
    });
    p(["","","","","",""],"vacio");
  }

  // ── Sección 5: Resumen por participante ──────────────────────────
  p(["📊  RESUMEN POR PARTICIPANTE","","","","",""],"sec_ok");
  p(["Participante","Días completos","Sin salida","Sin entrada","Estado",""],"enc");
  Object.keys(resumenPart).sort().forEach(function(nombre){
    var r = resumenPart[nombre];
    var estado = (r.sinSalida===0 && r.sinEntrada===0) ? "✅ OK" : "⚠️ Revisar";
    p([nombre, r.dias, r.sinSalida||"", r.sinEntrada||"", estado,""],
      (r.sinSalida>0||r.sinEntrada>0)?"warn":"ok");
  });

  // ── Escribir hoja ─────────────────────────────────────────────────
  hD.getRange(1,1,filas.length,6).setValues(filas);

  var COLORES = { titulo:"#1a237e", sub:"#e8eaf6", vacio:"#ffffff", sec_warn:"#e65100",
                  sec_ok:"#1b5e20", enc:"#37474f", warn:"#fff3e0", ok:"#f1f8e9", nota:"#fce4ec" };
  tipos.forEach(function(t,i){
    var r = hD.getRange(i+1,1,1,6);
    r.setFontFamily("Arial").setFontSize(10);
    if (t==="titulo") r.merge().setFontSize(13).setFontWeight("bold").setBackground(COLORES.titulo).setFontColor("#fff").setHorizontalAlignment("center");
    else if (t==="sub") r.merge().setBackground(COLORES.sub).setFontColor("#283593").setHorizontalAlignment("center");
    else if (t==="sec_warn") r.merge().setFontWeight("bold").setBackground(COLORES.sec_warn).setFontColor("#fff");
    else if (t==="sec_ok")  r.merge().setFontWeight("bold").setBackground(COLORES.sec_ok).setFontColor("#fff");
    else if (t==="enc")  r.setFontWeight("bold").setBackground(COLORES.enc).setFontColor("#fff");
    else if (t==="warn") r.setBackground(COLORES.warn);
    else if (t==="ok")   r.setBackground(COLORES.ok).setFontColor("#1b5e20");
    else if (t==="nota") r.merge().setFontStyle("italic").setBackground(COLORES.nota).setFontColor("#880e4f");
  });

  hD.setColumnWidth(1,50); hD.setColumnWidth(2,250); hD.setColumnWidth(3,100);
  hD.setColumnWidth(4,280); hD.setColumnWidth(5,160); hD.setColumnWidth(6,180);
  hD.activate();

  _alert(
    "✅ Diagnóstico generado en hoja 'Diagnóstico_Kobo'\n\n" +
    "Resumen:\n" +
    "• Pares OK:           " + pareados + " días\n" +
    "• Sin salida:         " + sinSalida.length + (sinSalida.length ? " ⚠️" : " ✅") + "\n" +
    "• Sin entrada:        " + sinEntrada.length + (sinEntrada.length ? " ⚠️" : " ✅") + "\n" +
    "• Slugs sin normalizar: " + slugs.length + (slugs.length ? " ⚠️" : " ✅") + "\n" +
    "• Sin tipo válido:    " + sinTipo.length + (sinTipo.length ? " ⚠️" : " ✅") + "\n\n" +
    (sinSalida.length ? "Para recuperar horas perdidas: 'Reparar Datos Kobo'" : "")
  );
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
