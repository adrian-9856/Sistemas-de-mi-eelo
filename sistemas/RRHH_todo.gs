// ============================================================
// SISTEMA RRHH — mi eelo
// ============================================================

const CFG = {
  ORG:          "mi eelo",   // programa
  PROYECTO:     "Textil",    // proyecto padre
  CORREO_ADMIN:    "adrian@creamosguatemala.org",
  CORREO_CHEQUES:  "adrian@creamosguatemala.org",  // encargado/a de cheques — cambiar cuando esté listo
  CORREO_PLANILLA: "adrian@creamosguatemala.org",  // encargado/a de planilla transferencias — cambiar cuando esté listo
  DRIVE_FOLDER_PAGOS: "Pagos RRHH mi eelo",        // carpeta en Drive donde se guardan los PDFs
  // Tarifas por categoría (Q por hora) — A=Q16.50 B=Q15.75 C=Q15.00 D=Q14.00
  CATEGORIAS:   { A: 16.50, B: 15.75, C: 15.00, D: 14.00 },
  // Colores oficiales por categoría — usados en TODOS los reportes y hojas
  COLORES_CAT: {
    A: { bg: "#2e7d32", fg: "#ffffff", bgClaro: "#c8e6c9" }, // verde
    B: { bg: "#f9a825", fg: "#212121", bgClaro: "#fff9c4" }, // amarillo/dorado
    C: { bg: "#e65100", fg: "#ffffff", bgClaro: "#ffe0b2" }, // naranja
    D: { bg: "#c62828", fg: "#ffffff", bgClaro: "#ffcdd2" }  // rojo/rosado
  },
  IVA_PCT:      0.05,  // 5% Pequeño Contribuyente Guatemala (solo quien tiene factura)
  HORAS_JORNADA_NORMAL: 7,
  KOBO_URL_CSV: "https://kf.kobotoolbox.org/api/v2/assets/agi395bJj6ojXJzPPDT9n6/export-settings/esFyGoVugvB2pNtpgngLSGD/data.csv",
  KOBO_URL_ESTIPENDIO: "https://kf.kobotoolbox.org/api/v2/assets/aXKv6g3zzpTN2byxajnxyp/export-settings/es5ZJRFNHmd8EDEQdoDu5jj/data.csv",
  KOBO_TIPO_ENTRADA: "🟢 Entrada",   // valor normalizado interno (no el label de Kobo)
  KOBO_TIPO_SALIDA:  "🔴 Salida",    // Kobo exporta "Entrada"/"Salida" — detectarColumnas detecta ambos
  HOJAS: {
    PARTICIPANTES: "PARTICIPANTES",
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
  TIMEZONE: "America/Guatemala",
};

/*
 PARTICIPANTES — 23 columnas (A–W)
 A  Creamos_ID (0)           I  Etapa (8)             — Inscritx/Retiradx/Empleadx/Ciclo de Vida Terminado
 B  Nombre (1)               J  Educacion (9)
 C  Fecha_Nacimiento (2)     K  Apoyo_Emocional (10)
 D  Edad (3)                 L  Inclusion_Laboral (11)
 E  Genero (4)               M  Categoria (12)        — A/B/C/D
 F  Ano_Entrada_Creamos (5)  N  Tarifa_Hora (13)      — Q/hr (auto desde Categoria)
 G  Proyecto (6)             O  Tiene_Factura (14)    — Sí/No (aplica IVA 5%)
 H  Programa (7)             P  DPI (15)
                             Q  NIT (16)
                             R  Correo (17)
                             S  Banco (18)
                             T  Tipo_Cuenta (19)
                             U  Num_Cuenta (20)
                             V  Forma_Pago (21)
                             W  URL_Doc_Proceso (22)
 C-F se llenan automático desde "Copy of CREAMOS ID nuevo" vía sincronizarDesdeCreamos()
*/

// ── Helpers ──────────────────────────────────────────────────

function _sh(nombre) {
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!h) throw new Error('Hoja "' + nombre + '" no encontrada. Ejecuta "Instalación completa" desde el menú 👥 RRHH.');
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

// ── Helpers de sincronización entre hojas ────────────────────

/**
 * Lee PARTICIPANTES y devuelve un mapa {nombre: {id, cat}} con IDs y categorías actuales.
 * Si PARTICIPANTES no existe o está vacía, cae back a LISTA_OFICIAL.
 */
function _mapaDatosParticipantes(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var mapa = {};
  var hP = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES);
  if (hP && hP.getLastRow() >= 2) {
    var datos = hP.getRange(2, 1, hP.getLastRow() - 1, 13).getValues(); // A–M
    datos.forEach(function(f) {
      var nombre = String(f[1] || "").trim();
      if (!nombre) return;
      var id  = _esCreamos_ID_real(String(f[0] || "").trim()) ? String(f[0]).trim() : "";
      var cat = String(f[12] || "").trim().toUpperCase(); // col M = Categoria (idx 12)
      mapa[nombre] = { id: id, cat: cat };
    });
  }
  // Fallback: completar con LISTA_OFICIAL para quienes no están en PARTICIPANTES
  LISTA_OFICIAL.forEach(function(it) {
    var nombre = it[1];
    if (!mapa[nombre]) mapa[nombre] = { id: it[3] || "", cat: it[2] || "" };
  });
  return mapa;
}

/**
 * Colorea todas las filas de una hoja auxiliar según la categoría del participante.
 * @param {Sheet} hoja
 * @param {number} colNombre  índice 1-based de la columna con el nombre
 * @param {number} nCols      total de columnas a colorear
 * @param {Object} mapaDatos  resultado de _mapaDatosParticipantes()
 */
function _colorearHojaApoyo(hoja, colNombre, nCols, mapaDatos) {
  if (!hoja || hoja.getLastRow() < 2) return;
  var nombres = hoja.getRange(2, colNombre, hoja.getLastRow() - 1, 1).getValues();
  nombres.forEach(function(f, i) {
    var nombre = String(f[0] || "").trim();
    if (!nombre) return;
    var cat    = (mapaDatos[nombre] || {}).cat || "";
    var color  = (CFG.COLORES_CAT[cat] || {}).bgClaro || "#ffffff";
    hoja.getRange(i + 2, 1, 1, nCols).setBackground(color);
  });
}

/**
 * Actualiza Creamos_ID (col A) en una hoja auxiliar leyendo desde mapaDatos.
 * @param {Sheet} hoja
 * @param {number} colNombre  índice 1-based de la columna con el nombre
 * @param {Object} mapaDatos
 */
function _actualizarIDsEnHoja(hoja, colNombre, mapaDatos) {
  if (!hoja || hoja.getLastRow() < 2) return;
  var filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, colNombre).getValues();
  filas.forEach(function(f, i) {
    var nombre = String(f[colNombre - 1] || "").trim();
    if (!nombre) return;
    var id = (mapaDatos[nombre] || {}).id || "";
    var actual = String(f[0] || "").trim();
    if (id && actual !== id) hoja.getRange(i + 2, 1).setValue(id);
  });
}

/**
 * Re-aplica colores e IDs en todas las hojas auxiliares existentes.
 * Llámalo después de cargar la lista, sincronizar IDs, o cambiar categorías.
 */
function actualizarColoresYIDs() { _run(function() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var mapa = _mapaDatosParticipantes(ss);
  var log  = [];

  var hDE = ss.getSheetByName("DiasEstudio");
  if (hDE && hDE.getLastRow() >= 2) {
    _actualizarIDsEnHoja(hDE, 2, mapa);
    _colorearHojaApoyo(hDE, 2, 11, mapa);
    log.push("📚 DiasEstudio: IDs y colores actualizados");
  }

  var hLT = ss.getSheetByName("ListaTerapias");
  if (hLT && hLT.getLastRow() >= 2) {
    _actualizarIDsEnHoja(hLT, 2, mapa);
    _colorearHojaApoyo(hLT, 2, 4, mapa);
    log.push("🧘 ListaTerapias: IDs y colores actualizados");
  }

  var hIL = ss.getSheetByName("InclusionLaboral");
  if (hIL && hIL.getLastRow() >= 2) {
    _actualizarIDsEnHoja(hIL, 2, mapa);
    _colorearHojaApoyo(hIL, 2, 4, mapa);
    log.push("💼 InclusionLaboral: IDs y colores actualizados");
  }

  var hHC = ss.getSheetByName("HijosCCI");
  if (hHC && hHC.getLastRow() >= 2) {
    _actualizarIDsEnHoja(hHC, 2, mapa);
    _colorearHojaApoyo(hHC, 2, 5, mapa);
    log.push("👶 HijosCCI: IDs y colores actualizados");
  }

  var hBon = ss.getSheetByName("Bonos");
  if (hBon && hBon.getLastRow() >= 2) {
    log.push("💵 Bonos: hoja presente (" + (hBon.getLastRow()-1) + " registros)");
  }

  _alert("✅ Sincronización completa\n\n" + log.join("\n") +
    "\n\nLos colores y Creamos IDs están al día en todas las hojas.");
}); }

// ── Menú ─────────────────────────────────────────────────────

function onOpen() {
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch(_) { return; }

  var menuAsistencia = ui.createMenu("📥 Asistencia")
    .addItem("📥 Importar desde Kobo",                   "importarDesdeKobo")
    .addSeparator()
    .addItem("🔍 Diagnosticar registros Kobo",           "diagnosticarDatosKobo")
    .addItem("🟡 Marcar filas sospechosas",              "marcarFilasSospechosas")
    .addItem("🗑️ Eliminar filas marcadas en rojo",       "eliminarFilasRojas")
    .addItem("🔧 Reparar datos Kobo",                    "repararDatosKobo")
    .addItem("🧹 Eliminar columnas innecesarias",         "eliminarColumnasKobo")
    .addSeparator()
    .addItem("🧽 Limpiar datos antes del período activo", "limpiarDatosKoboAntesDePeriodo")
    .addItem("🧽 Eliminar registros anteriores a 2026",   "limpiarFilasAnteriores2026")
    .addSeparator()
    .addItem("🔁 Limpiar y reimportar DatosKobo",        "reimportarTodoDesdeKobo");

  // ══════════════════════════════════════════════════════════
  // BLOQUE 2: QUINCENA
  // ══════════════════════════════════════════════════════════
  var menuQuincena = ui.createMenu("📅 Quincena")
    .addItem("📊 Ver / actualizar quincena actual",      "verQuincenaActual")
    .addItem("📋 Actualizar detalle de quincena",        "actualizarDetalleQuincena")
    .addItem("✅ Cerrar y abrir siguiente",               "cerrarQuincenaYCrearSiguiente")
    .addSeparator()
    .addItem("🗓️ Nueva quincena (manual)",               "configurarNuevaQuincena")
    .addItem("🔍 Reporte de quincena pasada",            "generarReporteQuincenaPasada")
    .addItem("📊 Reporte por rango de fechas",           "generarReportePorRango")
    .addSeparator()
    .addItem("💳 Registrar pagos de quincena",           "registrarPagosQuincena")
    .addSeparator()
    .addItem("🩺 Diagnosticar cálculo de horas",         "diagnosticarCalculoHoras")
    .addItem("🔎 Detalle participantes en Kobo",         "diagnosticarParticipantesKobo");

  // ══════════════════════════════════════════════════════════
  // BLOQUE 3: ADMIN — submenús para mantener orden
  // ══════════════════════════════════════════════════════════

  // Submenú: Participantes (cambios individuales)
  var subParticipantes = ui.createMenu("👤 Participantes")
    .addItem("➕ Nuevo participante",                     "nuevoParticipante")
    .addItem("🔼 Cambiar categoría",                     "cambiarCategoriaParticipante")
    .addItem("✏️ Cambiar nombre",                        "cambiarNombreParticipante")
    .addItem("🗑️ Eliminar participante completo",         "eliminarParticipanteCompleto")
    .addSeparator()
    .addItem("📄 Generar DP (fila activa)",              "generarDpFilaActiva")
    .addItem("👥 Directorio de participantes",           "generarDirectorioParticipantes");

  // Submenú: Facturación y pagos
  var subFacturacion = ui.createMenu("🧾 Facturación")
    .addItem("🧾 Configurar IVA (quién tiene factura)",  "configurarFacturacion")
    .addItem("✅ Guardar cambios de facturación",         "aplicarCambiosFacturacion")
    .addSeparator()
    .addItem("🟣 Importar Estipendio desde Kobo",         "importarEstipendioDesdeKobo");

  // Submenú: Hojas del sistema (creación/recreación — uso ocasional)
  var subHojas = ui.createMenu("📋 Crear hojas")
    .addItem("📚 Días de estudio",                       "crearHojaDiasEstudio")
    .addItem("🧘 Lista de terapias",                     "crearHojaListaTerapias")
    .addItem("💼 Inclusión Laboral",                     "crearHojaInclusionLaboral")
    .addItem("🔴 Retiradx",                              "crearHojaRetiradx")
    .addItem("🔵 CiclosVida",                            "crearHojaCiclosVida")
    .addItem("📊 Historial de Quincenas",                "crearHojaHistorialQuincenas")
    .addItem("💵 Bonos",                                 "crearHojaBonos")
    .addItem("🟣 Estipendio",                            "crearHojaEstipendio")
    .addItem("🏦 Cheques",                               "crearHojaCheques")
    .addItem("🔄 Transferencias",                        "crearHojaTransferencias")
    .addItem("👶 Hijos CCI",                             "crearHojaHijosCCI");

  // Submenú: Mantenimiento / setup (raro uso en sistema activo)
  var subMant = ui.createMenu("🔧 Mantenimiento")
    .addItem("📋 Cargar lista oficial",                  "cargarListaParticipantes")
    .addItem("🔄 Sincronizar desde Creamos DB",          "sincronizarDesdeCreamos")
    .addItem("📲 Sincronizar IDs desde DatosKobo",       "sincronizarIDsDesdeKobo")
    .addItem("🔍 Diagnosticar IDs no encontrados",       "diagnosticarBusquedaCreamos")
    .addSeparator()
    .addItem("🔗 Sincronizar participación",             "sincronizarParticipacion")
    .addItem("🎨 Actualizar colores e IDs",              "actualizarColoresYIDs")
    .addItem("🧹 Limpiar duplicados en PARTICIPANTES",   "limpiarDuplicadosParticipantes")
    .addItem("📄 Actualizar todos los DPs",              "actualizarTodosLosDps")
    .addSeparator()
    .addItem("📊 Cargar historial de quincenas pasadas",  "backfillHistorialQuincenas")
    .addSeparator()
    .addItem("⚡ Activar automatizaciones",              "configurarTriggers")
    .addItem("⬆️ Migrar sistema",                        "migrarSistema")
    .addItem("🔽 Reparar dropdowns",                     "repararDropdownsParticipantes")
    .addItem("🔢 Reparar formato DPI y NIT",             "repararFormatoDPI")
    .addItem("🔄 Recalcular tarifas",                    "recalcularTarifas")
    .addItem("📤 Exportar para PowerBI",                 "exportarParaPowerBI")
    .addItem("📖 Guía de Uso",                           "crearGuiaUso")
    .addSeparator()
    .addItem("🗑️ Reinstalar sistema (borra TODO)",       "reinstalarSistema");

  // Admin principal — solo lo de uso frecuente
  var menuAdmin = ui.createMenu("⚙️ Admin")
    .addSubMenu(subParticipantes)
    .addSubMenu(subFacturacion)
    .addSeparator()
    .addItem("🔄 Sincronizar desde Creamos DB",          "sincronizarDesdeCreamos")
    .addItem("🎨 Actualizar colores e IDs",              "actualizarColoresYIDs")
    .addSeparator()
    .addItem("📊 Actualizar Dashboard Visual",           "actualizarDashboardVisual")
    .addSeparator()
    .addSubMenu(subHojas)
    .addSubMenu(subMant);

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

  try { actualizarDashboardVisual(); } catch(_) {}
}

// onEdit: col I = Categoria → auto-llenar Tarifa_Hora en PARTICIPANTES
//         col I = Etapa → si "Retiradx" dispara flujo de retiro
//         DiasEstudio cols C-I → Educacion (col J) en PARTICIPANTES  (A=ID, B=Nombre, C-I=días)
//         ListaTerapias col C  → Apoyo_Emocional (col K) en PARTICIPANTES
//         InclusionLaboral col C → Inclusion_Laboral (col L) en PARTICIPANTES
//         HijosCCI cols C-D → datos propios de hoja auxiliar (no se reflejan en PARTICIPANTES)
function onEdit(e) {
  var sheet = e.range.getSheet();
  var nombre = sheet.getName();
  var col    = e.range.getColumn();
  var fila   = e.range.getRow();

  // PARTICIPANTES — Categoria (col M=13) cambia → auto-llenar Tarifa_Hora (col N=14)
  if (nombre === CFG.HOJAS.PARTICIPANTES && col === 13 && fila >= 2) {
    var cat = String(e.range.getValue()).trim().toUpperCase();
    var tarifa = CFG.CATEGORIAS[cat];
    if (tarifa) sheet.getRange(fila, 14).setValue(tarifa);
  }

  // PARTICIPANTES — Etapa (col I=9) cambia a "Retiradx" o "Ciclo de Vida Terminado"
  if (nombre === CFG.HOJAS.PARTICIPANTES && col === 9 && fila >= 2) {
    var etapa = String(e.range.getValue()).trim();
    if (etapa === "Retiradx") {
      try { _iniciarRetiro(sheet, fila); } catch(err) { Logger.log("Error retiro: " + err.message); }
    } else if (etapa === "Ciclo de Vida Terminado") {
      try { _registrarCicloVida(sheet, fila); } catch(err) { Logger.log("Error ciclo vida: " + err.message); }
    }
  }

  // DIASESTUDIO — cualquier día (cols C-I = 3-9; A=ID, B=Nombre) → Educacion en PARTICIPANTES
  if (nombre === "DiasEstudio" && col >= 3 && col <= 9 && fila >= 2) {
    try { _syncEducacionFila(sheet, fila); } catch(_) {}
  }

  // LISTATERAPIAS — col C (Recibe Terapia) → Apoyo_Emocional en PARTICIPANTES
  if (nombre === "ListaTerapias" && col === 3 && fila >= 2) {
    try { _syncApoyoFila(sheet, fila); } catch(_) {}
  }

  // INCLUSIONLABORAL — col C (Participa) → Inclusion_Laboral en PARTICIPANTES
  if (nombre === "InclusionLaboral" && col === 3 && fila >= 2) {
    try { _syncInclusionFila(sheet, fila); } catch(_) {}
  }

  // HIJOSCCI — datos propios de la hoja (no se reflejan en PARTICIPANTES, esquema 19 cols)
  // _syncHijosCCIFila es no-op, se conserva para compatibilidad

}

/**
 * onEdit INSTALABLE — maneja status de pagos (Cheques/Transferencias).
 * Necesita trigger instalable (configurarTriggers) para usar MailApp y UI.
 */
function onEditInstalable(e) {
  try {
    var sheet  = e.range.getSheet();
    var nombre = sheet.getName();
    var col    = e.range.getColumn();
    var fila   = e.range.getRow();
    if (fila < 2) return;

    // CHEQUES — col K (Status=11) → "Cobrado"
    if (nombre === "Cheques" && col === 11) {
      var statusChq = String(e.range.getValue()).trim();
      if (statusChq === "Cobrado") {
        _verificarYArchivarPagos("Cheques");
      }
    }

    // TRANSFERENCIAS — col M (Status=13) → "Transferencias Subidas"
    if (nombre === "Transferencias" && col === 13) {
      var statusTr = String(e.range.getValue()).trim();
      if (statusTr === "Transferencias Subidas") {
        _enviarEmailPago("Transferencias");
        _verificarYArchivarPagos("Transferencias");
      }
    }
  } catch(err) {
    Logger.log("onEditInstalable error: " + err.message);
  }
}

/**
 * Verifica si todos los registros activos en Cheques (Status=col 11) o
 * Transferencias (Status=col 13) están en estado final.
 * Si sí → archiva la hoja con nombre "Cheques_MesAño" y crea hoja nueva con headers.
 */
function _verificarYArchivarPagos(nombreHoja) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hoja  = ss.getSheetByName(nombreHoja);
  if (!hoja || hoja.getLastRow() < 2) return;

  var esCheque  = (nombreHoja === "Cheques");
  var colStatus = esCheque ? 11 : 13;  // K o M
  var estadoFinal = esCheque ? "Cobrado" : "Transferencias Subidas";

  var datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, colStatus).getValues();
  var totalActivos = 0, totalFinales = 0;
  var mesArchivo = "", anioArchivo = "";

  datos.forEach(function(r) {
    var nombre = String(r[0] || "").trim();
    if (!nombre) return;
    totalActivos++;
    var status = String(r[colStatus - 1] || "").trim();
    if (status === estadoFinal) totalFinales++;
    // Tomar mes/año de la primera fila con datos
    if (!mesArchivo && r[esCheque ? 2 : 10]) {
      mesArchivo  = String(r[esCheque ? 2 : 10] || "").trim();
      anioArchivo = String(r[esCheque ? 3 : 11] || "").trim();
    }
  });

  if (totalActivos === 0 || totalFinales < totalActivos) return; // aún hay pendientes

  // Todos en estado final → archivar
  var sufijo = (mesArchivo || "mes") + "_" + (anioArchivo || new Date().getFullYear());
  var nombreArchivo = nombreHoja + "_" + sufijo;

  // Crear copia archivada
  var copia = hoja.copyTo(ss);
  copia.setName(nombreArchivo);
  copia.setTabColor("#78909c"); // gris = archivado

  // Limpiar hoja original (conservar encabezado)
  if (hoja.getLastRow() > 1) hoja.deleteRows(2, hoja.getLastRow() - 1);

  SpreadsheetApp.getUi().alert(
    "✅ PAGOS ARCHIVADOS\n\n" +
    "📋 Hoja archivada como: " + nombreArchivo + "\n" +
    "📄 Hoja '" + nombreHoja + "' lista para el próximo mes."
  );
}

/**
 * Envía email de notificación cuando los pagos pasan a estado final.
 * tipo: "Cheques" → CORREO_CHEQUES, "Transferencias" → CORREO_PLANILLA
 */
function _enviarEmailPago(tipo) {
  var dest = (tipo === "Cheques") ? CFG.CORREO_CHEQUES : CFG.CORREO_PLANILLA;
  if (!dest) return;
  var asunto = "[" + CFG.ORG + "] " + tipo + " listos para procesar";
  var cuerpo = "Hola,\n\nLos " + tipo + " del período están listos para ser procesados.\n\n" +
    "Por favor revisa la hoja '" + tipo + "' en el archivo de RRHH mi eelo.\n\n" +
    "Saludos,\n" + CFG.ORG;
  try {
    MailApp.sendEmail({ to: dest, subject: asunto, body: cuerpo });
  } catch(e) { Logger.log("Email error: " + e.message); }
}

/** Sincroniza la fila de DiasEstudio hacia col F (Educacion=6) de PARTICIPANTES */
function _syncEducacionFila(hDE, fila) {
  var fila_ = hDE.getRange(fila, 1, 1, 9).getValues()[0]; // A=ID, B=Nombre, C-I=días
  var participante = String(fila_[1] || "").trim(); // col B = Nombre
  if (!participante) return;

  var DIAS_NOM = ["Lun","Mar","Mié","Jue","Vie","Sáb","Dom"];
  var activos = [];
  for (var d = 0; d < 7; d++) {
    if (String(fila_[d + 2] || "").trim().toUpperCase() === "X") activos.push(DIAS_NOM[d]); // días en cols C-I (índice 2-8)
  }
  var texto = activos.length > 0 ? "Sí — " + activos.join(", ") : "No";
  _actualizarColParticipante(participante, 10, texto); // col J = Educacion (idx 9, col 10)
}

/** Sincroniza la fila de ListaTerapias hacia col G (Apoyo_Emocional=7) de PARTICIPANTES */
function _syncApoyoFila(hLT, fila) {
  var fila_ = hLT.getRange(fila, 1, 1, 3).getValues()[0];
  var participante = String(fila_[1] || "").trim(); // col B = Participante
  if (!participante) return;
  var texto = String(fila_[2] || "").trim().toUpperCase() === "X" ? "Sí" : "No";
  _actualizarColParticipante(participante, 11, texto); // col K = Apoyo_Emocional
}

/** Sincroniza la fila de InclusionLaboral hacia col H (Inclusion_Laboral=8) de PARTICIPANTES */
function _syncInclusionFila(hIL, fila) {
  var fila_ = hIL.getRange(fila, 1, 1, 3).getValues()[0];
  var participante = String(fila_[1] || "").trim(); // col B = Participante
  if (!participante) return;
  var texto = String(fila_[2] || "").trim().toUpperCase() === "X" ? "Sí" : "No";
  _actualizarColParticipante(participante, 12, texto); // col L = Inclusion_Laboral
}

/** HijosCCI ya no se almacena en PARTICIPANTES (esquema 19 cols) — función conservada por compatibilidad */
function _syncHijosCCIFila(hHC, fila) {
  // Hijos_CCI y Num_Hijos_CCI fueron eliminados de PARTICIPANTES en el esquema de 19 cols.
  // Los datos de HijosCCI viven únicamente en la hoja auxiliar HijosCCI.
  return;
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
  var hIL = ss.getSheetByName("InclusionLaboral");
  var hHC = ss.getSheetByName("HijosCCI");
  var hP  = _sh(CFG.HOJAS.PARTICIPANTES);
  if (!hP || hP.getLastRow() < 2) { _alert("No hay participantes cargados."); return; }

  var DIAS_NOM = ["Lun","Mar","Mié","Jue","Vie","Sáb","Dom"];
  var actDE = 0, actLT = 0, actIL = 0, actHC = 0;

  // DiasEstudio → Educacion (col J = 10)  — A=ID, B=Nombre, C-I=días
  if (hDE && hDE.getLastRow() >= 2) {
    var datDE = hDE.getRange(2, 1, hDE.getLastRow() - 1, 9).getValues();
    datDE.forEach(function(f) {
      var participante = String(f[1] || "").trim(); // col B = Nombre
      if (!participante) return;
      var activos = [];
      for (var d = 0; d < 7; d++) {
        if (String(f[d + 2] || "").trim().toUpperCase() === "X") activos.push(DIAS_NOM[d]);
      }
      var texto = activos.length > 0 ? "Sí — " + activos.join(", ") : "No";
      _actualizarColParticipante(participante, 10, texto);
      if (activos.length) actDE++;
    });
  }

  // ListaTerapias → Apoyo_Emocional (col K = 11)
  if (hLT && hLT.getLastRow() >= 2) {
    var datLT = hLT.getRange(2, 1, hLT.getLastRow() - 1, 3).getValues();
    datLT.forEach(function(f) {
      var participante = String(f[1] || "").trim(); // col B
      if (!participante) return;
      var texto = String(f[2] || "").trim().toUpperCase() === "X" ? "Sí" : "No";
      _actualizarColParticipante(participante, 11, texto);
      if (texto === "Sí") actLT++;
    });
  }

  // InclusionLaboral → Inclusion_Laboral (col L = 12)
  if (hIL && hIL.getLastRow() >= 2) {
    var datIL = hIL.getRange(2, 1, hIL.getLastRow() - 1, 3).getValues();
    datIL.forEach(function(f) {
      var participante = String(f[1] || "").trim(); // col B
      if (!participante) return;
      var texto = String(f[2] || "").trim().toUpperCase() === "X" ? "Sí" : "No";
      _actualizarColParticipante(participante, 12, texto);
      if (texto === "Sí") actIL++;
    });
  }

  // HijosCCI: datos se mantienen en hoja auxiliar, no se sincronizan a PARTICIPANTES (esquema 19 cols)
  if (hHC && hHC.getLastRow() >= 2) {
    actHC = hHC.getLastRow() - 1; // solo contar
  }

  // Actualizar colores e IDs en todas las hojas auxiliares
  var mapa = _mapaDatosParticipantes(ss);
  var hDE2 = ss.getSheetByName("DiasEstudio");
  var hLT2 = ss.getSheetByName("ListaTerapias");
  var hIL2 = ss.getSheetByName("InclusionLaboral");
  var hHC2 = ss.getSheetByName("HijosCCI");
  if (hDE2) { _actualizarIDsEnHoja(hDE2, 2, mapa); _colorearHojaApoyo(hDE2, 2, 11, mapa); }
  if (hLT2) { _actualizarIDsEnHoja(hLT2, 2, mapa); _colorearHojaApoyo(hLT2, 2, 4, mapa); }
  if (hIL2) { _actualizarIDsEnHoja(hIL2, 2, mapa); _colorearHojaApoyo(hIL2, 2, 4, mapa); }
  if (hHC2) { _actualizarIDsEnHoja(hHC2, 2, mapa); _colorearHojaApoyo(hHC2, 2, 5, mapa); }

  _alert(
    "✅ Participación sincronizada\n\n" +
    "📚 Educación (días de estudio): " + actDE + " con asistencia marcada\n" +
    "🧘 Apoyo Emocional (terapias): " + actLT + " con terapia marcada\n" +
    "💼 Inclusión Laboral: " + actIL + " participando\n" +
    "👶 Hijos CCI (hoja auxiliar): " + actHC + " filas registradas\n\n" +
    "🎨 Colores e IDs actualizados en todas las hojas"
  );
}); }

// ── Crear hojas ───────────────────────────────────────────────

function crearHojas() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // (CLASIFICACION eliminada — la info de categorías está en el Directorio)

  // PARTICIPANTES — 23 cols (A–W)
  // A=Creamos_ID, B=Nombre, C=Fecha_Nacimiento, D=Edad, E=Genero, F=Ano_Entrada_Creamos,
  // G=Proyecto, H=Programa, I=Etapa, J=Educacion, K=Apoyo_Emocional, L=Inclusion_Laboral,
  // M=Categoria, N=Tarifa_Hora, O=Tiene_Factura, P=DPI, Q=NIT, R=Correo,
  // S=Banco, T=Tipo_Cuenta, U=Num_Cuenta, V=Forma_Pago, W=URL_Doc_Proceso
  var hP = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES) || ss.insertSheet(CFG.HOJAS.PARTICIPANTES);
  var esNuevaP = hP.getLastRow() === 0;
  if (esNuevaP) {
    hP.appendRow([
      "Creamos_ID","Nombre","Fecha_Nacimiento","Edad","Genero","Ano_Entrada_Creamos",
      "Proyecto","Programa","Etapa",
      "Educacion","Apoyo_Emocional","Inclusion_Laboral",
      "Categoria","Tarifa_Hora","Tiene_Factura",
      "DPI","NIT","Correo",
      "Banco","Tipo_Cuenta","Num_Cuenta","Forma_Pago","URL_Doc_Proceso"
    ]);
  }
  _fmtEnc(hP, "#639922");
  var vEtapa   = SpreadsheetApp.newDataValidation().requireValueInList(["Inscritx","Retiradx","Empleadx","Ciclo de Vida Terminado"],true).build();
  var vCat     = SpreadsheetApp.newDataValidation().requireValueInList(["A","B","C","D"],true).build();
  var vSiNo    = SpreadsheetApp.newDataValidation().requireValueInList(["Sí","No"],true).build();
  var vBanco   = SpreadsheetApp.newDataValidation().requireValueInList([
    "Banrural","Industrial","BAC Credomatic","G&T Continental","Banco Azteca","N/A","Otro"
  ],true).build();
  var vTipoCta = SpreadsheetApp.newDataValidation().requireValueInList(["Monetaria","Ahorro",""],true).build();
  var vPago    = SpreadsheetApp.newDataValidation().requireValueInList(["Transferencia","Cheque"],true).build();
  hP.getRange("I2:I500").setDataValidation(vEtapa);   // col I = Etapa
  hP.getRange("M2:M500").setDataValidation(vCat);     // col M = Categoria
  hP.getRange("O2:O500").setDataValidation(vSiNo);    // col O = Tiene_Factura
  hP.getRange("S2:S500").setDataValidation(vBanco);   // col S = Banco
  hP.getRange("T2:T500").setDataValidation(vTipoCta); // col T = Tipo_Cuenta
  hP.getRange("V2:V500").setDataValidation(vPago);    // col V = Forma_Pago
  hP.getRange("N2:N500").setNumberFormat("Q#,##0.00"); // col N = Tarifa_Hora
  hP.getRange("C2:C500").setNumberFormat("dd/MM/yyyy"); // col C = Fecha_Nacimiento
  hP.getRange("P2:P500").setNumberFormat("@");          // col P = DPI (texto plano)
  hP.getRange("Q2:Q500").setNumberFormat("@");          // col Q = NIT (texto plano)
  if (esNuevaP) {
    hP.setColumnWidth(2, 220);  // Nombre
    hP.setColumnWidth(3, 110);  // Fecha_Nacimiento
    hP.setColumnWidth(4, 60);   // Edad
    hP.setColumnWidth(5, 80);   // Genero
    hP.setColumnWidth(6, 80);   // Ano_Entrada
    hP.setColumnWidth(13, 100); // Categoria
    hP.setColumnWidth(19, 120); // Banco
    hP.setColumnWidth(20, 100); // Tipo_Cuenta
    hP.setColumnWidth(21, 130); // Num_Cuenta
    hP.setColumnWidth(22, 120); // Forma_Pago
    hP.setColumnWidth(23, 300); // URL_Doc_Proceso
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
    "• PARTICIPANTES (19 cols — Etapa, Banco, Cuenta, sin Hijos CCI ni Estipendio)\n" +
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
  var part = { id:id, nombre:nombre, proyecto:CFG.PROYECTO, programa:CFG.ORG, etapa:"Inscritx",
               educacion:"", apoyoEmocional:"", inclusionLaboral:"",
               categoria:"", tarifa:"", tieneFactura:"Sí",
               dpi:dpi||db.dpi||"", nit:"", correo:"", banco:"", tipoCuenta:"", numCuenta:"", formaPago:"" };
  _escribirContenidoDP(doc, part, [], []);

  hP.appendRow([id, nombre, db.fechaNac||"", db.edad||"", db.genero||"", db.anioEntrada||"",
               CFG.PROYECTO, CFG.ORG, "Inscritx", "", "", "",
               "", 0, "Sí", dpi, "", "", "", "", "", "", url]);
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
  [ 2, "Rosaura Jeannette Saquic Lopez",     "C", "ROSA070794", "rosaura_jeannette_saquic_lopez"],
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
  [16, "Brenda Azucena del Cid Urrea",       "D", "BRDE260486", "brenda_azucena_del_cid_urrea"],
  [17, "Vilma Elizabeth Lopez Vasquez",      "C", "VILO040971", "vilma_elizabeth_lopez_vasquez"],
  [18, "Angélica Maribel Cuxe Pérez",        "D", "ANCU300380", "angelica_maribel_cuxe_perez"],
  [19, "Sara Evilia Raymundo Rivera",        "B", "SARA010779", "sara_evilia_raymundo_rivera"],
  [20, "Helen Melany Rodas López",           "C", "HERO171201", "helen_melany_rodas_lopez"],
  [21, "Elendi Nicol Pedroza Cuxé",          "D", "ELPE261105", "elendi_nicol_pedroza_cuxe"],
  [22, "Emily Cristina Zacarías Morales",    "B", "EMZA021099", "emily_cristina_zacarias_morales"],
  [23, "Juana del Rosario Vicente Choy",     "B", "JUVI281187", "juana_del_rosario_vicente_choy"],
  [24, "Mayra Lorena Cifuentes García",      "C", "MACI030373", "mayra_lorena_cifuentes_garcia"],
  [25, "Ana Rebeca Larios Perez",            "D", "ANLA060686", "ana_rebeca_larios_perez"],
  [26, "Jeimy Suceli Barrientos",            "D", "JEBA011090", "jeimy_suceli_barrientos"],
  [27, "María Aidé Alvarado Cortéz",         "D", "MAAL070490", "maria_aide_alvarado_cortez"],
  [28, "Yocelin Yajaira Celada Rodriguez",   "C", "YOCE041291", "yocelin_yajaira_celada_rodriguez"],
  [29, "Ruth Saraí Pivaral Sequen",          "C", "RUPI170992", "ruth_sarai_pivaral_sequen"],
  [30, "Laura Elizabeth Gonzalez Figueroa",  "C", "LAGO091289", "laura_elizabeth_gonzalez_figueroa"],
  [31, "Sandra Aracely Vicente Cortéz",      "C", "SADI081176", "sandra_aracely_vicente_cortez"],
  [32, "Heidy Yessenía Morales Lázaro",      "C", "HEMO120203", "heidy_yessenia_morales_lazaro"],
  [33, "Anaid Lluleydi Mateo Morales",       "C", "ANMA100605", "anaid_lluleydi_mateo_morales"]
];

// Aliases Kobo → nombre oficial (nombres cortos, typos, slugs parciales)
// Clave: nombre ya limpio (sin _, sin tildes, minúsculas) → nombre oficial exacto en PARTICIPANTES
var ALIASES_KOBO = [
  ["leticia rodriguez",       "Mirna Leticia Rodriguez Paniagua"],
  ["ericka vasquez",          "Erika Vásquez Tocay de López"],
  ["otiilia turuy paz",       "Otilia Turuy Paz"],          // typo doble i en Kobo
  ["jeanette saquic",         "Rosaura Jeannette Saquic Lopez"],
  ["mar a aide alvarado",     "María Aidé Alvarado Cortéz"], // "á" → "a" en slug
  ["mayra cifuentes",         "Mayra Lorena Cifuentes García"],
  ["ruth sara pivaral",       "Ruth Saraí Pivaral Sequen"],
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

  // Borrar datos anteriores (preservar encabezado) — deleteRows para no dejar filas vacías
  var lastRow = hP.getLastRow();
  if (lastRow > 1) hP.deleteRows(2, lastRow - 1);

  // Construir filas buscando cada participante en la DB de Creamos
  SpreadsheetApp.getActiveSpreadsheet().toast("Buscando en base de datos Creamos...", "⏳", -1);
  var filas      = [];
  var encontrados = 0, sinEncontrar = [];

  LISTA_OFICIAL.forEach(function(item) {
    var nombre = item[1], cat = item[2];
    var idLista = item[3] || "";               // ID hardcodeado en LISTA_OFICIAL
    var db      = _buscarEnCreamos_DB(nombre); // busca ID, DPI y datos extra en la DB
    var id      = db.id || idLista;            // DB tiene prioridad; si falla, usar LISTA_OFICIAL
    var dpi     = db.dpi || "";
    var tarifa = CFG.CATEGORIAS[cat] || 0;

    if (id) encontrados++;
    else    sinEncontrar.push(item[0] + ". " + nombre);

    filas.push([
      id,                 // A: Creamos_ID
      nombre,             // B: Nombre
      db.fechaNac || "",  // C: Fecha_Nacimiento
      db.edad     || "",  // D: Edad
      db.genero   || "",  // E: Genero
      db.anioEntrada || "", // F: Ano_Entrada_Creamos
      CFG.PROYECTO,       // G: Proyecto
      CFG.ORG,            // H: Programa
      "Inscritx",         // I: Etapa
      "", "", "",         // J–L: Educacion, Apoyo_Emocional, Inclusion_Laboral
      cat,                // M: Categoria
      tarifa,             // N: Tarifa_Hora
      "Sí",               // O: Tiene_Factura
      dpi,                // P: DPI
      "", "", "", "", "", "", ""  // Q–W: NIT, Correo, Banco, Tipo_Cuenta, Num_Cuenta, Forma_Pago, URL
    ]);
  });

  if (filas.length > 0) {
    hP.getRange(2, 1, filas.length, 23).setValues(filas);
    hP.getRange(2, 14, filas.length, 1).setNumberFormat("Q#,##0.00"); // col N = Tarifa_Hora
    _colorearParticipantes(hP, filas.length);
    // Marcar celdas sin Creamos ID (no se encontró en la DB)
    filas.forEach(function(f, i) {
      if (!f[0]) {
        hP.getRange(i + 2, 1)
          .setValue("⚠️ Crear perfil en Salesforce")
          .setBackground("#fff3cd").setFontColor("#856404").setFontStyle("italic");
      }
    });
  }

  // Propagar IDs y colores a todas las hojas auxiliares
  var ss2 = SpreadsheetApp.getActiveSpreadsheet();
  var mapa2 = _mapaDatosParticipantes(ss2);
  var hDE2 = ss2.getSheetByName("DiasEstudio");
  var hLT2 = ss2.getSheetByName("ListaTerapias");
  var hIL2 = ss2.getSheetByName("InclusionLaboral");
  var hHC2 = ss2.getSheetByName("HijosCCI");
  if (hDE2) { _actualizarIDsEnHoja(hDE2, 2, mapa2); _colorearHojaApoyo(hDE2, 2, 11, mapa2); }
  if (hLT2) { _actualizarIDsEnHoja(hLT2, 2, mapa2); _colorearHojaApoyo(hLT2, 2, 4, mapa2); }
  if (hIL2) { _actualizarIDsEnHoja(hIL2, 2, mapa2); _colorearHojaApoyo(hIL2, 2, 4, mapa2); }
  if (hHC2) { _actualizarIDsEnHoja(hHC2, 2, mapa2); _colorearHojaApoyo(hHC2, 2, 5, mapa2); }

  _alert(
    "✅ Lista oficial cargada — " + filas.length + " participantes.\n\n" +
    "• Encontrados en base Creamos: " + encontrados + "\n" +
    (sinEncontrar.length
      ? "• Sin perfil en Creamos (" + sinEncontrar.length + "):\n  " + sinEncontrar.join("\n  ") +
        "\n\n⚠️ Marcadas en amarillo — cuando creen el perfil en Salesforce\n" +
        "usa Administración → Sincronizar desde Creamos DB."
      : "• Todos tienen Creamos ID ✅") +
    "\n\n🎨 Colores e IDs propagados a DiasEstudio, ListaTerapias e InclusionLaboral."
  );
  try { actualizarDashboardVisual(); } catch(_) {}
}); }

/**
 * Genera ID en formato: 2 letras del primer nombre + 2 letras del primer apellido + seq 3 dígitos
 * Ejemplo: ANGELICA VELIZ → ANVE001
 */
/**
 * Busca un participante en la base de datos oficial "Copy of CREAMOS ID nuevo".
 * Estrategias (de más a menos estricta):
 *   1) Exacto  2) Normalizado (sin acentos/mayúsculas)  3) Primeras 2 palabras
 *   4) Primer nombre + primer apellido en cualquier orden
 *   5) Todas las palabras del nombre buscado aparecen en el nombre de la DB
 *   6) Coincidencia por prefijo de 4 letras de cada palabra
 * Retorna: { id, dpi, edad, genero, fechaNac, anioEntrada } o {} si no encontrado.
 */
function _buscarEnCreamos_DB(nombre) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hDB  = ss.getSheetByName(CFG.HOJAS.CREAMOS_DB);
  if (!hDB || hDB.getLastRow() < 2) return {};

  var datos = hDB.getDataRange().getValues();
  var enc   = datos[0];

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

  var normBuscar  = textoParaComparar(nombre);
  // Filtrar palabras conectoras para evitar que "del", "de", "la", "el" rompan el match
  var CONECTORAS  = { de:1, del:1, la:1, el:1, los:1, las:1, y:1, e:1 };
  var palabrasBus = normBuscar.split(/\s+/).filter(function(p){
    return p.length > 1 && !CONECTORAS[p];
  });
  var p2 = palabrasBus.slice(0, 2).join(" ");

  // Construir cache normalizado una sola vez
  var normDB = [];
  for (var i = 1; i < datos.length; i++) {
    normDB[i] = textoParaComparar(String(datos[i][iNombre] || ""));
  }

  // Estrategia 1: Exacto
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][iNombre]).trim() === nombre) return extraer(datos[i]);
  }
  // Estrategia 2: Normalizado completo
  for (var i = 1; i < datos.length; i++) {
    if (normDB[i] === normBuscar) return extraer(datos[i]);
  }
  // Estrategia 3: Primeras 2 palabras coinciden al inicio
  for (var i = 1; i < datos.length; i++) {
    var db2 = normDB[i].split(/\s+/).slice(0, 2).join(" ");
    if (normDB[i].indexOf(p2) === 0 || p2.indexOf(db2) === 0) return extraer(datos[i]);
  }
  // Estrategia 4: Primer nombre + primer apellido en cualquier posición
  if (palabrasBus.length >= 2) {
    var pNombre   = palabrasBus[0];
    var pApellido = palabrasBus[1];
    for (var i = 1; i < datos.length; i++) {
      var dbPals = normDB[i].split(/\s+/);
      var tieneNombre   = dbPals.some(function(p){ return p === pNombre; });
      var tieneApellido = dbPals.some(function(p){ return p === pApellido; });
      if (tieneNombre && tieneApellido) return extraer(datos[i]);
    }
  }
  // Estrategia 5: Todas las palabras del buscado aparecen en el nombre DB
  if (palabrasBus.length >= 3) {
    for (var i = 1; i < datos.length; i++) {
      var dbPals = normDB[i].split(/\s+/);
      var todas = palabrasBus.every(function(p) {
        return dbPals.some(function(d){ return d === p; });
      });
      if (todas) return extraer(datos[i]);
    }
  }
  // Estrategia 6: Coincidencia por prefijo de 4 letras de las primeras 2 palabras sustantivas
  if (palabrasBus.length >= 2) {
    var pref1 = palabrasBus[0].substring(0, 4);
    var pref2 = palabrasBus[1].substring(0, 4);
    var mejorPuntaje = 0, mejorFila = -1;
    for (var i = 1; i < datos.length; i++) {
      var dbPals = normDB[i].split(/\s+/);
      var tiene1 = dbPals.some(function(p){ return p.substring(0,4) === pref1; });
      var tiene2 = dbPals.some(function(p){ return p.substring(0,4) === pref2; });
      var puntaje = (tiene1 ? 1 : 0) + (tiene2 ? 1 : 0);
      if (puntaje > mejorPuntaje) { mejorPuntaje = puntaje; mejorFila = i; }
    }
    if (mejorPuntaje === 2) return extraer(datos[mejorFila]);
  }
  // Estrategia 7: Prefijo de 3 letras del primer nombre + cualquier palabra del apellido
  // Captura variaciones como Erika/Ericka, Emily/Emili, etc.
  if (palabrasBus.length >= 2) {
    var pref3 = palabrasBus[0].substring(0, 3);
    var pApell = palabrasBus[palabrasBus.length - 1]; // último apellido (más único)
    var mejorP7 = 0, mejorF7 = -1;
    for (var i = 1; i < datos.length; i++) {
      var dbP = normDB[i].split(/\s+/);
      var m3 = dbP.some(function(p){ return p.substring(0,3) === pref3; });
      var mA = dbP.some(function(p){ return p === pApell || p.substring(0,4) === pApell.substring(0,4); });
      var pts = (m3?1:0) + (mA?1:0);
      if (pts > mejorP7) { mejorP7 = pts; mejorF7 = i; }
    }
    if (mejorP7 === 2) return extraer(datos[mejorF7]);
  }
  // Estrategia 8: Palabras sustantivas — al menos primer nombre + un apellido coinciden (ignora palabras cortas y conectoras)
  if (palabrasBus.length >= 3) {
    var sustantivas = palabrasBus.filter(function(p){ return p.length >= 4; });
    if (sustantivas.length >= 2) {
      var pS1 = sustantivas[0], pS2 = sustantivas[sustantivas.length - 1];
      for (var i = 1; i < datos.length; i++) {
        var dbP = normDB[i].split(/\s+/);
        var c1 = dbP.some(function(p){ return p === pS1 || p.substring(0,4) === pS1.substring(0,4); });
        var c2 = dbP.some(function(p){ return p === pS2 || p.substring(0,4) === pS2.substring(0,4); });
        if (c1 && c2) return extraer(datos[i]);
      }
    }
  }
  // Estrategia 9: cualquier 2 palabras largas (>=4 letras) del buscado aparecen en la fila DB
  // Captura casos donde el nombre tiene palabras en distinto orden o incompleto
  if (palabrasBus.length >= 2) {
    var largas = palabrasBus.filter(function(p){ return p.length >= 4; });
    if (largas.length >= 2) {
      var mejorE9 = 0, mejorI9 = -1;
      for (var i = 1; i < datos.length; i++) {
        var dbP = normDB[i].split(/\s+/);
        var hits = largas.filter(function(p) {
          return dbP.some(function(d){ return d.indexOf(p) !== -1 || p.indexOf(d) !== -1; });
        }).length;
        if (hits > mejorE9) { mejorE9 = hits; mejorI9 = i; }
      }
      if (mejorE9 >= 2) return extraer(datos[mejorI9]);
    }
  }
  return { noEncontrado: true };
}

function _buscarEnCreamos_DB_porID(id) {
  if (!id) return { noEncontrado: true };
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var hDB = ss.getSheetByName(CFG.HOJAS.CREAMOS_DB);
  if (!hDB || hDB.getLastRow() < 2) return { noEncontrado: true };
  var datos = hDB.getDataRange().getValues();
  var enc   = datos[0];
  var iID = -1, iNombre = -1, iAnio = -1, iEdad = -1, iGenero = -1, iFechaNac = -1, iDPI = -1;
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
  if (iID === -1) return { noEncontrado: true };
  var idBusc = id.trim().toUpperCase();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][iID] || "").trim().toUpperCase() === idBusc) {
      return {
        id:          String(datos[i][iID]      || "").trim(),
        dpi:         iDPI     >= 0 ? String(datos[i][iDPI]     || "").trim() : "",
        edad:        iEdad    >= 0 ? String(datos[i][iEdad]    || "").trim() : "",
        genero:      iGenero  >= 0 ? String(datos[i][iGenero]  || "").trim() : "",
        fechaNac:    iFechaNac>= 0 ? datos[i][iFechaNac]                     : "",
        anioEntrada: iAnio    >= 0 ? String(datos[i][iAnio]   || "").trim() : ""
      };
    }
  }
  return { noEncontrado: true };
}

/**
 * Diagnóstico: para cada participante sin Creamos_ID muestra los 5 mejores
 * candidatos que encontró en la base de datos. Útil cuando la sincronización
 * reporta "sin perfil" aunque la persona sí está en la DB.
 */
function diagnosticarBusquedaCreamos() { _run(function() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var hP  = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES);
  var hDB = ss.getSheetByName(CFG.HOJAS.CREAMOS_DB);
  if (!hP || !hDB || hDB.getLastRow() < 2) {
    _alert("Faltan hojas PARTICIPANTES o Creamos DB.");
    return;
  }

  var partDatos = hP.getRange(2, 1, Math.max(hP.getLastRow()-1,1), 2).getValues();
  var dbDatos   = hDB.getDataRange().getValues();

  var enc = dbDatos[0];
  var iNombre = -1, iID = -1;
  enc.forEach(function(h, i) {
    var hl = String(h).toLowerCase().trim();
    if (hl.indexOf("nombre") !== -1) iNombre = i;
    if ((hl.indexOf("creamos") !== -1 && hl.indexOf("id") !== -1) || hl === "id") iID = i;
  });
  if (iNombre === -1) { _alert("No se encontró columna 'Nombre' en la DB."); return; }

  var dbNombres = [];
  for (var i = 1; i < dbDatos.length; i++) {
    var nm = String(dbDatos[i][iNombre] || "").trim();
    if (nm) dbNombres.push({
      original: nm,
      norm: textoParaComparar(nm),
      id: iID >= 0 ? String(dbDatos[i][iID] || "").trim() : ""
    });
  }

  var lineas = [];
  partDatos.forEach(function(f) {
    var id     = String(f[0] || "").trim();
    var nombre = String(f[1] || "").trim();
    if (!nombre) return;
    if (_esCreamos_ID_real(id)) return;

    var normBusc = textoParaComparar(nombre);
    var palabras = normBusc.split(/\s+/).filter(function(p){ return p.length >= 3; });

    var scores = dbNombres.map(function(d) {
      var dbP = d.norm.split(/\s+/);
      var hits = palabras.filter(function(p) {
        return dbP.some(function(dp){ return dp.indexOf(p) !== -1 || p.indexOf(dp) !== -1; });
      }).length;
      return { nombre: d.original, id: d.id, score: hits };
    }).sort(function(a, b){ return b.score - a.score; }).slice(0, 5);

    lineas.push("❓ " + nombre);
    scores.forEach(function(s, idx) {
      lineas.push("  " + (idx+1) + ". [" + s.score + " pts] " + s.nombre +
        (s.id ? "  ←  " + s.id : ""));
    });
    lineas.push("");
  });

  if (lineas.length === 0) {
    _alert("✅ Todos los participantes tienen Creamos ID. Nada que diagnosticar.");
    return;
  }
  _alert("🔍 Diagnóstico — mejores candidatos en DB:\n\n" + lineas.join("\n") +
    "\nSi el nombre correcto aparece en la lista, anota su Creamos ID\n" +
    "y ponlo manualmente en col A de PARTICIPANTES.");
}); }

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

/** Colorea filas de PARTICIPANTES según categoría (no sobreescribe filas Retiradx) */
function _colorearParticipantes(hP, total) {
  for (var i = 0; i < total; i++) {
    var etapa = String(hP.getRange(i + 2, 9).getValue()).trim(); // col I = Etapa
    if (etapa === "Retiradx") continue;
    var cat   = String(hP.getRange(i + 2, 13).getValue()).trim().toUpperCase(); // col M = Categoria
    var color = (CFG.COLORES_CAT[cat] || {}).bgClaro || "#ffffff";
    hP.getRange(i + 2, 1, 1, 23).setBackground(color);
  }
}

// ── Flujo de retiro ───────────────────────────────────────────

var RAZONES_RETIRO = [
  "Otras prioridades",
  "Horario laboral",
  "Retos/problemas familiares",
  "Violencia de parte de la pareja/violencia de género",
  "Migración (por motivos económicos/por violencia)",
  "Embarazo",
  "Retos/problemas de salud física",
  "Retos/problemas de salud mental",
  "Retos/Problemas legales/Privación de libertad",
  "Falta de apoyo",
  "Compromisos religiosos",
  "Problemas financieros",
  "Violencia comunitaria",
  "No querer continuar en el proceso",
  "Descontento con la organización",
  "Asesinato/Fallecimiento",
  "Cuidado de terceras personas",
  "Falta de adaptabilidad",
  "Pérdida de contacto / Inaccesibilidad",
  "Expectativas no alineadas con el programa",
  "Cambio de prioridad personal",
  "Sobrecarga personal / Dificultad para sostener el proceso"
];

function _iniciarRetiro(sheet, fila) {
  var ui = SpreadsheetApp.getUi();
  var datos = sheet.getRange(fila, 1, 1, 13).getValues()[0];
  var id = String(datos[0] || "").trim();
  var nombre = String(datos[1] || "").trim();
  var cat = String(datos[12] || "").trim(); // col M = Categoria (idx 12)

  var lista = RAZONES_RETIRO.map(function(r, i) { return (i + 1) + ". " + r; });
  var r = ui.prompt(
    "Razón de retiro — " + nombre,
    lista.join("\n") + "\n\nEscribe el número de la razón:",
    ui.ButtonSet.OK_CANCEL
  );

  if (r.getSelectedButton() !== ui.Button.OK) {
    sheet.getRange(fila, 9).setValue("Inscritx"); // revertir — col I = Etapa
    return;
  }

  var num = parseInt(r.getResponseText().trim(), 10);
  var razon = (num >= 1 && num <= RAZONES_RETIRO.length)
    ? RAZONES_RETIRO[num - 1]
    : "Otra";

  // Colorear fila rojo vivo
  sheet.getRange(fila, 1, 1, 23).setBackground("#ff1744").setFontColor("#ffffff");

  // Registrar en hoja Retiradx
  var hR = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Retiradx");
  if (hR) hR.appendRow([new Date(), id, nombre, cat, razon, ""]);
}

/**
 * Registra una participante en la hoja CiclosVida cuando su Etapa cambia a
 * "Ciclo de Vida Terminado" en PARTICIPANTES.
 */
function _registrarCicloVida(sheet, fila) {
  try {
    var datos = sheet.getRange(fila, 1, 1, 13).getValues()[0];
    var id = datos[0], nombre = datos[1], cat = datos[12]; // col M = Categoria (idx 12)
    var hCV = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CiclosVida");
    if (!hCV) return;
    hCV.appendRow([new Date(), id, nombre, cat]);
    var nr = hCV.getLastRow();
    hCV.getRange(nr, 1).setNumberFormat("dd/MM/yyyy");
    // Colorear la fila en PARTICIPANTES con el color de Ciclo de Vida
    sheet.getRange(fila, 1, 1, 23).setBackground("#b0bec5").setFontColor("#212121");
  } catch(e) { Logger.log("_registrarCicloVida error: " + e.message); }
}

function crearHojaRetiradx() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Retiradx");
  var esNueva = !hoja;
  if (esNueva) hoja = ss.insertSheet("Retiradx");

  hoja.getRange(1,1,1,6).setValues([["Fecha","Creamos_ID","Nombre","Categoria","Razón_Retiro","Notas"]])
    .setFontWeight("bold").setBackground("#c62828").setFontColor("#fff").setHorizontalAlignment("center");
  hoja.setFrozenRows(1);

  if (esNueva) {
    hoja.setColumnWidth(1,130); hoja.setColumnWidth(2,120);
    hoja.setColumnWidth(3,240); hoja.setColumnWidth(4,90);
    hoja.setColumnWidth(5,280); hoja.setColumnWidth(6,300);
  }
  hoja.activate();
  _alert("✅ Hoja 'Retiradx' creada.\n\nSe llena automáticamente cuando se marca una participante como Retiradx en PARTICIPANTES.");
}); }

/**
 * Hoja "CiclosVida" — registra participantes que completan su ciclo de vida en el programa.
 * Columnas: Fecha | Creamos_ID | Nombre | Categoria
 * Se llena automáticamente cuando Etapa = "Ciclo de Vida Terminado" en PARTICIPANTES.
 */
function crearHojaCiclosVida() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("CiclosVida");
  var esNueva = !hoja;
  if (esNueva) hoja = ss.insertSheet("CiclosVida");

  hoja.getRange(1,1,1,4).setValues([["Fecha","Creamos_ID","Nombre","Categoria"]])
    .setFontWeight("bold").setBackground("#1a237e").setFontColor("#ffffff").setHorizontalAlignment("center");
  hoja.setFrozenRows(1);

  if (esNueva) {
    hoja.setColumnWidth(1, 130);
    hoja.setColumnWidth(2, 120);
    hoja.setColumnWidth(3, 250);
    hoja.setColumnWidth(4, 90);
    hoja.getRange("A2:A500").setNumberFormat("dd/MM/yyyy");
  }

  hoja.activate();
  _alert(
    "✅ Hoja 'CiclosVida' creada.\n\n" +
    "Columnas: Fecha | Creamos_ID | Nombre | Categoria\n\n" +
    "Se llena automáticamente cuando se cambia la Etapa de una\n" +
    "participante a 'Ciclo de Vida Terminado' en PARTICIPANTES."
  );
}); }

/**
 * Hoja "Bonos" — registra bonos individuales que se suman al pago de la quincena correspondiente.
 * Columnas: Fecha | Creamos_ID | Participante | Monto | Tipo_Bono | Notas
 */
function crearHojaBonos() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Bonos");
  var esNueva = !hoja;
  if (esNueva) hoja = ss.insertSheet("Bonos");

  hoja.getRange(1,1,1,6).setValues([["Fecha","Creamos_ID","Participante","Monto","Tipo_Bono","Notas"]])
    .setFontWeight("bold").setBackground("#1565c0").setFontColor("#fff").setHorizontalAlignment("center");
  hoja.setFrozenRows(1);

  if (esNueva) {
    hoja.setColumnWidth(1, 120); hoja.setColumnWidth(2, 110);
    hoja.setColumnWidth(3, 240); hoja.setColumnWidth(4, 110);
    hoja.setColumnWidth(5, 160); hoja.setColumnWidth(6, 300);
    hoja.getRange("A2:A500").setNumberFormat("dd/MM/yyyy");
    hoja.getRange("D2:D500").setNumberFormat('"Q"#,##0.00');
  }

  hoja.activate();
  _alert(
    "💵 BONOS — Para qué sirve:\n\n" +
    "Registra bonos o incentivos individuales que se suman al pago del período.\n\n" +
    "Columnas:\n" +
    "• Fecha: fecha del bono (debe caer dentro de la quincena)\n" +
    "• Participante: nombre exacto como aparece en PARTICIPANTES\n" +
    "• Monto: Q a pagar adicionalmente\n" +
    "• Tipo_Bono: ej. 'Desempeño', 'Asistencia perfecta', 'Puntualidad', etc.\n\n" +
    "El bono aparece automáticamente en el reporte de quincena (col Bono) " +
    "y se suma al total que paga la organización."
  );
}); }

/**
 * Hoja "Estipendio" — registra estipendios fijos individuales por quincena.
 * Columnas: Fecha | Creamos_ID | Participante | Monto | Notas
 * La fecha debe caer dentro de la quincena para que se sume al pago.
 */
function crearHojaEstipendio() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Estipendio");
  var esNueva = !hoja;
  if (esNueva) hoja = ss.insertSheet("Estipendio");

  hoja.getRange(1,1,1,5).setValues([["Fecha","Creamos_ID","Participante","Monto","Notas"]])
    .setFontWeight("bold").setBackground("#4527a0").setFontColor("#fff").setHorizontalAlignment("center");
  hoja.setFrozenRows(1);

  if (esNueva) {
    hoja.setColumnWidth(1, 120); hoja.setColumnWidth(2, 110);
    hoja.setColumnWidth(3, 240); hoja.setColumnWidth(4, 110);
    hoja.setColumnWidth(5, 300);
    hoja.getRange("A2:A500").setNumberFormat("dd/MM/yyyy");
    hoja.getRange("D2:D500").setNumberFormat('"Q"#,##0.00');
  }

  hoja.activate();
  _alert(
    "🟣 ESTIPENDIO — Para qué sirve:\n\n" +
    "Registra estipendios fijos (ej: transporte, alimentación) por persona y quincena.\n\n" +
    "Columnas:\n" +
    "• Fecha: debe caer dentro de la quincena activa\n" +
    "• Creamos_ID: código de la participante (ej: ANVE241097)\n" +
    "• Participante: nombre como aparece en PARTICIPANTES\n" +
    "• Monto: Q a pagar\n" +
    "• Notas: descripción (ej: 'Transporte Q1 junio')\n\n" +
    "El estipendio aparece en col Estipendio del reporte de quincena y se suma al total."
  );
}); }

// Importar estipendios desde Kobo — solo 2026, solo período reciente, sin duplicar
function importarEstipendioDesdeKobo(silencioso) { _run(function() {
  function _msg(m) { if (!silencioso) _alert(m); else Logger.log(m); }
  var resp = UrlFetchApp.fetch(CFG.KOBO_URL_ESTIPENDIO, { muteHttpExceptions: true });
  var code = resp.getResponseCode();
  if (code === 503) { _msg("⏳ Kobo ocupado (503). Espera 2 min e intenta de nuevo."); return; }
  if (code !== 200) throw new Error("Error Kobo HTTP " + code);

  var datos = Utilities.parseCsv(resp.getContentText(), ";");
  if (datos.length < 2) { _msg("Kobo no devolvió registros de estipendio."); return; }

  var enc = datos[0].map(function(h){ return String(h).trim().toLowerCase(); });

  // Detectar columnas clave
  function _idx(palabras) {
    for (var pi = 0; pi < palabras.length; pi++) {
      for (var ei = 0; ei < enc.length; ei++) {
        if (enc[ei].indexOf(palabras[pi]) !== -1) return ei;
      }
    }
    return -1;
  }
  var iStart  = _idx(["start"]);
  var iPart   = _idx(["participante","nombre","participant"]);
  var iMonto  = _idx(["monto","amount","valor","total","estipendio","q_"]);
  var iID     = _idx(["c_id","creamos","cid","codigo"]);
  var iNotas  = _idx(["nota","note","descripcion","motivo","concepto"]);

  if (iStart < 0 || iPart < 0 || iMonto < 0) {
    _msg("❌ No se detectaron columnas en el CSV de estipendio.\n" +
      "Cabeceras encontradas:\n" + datos[0].join(", "));
    return;
  }

  // Filtrar solo registros recientes (desde 60 días antes del período activo)
  var periodo = _periodoActivo();
  var fechaMinima;
  if (periodo && periodo.fi) {
    fechaMinima = new Date(periodo.fi);
    fechaMinima.setDate(fechaMinima.getDate() - 60);
  } else {
    fechaMinima = new Date(); fechaMinima.setDate(fechaMinima.getDate() - 90);
  }

  var mapeoNombres = cargarMapeoNombres();
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Estipendio");
  if (!hoja) { crearHojaEstipendio(); hoja = ss.getSheetByName("Estipendio"); }

  // Leer UUIDs ya en la hoja para evitar duplicados
  var iUUID = _idx(["_uuid","uuid"]);
  var uuidsExist = {};
  if (iUUID >= 0 && hoja.getLastRow() >= 2) {
    // Estipendio sheet no guarda UUID, usamos Fecha+Nombre como clave
  }
  // Usar Fecha+Participante como clave de dedup
  var clavesExist = {};
  if (hoja.getLastRow() >= 2) {
    var exRows = hoja.getRange(2, 1, hoja.getLastRow()-1, 3).getValues();
    exRows.forEach(function(r) {
      var k = String(r[0]||"").substring(0,10) + "|" + String(r[2]||"").trim().toLowerCase();
      clavesExist[k] = true;
    });
  }

  var nuevas = [];
  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var rawStart = String(fila[iStart] || "").trim();
    var tsRow = new Date(rawStart);
    if (isNaN(tsRow)) {
      var pm = rawStart.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
      if (pm) tsRow = new Date(pm[3], pm[2]-1, pm[1]);
    }
    if (isNaN(tsRow) || tsRow < fechaMinima) continue;

    var anio = tsRow.getFullYear();
    if (anio < 2026) continue;

    var rawNombre = String(fila[iPart] || "").trim();
    if (!rawNombre) continue;
    var nombre = normalizarNombre(rawNombre, mapeoNombres) || limpiarNombre(rawNombre) || rawNombre;
    var monto  = parseFloat(String(fila[iMonto]||"").replace(",",".")) || 0;
    if (!monto) continue;

    var fechaFmt = Utilities.formatDate(tsRow, CFG.TIMEZONE, "yyyy-MM-dd");
    var clave = fechaFmt + "|" + nombre.trim().toLowerCase();
    if (clavesExist[clave]) continue;

    var id    = iID >= 0 ? String(fila[iID]||"").trim() : "";
    var notas = iNotas >= 0 ? String(fila[iNotas]||"").trim() : "";
    nuevas.push([tsRow, id, nombre, monto, notas]);
    clavesExist[clave] = true;
  }

  if (nuevas.length === 0) { _msg("✅ Sin estipendios nuevos para importar."); return; }
  hoja.getRange(hoja.getLastRow()+1, 1, nuevas.length, 5).setValues(nuevas);
  hoja.getRange(hoja.getLastRow()-nuevas.length+2, 1, nuevas.length, 1).setNumberFormat("dd/MM/yyyy");
  hoja.getRange(hoja.getLastRow()-nuevas.length+2, 4, nuevas.length, 1).setNumberFormat('"Q"#,##0.00');
  _msg("✅ " + nuevas.length + " estipendios importados desde Kobo.");
}); }

// ── Hojas de pago: Cheques y Transferencias ──────────────────────

/**
 * Hoja "Cheques" — una fila por participante por mes, con Q1 y Q2.
 * Columnas: Nombre | Programa | Mes | Año | Q1_Monto | Q2_Monto | Total | #Cheque_Q1 | #Cheque_Q2 | Cta_Cheques | Status
 */
function crearHojaCheques() { _run(function() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Cheques");
  var esNueva = !hoja;
  if (esNueva) hoja = ss.insertSheet("Cheques");

  var enc = ["Nombre","Programa","Mes","Año","Q1_Monto","Q2_Monto","Total",
             "#Cheque_Q1","#Cheque_Q2","Cta_Cheques","Status"];
  hoja.getRange(1, 1, 1, enc.length).setValues([enc])
    .setBackground("#1a237e").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center");
  hoja.setFrozenRows(1);

  var vStatus = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pendiente","Cobrado","Cancelado"], true).build();
  hoja.getRange("K2:K1000").setDataValidation(vStatus);

  hoja.getRange("E2:G1000").setNumberFormat('"Q"#,##0.00'); // Q1, Q2, Total

  if (esNueva) {
    hoja.setColumnWidth(1, 250);  // Nombre
    hoja.setColumnWidth(2, 140);  // Programa
    hoja.setColumnWidth(3, 110);  // Mes
    hoja.setColumnWidth(4, 70);   // Año
    hoja.setColumnWidth(5, 110);  // Q1_Monto
    hoja.setColumnWidth(6, 110);  // Q2_Monto
    hoja.setColumnWidth(7, 110);  // Total
    hoja.setColumnWidth(8, 120);  // #Cheque_Q1
    hoja.setColumnWidth(9, 120);  // #Cheque_Q2
    hoja.setColumnWidth(10, 160); // Cta_Cheques
    hoja.setColumnWidth(11, 130); // Status
  }

  hoja.activate();
  _alert(
    "🏦 CHEQUES — Hoja lista.\n\n" +
    "Una fila por participante por mes. Columnas:\n" +
    "• Nombre / Programa / Mes / Año\n" +
    "• Q1_Monto / Q2_Monto: montos de cada quincena (auto)\n" +
    "• Total: suma Q1+Q2 (auto)\n" +
    "• #Cheque_Q1 / #Cheque_Q2: número de cheque (llenar manualmente)\n" +
    "• Cta_Cheques: cuenta de origen\n" +
    "• Status: Pendiente → Cobrado\n\n" +
    "Cuando TODOS los cheques estén en 'Cobrado', la hoja se archiva automáticamente."
  );
}); }

/**
 * Hoja "Transferencias" — registra transferencias electrónicas por mes.
 * Columnas: Nombre | Tipo_Pago | Servicio | Banco | Tipo_Cuenta | Num_Cta |
 *           Cta_Pago | Quincena_1 | Quincena_2 | Total_Mes | Mes | Año
 */
function crearHojaTransferencias() { _run(function() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Transferencias");
  var esNueva = !hoja;
  if (esNueva) hoja = ss.insertSheet("Transferencias");

  var enc = ["Nombre","Tipo_Pago","Servicio","Banco","Tipo_Cuenta","Num_Cta",
             "Cta_Pago","Quincena_1","Quincena_2","Total_Mes","Mes","Año","Status"];
  hoja.getRange(1, 1, 1, enc.length).setValues([enc])
    .setBackground("#1b5e20").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center");
  hoja.setFrozenRows(1);

  hoja.getRange("H2:J1000").setNumberFormat('"Q"#,##0.00'); // Q1, Q2, Total_Mes

  if (esNueva) {
    hoja.setColumnWidth(1, 250);  // Nombre
    hoja.setColumnWidth(2, 120);  // Tipo_Pago
    hoja.setColumnWidth(3, 160);  // Servicio
    hoja.setColumnWidth(4, 140);  // Banco
    hoja.setColumnWidth(5, 110);  // Tipo_Cuenta
    hoja.setColumnWidth(6, 160);  // Num_Cta
    hoja.setColumnWidth(7, 160);  // Cta_Pago
    hoja.setColumnWidth(8, 110);  // Quincena_1
    hoja.setColumnWidth(9, 110);  // Quincena_2
    hoja.setColumnWidth(10, 110); // Total_Mes
    hoja.setColumnWidth(11, 110); // Mes
    hoja.setColumnWidth(12, 80);   // Año
    hoja.setColumnWidth(13, 160);  // Status
  }

  var vStatus = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pendiente","Transferencias Subidas"], true).build();
  hoja.getRange("M2:M1000").setDataValidation(vStatus);

  hoja.activate();
  _alert(
    "🔄 TRANSFERENCIAS — Hoja lista.\n\n" +
    "Columnas:\n" +
    "• Nombre: participante\n" +
    "• Tipo_Pago / Servicio: Transferencia / Textil\n" +
    "• Banco / Tipo_Cuenta / Num_Cta: datos bancarios del participante\n" +
    "• Cta_Pago: cuenta de la organización origen del pago\n" +
    "• Quincena_1 / Quincena_2: montos por quincena (auto)\n" +
    "• Total_Mes: suma automática de Q1+Q2\n" +
    "• Mes / Año: período mensual\n" +
    "• Status: Pendiente → Transferencias Subidas (archiva y cierra)\n\n" +
    "Cuando cambias Status a 'Transferencias Subidas', la hoja se archiva automáticamente."
  );
}); }

/**
 * Registra los pagos de una quincena seleccionada en las hojas Cheques y Transferencias.
 * Lee el reporte de quincena ya calculado (hoja con tab name de PERIODOS col G),
 * y para cada participante con monto > 0 lo enruta según Forma_Pago en PARTICIPANTES.
 */
function registrarPagosQuincena() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = CFG.TIMEZONE;

  // ── Paso 1: Obtener lista de quincenas desde PERIODOS ─────────
  var hPeriodos = ss.getSheetByName(CFG.HOJAS.PERIODOS);
  var periodos  = [];
  if (hPeriodos && hPeriodos.getLastRow() >= 2) {
    var datPer = hPeriodos.getRange(2, 1, hPeriodos.getLastRow() - 1, 7).getValues();
    datPer.forEach(function(r, i) {
      var label  = String(r[1] || "").trim();
      var estado = String(r[4] || "").trim();
      var tab    = String(r[6] || "").trim();
      if (!label || !tab) return;
      periodos.push({ fila: i + 2, label: label, estado: estado, tab: tab });
    });
  }

  // ── Paso 2: Mostrar lista y pedir selección ───────────────────
  var promptMsg;
  if (periodos.length === 0) {
    promptMsg = "No hay quincenas en PERIODOS.\nEscribe el nombre de la hoja del reporte manualmente:";
  } else {
    var lista = periodos.map(function(p, i) {
      return (i + 1) + ". " + p.label + " [" + p.estado + "]";
    }).join("\n");
    promptMsg = "Quincenas disponibles:\n\n" + lista +
      "\n\nEscribe el número de la quincena a registrar:";
  }

  var resp = ui.prompt("💳 Registrar pagos de quincena", promptMsg, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var entrada = resp.getResponseText().trim();
  if (!entrada) { _alert("No ingresaste ningún valor."); return; }

  var periodoSel, tabNombre;
  if (periodos.length > 0) {
    var idx = parseInt(entrada);
    if (!isNaN(idx) && idx >= 1 && idx <= periodos.length) {
      periodoSel = periodos[idx - 1];
      tabNombre  = periodoSel.tab;
    } else {
      // Intento de coincidencia por texto
      var entradaNorm = entrada.toLowerCase();
      for (var pi = 0; pi < periodos.length; pi++) {
        if (periodos[pi].label.toLowerCase().indexOf(entradaNorm) !== -1 ||
            periodos[pi].tab.toLowerCase().indexOf(entradaNorm) !== -1) {
          periodoSel = periodos[pi];
          tabNombre  = periodoSel.tab;
          break;
        }
      }
      if (!tabNombre) { tabNombre = entrada; } // Usar como nombre de hoja directo
    }
  } else {
    tabNombre = entrada;
  }

  // ── Paso 3: Leer reporte de quincena ─────────────────────────
  var hReporte = ss.getSheetByName(tabNombre);
  if (!hReporte) {
    _alert("No se encontró la hoja de reporte: '" + tabNombre + "'\n\n" +
           "Verifica que la quincena haya sido calculada (📅 Quincena → Ver / actualizar quincena).");
    return;
  }

  // El reporte tiene: fila 1=tarifas, 2=título, 3=encabezados, 4+=datos
  // Col B (idx 1)=Participante, Col D (idx 3)=ID, Col O (idx 14)=Neto part.
  var repData = hReporte.getDataRange().getValues();
  var pagosReporte = {}; // {nombre: {neto, id}}
  for (var ri = 3; ri < repData.length; ri++) {
    var nombre = String(repData[ri][1] || "").trim();
    var idRep  = String(repData[ri][3] || "").trim();
    var neto   = parseFloat(repData[ri][14]) || 0;
    if (!nombre || neto <= 0) continue;
    // Ignorar filas de totales/leyenda (sin número en col A)
    var numFila = repData[ri][0];
    if (!numFila || isNaN(Number(numFila))) continue;
    pagosReporte[nombre] = { neto: neto, id: idRep };
  }

  if (Object.keys(pagosReporte).length === 0) {
    _alert("No se encontraron participantes con pago > 0 en la hoja '" + tabNombre + "'.\n" +
           "Verifica que el reporte esté calculado correctamente.");
    return;
  }

  // ── Paso 4: Leer datos de pago de PARTICIPANTES ───────────────
  // Cols: A(1)=ID, B(2)=Nombre, D(4)=Programa, F(6)=Educacion, G(7)=Apoyo_Emocional,
  //       H(8)=Inclusion_Laboral, I(9)=Categoria
  //       O(15)=Banco, P(16)=Tipo_Cuenta, Q(17)=Num_Cuenta, R(18)=Forma_Pago
  var hP = _sh(CFG.HOJAS.PARTICIPANTES);
  if (hP.getLastRow() < 2) { _alert("PARTICIPANTES está vacía."); return; }
  var datosP = hP.getRange(2, 1, hP.getLastRow() - 1, 22).getValues();
  // mapaPago indexado por: ID, nombre exacto y nombre normalizado
  var mapaPagoById   = {}; // {cremos_id: info}
  var mapaPagoByNorm = {}; // {nombre_normalizado: info}
  datosP.forEach(function(r) {
    var pid    = String(r[0] || "").trim();       // col A = Creamos_ID
    var nombre = String(r[1] || "").trim();
    if (!nombre) return;
    var programa   = String(r[7]  || "").trim();  // col H = Programa (idx 7)
    var educacion  = String(r[9]  || "").trim();  // col J = Educacion (idx 9)
    var apoyo      = String(r[10] || "").trim();  // col K = Apoyo_Emocional (idx 10)
    var inclusion  = String(r[11] || "").trim();  // col L = Inclusion_Laboral (idx 11)
    var banco      = String(r[18] || "").trim();  // col S = Banco (idx 18)
    var tipoCuenta = String(r[19] || "").trim();  // col T = Tipo_Cuenta (idx 19)
    var numCuenta  = String(r[20] || "").trim();  // col U = Num_Cuenta (idx 20)
    var formaPago  = String(r[21] || "").trim();  // col V = Forma_Pago (idx 21)

    var servicio = "Textil"; // siempre Textil para el taller

    var info = { formaPago: formaPago, banco: banco, tipoCuenta: tipoCuenta,
                 numCuenta: numCuenta, programa: programa, servicio: servicio,
                 nombreOficial: nombre };
    if (pid) mapaPagoById[pid] = info;
    mapaPagoByNorm[textoParaComparar(nombre)] = info;
  });

  // Helper: busca info de pago por ID primero, luego nombre normalizado
  function _buscarInfoPago(nombre, id) {
    if (id && mapaPagoById[id]) return mapaPagoById[id];
    var norm = textoParaComparar(nombre);
    if (mapaPagoByNorm[norm]) return mapaPagoByNorm[norm];
    // Búsqueda parcial: primera palabra significativa del nombre
    var palabras = norm.split(/\s+/).filter(function(p){ return p.length >= 4; });
    if (palabras.length >= 2) {
      var keys = Object.keys(mapaPagoByNorm);
      for (var ki = 0; ki < keys.length; ki++) {
        var kn = keys[ki];
        var matches = palabras.filter(function(p){ return kn.indexOf(p) !== -1; }).length;
        if (matches >= 2) return mapaPagoByNorm[kn];
      }
    }
    return null;
  }

  // ── Paso 5: Determinar mes, año y quincena_label ──────────────
  var labelPeriodo = periodoSel ? periodoSel.label : tabNombre;
  // Inferir mes/año desde la fecha_inicio del período
  var mesNombre = "—", anioNum = new Date().getFullYear(), quincenaLabel = "Q1";
  if (periodoSel) {
    var datPerFull = hPeriodos.getRange(2, 1, hPeriodos.getLastRow() - 1, 4).getValues();
    var filaIdx = periodoSel.fila - 2;
    if (filaIdx >= 0 && filaIdx < datPerFull.length) {
      var fecIni = new Date(datPerFull[filaIdx][2]); // col C = Fecha_Inicio
      var fecFin = new Date(datPerFull[filaIdx][3]); // col D = Fecha_Fin
      if (!isNaN(fecIni)) {
        var mesIdx = fecIni.getMonth();
        mesNombre = CFG.MESES[mesIdx];
        anioNum   = fecIni.getFullYear();
        // fecIni >= 20 (normalmente 26) → PRIMERA quincena Q1 (del 26 al 10)
        // fecIni < 20  (normalmente 11) → SEGUNDA quincena Q2 (del 11 al 25)
        quincenaLabel = fecIni.getDate() >= 20 ? "Q1" : "Q2";
      }
    }
  }

  // ── Paso 6: Crear hojas si no existen ────────────────────────
  if (!ss.getSheetByName("Cheques"))       crearHojaCheques();
  if (!ss.getSheetByName("Transferencias")) crearHojaTransferencias();
  var hCheques = ss.getSheetByName("Cheques");
  var hTransf  = ss.getSheetByName("Transferencias");

  // Leer datos existentes para deduplicación
  var existCheques = [];
  if (hCheques.getLastRow() >= 2) {
    existCheques = hCheques.getRange(2, 1, hCheques.getLastRow() - 1, 9).getValues();
  }
  var existTransf = [];
  if (hTransf.getLastRow() >= 2) {
    existTransf = hTransf.getRange(2, 1, hTransf.getLastRow() - 1, 12).getValues();
  }

  // ── Paso 7: Enrutar pagos ─────────────────────────────────────
  var hoy = new Date();
  var cuentaPago = "Cuenta mi eelo";
  var nCheques = 0, nTransferencias = 0;
  var erroresPago = [];

  Object.keys(pagosReporte).forEach(function(nombre) {
    var entrada = pagosReporte[nombre];
    var monto   = entrada.neto;
    var idRep   = entrada.id;
    var info    = _buscarInfoPago(nombre, idRep);
    if (!info) {
      erroresPago.push("⚠️ " + nombre + ": no encontrado en PARTICIPANTES (sin Forma_Pago)");
      return;
    }
    var forma = info.formaPago;
    var nombreOficial = info.nombreOficial || nombre;

    if (forma === "Cheque") {
      // Una fila por participante por mes — misma lógica que Transferencias
      // Cols: Nombre(0) Programa(1) Mes(2) Año(3) Q1_Monto(4) Q2_Monto(5) Total(6) #Chq_Q1(7) #Chq_Q2(8) Cta(9) Status(10)
      var normNomChq = textoParaComparar(nombreOficial);
      var filaChqExist = -1;
      for (var ci = 0; ci < existCheques.length; ci++) {
        if (textoParaComparar(String(existCheques[ci][0]||"")) === normNomChq &&
            String(existCheques[ci][2]).trim() === mesNombre &&
            String(existCheques[ci][3]).toString().trim() === String(anioNum)) {
          filaChqExist = ci;
          break;
        }
      }

      if (filaChqExist >= 0) {
        var filaRealChq = filaChqExist + 2;
        if (quincenaLabel === "Q1") {
          hCheques.getRange(filaRealChq, 5).setValue(monto);
          existCheques[filaChqExist][4] = monto;
        } else {
          hCheques.getRange(filaRealChq, 6).setValue(monto);
          existCheques[filaChqExist][5] = monto;
        }
        var q1c = parseFloat(existCheques[filaChqExist][4]) || 0;
        var q2c = parseFloat(existCheques[filaChqExist][5]) || 0;
        hCheques.getRange(filaRealChq, 7).setValue(q1c + q2c);
        existCheques[filaChqExist][6] = q1c + q2c;
        nCheques++;
      } else {
        var q1cNew = quincenaLabel === "Q1" ? monto : 0;
        var q2cNew = quincenaLabel === "Q2" ? monto : 0;
        var nuevaFilaChq = [nombreOficial, info.programa, mesNombre, anioNum,
                            q1cNew, q2cNew, q1cNew + q2cNew, "", "", "", "Pendiente"];
        hCheques.appendRow(nuevaFilaChq);
        var newRowChq = hCheques.getLastRow();
        hCheques.getRange(newRowChq, 5, 1, 3).setNumberFormat('"Q"#,##0.00');
        existCheques.push(nuevaFilaChq);
        nCheques++;
      }

    } else if (forma === "Transferencia") {
      // Buscar fila existente en Transferencias para (nombre, mes, año)
      var normNomTr = textoParaComparar(nombreOficial);
      var filaExist = -1;
      for (var ti = 0; ti < existTransf.length; ti++) {
        if (textoParaComparar(String(existTransf[ti][0]||"")) === normNomTr &&
            String(existTransf[ti][10]).trim() === mesNombre &&
            String(existTransf[ti][11]).toString().trim() === String(anioNum)) {
          filaExist = ti;
          break;
        }
      }

      if (filaExist >= 0) {
        // Actualizar fila existente: llenar Q1 o Q2
        var filaReal = filaExist + 2; // 1-indexed + encabezado
        if (quincenaLabel === "Q1") {
          hTransf.getRange(filaReal, 8).setValue(monto);
          existTransf[filaExist][7] = monto;
        } else {
          hTransf.getRange(filaReal, 9).setValue(monto);
          existTransf[filaExist][8] = monto;
        }
        // Recalcular Total_Mes
        var q1 = parseFloat(existTransf[filaExist][7]) || 0;
        var q2 = parseFloat(existTransf[filaExist][8]) || 0;
        hTransf.getRange(filaReal, 10).setValue(q1 + q2);
        existTransf[filaExist][9] = q1 + q2;
        nTransferencias++;

      } else {
        // Insertar nueva fila
        var q1new = quincenaLabel === "Q1" ? monto : 0;
        var q2new = quincenaLabel === "Q2" ? monto : 0;
        var nuevaFila = [nombreOficial, "Transferencia", info.servicio, info.banco,
                         info.tipoCuenta, info.numCuenta, cuentaPago,
                         q1new, q2new, q1new + q2new, mesNombre, anioNum];
        hTransf.appendRow(nuevaFila);
        var newRowT = hTransf.getLastRow();
        hTransf.getRange(newRowT, 8, 1, 3).setNumberFormat('"Q"#,##0.00');
        existTransf.push(nuevaFila);
        nTransferencias++;
      }

    } else {
      erroresPago.push("⚠️ " + nombre + ": Forma_Pago '" + forma + "' no reconocida (no es Cheque ni Transferencia)");
    }
  });

  // ── Paso 8: Resumen ───────────────────────────────────────────
  var msg = "✅ Pagos de quincena registrados\n\n" +
    "Período: " + labelPeriodo + "\n" +
    "Quincena: " + quincenaLabel + " — " + mesNombre + " " + anioNum + "\n\n" +
    "🏦 Cheques registrados/actualizados: " + nCheques + "\n" +
    "🔄 Transferencias registradas/actualizadas: " + nTransferencias;
  if (erroresPago.length > 0) {
    msg += "\n\n⚠️ Avisos:\n" + erroresPago.join("\n");
  }

  // ── Paso 9: Si es Q2 → exportar PDF + eliminar hoja + marcar Pagado ──
  if (quincenaLabel === "Q2") {
    msg += "\n\n📄 Es Q2 — generando PDF y cerrando período...";
    _alert(msg);
    try {
      _cerrarPeriodoQ2(ss, periodoSel, hPeriodos, tabNombre, mesNombre, anioNum, tz);
    } catch(e) {
      _alert("⚠️ Pagos guardados. Error al cerrar período:\n" + e.message +
             "\nEl PDF/cierre puede hacerse manualmente.");
    }
  } else {
    msg += "\n\nℹ️ Es Q1 — cuando registres Q2, el período se cerrará automáticamente.";
    _alert(msg);
  }
}); }

/**
 * Cierra el período Q2: guarda PDF del reporte en Drive, elimina hoja del reporte,
 * busca el reporte Q1 y lo elimina también si existe, marca período como Pagado.
 */
function _cerrarPeriodoQ2(ss, periodoSel, hPeriodos, tabQ2, mesNombre, anioNum, tz) {
  var ssId = ss.getId();
  var token = ScriptApp.getOAuthToken();

  // Buscar hoja Q1 del mismo mes (nombre contiene el mes/año pero fecha inicio ≤ 15)
  var hojaQ2 = ss.getSheetByName(tabQ2);
  var hojaQ1 = null;
  if (hPeriodos && hPeriodos.getLastRow() >= 2) {
    var allPer = hPeriodos.getRange(2, 1, hPeriodos.getLastRow() - 1, 7).getValues();
    allPer.forEach(function(r) {
      var tab   = String(r[6] || "").trim();
      var fIni  = new Date(r[2]);
      if (!tab || tab === tabQ2) return;
      if (!isNaN(fIni) && fIni.getMonth() === new Date().getMonth()) {
        // Mismo mes → es el Q1 compañero
        var h = ss.getSheetByName(tab);
        if (h) hojaQ1 = h;
      }
    });
  }

  // Exportar PDFs a Drive
  var folder;
  try {
    var it = DriveApp.getFoldersByName(CFG.DRIVE_FOLDER_PAGOS);
    folder = it.hasNext() ? it.next() : DriveApp.createFolder(CFG.DRIVE_FOLDER_PAGOS);
  } catch(e) { folder = DriveApp.getRootFolder(); }

  function _exportPDF(hoja, nombre) {
    if (!hoja) return;
    var url = "https://docs.google.com/spreadsheets/d/" + ssId +
      "/export?format=pdf&gid=" + hoja.getSheetId() +
      "&portrait=false&size=A4&fitw=true&top_margin=0.5&bottom_margin=0.5" +
      "&left_margin=0.5&right_margin=0.5&sheetnames=false&printtitle=false";
    try {
      var resp = UrlFetchApp.fetch(url, {
        headers: { Authorization: "Bearer " + token },
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) {
        folder.createFile(resp.getBlob().setName(nombre + ".pdf"));
      }
    } catch(e) { /* drive export falla silenciosamente */ }
  }

  var prefijoPDF = mesNombre + "_" + anioNum;
  if (hojaQ1) _exportPDF(hojaQ1, "Q1_" + prefijoPDF);
  _exportPDF(hojaQ2, (hojaQ1 ? "Q2_" : "Q2_solo_") + prefijoPDF);

  // Eliminar hojas de reporte
  if (hojaQ1) { try { ss.deleteSheet(hojaQ1); } catch(e) {} }
  if (hojaQ2) { try { ss.deleteSheet(hojaQ2); } catch(e) {} }

  // Marcar período(s) como Pagado en PERIODOS
  if (hPeriodos && hPeriodos.getLastRow() >= 2) {
    var datPer2 = hPeriodos.getRange(2, 1, hPeriodos.getLastRow() - 1, 7).getValues();
    datPer2.forEach(function(r, i) {
      var tab = String(r[6] || "").trim();
      if (tab === tabQ2 || (hojaQ1 && tab === hojaQ1.getName())) {
        hPeriodos.getRange(i + 2, 5).setValue("Pagado");
      }
    });
  }

  SpreadsheetApp.getUi().alert(
    "✅ PERÍODO CERRADO\n\n" +
    "📄 PDF guardado en Drive → " + CFG.DRIVE_FOLDER_PAGOS + "\n" +
    (hojaQ1 ? "✓ Reporte Q1 archivado y eliminado\n" : "⚠️ Reporte Q1 no encontrado\n") +
    "✓ Reporte Q2 archivado y eliminado\n" +
    "✓ Período marcado como 'Pagado' en PERIODOS"
  );
}

/**
 * Hoja "HijosCCI" — registra si cada participante tiene hijos en CCI y cuántos.
 * Se sincroniza automáticamente con cols U/V de PARTICIPANTES.
 */
function crearHojaHijosCCI() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("HijosCCI");
  var esNueva = !hoja;
  if (esNueva) hoja = ss.insertSheet("HijosCCI");

  hoja.getRange(1,1,1,5).setValues([["Creamos_ID","Participante","¿Tiene hijos en CCI? (X)","¿Cuántos?","Notas"]])
    .setFontWeight("bold").setBackground("#6a1b9a").setFontColor("#fff").setHorizontalAlignment("center");
  hoja.setFrozenRows(1);

  var mapa = _mapaDatosParticipantes(ss);

  if (esNueva) {
    var filas = LISTA_OFICIAL.map(function(it) {
      var id  = (mapa[it[1]] || {}).id || it[3] || "";
      // Leer valor actual desde PARTICIPANTES si existe
      return [id, it[1], "", "", ""];
    });
    hoja.getRange(2, 1, filas.length, 5).setValues(filas);
    var vX = SpreadsheetApp.newDataValidation().requireValueInList(["X",""],true).build();
    hoja.getRange(2, 3, filas.length, 1).setDataValidation(vX).setHorizontalAlignment("center");
    hoja.getRange(2, 4, filas.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).build()
    ).setHorizontalAlignment("center");
  } else {
    _actualizarIDsEnHoja(hoja, 2, mapa);
  }

  _colorearHojaApoyo(hoja, 2, 5, mapa);

  hoja.setColumnWidth(1, 110); hoja.setColumnWidth(2, 260);
  hoja.setColumnWidth(3, 160); hoja.setColumnWidth(4, 110);
  hoja.setColumnWidth(5, 300);
  hoja.activate();
  _alert(
    "👶 HIJOS CCI — Para qué sirve:\n\n" +
    "Registra si la participante tiene hijos en el CCI (Seguro Social) y cuántos.\n\n" +
    "• Marca X en col C si tiene hijos en CCI\n" +
    "• Escribe el número en col D\n\n" +
    "Se sincroniza automáticamente con las columnas I (Hijos_CCI) y J (Num_Hijos_CCI) " +
    "de PARTICIPANTES cada vez que editas una celda."
  );
}); }

function crearHojaInclusionLaboral() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("InclusionLaboral");
  var esNueva = !hoja;
  if (esNueva) hoja = ss.insertSheet("InclusionLaboral");

  hoja.getRange(1,1,1,4).setValues([["Creamos_ID","Participante","Participa (X)","Notas"]])
    .setFontWeight("bold").setBackground("#f57c00").setFontColor("#fff").setHorizontalAlignment("center");
  hoja.setFrozenRows(1);

  var mapa = _mapaDatosParticipantes(ss);

  if (esNueva) {
    var filas = LISTA_OFICIAL.map(function(it) {
      var id = (mapa[it[1]] || {}).id || it[3] || "";
      return [id, it[1], "", ""];
    });
    hoja.getRange(2, 1, filas.length, 4).setValues(filas);
    var vX = SpreadsheetApp.newDataValidation().requireValueInList(["X",""],true).build();
    hoja.getRange(2, 3, filas.length, 1).setDataValidation(vX).setHorizontalAlignment("center");
  } else {
    _actualizarIDsEnHoja(hoja, 2, mapa);
  }

  // Colores por categoría — siempre desde PARTICIPANTES
  _colorearHojaApoyo(hoja, 2, 4, mapa);

  hoja.setColumnWidth(1,120); hoja.setColumnWidth(2,260);
  hoja.setColumnWidth(3,120); hoja.setColumnWidth(4,300);
  hoja.activate();
  _alert(
    "💼 INCLUSIÓN LABORAL — Para qué sirve:\n\n" +
    "Marca con X a las participantes que están activas en el programa de Inclusión Laboral.\n" +
    "Se sincroniza automáticamente con col H (Inclusion_Laboral) de PARTICIPANTES.\n\n" +
    "No cuenta como días/horas — solo indica participación en el programa."
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

  var datos = hP.getRange(2, 1, lastRow - 1, 14).getValues();

  // Agrupar por categoría
  var porCat = { A: [], B: [], C: [], D: [] };
  datos.forEach(function(r) {
    var id      = String(r[0]).trim();
    var nombre  = String(r[1]).trim();
    var cat     = String(r[12]).trim().toUpperCase(); // col M = Categoria (idx 12)
    var tarifa  = parseFloat(r[13]) || CFG.CATEGORIAS[cat] || 0; // col N = Tarifa (idx 13)
    var estado  = String(r[8]).trim() || "Inscritx";  // col I = Etapa (idx 8)
    if (!nombre || !porCat[cat]) return;
    porCat[cat].push({ id: id, nombre: nombre, tarifa: tarifa, etapa: estado });
  });

  // Recrear hoja Directorio
  var HOJA = "Directorio";
  var hD = ss.getSheetByName(HOJA);
  if (hD) ss.deleteSheet(hD);
  hD = ss.insertSheet(HOJA);

  var filas = [], tipos = [];

  function push(fila, tipo) { filas.push(fila); tipos.push(tipo); }

  var ahora = Utilities.formatDate(new Date(), CFG.TIMEZONE, "dd/MM/yyyy HH:mm");
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
      push([num++, p.id, p.nombre, cat, "Q" + tarifa.toFixed(2), p.etapa],
           p.etapa === "Retiradx" ? "inactivo" : "activo");
    });

    push(["", "", "Subtotal categoría " + cat + ": " + lista.length, "", "", ""], "subtotal");
    push(["", "", "", "", "", ""], "vacio");
  });

  // Total general
  var totalActivos = Object.keys(porCat).reduce(function(s, c) {
    return s + porCat[c].filter(function(p){ return p.etapa !== "Retiradx"; }).length;
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
    proyecto:         String(f[6]  || ""), programa:         String(f[7]  || ""),
    etapa:            String(f[8]  || ""), educacion:        String(f[9]  || ""),
    apoyoEmocional:   String(f[10] || ""), inclusionLaboral: String(f[11] || ""),
    categoria:        String(f[12] || ""), tarifa:           String(f[13] || ""),
    tieneFactura:     String(f[14] || ""),
    dpi:              String(f[15] || ""), nit:              String(f[16] || ""),
    correo:           String(f[17] || ""), banco:            String(f[18] || ""),
    tipoCuenta:       String(f[19] || ""), numCuenta:        String(f[20] || ""),
    formaPago:        String(f[21] || "")
  };
  var urlActual = String(f[22] || "");

  var carpeta = _carpetaDP();
  var doc = _abrirOCrearDocProceso(id, nombre, carpeta, urlActual);
  _escribirContenidoDP(doc, part);

  var urlNueva = doc.getUrl();
  if (urlNueva !== urlActual) hP.getRange(fila, 23).setValue(urlNueva); // col W = URL_Doc_Proceso
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
    // Documento inaccesible/borrado → crear uno nuevo
    if (!doc) {
      doc = DocumentApp.create(titulo);
      DriveApp.getFileById(doc.getId()).moveTo(carpeta);
    } else {
      doc.setName(titulo);
    }
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
  var tz  = CFG.TIMEZONE;
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
    ["Creamos ID",     part.id          ||"—", "Etapa",         part.etapa       ||"—"],
    ["Nombre",         part.nombre      ||"—", "Categoría",     part.categoria   ||"—"],
    ["Proyecto",       part.proyecto    ||"—", "Tiene Factura", part.tieneFactura||"—"],
    ["Programa",       part.programa    ||"—", "Tarifa/hora",   tarifahora],
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
 * ⚡ IMPORTAR TODO (1 clic)
 * Hace en secuencia: importar Kobo → normalizar nombres → actualizar quincena
 */
function importarYEmparejar() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [], errores = [];

  ss.toast("Paso 1/2: Importando desde Kobo...", "⚡", -1);
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

  ss.toast("Paso 2/2: Actualizando quincena actual...", "⚡", -1);
  try {
    var periodo = _periodoActivo();
    if (periodo) {
      var fi = new Date(periodo.fi), ff = new Date(periodo.ff);
      var horasReponer = _leerHorasReponerExistentes(periodo.tab);
      _generarReporteQuincena(fi, ff, periodo.label, periodo.tab, horasReponer);
      log.push("✅ Paso 2: Quincena '" + periodo.label + "' actualizada");
    } else {
      log.push("ℹ️ Paso 2: Sin quincena activa — ve a 📅 Quincena → Nueva quincena");
    }
  } catch(e) { errores.push("❌ Paso 2: " + e.message); }

  _alert(
    "⚡ IMPORTAR TODO\n\n" +
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

// Normaliza un valor de celda (Date obj o ISO string) a ms-epoch como string.
// Permite comparar timestamps sin importar si vienen de getValues() o parseCsv().
function _ckTs(val) {
  if (!val && val !== 0) return "";
  if (val instanceof Date) return isNaN(val) ? "" : String(val.getTime());
  var d = new Date(String(val).trim());
  return isNaN(d) ? String(val).trim().substring(0, 19) : String(d.getTime());
}

// Normaliza nombre de participante para dedup: minúsculas + guiones→espacios.
// Kobo devuelve "sindy_paola_lazaro_diaz"; la hoja guarda "Sindy Paola Lazaro Diaz".
function _ckPart(val) {
  return String(val||"").trim().toLowerCase().replace(/_/g, " ");
}

function _autoImportarKobo() {
  try { _importarKoboCore(true); } catch(e) { Logger.log("Error auto-import Kobo: " + e); }
}

function importarDesdeKobo() { _run(function() { _importarKoboCore(false); }); }

function _importarKoboCore(silencioso) {
  function _msg(m) { if (!silencioso) _alert(m); else Logger.log(m); }

  var resp = UrlFetchApp.fetch(CFG.KOBO_URL_CSV, { muteHttpExceptions: true });
  var code = resp.getResponseCode();
  if (code === 503) { _msg("⏳ Kobo ocupado (503). Espera 2 min e intenta de nuevo."); return; }
  if (code !== 200) throw new Error("Error Kobo HTTP " + code + ": " + resp.getContentText().substring(0,200));

  var datosRaw = Utilities.parseCsv(resp.getContentText(), ";");
  if (datosRaw.length < 2) { _msg("Kobo no devolvió registros."); return; }

  // Sin filtro de fecha — todo el histórico; dedup previene duplicados.

  // Excluir registros de Manufactura (columna "Destino" del formulario Kobo).
  // Estas personas no pertenecen a mi eelo y no deben contarse en horas/pagos.
  var colsDestino = detectarColumnas(datosRaw[0], []);
  var descartadosMan = 0;
  if (colsDestino.destino !== undefined) {
    var filtradoDestino = [datosRaw[0]];
    for (var fd = 1; fd < datosRaw.length; fd++) {
      var valDestino = String(datosRaw[fd][colsDestino.destino] || "").trim().toLowerCase();
      if (valDestino === "manufactura") { descartadosMan++; continue; }
      filtradoDestino.push(datosRaw[fd]);
    }
    datosRaw = filtradoDestino;
  }

  // Filtrar participantes oficiales
  var nombresOficiales2 = {};
  LISTA_OFICIAL.forEach(function(it) { nombresOficiales2[it[1]] = true; });
  var mapeoNombres2 = cargarMapeoNombres();
  var colsHdr2 = detectarColumnas(datosRaw[0], []);
  if (colsHdr2.participante !== undefined) {
    var filtrado2 = [datosRaw[0]];
    for (var fi2 = 1; fi2 < datosRaw.length; fi2++) {
      var rawN2 = String(datosRaw[fi2][colsHdr2.participante] || "").trim();
      var normN2 = rawN2 ? normalizarNombre(rawN2, mapeoNombres2) : "";
      if (_esValorAccion(rawN2) || normN2 && nombresOficiales2[normN2]) filtrado2.push(datosRaw[fi2]);
    }
    datosRaw = filtrado2;
  }

  // Filtrar columnas necesarias
  var datosNuevos = _filtrarColumnasKobo(datosRaw);

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
    _msg("✅ Importación inicial: " + (datosNuevos.length-1) + " registros.");
    return;
  }

  // Incremental — dedup por UUID; si UUID ausente, fallback a timestamp+participante
  var encNuevos = datosNuevos[0];
  var uuidColN  = _buscarIndice(encNuevos, "_uuid");
  var partColN  = detectarColumnas(encNuevos, []).participante;
  var datosEx   = hoja.getDataRange().getValues();
  var uuidColE  = _buscarIndice(datosEx[0], "_uuid");
  var partColE  = detectarColumnas(datosEx[0], []).participante;
  var uuidsExist = {}, compositeExist = {};
  for (var i=1; i<datosEx.length; i++) {
    var u = uuidColE >= 0 ? String(datosEx[i][uuidColE]||"").trim() : "";
    if (u) uuidsExist[u] = true;
    var tsK = _ckTs(datosEx[i][0]);
    var ptK = _ckPart(partColE >= 0 ? datosEx[i][partColE] : "");
    if (tsK && ptK) compositeExist[tsK+"|"+ptK] = true;
  }
  var filasNuevas = [];
  for (var j=1; j<datosNuevos.length; j++) {
    var uid = uuidColN >= 0 ? String(datosNuevos[j][uuidColN]||"").trim() : "";
    if (uid && uuidsExist[uid]) continue;
    var tsN = _ckTs(datosNuevos[j][0]);
    var ptN = _ckPart(partColN >= 0 ? datosNuevos[j][partColN] : "");
    if (tsN && ptN && compositeExist[tsN+"|"+ptN]) continue;
    filasNuevas.push(datosNuevos[j]);
  }
  if (filasNuevas.length === 0) {
    _msg("✅ Ya está al día. Sin registros nuevos." +
      (descartadosMan > 0 ? " (" + descartadosMan + " de Manufactura ignorados)" : ""));
    return;
  }

  hoja.getRange(hoja.getLastRow()+1,1,filasNuevas.length,filasNuevas[0].length).setValues(filasNuevas);
  _normalizarAccionSilencioso(hoja);
  _msg("✅ " + filasNuevas.length + " registros nuevos importados." +
    (descartadosMan > 0 ? "\n(" + descartadosMan + " registros de Manufactura ignorados)" : ""));
  try { actualizarDashboardVisual(); } catch(_) {}
}

// Llamado por el trigger instalable onOpen (tiene permisos completos)
function importarAlAbrir() {
  // Kobo asistencia: NO se importa al abrir — descarga miles de filas y tarda >10s.
  // El trigger de 6h (_autoImportarKobo) lo hace silencioso en segundo plano.
  // Importación manual: Admin → 📥 Datos Kobo → 📥 Importar desde Kobo
  try { importarEstipendioDesdeKobo(true); } catch(_) {}  // estipendio: ligero, OK al abrir
  try { _upsertHistorialActivo(); } catch(_) {}
  try { actualizarDetalleQuincena(); } catch(_) {}
  try { actualizarQuincenaActual(); } catch(_) {}
}

/*
 * Regenera la hoja "📋 Detalle Quincena" con el período activo actual.
 * Se llama en onOpen y desde el menú.
 */
function actualizarDetalleQuincena() {
  var periodo = _periodoActivo();
  if (!periodo) return;
  var fi  = new Date(periodo.fi);
  var ff  = new Date(periodo.ff);
  generarReporte("rango", fi, ff, null, "📋 Detalle Quincena");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var h  = ss.getSheetByName("📋 Detalle Quincena");
  if (h) ss.setActiveSheet(h);
}

// Elimina filas de DatosKobo anteriores al inicio del período activo (sin descargar nada)
function limpiarDatosKoboAntesDePeriodo() { _run(function() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var hoja    = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (!hoja || hoja.getLastRow() < 2) { _alert("DatosKobo está vacío."); return; }

  var periodo = _periodoActivo();
  if (!periodo) { _alert("No hay período activo. Configura la quincena primero."); return; }

  var corte = new Date(periodo.fi);
  corte.setDate(corte.getDate() - 7); // conservar 7 días de buffer antes del período
  corte.setHours(0,0,0,0);

  var total  = hoja.getLastRow() - 1;
  var datos  = hoja.getRange(2, 1, total, 1).getValues();
  var borrar = [];
  for (var i = datos.length - 1; i >= 0; i--) {
    var s = String(datos[i][0] || "").trim();
    var ts = new Date(s);
    if (isNaN(ts)) { var m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/); if (m) ts = new Date(m[3],m[2]-1,m[1]); }
    if (!isNaN(ts) && ts < corte) borrar.push(i + 2);
  }

  if (borrar.length === 0) { _alert("✅ No hay filas anteriores al período activo."); return; }

  // Borrar en bloques consecutivos de abajo hacia arriba
  var ini = borrar[0], fin = borrar[0], eliminadas = 0;
  for (var j = 1; j < borrar.length; j++) {
    if (borrar[j] === ini - 1) { ini = borrar[j]; }
    else { hoja.deleteRows(ini, fin-ini+1); eliminadas += fin-ini+1; ini = borrar[j]; fin = borrar[j]; }
  }
  hoja.deleteRows(ini, fin-ini+1); eliminadas += fin-ini+1;
  _alert("✅ " + eliminadas + " filas anteriores al " +
    Utilities.formatDate(corte, CFG.TIMEZONE, "dd/MM/yyyy") + " eliminadas.\nQuedan " + (hoja.getLastRow()-1) + " registros.");
}); }

// Elimina en-lugar las filas con fecha anterior a 2026 sin descargar nada de Kobo
function limpiarFilasAnteriores2026() { _run(function() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (!hoja || hoja.getLastRow() < 2) { _alert("DatosKobo está vacío."); return; }

  var total  = hoja.getLastRow() - 1; // sin encabezado
  var datos  = hoja.getRange(2, 1, total, 1).getValues(); // solo col A (start)
  var borrar = [];
  for (var i = datos.length - 1; i >= 0; i--) {
    var s    = String(datos[i][0] || "").trim();
    // Soporta ISO "2025-09-01..." y local "01/09/2025..."
    var anio = NaN;
    if (s.length >= 4) {
      var p4 = parseInt(s.substring(0, 4), 10);
      anio = (p4 >= 2000 && p4 <= 2099) ? p4 : NaN;
    }
    if (isNaN(anio)) {
      // Intenta extraer año de cualquier parte con regex
      var m = s.match(/\b(20\d{2})\b/);
      if (m) anio = parseInt(m[1], 10);
    }
    if (!isNaN(anio) && anio < 2026) borrar.push(i + 2); // +2 = fila real (1-based + encabezado)
  }

  if (borrar.length === 0) { _alert("✅ No hay filas anteriores a 2026. Todo limpio."); return; }

  // Borrar en bloques consecutivos (de abajo hacia arriba para no desplazar índices)
  var inicio = borrar[0], fin = borrar[0], eliminadas = 0;
  for (var j = 1; j < borrar.length; j++) {
    if (borrar[j] === inicio - 1) { inicio = borrar[j]; }
    else {
      hoja.deleteRows(inicio, fin - inicio + 1);
      eliminadas += fin - inicio + 1;
      inicio = borrar[j]; fin = borrar[j];
    }
  }
  hoja.deleteRows(inicio, fin - inicio + 1);
  eliminadas += fin - inicio + 1;

  _alert("✅ " + eliminadas + " filas anteriores a 2026 eliminadas.\nQuedan " + (hoja.getLastRow() - 1) + " registros.");
}); }

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

  // Sin filtro de fecha — reimporta TODO el histórico desde Kobo.
  // Se descartan: participantes no oficiales y registros de Manufactura.

  // Excluir registros de Manufactura (columna "Destino" del formulario Kobo)
  var colsDestinoR = detectarColumnas(datos[0], []);
  var descartadosMan2 = 0;
  if (colsDestinoR.destino !== undefined) {
    var filtradoDestinoR = [datos[0]];
    for (var fdr = 1; fdr < datos.length; fdr++) {
      var valDestinoR = String(datos[fdr][colsDestinoR.destino] || "").trim().toLowerCase();
      if (valDestinoR === "manufactura") { descartadosMan2++; continue; }
      filtradoDestinoR.push(datos[fdr]);
    }
    datos = filtradoDestinoR;
  }

  // Filtrar participantes no oficiales (solo las 33 de LISTA_OFICIAL)
  var nombresOficiales = {};
  LISTA_OFICIAL.forEach(function(it) { nombresOficiales[it[1]] = true; });
  var mapeoNombres = cargarMapeoNombres();
  var colsHdr = detectarColumnas(datos[0], []);
  var descartadosNP = 0;
  if (colsHdr.participante !== undefined) {
    var filtrado2 = [datos[0]];
    for (var fi2 = 1; fi2 < datos.length; fi2++) {
      var rawN = String(datos[fi2][colsHdr.participante] || "").trim();
      var normN = rawN ? normalizarNombre(rawN, mapeoNombres) : "";
      if (!normN || !nombresOficiales[normN]) { descartadosNP++; continue; }
      filtrado2.push(datos[fi2]);
    }
    datos = filtrado2;
  }

  // Filtrar columnas: solo conservar las necesarias para el sistema
  datos = _filtrarColumnasKobo(datos);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hOld = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  var hoja;
  if (hOld) {
    hOld.clearContents();
    hOld.clearFormats();
    hoja = hOld;
  } else {
    hoja = ss.insertSheet(CFG.HOJAS.DATOS_KOBO);
  }
  hoja.getRange(1,1,datos.length,datos[0].length).setValues(datos);
  hoja.getRange(1,1,1,datos[0].length).setFontWeight("bold").setBackground("#4a86e8").setFontColor("#fff");
  hoja.setFrozenRows(1);
  _limpiarColumnasKobo(hoja, datos[0]);
  _normalizarAccionSilencioso(hoja);
  _alert("✅ Reimportación completa: " + (datos.length-1) + " registros importados.\n" +
    (descartadosNP > 0 ? "(" + descartadosNP + " registros de participantes no oficiales eliminados)\n" : "✅ Todos los participantes son oficiales\n") +
    (descartadosMan2 > 0 ? "(" + descartadosMan2 + " registros de Manufactura ignorados)" : "✅ Sin registros de Manufactura"));
}); }

// Palabras clave que identifican columnas necesarias para el sistema
var _COLS_KOBO_NECESARIAS = [
  "start", "end",
  "ingreso", "egreso", "entrada", "salida", "accion", "acción",
  "seleccione", "ingreso_egreso",
  "participante", "nombre",
  "c_id", "_id",
  "terapia", "permiso", "comput", "subtipo"
];

// Devuelve una copia del array datos con solo las columnas necesarias
function _filtrarColumnasKobo(datos) {
  if (!datos || datos.length < 1) return datos;
  var enc = datos[0];
  var idxKeep = [];
  for (var i = 0; i < enc.length; i++) {
    var h = String(enc[i]).trim().toLowerCase();
    if (h && _COLS_KOBO_NECESARIAS.some(function(p){ return h.indexOf(p) !== -1; })) {
      idxKeep.push(i);
    }
  }
  if (idxKeep.length === enc.length) return datos; // nada que filtrar
  return datos.map(function(fila) {
    return idxKeep.map(function(i){ return fila[i]; });
  });
}

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

  // Normalizar columna Participante y poblar c_id desde LISTA_OFICIAL
  var cols = detectarColumnas(enc, []);
  if (cols.participante === undefined) return;
  var mapeo = cargarMapeoNombres();
  var lastRow = hoja.getLastRow();
  if (lastRow < 2) return;

  // Mapa inverso nombre_oficial → Creamos_ID desde LISTA_OFICIAL
  var mapaNameAId = {};
  LISTA_OFICIAL.forEach(function(it) { if (it[3]) mapaNameAId[it[1]] = it[3]; });

  // Si c_id no existe en la hoja, agregar la columna al final
  if (cols.creamos_id === undefined) {
    var nc = hoja.getLastColumn() + 1;
    hoja.getRange(1, nc).setValue("c_id")
        .setFontWeight("bold").setBackground("#4a86e8").setFontColor("#fff");
    cols.creamos_id = nc - 1; // 0-based
  }

  var nCols = Math.max(cols.participante, cols.creamos_id) + 1;
  var bloque = hoja.getRange(2, 1, lastRow - 1, nCols).getValues();
  var cambiadosP = 0, cambiadosID = 0;
  var nuevosP  = [];
  var nuevosID = [];

  bloque.forEach(function(fila) {
    var raw = String(fila[cols.participante] || "").trim();
    var cid = String(fila[cols.creamos_id]   || "").trim();

    // Resolver nombre oficial
    var porId = cid ? mapeo["id:" + cid] : null;
    var nombreFinal = porId || (raw ? normalizarNombre(raw, mapeo) : "");
    if (nombreFinal !== raw) cambiadosP++;

    // Resolver Creamos_ID desde el nombre oficial
    var idFinal = mapaNameAId[nombreFinal] || cid || "";
    if (idFinal !== cid) cambiadosID++;

    nuevosP.push([nombreFinal]);
    nuevosID.push([idFinal]);
  });

  var colP  = cols.participante + 1;
  var colID = cols.creamos_id   + 1;
  if (cambiadosP  > 0) hoja.getRange(2, colP,  lastRow - 1, 1).setValues(nuevosP);
  if (cambiadosID > 0) hoja.getRange(2, colID, lastRow - 1, 1).setValues(nuevosID);
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
        // Si había intercambio (participante col tiene acción), copiar nombre al lugar correcto
        if (cols.participante !== undefined) {
          var partVal = String(datos[f][cols.participante]||"").trim();
          if (_esValorAccion(partVal) && valRaw && !_esValorAccion(valRaw)) {
            hoja.getRange(f+1, cols.participante+1).setValue(valRaw);
          }
        }
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
  // DiasEstudio col B (idx 1) = Nombre; ListaTerapias col B (idx 1) = Nombre
  [{nm:"DiasEstudio",col:2},{nm:"ListaTerapias",col:2}].forEach(function(cfg){
    var h=ss.getSheetByName(cfg.nm);if(!h)return;
    var d=h.getDataRange().getValues(),cam=0;
    for(var f=1;f<d.length;f++){var n=String(d[f][cfg.col-1]||"").trim();if(n&&mapeo[n]&&mapeo[n]!==n){h.getRange(f+1,cfg.col).setValue(mapeo[n]);cam++;}}
    if(cam>0) act.push(cfg.nm+"("+cam+")");
  });
  return act.join(", ");
}

// ── Timestamp helper: corrige registros cuyo formulario se abrió un día y se envió otro ──
// Cuando start y end caen en días distintos → el día real es el de end (envío del formulario).
// Evita entradas "fantasma" de días anteriores por formularios abiertos y no cerrados.
function _resolverTsKobo(fila, cols) {
  var iS = cols.start, iE = cols.end;
  if (iS === undefined) return (iE !== undefined) ? new Date(fila[iE]) : null;
  var tsS = new Date(fila[iS]);
  if (isNaN(tsS)) return (iE !== undefined) ? new Date(fila[iE]) : null;
  if (iE === undefined) return tsS;
  var tsE = new Date(fila[iE]);
  if (isNaN(tsE)) return tsS;
  // mismo día → usar start (hora real de apertura del formulario)
  if (tsS.getFullYear()===tsE.getFullYear() &&
      tsS.getMonth()===tsE.getMonth() &&
      tsS.getDate()===tsE.getDate()) return tsS;
  // días distintos → usar end (día y hora de envío = llegada real)
  return tsE;
}

// ── Detección de columnas Kobo ────────────────────────────────

function detectarColumnas(encabezados, datosEjemplo) {
  var cols={};
  for(var i=0;i<encabezados.length;i++){
    var h=String(encabezados[i]).trim(), hLow=h.toLowerCase();
    if(hLow==="start"){cols.start=i;continue;}
    if(hLow==="end"){cols.end=i;continue;}
    if(hLow==="_submission_time"||hLow==="submission_time"){cols.submissionTime=i;continue;}
    if(hLow==="_uuid"){cols.uuid=i;continue;}
    if(hLow.indexOf("uuid")!==-1&&cols.uuid===undefined){cols.uuid=i;continue;}
    if(hLow==="c_id"||hLow==="_c_id"){cols.creamos_id=i;continue;}
    if(hLow.indexOf("participante")!==-1||hLow.indexOf("nombre")!==-1||hLow.indexOf("seleccione")!==-1){
      if(cols.participante===undefined)cols.participante=i; else if(cols.participante2===undefined)cols.participante2=i; continue;
    }
    if((hLow.indexOf("ingreso")!==-1||hLow.indexOf("entrada")!==-1)&&(hLow.indexOf("egreso")!==-1||hLow.indexOf("salida")!==-1)){cols.accionUnificada=i;continue;}
    if(hLow.indexOf("accion")!==-1||hLow.indexOf("acción")!==-1||hLow==="type"||hLow.indexOf("marcar")!==-1){cols.accionUnificada=i;continue;}
    if(hLow==="subtipo_egreso"){cols.subtipoEgreso=i;continue;}
    if(hLow==="destino"||hLow.indexOf("dirige")!==-1){cols.destino=i;continue;}
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
    // Columnas intercambiadas: acción contiene nombre → buscar acción en columna participante
    var looksLikeAction = _esValorAccion(vR) ||
      v.indexOf("terapia")!==-1 || v.indexOf("permiso")!==-1 || v.indexOf("comput")!==-1;
    if (!looksLikeAction && cols.participante !== undefined) {
      var altAccion = String(fila[cols.participante]||"").trim();
      if (_esValorAccion(altAccion)) { vR = altAccion; v = altAccion.toLowerCase(); }
    }
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

// Detecta si un string parece un valor de acción Kobo (no un nombre de persona)
function _esValorAccion(s) {
  var low = String(s).toLowerCase();
  return low.indexOf("entrada") !== -1 || low.indexOf("salida") !== -1 ||
         low.indexOf("ingreso") !== -1 || low.indexOf("egreso") !== -1 ||
         String(s).indexOf("🟢") !== -1 || String(s).indexOf("🔴") !== -1;
}

function obtenerParticipanteFila(fila, cols) {
  var n1=cols.participante!==undefined?String(fila[cols.participante]||"").trim():"";
  var n2=cols.participante2!==undefined?String(fila[cols.participante2]||"").trim():"";
  // Columnas intercambiadas: participante tiene acción → buscar nombre en columna acción
  if (n1 && _esValorAccion(n1) && cols.accionUnificada !== undefined) {
    var deAccion = String(fila[cols.accionUnificada]||"").trim();
    if (deAccion && !_esValorAccion(deAccion)) return deAccion;
  }
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
    .replace(/^[A-ZÁÉÍÓÚÑÜ]{4}\d{3,6}\s*/i, "")        // ID al inicio
    .replace(/\s*\([A-ZÁÉÍÓÚÑÜ]{4}\d{3,6}\)\s*/i, "")  // ID en paréntesis
    .replace(/^[_•\s]+/, "")                             // guiones bajos/bullets al inicio
    .replace(/_/g, " ");                                  // slugs: _ → espacio
  return s.replace(/\s+/g, " ").trim();
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

  // 0. Aliases hardcoded (nombres cortos, typos, slugs parciales de Kobo)
  ALIASES_KOBO.forEach(function(par) {
    var alias   = par[0]; // ya en minúsculas sin tildes
    var oficial = par[1];
    m[alias] = oficial;
    m[alias.replace(/\s+/g, "_")] = oficial; // versión slug también
  });

  // 1. Desde LISTA_OFICIAL: slug kobo + Creamos_ID → nombre oficial
  LISTA_OFICIAL.forEach(function(item) {
    var nombreOficial = item[1];
    var id            = item[3] || "";
    var slug          = item[4] || "";
    if (slug) {
      m[slug] = nombreOficial;
      m[slug.replace(/_/g, " ")] = nombreOficial;
    }
    if (id) m["id:" + id] = nombreOficial; // lookup por ID desde LISTA_OFICIAL
    var sinTildes = textoParaComparar(item[1]);
    if (sinTildes !== nombreOficial.toLowerCase()) m[sinTildes] = nombreOficial;
  });

  // 2. Desde PARTICIPANTES: mapear Creamos_ID → nombre oficial
  //    Permite resolver entradas Kobo como "Mayra García (MACI030373)"
  var hP = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.HOJAS.PARTICIPANTES);
  if (hP && hP.getLastRow() >= 2) {
    var datP = hP.getRange(2, 1, hP.getLastRow() - 1, 2).getValues();
    datP.forEach(function(r) {
      var id  = String(r[0] || "").trim();
      var nom = String(r[1] || "").trim();
      if (id && nom && _esCreamos_ID_real(id)) m["id:" + id] = nom;
    });
  }

  // 3. Desde hoja NombresCanonicos (aliases manuales extras)
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
  // Intento exacto con nombre original
  if (mapeo[nombre]) return mapeo[nombre];
  // Limpiar primero (strip _ iniciales, _ → espacios, quitar ID prefix)
  var limpio = limpiarNombre(nombre);
  if (mapeo[limpio]) return mapeo[limpio];
  // Intento sin tildes / minúsculas sobre nombre limpio
  var norm = textoParaComparar(limpio);
  if (mapeo[norm]) return mapeo[norm];
  // Intento slug (reemplazar espacios por _)
  var slug = norm.replace(/\s+/g, "_");
  if (mapeo[slug]) return mapeo[slug];
  // Intento por Creamos_ID extraído del nombre (ej. "Nombre (MACI030373)")
  var cod = extraerCodigo(nombre);
  if (cod && mapeo["id:" + cod]) return mapeo["id:" + cod];
  return limpio; // fallback: nombre limpio sin underscores ni ID
}

// ── Días de estudio y terapias ────────────────────────────────

function obtenerDiasEstudio() {
  // Esquema: A=Creamos_ID(0), B=Participante(1), C-I=días(2-8), J=Fecha_Inicio(9), K=Fecha_Fin(10)
  var mapa={},h=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DiasEstudio");
  if(!h)return mapa;
  var d=h.getDataRange().getValues();
  for(var f=1;f<d.length;f++){
    var p=String(d[f][1]||"").trim();if(!p)continue; // col B = Participante
    mapa[p]={
      dias:{1:d[f][2]==="X"||d[f][2]==="x",2:d[f][3]==="X"||d[f][3]==="x",
             3:d[f][4]==="X"||d[f][4]==="x",4:d[f][5]==="X"||d[f][5]==="x",
             5:d[f][6]==="X"||d[f][6]==="x",6:d[f][7]==="X"||d[f][7]==="x",
             0:d[f][8]==="X"||d[f][8]==="x"},
      fechaInicio:d[f][9]?new Date(d[f][9]):null,
      fechaFin:d[f][10]?new Date(d[f][10]):null
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
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var hojaP    = _sh(CFG.HOJAS.PARTICIPANTES);
  var hojaF    = _sh(CFG.HOJAS.FACTURACION);
  var nombreMes = CFG.MESES[mes-1];

  // Mapa de participantes: id → { tarifa, tieneFactura }
  var partRows = hojaP.getDataRange().getValues();
  var partMap  = {};
  for (var pi=1; pi<partRows.length; pi++) {
    var pid = String(partRows[pi][0]).trim();
    if (!pid) continue;
    var tarifa = parseFloat(partRows[pi][9]); // col J = Tarifa_Hora
    if (isNaN(tarifa) || tarifa <= 0) {
      var cat = String(partRows[pi][8]).trim().toUpperCase(); // col I = Categoria
      tarifa = CFG.CATEGORIAS[cat] || CFG.CATEGORIAS.C;
    }
    var t = String(partRows[pi][10]).trim().toLowerCase(); // col K = Tiene_Factura
    partMap[pid] = {
      nombre:       String(partRows[pi][1]).trim(),
      tarifa:       tarifa,
      tieneFactura: t === "sí" || t === "si"
    };
  }

  // Sumar horas a pagar por participante y quincena — leer directamente de DatosKobo
  var horas = {"1":{}, "2":{}};
  var hK = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (hK && hK.getLastRow() >= 2) {
    var enc  = hK.getRange(1, 1, 1, hK.getLastColumn()).getValues()[0];
    var raw  = hK.getRange(2, 1, hK.getLastRow()-1, hK.getLastColumn()).getValues();
    var cols = detectarColumnas(enc, raw.slice(0, 50));
    var mapeoNombres = cargarMapeoNombres();
    var diasEstudioMap = obtenerDiasEstudio();
    var listaTerapias  = obtenerListaTerapias();
    // Agrupar registros por participante+día
    var porDia = {};
    raw.forEach(function(fila) {
      var ts = _resolverTsKobo(fila, cols);
      if (!ts || isNaN(ts)) return;
      if (ts.getMonth()+1 !== mes || ts.getFullYear() !== anio) return;
      var nombreRaw = obtenerParticipanteFila(fila, cols);
      if (!nombreRaw) return;
      var nombre = normalizarNombre(nombreRaw, mapeoNombres) || limpiarNombre(nombreRaw);
      if (!nombre) return;
      var tipo = obtenerTipoRegistro(fila, cols);
      if (!tipo.esIngreso && !tipo.esEgreso) return;
      var dClave = nombre + "|" + ts.getFullYear() + "-" + ts.getMonth() + "-" + ts.getDate();
      if (!porDia[dClave]) porDia[dClave] = { nombre: nombre, ts: ts, regs: [] };
      var tsEnd = (cols.end !== undefined) ? new Date(fila[cols.end]) : null;
      porDia[dClave].regs.push({ ts: ts, tsEnd: tsEnd, tipo: tipo, esTerapia: tipo.esTerapia });
    });

    Object.keys(porDia).forEach(function(dClave) {
      var grupo  = porDia[dClave];
      var nombre = grupo.nombre;
      var diaTs  = grupo.ts;
      var esDiaEst = esDiaDeEstudio(nombre, diaTs, diasEstudioMap);
      var esTer    = !!(listaTerapias[nombre]);
      // Si es día de estudio o terapia, no se cuentan horas a pagar
      if (esDiaEst || esTer) return;

      var regs = grupo.regs.sort(function(a,b){ return a.ts - b.ts; });
      var totalH = 0;
      var entrada = null;
      regs.forEach(function(reg) {
        if (reg.tipo.esIngreso && !entrada) {
          entrada = reg.ts;
        } else if (reg.tipo.esEgreso && entrada) {
          var tSalida = (reg.tsEnd && !isNaN(reg.tsEnd) &&
                         (reg.tsEnd - entrada)/3600000 > 0 &&
                         (reg.tsEnd - entrada)/3600000 < 16)
                        ? reg.tsEnd : reg.ts;
          var diffH = (tSalida - entrada) / 3600000;
          if (diffH > 0 && diffH <= 16) totalH += diffH;
          entrada = null;
        }
      });
      if (entrada) totalH += CFG.HORAS_JORNADA_NORMAL; // entrada sin salida → estimar

      if (totalH <= 0) return;
      totalH = Math.round(totalH * 100) / 100;

      // Buscar id del participante
      var id = "";
      for (var pid in partMap) {
        if (partMap[pid].nombre === nombre) { id = pid; break; }
      }
      var q = diaTs.getDate() <= 15 ? "1" : "2";
      var k = id || nombre;
      horas[q][k] = Math.round(((horas[q][k]||0) + totalH)*100)/100;
    });
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
  try { actualizarDashboardVisual(); } catch(_) {}
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
  var tz=CFG.TIMEZONE;
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
  var tz=CFG.TIMEZONE;
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
    if (n) catMap[n] = String(r[8]).trim().toUpperCase() || "?"; // col I = Categoria
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
  push(["Generado: " + Utilities.formatDate(new Date(), CFG.TIMEZONE, "dd/MM/yyyy HH:mm"),
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
    "2️⃣  Calcular facturación del mes actual\n" +
    "3️⃣  Generar recibos de pago\n" +
    "4️⃣  Generar resumen de facturación en hoja\n\n" +
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

  // Paso 2: Calcular facturación
  try { _calcular(mes, anio); log.push("✅ 2️⃣  Facturación calculada — " + nombreMes + " " + anio); }
  catch(e) { log.push("⚠️ 2️⃣  Facturación: " + e.message); }

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
    log.push("✅ 3️⃣  " + generados + " recibos generados");
  } catch(e) { log.push("⚠️ 3️⃣  Recibos: " + e.message); }

  // Paso 4: Resumen en hoja
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var titulo = "Fact_" + nombreMes + "_" + anio;
    // Llamar directamente la lógica (sin _run para no anidar)
    generarResumenFacturacionEnHoja(mes, anio);
    log.push("✅ 4️⃣  Resumen generado en hoja '" + titulo + "'");
  } catch(e) { log.push("⚠️ 4️⃣  Resumen: " + e.message); }

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
  var tz   = CFG.TIMEZONE;
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
  var tz  = CFG.TIMEZONE;

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
// Crea o actualiza la hoja HistorialQuincenas — siempre reescribe el encabezado
function crearHojaHistorialQuincenas() { _run(function() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("HistorialQuincenas");
  if (!hoja) hoja = ss.insertSheet("HistorialQuincenas");

  // Asegurar columnas suficientes
  while (hoja.getMaxColumns() < 17) hoja.insertColumnsAfter(hoja.getMaxColumns(), 1);

  // Siempre actualizar encabezado (permite agregar columnas nuevas sin borrar datos)
  var enc = ["#","Período","Fecha_Inicio","Fecha_Fin","Fecha_Cierre",
             "Total_Q","Total_Horas","Participantes_Con_Horas","Participantes_Sin_Horas",
             "Retiradx_Período","Dias_Estudio_Período","Estipendio_Total",
             "Ciclos_Vida_Año","Ingreso_Promedio_Q","Horas_Formacion_Período","Promedio_Horas_Mes",
             "Notas"];
  hoja.getRange(1,1,1,enc.length).setValues([enc])
      .setFontWeight("bold").setBackground("#1a237e").setFontColor("#fff").setHorizontalAlignment("center");
  hoja.setFrozenRows(1);

  // Anchos de columna
  hoja.setColumnWidth(1, 40);  hoja.setColumnWidth(2, 160);
  hoja.setColumnWidth(3, 110); hoja.setColumnWidth(4, 110); hoja.setColumnWidth(5, 110);
  hoja.setColumnWidth(6, 110); hoja.setColumnWidth(7, 100);
  hoja.setColumnWidth(8, 130); hoja.setColumnWidth(9, 130);
  hoja.setColumnWidth(10,110); hoja.setColumnWidth(11,120); hoja.setColumnWidth(12,110);
  hoja.setColumnWidth(13,100); hoja.setColumnWidth(14,120);
  hoja.setColumnWidth(15,130); hoja.setColumnWidth(16,130); hoja.setColumnWidth(17,200);

  // Formatos de datos
  hoja.getRange("C2:E500").setNumberFormat("dd/MM/yyyy");
  hoja.getRange("F2:F500").setNumberFormat('"Q"#,##0.00');
  hoja.getRange("G2:G500").setNumberFormat("0.0");
  hoja.getRange("L2:L500").setNumberFormat('"Q"#,##0.00');
  hoja.getRange("N2:N500").setNumberFormat('"Q"#,##0.00');
  hoja.getRange("O2:O500").setNumberFormat("0.0");
  hoja.getRange("P2:P500").setNumberFormat("0.0");

  hoja.activate();
  _alert("✅ Hoja HistorialQuincenas actualizada con 17 columnas.\nUsa 'Cargar historial de quincenas pasadas' para llenar las filas existentes.");
}); }

// Calcula todos los KPIs para un período — usado por _guardar y _upsert
function _calcularKPIsHistorial(fi, ff) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var dIni = new Date(fi.getFullYear(), fi.getMonth(), fi.getDate());
  var dFin = new Date(ff.getFullYear(), ff.getMonth(), ff.getDate());
  var anio = ff.getFullYear();

  // Horas y pago del período
  var resumen = _calcularResumenPeriodo(fi, ff);
  var totalQ = 0, totalHrs = 0, conHoras = 0, sinHoras = 0;
  Object.keys(resumen).forEach(function(n) {
    var d = resumen[n];
    var base = Math.round(d.horas * d.tarifa * 100) / 100;
    var iva  = d.tieneFactura ? Math.round(base * CFG.IVA_PCT * 100) / 100 : 0;
    var bono = Math.round((d.bono || 0) * 100) / 100;
    totalQ   += Math.round((base + iva + bono) * 100) / 100;
    totalHrs += d.horas;
    if (d.horas > 0) conHoras++; else sinHoras++;
  });

  // Retiradx en este período
  var kRetiradx = 0;
  var hRet = ss.getSheetByName("Retiradx");
  if (hRet && hRet.getLastRow() > 1) {
    hRet.getRange(2,1,hRet.getLastRow()-1,1).getValues().forEach(function(r) {
      var d = new Date(r[0]); var dia = new Date(d.getFullYear(),d.getMonth(),d.getDate());
      if (!isNaN(d) && dia >= dIni && dia <= dFin) kRetiradx++;
    });
  }

  // Días de estudio distintos en este período (conteo de fechas únicas en DiasEstudio)
  var kDiasEst = 0;
  var hDE = ss.getSheetByName("DiasEstudio");
  if (hDE && hDE.getLastRow() > 1) {
    hDE.getRange(2,1,hDE.getLastRow()-1,1).getValues().forEach(function(r) {
      var d = new Date(r[0]); var dia = new Date(d.getFullYear(),d.getMonth(),d.getDate());
      if (!isNaN(d) && dia >= dIni && dia <= dFin) kDiasEst++;
    });
  }

  // Estipendio total en este período
  var kEstip = 0;
  var hES = ss.getSheetByName("Estipendio");
  if (hES && hES.getLastRow() > 1) {
    hES.getRange(2,1,hES.getLastRow()-1,4).getValues().forEach(function(r) {
      var d = new Date(r[0]); var dia = new Date(d.getFullYear(),d.getMonth(),d.getDate());
      if (!isNaN(d) && dia >= dIni && dia <= dFin) kEstip += parseFloat(r[3])||0;
    });
  }

  // Indicador 1: Ciclos de Vida completados EN EL AÑO (no solo en el período)
  var kCiclosAnio = 0;
  var hCV = ss.getSheetByName("CiclosVida");
  if (hCV && hCV.getLastRow() > 1) {
    hCV.getRange(2,1,hCV.getLastRow()-1,1).getValues().forEach(function(r) {
      var d = new Date(r[0]);
      if (!isNaN(d) && d.getFullYear() === anio) kCiclosAnio++;
    });
  }

  // Indicador 2: Ingreso bruto promedio por participante este período
  var kIngresoProm = conHoras > 0 ? Math.round((totalQ / conHoras) * 100) / 100 : 0;

  // Indicador 5: Horas de formación en el período
  // = sesiones de DatosKobo en días de estudio × 7h (usando esDiaDeEstudio por participante)
  var kHrsFormacion = 0;
  try {
    var hK = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
    if (hK && hK.getLastRow() > 1) {
      var encK  = hK.getRange(1,1,1,hK.getLastColumn()).getValues()[0];
      var rawK  = hK.getRange(2,1,hK.getLastRow()-1,hK.getLastColumn()).getValues();
      var colsK = detectarColumnas(encK, rawK.slice(0,50));
      var diasEstMap = obtenerDiasEstudio();
      var mapeoN = cargarMapeoNombres();
      var vistos = {};
      rawK.forEach(function(fila) {
        var ts = _resolverTsKobo(fila, colsK);
        if (!ts || isNaN(ts)) return;
        var dia = new Date(ts.getFullYear(), ts.getMonth(), ts.getDate());
        if (dia < dIni || dia > dFin) return;
        var nRaw = obtenerParticipanteFila(fila, colsK);
        if (!nRaw) return;
        var nombre = normalizarNombre(nRaw, mapeoN) || limpiarNombre(nRaw);
        if (!nombre) return;
        if (!esDiaDeEstudio(nombre, ts, diasEstMap)) return;
        var clave = nombre + "|" + ts.getFullYear() + "-" + ts.getMonth() + "-" + ts.getDate();
        vistos[clave] = true;
      });
      kHrsFormacion = Math.round(Object.keys(vistos).length * CFG.HORAS_JORNADA_NORMAL * 10) / 10;
    }
  } catch(e) { Logger.log("kHrsFormacion: " + e.message); }

  // Indicador 6: Promedio horas laborales mensuales
  // Quincena ≈ 15 días ≈ 0.5 mes → promedio mensual = (totalHrs / conHoras) × 2
  var diasPeriodo = Math.max(1, Math.round((dFin - dIni) / 86400000) + 1);
  var factorMensual = 30 / diasPeriodo;
  var kPromHrsMes = conHoras > 0
    ? Math.round((totalHrs / conHoras) * factorMensual * 10) / 10
    : 0;

  return {
    totalQ: Math.round(totalQ*100)/100, totalHrs: Math.round(totalHrs*100)/100,
    conHoras: conHoras, sinHoras: sinHoras,
    kRetiradx: kRetiradx, kDiasEst: kDiasEst,
    kEstip: Math.round(kEstip*100)/100,
    kCiclosAnio: kCiclosAnio, kIngresoProm: kIngresoProm,
    kHrsFormacion: kHrsFormacion, kPromHrsMes: kPromHrsMes
  };
}

// Aplica formatos a la fila recién escrita en HistorialQuincenas
function _formatearFilaHistorial(hH, fila) {
  hH.getRange(fila, 3, 1, 3).setNumberFormat("dd/MM/yyyy");
  hH.getRange(fila, 6).setNumberFormat('"Q"#,##0.00');
  hH.getRange(fila, 7).setNumberFormat("0.0");
  hH.getRange(fila, 12).setNumberFormat('"Q"#,##0.00');
  hH.getRange(fila, 14).setNumberFormat('"Q"#,##0.00');
  hH.getRange(fila, 15).setNumberFormat("0.0");
  hH.getRange(fila, 16).setNumberFormat("0.0");
}

// Guarda una fila en HistorialQuincenas con todos los datos del período que se está cerrando
function _guardarHistorialQuincena(periodo) {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var hH  = ss.getSheetByName("HistorialQuincenas");
    if (!hH) { crearHojaHistorialQuincenas(); hH = ss.getSheetByName("HistorialQuincenas"); }
    if (!hH) return;
    var fi = new Date(periodo.fi), ff = new Date(periodo.ff);
    var k  = _calcularKPIsHistorial(fi, ff);
    var num = hH.getLastRow();
    var filaArr = [num, periodo.label, fi, ff, new Date(),
      k.totalQ, k.totalHrs, k.conHoras, k.sinHoras,
      k.kRetiradx, k.kDiasEst, k.kEstip,
      k.kCiclosAnio, k.kIngresoProm, k.kHrsFormacion, k.kPromHrsMes, ""];
    hH.appendRow(filaArr);
    var lr = hH.getLastRow();
    _formatearFilaHistorial(hH, lr);
    var bg = (num % 2 === 0) ? "#e8eaf6" : "#ffffff";
    hH.getRange(lr, 1, 1, filaArr.length).setBackground(bg);
  } catch(e) { Logger.log("_guardarHistorialQuincena: " + e.message); }
}

// Llena HistorialQuincenas con todas las quincenas cerradas que no estén ya registradas
function backfillHistorialQuincenas() { _run(function() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var hP  = ss.getSheetByName(CFG.HOJAS.PERIODOS);
  if (!hP || hP.getLastRow() < 2) { _alert("No hay períodos en PERIODOS."); return; }

  var hH = ss.getSheetByName("HistorialQuincenas");
  if (!hH) { crearHojaHistorialQuincenas(); hH = ss.getSheetByName("HistorialQuincenas"); }

  // Mapa label → fila en HistorialQuincenas (para detectar filas incompletas)
  var mapaFilas = {};
  if (hH.getLastRow() > 1) {
    var existData = hH.getRange(2, 2, hH.getLastRow()-1, 16).getValues();
    existData.forEach(function(r, idx) {
      var label = String(r[0]).trim();
      if (label) mapaFilas[label] = { fila: idx + 2, ciclos: r[11] }; // col 13 = idx 11 en 0-based desde col B
    });
  }

  var datos = hP.getDataRange().getValues();
  var agregados = 0, actualizados = 0;
  for (var i = 1; i < datos.length; i++) {
    var estado = String(datos[i][4]).trim();
    var label  = String(datos[i][1]).trim();
    if (estado !== "Cerrado") continue;
    var fi = datos[i][2], ff = datos[i][3];
    if (!fi || !ff) continue;

    if (!mapaFilas[label]) {
      // Fila nueva
      _guardarHistorialQuincena({ label: label, fi: fi, ff: ff, fila: i+1 });
      agregados++;
    } else if (mapaFilas[label].ciclos === "" || mapaFilas[label].ciclos === null || mapaFilas[label].ciclos === undefined) {
      // Fila existente pero con columnas nuevas vacías → rellenar solo cols 13-16
      var k = _calcularKPIsHistorial(new Date(fi), new Date(ff));
      var filaH = mapaFilas[label].fila;
      hH.getRange(filaH, 13, 1, 4).setValues([[k.kCiclosAnio, k.kIngresoProm, k.kHrsFormacion, k.kPromHrsMes]]);
      _formatearFilaHistorial(hH, filaH);
      actualizados++;
    }
  }
  _alert("✅ Historial actualizado.\n" +
    (agregados  > 0 ? agregados  + " quincena(s) nueva(s) agregadas.\n" : "") +
    (actualizados > 0 ? actualizados + " fila(s) con indicadores completados.\n" : "") +
    (agregados + actualizados === 0 ? "Todo ya estaba al día." : ""));
}); }

// Actualiza (o inserta) la fila de la quincena activa en HistorialQuincenas — silencioso
function _upsertHistorialActivo() {
  try {
    var periodo = _periodoActivo();
    if (!periodo) return;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hH = ss.getSheetByName("HistorialQuincenas");
    if (!hH) return;
    var fi = new Date(periodo.fi), ff = new Date(periodo.ff);
    var k  = _calcularKPIsHistorial(fi, ff);
    var valores = [periodo.label, fi, ff, "— activo —",
                   k.totalQ, k.totalHrs, k.conHoras, k.sinHoras,
                   k.kRetiradx, k.kDiasEst, k.kEstip,
                   k.kCiclosAnio, k.kIngresoProm, k.kHrsFormacion, k.kPromHrsMes, ""];
    // Buscar fila existente con mismo label
    var filaExist = -1;
    if (hH.getLastRow() > 1) {
      var labels = hH.getRange(2, 2, hH.getLastRow()-1, 1).getValues();
      for (var i = 0; i < labels.length; i++) {
        if (String(labels[i][0]).trim() === periodo.label) { filaExist = i + 2; break; }
      }
    }
    if (filaExist > 0) {
      hH.getRange(filaExist, 2, 1, valores.length).setValues([valores]);
      _formatearFilaHistorial(hH, filaExist);
    } else {
      var num = hH.getLastRow();
      hH.appendRow([num].concat(valores));
      var lr = hH.getLastRow();
      hH.getRange(lr, 1, 1, valores.length+1).setBackground("#e8f5e9"); // verde = activa
      _formatearFilaHistorial(hH, lr);
    }
  } catch(e) { Logger.log("_upsertHistorialActivo: " + e.message); }
}

function cerrarQuincenaYCrearSiguiente() { _run(function() {
  var periodo = _periodoActivo();
  if (!periodo) { _alert("No hay quincena activa. Usa 'Configurar nueva quincena'."); return; }

  var ui  = SpreadsheetApp.getUi();
  var tz  = CFG.TIMEZONE;
  var hP  = _sh(CFG.HOJAS.PERIODOS);

  var r = ui.alert("Cerrar quincena",
    "¿Cerrar el período " + periodo.label + "?\n\n" +
    "Total: Q" + (periodo.total || "—") + "\n\n" +
    "Se marcará como Cerrado y se configurará la siguiente quincena.",
    ui.ButtonSet.OK_CANCEL);
  if (r !== ui.Button.OK) return;

  // Guardar historial antes de cerrar
  _guardarHistorialQuincena(periodo);

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
  var tz  = CFG.TIMEZONE;
  return Utilities.formatDate(fi, tz, "dd/MM/yyyy") + " al " +
         Utilities.formatDate(ff, tz, "dd/MM/yyyy");
}

function _tabNombreQuincena(fi, ff) {
  var tz = CFG.TIMEZONE;
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
 * luego suma horas reales desde DatosKobo.
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
      codigo:       _esCreamos_ID_real(info.id) ? info.id : "",
      estipendio:   info.estipendio || 0,
      bono:         0
    };
  });

  // Construir mapa inverso ID → nombre para matching robusto en Bonos y Estipendio
  var mapaIdANombre = {};
  Object.keys(resultado).forEach(function(nom) {
    var id = resultado[nom].id;
    if (id) mapaIdANombre[id] = nom;
  });
  function _resolverNombreAux(id, nombre) {
    if (id && mapaIdANombre[id]) return mapaIdANombre[id];
    return nombre || "";
  }

  // Sumar bonos del período desde hoja "Bonos"
  var hBonos = ss.getSheetByName("Bonos");
  if (hBonos && hBonos.getLastRow() >= 2) {
    var datBonos = hBonos.getRange(2, 1, hBonos.getLastRow() - 1, 4).getValues();
    datBonos.forEach(function(b) {
      var fecha = new Date(b[0]);
      if (isNaN(fecha)) return;
      var dia = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate());
      if (dia < dIni || dia > dFin) return;
      var nombreBono = _resolverNombreAux(String(b[1]||"").trim(), String(b[2]||"").trim());
      var monto = parseFloat(b[3]) || 0;
      if (!nombreBono || !monto) return;
      if (resultado[nombreBono]) {
        resultado[nombreBono].bono = Math.round((resultado[nombreBono].bono + monto) * 100) / 100;
      }
    });
  }

  // Sumar estipendios del período desde hoja "Estipendio"
  var hEstip = ss.getSheetByName("Estipendio");
  if (hEstip && hEstip.getLastRow() >= 2) {
    var datEstip = hEstip.getRange(2, 1, hEstip.getLastRow() - 1, 4).getValues();
    datEstip.forEach(function(e) {
      var fecha = new Date(e[0]);
      if (isNaN(fecha)) return;
      var dia = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate());
      if (dia < dIni || dia > dFin) return;
      var nombreEstip = _resolverNombreAux(String(e[1]||"").trim(), String(e[2]||"").trim());
      var monto = parseFloat(e[3]) || 0;
      if (!nombreEstip || !monto) return;
      if (resultado[nombreEstip]) {
        resultado[nombreEstip].estipendio = Math.round((resultado[nombreEstip].estipendio + monto) * 100) / 100;
      }
    });
  }

  // Sumar horas reales desde DatosKobo — procesamiento por participante+DÍA
  // (evita emparejar entrada de día 2 con salida de día 3, y cuenta sin-salida correctamente)
  var hK = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (!hK || hK.getLastRow() < 2) return resultado;

  var enc  = hK.getRange(1, 1, 1, hK.getLastColumn()).getValues()[0];
  var raw  = hK.getRange(2, 1, hK.getLastRow()-1, hK.getLastColumn()).getValues();
  var cols = detectarColumnas(enc, raw.slice(0, 50));
  var mapeoNombres = cargarMapeoNombres();

  // Agrupar por participante+día
  var porPartDia = {};

  raw.forEach(function(fila) {
    var ts = _resolverTsKobo(fila, cols);
    if (!ts || isNaN(ts)) return;

    var dia = new Date(ts.getFullYear(), ts.getMonth(), ts.getDate());
    if (dia < dIni || dia > dFin) return;

    var nombreRaw = obtenerParticipanteFila(fila, cols);
    if (!nombreRaw) return;
    var nombre = normalizarNombre(nombreRaw, mapeoNombres) || limpiarNombre(nombreRaw);
    if (!nombre) return;

    var tipo = obtenerTipoRegistro(fila, cols);
    if (!tipo.esIngreso && !tipo.esEgreso) return;

    var dClave = nombre + "|" + ts.getFullYear() + "-" + ts.getMonth() + "-" + ts.getDate();
    if (!porPartDia[dClave]) porPartDia[dClave] = { nombre: nombre, ing: [], egr: [] };

    var tsEnd = (cols.end !== undefined) ? new Date(fila[cols.end]) : null;
    if (tipo.esIngreso) {
      porPartDia[dClave].ing.push(ts);
    } else {
      // Para salidas usar tsEnd si razonable (<16h)
      var tSal = (tsEnd && !isNaN(tsEnd) && (tsEnd - ts) >= 0 && (tsEnd - ts) < 57600000) ? tsEnd : ts;
      porPartDia[dClave].egr.push(tSal);
    }
  });

  // Rescatar sesiones huérfanas: entrada con nombre falso ("Entrada" u otro valor de acción)
  // + salida registrada bajo el nombre real del participante en el mismo día.
  // Construir pool de entradas falsas por fecha
  var fakePorDia = {};
  Object.keys(porPartDia).forEach(function(dClave) {
    var nomPart = porPartDia[dClave].nombre;
    if (!_esValorAccion(nomPart)) return;
    var datePart = dClave.split("|").slice(1).join("|");
    if (!fakePorDia[datePart]) fakePorDia[datePart] = [];
    porPartDia[dClave].ing.forEach(function(ts) { fakePorDia[datePart].push(ts); });
    delete porPartDia[dClave]; // eliminar participante falso
  });
  // Asignar entradas falsas a participantes reales que tienen salida sin entrada ese día
  Object.keys(porPartDia).forEach(function(dClave) {
    var d = porPartDia[dClave];
    if (d.ing.length > 0 || d.egr.length === 0) return;
    var datePart = dClave.split("|").slice(1).join("|");
    var fakes = fakePorDia[datePart];
    if (!fakes || fakes.length === 0) return;
    fakes.sort(function(a,b){return a-b;});
    d.ing.push(fakes.shift());
  });

  // Por cada día: earliest entrada + latest salida → horas del día
  Object.keys(porPartDia).forEach(function(dClave) {
    var d = porPartDia[dClave];
    var nombre = d.nombre;

    if (!resultado[nombre]) {
      var info2b = _buscarInfoParticipante(mapa, nombre);
      resultado[nombre] = { horas: 0, tarifa: info2b.tarifa || CFG.CATEGORIAS.C,
        tieneFactura: info2b.tieneFactura || false, categoria: info2b.categoria || "?",
        id: info2b.id || "", codigo: info2b.id || "",
        estipendio: info2b.estipendio || 0, bono: 0 };
    }

    if (d.ing.length === 0) return; // sin entrada → no se cuenta

    d.ing.sort(function(a,b){return a-b;});
    var entrada = d.ing[0]; // earliest entrada del día

    if (d.egr.length === 0) {
      // Sin salida → no se cuenta (entrada sola no es sesión completa)
      return;
    }

    d.egr.sort(function(a,b){return b-a;});
    var salida = d.egr[0]; // latest salida del día

    var diffH = (salida - entrada) / 3600000;
    if (diffH > 0 && diffH <= 16) {
      resultado[nombre].horas = Math.round((resultado[nombre].horas + diffH) * 100) / 100;
    }
    // Si diferencia es 0 o incoherente (>16h) → no se cuenta ese día
  });


  return resultado;
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
  var tz      = CFG.TIMEZONE;
  var ts      = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm");

  // Recrear o limpiar hoja
  var h = ss.getSheetByName(tabNombre);
  if (h) {
    h.clearContents();
    h.clearFormats();
  } else {
    h = ss.insertSheet(tabNombre);
  }

  // Asegurar suficientes columnas (ahora usamos 15)
  if (h.getMaxColumns() < 15) h.insertColumnsAfter(h.getMaxColumns(), 15 - h.getMaxColumns());

  // ── Fila 1: referencia de tarifas ──────────────────────────────
  h.getRange(1, 1, 1, 15).setValues([["Q16.50","Q15.75","Q15.00","Q14.00","","","","","","","","","","",""]]);
  ["#639922","#4285f4","#fbbc04","#ea4335"].forEach(function(c, i) {
    h.getRange(1, i+1).setBackground(c).setFontColor("#ffffff").setFontWeight("bold")
     .setHorizontalAlignment("center");
  });

  // ── Fila 2: título del período ─────────────────────────────────
  h.getRange(2, 1, 1, 15).merge()
   .setValue("Período: " + label)
   .setBackground("#f8f9fa").setFontWeight("bold").setFontSize(11)
   .setHorizontalAlignment("center")
   .setBorder(true,true,true,true,null,null,"#dadce0",SpreadsheetApp.BorderStyle.SOLID);

  // ── Fila 3: encabezados de columnas ───────────────────────────
  // A=#  B=Participante  C=Fact.  D=ID  E=spacer
  // F=Total hrs  G=Hrs reponer  H=Total a pagar
  // I=Monto Base  J=IVA 5%  K=Bono  L=Estipendio  M=Total org paga  N=Redondeo  O=Neto part.
  var encabezados = ["#","Participante","Fact.","ID","",
    "Total hrs","Hrs reponer","Total a pagar",
    "Monto Base","IVA 5%","Bono","Estipendio","Total org paga","Redondeo","Neto part."];
  h.getRange(3, 1, 1, 15).setValues([encabezados])
   .setBackground("#546e7a").setFontColor("#ffffff").setFontWeight("bold")
   .setHorizontalAlignment("center")
   .setBorder(true,true,true,true,null,null,"#37474f",SpreadsheetApp.BorderStyle.SOLID);

  // Colores de fondo por categoría (claro = lectura fácil)
  var BG_CAT = { A: CFG.COLORES_CAT.A.bgClaro, B: CFG.COLORES_CAT.B.bgClaro, C: CFG.COLORES_CAT.C.bgClaro, D: CFG.COLORES_CAT.D.bgClaro, "?": "#f3f3f3" };

  // ── Filas de datos: una por participante ───────────────────────
  var nombres = Object.keys(resumen).sort(function(a,b){ return a.localeCompare(b,"es"); });
  var filaActual = 4;
  var sumBase = 0, sumIVA = 0, sumBono = 0, sumEstip = 0, sumTotal = 0, totalGeneral = 0, sumHoras = 0;
  var num = 1;
  var conFactura = 0, sinFactura = 0;

  // Acumular todo en un array para un solo setValues masivo (más rápido)
  var bloqueValores = [];
  var bloqueFilas   = [];   // {fila, cat, tieneFactura, tieneBono, tieneEstip}

  nombres.forEach(function(nombre) {
    var d = resumen[nombre];
    var hReponer = prevHorasReponer[nombre] || 0;

    // ── Cálculo exacto (2 decimales en cada paso) ─────────────────
    var horas   = Math.round(d.horas   * 100) / 100;
    var hTotal  = Math.round((horas + hReponer) * 100) / 100;
    var base    = Math.round(hTotal * d.tarifa * 100) / 100;
    var iva     = d.tieneFactura ? Math.round(base * CFG.IVA_PCT * 100) / 100 : 0;
    var bono    = Math.round((d.bono || 0) * 100) / 100;
    var estip   = Math.round((d.estipendio || 0) * 100) / 100;
    var orgPaga = Math.round((base + iva + bono) * 100) / 100;  // estipendio NO entra al total org
    var redond  = Math.round(orgPaga);
    var neto    = Math.round((base + bono) * 100) / 100; // participante retiene base+bono; IVA va a SAT

    sumBase  += base;
    sumIVA   += iva;
    sumBono  += bono;
    sumEstip += estip;
    sumTotal += redond;
    totalGeneral += redond;
    sumHoras += horas;
    if (d.tieneFactura) conFactura++; else sinFactura++;

    var factInd = d.tieneFactura ? "★ Sí" : "—";

    bloqueValores.push([
      num++, nombre, factInd, d.codigo || "", "",
      horas || "",
      hReponer > 0 ? hReponer : "",
      hTotal || "",
      base || "",
      iva > 0 ? iva : "",
      bono > 0 ? bono : "",
      estip > 0 ? estip : "",
      orgPaga || "",
      redond || "",
      neto || ""
    ]);
    bloqueFilas.push({ fila: filaActual, cat: d.categoria || "?", tieneFactura: d.tieneFactura,
                       tieneBono: bono > 0, tieneEstip: estip > 0, sinHoras: horas === 0 });
    filaActual++;
  });

  // Escribir todos los valores de una vez
  if (bloqueValores.length > 0) {
    h.getRange(4, 1, bloqueValores.length, 15).setValues(bloqueValores);
  }

  // Aplicar formato fila a fila (colores por categoría + números)
  bloqueFilas.forEach(function(bf) {
    // Filas sin horas = gris claro + itálica para distinguirlas
    var bg = bf.sinHoras ? "#f5f5f5" : (BG_CAT[bf.cat] || BG_CAT["?"]);
    h.getRange(bf.fila, 1, 1, 15)
     .setBackground(bg).setFontFamily("Arial").setFontSize(10).setVerticalAlignment("middle")
     .setFontStyle(bf.sinHoras ? "italic" : "normal")
     .setFontColor(bf.sinHoras ? "#9e9e9e" : "#000000");
    h.getRange(bf.fila, 1).setHorizontalAlignment("center").setFontWeight("bold");
    h.getRange(bf.fila, 2, 1, 3).setHorizontalAlignment("left");
    h.getRange(bf.fila, 6, 1, 10).setHorizontalAlignment("right");
    // Col C (Fact.) — color especial si tiene factura
    if (bf.tieneFactura) {
      h.getRange(bf.fila, 3)
       .setBackground("#fce8e6").setFontColor("#c5221f").setFontWeight("bold")
       .setHorizontalAlignment("center");
    } else {
      h.getRange(bf.fila, 3).setFontColor("#9aa0a6").setHorizontalAlignment("center");
    }
    // Bono (col 11) — azul si tiene bono
    if (bf.tieneBono) {
      h.getRange(bf.fila, 11).setBackground("#e3f2fd").setFontColor("#1565c0").setFontWeight("bold");
    }
    // Estipendio (col 12) — verde si tiene estipendio
    if (bf.tieneEstip) {
      h.getRange(bf.fila, 12).setBackground("#e8f5e9").setFontColor("#2e7d32").setFontWeight("bold");
    }
    // Formato moneda cols I–O (9–15)
    h.getRange(bf.fila, 9, 1, 7).setNumberFormat('"Q"#,##0.00');
    h.getRange(bf.fila, 1, 1, 15)
     .setBorder(null,null,true,null,null,null,"#cccccc",SpreadsheetApp.BorderStyle.SOLID);
  });

  // ── Bloque de totales desglosados ──────────────────────────────
  filaActual++;

  // Fila: total de horas
  h.getRange(filaActual, 1, 1, 15).setValues([
    ["","Total de horas registradas","","","","",
     Math.round(sumHoras*100)/100,"","","","","","","",""]
  ]);
  h.getRange(filaActual,2).setFontStyle("italic").setFontColor("#555555");
  h.getRange(filaActual,7).setBackground("#e8eaf6").setFontWeight("bold")
   .setHorizontalAlignment("right").setNumberFormat('#,##0.00" hrs"');
  filaActual++;

  // Fila: subtotal monto base
  h.getRange(filaActual, 1, 1, 15).setValues([
    ["","Monto base (sin IVA)","","","","","","",
     Math.round(sumBase*100)/100,"","","","","",""]
  ]);
  h.getRange(filaActual,2).setFontStyle("italic").setFontColor("#555555");
  h.getRange(filaActual,9).setNumberFormat('"Q"#,##0.00').setBackground("#f8f9fa");
  filaActual++;

  // Fila: total IVA (Declaraguate)
  h.getRange(filaActual, 1, 1, 15).setValues([
    ["","IVA 5% total a declarar (Declaraguate)","","","","","","",
     "","","","",Math.round(sumIVA*100)/100,"",""]
  ]);
  h.getRange(filaActual,2).setFontWeight("bold").setFontColor("#c5221f");
  h.getRange(filaActual,13)
   .setNumberFormat('"Q"#,##0.00').setBackground("#fce8e6").setFontColor("#c5221f").setFontWeight("bold");
  filaActual++;

  // Fila: total bonos (si hay)
  if (sumBono > 0) {
    h.getRange(filaActual, 1, 1, 15).setValues([
      ["","Total bonos","","","","","","",
       "","",Math.round(sumBono*100)/100,"","","",""]
    ]);
    h.getRange(filaActual,2).setFontStyle("italic").setFontColor("#1565c0");
    h.getRange(filaActual,11).setNumberFormat('"Q"#,##0.00').setBackground("#e3f2fd").setFontColor("#1565c0");
    filaActual++;
  }

  // Fila: total estipendios (si hay)
  if (sumEstip > 0) {
    h.getRange(filaActual, 1, 1, 15).setValues([
      ["","Total estipendios","","","","","","",
       "","","",Math.round(sumEstip*100)/100,"","",""]
    ]);
    h.getRange(filaActual,2).setFontStyle("italic").setFontColor("#2e7d32");
    h.getRange(filaActual,12).setNumberFormat('"Q"#,##0.00').setBackground("#e8f5e9").setFontColor("#2e7d32");
    filaActual++;
  }

  // Fila: GRAN TOTAL (lo que paga la organización)
  h.getRange(filaActual, 1, 1, 15).setValues([
    ["","TOTAL QUE PAGA LA ORGANIZACIÓN","","","","","","",
     "","","","","",totalGeneral,""]
  ]);
  h.getRange(filaActual,2).setFontWeight("bold").setFontSize(11);
  h.getRange(filaActual,14)
   .setBackground("#00c853").setFontColor("#ffffff").setFontWeight("bold")
   .setFontSize(12).setHorizontalAlignment("center").setNumberFormat('"Q"#,##0.00');
  filaActual++;

  // ── Fila leyenda + resumen ─────────────────────────────────────
  filaActual++;
  h.getRange(filaActual, 1, 1, 15).setValues([[
    "A=Q16.50","B=Q15.75","C=Q15.00","D=Q14.00","",
    "★ = emite factura","","",
    "Con factura: "+conFactura,"Sin factura: "+sinFactura,"","","","",""
  ]]);
  [BG_CAT.A,BG_CAT.B,BG_CAT.C,BG_CAT.D].forEach(function(c,i){
    h.getRange(filaActual,i+1).setBackground(c).setFontSize(9)
     .setHorizontalAlignment("center").setFontWeight("bold");
  });
  h.getRange(filaActual,6,1,4).setFontSize(9).setFontColor("#555555");

  // ── Fila timestamp ─────────────────────────────────────────────
  filaActual++;
  h.getRange(filaActual, 1, 1, 15).merge()
   .setValue("Actualizado: " + ts + "  |  " + nombres.length + " participantes  |  " +
             nombres.filter(function(n){ return resumen[n].horas > 0; }).length + " con horas registradas")
   .setFontSize(8).setFontColor("#9aa0a6").setHorizontalAlignment("right");

  // ── Ancho de columnas (15 cols) ────────────────────────────────
  [35, 220, 60, 20, 20, 90, 100, 95, 95, 80, 90, 90, 95, 90, 90]
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

  // Leer PARTICIPANTES (A→K = cols 1-11)
  var datos = hP.getRange(2, 1, lastRow - 1, 15).getValues(); // A→O

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
    var cat    = String(r[12] || "").trim().toUpperCase();  // col M = Categoria (idx 12)
    var tieneFact = String(r[14]||"").trim();               // col O = Tiene_Factura (idx 14)
    if (!nombre) return;
    filas.push([i+1, id, nombre, tieneFact||"Sí", cat]);
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
    var actual = String(hP.getRange(filaP, 15).getValue()).trim(); // col O = Tiene_Factura
    if (actual !== valor) {
      hP.getRange(filaP, 15).setValue(valor);
      actualizados++;
    }
  }

  // Contar totales
  var total = hP.getLastRow()-1;
  var conFact = 0;
  if (total > 0) {
    var vals = hP.getRange(2,15,total,1).getValues(); // col O = Tiene_Factura
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
  var tz  = CFG.TIMEZONE;

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
      var bono = Math.round((d.bono || 0) * 100) / 100;
      var estip = Math.round((d.estipendio || 0) * 100) / 100;
      var org  = Math.round((base + iva + bono) * 100) / 100;  // estipendio NO entra al total org
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
    var bono2 = Math.round((d.bono || 0) * 100) / 100;
    var estip2 = Math.round((d.estipendio || 0) * 100) / 100;
    gBase += base; gIVA += iva; gOrg += Math.round((base+iva+bono2)*100)/100; // estipendio NO en total
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

// Wrapper sin argumentos para trigger de tiempo (triggers no pueden pasar params)
function _autoImportarEstipendio() {
  try { importarEstipendioDesdeKobo(true); } catch(e) { Logger.log("Error auto-estipendio: " + e); }
}

function configurarTriggers() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Eliminar triggers manejados por este sistema
  ScriptApp.getProjectTriggers().forEach(function(t){
    var h = t.getHandlerFunction();
    if (h === "importarDesdeKobo" || h === "_autoImportarKobo" ||
        h === "_autoImportarEstipendio" || h === "importarAlAbrir" ||
        h === "actualizarQuincenaActual" || h === "actualizarDashboardVisual" ||
        h === "onEditInstalable") {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Al abrir: importa Kobo + estipendio silencioso, actualiza reportes
  ScriptApp.newTrigger("importarAlAbrir").forSpreadsheet(ss).onOpen().create();
  // Kobo asistencia: auto-importar cada 6 horas silencioso (dedup previene duplicados)
  ScriptApp.newTrigger("_autoImportarKobo").timeBased().everyHours(6).create();
  // Estipendio: auto-importar cada 6 horas silencioso
  ScriptApp.newTrigger("_autoImportarEstipendio").timeBased().everyHours(6).create();
  // Actualizar reporte de quincena activa cada día a las 7am
  ScriptApp.newTrigger("actualizarQuincenaActual")
    .timeBased().everyDays(1).atHour(7).create();
  // Actualizar Dashboard Visual cada 10 minutos
  ScriptApp.newTrigger("actualizarDashboardVisual")
    .timeBased().everyMinutes(10).create();
  // onEdit instalable para email + archivo de pagos (requiere autorización)
  ScriptApp.newTrigger("onEditInstalable")
    .forSpreadsheet(ss).onEdit().create();

  _alert("✅ Automatizaciones activadas:\n\n" +
    "• 📥 Kobo asistencia: auto cada 6h + al abrir (silencioso, sin duplicados)\n" +
    "• 💵 Estipendio: auto cada 6h + al abrir (silencioso)\n" +
    "• 📅 Quincena activa: se actualiza cada día a las 7am\n" +
    "• 📊 Dashboard Visual: se actualiza cada 10 minutos\n\n" +
    "onEdit (automático):\n" +
    "• Categoría → auto-llena Tarifa\n" +
    "• Pagado → actualiza Dashboard");
}); }

// ── Sincronizar IDs desde DatosKobo ───────────────────────────

/**
 * Lee columna Participante de DatosKobo.
 * Para cada entrada con formato "Nombre (CODIGO)", extrae el código y
 * actualiza col A de PARTICIPANTES si ese participante aún no tiene ID.
 * Para entradas sin código, busca en Creamos DB por nombre.
 */
function sincronizarIDsDesdeKobo() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hK = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  var hP = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES);
  if (!hK || !hP) { _alert("Faltan hojas DatosKobo o PARTICIPANTES."); return; }

  // Detectar columna Participante en Kobo
  var enc  = hK.getRange(1, 1, 1, hK.getLastColumn()).getValues()[0];
  var cols = detectarColumnas(enc, hK.getRange(2, 1, Math.min(50, hK.getLastRow()-1), hK.getLastColumn()).getValues());
  var iParticipante = (cols.participante !== undefined) ? cols.participante : -1;
  if (iParticipante < 0) { _alert("No se detectó columna Participante en DatosKobo."); return; }

  // Recopilar nombres únicos de Kobo
  var rawNombres = hK.getRange(2, iParticipante + 1, hK.getLastRow() - 1, 1).getValues();
  var uniqueRaw  = {};
  rawNombres.forEach(function(r) {
    var v = String(r[0] || "").trim();
    if (v) uniqueRaw[v] = true;
  });

  // Leer PARTICIPANTES: construir mapa por nombre normalizado → {fila, id}
  var datP    = hP.getRange(2, 1, hP.getLastRow() - 1, 2).getValues();
  var mapaP   = {}; // norm → {fila (1-based), id}
  datP.forEach(function(r, i) {
    var nom = String(r[1] || "").trim();
    if (!nom) return;
    mapaP[textoParaComparar(nom)] = { fila: i + 2, id: String(r[0] || "").trim() };
  });

  var actualizados = 0, sinEncontrar = [];

  Object.keys(uniqueRaw).forEach(function(raw) {
    var codigo    = extraerCodigo(raw);
    var nombreLimpio = limpiarNombre(raw);
    var normNom   = textoParaComparar(nombreLimpio);

    // Buscar fila en PARTICIPANTES por nombre normalizado
    var entrada = mapaP[normNom];
    if (!entrada) {
      // Búsqueda parcial: 2+ palabras largas
      var palabras = normNom.split(/\s+/).filter(function(p){ return p.length >= 4; });
      if (palabras.length >= 2) {
        var keys = Object.keys(mapaP);
        for (var ki = 0; ki < keys.length; ki++) {
          var hits = palabras.filter(function(p){ return keys[ki].indexOf(p) !== -1; }).length;
          if (hits >= 2) { entrada = mapaP[keys[ki]]; break; }
        }
      }
    }

    if (!entrada) { sinEncontrar.push(nombreLimpio + (codigo ? " [" + codigo + "]" : "")); return; }
    if (_esCreamos_ID_real(entrada.id)) return; // ya tiene ID, no sobreescribir

    // Tiene código de Kobo → usarlo
    if (codigo) {
      hP.getRange(entrada.fila, 1).setValue(codigo);
      entrada.id = codigo;
      actualizados++;
      return;
    }

    // Sin código → buscar en Creamos DB
    var db = _buscarEnCreamos_DB(nombreLimpio);
    if (db.id) {
      hP.getRange(entrada.fila, 1).setValue(db.id);
      entrada.id = db.id;
      actualizados++;
    } else {
      sinEncontrar.push(nombreLimpio + " (sin código en Kobo, no encontrado en DB)");
    }
  });

  var msg = "✅ IDs actualizados desde DatosKobo: " + actualizados;
  if (sinEncontrar.length > 0)
    msg += "\n\n⚠️ Sin resolver (" + sinEncontrar.length + "):\n" + sinEncontrar.join("\n");
  _alert(msg);
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

  var datos = hP.getRange(2, 1, hP.getLastRow() - 1, 23).getValues();
  var actualizados = [], sinEncontrar = [];
  var mapaNombreAID = {};

  datos.forEach(function(fila, i) {
    var idActual    = String(fila[0] || "").trim();
    var nombre      = String(fila[1] || "").trim();
    var fechaNacAct = String(fila[2] || "").trim(); // col C = Fecha_Nacimiento
    var edadAct     = String(fila[3] || "").trim(); // col D = Edad
    var generoAct   = String(fila[4] || "").trim(); // col E = Genero
    var anioAct     = String(fila[5] || "").trim(); // col F = Ano_Entrada_Creamos
    if (!nombre) return;

    // Si ya tiene ID real Y todos los campos personales están llenos → solo registrar en mapa
    var yaCompleto = _esCreamos_ID_real(idActual) &&
                     fechaNacAct && edadAct && generoAct && anioAct;
    if (yaCompleto) {
      mapaNombreAID[nombre] = idActual;
      return;
    }

    ss.toast("Buscando: " + nombre, "🔍", -1);
    var db = _buscarEnCreamos_DB(nombre);
    // Si no encontró por nombre pero tiene ID real → buscar por ID como fallback
    if ((db.noEncontrado || !db.id) && _esCreamos_ID_real(idActual)) {
      db = _buscarEnCreamos_DB_porID(idActual);
    }
    if (db.noEncontrado || !db.id) {
      if (_esCreamos_ID_real(idActual)) mapaNombreAID[nombre] = idActual;
      else sinEncontrar.push(nombre);
      return;
    }

    var fueActualizado = false;

    // Actualizar ID solo si no tenía uno real
    if (!_esCreamos_ID_real(idActual)) {
      hP.getRange(i + 2, 1).setValue(db.id).setBackground(null).setFontColor(null).setFontStyle("normal");
      fueActualizado = true;
    }

    // Llenar cols C-F solo si están vacías
    if (!fechaNacAct && db.fechaNac)    { hP.getRange(i + 2, 3).setValue(db.fechaNac);    fueActualizado = true; }
    if (!edadAct     && db.edad)        { hP.getRange(i + 2, 4).setValue(db.edad);         fueActualizado = true; }
    if (!generoAct   && db.genero)      { hP.getRange(i + 2, 5).setValue(db.genero);       fueActualizado = true; }
    if (!anioAct     && db.anioEntrada) { hP.getRange(i + 2, 6).setValue(db.anioEntrada);  fueActualizado = true; }
    // DPI: llenar si vacío
    var dpiAct = String(fila[15] || "").trim(); // col P = DPI
    if (!dpiAct && db.dpi)              { hP.getRange(i + 2, 16).setValue(db.dpi);         fueActualizado = true; }

    mapaNombreAID[nombre] = db.id || idActual;
    if (fueActualizado) actualizados.push(nombre + " → " + (db.id || idActual));
  });

  ss.toast("", "", 1);
  _colorearParticipantes(hP, datos.length);

  // Propagar IDs recién actualizados a ListaTerapias e InclusionLaboral
  _propagarIDsAHojas(ss, mapaNombreAID);

  _alert(
    "🔄 Sincronización completada\n\n" +
    (actualizados.length
      ? "✅ Actualizados (" + actualizados.length + "):\n  " + actualizados.join("\n  ") + "\n\n"
      : "") +
    (sinEncontrar.length
      ? "⚠️ Aún sin perfil (" + sinEncontrar.length + "):\n  " + sinEncontrar.join("\n  ")
      : "Todos los participantes tienen Creamos ID ✅")
  );
  try { actualizarDashboardVisual(); } catch(_) {}
}); }

function _propagarIDsAHojas(ss, mapaNombreAID) {
  ["ListaTerapias", "InclusionLaboral", "HijosCCI"].forEach(function(nm) {
    var h = ss.getSheetByName(nm);
    if (!h || h.getLastRow() < 2) return;
    var datos = h.getRange(2, 1, h.getLastRow() - 1, 2).getValues();
    datos.forEach(function(f, i) {
      var nombre = String(f[1] || "").trim();
      if (nombre && mapaNombreAID[nombre]) {
        h.getRange(i + 2, 1).setValue(mapaNombreAID[nombre]);
      }
    });
  });
}

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
  var tz   = CFG.TIMEZONE;
  var ahora = new Date();
  var anio  = ahora.getFullYear();
  var ts    = Utilities.formatDate(ahora, tz, "dd/MM/yyyy HH:mm");

  // ── Calcular KPIs ─────────────────────────────────────────────

  // KPI 1: Ciclos de Vida completados en el año en curso (fuente: hoja CiclosVida)
  var kpi1 = 0;
  var hCV = ss.getSheetByName("CiclosVida");
  if (hCV && hCV.getLastRow() > 1) {
    var datCV = hCV.getRange(2, 1, hCV.getLastRow() - 1, 1).getValues();
    datCV.forEach(function(r) {
      var f = r[0];
      var d = (f instanceof Date) ? f : new Date(f);
      if (!isNaN(d) && d.getFullYear() === anio) kpi1++;
    });
  }

  // KPI 2: Monto promedio ingresos brutos mensual (año en curso)
  // = total_base_año / número_de_meses_con_datos
  var kpi2 = 0;
  var hojaF = ss.getSheetByName(CFG.HOJAS.FACTURACION);
  if (hojaF && hojaF.getLastRow() > 1) {
    var datF = hojaF.getDataRange().getValues().slice(1);
    var baseParMes = {}; // mes_index → total base
    datF.forEach(function(f) {
      if (Number(f[3]) !== anio) return;
      var mesIdx = CFG.MESES.indexOf(String(f[2]).trim());
      if (mesIdx < 0) return;
      var base = parseFloat(f[9]) || 0; // col 9 = Base/Monto_Base
      var hrs  = parseFloat(f[5]) || 0; // col 5 = HorasTrab
      if (hrs <= 0) return; // solo meses con horas
      baseParMes[mesIdx] = (baseParMes[mesIdx] || 0) + base;
    });
    var mesesConDatos = Object.keys(baseParMes).length;
    if (mesesConDatos > 0) {
      var totalBase = Object.keys(baseParMes).reduce(function(s, k) { return s + baseParMes[k]; }, 0);
      kpi2 = totalBase / mesesConDatos;
    }
  }

  // KPI 3: Horas totales de formación (año en curso) desde DatosKobo
  // Deduplica por participante+día; cuenta esDiaDeEstudio OR esTerapia
  var kpi3 = 0;
  var hK = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (hK && hK.getLastRow() > 1) {
    try {
      var encK = hK.getRange(1, 1, 1, hK.getLastColumn()).getValues()[0];
      var rawK = hK.getRange(2, 1, hK.getLastRow() - 1, hK.getLastColumn()).getValues();
      var colsK = detectarColumnas(encK, rawK.slice(0, 50));
      var diasEstMap  = obtenerDiasEstudio();
      var terapiasMap = obtenerListaTerapias();
      var mapeoN = cargarMapeoNombres();
      var diasFormacion = {}; // "nombre|yyyy-mm-dd" → true
      rawK.forEach(function(fila) {
        var ts2 = _resolverTsKobo(fila, colsK);
        if (!ts2 || isNaN(ts2) || ts2.getFullYear() !== anio) return;
        var nombreRaw = obtenerParticipanteFila(fila, colsK);
        if (!nombreRaw) return;
        var nombre = normalizarNombre(nombreRaw, mapeoN) || limpiarNombre(nombreRaw);
        if (!nombre) return;
        var esDE  = esDiaDeEstudio(nombre, ts2, diasEstMap);
        var esTer = !!terapiasMap[nombre];
        if (!esDE && !esTer) return;
        var claveD = nombre + "|" + _dClave(ts2);
        diasFormacion[claveD] = true;
      });
      kpi3 = Object.keys(diasFormacion).length * CFG.HORAS_JORNADA_NORMAL;
    } catch(e) { Logger.log("KPI3 error: " + e.message); }
  }

  // KPI 4: Promedio horas laborales mensuales (año en curso) desde FACTURACION
  // = total_HorasTrab_año / conteo_de_participante_meses_con_horas
  var kpi4 = 0;
  if (hojaF && hojaF.getLastRow() > 1) {
    var totalHrsAnio = 0;
    var conteoPartMeses = 0;
    hojaF.getDataRange().getValues().slice(1).forEach(function(f) {
      if (Number(f[3]) !== anio) return;
      var hrs = parseFloat(f[5]) || 0;
      if (hrs > 0) { totalHrsAnio += hrs; conteoPartMeses++; }
    });
    if (conteoPartMeses > 0) kpi4 = totalHrsAnio / conteoPartMeses;
  }

  // ── Layout: 2 columnas, cada KPI = label + value ──────────────
  // Columnas: A=label, B=valor
  dash.clearContents();
  dash.clearFormats();
  if (dash.getMaxColumns() < 2) dash.insertColumnsAfter(dash.getMaxColumns(), 2 - dash.getMaxColumns());

  var filas = [], tipos = [];
  function push(f, t) { filas.push(f); tipos.push(t); }

  // Título
  push(["DASHBOARD — mi eelo", ""], "titulo");
  push(["", ""], "vacio");

  // KPI 1 — Ciclos de Vida completados (año en curso)
  push(["Ciclos de Vida completados (año " + anio + ")", ""], "label1");
  push([kpi1 > 0 ? String(kpi1) : "0", ""], "val1");
  push(["", ""], "vacio");

  // KPI 2 — Monto promedio ingresos brutos mensual (año en curso)
  push(["Monto promedio ingresos brutos mensual (año " + anio + ")", ""], "label2");
  push([kpi2 > 0 ? "Q " + kpi2.toFixed(2) : "—", ""], "val2");
  push(["", ""], "vacio");

  // KPI 3 — Horas totales de formación (año en curso)
  push(["Horas totales de formación (año " + anio + ")", ""], "label3");
  push([kpi3 > 0 ? kpi3.toFixed(0) + " hrs" : "0 hrs", ""], "val3");
  push(["", ""], "vacio");

  // KPI 4 — Promedio horas laborales mensuales (año en curso)
  push(["Promedio horas laborales mensuales (año " + anio + ")", ""], "label4");
  push([kpi4 > 0 ? kpi4.toFixed(1) + " hrs" : "—", ""], "val4");
  push(["", ""], "vacio");

  // Timestamp
  push(["Actualizado: " + ts, ""], "ts");

  dash.getRange(1, 1, filas.length, 2).setValues(filas);

  // ── Formato ───────────────────────────────────────────────────
  tipos.forEach(function(tipo, idx) {
    var r = dash.getRange(idx + 1, 1, 1, 2);
    r.setFontFamily("Arial").setFontSize(11).setVerticalAlignment("middle");

    if (tipo === "titulo") {
      r.merge().setFontSize(16).setFontWeight("bold")
       .setBackground("#1a237e").setFontColor("#ffffff")
       .setHorizontalAlignment("center");
      dash.setRowHeight(idx + 1, 44);
    } else if (tipo === "vacio") {
      r.setBackground("#ffffff");
      dash.setRowHeight(idx + 1, 12);
    } else if (tipo === "label1") {
      r.merge().setFontWeight("bold").setFontSize(12)
       .setBackground("#1a237e").setFontColor("#ffffff")
       .setHorizontalAlignment("left");
      dash.setRowHeight(idx + 1, 32);
    } else if (tipo === "val1") {
      r.merge().setFontSize(28).setFontWeight("bold")
       .setBackground("#e8eaf6").setFontColor("#1a237e")
       .setHorizontalAlignment("center");
      dash.setRowHeight(idx + 1, 56);
    } else if (tipo === "label2") {
      r.merge().setFontWeight("bold").setFontSize(12)
       .setBackground("#1b5e20").setFontColor("#ffffff")
       .setHorizontalAlignment("left");
      dash.setRowHeight(idx + 1, 32);
    } else if (tipo === "val2") {
      r.merge().setFontSize(28).setFontWeight("bold")
       .setBackground("#e8f5e9").setFontColor("#1b5e20")
       .setHorizontalAlignment("center");
      dash.setRowHeight(idx + 1, 56);
    } else if (tipo === "label3") {
      r.merge().setFontWeight("bold").setFontSize(12)
       .setBackground("#e65100").setFontColor("#ffffff")
       .setHorizontalAlignment("left");
      dash.setRowHeight(idx + 1, 32);
    } else if (tipo === "val3") {
      r.merge().setFontSize(28).setFontWeight("bold")
       .setBackground("#fff3e0").setFontColor("#e65100")
       .setHorizontalAlignment("center");
      dash.setRowHeight(idx + 1, 56);
    } else if (tipo === "label4") {
      r.merge().setFontWeight("bold").setFontSize(12)
       .setBackground("#4a0072").setFontColor("#ffffff")
       .setHorizontalAlignment("left");
      dash.setRowHeight(idx + 1, 32);
    } else if (tipo === "val4") {
      r.merge().setFontSize(28).setFontWeight("bold")
       .setBackground("#f3e5f5").setFontColor("#4a0072")
       .setHorizontalAlignment("center");
      dash.setRowHeight(idx + 1, 56);
    } else if (tipo === "ts") {
      r.merge().setFontSize(9).setFontColor("#9e9e9e")
       .setBackground("#fafafa").setHorizontalAlignment("right")
       .setFontStyle("italic");
      dash.setRowHeight(idx + 1, 24);
    }
  });

  // Anchos de columna
  dash.setColumnWidth(1, 380);
  dash.setColumnWidth(2, 20);
  dash.setFrozenRows(1);

}); }

/*
 * Calcula todos los KPIs del mes en curso (legacy — mantenido para compatibilidad).
 * Retorna objeto con valores básicos de resumen.
 */
function _calcularKpis(ss) {
  var ahora    = new Date();
  var mesNum   = ahora.getMonth()+1;
  var anio     = ahora.getFullYear();
  var nombreMes = CFG.MESES[mesNum-1];
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();

  var kpi = {
    mes: nombreMes, anio: anio,
    ciclosCompletados: 0,
    ingresoPromedio: 0,
    horasFormacion: 0,
    promedioHorasLaborales: 0,
    activos: 0, sumBase: 0, sumIVA: 0, sumTot: 0,
    pagadas: 0, pendientes: 0, totalHorasMes: 0
  };

  var hojaP = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES);
  if (hojaP && hojaP.getLastRow() > 1) {
    hojaP.getDataRange().getValues().slice(1).forEach(function(r) {
      var etapa = String(r[4]).trim().toLowerCase();
      if (etapa !== "retiradx") kpi.activos++;
    });
  }

  var hojaF = ss.getSheetByName(CFG.HOJAS.FACTURACION);
  if (hojaF && hojaF.getLastRow() > 1) {
    var participantesMes = {};
    hojaF.getDataRange().getValues().slice(1).forEach(function(f) {
      if (f[2] !== nombreMes || Number(f[3]) !== anio) return;
      kpi.sumBase += parseFloat(f[9])  || 0;
      kpi.sumIVA  += parseFloat(f[11]) || 0;
      kpi.sumTot  += parseFloat(f[12]) || 0;
      if (String(f[17]) === "Sí") kpi.pagadas++; else kpi.pendientes++;
      var nombre = String(f[1]).trim();
      if (!participantesMes[nombre]) participantesMes[nombre] = 0;
      participantesMes[nombre] += parseFloat(f[5]) || 0;
      kpi.totalHorasMes += parseFloat(f[5]) || 0;
    });
    var numPart = Object.keys(participantesMes).length;
    if (numPart > 0) {
      var hVals = Object.keys(participantesMes).map(function(n){ return participantesMes[n]; });
      kpi.promedioHorasLaborales = hVals.reduce(function(s,v){ return s+v; },0) / hVals.length;
    }
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
  var tz              = CFG.TIMEZONE;

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
    if (!emp || _esValorAccion(emp)) continue; // ignorar entradas/salidas como nombre
    if (filtroParticipante && emp.toLowerCase() !== filtroParticipante.toLowerCase()) continue;

    var tipoReg = obtenerTipoRegistro(fila, cols);
    if (!tipoReg.esIngreso && !tipoReg.esEgreso) continue;

    var ts = _resolverTsKobo(fila, cols);
    if (!ts || isNaN(ts)) continue;

    var claveReg = emp + "|" + ts.getTime() + "|" + (tipoReg.esIngreso ? "E" : "S");
    if (regVistos[claveReg]) continue; regVistos[claveReg] = true;

    if (!_validarEnRango(tipo, ts, fechaInicio, fechaFin)) continue;

    if (!regPorEmp[emp]) regPorEmp[emp] = [];
    // fechaEnd: raw "end" column for EGRESO — used to recover real departure when "chained forms"
    // produce horas=0 (SALIDA opened at exact moment cross-day ENTRADA was submitted).
    var fechaEndRaw = (tipoReg.esEgreso && cols.end !== undefined) ? new Date(fila[cols.end]) : null;
    regPorEmp[emp].push({
      fecha: ts,
      fechaEnd: fechaEndRaw,
      esIngreso:    tipoReg.esIngreso,
      esEgreso:     tipoReg.esEgreso,
      esTerapia:    tipoReg.esTerapia,
      esPermiso:    tipoReg.esPermiso,
      esComputacion:tipoReg.esComputacion
    });
  }

  // Deduplicar por participante+día: un Kobo submit doble genera 2 filas iguales.
  // Conservar ingreso más temprano y egreso más tardío por día.
  Object.keys(regPorEmp).forEach(function(emp) {
    var byDia = {};
    regPorEmp[emp].forEach(function(r) {
      var k = r.fecha.getFullYear() + "-" + r.fecha.getMonth() + "-" + r.fecha.getDate();
      if (!byDia[k]) byDia[k] = { ing: [], egr: [] };
      if (r.esIngreso) byDia[k].ing.push(r);
      else if (r.esEgreso) byDia[k].egr.push(r);
    });
    var clean = [];
    Object.keys(byDia).forEach(function(k) {
      var d = byDia[k];
      if (d.ing.length) { d.ing.sort(function(a,b){return a.fecha-b.fecha;}); clean.push(d.ing[0]); }
      if (d.egr.length) { d.egr.sort(function(a,b){return b.fecha-a.fecha;}); clean.push(d.egr[0]); }
    });
    regPorEmp[emp] = clean;
  });

  var listaEmps = Object.keys(regPorEmp).sort();
  if (listaEmps.length === 0) {
    _alert("No hay datos para el período seleccionado.\n\n" +
      "Verifica que:\n• Hayas importado datos de Kobo\n• El período tenga registros\n• El nombre del filtro sea exacto");
    return;
  }

  // Crear / limpiar pestaña
  var hoja;
  if (nuevaHoja === false || typeof nuevaHoja === "string") {
    var sheetNm = (typeof nuevaHoja === "string") ? nuevaHoja : "Reporte_Fijo";
    var nh = ss.getSheetByName(sheetNm);
    hoja = nh || ss.insertSheet(sheetNm);
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

      // Resetear al cambiar de día — entrada sin salida NO se cuenta, se descarta
      if (fechaAnt && !_esMismaFecha(fechaAnt, reg.fecha)) {
        currentIngreso = null; // entrada sola del día anterior → descartada
        lastEgreso = null;
      }
      fechaAnt = reg.fecha;

      if (reg.esIngreso) {
        if (currentIngreso && _esMismaFecha(currentIngreso.fecha, reg.fecha)) continue;
        currentIngreso = reg; lastEgreso = null;
      } else if (reg.esEgreso && currentIngreso) {
        var horas  = (reg.fecha - currentIngreso.fecha) / 3600000;

        if (horas <= 0) {
          // "Chained form": SALIDA abierta en el instante exacto que ENTRADA cross-day resolvió.
          // Intentar con reg.fechaEnd (hora real de envío del formulario).
          var isDiaEstInv = esDiaDeEstudio(empId, currentIngreso.fecha, diasEstudioMapa);
          var salidaFin = reg.fechaEnd;
          var horasReal = (salidaFin && !isNaN(salidaFin)) ? (salidaFin - currentIngreso.fecha) / 3600000 : 0;
          if (horasReal >= 0.5 && horasReal < 14) {
            var tipoLblR = isDiaEstInv ? "Día de Estudio" : "Normal";
            var porcR    = isDiaEstInv ? 0 : 100;
            _addFila(currentIngreso.fecha, currentIngreso.fecha, salidaFin, tipoLblR, horasReal, porcR);
          }
          // Si horasReal < 0.5 o incoherente → sesión descartada (no se estima)
          currentIngreso = null; lastEgreso = reg;
        } else if (horas < 0.5) {
          // Sesión menor a 30 minutos → probablemente error de registro, se descarta
          currentIngreso = null; lastEgreso = reg;
        } else {
          // Salida normal (horas >= 0.5). Verificar si el form fue enviado mucho después
          // de que fue abierto — en ese caso usar fechaEnd como hora real de salida.
          var salidaR = reg.fecha;
          if (reg.fechaEnd && !isNaN(reg.fechaEnd)) {
            var formDelay = (reg.fechaEnd - reg.fecha) / 3600000; // cuánto tardó en enviar
            var horasConEnd = (reg.fechaEnd - currentIngreso.fecha) / 3600000;
            // Si el form tardó >1h en enviarse Y la sesión resultante es razonable: usar fechaEnd
            if (formDelay > 1 && horasConEnd > horas && horasConEnd <= 16) {
              salidaR = reg.fechaEnd;
              horas   = horasConEnd;
            }
          }
          var isDiaEst2 = esDiaDeEstudio(empId, currentIngreso.fecha, diasEstudioMapa);
          var tipoLbl, porc;
          if      (isDiaEst2)      { tipoLbl = "Día de Estudio"; porc = 0; }
          else if (reg.esPermiso)  { tipoLbl = "Permiso"; porc = 0; }
          else if (reg.esTerapia || listaTerapias[empId]) { tipoLbl = "Terapia"; porc = 100; }
          else if (reg.esComputacion) { tipoLbl = "Computación"; porc = 50; }
          else                     { tipoLbl = "Normal"; porc = 100; }
          _addFila(currentIngreso.fecha, currentIngreso.fecha, salidaR, tipoLbl, horas, porc);
          currentIngreso = null; lastEgreso = reg;
        }
      }
    }

    // Ingreso sin salida al final del set → descartado, no se estima
    // (currentIngreso abierto = entrada sin salida registrada, no cuenta)

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
  // Forzar formato numérico en celdas de horas y monto (evita que GAS aplique formato "Time")
  hoja.getRange(filaActual,  4).setNumberFormat("0.00");   // Horas Lab.
  hoja.getRange(filaActual+1,2).setNumberFormat("0.00");   // Total Terapias
  hoja.getRange(filaActual+1,4).setNumberFormat("0.00");   // Horas a Pagar
  hoja.getRange(filaActual+2,4).setNumberFormat('"Q "#,##0.00'); // Monto
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
  var tz = CFG.TIMEZONE;
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
  var tz = CFG.TIMEZONE;
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
  var map = {};       // clave: nombre original
  var normMap = {};   // clave: nombre normalizado → nombre original en map
  var hojaP = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.HOJAS.PARTICIPANTES);
  if (!hojaP) return map;
  var datos = hojaP.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var nombre = String(datos[i][1]||"").trim();
    if (!nombre) continue;

    var norm = textoParaComparar(nombre);
    // Si ya existe una entrada con el mismo nombre normalizado, conservar la que tenga ID real
    if (normMap[norm]) {
      var existing = map[normMap[norm]];
      var newId = String(datos[i][0]||"").trim();
      if (!_esCreamos_ID_real(existing.id) && _esCreamos_ID_real(newId)) {
        // La nueva fila tiene mejor ID — reemplazar
        delete map[normMap[norm]];
        normMap[norm] = nombre;
      } else {
        continue; // conservar la existente
      }
    } else {
      normMap[norm] = nombre;
    }

    var tarifa = parseFloat(datos[i][13]); // col N = Tarifa_Hora (idx 13)
    if (isNaN(tarifa) || tarifa <= 0) {
      var cat = String(datos[i][12]).trim().toUpperCase(); // col M = Categoria (idx 12)
      tarifa = CFG.CATEGORIAS[cat] || CFG.CATEGORIAS.C;
    }
    var t = String(datos[i][14]).trim().toLowerCase(); // col O = Tiene_Factura (idx 14)
    map[nombre] = {
      id:           String(datos[i][0]||"").trim(),
      tarifa:       tarifa,
      categoria:    String(datos[i][12]).trim().toUpperCase(),
      tieneFactura: t === "sí" || t === "si",
      estipendio:   0
    };
  }
  return map;
}

// ══════════════════════════════════════════════════════════════════
// MIGRACIÓN — actualiza instalación existente sin borrar datos
// ══════════════════════════════════════════════════════════════════

/**
 * Recalcula Tarifa_Hora (col J) desde Categoria (col I) para todos los participantes.
 * Si Tiene_Factura (col K) está vacío → pone "Sí" por defecto.
 * Seguro correr en cualquier momento; no borra ningún otro dato.
 */
function recalcularTarifas() { _run(function() {
  var hP = _sh(CFG.HOJAS.PARTICIPANTES);
  if (!hP || hP.getLastRow() < 2) { _alert("Sin datos en PARTICIPANTES."); return; }

  var nRows = hP.getLastRow() - 1;
  var datos = hP.getRange(2, 1, nRows, 15).getValues();

  var tarifasNueva  = [];
  var facturaNueva  = [];
  var nTarifa = 0, nFact = 0, sinCat = [];

  datos.forEach(function(r, i) {
    var nombre = String(r[1] || "").trim();
    var cat    = String(r[12] || "").trim().toUpperCase(); // col M = Categoria (idx 12)
    var tarifa = CFG.CATEGORIAS[cat];

    if (tarifa) {
      tarifasNueva.push([tarifa]);
      nTarifa++;
    } else {
      tarifasNueva.push([r[13] || ""]); // conservar existente si no hay cat válida
      if (nombre) sinCat.push(nombre);
    }

    var fact = String(r[14] || "").trim(); // col O = Tiene_Factura (idx 14)
    if (!fact || (fact !== "Sí" && fact !== "No")) {
      facturaNueva.push(["Sí"]);
      nFact++;
    } else {
      facturaNueva.push([fact]);
    }
  });

  hP.getRange(2, 14, nRows, 1).setValues(tarifasNueva); // col N = Tarifa_Hora
  hP.getRange(2, 15, nRows, 1).setValues(facturaNueva); // col O = Tiene_Factura

  var msg = "✅ Tarifas y factura actualizadas\n\n" +
    "• " + nTarifa + " tarifas calculadas desde Categoria (col M)\n" +
    "• " + nFact + " Tiene_Factura vacíos → rellenados con 'Sí'";

  if (sinCat.length > 0) {
    msg += "\n\n⚠️ Participantes SIN categoría asignada (" + sinCat.length + "):\n" +
      sinCat.slice(0, 10).join("\n") +
      (sinCat.length > 10 ? "\n... y " + (sinCat.length - 10) + " más" : "") +
      "\n\nAsigna A / B / C / D en col I para que sus tarifas se calculen correctamente.";
  }

  _alert(msg);
}); }

/**
 * Actualiza la estructura del sistema a la versión más reciente.
 * NO borra datos existentes. Seguro correr sobre una instalación ya en uso.
 *
 * Qué hace:
 *  1. Si col I tiene "Categoria" (esquema viejo): reordena columnas moviendo
 *     Hijos_CCI y Num_Hijos_CCI desde U/V a I/J (antes de Categoria).
 *  2. Si col I ya tiene "Hijos_CCI" (esquema nuevo): sin cambios estructurales.
 *  3. Crea hoja "Bonos" si no existe
 *  4. Crea hoja "HijosCCI" si no existe
 *  5. Aplica validaciones y formatos en PARTICIPANTES
 *  6. Actualiza colores e IDs en todas las hojas auxiliares
 */

// ── Reparar dropdowns PARTICIPANTES ──────────────────────────────
function repararDropdownsParticipantes() { _run(function() {
  var hP = _sh(CFG.HOJAS.PARTICIPANTES);
  if (hP.getLastRow() < 1) { _alert("PARTICIPANTES está vacía."); return; }

  // Detectar esquema por col C
  var colCHeader = String(hP.getRange(1, 3, 1, 1).getValue() || "").trim();
  var es23 = (colCHeader === "Fecha_Nacimiento");
  var es19 = (colCHeader === "Proyecto");

  if (!es23 && !es19) {
    _alert("⚠️ No se reconoce el esquema (col C = '" + colCHeader + "').\n" +
           "Esperado: 'Fecha_Nacimiento' (23 cols) o 'Proyecto' (19 cols).\n" +
           "Ejecuta primero ⬆️ Migrar sistema.");
    return;
  }

  // Limpiar validaciones existentes en toda la zona de datos
  hP.getRange("A2:Z500").clearDataValidations();

  var vEtapa  = SpreadsheetApp.newDataValidation().requireValueInList(["Inscritx","Retiradx","Empleadx","Ciclo de Vida Terminado"],true).build();
  var vCat    = SpreadsheetApp.newDataValidation().requireValueInList(["A","B","C","D"],true).build();
  var vSiNo   = SpreadsheetApp.newDataValidation().requireValueInList(["Sí","No"],true).build();
  var vBanco  = SpreadsheetApp.newDataValidation().requireValueInList(["Banrural","Industrial","BAC Credomatic","G&T Continental","Banco Azteca","N/A","Otro"],true).build();
  var vTipoCt = SpreadsheetApp.newDataValidation().requireValueInList(["Monetaria","Ahorro",""],true).build();
  var vPago   = SpreadsheetApp.newDataValidation().requireValueInList(["Transferencia","Cheque"],true).build();

  if (es23) {
    // 23 cols A–W
    hP.getRange("I2:I500").setDataValidation(vEtapa);   // col I = Etapa
    hP.getRange("M2:M500").setDataValidation(vCat);     // col M = Categoria
    hP.getRange("O2:O500").setDataValidation(vSiNo);    // col O = Tiene_Factura
    hP.getRange("S2:S500").setDataValidation(vBanco);   // col S = Banco
    hP.getRange("T2:T500").setDataValidation(vTipoCt);  // col T = Tipo_Cuenta
    hP.getRange("V2:V500").setDataValidation(vPago);    // col V = Forma_Pago
    hP.getRange("N2:N500").setNumberFormat("Q#,##0.00"); // col N = Tarifa_Hora
    hP.getRange("C2:C500").setNumberFormat("dd/MM/yyyy"); // col C = Fecha_Nacimiento
    hP.getRange("P2:P500").setNumberFormat("@");          // col P = DPI (texto plano)
    hP.getRange("Q2:Q500").setNumberFormat("@");          // col Q = NIT (texto plano)
    _alert("✅ Dropdowns reparados — esquema 23 cols (A–W):\n" +
           "• I = Etapa\n• M = Categoría (A/B/C/D)\n• O = Tiene_Factura\n" +
           "• S = Banco\n• T = Tipo_Cuenta\n• V = Forma_Pago\n• P = DPI (texto)");
  } else {
    // 19 cols A–S
    hP.getRange("E2:E500").setDataValidation(vEtapa);   // col E = Etapa
    hP.getRange("I2:I500").setDataValidation(vCat);     // col I = Categoria
    hP.getRange("K2:K500").setDataValidation(vSiNo);    // col K = Tiene_Factura
    hP.getRange("O2:O500").setDataValidation(vBanco);   // col O = Banco
    hP.getRange("P2:P500").setDataValidation(vTipoCt);  // col P = Tipo_Cuenta
    hP.getRange("R2:R500").setDataValidation(vPago);    // col R = Forma_Pago
    hP.getRange("J2:J500").setNumberFormat("Q#,##0.00"); // col J = Tarifa_Hora
    hP.getRange("L2:L500").setNumberFormat("@");          // col L = DPI (texto plano)
    hP.getRange("M2:M500").setNumberFormat("@");          // col M = NIT (texto plano)
    _alert("✅ Dropdowns reparados — esquema 19 cols (A–S):\n" +
           "• E = Etapa\n• I = Categoría (A/B/C/D)\n• K = Tiene_Factura\n" +
           "• O = Banco\n• P = Tipo_Cuenta\n• R = Forma_Pago\n\n" +
           "ℹ️ Para actualizar a 23 cols usa ⬆️ Migrar sistema.");
  }
}); }

function repararFormatoDPI() { _run(function() {
  var hP = _sh(CFG.HOJAS.PARTICIPANTES);
  if (hP.getLastRow() < 2) { _alert("PARTICIPANTES vacía."); return; }

  var colCHeader = String(hP.getRange(1, 3, 1, 1).getValue() || "").trim();
  var es23 = (colCHeader === "Fecha_Nacimiento");

  // En 23-col: P=DPI(16), Q=NIT(17). En 19-col: L=DPI(12), M=NIT(13)
  var colDPI = es23 ? 16 : 12;
  var colNIT = es23 ? 17 : 13;
  var nRows  = hP.getLastRow() - 1;

  function limpiarCol(colNum) {
    var rango = hP.getRange(2, colNum, nRows, 1);
    rango.setNumberFormat("@"); // texto plano primero
    var vals = rango.getValues();
    var nuevos = vals.map(function(r) {
      var v = r[0];
      if (v === "" || v === null || v === undefined) return [v];
      // Si es número, convertir a string sin decimales ni formato
      if (typeof v === "number") return [String(Math.round(v))];
      // Si es string con Q, comas y .00 → limpiar
      var s = String(v).replace(/^Q/,"").replace(/,/g,"").replace(/\.00$/,"").trim();
      return [s];
    });
    rango.setValues(nuevos);
  }

  limpiarCol(colDPI);
  limpiarCol(colNIT);

  _alert("✅ DPI y NIT corregidos — ahora son texto plano sin formato moneda.\n" +
         "Esquema: " + (es23 ? "23 cols (P=DPI, Q=NIT)" : "19 cols (L=DPI, M=NIT)"));
}); }

function migrarSistema() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    "⬆️ Migrar sistema",
    "Esto actualiza la estructura a la versión más reciente.\n\n" +
    "• NO se borran participantes ni registros existentes\n" +
    "• Si el esquema es antiguo (Categoria en col I): mueve Hijos_CCI/Num_Hijos_CCI\n" +
    "  desde U/V a I/J y desplaza Categoria→K, etc.\n" +
    "• Se crean hojas Bonos y HijosCCI (si no existen)\n\n" +
    "¿Continuar?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var log = [], errores = [];

  // ── PASO 1: Reorganizar columnas de PARTICIPANTES si es necesario ─
  try {
    var hP = _sh(CFG.HOJAS.PARTICIPANTES);
    var lastCol = hP.getLastColumn();
    var lastRow = hP.getLastRow();
    var nDatos  = Math.max(0, lastRow - 1);

    var enc = lastRow >= 1
      ? hP.getRange(1, 1, 1, Math.max(lastCol, 19)).getValues()[0]
            .map(function(v){ return String(v||"").trim(); })
      : [];

    var colCHeader = enc[2] || ""; // índice 2 = col C
    var colIHeader = enc[8] || ""; // índice 8 = col I

    if (colCHeader === "Proyecto" && (colIHeader === "Categoria" || colIHeader === "Etapa")) {
      // ── MIGRACIÓN 19 → 23 cols: insertar Fecha_Nacimiento/Edad/Genero/Ano_Entrada después de col B ──
      hP.insertColumns(3, 4); // inserta 4 cols vacías en pos 3-6; datos existentes C-S → G-W automático
      hP.getRange(1, 3, 1, 4).setValues([["Fecha_Nacimiento","Edad","Genero","Ano_Entrada_Creamos"]]);
      _fmtEnc(hP, "#639922"); // reaplicar color encabezado (insertColumns puede romper formato)
      hP.getRange("C2:C500").setNumberFormat("dd/MM/yyyy");
      hP.setColumnWidth(3, 110); // Fecha_Nacimiento
      hP.setColumnWidth(4, 60);  // Edad
      hP.setColumnWidth(5, 80);  // Genero
      hP.setColumnWidth(6, 80);  // Ano_Entrada
      // Reaplicar validaciones en nuevas posiciones (las existentes se desplazaron con los datos)
      var vEtapaMig = SpreadsheetApp.newDataValidation().requireValueInList(["Inscritx","Retiradx","Empleadx","Ciclo de Vida Terminado"],true).build();
      var vCatMig   = SpreadsheetApp.newDataValidation().requireValueInList(["A","B","C","D"],true).build();
      var vSiNoMig  = SpreadsheetApp.newDataValidation().requireValueInList(["Sí","No"],true).build();
      var vBancoMig = SpreadsheetApp.newDataValidation().requireValueInList(["Banrural","Industrial","BAC Credomatic","G&T Continental","Banco Azteca","N/A","Otro"],true).build();
      var vTipoCMig = SpreadsheetApp.newDataValidation().requireValueInList(["Monetaria","Ahorro",""],true).build();
      var vPagoMig  = SpreadsheetApp.newDataValidation().requireValueInList(["Transferencia","Cheque"],true).build();
      hP.getRange("I2:I500").setDataValidation(vEtapaMig);
      hP.getRange("M2:M500").setDataValidation(vCatMig);
      hP.getRange("O2:O500").setDataValidation(vSiNoMig);
      hP.getRange("S2:S500").setDataValidation(vBancoMig);
      hP.getRange("T2:T500").setDataValidation(vTipoCMig);
      hP.getRange("V2:V500").setDataValidation(vPagoMig);
      hP.getRange("N2:N500").setNumberFormat("Q#,##0.00");
      log.push("✅ PARTICIPANTES: migrado a 23 cols — cols C-F (Fecha_Nacimiento/Edad/Genero/Año) insertadas" +
        (nDatos > 0 ? " — " + nDatos + " filas conservadas sin cambios" : ""));
      log.push("ℹ️ Ejecuta ⚙️ Admin → 🔄 Sincronizar desde Creamos DB para llenar las nuevas columnas");

    } else if (colCHeader === "Fecha_Nacimiento") {
      // ── Ya en esquema 23 cols ─────────────────────────────────────────────────────────────
      var nConCat23 = 0;
      if (nDatos > 0) {
        var chkDatos23 = hP.getRange(2, 13, nDatos, 1).getValues(); // col M = Categoria
        chkDatos23.forEach(function(r) {
          if (CFG.CATEGORIAS[String(r[0]||"").trim().toUpperCase()]) nConCat23++;
        });
      }
      log.push("ℹ️ PARTICIPANTES: esquema 23 cols (A–W) ya aplicado" +
        (nDatos > 0 ? " — " + nConCat23 + "/" + nDatos + " con Categoria válida" : ""));

    } else if (colIHeader === "Hijos_CCI" || colIHeader === "Num_Hijos_CCI") {
      // ── Esquema INTERMEDIO detectado (22 cols con Hijos_CCI): migrar a 19 cols ─────────
      var nFilas = Math.max(0, lastRow - 1);
      if (nFilas > 0) {
        var datosViejos = hP.getRange(2, 1, nFilas, Math.min(22, hP.getLastColumn())).getValues();
        var datosNuevos = datosViejos.map(function(r) {
          return [
            r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7],
            r[10], r[11], r[12],
            r[13], r[14], r[15], r[16], r[17], r[18], r[19], r[20]
          ];
        });
        hP.getRange(2, 1, nFilas, 19).setValues(datosNuevos);
      }
      hP.getRange(1, 1, 1, 19).setValues([[
        "Creamos_ID","Nombre","Proyecto","Programa","Etapa",
        "Educacion","Apoyo_Emocional","Inclusion_Laboral",
        "Categoria","Tarifa_Hora","Tiene_Factura",
        "DPI","NIT","Correo",
        "Banco","Tipo_Cuenta","Num_Cuenta","Forma_Pago","URL_Doc_Proceso"
      ]]);
      log.push("✅ PARTICIPANTES: migrado a esquema 19 cols (eliminados Hijos_CCI/Estipendio)" +
        (nDatos > 0 ? " (" + nDatos + " filas migradas)" : ""));

    } else if (colIHeader === "Categoria") {
      var nConCat = 0;
      if (nDatos > 0) {
        var chkDatos = hP.getRange(2, 9, nDatos, 1).getValues();
        chkDatos.forEach(function(r) {
          if (CFG.CATEGORIAS[String(r[0]||"").trim().toUpperCase()]) nConCat++;
        });
      }
      log.push("ℹ️ PARTICIPANTES: esquema 19 cols ya aplicado" +
        (nDatos > 0 ? " — " + nConCat + "/" + nDatos + " participantes con Categoria válida" : ""));

    } else {
      log.push("⚠️ PARTICIPANTES: col C = '" + colCHeader + "', col I = '" + colIHeader + "' — encabezado inesperado, sin cambios");
    }

    // Aplicar validaciones base (solo si esquema 19-col antiguo — para 23-col ya se aplican arriba)
    if (colCHeader === "Proyecto") {
      var vSiNo = SpreadsheetApp.newDataValidation().requireValueInList(["Sí","No"],true).build();
      var vCat  = SpreadsheetApp.newDataValidation().requireValueInList(["A","B","C","D"],true).build();
      hP.getRange("I2:I500").setDataValidation(vCat);
      hP.getRange("K2:K500").setDataValidation(vSiNo);
      hP.getRange("J2:J500").setNumberFormat('"Q"#,##0.00');
    }

  } catch(e) { errores.push("❌ PARTICIPANTES: " + e.message); }

  // ── PASO 2: Crear hoja Bonos si no existe ────────────────────
  try {
    if (!ss.getSheetByName("Bonos")) {
      crearHojaBonos();
      log.push("✅ Hoja 'Bonos' creada");
    } else {
      log.push("ℹ️ Hoja 'Bonos' ya existe — sin cambios");
    }
  } catch(e) { errores.push("❌ Bonos: " + e.message); }

  // ── PASO 3: Crear hoja HijosCCI si no existe ─────────────────
  try {
    if (!ss.getSheetByName("HijosCCI")) {
      crearHojaHijosCCI();
      log.push("✅ Hoja 'HijosCCI' creada");
    } else {
      log.push("ℹ️ Hoja 'HijosCCI' ya existe — sin cambios");
    }
  } catch(e) { errores.push("❌ HijosCCI: " + e.message); }

  // ── PASO 4: Actualizar colores e IDs en hojas auxiliares ─────
  try {
    var mapa = _mapaDatosParticipantes(ss);
    ["DiasEstudio","ListaTerapias","InclusionLaboral","HijosCCI"].forEach(function(nm) {
      var h = ss.getSheetByName(nm);
      if (!h || h.getLastRow() < 2) return;
      var nCols = (nm === "DiasEstudio") ? 11 : (nm === "HijosCCI") ? 5 : 4;
      _actualizarIDsEnHoja(h, 2, mapa);
      _colorearHojaApoyo(h, 2, nCols, mapa);
    });
    log.push("✅ Colores e IDs actualizados en hojas auxiliares");
  } catch(e) { errores.push("❌ Colores/IDs: " + e.message); }

  // ── PASO 5: Recalcular Tarifa_Hora desde Categoria ──────────────
  try {
    var hP2 = _sh(CFG.HOJAS.PARTICIPANTES);
    if (hP2 && hP2.getLastRow() > 1) {
      var nRows2 = hP2.getLastRow() - 1;
      // Detectar esquema por encabezado col C
      var encP2 = hP2.getRange(1, 3, 1, 1).getValue();
      var es23cols = String(encP2||"").trim() === "Fecha_Nacimiento";
      var iCat  = es23cols ? 12 : 8;  // 0-based: col M(13)=12 en 23-col, col I(9)=8 en 19-col
      var iTar  = es23cols ? 13 : 9;
      var iFact = es23cols ? 14 : 10;
      var colTar  = es23cols ? 14 : 10; // 1-based col N o J
      var colFact = es23cols ? 15 : 11;
      var nCols2  = es23cols ? 15 : 11;
      var datos2 = hP2.getRange(2, 1, nRows2, nCols2).getValues();
      var tarifs2 = [], facts2 = [];
      datos2.forEach(function(r) {
        var cat2    = String(r[iCat]||"").trim().toUpperCase();
        var tarifa2 = CFG.CATEGORIAS[cat2];
        tarifs2.push([tarifa2 || (r[iTar] || "")]);
        var fact2 = String(r[iFact]||"").trim();
        facts2.push([(fact2 === "Sí" || fact2 === "No") ? fact2 : "Sí"]);
      });
      hP2.getRange(2, colTar,  nRows2, 1).setValues(tarifs2);
      hP2.getRange(2, colFact, nRows2, 1).setValues(facts2);
      log.push("✅ Tarifas y Tiene_Factura recalculadas desde Categoria (" + (es23cols ? "23 cols" : "19 cols") + ")");
    }
  } catch(e) { errores.push("❌ Tarifas: " + e.message); }

  _alert(
    "⬆️ MIGRACIÓN COMPLETADA\n\n" +
    log.join("\n") +
    (errores.length ? "\n\n❌ ERRORES:\n" + errores.join("\n") : "") +
    "\n\n✅ Tus datos existentes no fueron modificados.\n\n" +
    "Esquema actual de PARTICIPANTES (23 cols A–W):\n" +
    "• Cols C–F = Fecha_Nacimiento / Edad / Genero / Año Entrada Creamos\n" +
    "• Col I = Etapa  |  Col M = Categoria  |  Col N = Tarifa\n" +
    "• Col O = Tiene_Factura  |  Col P = DPI  |  Col W = URL Doc\n\n" +
    "Siguiente paso: ⚙️ Admin → 🔄 Sincronizar desde Creamos DB\n" +
    "para llenar los datos de Fecha_Nacimiento, Edad, Genero y Año."
  );
}); }

// ══════════════════════════════════════════════════════════════════
// INSTALACIÓN COMPLETA — wizard de 3 pasos
// ══════════════════════════════════════════════════════════════════

function instalarTodo() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert("🚀 INSTALACIÓN COMPLETA — " + CFG.PROYECTO + " / " + CFG.ORG,
    "Se ejecutarán 9 pasos automáticamente:\n\n" +
    "1 — Crear hojas: PARTICIPANTES (19 cols), DatosKobo, PERIODOS\n" +
    "2 — Importar datos desde Kobo\n" +
    "3 — Crear estructura en Drive (Docs_Proceso, Reportes)\n" +
    "4 — Crear hoja Días de Estudio\n" +
    "5 — Crear hoja Lista de Terapias\n" +
    "6 — Crear hojas Inclusión Laboral y Retiradx\n" +
    "7 — Crear hojas Bonos y HijosCCI\n" +
    "8 — Activar automatizaciones (Kobo SOLO manual + estipendio auto cada 6h)\n" +
    "9 — Crear Guía de Uso\n\n" +
    "Los datos existentes NO se borran.\n\n¿Continuar?",
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [], errores = [];

  // PASO 1
  try {
    ss.toast("Paso 1/7: Creando hojas (PARTICIPANTES, DatosKobo, PERIODOS)...", "🚀", -1);
    crearHojas();
    log.push("✅ Paso 1: Hojas del sistema creadas/verificadas");
  } catch(e) { errores.push("❌ Paso 1: " + e.message); }
  Utilities.sleep(500);

  // PASO 2: Importar Kobo — historial completo, dedup por UUID + clave compuesta
  try {
    ss.toast("Paso 2/9: Importando datos de Kobo (historial completo, sin duplicados)...", "🚀", -1);
    var res = UrlFetchApp.fetch(CFG.KOBO_URL_CSV, { muteHttpExceptions: true });
    var code = res.getResponseCode();
    if (code === 503) {
      log.push("⏳ Paso 2: Kobo ocupado (503) — importa manualmente desde el menú después");
    } else if (code !== 200) {
      errores.push("❌ Paso 2: Error Kobo HTTP " + code);
    } else {
      var datosRaw2 = Utilities.parseCsv(res.getContentText(), ";");
      if (datosRaw2.length < 2) {
        log.push("⚠️ Paso 2: Kobo sin registros");
      } else {
        // Sin filtro de fecha — todo el histórico; dedup previene duplicados
        // Excluir registros de Manufactura (columna "Destino")
        var colsDestinoI = detectarColumnas(datosRaw2[0], []);
        if (colsDestinoI.destino !== undefined) {
          var filtradoDestinoI = [datosRaw2[0]];
          for (var fdi = 1; fdi < datosRaw2.length; fdi++) {
            var valDestinoI = String(datosRaw2[fdi][colsDestinoI.destino] || "").trim().toLowerCase();
            if (valDestinoI !== "manufactura") filtradoDestinoI.push(datosRaw2[fdi]);
          }
          datosRaw2 = filtradoDestinoI;
        }
        var datosN2 = _filtrarColumnasKobo(datosRaw2);
        var hK = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
        if (!hK) {
          // Instalación nueva: crear hoja y escribir todo el histórico
          hK = ss.insertSheet(CFG.HOJAS.DATOS_KOBO);
          hK.getRange(1,1,datosN2.length,datosN2[0].length).setValues(datosN2);
          hK.getRange(1,1,1,datosN2[0].length).setFontWeight("bold").setBackground("#4a86e8").setFontColor("#fff");
          hK.setFrozenRows(1);
          _limpiarColumnasKobo(hK, datosN2[0]);
          _normalizarAccionSilencioso(hK);
          log.push("✅ Paso 2: " + (datosN2.length-1) + " registros importados desde Kobo");
        } else {
          // Reinstalación: solo agregar registros nuevos (UUID + clave compuesta)
          var encN2 = datosN2[0];
          var uuidCN2 = _buscarIndice(encN2, "_uuid");
          var partCN2 = detectarColumnas(encN2, []).participante;
          var datosEx2 = hK.getDataRange().getValues();
          var uuidCE2 = _buscarIndice(datosEx2[0], "_uuid");
          var partCE2 = detectarColumnas(datosEx2[0], []).participante;
          var uuidsEx2 = {}, compositeEx2 = {};
          for (var ue2=1; ue2<datosEx2.length; ue2++) {
            var uv2 = uuidCE2>=0 ? String(datosEx2[ue2][uuidCE2]||"").trim() : "";
            if (uv2) uuidsEx2[uv2] = true;
            var ts2e = _ckTs(datosEx2[ue2][0]);
            var pt2e = _ckPart(partCE2>=0 ? datosEx2[ue2][partCE2] : "");
            if (ts2e && pt2e) compositeEx2[ts2e+"|"+pt2e] = true;
          }
          var nuevas2 = [];
          for (var jj2=1; jj2<datosN2.length; jj2++) {
            var uid2 = uuidCN2>=0 ? String(datosN2[jj2][uuidCN2]||"").trim() : "";
            if (uid2 && uuidsEx2[uid2]) continue;
            var ts2n = _ckTs(datosN2[jj2][0]);
            var pt2n = _ckPart(partCN2>=0 ? datosN2[jj2][partCN2] : "");
            if (ts2n && pt2n && compositeEx2[ts2n+"|"+pt2n]) continue;
            nuevas2.push(datosN2[jj2]);
          }
          if (nuevas2.length > 0) {
            hK.getRange(hK.getLastRow()+1,1,nuevas2.length,nuevas2[0].length).setValues(nuevas2);
            _normalizarAccionSilencioso(hK);
            log.push("✅ Paso 2: " + nuevas2.length + " registros nuevos agregados");
          } else {
            log.push("ℹ️ Paso 2: DatosKobo ya al día, sin registros nuevos");
          }
        }
      }
    }
  } catch(e) { errores.push("❌ Paso 2: " + e.message); }
  Utilities.sleep(500);

  // PASO 3: Drive
  try {
    ss.toast("Paso 3/7: Creando estructura en Drive...", "🚀", -1);
    crearEstructuraDrive();
    log.push("✅ Paso 3: Estructura Drive creada/verificada");
  } catch(e) { errores.push("❌ Paso 3: " + e.message); }
  Utilities.sleep(500);

  // PASO 4: DiasEstudio
  try {
    ss.toast("Paso 4/7: Configurando Días de Estudio...", "🚀", -1);
    if (!ss.getSheetByName("DiasEstudio")) { crearHojaDiasEstudio(); log.push("✅ Paso 4: Hoja DiasEstudio creada"); }
    else { log.push("ℹ️ Paso 4: Hoja DiasEstudio ya existe"); }
  } catch(e) { errores.push("❌ Paso 4: " + e.message); }
  Utilities.sleep(300);

  // PASO 5: ListaTerapias
  try {
    ss.toast("Paso 5/7: Configurando Lista de Terapias...", "🚀", -1);
    if (!ss.getSheetByName("ListaTerapias")) { crearHojaListaTerapias(); log.push("✅ Paso 5: Hoja ListaTerapias creada"); }
    else { log.push("ℹ️ Paso 5: Hoja ListaTerapias ya existe"); }
  } catch(e) { errores.push("❌ Paso 5: " + e.message); }
  Utilities.sleep(300);

  // PASO 6: InclusionLaboral + Retiradx + CiclosVida
  try {
    ss.toast("Paso 6/9: Creando hojas Inclusión Laboral, Retiradx y CiclosVida...", "🚀", -1);
    if (!ss.getSheetByName("InclusionLaboral")) { crearHojaInclusionLaboral(); log.push("✅ Paso 6a: Hoja InclusionLaboral creada"); }
    else { log.push("ℹ️ Paso 6a: Hoja InclusionLaboral ya existe"); }
    if (!ss.getSheetByName("Retiradx")) { crearHojaRetiradx(); log.push("✅ Paso 6b: Hoja Retiradx creada"); }
    else { log.push("ℹ️ Paso 6b: Hoja Retiradx ya existe"); }
    if (!ss.getSheetByName("CiclosVida")) { crearHojaCiclosVida(); log.push("✅ Paso 6c: Hoja CiclosVida creada"); }
    else { log.push("ℹ️ Paso 6c: Hoja CiclosVida ya existe"); }
  } catch(e) { errores.push("❌ Paso 6: " + e.message); }
  Utilities.sleep(300);

  // PASO 7: Bonos + HijosCCI + Cheques + Transferencias
  try {
    ss.toast("Paso 7/9: Creando hojas Bonos, Estipendio, HijosCCI, Cheques y Transferencias...", "🚀", -1);
    if (!ss.getSheetByName("Bonos")) { crearHojaBonos(); log.push("✅ Paso 7a: Hoja Bonos creada"); }
    else { log.push("ℹ️ Paso 7a: Hoja Bonos ya existe"); }
    if (!ss.getSheetByName("Estipendio")) { crearHojaEstipendio(); log.push("✅ Paso 7b: Hoja Estipendio creada"); }
    else { log.push("ℹ️ Paso 7b: Hoja Estipendio ya existe"); }
    if (!ss.getSheetByName("HijosCCI")) { crearHojaHijosCCI(); log.push("✅ Paso 7c: Hoja HijosCCI creada"); }
    else { log.push("ℹ️ Paso 7c: Hoja HijosCCI ya existe"); }
    if (!ss.getSheetByName("Cheques")) { crearHojaCheques(); log.push("✅ Paso 7d: Hoja Cheques creada"); }
    else { log.push("ℹ️ Paso 7c: Hoja Cheques ya existe"); }
    if (!ss.getSheetByName("Transferencias")) { crearHojaTransferencias(); log.push("✅ Paso 7d: Hoja Transferencias creada"); }
    else { log.push("ℹ️ Paso 7d: Hoja Transferencias ya existe"); }
  } catch(e) { errores.push("❌ Paso 7: " + e.message); }
  Utilities.sleep(300);

  // PASO 8: Triggers
  try {
    ss.toast("Paso 8/9: Activando automatizaciones...", "🚀", -1);
    configurarTriggers();
    log.push("✅ Paso 8: Triggers activados (estipendio cada 6h + dashboard 10min + al abrir)");
  } catch(e) { errores.push("❌ Paso 8: " + e.message); }

  // PASO 9: Guía de Uso
  try {
    ss.toast("Paso 9/9: Creando Guía de Uso...", "📖", -1);
    crearGuiaUso();
    log.push("✅ Paso 9: Guía de Uso creada");
  } catch(e) { errores.push("❌ Paso 9: " + e.message); }

  ss.toast("", "", 1);
  var resumen = "🚀 INSTALACIÓN COMPLETA — " + CFG.ORG + "\n\n";
  resumen += log.join("\n");
  if (errores.length) resumen += "\n\n--- PROBLEMAS ---\n" + errores.join("\n");
  resumen += "\n\n--- HOJAS DEL SISTEMA ---\n";
  resumen += "• PARTICIPANTES     — lista maestra (19 cols: Etapa, Banco, cuenta — sin Hijos CCI/Estipendio)\n";
  resumen += "• DatosKobo         — datos de Kobo (se actualiza automático cada hora)\n";
  resumen += "• PERIODOS          — control de quincenas\n";
  resumen += "• DiasEstudio       — qué días estudia cada participante\n";
  resumen += "• ListaTerapias     — quién recibe terapia\n";
  resumen += "• InclusionLaboral  — quién participa en Inclusión Laboral\n";
  resumen += "• Retiradx          — registro de retiros con razón\n";
  resumen += "• Bonos             — bonos individuales (se suman al pago de quincena)\n";
  resumen += "• HijosCCI         — hijos en CCI (se refleja en PARTICIPANTES)\n";
  resumen += "• Q_[fecha]         — reporte generado por quincena (15 cols + bono/estip)\n";
  resumen += "\nTarifas: A=Q16.50 | B=Q15.75 | C=Q15.00 | D=Q14.00\n";
  resumen += "IVA 5%: Admin → Configurar IVA (quién tiene factura)\n";
  resumen += "Etapa: Inscritx / Retiradx / Empleadx / Ciclo de Vida Terminado\n";
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

  // Encabezado siempre — A=Creamos_ID, B=Participante, C-I=días, J-K=fechas
  var enc = ["Creamos_ID","Participante","Lun","Mar","Mié","Jue","Vie","Sáb","Dom","Fecha_Inicio","Fecha_Fin"];
  hoja.getRange(1,1,1,enc.length).setValues([enc])
    .setFontWeight("bold").setBackground("#7b1fa2").setFontColor("#fff").setHorizontalAlignment("center");
  hoja.setFrozenRows(1);

  var mapa = _mapaDatosParticipantes(ss);

  if (esNueva) {
    var filas = LISTA_OFICIAL.map(function(it) {
      var id = (mapa[it[1]] || {}).id || it[3] || "";
      return [id, it[1], "","","","","","","","",""];
    });
    hoja.getRange(2, 1, filas.length, 11).setValues(filas);
    var vX = SpreadsheetApp.newDataValidation().requireValueInList(["X",""],true).build();
    hoja.getRange(2,3,filas.length,7).setDataValidation(vX).setHorizontalAlignment("center"); // cols C-I
  } else {
    _actualizarIDsEnHoja(hoja, 2, mapa); // hoja existente: actualizar IDs
  }

  // Colores por categoría — siempre desde PARTICIPANTES (col B = Nombre)
  _colorearHojaApoyo(hoja, 2, 11, mapa);

  hoja.setColumnWidth(1, 110); // Creamos_ID
  hoja.setColumnWidth(2, 260); // Participante
  for (var c=3;c<=9;c++) hoja.setColumnWidth(c,55);
  hoja.setColumnWidth(10,110); hoja.setColumnWidth(11,110);
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

  hoja.getRange(1,1,1,4).setValues([["Creamos_ID","Participante","Recibe Terapia (X)","Notas"]])
    .setFontWeight("bold").setBackground("#00897b").setFontColor("#fff").setHorizontalAlignment("center");
  hoja.setFrozenRows(1);

  var mapa = _mapaDatosParticipantes(ss);

  if (esNueva) {
    var filas = LISTA_OFICIAL.map(function(it) {
      var id = (mapa[it[1]] || {}).id || it[3] || "";
      return [id, it[1], "", ""];
    });
    hoja.getRange(2, 1, filas.length, 4).setValues(filas);
    var vX = SpreadsheetApp.newDataValidation().requireValueInList(["X",""],true).build();
    hoja.getRange(2,3,filas.length,1).setDataValidation(vX).setHorizontalAlignment("center");
  } else {
    // Hoja ya existe: actualizar IDs
    _actualizarIDsEnHoja(hoja, 2, mapa);
  }

  // Colores por categoría (siempre, nueva o existente)
  _colorearHojaApoyo(hoja, 2, 4, mapa);

  hoja.setColumnWidth(1,120); hoja.setColumnWidth(2,260);
  hoja.setColumnWidth(3,150); hoja.setColumnWidth(4,300);
  hoja.activate();

  _alert(
    "🧘 LISTA DE TERAPIAS — Para qué sirve:\n\n" +
    "Marca con X a las personas que reciben sesiones de terapia.\n" +
    "Cuando llegue un registro de Kobo con tipo 'Terapia' o 'Salida Terapia',\n" +
    "el sistema lo contabiliza como horas trabajadas.\n\n" +
    "Colores y Creamos IDs se sincronizan desde PARTICIPANTES automáticamente."
  );
}); }

// ══════════════════════════════════════════════════════════════════
// DIAGNÓSTICO, REPARACIÓN Y CAMBIO DE NOMBRE
// ══════════════════════════════════════════════════════════════════

/**
 * Marca en amarillo/rojo las filas sospechosas de DatosKobo:
 *   🟡 Amarillo: par ENTRADA+SALIDA con < 30 min de diferencia (posible error)
 *   🔴 Rojo:     ENTRADA cruzó medianoche (start PM ≥ 12, end día distinto) — stale form
 *
 * Después revisa las filas rojas y usa "Eliminar filas rojas" para borrarlas.
 */
function marcarFilasSospechosas() { _run(function() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hojaK = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (!hojaK || hojaK.getLastRow() < 2) {
    _alert("DatosKobo vacío o inexistente."); return;
  }

  var datos   = hojaK.getDataRange().getValues();
  var cols    = detectarColumnas(datos[0], datos.slice(1));
  var nCols   = hojaK.getLastColumn();

  if (cols.start === undefined) {
    _alert("No se detectó columna 'start' en DatosKobo. Importa desde Kobo primero."); return;
  }

  // Limpiar marcas previas
  hojaK.getRange(2, 1, hojaK.getLastRow()-1, nCols).setBackground(null);

  var nRojo = 0, nAmarillo = 0;

  for (var i = 1; i < datos.length; i++) {
    var fila   = datos[i];
    var tsS    = new Date(fila[cols.start]);
    var tsE    = cols.end !== undefined ? new Date(fila[cols.end]) : null;
    if (isNaN(tsS)) continue;

    var tipo   = obtenerTipoRegistro(fila, cols);
    if (!tipo.esIngreso && !tipo.esEgreso) continue;

    var bg = null;

    if (tsE && !isNaN(tsE)) {
      var mismoD = (tsS.getFullYear()===tsE.getFullYear() &&
                    tsS.getMonth()===tsE.getMonth() &&
                    tsS.getDate()===tsE.getDate());
      var startH = parseInt(Utilities.formatDate(tsS, CFG.TIMEZONE, "H"), 10);
      var diffMin = (tsE - tsS) / 60000;

      if (!mismoD && tipo.esIngreso && startH >= 12) {
        // ENTRADA abierta en la tarde de un día, enviada al día siguiente — stale form
        bg = "#ffcdd2"; // rojo claro
        nRojo++;
      } else if (mismoD && diffMin >= 0 && diffMin < 30) {
        // start y end en < 30 minutos — posible duplicado o error
        bg = "#fff9c4"; // amarillo claro
        nAmarillo++;
      }
    }

    if (bg) hojaK.getRange(i+1, 1, 1, nCols).setBackground(bg);
  }

  _alert(
    "✅ Marcado completado\n\n" +
    "🔴 Rojas: " + nRojo + " filas (ENTRADA stale — form abierto el día anterior en la tarde)\n" +
    "🟡 Amarillas: " + nAmarillo + " filas (start ↔ end < 30 min — revisar)\n\n" +
    "Pasos:\n" +
    "1. Revisa las filas amarillas: ¿son válidas o errores?\n" +
    "2. Las filas rojas casi siempre son basura → usa\n" +
    "   📥 Asistencia → 🗑️ Eliminar filas marcadas en rojo"
  );
}); }

/**
 * Elimina de DatosKobo todas las filas con fondo rojo (marcadas por marcarFilasSospechosas).
 */
function eliminarFilasRojas() { _run(function() {
  var ui    = SpreadsheetApp.getUi();
  var hojaK = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (!hojaK || hojaK.getLastRow() < 2) { _alert("DatosKobo vacío."); return; }

  var lastRow = hojaK.getLastRow();
  var nCols   = hojaK.getLastColumn();
  var bgs     = hojaK.getRange(2, 1, lastRow-1, nCols).getBackgrounds();

  // Contar filas rojas
  var filasRojas = [];
  for (var i = 0; i < bgs.length; i++) {
    if (bgs[i][0] === "#ffcdd2") filasRojas.push(i+2); // 1-indexed
  }

  if (filasRojas.length === 0) {
    _alert("No hay filas rojas. Usa primero 'Marcar filas sospechosas'."); return;
  }

  var conf = ui.alert(
    "🗑️ Eliminar filas rojas",
    "Se borrarán " + filasRojas.length + " filas marcadas en rojo de DatosKobo.\n\n¿Continuar?",
    ui.ButtonSet.YES_NO
  );
  if (conf !== ui.Button.YES) return;

  // Borrar de abajo hacia arriba para no afectar índices
  for (var j = filasRojas.length-1; j >= 0; j--) {
    hojaK.deleteRow(filasRojas[j]);
  }

  _alert("✅ " + filasRojas.length + " filas eliminadas.\n\nRegenetra el reporte para ver los cambios.");
}); }

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
  var tz     = CFG.TIMEZONE;

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

    var _tsCol2 = (cols.submissionTime !== undefined) ? cols.submissionTime : (cols.end !== undefined) ? cols.end : cols.start;
    var tsRaw = _tsCol2 !== undefined ? fila[_tsCol2] : null;
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

  // ── Sección 4b: Nombres que no coinciden con ninguna PARTICIPANTE ─
  var hP2 = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES);
  var nombresParticipantes = {};
  if (hP2 && hP2.getLastRow() >= 2) {
    hP2.getRange(2, 2, hP2.getLastRow()-1, 1).getValues().forEach(function(r) {
      var n = String(r[0]||"").trim();
      if (n) nombresParticipantes[n.toLowerCase()] = n;
    });
  }
  var nomNoMatch = [];
  Object.keys(resumenPart).forEach(function(nom) {
    var k = nom.toLowerCase().replace(/[áàä]/g,"a").replace(/[éèë]/g,"e")
               .replace(/[íìï]/g,"i").replace(/[óòö]/g,"o").replace(/[úùü]/g,"u");
    var match = nombresParticipantes[nom.toLowerCase()] ||
      Object.keys(nombresParticipantes).some(function(pk){
        return pk.indexOf(k) !== -1 || k.indexOf(pk) !== -1;
      });
    if (!match) nomNoMatch.push(nom);
  });
  if (nomNoMatch.length > 0) {
    p(["❌  NOMBRES KOBO SIN COINCIDENCIA EN PARTICIPANTES — horas perdidas","","","","",""],"sec_warn");
    p(["Nombre en Kobo (normalizado)","Acción","","","",""],"enc");
    nomNoMatch.forEach(function(n){
      p([n, "Agregar a PARTICIPANTES o a hoja NombresCanonicos","","","",""],"warn");
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

function eliminarColumnasKobo() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var hojaK = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (!hojaK) { _alert("No existe DatosKobo."); return; }

  var enc = hojaK.getRange(1, 1, 1, hojaK.getLastColumn()).getValues()[0];
  var aEliminar = [];
  for (var i = 0; i < enc.length; i++) {
    var h = String(enc[i]).trim().toLowerCase();
    var esNecesaria = h && _COLS_KOBO_NECESARIAS.some(function(p){ return h.indexOf(p) !== -1; });
    if (!esNecesaria) aEliminar.push(i + 1); // 1-based
  }

  if (!aEliminar.length) { ui.alert("✅ DatosKobo ya está limpia — no hay columnas innecesarias."); return; }

  var nombresAEliminar = aEliminar.map(function(c){ return String(enc[c-1]||"Col "+c).trim(); });
  if (ui.alert("🧹 ELIMINAR COLUMNAS",
    "Se eliminarán " + aEliminar.length + " columnas:\n" + nombresAEliminar.join(", ") +
    "\n\nEsta acción no se puede deshacer. ¿Continuar?",
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  // Eliminar de derecha a izquierda para no desplazar índices
  for (var ri = aEliminar.length - 1; ri >= 0; ri--) {
    hojaK.deleteColumn(aEliminar[ri]);
  }
  ui.alert("✅ Listo. Se eliminaron " + aEliminar.length + " columnas de DatosKobo.");
}); }

function repararDatosKobo() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaK = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (!hojaK) { _alert("No existe DatosKobo."); return; }

  var datos = hojaK.getDataRange().getValues();
  var cols = detectarColumnas(datos[0], datos.slice(1));
  var mapeoN = cargarMapeoNombres();

  // Fase 0: detectar participantes no oficiales
  var nombresOficialesSet = {};
  LISTA_OFICIAL.forEach(function(it) { nombresOficialesSet[it[1]] = true; });
  var filasNoOficiales = []; // índices 1-based en la hoja
  if (cols.participante !== undefined) {
    for (var f=1; f<datos.length; f++) {
      var rawN = String(datos[f][cols.participante]||"").trim();
      var normN = rawN ? normalizarNombre(rawN, mapeoN) : "";
      if (rawN && !nombresOficialesSet[normN]) filasNoOficiales.push(f+1); // +1 = fila hoja
    }
  }

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
      if (n) { var norm = normalizarNombre(n, mapeoN); if (norm !== n) nombresANorm++; }
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
  if (filasNoOficiales.length) msg += "\n🚫 Filas con participantes NO oficiales: "+filasNoOficiales.length+" (se eliminarán)\n";

  if (!Object.keys(desconocidos).length && !nombresANorm && !filasNoOficiales.length) { ui.alert(msg+"\nNo hay nada que reparar."); return; }

  if (ui.alert("🔧 Reparar Datos", msg+"\n¿Aplicar correcciones?", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  var cam1=0, cam2=0;
  if (cols.accionUnificada !== undefined) {
    for (var f=1; f<datos.length; f++) {
      var v = String(datos[f][cols.accionUnificada]||"").trim();
      if (v && correcciones[v]) { hojaK.getRange(f+1,cols.accionUnificada+1).setValue(correcciones[v]); cam1++; }
    }
  }
  // Fase 2: normalizar nombres y poblar c_id
  datos = hojaK.getDataRange().getValues();
  cols = detectarColumnas(datos[0], datos.slice(1));
  var mapaNameAId = {};
  LISTA_OFICIAL.forEach(function(it) { if (it[3]) mapaNameAId[it[1]] = it[3]; });

  // Si c_id no existe, crear columna
  if (cols.creamos_id === undefined) {
    var nc = hojaK.getLastColumn() + 1;
    hojaK.getRange(1, nc).setValue("c_id")
        .setFontWeight("bold").setBackground("#4a86e8").setFontColor("#fff");
    cols.creamos_id = nc - 1;
    datos = hojaK.getDataRange().getValues(); // recargar con nueva col
  }

  if (cols.participante !== undefined) {
    var cam2b = 0;
    for (var f=1; f<datos.length; f++) {
      var n = String(datos[f][cols.participante]||"").trim();
      if (!n) continue;
      var norm = normalizarNombre(n, mapeoN);
      if (norm !== n) { hojaK.getRange(f+1, cols.participante+1).setValue(norm); cam2++; }
      var idActual = String(datos[f][cols.creamos_id]||"").trim();
      var idCorrecto = mapaNameAId[norm] || mapaNameAId[n] || "";
      if (idCorrecto && idCorrecto !== idActual) {
        hojaK.getRange(f+1, cols.creamos_id+1).setValue(idCorrecto); cam2b++;
      }
    }
    cam2 = cam2 + (cam2b > 0 ? " (+" + cam2b + " c_id)" : "");
  }
  // Fase 3: eliminar filas no oficiales (de abajo hacia arriba para no desplazar índices)
  var cam3 = filasNoOficiales.length;
  for (var ri = filasNoOficiales.length - 1; ri >= 0; ri--) {
    hojaK.deleteRow(filasNoOficiales[ri]);
  }
  ui.alert("✅ REPARACIÓN COMPLETADA\n\nEntrada/Salida corregidos: "+cam1+"\nNombres/IDs normalizados: "+cam2+"\nFilas no oficiales eliminadas: "+cam3);
}); }

/**
 * Elimina filas duplicadas de PARTICIPANTES.
 * Agrupa por nombre normalizado, conserva la fila con Creamos_ID real;
 * si ninguna lo tiene, conserva la primera. Borra las demás.
 */
function limpiarDuplicadosParticipantes() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var hP = _sh(CFG.HOJAS.PARTICIPANTES);
  if (hP.getLastRow() < 3) { _alert("No hay datos que limpiar."); return; }

  var datos = hP.getRange(2, 1, hP.getLastRow()-1, 23).getValues();
  var normMap = {}; // norma → índice ganador (0-based en datos)

  datos.forEach(function(f, i) {
    var nombre = String(f[1]||"").trim();
    if (!nombre) return;
    var norm = textoParaComparar(nombre);
    if (normMap[norm] === undefined) {
      normMap[norm] = i;
    } else {
      // Conservar quien tenga Creamos_ID real
      var winner = normMap[norm];
      var idNew  = String(f[0]||"").trim();
      var idOld  = String(datos[winner][0]||"").trim();
      if (_esCreamos_ID_real(idNew) && !_esCreamos_ID_real(idOld)) {
        normMap[norm] = i; // nuevo gana
      }
      // Si ambos tienen ID o ninguno, conservar el primero (winner queda igual)
    }
  });

  var filasAConservar = {};
  Object.keys(normMap).forEach(function(k){ filasAConservar[normMap[k]] = true; });

  // Recolectar filas a eliminar (en orden descendente para no desplazar índices)
  var aEliminar = [];
  datos.forEach(function(f, i) {
    var nombre = String(f[1]||"").trim();
    if (!nombre) { aEliminar.push(i); return; } // vacías también
    if (!filasAConservar[i]) aEliminar.push(i);
  });

  if (aEliminar.length === 0) { _alert("✅ No hay duplicados. PARTICIPANTES está limpio."); return; }

  var resp = ui.alert("🧹 Limpiar duplicados",
    "Se encontraron " + aEliminar.length + " fila(s) duplicadas o vacías en PARTICIPANTES.\n\n¿Eliminarlas?",
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  // Borrar de abajo hacia arriba
  aEliminar.reverse().forEach(function(i) {
    hP.deleteRow(i + 2); // +2 porque la fila 1 es encabezado
  });

  _alert("✅ " + aEliminar.length + " filas duplicadas/vacías eliminadas.\n\nRecalcula la quincena para ver el reporte limpio.");
}); }

function cambiarCategoriaParticipante() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var hP = _sh(CFG.HOJAS.PARTICIPANTES);
  if (!hP || hP.getLastRow() < 2) { _alert("No hay participantes cargados."); return; }

  // Construir lista de nombres
  var datos  = hP.getRange(2, 1, hP.getLastRow() - 1, 13).getValues();
  var lista  = datos.map(function(f, i) {
    return (i + 1) + ". " + String(f[1] || "").trim() + "  [" + String(f[12] || "?") + "]"; // f[12]=Categoria
  }).filter(function(s) { return s.indexOf(". ") !== -1 && s.length > 5; });

  var r1 = ui.prompt(
    "🔼 Cambiar categoría",
    "Escribe el número o nombre de la participante:\n\n" + lista.slice(0, 33).join("\n"),
    ui.ButtonSet.OK_CANCEL
  );
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var buscar = r1.getResponseText().trim();
  if (!buscar) return;

  // Encontrar la fila
  var filaEncontrada = -1, nombreEncontrado = "", catActual = "";
  var numBuscar = parseInt(buscar, 10);
  for (var i = 0; i < datos.length; i++) {
    var nombre = String(datos[i][1] || "").trim();
    if (!nombre) continue;
    var coincide = (!isNaN(numBuscar) && numBuscar === i + 1) ||
                   textoParaComparar(nombre).indexOf(textoParaComparar(buscar)) !== -1;
    if (coincide) {
      filaEncontrada = i + 2;
      nombreEncontrado = nombre;
      catActual = String(datos[i][12] || "").trim().toUpperCase(); // col M = Categoria (idx 12)
      break;
    }
  }
  if (filaEncontrada < 0) { _alert("No se encontró: " + buscar); return; }

  var tarifaActual = CFG.CATEGORIAS[catActual] || 0;
  var r2 = ui.prompt(
    "🔼 Nueva categoría para " + nombreEncontrado,
    "Categoría actual: " + catActual + " (Q" + tarifaActual.toFixed(2) + "/hr)\n\n" +
    "Escribe la nueva categoría:\n" +
    "  A = Q" + CFG.CATEGORIAS.A.toFixed(2) + "/hr\n" +
    "  B = Q" + CFG.CATEGORIAS.B.toFixed(2) + "/hr\n" +
    "  C = Q" + CFG.CATEGORIAS.C.toFixed(2) + "/hr\n" +
    "  D = Q" + CFG.CATEGORIAS.D.toFixed(2) + "/hr",
    ui.ButtonSet.OK_CANCEL
  );
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var nuevaCat = r2.getResponseText().trim().toUpperCase();
  if (!CFG.CATEGORIAS[nuevaCat]) { _alert("Categoría inválida: " + nuevaCat + "\nDebe ser A, B, C o D."); return; }
  if (nuevaCat === catActual) { _alert("La categoría ya es " + catActual + ". No hubo cambio."); return; }

  var nuevaTarifa = CFG.CATEGORIAS[nuevaCat];
  hP.getRange(filaEncontrada, 13).setValue(nuevaCat);    // col M = Categoria
  hP.getRange(filaEncontrada, 14).setValue(nuevaTarifa); // col N = Tarifa_Hora

  // Actualizar LISTA_OFICIAL en memoria (no persiste pero refleja el cambio visual)
  _colorearParticipantes(hP, hP.getLastRow() - 1);

  _alert(
    "✅ Categoría actualizada\n\n" +
    "Participante: " + nombreEncontrado + "\n" +
    catActual + " → " + nuevaCat + "\n" +
    "Q" + tarifaActual.toFixed(2) + " → Q" + nuevaTarifa.toFixed(2) + " / hora\n\n" +
    "El cambio aplica a partir de la próxima quincena."
  );
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

// ══════════════════════════════════════════════════════════════════
// ELIMINAR PARTICIPANTE COMPLETO
// ══════════════════════════════════════════════════════════════════

/**
 * Elimina un participante y TODOS sus registros de todas las hojas relevantes.
 * Busca por Creamos_ID (exacto) o por Nombre (normalizado).
 */
function eliminarParticipanteCompleto() { _run(function() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Paso 1: solicitar nombre o ID
  var r1 = ui.prompt(
    "🗑️ Eliminar participante completo",
    "Nombre o Creamos_ID del participante a eliminar:",
    ui.ButtonSet.OK_CANCEL
  );
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var busqueda = r1.getResponseText().trim();
  if (!busqueda) { _alert("No ingresaste ningún valor."); return; }

  // Paso 2: buscar en PARTICIPANTES
  var hP = _sh(CFG.HOJAS.PARTICIPANTES);
  if (hP.getLastRow() < 2) { _alert("PARTICIPANTES está vacía."); return; }

  var datosP = hP.getRange(2, 1, hP.getLastRow() - 1, 19).getValues();
  var normBusqueda = textoParaComparar(busqueda);
  var filaEncontrada = -1;
  var nombreEncontrado = "", idEncontrado = "";

  for (var i = 0; i < datosP.length; i++) {
    var idFila     = String(datosP[i][0] || "").trim();
    var nombreFila = String(datosP[i][1] || "").trim();
    if (!nombreFila) continue;
    var coincideID     = idFila && idFila.toUpperCase() === busqueda.toUpperCase();
    var coincideNombre = textoParaComparar(nombreFila) === normBusqueda;
    if (coincideID || coincideNombre) {
      filaEncontrada  = i + 2; // 1-indexed (fila 1 = encabezado)
      nombreEncontrado = nombreFila;
      idEncontrado    = idFila;
      break;
    }
  }

  if (filaEncontrada < 0) {
    _alert("No se encontró ningún participante con: \"" + busqueda + "\".\n\nVerifica el nombre o ID exacto en PARTICIPANTES.");
    return;
  }

  // Paso 3: confirmación
  var HOJAS_AFECTADAS = [
    "PARTICIPANTES", "FACTURACION", "DiasEstudio", "ListaTerapias",
    "InclusionLaboral", "HijosCCI", "Bonos", "Cheques", "Transferencias",
    "Retiradx", "CiclosVida"
  ];
  var r2 = ui.alert(
    "⚠️ CONFIRMAR ELIMINACIÓN",
    "Se eliminará a:\n\n" +
    "Nombre:     " + nombreEncontrado + "\n" +
    "Creamos ID: " + (idEncontrado || "—") + "\n\n" +
    "Y TODOS sus registros en:\n" +
    HOJAS_AFECTADAS.join(", ") + "\n\n" +
    "Esta acción NO se puede deshacer.\n\n¿Continuar?",
    ui.ButtonSet.YES_NO
  );
  if (r2 !== ui.Button.YES) return;

  // Paso 4: eliminar en todas las hojas
  var resumen = [];

  HOJAS_AFECTADAS.forEach(function(nombreHoja) {
    var hoja = ss.getSheetByName(nombreHoja);
    if (!hoja || hoja.getLastRow() < 2) return;

    var borradas;
    if (nombreHoja === "PARTICIPANTES") {
      // En PARTICIPANTES: col A = ID (1-based), col B = Nombre (2-based)
      borradas = _eliminarFilasPorParticipante(hoja, 2, 1, nombreEncontrado, idEncontrado);
    } else if (nombreHoja === "FACTURACION") {
      // FACTURACION: col 1(1-based) = Creamos_ID, col 2(1-based) = Nombre
      borradas = _eliminarFilasPorParticipante(hoja, 2, 1, nombreEncontrado, idEncontrado);
    } else if (nombreHoja === "Cheques") {
      // Cheques: col 3(1-based) = Nombre, sin col ID
      borradas = _eliminarFilasPorParticipante(hoja, 3, -1, nombreEncontrado, idEncontrado);
    } else if (nombreHoja === "Transferencias") {
      // Transferencias: col 1(1-based) = Nombre, sin col ID
      borradas = _eliminarFilasPorParticipante(hoja, 1, -1, nombreEncontrado, idEncontrado);
    } else {
      // DiasEstudio, ListaTerapias, InclusionLaboral, HijosCCI, Bonos, Retiradx, CiclosVida:
      // col 1(1-based) = Creamos_ID, col 2(1-based) = Nombre (o Participante)
      borradas = _eliminarFilasPorParticipante(hoja, 2, 1, nombreEncontrado, idEncontrado);
    }

    if (borradas > 0) resumen.push(nombreHoja + ": " + borradas + " fila(s)");
  });

  var totalBorradas = resumen.reduce(function(s, item) {
    return s + parseInt(item.match(/(\d+)/)[1]);
  }, 0);

  var msg = "✅ Eliminado: " + nombreEncontrado;
  if (idEncontrado) msg += " (" + idEncontrado + ")";
  msg += "\nRegistros borrados: " + totalBorradas + " en " + resumen.length + " hoja(s).";
  if (resumen.length > 0) msg += "\n\nDetalle:\n• " + resumen.join("\n• ");
  _alert(msg);
}); }

/**
 * Elimina filas de una hoja que coincidan con el nombre o ID del participante.
 * Borra de abajo hacia arriba para no desplazar índices.
 *
 * @param {Sheet}  hoja           - Hoja de Google Sheets
 * @param {number} colNombre_1b   - Columna del nombre (1-based)
 * @param {number} colId_1b       - Columna del ID (1-based); -1 si no aplica
 * @param {string} nombre         - Nombre oficial del participante
 * @param {string} id             - Creamos_ID del participante
 * @returns {number}              - Número de filas eliminadas
 */
function _eliminarFilasPorParticipante(hoja, colNombre_1b, colId_1b, nombre, id) {
  if (!hoja || hoja.getLastRow() < 2) return 0;
  var nCols = Math.max(colNombre_1b, colId_1b > 0 ? colId_1b : 0);
  var datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, nCols).getValues();
  var normNombre = textoParaComparar(nombre);
  var idNorm     = id ? id.toUpperCase() : "";

  var aEliminar = [];
  datos.forEach(function(fila, i) {
    var nombreFila = String(fila[colNombre_1b - 1] || "").trim();
    var idFila     = (colId_1b > 0) ? String(fila[colId_1b - 1] || "").trim().toUpperCase() : "";

    var coincideNombre = textoParaComparar(nombreFila) === normNombre;
    var coincideId     = idNorm && idFila === idNorm;

    if (coincideNombre || coincideId) {
      aEliminar.push(i + 2); // +2 por encabezado en fila 1
    }
  });

  // Eliminar de abajo hacia arriba
  aEliminar.reverse().forEach(function(filaReal) {
    hoja.deleteRow(filaReal);
  });

  return aEliminar.length;
}

// ══════════════════════════════════════════════════════════════════
// FEATURE 1 — DASHBOARD VISUAL
// ══════════════════════════════════════════════════════════════════

function actualizarDashboardVisual() { _run(function() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var tz   = CFG.TIMEZONE;
  var ahora = new Date();
  var anio  = ahora.getFullYear();
  var ts    = Utilities.formatDate(ahora, tz, "dd/MM/yyyy HH:mm");

  var NOMBRE_HOJA = "DASHBOARD_VISUAL";
  var hDV = ss.getSheetByName(NOMBRE_HOJA);
  if (!hDV) {
    hDV = ss.insertSheet(NOMBRE_HOJA);
  } else {
    hDV.clearContents();
    hDV.clearFormats();
  }

  // Asegurar columnas suficientes (A-G = 7 cols)
  while (hDV.getMaxColumns() < 7) hDV.insertColumnsAfter(hDV.getMaxColumns(), 1);

  // ── Anchos de columna ────────────────────────────────────────
  // A-F = 130px cada una (6 cols de KPI), G = 90px (botón)
  hDV.setColumnWidth(1, 130);
  hDV.setColumnWidth(2, 130);
  hDV.setColumnWidth(3, 130);
  hDV.setColumnWidth(4, 130);
  hDV.setColumnWidth(5, 130);
  hDV.setColumnWidth(6, 130);
  hDV.setColumnWidth(7, 90);

  // ── Calcular KPIs ────────────────────────────────────────────

  // Participantes activas (Etapa = "Inscritx") + categorías + forma de pago
  var kpiActivas = 0;
  var catCount = { A: 0, B: 0, C: 0, D: 0 };
  var pagoCheque = 0, pagoTransf = 0;
  var hP = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES);
  if (hP && hP.getLastRow() > 1) {
    hP.getRange(2, 1, hP.getLastRow() - 1, 22).getValues().forEach(function(r) {
      var etapa = String(r[8]  || "").trim();                  // col I = Etapa (idx 8)
      var cat   = String(r[12] || "").trim().toUpperCase();    // col M = Categoria (idx 12)
      var pago  = String(r[21] || "").trim();                  // col V = Forma_Pago (idx 21)
      if (etapa === "Inscritx") {
        kpiActivas++;
        if (catCount[cat] !== undefined) catCount[cat]++;
        if (pago === "Transferencia") pagoTransf++;
        else if (pago === "Cheque") pagoCheque++;
      }
    });
  }

  // Retiradx este año
  var kpiRetiradx = 0;
  var hRet = ss.getSheetByName("Retiradx");
  if (hRet && hRet.getLastRow() > 1) {
    hRet.getRange(2, 1, hRet.getLastRow() - 1, 1).getValues().forEach(function(r) {
      var f = r[0];
      var d = (f instanceof Date) ? f : new Date(f);
      if (!isNaN(d) && d.getFullYear() === anio) kpiRetiradx++;
    });
  }

  // Ciclos de Vida completados este año
  var kpiCiclos = 0;
  var hCV = ss.getSheetByName("CiclosVida");
  if (hCV && hCV.getLastRow() > 1) {
    hCV.getRange(2, 1, hCV.getLastRow() - 1, 1).getValues().forEach(function(r) {
      var f = r[0];
      var d = (f instanceof Date) ? f : new Date(f);
      if (!isNaN(d) && d.getFullYear() === anio) kpiCiclos++;
    });
  }

  // Indicadores financieros desde FACTURACION
  var kpiPromMensual = 0, kpiTotalAnio = 0;
  var hojaF = ss.getSheetByName(CFG.HOJAS.FACTURACION);
  if (hojaF && hojaF.getLastRow() > 1) {
    var datF = hojaF.getDataRange().getValues().slice(1);
    var baseParMes = {};
    var totalNetoAnio = 0;
    datF.forEach(function(f) {
      if (Number(f[3]) !== anio) return;
      var mesIdx = CFG.MESES.indexOf(String(f[2]).trim());
      var base = parseFloat(f[9]) || 0;   // col index 9 = Base
      var neto = parseFloat(f[13]) || 0;  // col index 13 = Neto
      var hrs  = parseFloat(f[5]) || 0;   // col index 5 = HorasTrab
      totalNetoAnio += neto;
      if (mesIdx >= 0 && hrs > 0) {
        baseParMes[mesIdx] = (baseParMes[mesIdx] || 0) + base;
      }
    });
    kpiTotalAnio = totalNetoAnio;
    var mesesConDatos = Object.keys(baseParMes).length;
    if (mesesConDatos > 0) {
      var totalBase = Object.keys(baseParMes).reduce(function(s, k) { return s + baseParMes[k]; }, 0);
      kpiPromMensual = totalBase / mesesConDatos;
    }
  }

  // Horas formación este año — DatosKobo study+therapy days × CFG.HORAS_JORNADA_NORMAL
  var kpiHrsFormacion = 0;
  var hK = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (hK && hK.getLastRow() > 1) {
    try {
      var encK = hK.getRange(1, 1, 1, hK.getLastColumn()).getValues()[0];
      var rawK = hK.getRange(2, 1, hK.getLastRow() - 1, hK.getLastColumn()).getValues();
      var colsK = detectarColumnas(encK, rawK.slice(0, 50));
      var diasEstMap  = obtenerDiasEstudio();
      var terapiasMap = obtenerListaTerapias();
      var mapeoN = cargarMapeoNombres();
      var diasFormacion = {};
      rawK.forEach(function(fila) {
        var ts2 = _resolverTsKobo(fila, colsK);
        if (!ts2 || isNaN(ts2) || ts2.getFullYear() !== anio) return;
        var nombreRaw = obtenerParticipanteFila(fila, colsK);
        if (!nombreRaw) return;
        var nombre = normalizarNombre(nombreRaw, mapeoN) || limpiarNombre(nombreRaw);
        if (!nombre) return;
        var esDE  = esDiaDeEstudio(nombre, ts2, diasEstMap);
        var esTer = !!terapiasMap[nombre];
        if (!esDE && !esTer) return;
        diasFormacion[nombre + "|" + _dClave(ts2)] = true;
      });
      kpiHrsFormacion = Object.keys(diasFormacion).length * CFG.HORAS_JORNADA_NORMAL;
    } catch(e) { Logger.log("DashVisual KPI hrs formacion: " + e.message); }
  }

  // Promedio horas laborales mensuales — FACTURACION col 5 (HorasTrab)
  var kpiPromHrsMes = 0;
  if (hojaF && hojaF.getLastRow() > 1) {
    var totalHrs = 0, cuentaPartMes = 0;
    hojaF.getDataRange().getValues().slice(1).forEach(function(f) {
      if (Number(f[3]) !== anio) return;
      var hrs = parseFloat(f[5]) || 0;
      if (hrs > 0) { totalHrs += hrs; cuentaPartMes++; }
    });
    if (cuentaPartMes > 0) kpiPromHrsMes = totalHrs / cuentaPartMes;
  }

  // ── Helpers de layout ────────────────────────────────────────

  /**
   * Bloque KPI que ocupa nCols columnas, comenzando en colInicio.
   * Fila fila   = label (altura 28px)
   * Fila fila+1 = valor (altura 55px, font 32)
   */
  function kpiBlock(fila, colInicio, nCols, label, valor, bgColor, fgColor) {
    var rLabel = hDV.getRange(fila,     colInicio, 1, nCols);
    var rVal   = hDV.getRange(fila + 1, colInicio, 1, nCols);
    rLabel.merge().setValue(label)
      .setBackground(bgColor).setFontColor(fgColor)
      .setFontWeight("bold").setFontSize(10)
      .setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setWrap(true);
    hDV.setRowHeight(fila, 28);
    rVal.merge().setValue(valor)
      .setBackground(bgColor).setFontColor(fgColor)
      .setFontWeight("bold").setFontSize(32)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    hDV.setRowHeight(fila + 1, 55);
    rVal.setBorder(true, true, true, true, null, null, "#ffffff", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }

  /** Fila de título de sección — siempre merge cols 1-6 */
  function seccionTitle(fila, texto, bgColor) {
    hDV.getRange(fila, 1, 1, 6).merge()
      .setValue(texto)
      .setBackground(bgColor).setFontColor("#ffffff")
      .setFontWeight("bold").setFontSize(11)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    hDV.setRowHeight(fila, 32);
  }

  /** Fila espaciadora entre secciones — height 10px, blanco */
  function spacerRow(fila) {
    hDV.getRange(fila, 1, 1, 7).setBackground("#ffffff");
    hDV.setRowHeight(fila, 10);
  }

  // ── Fila 1: Título principal — merge A-G ────────────────────
  hDV.getRange(1, 1, 1, 7).merge()
    .setValue("📊 DASHBOARD — mi eelo | Textil")
    .setBackground("#1a237e").setFontColor("#ffffff")
    .setFontWeight("bold").setFontSize(15)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  hDV.setRowHeight(1, 48);

  // ── Fila 2: Fecha actualización (A-F) + botón (G) ───────────
  hDV.getRange(2, 1, 1, 6).merge()
    .setValue("Última actualización: " + ts)
    .setBackground("#eceff1").setFontColor("#546e7a")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontStyle("italic");
  hDV.getRange(2, 7)
    .setValue("▶ Actualizar")
    .setBackground("#1565c0").setFontColor("#ffffff")
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setNote("Para asignar: clic derecho en esta celda → Asignar macro → actualizarDashboardVisual");
  hDV.setRowHeight(2, 30);

  // ── Fila 3: Espaciador ───────────────────────────────────────
  spacerRow(3);

  // ── SECCIÓN 1: RESUMEN GENERAL — 3 KPIs, 2 cols cada uno (filas 4-6) ──
  seccionTitle(4, "👥 RESUMEN GENERAL", "#37474f");
  // Block1: cols 1-2 (A-B)
  kpiBlock(5, 1, 2, "Participantes Activas", kpiActivas, "#1565c0", "#ffffff");
  // Block2: cols 3-4 (C-D)
  kpiBlock(5, 3, 2, "Retiradx Este Año", kpiRetiradx, "#b71c1c", "#ffffff");
  // Block3: cols 5-6 (E-F)
  kpiBlock(5, 5, 2, "Ciclos de Vida Este Año", kpiCiclos, "#1b5e20", "#ffffff");
  spacerRow(7);

  // ── SECCIÓN 2: INDICADORES FINANCIEROS — 2 KPIs, 3 cols cada uno (filas 8-10) ──
  seccionTitle(8, "💰 INDICADORES FINANCIEROS", "#004d40");
  // Block1: cols 1-3 (A-C)
  kpiBlock(9, 1, 3, "Monto Promedio Mensual (Q)", "Q " + kpiPromMensual.toFixed(2), "#00695c", "#ffffff");
  // Block2: cols 4-6 (D-F)
  kpiBlock(9, 4, 3, "Total Pagado Este Año (Q)", "Q " + kpiTotalAnio.toFixed(2), "#2e7d32", "#ffffff");
  spacerRow(11);

  // ── SECCIÓN 3: FORMACIÓN Y HORAS — 2 KPIs, 3 cols cada uno (filas 12-14) ──
  seccionTitle(12, "📚 FORMACIÓN Y HORAS", "#4a148c");
  // Block1: cols 1-3 (A-C)
  kpiBlock(13, 1, 3, "Horas Formación Este Año", kpiHrsFormacion > 0 ? kpiHrsFormacion.toFixed(0) + " hrs" : "0 hrs", "#bf360c", "#ffffff");
  // Block2: cols 4-6 (D-F)
  kpiBlock(13, 4, 3, "Promedio Hrs Laborales/Mes", kpiPromHrsMes > 0 ? kpiPromHrsMes.toFixed(1) + " hrs" : "0.0 hrs", "#e65100", "#ffffff");
  spacerRow(15);

  // ── SECCIÓN 4: PARTICIPANTES POR CATEGORÍA — 4 KPIs, 1 col cada uno (filas 16-18) ──
  seccionTitle(16, "🏷️ PARTICIPANTES POR CATEGORÍA", "#263238");
  var catCfg = [
    { cat: "A", bg: CFG.COLORES_CAT.A.bg, fg: CFG.COLORES_CAT.A.fg },
    { cat: "B", bg: CFG.COLORES_CAT.B.bg, fg: CFG.COLORES_CAT.B.fg },
    { cat: "C", bg: CFG.COLORES_CAT.C.bg, fg: CFG.COLORES_CAT.C.fg },
    { cat: "D", bg: CFG.COLORES_CAT.D.bg, fg: CFG.COLORES_CAT.D.fg }
  ];
  // 4 KPIs, each 1 col wide, cols 1-4 (A-D); cols E-F left white
  catCfg.forEach(function(cfg, idx) {
    var colI = idx + 1;
    hDV.getRange(17, colI)
      .setValue("Cat. " + cfg.cat)
      .setBackground(cfg.bg).setFontColor(cfg.fg)
      .setFontWeight("bold").setFontSize(10)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    hDV.getRange(18, colI)
      .setValue(catCount[cfg.cat])
      .setBackground(cfg.bg).setFontColor(cfg.fg)
      .setFontWeight("bold").setFontSize(32)
      .setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setBorder(true, true, true, true, null, null, "#ffffff", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    hDV.setRowHeight(17, 28);
    hDV.setRowHeight(18, 55);
  });
  // Cols E-F (5-6) blanco en filas 17-18
  hDV.getRange(17, 5, 2, 2).setBackground("#ffffff");
  spacerRow(19);

  // ── SECCIÓN 5: FORMA DE PAGO — 2 KPIs, 3 cols cada uno (filas 20-22) ──
  seccionTitle(20, "💳 FORMA DE PAGO", "#1a237e");
  // Block1: cols 1-3 (A-C)
  kpiBlock(21, 1, 3, "Transferencias", pagoTransf, "#1565c0", "#ffffff");
  // Block2: cols 4-6 (D-F)
  kpiBlock(21, 4, 3, "Cheques", pagoCheque, "#37474f", "#ffffff");
  spacerRow(23);

  // ── Final ────────────────────────────────────────────────────
  hDV.setHiddenGridlines(true);
  hDV.setTabColor("#1a237e");
  hDV.setFrozenRows(1);

  ss.toast("✅ Dashboard Visual actualizado", "📊", 4);
}); }

// ══════════════════════════════════════════════════════════════════
// FEATURE 2 — EXPORTAR PARA POWERBI
// ══════════════════════════════════════════════════════════════════

function exportarParaPowerBI() { _run(function() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var tz    = CFG.TIMEZONE;
  var hoy   = new Date();

  var NOMBRE_HOJA = "PowerBI_Export";
  var hExp = ss.getSheetByName(NOMBRE_HOJA);
  if (!hExp) {
    hExp = ss.insertSheet(NOMBRE_HOJA);
  } else {
    hExp.clearContents();
    hExp.clearFormats();
  }

  // Encabezados
  var HEADERS = [
    "Creamos_ID","Nombre","Proyecto","Programa","Etapa","Categoria",
    "Tarifa_Hora","Tiene_Factura","Banco","Forma_Pago","Hijos_CCI",
    "Mes","Año","Quincena","Horas_Trabajadas","Monto_Base",
    "IVA","Total_Org_Paga","Neto_Participante","Estipendio","Bono","Fecha_Export"
  ];

  // Leer PARTICIPANTES
  var hP = ss.getSheetByName(CFG.HOJAS.PARTICIPANTES);
  var mapaParticipantes = {}; // nombre → {id,proyecto,programa,etapa,categoria,tarifa,tieneFactura,banco,formaPago,estipendio}
  if (hP && hP.getLastRow() > 1) {
    hP.getRange(2, 1, hP.getLastRow() - 1, 19).getValues().forEach(function(r) {
      var nombre = String(r[1] || "").trim();
      if (!nombre) return;
      mapaParticipantes[nombre] = {
        id:           String(r[0]  || "").trim(),
        proyecto:     String(r[2]  || "").trim(),
        programa:     String(r[3]  || "").trim(),
        etapa:        String(r[4]  || "").trim(),
        categoria:    String(r[8]  || "").trim(),
        tarifa:       parseFloat(r[9])  || 0,
        tieneFactura: String(r[10] || "").trim(),
        banco:        String(r[14] || "").trim(),
        formaPago:    String(r[17] || "").trim(),
        estipendio:   0
      };
    });
  }

  // Leer FACTURACION
  var hojaF = ss.getSheetByName(CFG.HOJAS.FACTURACION);
  var filasExport = [];
  var nombresConFacturacion = {};

  if (hojaF && hojaF.getLastRow() > 1) {
    // Leer bonos
    var mapaBonosF = {};
    var hBonos = ss.getSheetByName("Bonos");
    if (hBonos && hBonos.getLastRow() > 1) {
      hBonos.getRange(2, 1, hBonos.getLastRow() - 1, hBonos.getLastColumn()).getValues().forEach(function(b) {
        var nombre = String(b[2] || "").trim();
        var monto  = parseFloat(b[3]) || 0;
        if (nombre && monto) {
          mapaBonosF[nombre] = (mapaBonosF[nombre] || 0) + monto;
        }
      });
    }

    hojaF.getDataRange().getValues().slice(1).forEach(function(f) {
      var nombre = String(f[1] || "").trim();
      if (!nombre) return;
      nombresConFacturacion[nombre] = true;

      var p = mapaParticipantes[nombre] || {
        id: String(f[0] || "").trim(),
        proyecto: "", programa: "", etapa: "", categoria: "",
        tarifa: parseFloat(f[8]) || 0, tieneFactura: String(f[10] || ""),
        banco: "", formaPago: "", hijosCCI: "", estipendio: 0
      };

      var mes      = String(f[2] || "").trim();
      var anioFila = String(f[3] || "").trim();
      var quincena = String(f[4] || "").trim();
      var hrsTrab  = parseFloat(f[5]) || 0;
      var base     = parseFloat(f[9]) || 0;
      var iva      = parseFloat(f[11]) || 0;
      var totalOrg = parseFloat(f[12]) || 0;
      var neto     = parseFloat(f[13]) || 0;
      var bono     = mapaBonosF[nombre] || 0;

      filasExport.push([
        p.id, nombre, p.proyecto, p.programa, p.etapa, p.categoria,
        p.tarifa, p.tieneFactura, p.banco, p.formaPago, p.hijosCCI,
        mes, anioFila, quincena, hrsTrab, base,
        iva, totalOrg, neto, p.estipendio, bono, hoy
      ]);
    });
  }

  // Participantes sin facturación — incluir con valores 0
  Object.keys(mapaParticipantes).forEach(function(nombre) {
    if (nombresConFacturacion[nombre]) return;
    var p = mapaParticipantes[nombre];
    filasExport.push([
      p.id, nombre, p.proyecto, p.programa, p.etapa, p.categoria,
      p.tarifa, p.tieneFactura, p.banco, p.formaPago, p.hijosCCI,
      "", "", "", 0, 0, 0, 0, 0, p.estipendio, 0, hoy
    ]);
  });

  // Escribir en hoja
  hExp.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  if (filasExport.length > 0) {
    hExp.getRange(2, 1, filasExport.length, HEADERS.length).setValues(filasExport);
  }

  // Formato encabezado
  hExp.getRange(1, 1, 1, HEADERS.length)
    .setBackground("#1a237e").setFontColor("#ffffff")
    .setFontWeight("bold");
  hExp.setFrozenRows(1);

  // Formato columnas de fecha (col 22 = Fecha_Export)
  if (filasExport.length > 0) {
    hExp.getRange(2, 22, filasExport.length, 1).setNumberFormat("dd/MM/yyyy");
  }

  // Formato columnas Q (cols 7=Tarifa, 16=Base, 17=IVA, 18=Total, 19=Neto, 20=Estipendio, 21=Bono)
  var colsQ = [7, 16, 17, 18, 19, 20, 21];
  if (filasExport.length > 0) {
    colsQ.forEach(function(c) {
      hExp.getRange(2, c, filasExport.length, 1).setNumberFormat('"Q"#,##0.00');
    });
  }

  // Auto-resize columnas
  for (var ci = 1; ci <= HEADERS.length; ci++) {
    hExp.autoResizeColumn(ci);
  }

  var totalFilas = filasExport.length;
  ss.toast("✅ PowerBI_Export actualizado — " + totalFilas + " filas", "📤", 5);
}); }

// ══════════════════════════════════════════════════════════════════
// FEATURE 3 — GUÍA DE USO
// ══════════════════════════════════════════════════════════════════

function crearGuiaUso() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var NOMBRE_HOJA = "Guía de Uso";
  var h = ss.getSheetByName(NOMBRE_HOJA);
  if (!h) {
    h = ss.insertSheet(NOMBRE_HOJA);
  } else {
    h.clearContents();
    h.clearFormats();
  }

  // Asegurar 4 columnas
  while (h.getMaxColumns() < 4) h.insertColumnsAfter(h.getMaxColumns(), 1);

  // ── Anchos de columna ────────────────────────────────────────
  h.setColumnWidth(1, 40);
  h.setColumnWidth(2, 260);
  h.setColumnWidth(3, 420);
  h.setColumnWidth(4, 180);

  // ── Fecha actual para el pie ──────────────────────────────────
  var tz = CFG.TIMEZONE;
  var fechaGuia = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy");

  // Acumulador de filas: [col A, col B, col C, col D]
  var filas = [];
  var tipos = [];

  function push(f, t) { filas.push(f); tipos.push(t); }

  // ── Fila 1: Título principal ──────────────────────────────────
  push(["📋 GUÍA DE PAGOS — Sistema RRHH mi eelo", "", "", ""], "titulo");

  // ── Fila 2: Banner para quién es ─────────────────────────────
  push(["👤 Esta guía es para la persona encargada de registrar horas y gestionar los pagos quincenales.", "", "", ""], "banner");

  // ── Fila 3: Espacio ───────────────────────────────────────────
  push(["", "", "", ""], "vacio");

  // ══════════════════════════════════════════════════════════════
  // SECCIÓN 1 — FLUJO COMPLETO DE PAGOS
  // ══════════════════════════════════════════════════════════════
  push(["🗓️ CADA QUINCENA: FLUJO COMPLETO DE PAGOS", "", "", ""], "seccion");
  push(["Paso", "Acción", "Cómo hacerlo", "Qué verificar"], "enc_tabla");
  push(["1", "Importar las horas desde Kobo",
        "Menú 📥 Asistencia → Importar desde Kobo",
        "Que aparezca el mensaje \"X registros importados\". Si da error, esperar 5 minutos y repetir."], "fila");
  push(["2", "Calcular el reporte de quincena",
        "Menú 📅 Quincena → Ver / actualizar quincena actual",
        "Se abre una hoja nueva con el período. Revisa que aparezcan todas las participantes activas."], "fila");
  push(["3", "Revisar el reporte de quincena",
        "Abrir la hoja del período (ej: \"26/05 al 04/06\")",
        "Verificar: (a) que no haya personas duplicadas, (b) que las horas sean razonables (no más de 9 hrs/día), (c) que los montos en \"Neto part.\" sean correctos."], "fila");
  push(["4", "Registrar los pagos",
        "Menú 💳 Quincena → Registrar pagos de quincena",
        "Seleccionar el número de la quincena en la lista. Confirmar."], "fila");
  push(["5", "Verificar hoja Cheques",
        "Abrir hoja \"Cheques\"",
        "Deben aparecer las participantes que tienen Forma_Pago = Cheque. El número de cheque (col B) se llena a mano."], "fila");
  push(["6", "Verificar hoja Transferencias",
        "Abrir hoja \"Transferencias\"",
        "Deben aparecer las participantes con Transferencia. Si es Q1, se llena columna Q1. Si es Q2, se llena Q2 y se calcula el Total automáticamente."], "fila");
  push(["7", "Cerrar la quincena",
        "Menú 📅 Quincena → Cerrar quincena y abrir siguiente",
        "Confirmar el cierre. Se archiva la quincena actual y se prepara la siguiente."], "fila");

  push(["", "", "", ""], "vacio");

  // ══════════════════════════════════════════════════════════════
  // SECCIÓN 2 — PROBLEMAS COMUNES
  // ══════════════════════════════════════════════════════════════
  push(["⚠️ PROBLEMAS COMUNES Y CÓMO RESOLVERLOS", "", "", ""], "seccion_rojo");
  push(["Paso", "Problema", "Causa más probable", "Solución"], "enc_tabla");
  push(["—", "Una participante no aparece en el reporte",
        "Su nombre en Kobo no coincide con PARTICIPANTES",
        "Admin → 📲 Sincronizar IDs desde DatosKobo. Si persiste: Admin → 🔍 Diagnosticar IDs."], "fila");
  push(["—", "Aparecen personas duplicadas",
        "Hay filas repetidas en PARTICIPANTES",
        "Admin → 🧹 Limpiar duplicados en PARTICIPANTES. Luego recalcular."], "fila");
  push(["—", "Las horas parecen incorrectas",
        "Registro de Kobo tiene entradas sin salida",
        "Admin → 🔧 Reparar datos Kobo. Luego reimportar."], "fila");
  push(["—", "El botón \"Registrar pagos\" dice \"no encontrado en PARTICIPANTES\"",
        "La participante no tiene Forma_Pago configurada",
        "Ir a PARTICIPANTES → col T → escribir \"Transferencia\" o \"Cheque\" para esa persona."], "fila");
  push(["—", "Una participante tiene Forma_Pago pero no llega a Cheques/Transferencias",
        "El nombre en el reporte difiere del nombre en PARTICIPANTES",
        "Admin → 🔄 Sincronizar desde Creamos DB. O asegurarse que el Creamos ID esté en col A de PARTICIPANTES."], "fila");
  push(["—", "Error \"Hoja no encontrada\"",
        "El sistema no está instalado completo",
        "Menú 👥 RRHH → 🚀 Instalación completa."], "fila");

  push(["", "", "", ""], "vacio");

  // ══════════════════════════════════════════════════════════════
  // SECCIÓN 3 — CONFIGURACIÓN INICIAL
  // ══════════════════════════════════════════════════════════════
  push(["📋 ANTES DE EMPEZAR: CONFIGURACIÓN INICIAL", "", "", ""], "seccion_verde");
  push(["Paso", "Acción", "Cómo hacerlo", "Notas"], "enc_tabla");
  push(["1", "Instalar el sistema",
        "Menú 👥 RRHH → 🚀 Instalación completa",
        "Solo se hace una vez. Crea todas las hojas automáticamente."], "fila");
  push(["2", "Cargar las 33 participantes",
        "Admin → 📋 Cargar lista oficial",
        "Carga los datos base de todas las participantes del programa."], "fila");
  push(["3", "Sincronizar los IDs",
        "Admin → 🔄 Sincronizar desde Creamos DB",
        "Busca el Creamos ID de cada participante. Puede tardar 1-2 minutos."], "fila");
  push(["4", "Configurar forma de pago",
        "Abrir hoja PARTICIPANTES → columna T (Forma_Pago)",
        "Para cada participante, elegir \"Transferencia\" o \"Cheque\" del desplegable."], "fila");
  push(["5", "Configurar datos bancarios",
        "En PARTICIPANTES, cols Q (Banco), R (Tipo Cuenta), S (Número de Cuenta)",
        "Llenar los datos bancarios de cada participante para las transferencias."], "fila");
  push(["6", "Configurar quincena",
        "Menú 📅 Quincena → Configurar nueva quincena",
        "Ingresar la fecha de inicio y fin del primer período de pago."], "fila");
  push(["7", "Activar automatizaciones",
        "Admin → ⚡ Activar automatizaciones",
        "Activa la importación automática de Kobo cada hora. Solo se hace una vez."], "fila");

  push(["", "", "", ""], "vacio");

  // ══════════════════════════════════════════════════════════════
  // SECCIÓN 4 — DATOS IMPORTANTES
  // ══════════════════════════════════════════════════════════════
  push(["💡 DATOS IMPORTANTES DEL SISTEMA", "", "", ""], "seccion_gris");
  push(["Dato", "Información", "Detalle", "Nota"], "enc_tabla");
  push(["Tarifa A", "Q16.50 / hora", "Categoría A — nivel más avanzado", "Ver col K en PARTICIPANTES"], "fila");
  push(["Tarifa B", "Q15.75 / hora", "Categoría B", ""], "fila");
  push(["Tarifa C", "Q15.00 / hora", "Categoría C — nivel estándar", ""], "fila");
  push(["Tarifa D", "Q14.00 / hora", "Categoría D — nivel inicial", ""], "fila");
  push(["Col T en PARTICIPANTES", "Forma de Pago",
        "\"Transferencia\" o \"Cheque\" — sin esto, la participante NO aparece en ninguna hoja de pago.",
        "Llenar para todas antes de registrar pagos."], "fila");
  push(["Col V en PARTICIPANTES", "Estipendio",
        "Monto fijo adicional que recibe cada quincena además de sus horas. Si no tiene, dejar en 0.",
        "Se suma automáticamente al calcular la quincena."], "fila");
  push(["Días de Estudio", "Horas NO pagadas",
        "Los días marcados en la hoja DiasEstudio no se cuentan como horas de trabajo.",
        "Verificar que estén correctamente configurados."], "fila");
  push(["Terapias", "Horas NO pagadas",
        "Los días de terapia tampoco cuentan como horas de trabajo.",
        "Igual que días de estudio."], "fila");
  push(["IVA 5%", "Factura pequeño contribuyente",
        "Solo se aplica a quien tiene \"Sí\" en col K (Tiene_Factura) de PARTICIPANTES.",
        "Admin → Configurar IVA para cambiar esto."], "fila");

  push(["", "", "", ""], "vacio");

  // ══════════════════════════════════════════════════════════════
  // SECCIÓN 5 — GLOSARIO RÁPIDO
  // ══════════════════════════════════════════════════════════════
  push(["📞 GLOSARIO RÁPIDO", "", "", ""], "seccion_morado");
  push(["Término", "Significado", "—", "—"], "enc_tabla");
  push(["Quincena",       "Período de pago. Normalmente del 1 al 15 o del 16 al fin de mes.", "—", "—"], "fila");
  push(["Creamos ID",     "Código único de cada participante en el sistema Creamos/Salesforce. Ej: MACI030373", "—", "—"], "fila");
  push(["Q1 / Q2",        "Primera o segunda quincena del mes", "—", "—"], "fila");
  push(["Kobo / DatosKobo", "Sistema donde las participantes registran su entrada y salida. Los datos se importan automáticamente.", "—", "—"], "fila");
  push(["Etapa \"Inscritx\"", "Participante activa en el programa", "—", "—"], "fila");
  push(["Etapa \"Retiradx\"", "Dejó el programa. Al cambiar a esta etapa, el sistema pide la razón.", "—", "—"], "fila");
  push(["IVA 5%",         "Impuesto que se descuenta a quien emite factura", "—", "—"], "fila");
  push(["Neto participante", "Lo que realmente recibe la participante después de IVA", "—", "—"], "fila");
  push(["Monto Base",     "Horas trabajadas × tarifa por hora", "—", "—"], "fila");

  push(["", "", "", ""], "vacio");

  // ── Pie: fecha de actualización ───────────────────────────────
  push(["📅 Última actualización de esta guía: " + fechaGuia, "", "", ""], "pie");

  // ── Escribir todas las filas ──────────────────────────────────
  h.getRange(1, 1, filas.length, 4).setValues(filas);

  // ── Activar wrap de texto en toda la hoja ────────────────────
  h.getRange(1, 1, filas.length, 4).setWrap(true);

  // ── Aplicar formatos fila a fila ─────────────────────────────
  tipos.forEach(function(tipo, i) {
    var fila1 = i + 1;
    var rAll  = h.getRange(fila1, 1, 1, 4);

    if (tipo === "titulo") {
      h.getRange(fila1, 1, 1, 4).merge()
        .setBackground("#1a237e").setFontColor("#ffffff")
        .setFontWeight("bold").setFontSize(15)
        .setHorizontalAlignment("center").setVerticalAlignment("middle");
      h.setRowHeight(fila1, 44);

    } else if (tipo === "banner") {
      h.getRange(fila1, 1, 1, 4).merge()
        .setBackground("#e8f5e9").setFontColor("#1b5e20")
        .setFontWeight("bold").setFontSize(11)
        .setHorizontalAlignment("center").setVerticalAlignment("middle");
      h.setRowHeight(fila1, 35);

    } else if (tipo === "seccion") {
      h.getRange(fila1, 1, 1, 4).merge()
        .setBackground("#1a237e").setFontColor("#ffffff")
        .setFontWeight("bold").setFontSize(12)
        .setHorizontalAlignment("left").setVerticalAlignment("middle");
      h.setRowHeight(fila1, 34);

    } else if (tipo === "seccion_rojo") {
      h.getRange(fila1, 1, 1, 4).merge()
        .setBackground("#b71c1c").setFontColor("#ffffff")
        .setFontWeight("bold").setFontSize(12)
        .setHorizontalAlignment("left").setVerticalAlignment("middle");
      h.setRowHeight(fila1, 34);

    } else if (tipo === "seccion_verde") {
      h.getRange(fila1, 1, 1, 4).merge()
        .setBackground("#1b5e20").setFontColor("#ffffff")
        .setFontWeight("bold").setFontSize(12)
        .setHorizontalAlignment("left").setVerticalAlignment("middle");
      h.setRowHeight(fila1, 34);

    } else if (tipo === "seccion_gris") {
      h.getRange(fila1, 1, 1, 4).merge()
        .setBackground("#37474f").setFontColor("#ffffff")
        .setFontWeight("bold").setFontSize(12)
        .setHorizontalAlignment("left").setVerticalAlignment("middle");
      h.setRowHeight(fila1, 34);

    } else if (tipo === "seccion_morado") {
      h.getRange(fila1, 1, 1, 4).merge()
        .setBackground("#4a148c").setFontColor("#ffffff")
        .setFontWeight("bold").setFontSize(12)
        .setHorizontalAlignment("left").setVerticalAlignment("middle");
      h.setRowHeight(fila1, 34);

    } else if (tipo === "enc_tabla") {
      rAll.setBackground("#90a4ae").setFontColor("#ffffff")
        .setFontWeight("bold").setFontSize(10)
        .setHorizontalAlignment("left").setVerticalAlignment("middle");
      h.getRange(fila1, 1).setHorizontalAlignment("center");
      h.setRowHeight(fila1, 24);

    } else if (tipo === "fila") {
      rAll.setBackground("#ffffff").setFontColor("#212121").setFontSize(10)
        .setVerticalAlignment("top");
      h.getRange(fila1, 1).setHorizontalAlignment("center").setFontWeight("bold");
      rAll.setBorder(null, null, true, null, null, null, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
      h.setRowHeight(fila1, 52);

    } else if (tipo === "vacio") {
      rAll.setBackground("#f9f9f9");
      h.setRowHeight(fila1, 8);

    } else if (tipo === "pie") {
      h.getRange(fila1, 1, 1, 4).merge()
        .setBackground("#ffffff").setFontColor("#9e9e9e")
        .setFontStyle("italic").setFontSize(10)
        .setHorizontalAlignment("center").setVerticalAlignment("middle");
      h.setRowHeight(fila1, 28);
    }
  });

  // ── Congelar fila 1 ──────────────────────────────────────────
  h.setFrozenRows(1);
  ss.setActiveSheet(h);
  ss.toast("✅ Guía de Pagos creada/actualizada", "📋", 4);
}); }

// ══════════════════════════════════════════════════════════════════
// DIAGNÓSTICO DE HORAS
// ══════════════════════════════════════════════════════════════════
function diagnosticarCalculoHoras() { _run(function() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var lin = [];
  var ok  = true;

  // ── 1. ¿Hay período activo? ───────────────────────────────────
  var periodo = _periodoActivo();
  if (!periodo) {
    lin.push("❌ PASO 1 — NO HAY PERÍODO ACTIVO en hoja PERIODOS.");
    lin.push("   → Menú 📅 Quincena → 🗓️ Nueva quincena (manual)");
    lin.push("   → Pon las fechas de la quincena actual y ponla como 'Activo'.");
    ok = false;
  } else {
    var tz = CFG.TIMEZONE;
    lin.push("✅ PASO 1 — Período activo: " + periodo.label +
             "\n   Hoja reporte: '" + periodo.tab + "'");
  }

  // ── 2. ¿DatosKobo tiene datos? ───────────────────────────────
  var hK = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (!hK || hK.getLastRow() < 2) {
    lin.push("\n❌ PASO 2 — DatosKobo está VACÍO o no existe.");
    lin.push("   → Menú 📥 Asistencia → 📥 Importar desde Kobo");
    ok = false;
  } else {
    var nRegistros = hK.getLastRow() - 1;
    lin.push("\n✅ PASO 2 — DatosKobo tiene " + nRegistros + " registros.");
  }

  // ── 3. ¿Se detectan columnas correctamente? ──────────────────
  if (hK && hK.getLastRow() >= 2) {
    var enc = hK.getRange(1, 1, 1, hK.getLastColumn()).getValues()[0];
    var raw50 = hK.getRange(2, 1, Math.min(50, hK.getLastRow()-1), hK.getLastColumn()).getValues();
    var cols = detectarColumnas(enc, raw50);

    var colInfo = [];
    if (cols.start          !== undefined) colInfo.push("start(col "+(cols.start+1)+")");
    if (cols.end            !== undefined) colInfo.push("end(col "+(cols.end+1)+")");
    if (cols.participante   !== undefined) colInfo.push("participante(col "+(cols.participante+1)+")");
    if (cols.accionUnificada!== undefined) colInfo.push("acción(col "+(cols.accionUnificada+1)+")");
    else if (cols.ingreso   !== undefined) colInfo.push("ingreso(col "+(cols.ingreso+1)+")");

    if (cols.participante === undefined) {
      lin.push("\n❌ PASO 3 — No se detectó columna de PARTICIPANTE en DatosKobo.");
      lin.push("   Encabezados encontrados: " + enc.slice(0,10).join(", "));
      ok = false;
    } else if (cols.accionUnificada === undefined && cols.ingreso === undefined) {
      lin.push("\n❌ PASO 3 — No se detectó columna de ACCIÓN (Entrada/Salida).");
      lin.push("   Encabezados: " + enc.slice(0,10).join(", "));
      ok = false;
    } else {
      lin.push("\n✅ PASO 3 — Columnas detectadas: " + colInfo.join(", "));
    }

    // ── 4. Muestra de registros en período activo ─────────────
    if (periodo && ok) {
      var fi = new Date(periodo.fi);
      var ff = new Date(periodo.ff);
      var dIni = new Date(fi.getFullYear(), fi.getMonth(), fi.getDate());
      var dFin = new Date(ff.getFullYear(), ff.getMonth(), ff.getDate());
      var raw = hK.getRange(2, 1, hK.getLastRow()-1, hK.getLastColumn()).getValues();
      var mapeo = cargarMapeoNombres();

      var enPeriodo = 0, entradas = 0, salidas = 0, sinNombre = 0;
      var nombresVis = {};
      raw.forEach(function(fila) {
        var ts = _resolverTsKobo(fila, cols);
        if (!ts || isNaN(ts)) return;
        var dia = new Date(ts.getFullYear(), ts.getMonth(), ts.getDate());
        if (dia < dIni || dia > dFin) return;
        enPeriodo++;
        var nRaw = obtenerParticipanteFila(fila, cols);
        if (!nRaw) { sinNombre++; return; }
        var nombre = normalizarNombre(nRaw, mapeo);
        nombresVis[nombre] = true;
        var tipo = obtenerTipoRegistro(fila, cols);
        if (tipo.esIngreso) entradas++;
        if (tipo.esEgreso)  salidas++;
      });

      lin.push("\n✅ PASO 4 — En el período activo:");
      lin.push("   Registros: " + enPeriodo + " (entradas: " + entradas + ", salidas: " + salidas + ")");
      lin.push("   Participantes con datos: " + Object.keys(nombresVis).length);
      if (sinNombre > 0) lin.push("   ⚠️ " + sinNombre + " registros sin nombre de participante");

      if (enPeriodo === 0) {
        lin.push("\n❌ No hay registros de Kobo dentro del período " + periodo.label);
        lin.push("   ¿Las fechas del período son correctas?");
        lin.push("   Primer registro en DatosKobo: " + (function(){
          var ts0 = _resolverTsKobo(raw[0], cols);
          return ts0 ? Utilities.formatDate(ts0, CFG.TIMEZONE, "dd/MM/yyyy") : "no detectado";
        })());
        lin.push("   Último registro: " + (function(){
          var ts0 = _resolverTsKobo(raw[raw.length-1], cols);
          return ts0 ? Utilities.formatDate(ts0, CFG.TIMEZONE, "dd/MM/yyyy") : "no detectado";
        })());
        ok = false;
      }
    }
  }

  // ── 5. ¿PARTICIPANTES tiene categorías? ──────────────────────
  var mapa = _construirMapaTarifas();
  var nPart = Object.keys(mapa).length;
  var nConTarifa = Object.keys(mapa).filter(function(k){ return mapa[k].tarifa > 0; }).length;
  if (nPart === 0) {
    lin.push("\n❌ PASO 5 — PARTICIPANTES sin datos de tarifa. ¿Está cargada la lista?");
    ok = false;
  } else {
    lin.push("\n✅ PASO 5 — " + nConTarifa + "/" + nPart + " participantes con tarifa válida.");
  }

  lin.push("\n" + (ok ? "✅ TODO LISTO — corre 📅 Quincena → Ver / actualizar quincena actual"
                      : "⚠️ Corrige los pasos marcados ❌ primero."));

  _alert("🔍 DIAGNÓSTICO DE HORAS\n\n" + lin.join("\n"));
}); }

/*
 * Muestra detalle por participante: quién tiene datos en DatosKobo en el período
 * activo, cuántas entradas/salidas, y si su nombre hace match con PARTICIPANTES.
 * Útil para identificar por qué algún participante no aparece en el reporte.
 */
function diagnosticarParticipantesKobo() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var periodo = _periodoActivo();
  if (!periodo) { _alert("No hay período activo. Configura la quincena primero."); return; }

  var hK = ss.getSheetByName(CFG.HOJAS.DATOS_KOBO);
  if (!hK || hK.getLastRow() < 2) { _alert("DatosKobo está vacío."); return; }

  var fi = new Date(periodo.fi), ff = new Date(periodo.ff);
  var dIni = new Date(fi.getFullYear(), fi.getMonth(), fi.getDate());
  var dFin = new Date(ff.getFullYear(), ff.getMonth(), ff.getDate());

  var enc  = hK.getRange(1, 1, 1, hK.getLastColumn()).getValues()[0];
  var raw  = hK.getRange(2, 1, hK.getLastRow()-1, hK.getLastColumn()).getValues();
  var cols = detectarColumnas(enc, raw.slice(0, 50));
  var mapeoNombres = cargarMapeoNombres();
  var mapa = _construirMapaTarifas();

  // Contar entradas/salidas por participante en el período
  var porPart = {};
  raw.forEach(function(fila) {
    var ts = _resolverTsKobo(fila, cols);
    if (!ts || isNaN(ts)) return;
    var dia = new Date(ts.getFullYear(), ts.getMonth(), ts.getDate());
    if (dia < dIni || dia > dFin) return;
    var nRaw = obtenerParticipanteFila(fila, cols);
    if (!nRaw) return;
    var nombre = normalizarNombre(nRaw, mapeoNombres) || limpiarNombre(nRaw);
    if (!nombre) return;
    if (!porPart[nombre]) porPart[nombre] = { ing: 0, egr: 0, raw: nRaw };
    var tipo = obtenerTipoRegistro(fila, cols);
    if (tipo.esIngreso) porPart[nombre].ing++;
    if (tipo.esEgreso)  porPart[nombre].egr++;
  });

  var nombres = Object.keys(porPart).sort(function(a,b){return a.localeCompare(b,"es");});
  if (nombres.length === 0) {
    _alert("No se encontraron registros de participantes en el período " + periodo.label);
    return;
  }

  var lin = ["📋 PARTICIPANTES EN PERÍODO: " + periodo.label + "\n"];
  var conHoras = 0, sinEntrada = 0, sinMatch = 0;

  nombres.forEach(function(nombre) {
    var d = porPart[nombre];
    var enPart = !!mapa[nombre];
    var estado = "";
    if (!enPart) {
      // Buscar match parcial
      var info = _buscarInfoParticipante(mapa, nombre);
      enPart = !!(info && info.tarifa);
      if (!enPart) { estado = "⚠️ SIN MATCH en PARTICIPANTES"; sinMatch++; }
      else { estado = "✅ match parcial"; }
    } else {
      estado = "✅ match exacto";
    }
    if (d.ing === 0) { estado += " | ❌ SIN ENTRADA (solo salida)"; sinEntrada++; }
    else { conHoras++; }
    lin.push((d.ing === 0 || !enPart ? "❌" : "✅") + " " + nombre +
             " — entradas:" + d.ing + " salidas:" + d.egr + " | " + estado);
  });

  lin.push("\n──────────────────────────────");
  lin.push("Total en período: " + nombres.length);
  lin.push("Con entradas (calculan horas): " + conHoras);
  lin.push("Sin entrada (no calculan): " + sinEntrada);
  lin.push("Sin match en PARTICIPANTES: " + sinMatch);

  _alert(lin.join("\n"));
}); }
