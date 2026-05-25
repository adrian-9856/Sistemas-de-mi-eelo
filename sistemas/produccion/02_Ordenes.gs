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
