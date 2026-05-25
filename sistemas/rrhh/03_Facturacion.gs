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
