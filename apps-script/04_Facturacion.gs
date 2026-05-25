// ============================================================
// 04_Facturacion.gs  —  Cálculo automático de facturación
// ============================================================

// Calcula horas y montos para todos los participantes del mes/año actuales.
function calcularFacturacionMes() {
  var ahora  = new Date();
  var mes    = ahora.getMonth() + 1;
  var anio   = ahora.getFullYear();
  _calcularFacturacion(mes, anio);
}

function _calcularFacturacion(mes, anio) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var hojaFact  = ss.getSheetByName(CONFIG.HOJAS.FACTURACION);
  var hojaAsist = ss.getSheetByName(CONFIG.HOJAS.ASISTENCIA);
  var hojaPart  = ss.getSheetByName(CONFIG.HOJAS.PARTICIPANTES);

  var asistencia    = hojaAsist.getDataRange().getValues();
  var participantes = hojaPart.getDataRange().getValues();

  // Agrupar horas por participante para el mes/año dado
  var horasPorID = {};

  for (var i = 1; i < asistencia.length; i++) {
    var fila = asistencia[i];
    var id   = fila[0];
    if (!id) continue;

    var fechaEntrada = new Date(fila[2]);
    if (isNaN(fechaEntrada.getTime())) continue;

    if (fechaEntrada.getMonth() + 1 !== mes) continue;
    if (fechaEntrada.getFullYear()   !== anio) continue;

    var horasAPagar = parseFloat(fila[9]) || 0;
    horasPorID[id]  = (horasPorID[id] || 0) + horasAPagar;
  }

  // Agregar o actualizar filas en FACTURACION_CONSOLIDADA
  var datosFact = hojaFact.getDataRange().getValues();
  var nombreMes = CONFIG.MESES[mes - 1];

  participantes.slice(1).forEach(function(p) {
    var id     = p[0];
    var nombre = p[2];
    if (!id) return;

    var horas  = Math.round((horasPorID[id] || 0) * 100) / 100;
    var monto  = Math.round(horas * CONFIG.TARIFA_HORA * 100) / 100;

    // Buscar si ya existe una fila para este participante/mes/año
    var filaExistente = -1;
    for (var j = 1; j < datosFact.length; j++) {
      if (datosFact[j][0] === id &&
          datosFact[j][2] === nombreMes &&
          datosFact[j][3] === anio) {
        filaExistente = j + 1; // número de fila en la hoja (1-based)
        break;
      }
    }

    if (filaExistente > 0) {
      // Actualiza solo horas y monto, respeta el resto (factura, pagado…)
      hojaFact.getRange(filaExistente, 6).setValue(horas);
      hojaFact.getRange(filaExistente, 7).setValue(monto);
    } else {
      // Nueva fila
      hojaFact.appendRow([
        id,        // A: ID_Participante
        nombre,    // B: Nombre
        nombreMes, // C: Mes
        anio,      // D: Año
        "",        // E: Quincena
        horas,     // F: Horas_Trabajadas
        monto,     // G: Monto_A_Pagar
        "No",      // H: Factura_Entregada
        "",        // I: Numero_Factura
        "No",      // J: Declaraguate
        "No",      // K: Pagado
        "",        // L: Fecha_Pago
        "",        // M: Comentarios
      ]);
    }
  });

  actualizarDashboard();
  SpreadsheetApp.getUi().alert(
    "✅ Facturación calculada\nMes: " + nombreMes + " " + anio + "\n" +
    "Participantes procesados: " + Object.keys(horasPorID).length
  );
}
