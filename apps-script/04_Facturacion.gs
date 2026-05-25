// ============================================================
// 04_Facturacion.gs  —  Cálculo y migración de facturación
//
// Estructura real de Facturación_Mi_eelo.xlsx:
//   Hoja "Datos de Facturación": DPI, NIT, banco, SAT
//   Hojas "Mayo" a "Abril": una por mes, con quincenas
//   → Columnas variables por mes, pero siempre tienen:
//     Nombre | Quincena1_Monto | Quincena1_Factura |
//     Quincena2_Monto | Quincena2_Factura | Total |
//     Declaraguate | Pagado
// ============================================================

// ── Importación histórica ──────────────────────────────────

// Importa la hoja "Datos de Facturación" (DPI, NIT, banco)
// a la hoja DATOS_FACTURACION del Sheets.
// Pega esa hoja en este Sheets con nombre "IMPORT_FACT_DATOS" primero.
function importarDatosFacturacion() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var hojaTemp = ss.getSheetByName("IMPORT_FACT_DATOS");
  var hojaDest = ss.getSheetByName("DATOS_FACTURACION");

  if (!hojaTemp) {
    SpreadsheetApp.getUi().alert(
      "Crea una hoja 'IMPORT_FACT_DATOS' y pega la hoja\n" +
      "'Datos de Facturación' del archivo Facturación_Mi_eelo.xlsx"
    );
    return;
  }
  if (!hojaDest) {
    hojaDest = ss.insertSheet("DATOS_FACTURACION");
    hojaDest.appendRow([
      "Creamos_ID","Nombre","DPI","NIT","Correo","Agencia_Virtual",
      "Status_SAT","Banco","Tipo_Cuenta","Numero_Cuenta","Notas"
    ]);
  }

  var datos = hojaTemp.getDataRange().getValues();
  var importados = 0;

  // Detecta la fila de encabezados (busca "Nombre" o similar)
  var inicio = 1;
  for (var i = 0; i < Math.min(datos.length, 5); i++) {
    if (String(datos[i][1]).toLowerCase().includes("nombre")) { inicio = i + 1; break; }
  }

  for (var i = inicio; i < datos.length; i++) {
    var nombre = datos[i][1];
    if (!nombre) continue;

    hojaDest.appendRow([
      "",            // A: Creamos_ID (se cruza con PARTICIPANTES después)
      nombre,        // B: Nombre
      datos[i][2],   // C: DPI
      datos[i][3],   // D: NIT
      datos[i][4],   // E: Correo
      datos[i][5],   // F: Agencia_Virtual
      datos[i][6],   // G: Status_SAT
      "",            // H: Banco (está en las hojas mensuales)
      "",            // I: Tipo_Cuenta
      "",            // J: Numero_Cuenta
      "",            // K: Notas
    ]);
    importados++;
  }

  SpreadsheetApp.getUi().alert("✅ " + importados + " registros de facturación importados.");
}

// Importa el historial de 14 meses a FACTURACION_CONSOLIDADA.
// Pega cada hoja mensual en este Sheets con nombre "IMPORT_[MES]" (ej: "IMPORT_Mayo").
function importarHistorialFacturacion() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var hojaDest = ss.getSheetByName(CONFIG.HOJAS.FACTURACION);

  var mesesArchivo = [
    "Mayo","Junio","Julio","Agosto","Septiembre",
    "Octubre","Noviembre","Diciembre","Enero","Febrero",
    "Marzo","Abril"
  ];

  var totalImportado = 0;

  mesesArchivo.forEach(function(mes) {
    var hojaTemp = ss.getSheetByName("IMPORT_" + mes);
    if (!hojaTemp) return; // si no está pegada, se salta

    var datos = hojaTemp.getDataRange().getValues();
    var anio  = (mes === "Enero" || mes === "Febrero" || mes === "Marzo" || mes === "Abril")
                ? 2026 : 2025;

    // Detecta fila de encabezados
    var inicio = _detectarFilaEncabezado(datos);

    for (var i = inicio + 1; i < datos.length; i++) {
      var fila   = datos[i];
      var nombre = fila[0]; // primera columna suele ser el nombre
      if (!nombre) continue;

      // Detecta columnas de monto por la presencia de valores numéricos
      var q1Monto     = _buscarNumero(fila, 1, 5);
      var q1Factura   = _buscarTexto(fila, 2, 6, ["Sí","No","si","no"]);
      var q2Monto     = _buscarNumero(fila, 5, 10);
      var q2Factura   = _buscarTexto(fila, 6, 11, ["Sí","No","si","no"]);
      var declaraguate = fila[fila.length - 3] || "";
      var pagado       = fila[fila.length - 2] || "";

      // Quincena 1
      if (q1Monto) {
        hojaDest.appendRow([
          "",        // A: ID (se cruza después)
          nombre,    // B: Nombre
          mes,       // C: Mes
          anio,      // D: Año
          "1",       // E: Quincena
          "",        // F: Horas_Trabajadas
          q1Monto,   // G: Monto_A_Pagar
          q1Factura, // H: Factura_Entregada
          "",        // I: Numero_Factura
          declaraguate, // J: Declaraguate
          pagado,    // K: Pagado
          "",        // L: Fecha_Pago
          "",        // M: Comentarios
        ]);
        totalImportado++;
      }

      // Quincena 2
      if (q2Monto) {
        hojaDest.appendRow([
          "", nombre, mes, anio, "2",
          "", q2Monto, q2Factura, "", declaraguate, pagado, "", "",
        ]);
        totalImportado++;
      }
    }
  });

  SpreadsheetApp.getUi().alert(
    "✅ " + totalImportado + " registros de facturación histórica importados.\n" +
    "Hojas importadas: las que encontró con prefijo IMPORT_"
  );
}

// ── Cálculo automático mensual ────────────────────────────

function calcularFacturacionMes() {
  var ahora    = new Date();
  _calcularFacturacion(ahora.getMonth() + 1, ahora.getFullYear());
}

function _calcularFacturacion(mes, anio) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var hojaFact  = ss.getSheetByName(CONFIG.HOJAS.FACTURACION);
  var hojaAsist = ss.getSheetByName(CONFIG.HOJAS.ASISTENCIA);
  var hojaPart  = ss.getSheetByName(CONFIG.HOJAS.PARTICIPANTES);

  var asistencia    = hojaAsist.getDataRange().getValues();
  var participantes = hojaPart.getDataRange().getValues();
  var nombreMes     = CONFIG.MESES[mes - 1];

  // Agrupamos horas a pagar por participante+quincena del mes dado
  var horasPorID = { "1": {}, "2": {} };

  for (var i = 1; i < asistencia.length; i++) {
    var f          = asistencia[i];
    var id         = f[0];
    var tipo       = f[3]; // "🟢 Entrada" | "🔴 Salida"
    var timestamp  = new Date(f[2]);
    var horasAPagar = parseFloat(f[8]) || 0;

    if (!id || tipo !== CONFIG.KOBO_ENTRADA) continue;
    if (isNaN(timestamp.getTime())) continue;
    if (timestamp.getMonth() + 1 !== mes || timestamp.getFullYear() !== anio) continue;

    var quincena = timestamp.getDate() <= 15 ? "1" : "2";
    horasPorID[quincena][id] = (horasPorID[quincena][id] || 0) + horasAPagar;
  }

  var datosFact   = hojaFact.getDataRange().getValues();
  var procesados  = 0;

  ["1","2"].forEach(function(q) {
    participantes.slice(1).forEach(function(p) {
      var id     = String(p[0]).trim();
      var nombre = p[1];
      if (!id || !nombre) return;

      var horas  = Math.round((horasPorID[q][id] || 0) * 100) / 100;
      var monto  = Math.round(horas * CONFIG.TARIFA_HORA * 100) / 100;

      // Busca si ya existe la fila para este ID/mes/año/quincena
      var filaExistente = -1;
      for (var j = 1; j < datosFact.length; j++) {
        if (String(datosFact[j][0]) === id &&
            datosFact[j][2] === nombreMes &&
            datosFact[j][3] === anio &&
            String(datosFact[j][4]) === q) {
          filaExistente = j + 1;
          break;
        }
      }

      if (filaExistente > 0) {
        hojaFact.getRange(filaExistente, 6).setValue(horas);
        hojaFact.getRange(filaExistente, 7).setValue(monto);
      } else {
        hojaFact.appendRow([
          id, nombre, nombreMes, anio, q,
          horas, monto,
          "No", "", "No", "No", "", "",
        ]);
      }
      procesados++;
    });
  });

  actualizarDashboard();
  SpreadsheetApp.getUi().alert(
    "✅ Facturación calculada — " + nombreMes + " " + anio + "\n" +
    procesados + " registros procesados."
  );
}

// ── Utilidades internas ──────────────────────────────────

function _detectarFilaEncabezado(datos) {
  var maxVal = 0, filaMax = 0;
  for (var i = 0; i < Math.min(datos.length, 5); i++) {
    var cnt = datos[i].filter(function(v){ return v !== null && v !== ""; }).length;
    if (cnt > maxVal) { maxVal = cnt; filaMax = i; }
  }
  return filaMax;
}

function _buscarNumero(fila, desde, hasta) {
  for (var i = desde; i < Math.min(fila.length, hasta); i++) {
    var v = parseFloat(fila[i]);
    if (!isNaN(v) && v > 0) return v;
  }
  return null;
}

function _buscarTexto(fila, desde, hasta, opciones) {
  for (var i = desde; i < Math.min(fila.length, hasta); i++) {
    var v = String(fila[i] || "").trim();
    if (opciones.indexOf(v) > -1) return v;
  }
  return "";
}
