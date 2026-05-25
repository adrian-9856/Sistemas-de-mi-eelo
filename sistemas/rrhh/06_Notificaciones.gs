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
