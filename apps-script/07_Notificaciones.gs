// ============================================================
// 07_Notificaciones.gs  —  Correos automáticos y alertas
// ============================================================

// Envía recordatorio a participantes con facturas pendientes de pago.
function enviarRecordatorioPagos() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var hojaFact = ss.getSheetByName(CONFIG.HOJAS.FACTURACION);
  var hojaPart = ss.getSheetByName(CONFIG.HOJAS.PARTICIPANTES);
  var datos    = hojaFact.getDataRange().getValues();
  var ahora    = new Date();
  var mes      = CONFIG.MESES[ahora.getMonth()];
  var anio     = ahora.getFullYear();

  var mapeoCorreo = _mapeoIDCorreo(hojaPart);
  var enviados = 0;

  for (var i = 1; i < datos.length; i++) {
    var f = datos[i];
    if (f[2] !== mes || f[3] !== anio) continue;
    if (f[10] === "Sí") continue;   // ya pagado
    if (!f[6] || f[6] === 0) continue;

    var id     = f[0];
    var nombre = f[1];
    var monto  = f[6];
    var correo = mapeoCorreo[id];

    if (!correo) continue;

    var asunto  = "[" + CONFIG.NOMBRE_ORGANIZACION + "] Recordatorio de pago — " + mes + " " + anio;
    var cuerpo  =
      "Hola " + nombre + ",\n\n" +
      "Te recordamos que tienes un pago pendiente por:\n\n" +
      "  Período: " + mes + " " + anio + "\n" +
      "  Monto:   Q " + monto.toFixed(2) + "\n\n" +
      "Para completar el proceso, recuerda entregar tu factura a la brevedad.\n\n" +
      "Saludos,\n" + CONFIG.NOMBRE_ORGANIZACION;

    MailApp.sendEmail({ to: correo, subject: asunto, body: cuerpo });
    enviados++;
  }

  SpreadsheetApp.getUi().alert(
    "✅ Recordatorios enviados: " + enviados + " correos."
  );
}

// Envía al admin un resumen consolidado del mes.
function enviarResumenMensual() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var hojaFact = ss.getSheetByName(CONFIG.HOJAS.FACTURACION);
  var datos    = hojaFact.getDataRange().getValues();
  var ahora    = new Date();
  var mes      = CONFIG.MESES[ahora.getMonth()];
  var anio     = ahora.getFullYear();

  var totalHoras = 0, totalMonto = 0, pagados = 0, pendientes = 0;
  var lineas = [];

  for (var i = 1; i < datos.length; i++) {
    var f = datos[i];
    if (f[2] !== mes || f[3] !== anio) continue;

    var horas = parseFloat(f[5]) || 0;
    var monto = parseFloat(f[6]) || 0;
    totalHoras += horas;
    totalMonto += monto;

    if (f[10] === "Sí") { pagados++; } else { pendientes++; }

    lineas.push(
      "  " + f[1] + " — " + horas + " hrs — Q " + monto.toFixed(2) +
      " — " + (f[10] === "Sí" ? "PAGADO" : "PENDIENTE")
    );
  }

  var cuerpo =
    "Resumen de facturación — " + mes + " " + anio + "\n" +
    "═══════════════════════════════════\n\n" +
    lineas.join("\n") + "\n\n" +
    "═══════════════════════════════════\n" +
    "Total horas:   " + totalHoras.toFixed(2) + " hrs\n" +
    "Total a pagar: Q " + totalMonto.toFixed(2) + "\n" +
    "Pagados:       " + pagados + "\n" +
    "Pendientes:    " + pendientes + "\n";

  MailApp.sendEmail({
    to:      CONFIG.CORREO_ADMIN,
    subject: "[" + CONFIG.NOMBRE_ORGANIZACION + "] Resumen " + mes + " " + anio,
    body:    cuerpo,
  });

  SpreadsheetApp.getUi().alert("✅ Resumen enviado a " + CONFIG.CORREO_ADMIN);
}

// Alerta interna cuando se crea una nueva orden de producción.
function notificarNuevaOrden(numeroOrden, cliente, descripcion) {
  var asunto = "[" + CONFIG.NOMBRE_ORGANIZACION + "] Nueva orden: " + numeroOrden;
  var cuerpo =
    "Se registró una nueva orden:\n\n" +
    "  Número:   " + numeroOrden + "\n" +
    "  Cliente:  " + cliente + "\n" +
    "  Descripción: " + descripcion + "\n\n" +
    "Revisa el sistema para asignar participantes.";

  MailApp.sendEmail({ to: CONFIG.CORREO_ADMIN, subject: asunto, body: cuerpo });
}

function _mapeoIDCorreo(hojaPart) {
  var datos = hojaPart.getDataRange().getValues();
  var mapa  = {};
  for (var i = 1; i < datos.length; i++) {
    var id     = datos[i][0]; // col A
    var correo = datos[i][8]; // col I = Correo_Electronico
    if (id && correo) mapa[id] = correo;
  }
  return mapa;
}
