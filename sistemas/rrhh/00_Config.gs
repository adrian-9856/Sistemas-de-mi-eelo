// ============================================================
// SISTEMA RRHH — Mi eelo
// Google Sheets: "Mi eelo · RRHH"
// Maneja: Participantes, Asistencia Kobo, Facturación
// ============================================================

const CFG = {
  ORG:          "Mi eelo",
  CORREO_ADMIN: "adrian@creamosguatemala.org",
  TARIFA_HORA:  25,   // Q por hora trabajada

  HOJAS: {
    PARTICIPANTES: "PARTICIPANTES",
    ASISTENCIA:    "ASISTENCIA",
    FACTURACION:   "FACTURACION",
    DASHBOARD:     "DASHBOARD",
  },

  // KoboToolbox
  KOBO_TOKEN:     "TU_TOKEN_AQUI",
  KOBO_ASSET_UID: "TU_ASSET_UID_AQUI",

  // Campo "Participante" en Kobo tiene formato: "Nombre (CREAMOS_ID)"
  // Ej: "Juana del Rosario Vicente Choy (JUVI281187)"
  KOBO_TIPO_ENTRADA: "🟢 Entrada",
  KOBO_TIPO_SALIDA:  "🔴 Salida",

  MESES: ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
          "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"],

  DRIVE: {
    RAIZ:      "",  // "Mi eelo · RRHH"
    FACTURAS:  "",  // subcarpeta Facturas/
  },
};
