// ============================================================
// SISTEMA PRODUCCIÓN — Mi eelo
// Google Sheets: "Mi eelo · Producción"
// Maneja: OPs, OMs, Clientes, Drive, Docs
// ============================================================

const CFG = {
  ORG:           "Mi eelo",
  DIRECCION:     "13 calle 2-90 Zona 7, Landívar, Ciudad de Guatemala",
  TELEFONO:      "(502) 3764 9769",
  CORREO_ADMIN:  "adrian@creamosguatemala.org",

  HOJAS: {
    ORDENES:    "ORDENES",
    CLIENTES:   "CLIENTES",
    DASHBOARD:  "DASHBOARD",
  },

  // IDs de carpetas Drive (se llenan automáticamente la primera vez)
  DRIVE: {
    RAIZ:     "",   // "Mi eelo · Producción"
    ORDENES:  "",   // subcarpeta Órdenes/
    CLIENTES: "",   // subcarpeta Órdenes/Clientes/
  },
};
