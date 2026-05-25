// ============================================================
// 00_Config.gs  —  Configuración central del sistema Mi eelo
// Edita SOLO este archivo para ajustar el sistema.
// ============================================================

const CONFIG = {

  // --- TARIFAS ---
  TARIFA_HORA: 25,

  // --- CORREOS ---
  CORREO_ADMIN: "adrian@creamosguatemala.org",
  NOMBRE_ORGANIZACION: "Mi eelo",

  // --- HOJAS DEL SPREADSHEET ---
  HOJAS: {
    PARTICIPANTES:  "PARTICIPANTES_MASTER",
    ASISTENCIA:     "ASISTENCIA_KOBO",
    FACTURACION:    "FACTURACION_CONSOLIDADA",
    ORDENES:        "ORDENES_PRODUCCION",
    VENTAS:         "COMERCIAL_VENTAS",
    COMPRAS:        "COMERCIAL_COMPRAS",
    DASHBOARD:      "DASHBOARD",
    NOMBRES_CANON:  "NOMBRES_CANONICOS",   // hoja auxiliar de mapeo
  },

  // --- KOBO ---
  // Campo "Participante" contiene: "Nombre Completo (CREAMOS_ID)"
  // Ej: "Juana del Rosario Vicente Choy (JUVI281187)"
  KOBO_API_TOKEN:  "TU_TOKEN_AQUI",
  KOBO_ASSET_UID:  "TU_ASSET_UID_AQUI",

  // Nombres exactos de columnas en la hoja DatosKobo (o en la API)
  KOBO_COL: {
    START:        1,   // col A: start (timestamp de envío del formulario)
    END:          2,   // col B: end
    TIPO:         4,   // col D: "Ingreso / Egreso"  →  "🟢 Entrada" | "🔴 Salida"
    PARTICIPANTE: 5,   // col E: "Nombre (CREAMOS_ID)"
    UUID:         12,  // col L: _uuid
  },

  // Valores exactos del campo Tipo
  KOBO_ENTRADA: "🟢 Entrada",
  KOBO_SALIDA:  "🔴 Salida",

  // --- GOOGLE DRIVE ---
  // Se llenan automáticamente la primera vez que ejecutas "Crear estructura en Drive"
  DRIVE: {
    CARPETA_RAIZ:     "",
    CARPETA_FACTURAS: "",
    CARPETA_REPORTES: "",
    CARPETA_ORDENES:  "",
  },

  // --- MESES ---
  MESES: [
    "Enero","Febrero","Marzo","Abril","Mayo","Junio",
    "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"
  ],
};
