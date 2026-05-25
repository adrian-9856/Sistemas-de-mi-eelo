// ============================================================
// 00_Config.gs  —  Configuración central del sistema Mi eelo
// Edita SOLO este archivo para ajustar el sistema.
// ============================================================

const CONFIG = {

  // --- TARIFAS ---
  TARIFA_HORA: 25,            // Quetzales por hora trabajada
  HORAS_QUINCENA_NORMAL: 80, // Horas esperadas por quincena

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
  },

  // --- KOBO ---
  KOBO_API_TOKEN: "TU_TOKEN_AQUI",       // Reemplaza con tu token de KoboToolbox
  KOBO_ASSET_UID: "TU_ASSET_UID_AQUI",  // UID del formulario en Kobo
  KOBO_CAMPO_NOMBRE: "nombre_participante",
  KOBO_CAMPO_ENTRADA: "hora_entrada",
  KOBO_CAMPO_SALIDA:  "hora_salida",
  KOBO_CAMPO_TIPO:    "tipo_registro",
  KOBO_CAMPO_UUID:    "_uuid",

  // --- GOOGLE DRIVE (IDs de carpetas) ---
  // Deja vacío la primera vez; el sistema las crea automáticamente.
  DRIVE: {
    CARPETA_RAIZ:     "",   // "Mi eelo - Sistema Unificado"
    CARPETA_FACTURAS: "",
    CARPETA_REPORTES: "",
    CARPETA_ORDENES:  "",
  },

  // --- MESES (para fórmulas y docs) ---
  MESES: [
    "Enero","Febrero","Marzo","Abril","Mayo","Junio",
    "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"
  ],
};
