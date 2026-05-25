# Guía de Instalación — Sistema Mi eelo

## Requisitos previos
- Google Sheets con las 6 hojas creadas (ver nombres en sección Hojas)
- Cuenta en KoboToolbox con el formulario de asistencia activo
- Token de API de Kobo

---

## Paso 1 — Obtener tu Token de KoboToolbox

1. Inicia sesión en [kc.kobotoolbox.org](https://kc.kobotoolbox.org)
2. Ve a **Tu perfil → API Token**
3. Copia el token

---

## Paso 2 — Obtener el Asset UID de tu formulario

1. En KoboToolbox, abre tu formulario de asistencia
2. La URL se ve así: `https://kc.kobotoolbox.org/api/v2/assets/`**aBcXyZ123**/
3. El código al final es tu Asset UID

---

## Paso 3 — Instalar el código en Apps Script

1. En tu Google Sheets: **Extensiones → Apps Script**
2. Borra el código que aparece por defecto
3. Crea archivos nuevos con el botón **"+"** (uno por cada archivo `.gs`)
4. Nombra y pega cada archivo en el mismo orden:

| Archivo en este repo      | Nombre en Apps Script    |
|---------------------------|--------------------------|
| `00_Config.gs`            | `Config`                 |
| `01_Menu.gs`              | `Menu`                   |
| `02_KoboImport.gs`        | `KoboImport`             |
| `03_Participantes.gs`     | `Participantes`          |
| `04_Facturacion.gs`       | `Facturacion`            |
| `05_Documentos.gs`        | `Documentos`             |
| `06_Drive.gs`             | `Drive`                  |
| `07_Notificaciones.gs`    | `Notificaciones`         |
| `08_Dashboard.gs`         | `Dashboard`              |
| `09_Triggers.gs`          | `Triggers`               |

---

## Paso 4 — Configurar `00_Config.gs`

Edita estas líneas con tus datos reales:

```javascript
KOBO_API_TOKEN: "TU_TOKEN_AQUI",     // token copiado en Paso 1
KOBO_ASSET_UID: "TU_ASSET_UID_AQUI", // UID copiado en Paso 2
CORREO_ADMIN:   "tu@correo.com",
TARIFA_HORA:    25,                   // ajusta a tu tarifa real
```

Si los nombres de campos en tu formulario Kobo son diferentes, ajusta también:
```javascript
KOBO_CAMPO_NOMBRE:   "nombre_participante",
KOBO_CAMPO_ENTRADA:  "hora_entrada",
KOBO_CAMPO_SALIDA:   "hora_salida",
```

---

## Paso 5 — Primera ejecución (en orden)

1. **Guardar** todos los archivos en Apps Script (Ctrl+S)
2. Recargar el Google Sheets — aparecerá el menú **🚀 Mi eelo**
3. Ir a **🚀 Mi eelo → 📁 Crear estructura en Drive** (crea las carpetas automáticamente)
4. Ir a **🚀 Mi eelo → 📥 Importar desde Kobo** (autoriza permisos cuando te pida)
5. Ir a **🚀 Mi eelo → 🔗 Normalizar IDs participantes**
6. Ir a **🚀 Mi eelo → 💰 Calcular facturación del mes**
7. Ir a **🚀 Mi eelo → 🔄 Actualizar Dashboard**
8. Ir a **🚀 Mi eelo → ⚙️ Configurar triggers automáticos** (una sola vez)

---

## Paso 6 — Verificar permisos

La primera vez que ejecutes cada función, Google pedirá autorización para:
- **KoboToolbox**: acceso a internet (`UrlFetchApp`)
- **Drive**: crear y mover archivos
- **Gmail**: enviar correos
- **Docs**: crear documentos

Acepta todos. Si aparece "Google no verificó esta app", haz clic en **Avanzado → Ir a Mi eelo (no seguro)** — es normal para scripts propios.

---

## Nombres de hojas requeridos

El sistema espera exactamente estos nombres (respeta mayúsculas):

| Hoja                     | Contenido                        |
|--------------------------|----------------------------------|
| `PARTICIPANTES_MASTER`   | Datos maestros de participantes  |
| `ASISTENCIA_KOBO`        | Registros de asistencia de Kobo  |
| `FACTURACION_CONSOLIDADA`| Horas y montos por participante  |
| `ORDENES_PRODUCCION`     | OPs y OMs                        |
| `COMERCIAL_VENTAS`       | Registro de ventas               |
| `COMERCIAL_COMPRAS`      | Registro de compras              |
| `DASHBOARD`              | Se crea automáticamente          |

---

## Estructura de carpetas en Drive (creada automáticamente)

```
Mi eelo - Sistema Unificado/
├── Facturas/
│   └── 2026/
│       ├── Enero/
│       ├── Febrero/
│       └── ...
├── Reportes/
│   └── 2026/
├── Ordenes_Produccion/
│   ├── OPs/
│   └── OMs/
└── Archivos_Historicos/
```

---

## Automatizaciones activas (después del Paso 8)

| Frecuencia          | Qué hace                                      |
|---------------------|-----------------------------------------------|
| Cada hora           | Importar asistencia desde Kobo                |
| Diario (00:00)      | Recalcular horas y montos de facturación      |
| Cada 30 minutos     | Actualizar Dashboard                          |
| Día 1 de cada mes   | Enviar resumen mensual al admin               |
| Cada viernes (9am)  | Recordatorio de pagos pendientes              |
