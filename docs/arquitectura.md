# Arquitectura del Sistema — Mi eelo

## Visión general

```
KoboToolbox API
      │  (cada hora, automático)
      ▼
┌─────────────────────────────────────────────────────┐
│              GOOGLE SHEETS                          │
│                                                     │
│  PARTICIPANTES_MASTER ◄──────────────────────────┐  │
│        │                                         │  │
│        │ ID lookup                               │  │
│        ▼                                         │  │
│  ASISTENCIA_KOBO                                 │  │
│        │                                         │  │
│        │ SUMIFS por participante/mes             │  │
│        ▼                                         │  │
│  FACTURACION_CONSOLIDADA                         │  │
│                                                  │  │
│  ORDENES_PRODUCCION ─────────────────────────────┘  │
│  COMERCIAL_VENTAS                                    │
│  COMERCIAL_COMPRAS                                   │
│  DASHBOARD (métricas en tiempo real)                 │
└─────────────┬───────────────────────┬────────────────┘
              │                       │
              ▼                       ▼
      GOOGLE DOCS                GOOGLE DRIVE
   ┌──────────────┐          ┌──────────────────┐
   │ • Facturas   │          │ Mi eelo - Sistema │
   │   individuales│         │   Unificado/      │
   │ • Reportes   │          │   ├── Facturas/   │
   │   mensuales  │          │   ├── Reportes/   │
   │ • Órdenes    │          │   └── Ordenes/    │
   └──────────────┘          └──────────────────┘
              │
              ▼
         GMAIL / EMAIL
   ┌──────────────────────┐
   │ • Recordatorio pagos │
   │   (viernes)          │
   │ • Resumen mensual    │
   │   (día 1)            │
   └──────────────────────┘
```

---

## Módulos del sistema

### `00_Config.gs` — Configuración central
Único archivo que necesitas editar para personalizar el sistema.
Contiene tarifas, correos, nombres de hojas, credenciales de Kobo.

### `01_Menu.gs` — Menú y hooks de edición
- Crea el menú **🚀 Mi eelo** en el Spreadsheet.
- `onEdit`: detecta cuando se marca un pago como completado y actualiza el Dashboard.

### `02_KoboImport.gs` — Importación de asistencia
- Llama a la API REST de KoboToolbox.
- Filtra registros ya importados por UUID (sin duplicados).
- Calcula horas trabajadas y porcentaje de pago por tipo de registro.
- Llama a `normalizarIDs()` automáticamente al terminar.

### `03_Participantes.gs` — Gestión de participantes
- Genera IDs secuenciales (P001, P002...).
- Mantiene el mapeo Nombre → ID para conectar todas las hojas.

### `04_Facturacion.gs` — Cálculo de pagos
- Agrega horas desde `ASISTENCIA_KOBO` por participante y mes.
- Crea o actualiza filas en `FACTURACION_CONSOLIDADA`.
- Respeta datos ya existentes (facturas marcadas, comentarios).

### `05_Documentos.gs` — Generación de Google Docs
- **Facturas**: un Doc por participante por mes, con tabla de horas/monto.
- **Reporte mensual**: Doc consolidado con todos los participantes.
- **Órdenes**: Doc de cada OP/OM con los datos de la hoja.
- Guarda el enlace del Doc en la hoja correspondiente.

### `06_Drive.gs` — Organización en Drive
- Crea la estructura de carpetas la primera vez.
- Mueve automáticamente cada Doc generado a su carpeta correcta.
- Usa `PropertiesService` para recordar los IDs de carpetas entre ejecuciones.

### `07_Notificaciones.gs` — Correos automáticos
- Recordatorio semanal a participantes con pagos pendientes.
- Resumen mensual al administrador el día 1 de cada mes.
- Alerta interna cuando se registra una orden nueva.

### `08_Dashboard.gs` — Métricas en tiempo real
- Hoja `DASHBOARD` siempre actualizada con las métricas clave.
- Colorea en rojo si hay pagos pendientes.
- Se actualiza cada 30 minutos y al marcar un pago como completado.

### `09_Triggers.gs` — Automatización
- Se ejecuta una sola vez para instalar todos los triggers.
- Borra triggers anteriores antes de crear los nuevos (evita duplicados).

---

## Flujo de datos típico (diario)

1. **Cada hora** → Kobo → nuevos registros en `ASISTENCIA_KOBO`
2. **Medianoche** → recalcula `FACTURACION_CONSOLIDADA`
3. **Cada 30 min** → actualiza `DASHBOARD`
4. **Viernes** → correo de recordatorio a participantes con pagos pendientes
5. **Día 1 del mes** → correo de resumen al admin

## Flujo manual (cuando se necesita)

1. Admin genera facturas individuales: **Menú → Generar facturas en Google Docs**
2. Se crean Docs en Drive: `Facturas/2026/Mayo/Factura_P001_Mayo_2026.gdoc`
3. El enlace queda guardado en la hoja para fácil acceso
4. Admin genera reporte mensual: **Menú → Generar reporte mensual**
5. Reporte queda en: `Reportes/2026/Reporte_Mayo_2026.gdoc`
