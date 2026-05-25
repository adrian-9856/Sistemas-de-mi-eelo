# Sistema Unificado — Mi eelo

Sistema centralizado en Google Sheets con automatizaciones via Apps Script.
Conecta KoboToolbox, Google Drive, Google Docs y correo en un solo flujo.

---

## ¿Qué hace este sistema?

| Tarea                          | Antes                        | Ahora                              |
|-------------------------------|------------------------------|------------------------------------|
| Asistencia                    | Manual en Excel              | Importación automática desde Kobo  |
| Facturación                   | 14 hojas separadas           | 1 hoja calculada automáticamente   |
| Facturas individuales         | Manual                       | Google Docs generados con 1 clic   |
| Reportes mensuales            | Manual                       | Google Doc automático al admin     |
| Órdenes de producción         | 141 hojas Excel              | 1 hoja con Doc por orden           |
| Notificaciones de pago        | Ninguna                      | Correo automático cada viernes     |
| Resumen al admin              | Manual                       | Correo automático el día 1         |
| Archivos organizados          | Carpetas desorganizadas      | Estructura Drive automática        |

---

## Estructura del repositorio

```
apps-script/
  00_Config.gs          ← EMPIEZA AQUÍ: edita tus credenciales y tarifas
  01_Menu.gs            ← Menú en Google Sheets
  02_KoboImport.gs      ← Importa asistencia desde KoboToolbox
  03_Participantes.gs   ← Gestión de IDs de participantes
  04_Facturacion.gs     ← Cálculo de horas y montos
  05_Documentos.gs      ← Genera facturas y reportes en Google Docs
  06_Drive.gs           ← Organiza archivos en Google Drive
  07_Notificaciones.gs  ← Envía correos automáticos
  08_Dashboard.gs       ← Actualiza métricas en tiempo real
  09_Triggers.gs        ← Configura automatizaciones (ejecutar 1 sola vez)

docs/
  instalacion.md        ← Guía paso a paso para instalar
  arquitectura.md       ← Cómo fluyen los datos entre módulos
```

---

## Inicio rápido

1. Lee [`docs/instalacion.md`](docs/instalacion.md) — guía completa paso a paso
2. Edita `apps-script/00_Config.gs` con tu token de Kobo y correo
3. Pega los scripts en **Extensiones → Apps Script** de tu Google Sheets
4. Ejecuta **🚀 Mi eelo → ⚙️ Configurar triggers automáticos** una sola vez

Listo. El sistema corre solo desde ese punto.

---

## Automatizaciones activas

| Frecuencia         | Acción                                     |
|--------------------|--------------------------------------------|
| Cada hora          | Importa asistencia desde KoboToolbox       |
| Diario (00:00)     | Recalcula facturación del mes              |
| Cada 30 minutos    | Actualiza Dashboard                        |
| Día 1 de cada mes  | Envía resumen mensual al administrador     |
| Cada viernes 9am   | Recordatorio de pagos pendientes           |

---

*Versión: 2.0 — Mayo 2026*
