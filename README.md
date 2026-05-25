# Sistemas Mi eelo

Dos sistemas independientes. Cada uno es un Google Sheets separado con sus propios scripts.

---

## Sistema 1 — Producción (OPs y OMs)
`sistemas/produccion/`

**Google Sheets:** "Mi eelo · Producción"

| Hoja | Contenido |
|------|-----------|
| ORDENES | Todas las OPs y OMs con 35 columnas: cliente, medidas, serigrafía, mockup, estado |
| CLIENTES | Lista de clientes con enlace directo a su carpeta en Drive |
| DASHBOARD | Métricas en tiempo real: órdenes por estado y por cliente |

**Lo que hace:**
- Nueva orden (OP o OM) con número automático `OP26-001`
- Genera el Google Doc del formulario real con 1 clic desde cualquier fila
- Organiza los Docs en Drive por cliente: `Órdenes/Clientes/Peace by Piece/`
- Filtra todas las órdenes de un cliente con 1 clic
- Dashboard que se actualiza automáticamente

**Scripts (pegar en Apps Script en este orden):**
```
00_Config.gs     ← edita aquí tu correo y nombre de org
01_Menu.gs
02_Ordenes.gs
03_Documentos.gs
04_Drive.gs
05_Dashboard.gs
06_Triggers.gs
```

---

## Sistema 2 — RRHH / Participantes
`sistemas/rrhh/`

**Google Sheets:** "Mi eelo · RRHH"

| Hoja | Contenido |
|------|-----------|
| PARTICIPANTES | 76 personas: Creamos ID, nombre, proyecto, datos bancarios |
| ASISTENCIA | Registro de entradas/salidas desde Kobo, horas calculadas |
| FACTURACION | Historial de pagos por quincena y mes |
| DASHBOARD | Métricas: participantes activos, total a pagar, pendientes |

**Lo que hace:**
- Importa asistencia desde KoboToolbox cada hora (automático)
- Empareja cada entrada con su salida para calcular horas reales
- Calcula montos por quincena automáticamente
- Genera recibo de pago en Google Docs por participante
- Correos automáticos: recordatorio de pagos (viernes) y resumen mensual (día 1)
- Drive organizado: `Facturas/2026/Mayo/Recibo_ANLA060686_Q1_Mayo_2026`

**Scripts (pegar en Apps Script en este orden):**
```
00_Config.gs     ← edita aquí tu token Kobo, correo y tarifa/hora
01_Menu.gs
02_Kobo.gs
03_Facturacion.gs
04_Documentos.gs
05_Drive.gs
06_Notificaciones.gs
07_Dashboard.gs
08_Participantes.gs
09_Triggers.gs
```

---

## Cómo instalar cada sistema

1. Crea un Google Sheets nuevo con el nombre correspondiente
2. Ve a **Extensiones → Apps Script**
3. Borra el código default y crea un archivo por cada `.gs` (botón "+")
4. Pega el contenido de cada archivo en el mismo orden
5. Edita `00_Config.gs` con tus datos reales
6. Recarga el Sheets — aparece el menú
7. Ejecuta **Crear estructura en Drive** (una vez)
8. Ejecuta **Configurar automatizaciones** (una vez)

La guía detallada de carga inicial está en `docs/carga-inicial.md`
