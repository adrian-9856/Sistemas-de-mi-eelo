# Sistemas Mi eelo

Dos sistemas independientes. Cada uno es **un solo archivo** para pegar en Google Apps Script.

---

## Sistema 1 — Producción (OPs y OMs)
**Archivo:** `sistemas/PRODUCCION_todo.gs`
**Google Sheets:** crea uno nuevo llamado "Mi eelo · Producción"

**Cómo instalar:**
1. Abre el Sheets → **Extensiones → Apps Script**
2. Borra el código default
3. Pega TODO el contenido de `PRODUCCION_todo.gs`
4. Edita las 3 líneas de configuración al inicio del archivo
5. Guarda → recarga el Sheets → aparece el menú **📦 Producción**
6. Ejecuta **Crear estructura en Drive** (una vez)

**Lo que hace:**
- Nueva OP/OM con número automático (OP26-001, OM26-001...)
- Genera el Google Doc del formulario completo con 1 clic
- Drive organizado por cliente: `Mi eelo · Producción / Órdenes / Clientes / [nombre] /`
- Filtra todas las órdenes de un cliente con 1 clic
- Dashboard con estado de órdenes en tiempo real

---

## Sistema 2 — RRHH / Participantes
**Archivo:** `sistemas/RRHH_todo.gs`
**Google Sheets:** crea uno nuevo llamado "Mi eelo · RRHH"

**Cómo instalar:**
1. Abre el Sheets → **Extensiones → Apps Script**
2. Borra el código default
3. Pega TODO el contenido de `RRHH_todo.gs`
4. Edita las 3 líneas de configuración al inicio (token Kobo, correo, tarifa)
5. Guarda → recarga el Sheets → aparece el menú **👥 RRHH**
6. Ejecuta **Crear estructura en Drive** (una vez)
7. Ejecuta **Configurar automatizaciones** (una vez)

**Lo que hace:**
- Importa asistencia desde Kobo cada hora (automático)
- Empareja cada entrada con su salida → calcula horas reales
- Calcula montos por quincena automáticamente
- Genera recibos de pago en Docs por participante
- Correos automáticos: recordatorio de pagos (viernes) y resumen (día 1)
- Drive: `Mi eelo · RRHH / Facturas / 2026 / Mayo / ...`
