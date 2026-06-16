# 📋 Guía de Conocimiento — Sistema RRHH mi eelo

> **Versión:** 1.3 | **Última actualización:** 2026-06-16 | **Responsable:** Adrian (adrian@creamosguatemala.org)

---

## 🗂️ Índice

1. [¿Qué es este sistema?](#1-qué-es-este-sistema)
2. [Dónde vive](#2-dónde-vive)
3. [Estructura](#3-estructura)
4. [Cómo está construido](#4-cómo-está-construido)
5. [Flujo de trabajo](#5-flujo-de-trabajo)
6. [Qué se puede tocar y qué no](#6-qué-se-puede-tocar-y-qué-no)
7. [Cómo agregar o extender](#7-cómo-agregar-o-extender)
8. [Problemas conocidos y soluciones](#8-problemas-conocidos-y-soluciones)
9. [Historial de versiones](#9-historial-de-versiones)
10. [Contacto y mantenimiento](#10-contacto-y-mantenimiento)
11. [Cómo mantener esta guía actualizada](#11-cómo-mantener-esta-guía-actualizada)

---

## 1. ¿Qué es este sistema?

El Sistema RRHH mi eelo es una herramienta de **gestión de recursos humanos** diseñada específicamente para el taller textil **mi eelo**, parte de la organización **Creamos Guatemala**.

### Problema que resuelve

Sin este sistema, calcular cuánto pagar a cada participante cada quincena requería:
- Revisar manualmente los registros de entrada/salida del taller
- Calcular horas en Excel sin validación
- Llevar en papel quién tiene factura (IVA) y quién no
- Hacer los pagos sin un registro centralizado

El sistema automatiza todo ese proceso.

### ¿Para quién es?

| Usuario | Rol |
|---------|-----|
| Coordinador/a de mi eelo | Gestiona asistencia, genera reportes, registra pagos |
| Administración Creamos | Consulta pagos, cheques, transferencias |
| Nuevo coordinador/a | Puede aprender el flujo con esta guía |

### Qué hace exactamente

- Importa registros de entrada/salida desde **KoboToolbox** (app en tableta)
- Calcula horas trabajadas por persona por quincena
- Aplica tarifas por categoría (A/B/C/D) y suma IVA si corresponde
- Genera reportes de pago listos para cheques y transferencias bancarias
- Lleva registro de quién se retiró y por qué
- Documenta inclusión laboral, terapias y días de estudio

---

## 2. Dónde vive

### Repositorio

| Elemento | Detalle |
|----------|---------|
| **Repositorio GitHub** | `adrian-9856/Sistemas-de-mi-eelo` |
| **Rama de desarrollo** | `claude/gifted-fermi-8fm3T` |
| **Archivo principal** | `sistemas/RRHH_todo.gs` (~7,300 líneas) |
| **Documentación** | `docs/GUIA_CONOCIMIENTO_RRHH.md` |

### Accesos necesarios

| Acceso | Para qué |
|--------|----------|
| Google Spreadsheet de mi eelo | Donde vive el sistema completo |
| Google Apps Script (GAS) | Para pegar/actualizar el código |
| KoboToolbox (cuenta Creamos) | Fuente de datos de asistencia |
| Google Drive (cuenta Creamos) | Carpeta "Pagos RRHH mi eelo" para PDFs |
| Gmail (cuenta Creamos) | Envío de notificaciones de pago |

### Cómo abrir el editor de código

1. Abrir el Google Spreadsheet
2. Menú superior → **Extensiones** → **Apps Script**
3. Se abre el editor con el archivo `RRHH_todo.gs`

---

## 3. Estructura

### Hojas del Spreadsheet

#### Hojas principales (no borrar nunca)

| Hoja | Propósito | Columnas |
|------|-----------|----------|
| `PARTICIPANTES` | Lista maestra de todas las participantes | 19 cols A–S |
| `DatosKobo` | Registros de entrada/salida importados desde Kobo | Variable (depende de Kobo) |
| `PERIODOS` | Control de quincenas abiertas y cerradas | 7 cols |
| `Copy of CREAMOS ID nuevo` | Base de datos oficial de Creamos/Salesforce | Protegida, solo lectura |

#### Hojas de seguimiento

| Hoja | Propósito |
|------|-----------|
| `DiasEstudio` | Qué días cada participante tiene clases (no se pagan) |
| `ListaTerapias` | Quiénes reciben terapia (se marca con X) |
| `InclusionLaboral` | Quiénes participan en inclusión laboral |
| `HijosCCI` | Participantes con hijos en CCI |
| `Retiradx` | Registro de participantes que dejaron el programa |
| `CiclosVida` | Participantes que terminaron su ciclo de vida |
| `Bonos` | Bonos especiales individuales |

#### Hojas de pago

| Hoja | Propósito |
|------|-----------|
| `Cheques` | Registro de pagos por cheque (una fila por persona por mes) |
| `Transferencias` | Registro de transferencias bancarias |

#### Hojas temporales (se crean y eliminan solas)

| Hoja | Cuándo aparece |
|------|---------------|
| `Q_26_05_10_06_2026` (ej.) | Al generar reporte de quincena |
| `Rep_DIA_...` / `Rep_MES_...` | Al generar reportes por período |
| `DASHBOARD` | Dashboard visual con KPIs |

### Esquema de PARTICIPANTES (19 columnas)

| Col | Campo | Descripción |
|-----|-------|-------------|
| A | `Creamos_ID` | ID oficial de Salesforce (ej: MACI030373) |
| B | `Nombre` | Nombre completo oficial |
| C | `Proyecto` | Siempre "Textil" |
| D | `Programa` | Siempre "mi eelo" |
| E | `Etapa` | Inscritx / Retiradx / Empleadx / Ciclo de Vida Terminado |
| F | `Educacion` | Sí/No (auto desde hoja DiasEstudio) |
| G | `Apoyo_Emocional` | Sí/No (auto desde hoja ListaTerapias) |
| H | `Inclusion_Laboral` | Sí/No (auto desde hoja InclusionLaboral) |
| I | `Categoria` | A / B / C / D |
| J | `Tarifa_Hora` | Q/hr (se llena automático al poner Categoria) |
| K | `Tiene_Factura` | Sí / No (determina si aplica IVA 5%) |
| L | `DPI` | Número de DPI |
| M | `NIT` | NIT fiscal |
| N | `Correo` | Email de contacto |
| O | `Banco` | Banco para transferencia |
| P | `Tipo_Cuenta` | Monetaria / Ahorro |
| Q | `Num_Cuenta` | Número de cuenta bancaria |
| R | `Forma_Pago` | Cheque / Transferencia |
| S | `URL_Doc_Proceso` | Link al Documento de Proceso en Drive |

### Tarifas por categoría

| Categoría | Tarifa | Color |
|-----------|--------|-------|
| A | Q 16.50 / hora | 🟩 Verde |
| B | Q 15.75 / hora | 🟨 Amarillo/Dorado |
| C | Q 15.00 / hora | 🟧 Naranja |
| D | Q 14.00 / hora | 🟥 Rojo |

### Períodos de quincena

| Nombre | Días | Cuándo se paga |
|--------|------|---------------|
| **Q1 — Primera quincena** | Del 26 al 10 del siguiente mes | Al registrar Q1 |
| **Q2 — Segunda quincena** | Del 11 al 25 del mismo mes | Al registrar Q2 → se genera PDF y se cierra |

---

## 4. Cómo está construido

### Tecnología

| Componente | Tecnología |
|------------|-----------|
| Base de código | Google Apps Script (JavaScript) |
| Base de datos | Google Sheets (hojas del Spreadsheet) |
| Fuente de asistencia | KoboToolbox (API CSV) |
| Documentos | Google Docs (Documentos de Proceso, reportes PDF) |
| Almacenamiento de archivos | Google Drive |
| Notificaciones | Gmail vía MailApp |

### Arquitectura

Todo el sistema vive en **un solo archivo**: `sistemas/RRHH_todo.gs`.

```
┌─────────────────────────────────────────────────────────────┐
│                  Google Spreadsheet                          │
│                                                             │
│  PARTICIPANTES ←──────────────────────────────────────────┐ │
│  (lista maestra)      ↑  sync automático                  │ │
│                       │                                   │ │
│  DiasEstudio ─────────┤                                   │ │
│  ListaTerapias ───────┤                                   │ │
│  InclusionLaboral ────┘                                   │ │
│                                                           │ │
│  DatosKobo ←── API KoboToolbox (import cada hora)        │ │
│     ↓ generarReporte()                                    │ │
│  Hoja Q_... (reporte quincena)                           │ │
│     ↓ registrarPagosQuincena()                            │ │
│  Cheques + Transferencias ──→ Email notificación          │ │
│     ↓ (si Q2)                                             │ │
│  PDF en Drive + PERIODOS marcado "Pagado"                 │ │
└─────────────────────────────────────────────────────────────┘
```

### Decisiones de diseño importantes

1. **Un solo archivo GAS**: Más fácil de copiar/pegar entre entornos. No requiere conocimiento de módulos o imports.

2. **`_resolverTsKobo()`**: Cuando un formulario Kobo se abre un día y se envía al siguiente (ej: se dejó abierto en la tarde), el sistema usa la fecha de **envío** (`end`), no de apertura (`start`). Esto evita entradas "fantasma" en días incorrectos.

3. **Hoja protegida `Copy of CREAMOS ID nuevo`**: Nunca se borra ni se edita. Es la fuente oficial de IDs de Salesforce. El código la busca al crear/actualizar participantes.

4. **Estimado de 7 horas**: Si una participante tiene registro de entrada pero no de salida, se asume una jornada normal de 7 horas (`HORAS_JORNADA_NORMAL`). Se marca con `*` en el reporte.

5. **Zona horaria fija**: `CFG.TIMEZONE = "America/Guatemala"`. Guatemala no usa horario de verano (siempre UTC-6). Hardcodeado para evitar errores de conversión.

6. **IVA 5%**: Solo para participantes con `Tiene_Factura = "Sí"`. Régimen de Pequeño Contribuyente de Guatemala.

---

## 5. Flujo de trabajo

### Ciclo quincenal (cada 15 días)

```
Semana 1-2: Trabajo normal
    ↓
Participantes registran entrada/salida en tablet (KoboToolbox)
    ↓
Sistema importa datos automáticamente cada hora
    ↓
Día de corte (10 o 25 del mes):
    ↓
1. Menú 📥 Asistencia → Marcar filas sospechosas → revisar → Eliminar rojas
2. Menú 📅 Quincena → Generar reporte de quincena activa
3. Revisar reporte: horas por persona ¿están correctas?
4. Menú 📅 Quincena → Registrar pagos de quincena
5. Revisar hojas Cheques y Transferencias
6. Cobrar/subir transferencias → cambiar Status
7. Sistema archiva y envía email de confirmación
```

### Paso a paso detallado

#### 1. Importar datos de Kobo
- **Menú:** 📥 Asistencia → 📥 Importar desde Kobo
- Descarga todos los registros de entrada/salida de KoboToolbox
- Los datos quedan en la hoja `DatosKobo`
- Pasa automáticamente cada hora

#### 2. Limpiar datos malos (antes de cada cierre)
- **Menú:** 📥 Asistencia → 🟡 Marcar filas sospechosas
- Marca en **rojo** formularios que se abrieron en la tarde de un día y se enviaron al día siguiente (dato inválido)
- Marca en **amarillo** formularios enviados en menos de 30 minutos (posible error)
- Revisar las amarillas manualmente
- **Menú:** 📥 Asistencia → 🗑️ Eliminar filas marcadas en rojo

#### 3. Generar reporte de quincena
- **Menú:** 📅 Quincena → 📊 Ver / actualizar quincena actual
- Genera una hoja con todas las entradas/salidas del período
- Cada persona tiene su subtotal con: horas, tarifa, monto base, IVA, total

#### 4. Registrar pagos
- **Menú:** 📅 Quincena → 💳 Registrar pagos de quincena
- Seleccionar la quincena
- El sistema llena automáticamente las hojas `Cheques` y `Transferencias`
- Si es **Q1**: solo escribe los montos, espera Q2
- Si es **Q2**: genera PDF, elimina las hojas de reporte, marca período como "Pagado"

#### 5. Cerrar pagos
- En hoja `Cheques`: cuando todos los cheques están cobrados → cambiar Status a `Cobrado`
- En hoja `Transferencias`: cuando las transferencias se suben → cambiar Status a `Transferencias Subidas`
- El sistema archiva automáticamente y envía email de notificación

### Acciones administrativas

| Acción | Cuándo hacerlo | Menú |
|--------|---------------|------|
| Agregar participante nueva | Al incorporar alguien nuevo | ⚙️ Admin → ➕ Nuevo participante |
| Retirar participante | Al salir del programa | Cambiar Etapa a "Retiradx" en PARTICIPANTES |
| Actualizar categoría | Si cambia el nivel | ⚙️ Admin → 🔼 Cambiar categoría |
| Sincronizar IDs de Salesforce | Mensualmente | ⚙️ Admin → 🔄 Sincronizar desde Creamos DB |
| Recalcular tarifas | Después de migrar o cambiar cats | ⚙️ Admin → 🔄 Recalcular tarifas y factura |

---

## 6. Qué se puede tocar y qué no

### ✅ Se puede modificar libremente

| Elemento | Cómo |
|----------|------|
| Categoría (A/B/C/D) en PARTICIPANTES col I | Escribir directo; tarifa se actualiza sola |
| Tiene_Factura (Sí/No) en PARTICIPANTES col K | Escribir directo |
| Datos bancarios (Banco, Tipo_Cuenta, Num_Cuenta) | Escribir directo en PARTICIPANTES |
| Correo y DPI de participante | Escribir directo en PARTICIPANTES |
| Filas en DiasEstudio (X o vacío) | Cambiar qué días tiene estudio |
| Filas en ListaTerapias (X o vacío) | Marcar quién recibe terapia |
| Filas en InclusionLaboral (X o vacío) | Marcar quién participa |
| Status en Cheques (Pendiente/Cobrado/Cancelado) | Cambiar cuando se cobra |
| Status en Transferencias | Cambiar cuando se sube |
| Notas en Retiradx | Agregar información adicional |

### ⚠️ Modificar con cuidado

| Elemento | Por qué cuidar |
|----------|---------------|
| Nombres en PARTICIPANTES | Deben coincidir exactamente con Kobo. Si cambia, actualizar mapeo |
| Fecha de inicio/fin en PERIODOS | Afecta cálculo de todas las quincenas |
| Columna `start` en DatosKobo | Es la fuente de verdad de asistencia; no editar sin razón |
| Código en Apps Script | Cambios mal hechos rompen todo el sistema |

### 🚫 Nunca tocar

| Elemento | Por qué |
|----------|---------|
| Hoja `Copy of CREAMOS ID nuevo` | Base de datos oficial de Salesforce. Si se borra o edita, se pierden todos los IDs |
| Función `reinstalarSistema()` en producción | Borra absolutamente todo. Solo usar en entorno de pruebas |
| Estructura de columnas de PARTICIPANTES (A–S) | El código usa índices numéricos. Si se mueven columnas, todo se rompe |
| Triggers instalados | No borrar los triggers de `onEditInstalable` o `importarDesdeKobo`; son los que hacen funcionar emails e importación automática |

---

## 7. Cómo agregar o extender

### Agregar una nueva participante

1. **Menú:** ⚙️ Admin → ➕ Nuevo participante
2. El sistema busca automáticamente en `Copy of CREAMOS ID nuevo`
3. Si no la encuentra: aparece advertencia "⚠️ Crear perfil en Salesforce"
4. Llenar manualmente: Categoría (col I), Tiene_Factura (col K), datos bancarios
5. Si tiene terapia, días de estudio o inclusión laboral: marcar en las hojas respectivas

### Agregar una nueva hoja de seguimiento

Seguir el patrón de `crearHojaInclusionLaboral()` en el código:

```javascript
function crearHojaNueva() { _run(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("NombreHoja");
  var esNueva = !hoja;
  if (esNueva) hoja = ss.insertSheet("NombreHoja");

  var enc = ["Creamos_ID","Participante","Valor","Notas"];
  hoja.getRange(1,1,1,4).setValues([enc])
    .setBackground("#COLOR").setFontColor("#ffffff").setFontWeight("bold");
  hoja.setFrozenRows(1);
  // ... pre-llenar desde PARTICIPANTES
  _alert("✅ Hoja NombreHoja creada.");
}); }
```

Luego agregar al menú en `onOpen()` y a `instalarTodo()`.

### Agregar una nueva funcionalidad al menú

En la función `onOpen()`, buscar el submenú correspondiente y agregar:

```javascript
.addItem("📌 Nombre visible en menú", "nombreDeLaFuncion")
```

### Cambiar las tarifas por categoría

Editar en el objeto `CFG` al inicio del archivo:

```javascript
CATEGORIAS: { A: 16.50, B: 15.75, C: 15.00, D: 14.00 },
```

Después de cambiar: **Menú ⚙️ Admin → 🔄 Recalcular tarifas y factura** para actualizar todos los valores en PARTICIPANTES.

### Agregar un nuevo correo de notificación

Editar en `CFG`:

```javascript
CORREO_CHEQUES:  "encargada.cheques@creamosguatemala.org",
CORREO_PLANILLA: "encargada.planilla@creamosguatemala.org",
```

### Cambiar la duración de la jornada estimada

Si se cambia la política de horas por día (actualmente 7):

```javascript
HORAS_JORNADA_NORMAL: 7,  // ← cambiar aquí
```

---

## 8. Problemas conocidos y soluciones

| # | Síntoma | Causa | Solución |
|---|---------|-------|----------|
| 1 | Reporte muestra entradas en día incorrecto | Formulario Kobo abierto el día anterior en la tarde y enviado al día siguiente | Menú → 🟡 Marcar sospechosas → 🗑️ Eliminar rojas → Regenerar |
| 2 | Filas con 0 horas (ej: 9:04–9:04) | ENTRADA y SALIDA con mismo timestamp (formularios cruzados) | Código los convierte automáticamente a "Normal (Estimado)* 7 hrs" |
| 3 | Categoría muestra "Cat. NO" en reporte | Col I de PARTICIPANTES vacía o con esquema viejo | ⚙️ Admin → ⬆️ Migrar sistema → llenar col I → 🔄 Recalcular tarifas |
| 4 | IVA = Q0 para quien debería tenerlo | `Tiene_Factura` no dice "Sí" exactamente | Verificar col K. ⚙️ Admin → 🔄 Recalcular tarifas rellena los vacíos con "Sí" |
| 5 | Email de pagos no se envía | Los triggers no están instalados | ⚙️ Admin → ⚡ Activar automatizaciones |
| 6 | Importación Kobo falla | URL de exportación cambió o venció el token | Actualizar `CFG.KOBO_URL_CSV` en el código |
| 7 | Participante sin Creamos_ID (muestra ⚠️) | No tiene perfil en Salesforce | Crear perfil en Salesforce y sincronizar |
| 8 | Nombre en Kobo no coincide con PARTICIPANTES | Kobo usa slug (ej: `ana_rebeca_...`) | ⚙️ Admin → 📲 Sincronizar IDs desde DatosKobo |
| 9 | Totales de reporte y quincena no coinciden | Se editó manualmente una de las hojas | No editar hojas generadas. Limpiar Kobo y regenerar ambas |
| 10 | PARTICIPANTES tiene columnas en orden incorrecto | Se usó código viejo (22 cols) con datos nuevos (19 cols) | ⚙️ Admin → ⬆️ Migrar sistema |
| 11 | Quincena marcada como Q1 siendo Q2 o viceversa | Error de configuración de período | El período del 26 al 10 = Q1. Del 11 al 25 = Q2. Verificar fechas en PERIODOS |
| 12 | Horas en reporte en horario UTC (6 horas de diferencia) | Script configurado en UTC en lugar de Guatemala | Verificar que `CFG.TIMEZONE = "America/Guatemala"` esté en el código |

---

## 9. Historial de versiones

| Fecha | Versión | Cambios principales |
|-------|---------|---------------------|
| 2026-06-16 | 1.3 | Fix entradas 0 horas → estimado automático. Marcar/eliminar filas sospechosas en Kobo. Filtro ENTRADA PM (≥12:00) mejorado con hora local Guatemala. |
| 2026-06-16 | 1.2 | `recalcularTarifas()` batch update. Mejoras en `migrarSistema()` con diagnóstico de categorías. Menú Admin: nuevo item "Recalcular tarifas y factura". |
| 2026-06-16 | 1.1 | Fix entradas fantasma: `_resolverTsKobo()` usa `end` cuando formulario cruzó medianoche. Fix umbral PM de 17→12 hora local. |
| 2026-06-15 | 1.0 | Schema PARTICIPANTES 22→19 cols (removidos Hijos_CCI, Num_Hijos_CCI, Estipendio). Nueva hoja InclusionLaboral. Nueva hoja Retiradx con 22 razones de retiro. Propagación IDs a hojas auxiliares. Defaults: Tiene_Factura="Sí", Etapa="Inscritx". Timezone fijo América/Guatemala. Timestamp usa `start` no `end`. |
| 2026-06-14 | 0.9 | Sistema de pagos: Cheques y Transferencias con flujo Q1/Q2. Email automático con trigger instalable. PDF al cerrar Q2. Corrección quincenas: Q1=26→10, Q2=11→25. |
| 2026-06-13 | 0.8 | Workflow completo de pago: hoja Cheques (Q1+Q2 por mes), hoja Transferencias, archivo PDF en Drive, marcado "Pagado" en PERIODOS. |
| 2026-06-12 | 0.7 | Deduplicación day-level: mantiene primera entrada y última salida del día. Fix reporte vs quincena: ambos usan agrupación por día. |
| 2026-06-10 | 0.5 | Sistema base: PARTICIPANTES, DatosKobo, reportes, integración Kobo, cálculo de horas y tarifas. |

---

## 10. Contacto y mantenimiento

| Rol | Persona | Contacto |
|-----|---------|----------|
| Responsable principal | Adrian | adrian@creamosguatemala.org |
| Desarrollo/soporte técnico | Claude Code (IA) | Vía sesión en claude.ai/code |
| Problemas de Salesforce/Creamos DB | Equipo Creamos | Contacto interno |
| Problemas de KoboToolbox | Cuenta Creamos | Acceso a kf.kobotoolbox.org |

### Cómo reportar un problema

1. Anotar exactamente qué hizo y qué resultado obtuvo
2. Si hay un mensaje de error, copiarlo completo
3. Tomar captura de pantalla del estado de las hojas
4. Enviar a adrian@creamosguatemala.org con asunto "BUG Sistema RRHH - [descripción breve]"

### Para hacer cambios al código

1. Abrir el repositorio en GitHub: `adrian-9856/Sistemas-de-mi-eelo`
2. Rama activa: `claude/gifted-fermi-8fm3T`
3. Los cambios siempre van a `sistemas/RRHH_todo.gs`
4. Después de cada cambio: copiar el contenido completo y pegarlo en Apps Script → Guardar
5. Actualizar esta guía en la sección de Historial de Versiones

---

## 11. Cómo mantener esta guía actualizada

### Cuándo actualizar la guía

Actualizar esta guía siempre que ocurra cualquiera de estas situaciones:

| Evento | Secciones a actualizar |
|--------|----------------------|
| Se agrega una nueva hoja | §3 Estructura, §5 Flujo de trabajo, §7 Cómo extender |
| Se cambian las tarifas | §3 Tarifas por categoría |
| Se cambia el esquema de PARTICIPANTES | §3 Esquema de columnas, §4 Decisiones de diseño |
| Se resuelve un bug conocido | §8 Problemas conocidos (actualizar la fila del error) |
| Se agrega una nueva función al menú | §5 Flujo de trabajo, §7 Cómo extender |
| Cambia un correo o configuración | §2 Accesos necesarios, §10 Contacto |
| Se descubre un problema nuevo | §8 Agregar fila nueva con causa y solución |

### Cómo registrar en el historial

Agregar una fila nueva en la tabla del §9 con este formato:

```markdown
| YYYY-MM-DD | X.Y | Descripción breve del cambio en una línea |
```

Incrementar la versión así:
- Cambio pequeño (fix, ajuste): incrementar el número decimal (1.2 → 1.3)
- Cambio importante (nueva funcionalidad, rediseño): incrementar el entero (1.x → 2.0)

### Actualizar la versión al inicio del documento

La primera línea de la guía tiene:

```
> **Versión:** X.Y | **Última actualización:** YYYY-MM-DD | **Responsable:** ...
```

Actualizarla con cada cambio.

### Regla de oro

> Si alguien nuevo puede leer esta guía y operar el sistema sin preguntar nada,
> la guía está al día. Si necesita preguntar, hay algo que documentar.

---

*Guía generada con Claude Code — claude.ai/code*
*Repositorio: `adrian-9856/Sistemas-de-mi-eelo` | Rama: `claude/gifted-fermi-8fm3T`*
