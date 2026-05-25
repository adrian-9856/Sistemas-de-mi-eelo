# 📋 PLAN DE IMPLEMENTACIÓN - SISTEMA MI EELO GOOGLE SHEETS

**Objetivo:** Centralizar todos tus archivos Excel en un solo Google Sheets automatizado con Apps Script.

**Tiempo estimado:** 4-6 horas (puedes hacerlo por fases)

---

## 🎯 RESUMEN EJECUTIVO

### LO QUE TIENES AHORA (Fragmentado)
- ❌ **6 archivos Excel separados** con información duplicada
- ❌ **141 hojas de OPs/OMs** (una por orden)
- ❌ **14 hojas de facturación** (una por mes)
- ✅ **Sistema de asistencia KoboToolbox** ya funcionando con Apps Script

### LO QUE TENDRÁS AL FINAL (Centralizado)
- ✅ **1 solo archivo Google Sheets** con 6 hojas principales
- ✅ **Facturación automática** desde horas trabajadas
- ✅ **Búsquedas instantáneas** con filtros
- ✅ **Sin duplicación** de datos
- ✅ **Actualización automática** cada hora desde Kobo

---

## 📁 FASE 1: CREAR LA ESTRUCTURA BASE (30 min)

### Paso 1.1: Crear nuevo Google Sheets
1. Ve a https://sheets.google.com
2. Clic en **"Nuevo" → "Hoja de cálculo en blanco"**
3. Nómbralo: **"Mi eelo - Sistema Unificado"**

### Paso 1.2: Crear las 6 hojas principales
Renombra la primera hoja y crea las otras 5:

1. **PARTICIPANTES_MASTER**
2. **ASISTENCIA_KOBO** 
3. **FACTURACION_CONSOLIDADA**
4. **ORDENES_PRODUCCION**
5. **COMERCIAL_VENTAS**
6. **COMERCIAL_COMPRAS**

**Cómo crear hojas:**
- Abajo a la izquierda, clic en el **"+"** para añadir hoja
- Clic derecho en la pestaña → "Renombrar"

---

## 👥 FASE 2: HOJA PARTICIPANTES_MASTER (1 hora)

### Paso 2.1: Crear encabezados (Fila 1)
Copia estos encabezados en la primera fila:

```
A: ID_Participante
B: Creamos_ID
C: Nombre_Completo
D: Proyecto
E: Division
F: Estado
G: DPI
H: NIT
I: Correo_Electronico
J: Datos_SAT_Usuario
K: Datos_SAT_Password
L: Forma_Pago
M: Banco
N: Numero_Cuenta
```

### Paso 2.2: Copiar datos desde tus archivos Excel

**Origen 1: Participantes___mi_eelo_26.xlsx**
1. Abre el archivo Excel
2. Copia toda la hoja "Listado Participantes"
3. Pega temporalmente en Google Sheets en columnas Q, R, S... (a la derecha)
4. Luego arrastra manualmente cada dato a su columna correcta:
   - Nombre → columna C
   - Creamos ID → columna B
   - Proyecto → columna D
   - División → columna E
   - Estado → columna F

**Origen 2: Facturación_Mi_eelo.xlsx → Hoja "Datos de Facturación"**
1. Abre el archivo Excel
2. Copia la hoja "Datos de Facturación"
3. Pega en columnas temporales
4. Arrastra cada dato:
   - Nombre → busca coincidencia en columna C (mismo nombre)
   - DPI → columna G
   - NIT → columna H
   - Correo → columna I
   - Datos SAT → columnas J, K

### Paso 2.3: Crear columna ID_Participante (columna A)
En la celda **A2** escribe:
```
=TEXTO(FILA()-1,"P000")
```
Arrastra la fórmula hacia abajo hasta la última persona.

Esto genera: P001, P002, P003...

### Paso 2.4: Formatear tabla
1. Selecciona toda la fila 1 (encabezados)
2. **Formato → Negrita + Color de fondo naranja**
3. **Datos → Crear filtro** (para poder filtrar después)
4. Ajusta ancho de columnas

---

## ⏰ FASE 3: HOJA ASISTENCIA_KOBO (15 min)

### Paso 3.1: Crear encabezados

```
A: ID_Participante
B: Nombre_Participante
C: Fecha_Hora_Entrada
D: Fecha_Hora_Salida
E: Tipo_Registro
F: Horas_Trabajadas
G: Es_Dia_Estudio
H: Es_Terapia
I: Porcentaje_Pago
J: Horas_A_Pagar
K: UUID_Kobo
```

### Paso 3.2: Copiar script actual
1. **Extensiones → Apps Script**
2. Copia el código de tu archivo `Planilla___mi_eelo.xlsx` que ya tienes
3. El script ya importa desde KoboToolbox automáticamente
4. **NO cambies nada todavía**, solo pégalo

### Paso 3.3: Conectar con PARTICIPANTES_MASTER
Más adelante añadiremos una columna que use VLOOKUP para traer el ID desde el nombre.

---

## 💰 FASE 4: HOJA FACTURACION_CONSOLIDADA (1.5 horas)

### Paso 4.1: Crear encabezados

```
A: ID_Participante
B: Nombre_Participante
C: Mes
D: Año
E: Quincena
F: Horas_Trabajadas
G: Monto_A_Pagar
H: Factura_Entregada
I: Numero_Factura
J: Declaraguate
K: Pagado
L: Fecha_Pago
M: Comentarios
```

### Paso 4.2: Copiar datos de facturación (las 14 hojas)
Abre **Facturación_Mi_eelo.xlsx** y por cada hoja (Mayo, Junio, Julio...):

1. Copia los datos de cada mes
2. Pega en FACTURACION_CONSOLIDADA
3. Añade manualmente en columna C el mes ("Mayo", "Junio"...)
4. Añade en columna D el año (2025 o 2026)
5. Repite con las 14 hojas

**Truco:** Pega todos seguidos uno debajo del otro. Al final tendrás ~300 filas.

### Paso 4.3: Conectar con PARTICIPANTES_MASTER
En la celda **A2** (primera fila de datos):
```
=VLOOKUP(B2, PARTICIPANTES_MASTER!C:A, 1, FALSE)
```

Esto busca el nombre en columna B y trae el ID_Participante.

En la celda **B2** (si no pegaste el nombre):
```
=VLOOKUP(A2, PARTICIPANTES_MASTER!A:C, 3, FALSE)
```

### Paso 4.4: Calcular horas automáticamente (FUTURO)
Más adelante, columna F usará:
```
=SUMIFS(ASISTENCIA_KOBO!J:J, ASISTENCIA_KOBO!A:A, A2, ASISTENCIA_KOBO!C:C, ">="&FECHA(D2,COINCIDIR(C2,{"Enero";"Febrero";"Marzo";"Abril";"Mayo";"Junio";"Julio";"Agosto";"Septiembre";"Octubre";"Noviembre";"Diciembre"},0),1))
```

**Por ahora déjalo vacío**, lo configuraremos después.

---

## 📦 FASE 5: HOJA ORDENES_PRODUCCION (2 horas)

### Paso 5.1: Crear encabezados

```
A: Numero_Orden
B: Tipo_Orden
C: Fecha_Creacion
D: Cliente
E: Contacto_Cliente
F: Descripcion
G: Cantidad
H: Estado
I: Participantes_Asignados
J: Horas_Dedicadas
K: Fecha_Entrega
L: Notas
```

### Paso 5.2: Copiar datos de OPs (73 hojas)
Abre **OPs_26____mi_eelo.xlsx**:

Para cada hoja (001, 002, 003...):
1. Identifica en cada hoja:
   - Cliente (generalmente fila 8)
   - Contacto (fila 11)
   - Descripción (fila 14)
2. Copia manualmente a ORDENES_PRODUCCION
3. En columna A escribe: OP-001, OP-002, OP-003...
4. En columna B escribe: OP

**Advertencia:** Esto es tedioso. Alternativa más rápida:
- Copia solo las 20 órdenes más recientes/importantes
- Las demás déjalas en el archivo Excel original como "archivo histórico"

### Paso 5.3: Copiar datos de OMs (68 hojas)
Igual que arriba pero:
- Columna A: OM-001, OM-002...
- Columna B: OM

### Paso 5.4: Añadir validación de datos
Selecciona columna H (Estado):
1. **Datos → Validación de datos**
2. **Criterios:** Lista de elementos
3. Escribe: `Pendiente, En Proceso, Completada, Cancelada`
4. **Guardar**

---

## 🛒 FASE 6: HOJAS COMERCIALES (30 min)

### Paso 6.1: COMERCIAL_VENTAS
Encabezados:
```
A: Fecha
B: Numero_Orden
C: Cliente
D: Producto
E: Cantidad
F: Precio_Unitario
G: Total
H: Tipo_Venta
I: Factura
J: Estado_Pago
```

Copia datos de **_26_Comercial___mi_eelo.xlsx → Hoja "Sales"**

### Paso 6.2: COMERCIAL_COMPRAS
Encabezados:
```
A: Fecha
B: Numero_Orden
C: Proveedor
D: Descripcion
E: Monto
F: Tipo_Compra
G: Factura_DTE
H: Pago_Realizado
```

Copia datos de **_26_Comercial___mi_eelo.xlsx → Hoja "Compras"**

---

## 🔗 FASE 7: CONECTAR TODO CON FÓRMULAS (1 hora)

### Paso 7.1: En ASISTENCIA_KOBO
**Columna A (ID_Participante):**
```
=VLOOKUP(B2, PARTICIPANTES_MASTER!C:A, 1, FALSE)
```
Busca el nombre en PARTICIPANTES_MASTER y trae el ID.

**Columna I (Porcentaje_Pago):**
```
=SI(G2="Sí", 0, SI(H2="Sí", 100, 100))
```
- Si es día de estudio → 0%
- Si es terapia → 100%
- Normal → 100%

**Columna J (Horas_A_Pagar):**
```
=F2 * (I2/100)
```

### Paso 7.2: En FACTURACION_CONSOLIDADA
**Columna F (Horas_Trabajadas) - automático desde ASISTENCIA:**
```
=SUMIFS(ASISTENCIA_KOBO!J:J, ASISTENCIA_KOBO!A:A, A2, MES(ASISTENCIA_KOBO!C:C), COINCIDIR(C2,{"Enero";"Febrero";"Marzo";"Abril";"Mayo";"Junio";"Julio";"Agosto";"Septiembre";"Octubre";"Noviembre";"Diciembre"},0))
```

**Columna G (Monto_A_Pagar):**
```
=F2 * 25
```
(Ajusta el 25 por tu tarifa por hora)

### Paso 7.3: En ORDENES_PRODUCCION
**Columna J (Horas_Dedicadas):**
```
=SUMIF(ASISTENCIA_KOBO!L:L, A2, ASISTENCIA_KOBO!J:J)
```
(Primero necesitas añadir columna L en ASISTENCIA_KOBO que diga en qué orden trabajaron)

---

## ⚙️ FASE 8: CONFIGURAR APPS SCRIPT (1 hora)

### Paso 8.1: Abrir editor
**Extensiones → Apps Script**

### Paso 8.2: Pegar tu código actual
Copia todo el código de `Planilla___mi_eelo.xlsx` que ya tienes funcionando.

### Paso 8.3: Modificar para que importe a la hoja correcta
Busca esta línea:
```javascript
var hoja = spreadsheet.getSheetByName("DatosKobo");
```

Cámbiala por:
```javascript
var hoja = spreadsheet.getSheetByName("ASISTENCIA_KOBO");
```

### Paso 8.4: Configurar trigger automático
1. En Apps Script, clic en **⏰ (Activadores)** en el menú izquierdo
2. **Añadir activador**
3. Configuración:
   - Función: `importarCSVdesdeKobo`
   - Origen del evento: **Basado en tiempo**
   - Tipo de activador: **Temporizador por horas**
   - Intervalo: **Cada hora**
4. **Guardar**

### Paso 8.5: Añadir función de normalización
Añade esta función nueva al final del script:

```javascript
function normalizarIDsParticipantes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaAsist = ss.getSheetByName("ASISTENCIA_KOBO");
  var hojaPart = ss.getSheetByName("PARTICIPANTES_MASTER");
  
  var datosAsist = hojaAsist.getDataRange().getValues();
  var datosPart = hojaPart.getDataRange().getValues();
  
  // Crear mapeo Nombre → ID
  var mapeo = {};
  for (var i = 1; i < datosPart.length; i++) {
    var id = datosPart[i][0];    // Columna A
    var nombre = datosPart[i][2]; // Columna C
    mapeo[nombre.toLowerCase().trim()] = id;
  }
  
  // Llenar columna A de ASISTENCIA_KOBO
  for (var i = 1; i < datosAsist.length; i++) {
    var nombreAsist = String(datosAsist[i][1]).toLowerCase().trim(); // Columna B
    var idEncontrado = mapeo[nombreAsist];
    
    if (idEncontrado) {
      hojaAsist.getRange(i + 1, 1).setValue(idEncontrado); // Columna A
    }
  }
  
  SpreadsheetApp.getUi().alert('✅ IDs normalizados: ' + (datosAsist.length - 1) + ' registros');
}
```

### Paso 8.6: Crear menú personalizado
Añade esta función al principio:

```javascript
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Mi eelo')
    .addItem('📥 Importar desde Kobo', 'importarCSVdesdeKobo')
    .addItem('🔗 Normalizar IDs', 'normalizarIDsParticipantes')
    .addItem('📊 Generar Reporte Mensual', 'generarReporteMensual')
    .addSeparator()
    .addItem('⚙️ Configurar Actualización Automática', 'configurarActualizacionAutomatica')
    .addToUi();
}
```

---

## ✅ FASE 9: PRUEBAS Y VALIDACIÓN (30 min)

### Paso 9.1: Prueba manual
1. Ve al menú **🚀 Mi eelo → 📥 Importar desde Kobo**
2. Espera 10-20 segundos
3. Verifica que aparezcan datos en ASISTENCIA_KOBO

### Paso 9.2: Prueba VLOOKUP
En PARTICIPANTES_MASTER, busca una persona:
1. Copia su nombre exacto
2. Ve a ASISTENCIA_KOBO
3. Busca (Ctrl+F) ese nombre
4. Verifica que en columna A aparezca su ID (ej: P015)

### Paso 9.3: Prueba suma de horas
1. Ve a FACTURACION_CONSOLIDADA
2. Verifica que columna F (Horas_Trabajadas) tenga números
3. Si sale `#N/A` o `#ERROR`, revisa las fórmulas

---

## 🎨 FASE 10: FORMATO Y DASHBOARD (30 min - OPCIONAL)

### Paso 10.1: Colores por hoja
- **PARTICIPANTES_MASTER:** Encabezado naranja (#ff6f00)
- **ASISTENCIA_KOBO:** Encabezado azul (#1f54a8)
- **FACTURACION_CONSOLIDADA:** Encabezado verde (#639922)
- **ORDENES_PRODUCCION:** Encabezado morado (#7F77DD)

### Paso 10.2: Crear hoja DASHBOARD (opcional)
Encabezados:
```
A: Métrica
B: Valor
```

Ejemplos:
```
A2: Total Participantes
B2: =CONTAR.SI(PARTICIPANTES_MASTER!F:F, "Activo")

A3: Horas Trabajadas (Mes Actual)
B3: =SUMAR.SI(MES(ASISTENCIA_KOBO!C:C), MES(HOY()), ASISTENCIA_KOBO!J:J)

A4: Órdenes Pendientes
B4: =CONTAR.SI(ORDENES_PRODUCCION!H:H, "Pendiente")
```

### Paso 10.3: Gráficos
1. Selecciona datos de FACTURACION_CONSOLIDADA (columnas C y F)
2. **Insertar → Gráfico**
3. Tipo: **Gráfico de columnas**
4. Título: "Horas trabajadas por mes"

---

## 📱 FASE 11: COMPARTIR Y PERMISOS (10 min)

### Paso 11.1: Compartir con tu equipo
1. **Compartir** (botón arriba a la derecha)
2. Añade correos de tu equipo
3. Permisos:
   - **Editores:** pueden modificar datos
   - **Lectores:** solo pueden ver

### Paso 11.2: Proteger hojas sensibles
1. Clic derecho en pestaña **PARTICIPANTES_MASTER**
2. **Proteger hoja**
3. **Rango:** Toda la hoja
4. **Permisos:** Solo tú puedes editar
5. **Listo**

---

## 🔄 MANTENIMIENTO DIARIO (5 min/día)

### Checklist diario:
- [ ] Verificar que el script importó datos nuevos (ver última fila de ASISTENCIA_KOBO)
- [ ] Revisar ausencias (filtrar ASISTENCIA_KOBO por fecha = hoy)
- [ ] Actualizar estado de órdenes en ORDENES_PRODUCCION

### Checklist quincenal:
- [ ] Verificar horas en FACTURACION_CONSOLIDADA
- [ ] Marcar facturas entregadas
- [ ] Generar reporte quincenal

---

## 🆘 SOLUCIÓN DE PROBLEMAS

### Problema: "El script no importa datos de Kobo"
**Solución:**
1. Ve a **Extensiones → Apps Script**
2. Clic en **▶️ Ejecutar** (arriba)
3. Si dice "Autorización requerida" → **Revisar permisos** → **Permitir**

### Problema: "VLOOKUP da error #N/A"
**Solución:**
1. Verifica que los nombres sean EXACTAMENTE iguales
2. En PARTICIPANTES_MASTER, columna C, elimina espacios extra:
   ```
   =ESPACIOS(C2)
   ```

### Problema: "Las horas no se suman bien"
**Solución:**
1. Verifica que las fechas en ASISTENCIA_KOBO tengan formato de fecha (no texto)
2. Selecciona columna C → **Formato → Número → Fecha y hora**

---

## 📚 RECURSOS ADICIONALES

### Tutoriales recomendados:
- VLOOKUP: https://support.google.com/docs/answer/3093318
- Apps Script básico: https://developers.google.com/apps-script/guides/sheets
- Fórmulas de Google Sheets: https://support.google.com/docs/table/25273

### Atajos de teclado útiles:
- **Ctrl + F:** Buscar
- **Ctrl + H:** Buscar y reemplazar
- **Ctrl + Alt + Shift + H:** Ver historial de versiones
- **Ctrl + K:** Insertar enlace

---

## ✅ CHECKLIST FINAL

Marca cuando completes cada fase:

- [ ] Fase 1: Estructura base creada (6 hojas)
- [ ] Fase 2: PARTICIPANTES_MASTER completa con 76 personas
- [ ] Fase 3: ASISTENCIA_KOBO configurada
- [ ] Fase 4: FACTURACION_CONSOLIDADA con datos de 14 meses
- [ ] Fase 5: ORDENES_PRODUCCION con OPs y OMs
- [ ] Fase 6: Hojas comerciales creadas
- [ ] Fase 7: Fórmulas VLOOKUP funcionando
- [ ] Fase 8: Apps Script configurado y funcionando
- [ ] Fase 9: Pruebas pasadas ✅
- [ ] Fase 10: Formato y colores aplicados
- [ ] Fase 11: Compartido con equipo

---

## 🎉 RESULTADO FINAL

Al completar todas las fases tendrás:

✅ **Un solo archivo Google Sheets** en lugar de 6 archivos Excel  
✅ **Importación automática** desde KoboToolbox cada hora  
✅ **Facturación automática** calculada desde horas trabajadas  
✅ **Búsquedas instantáneas** con filtros  
✅ **Trazabilidad completa** de participantes → asistencia → facturación → órdenes  
✅ **Acceso desde cualquier dispositivo** (PC, tablet, móvil)  
✅ **Historial de cambios** automático (Ctrl + Alt + Shift + H)  

---

**¿Necesitas ayuda?** Guarda este archivo y úsalo como guía paso a paso.

**Consejo:** No intentes hacer todo en un día. Completa 2-3 fases por día durante una semana.

---

*Documento creado: 2026-05-24*  
*Versión: 1.0*  
*Sistema: Mi eelo - Google Sheets Unificado*
