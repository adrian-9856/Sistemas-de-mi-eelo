# Carga inicial de datos históricos

Sigue este orden exacto la primera vez. Después todo es automático.

---

## Paso 1 — Participantes (desde Participantes___mi_eelo_26.xlsx)

1. Abre el archivo Excel
2. Copia toda la hoja "Listado Participantes" (incluyendo encabezados)
3. En tu Google Sheets, crea una hoja nueva llamada exactamente: `IMPORT_PART`
4. Pega los datos ahí (Ctrl+Shift+V → solo valores)
5. Ejecuta **🚀 Mi eelo → Carga inicial → 👥 Importar participantes desde Excel**
6. Cuando termine, puedes borrar la hoja `IMPORT_PART`

Resultado: 76 participantes en `PARTICIPANTES_MASTER`

---

## Paso 2 — Asistencia histórica (desde Planilla___mi_eelo.xlsx)

1. Abre el archivo Excel
2. Copia toda la hoja **DatosKobo** (tiene ~6,990 filas)
3. En tu Google Sheets, crea una hoja nueva llamada exactamente: `DatosKobo`
4. Pega los datos ahí (Ctrl+Shift+V → solo valores)
5. Ejecuta **🚀 Mi eelo → Carga inicial → 📋 Importar asistencia desde Planilla**
6. El sistema extrae el Creamos ID de cada fila y empareja entradas/salidas
7. Puedes dejar la hoja `DatosKobo` o borrarla después

También copia las hojas auxiliares del mismo Excel:
- `DiasEstudio` → copia tal cual al Google Sheets
- `ListaTerapias` → copia tal cual al Google Sheets
- `NombresCanonicos` → copia tal cual al Google Sheets

Estas hojas se usan para calcular el porcentaje de pago correcto.

---

## Paso 3 — Órdenes de Producción (desde OPs_26___mi_eelo.xlsx)

Las OPs tienen 73 hojas (001 a 073), una por orden. Tienes dos opciones:

### Opción A: Automática (recomendada si tienes acceso a copiar en lote)
1. Copia **todas las hojas** del archivo OPs al Google Sheets  
   (Clic derecho en cada pestaña → Copiar a → Tu spreadsheet)
2. Ejecuta **🚀 Mi eelo → Carga inicial → 📦 Importar OPs / OMs**
3. Escribe `OP` cuando te pregunte el tipo
4. Repite con el archivo OMs (escribe `OM`)

### Opción B: Manual (solo las más recientes)
- Copia solo las últimas 20 órdenes directamente a `ORDENES_PRODUCCION`
- Usa las columnas: Numero_Orden | Tipo | Fecha | Cliente | Contacto | Descripcion | Cantidad | Tela | Estado | Participantes | Horas | Fecha_Entrega | Proyecto | Comentarios

---

## Paso 4 — Comercial (desde _26_Comercial___mi_eelo.xlsx)

1. Copia la hoja **"Sales "** (con espacio) al Google Sheets → nómbrala `Sales`
2. Copia la hoja **"Compras"** al Google Sheets → déjala como `Compras`
3. Ejecuta **🚀 Mi eelo → Carga inicial → 💼 Importar ventas**
4. Ejecuta **🚀 Mi eelo → Carga inicial → 🛒 Importar compras**

---

## Paso 5 — Calcular facturación histórica

Una vez que tienes la asistencia importada:

1. Ejecuta **💰 Calcular facturación del mes** para el mes actual
2. Para meses anteriores: modifica temporalmente la fecha en la función
   `calcularFacturacionMes()` en `04_Facturacion.gs` y vuelve a ejecutar

---

## Paso 6 — Configurar automatizaciones

Ejecuta **⚙️ Configurar triggers automáticos** — solo una vez.

Desde ese momento:
- Kobo se importa cada hora
- La facturación se recalcula diariamente
- El dashboard se actualiza cada 30 min

---

## Mapa de archivos → hojas del sistema

| Archivo Excel original          | Hoja en Google Sheets           | Método de carga      |
|---------------------------------|---------------------------------|----------------------|
| Participantes___mi_eelo_26.xlsx | PARTICIPANTES_MASTER            | Función automática   |
| Planilla___mi_eelo.xlsx (DatosKobo) | ASISTENCIA_KOBO             | Función automática   |
| Planilla___mi_eelo.xlsx (DiasEstudio) | DiasEstudio              | Pegar manualmente    |
| Planilla___mi_eelo.xlsx (ListaTerapias) | ListaTerapias           | Pegar manualmente    |
| OPs_26___mi_eelo.xlsx (001-073) | ORDENES_PRODUCCION              | Función automática   |
| OMs_26___mi_eelo.xlsx (001-070) | ORDENES_PRODUCCION (cont.)      | Función automática   |
| _26_Comercial___mi_eelo.xlsx (Sales) | COMERCIAL_VENTAS           | Función automática   |
| _26_Comercial___mi_eelo.xlsx (Compras) | COMERCIAL_COMPRAS        | Función automática   |
