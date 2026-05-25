# Diseño del Sistema Completo — Mi eelo

## La idea central

```
                    ┌─────────────────────────────────────┐
                    │   GOOGLE SHEETS  (base de datos)     │
                    │                                      │
  KoboToolbox ───► │  PARTICIPANTES + ASISTENCIA +        │
                    │  FACTURACIÓN + ÓRDENES +             │
                    │  CLIENTES + PROVEEDORES + COMERCIAL  │
                    └──────────────┬──────────────────────┘
                                   │  botón / trigger
                    ┌──────────────▼──────────────────────┐
                    │   GOOGLE DOCS  (documentos)          │
                    │                                      │
                    │  • Factura de participante           │
                    │  • Orden de Producción (OP)          │
                    │  • Orden de Manufactura (OM)         │
                    │  • Reporte mensual                   │
                    └──────────────┬──────────────────────┘
                                   │  guarda automáticamente
                    ┌──────────────▼──────────────────────┐
                    │   GOOGLE DRIVE  (archivos)           │
                    │                                      │
                    │  Organizado por: Cliente / Proveedor │
                    │  Historial completo de cada uno      │
                    │  Mockups, fotos, PDFs                │
                    └─────────────────────────────────────┘
```

---

## Las 10 hojas del Google Sheets

| Hoja | Viene de | Para qué sirve |
|------|----------|----------------|
| **PARTICIPANTES_MASTER** | Participantes.xlsx | Datos de 76 personas: ID, nombre, proyecto, división |
| **DATOS_FACTURACION** | Facturación (hoja "Datos de Facturación") | DPI, NIT, banco, datos SAT — datos sensibles separados |
| **ASISTENCIA_KOBO** | Planilla.xlsx + Kobo API | Entradas/salidas diarias, horas calculadas |
| **FACTURACION_CONSOLIDADA** | Facturación (14 hojas) | Un solo historial por participante/mes/quincena |
| **ORDENES_PRODUCCION** | OPs.xlsx + OMs.xlsx | Todas las órdenes juntas: cliente, descripción, tela, medidas, serigrafía |
| **CLIENTES** | Nueva | Lista de clientes con historial de órdenes y carpeta en Drive |
| **PROVEEDORES** | Nueva | Lista de proveedores con historial de compras |
| **COMERCIAL_VENTAS** | Comercial (Sales) | Registro de ventas vinculado a órdenes |
| **COMERCIAL_COMPRAS** | Comercial (Compras) | Registro de compras vinculado a proveedores |
| **DASHBOARD** | Calculado | Métricas en tiempo real |

---

## Diseño de ORDENES_PRODUCCION (hoja principal)

Columnas que reflejan el formulario real de OP/OM:

```
A: Numero_Orden      (OP26-001 / OM25-001)
B: Tipo              (OP | OM)
C: Fecha_Creacion
D: Cliente
E: Contacto
F: Proyecto
G: Fecha_Promesa
H: Descripcion
I: Cantidad
J: Tela_Material
K: Comentarios_Confeccion
─── Medidas ───────────────────────────────
L: Ancho
M: Altura
N: Fuelle
O: Sistema_Medicion  (Inches | CM)
─── Especificaciones ──────────────────────
P: Bolsillo_Interno  (Sí | No)
Q: Bolsillo_Externo  (Sí | No)
R: Tirantes          (Sí | No)
S: Forros            (Sí | No)
T: Specs_Adicionales
─── Serigrafía ────────────────────────────
U: Serigrafia        (Sí | No)
V: Colores_Serigrafia
W: Pantones
X: Localizacion_Impresion
Y: Medidas_Impresion
─── Films / Mockup ────────────────────────
Z: Films_Entregados
AA: Codigos_Films
AB: Mockup           (Sí | No)
AC: Muestra_Bodega   (Sí | No)
─── Control ───────────────────────────────
AD: Estado           (Pendiente | En Proceso | Completada | Cancelada)
AE: Participantes_Asignados
AF: Horas_Dedicadas
AG: URL_Mockup_Drive
AH: URL_Documento_Drive
AI: Notas_Adicionales
```

---

## Diseño de CLIENTES (hoja nueva)

```
A: ID_Cliente        (CLI-001)
B: Nombre_Cliente
C: Contacto_Principal
D: Email
E: Telefono
F: Pais
G: Total_Ordenes     (=CONTAR.SI vinculado a ORDENES)
H: Ultima_Orden      (=MAX vinculado)
I: Total_Facturado   (=SUMAR.SI)
J: URL_Carpeta_Drive (link directo a su carpeta)
K: Notas
```

---

## Flujo: generar un OP o OM con un clic

```
Usuario selecciona fila en ORDENES_PRODUCCION
         │
         ▼
🚀 Mi eelo → Generar Doc de orden
         │
         ▼
Apps Script lee todos los campos (A hasta AI) de esa fila
         │
         ├── Crea Google Doc con el formato del formulario
         │   (mismo layout que el Excel: secciones de medidas,
         │    serigrafía, films, mockup)
         │
         ├── Si hay URL de mockup en Drive → lo inserta en el Doc
         │
         ├── Mueve el Doc a la carpeta del cliente en Drive:
         │   Órdenes / Clientes / [Nombre_Cliente] / [Numero_Orden].gdoc
         │
         └── Guarda el enlace en columna AH de la hoja
```

---

## Flujo: historial de un cliente

```
Filtrar ORDENES_PRODUCCION por columna D (Cliente)
   → ver todas sus OPs y OMs históricas

Abrir carpeta Drive del cliente:
   Órdenes / Clientes / Peace by Piece /
   ├── OP26-001_Pouch_mi_eelo.pdf
   ├── OP26-015_Bolsa_Tela.pdf
   └── Mockups /
       ├── pouch_cocoCanva.jpg
       └── bolsa_lino.jpg

Desde CLIENTES, columna J → clic en el enlace → carpeta directo
```

---

## Flujo: mismo proveedor, múltiples compras

```
PROVEEDORES (hoja)
   → cada proveedor tiene su URL_Carpeta_Drive

COMERCIAL_COMPRAS
   → filtrar por proveedor → historial completo

Carpeta Drive:
   Compras / Proveedores / [Nombre_Proveedor] /
   ├── Factura_Tela_Enero_2026.pdf
   ├── Factura_Hilo_Marzo_2026.pdf
   └── DTE_Mayo_2026.pdf
```

---

## Flujo: facturación participantes

```
Kobo → ASISTENCIA_KOBO (cada hora, automático)
         │
         ▼  (empareja entrada+salida → horas reales)
FACTURACION_CONSOLIDADA
         │
         ▼  (botón o automático el día 1)
Google Doc: "Factura_ANLA060686_Mayo_2026"
         │
         ▼
Drive: Facturas / 2026 / Mayo / Factura_ANLA060686.gdoc
         │
         ▼
Email automático a participante + link al Doc
```

---

## Estructura de carpetas en Drive

```
Mi eelo - Sistema/
├── Facturas/
│   └── 2026/
│       ├── Enero/ ... Diciembre/
├── Órdenes/
│   ├── Clientes/
│   │   ├── Peace by Piece/
│   │   │   ├── Mockups/
│   │   │   └── [OP/OM docs aquí]
│   │   ├── Home Collection/
│   │   └── [un folder por cliente]
│   └── Proveedores/
│       └── [un folder por proveedor]
├── Reportes/
│   └── 2026/
└── Archivos_Historicos/
    ├── OPs_Excel_Original/
    └── OMs_Excel_Original/
```

---

## Scripts del sistema (12 módulos)

| Archivo | Función |
|---------|---------|
| `00_Config.gs` | Configuración central |
| `01_Menu.gs` | Menú + onEdit |
| `02_KoboImport.gs` | Importar y emparejar asistencia |
| `03_Participantes.gs` | Gestión de participantes |
| `04_Facturacion.gs` | Calcular horas y montos |
| `05_Documentos.gs` | Generar Docs: facturas, reportes |
| `06_Drive.gs` | Estructura y organización en Drive |
| `07_Notificaciones.gs` | Correos automáticos |
| `08_Dashboard.gs` | Métricas en tiempo real |
| `09_Triggers.gs` | Automatizaciones |
| `10_ImportarOPs.gs` | Migración inicial OPs/OMs |
| `11_ImportarComercial.gs` | Migración inicial Comercial |
| `12_OrdenesDocs.gs` | **Generar Doc de OP/OM** (nuevo) |
| `13_Clientes.gs` | **Gestión de clientes y proveedores** (nuevo) |
