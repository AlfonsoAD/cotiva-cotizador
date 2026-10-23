# GH-8 · Fase 5 · PDF de la orden y vehículo en el PDF de la cotización

## Brief

**Objetivo:** que la orden de servicio salga en PDF (carta, una página) y que el PDF de la cotización ligada a una orden muestre el vehículo en el renglón 7.

**Entra**
- Hoja `PlantillaOrden` (visible y protegida, como `Plantilla`) con `DesignOrderTemplate`, `FillOrderTemplate` y `ExportOrderPdf` en `modTemplate`. Contenido: encabezado con logo, cliente, vehículo, entrada y salida, servicio solicitado, diagnóstico, condiciones editables y tres firmas (Cliente, Recibió, Entregó).
- Botón «Generar PDF» en `frmOrder`, con el mismo flujo que el de `frmQuote`: ofrece guardar, revisa los datos de la empresa y exporta.
- Renglón 7 de `Plantilla`: «Vehículo» en A7 y «Lincoln MKX 2010 · Placas ABC123» en B7:D7. Si no hay orden ligada, o la orden no tiene vehículo ni placas, queda vacío y con 10 pt de alto, como hoy.
- `EnsureWorkbookStructure` crea la hoja en libros anteriores, `ProtectAll` la protege, y el README documenta la macro nueva.

**No entra**
- Pruebas automáticas nuevas (lo pidió el humano). La verificación es manual: build, PDF de una orden, PDF de una cotización con y sin orden. El cambio a `OrderBannerText` ya lo cubren las pruebas que existen.
- Leyenda «Generado con Cotiva» y créditos de innobytes en el PDF de la orden. El pie solo lleva el número de página.
- Vista previa de la orden y cambios en los campos de `ServiceOrder` o de `BD_Ordenes`.

**Decisiones nuevas**
- [N] La orden se arma en `modTemplate`, que comparte con la cotización `PdfFolder`, `ConfigurePage` y `BottomBorder`. Descartado: un módulo aparte, que obligaría a hacer públicos esos helpers. La skill también dice que las plantillas viven en `modTemplate`.
- [N] Cuadrícula de 8 columnas (A-H) con tres pares etiqueta/valor por renglón y columnas C y F angostas como separación. Así las tres firmas quedan separadas. Descartado: las 4 columnas de la cotización, porque ahí tres firmas no caben parejas.
- [N] El logo de la orden se copia del logo de `Plantilla` cada vez que se llena la orden, así «Mi logo» sigue siendo la única fuente. Descartado: que `PlaceLogo` lo inserte en las dos hojas, porque los libros que ya tienen logo se quedarían sin él en la orden.
- [N] El vehículo va como etiqueta en A7 y valor en B7:D7, igual que «Cliente» en el renglón 6, en lugar de «Vehículo: …» en una sola celda. El alto pasa a 18 pt solo cuando hay texto. Las filas 8 a 34 no se mueven.
- [N] El PDF de la orden se llama `[Folio]_[Cliente].pdf` (por ejemplo `OS-0001_Juan Pérez.pdf`) y se guarda en la misma carpeta `Cotizaciones_PDF`. Descartado: una carpeta nueva solo para órdenes.
- [N] `ConfirmCompanyData` pasa de privada en `frmQuote` a pública en `modTemplate` para que la usen los dos formularios, sin cambiar lo que hace.

**Riesgos**
- Copiar el logo usa el portapapeles (`Shape.Copy` y `Worksheet.Paste`). Si falla, el PDF sale sin logo y sin error.
- Servicio solicitado y diagnóstico van en celdas combinadas de alto fijo (unos 5 y 8 renglones). Un texto más largo se corta en el PDF.
- `FitToPagesTall = 1` encoge la hoja si algo la desborda.

**Tamaño:** unas 2 h. Toca `modTemplate`, `modUtils`, `modConfig`, `modStrings`, `modApp`, `modSecurity`, `frmOrder`, `frmQuote` y el README.

## Discovery

<!-- Discovery completo; lo escribe bflow al cerrar el gate discovery. No lo edites. -->

## Requirements

Carril light: los criterios van dentro de cada tarea en Tasks.

## Design

Carril light: los nombres y el mapa de la hoja van en Tasks.

## Tasks

- [ ] **T1 · Vehículo en el PDF de la cotización**
  - `modUtils`: nueva `Public Function VehicleText(ByVal o As ServiceOrder) As String`. Regresa «Marca Modelo Año · Placas X», omite las partes vacías (año 0 = sin año, placas en mayúsculas) y regresa "" si `o Is Nothing` o si todo está vacío. `OrderBannerText` la usa en vez de su ciclo, sin cambiar su salida: las pruebas `Test_OrderBannerText_*` deben seguir pasando.
  - `modConfig`: `CELL_VEHICLE_LABEL = "A7"`, `CELL_VEHICLE = "B7"`, `VEHICLE_ROW_HEIGHT = 18`. `modStrings`: `TPL_VEHICLE = "Vehículo"`.
  - `DesignTemplate`: el renglón 7 sigue con 10 pt; A7 con el mismo formato que A6 (9 pt, gris, centrado vertical) y B7:D7 combinada (10 pt, ShrinkToFit, centrado vertical).
  - `ClearTemplateContent` limpia A7 y B7 y deja el renglón 7 en 10 pt. `FillTemplate`: si `q.OrderNumber` no está vacío, `LoadOrder` y `VehicleText`. Con texto, escribe `TPL_VEHICLE` en A7, el texto en B7 y pone el renglón en `VEHICLE_ROW_HEIGHT`.
  - Criterios: CUANDO la cotización tiene una orden ligada con vehículo, el PDF DEBE mostrar «Vehículo» y «Lincoln MKX 2010 · Placas ABC123» en el renglón 7, sin mover las filas 8 a 34. CUANDO no tiene orden, la orden no existe o no tiene vehículo ni placas, el renglón 7 DEBE quedar vacío y con 10 pt. CUANDO se carga otra cotización o se empieza una nueva, el vehículo anterior NO DEBE quedarse en la hoja.

- [ ] **T2 · Hoja `PlantillaOrden`: diseño**
  - `modConfig`: `SHEET_ORDER_TEMPLATE = "PlantillaOrden"`, `ORDER_PRINT_AREA = "$A$1:$H$29"` y una `CELL_ORD_*` por cada celda de datos del mapa. `modStrings`: `TPL_ORD_*` para el título «ORDEN DE SERVICIO», las secciones, las etiquetas, los pies de firma y `TPL_ORD_TERMS_DEFAULT`. El texto por defecto de las condiciones es genérico para cualquier taller y no nombra a ninguno: el taller no responde por objetos dejados en el vehículo, los trabajos extra se hacen solo con autorización del cliente, y los tiempos y montos son estimados. `MSG_ORDER_DESIGN_CONFIRM` y `MSG_ORDER_DESIGN_DONE`, si hacen falta textos distintos de los de la cotización.
  - Columnas: A, D y G (etiquetas) de 11; B, E y H (valores) de 19; C y F de 2.
  - Mapa:
    - 1: logo en A1:D1, título en E1:H1.
    - 2-4: empresa en A2:D2, A3:D3 y A4:D4. G2/H2 Folio, G3/H3 Estado, G4/H4 Prometida.
    - 5: línea de color.
    - 6: sección CLIENTE.
    - 7: Nombre en B7:E7, Teléfono en H7.
    - 8: Correo en B8:E8, Teléfono 2 en H8.
    - 9: Dirección en B9:H9.
    - 11: sección VEHÍCULO.
    - 12: Marca, Modelo, Año.
    - 13: Placas, Color, Motor.
    - 14: VIN, Kilometraje, Gasolina.
    - 16: sección ENTRADA Y SALIDA.
    - 17: Entrada (fecha y hora), Recibió, Prometida.
    - 18: Salida (fecha y hora), Entregó.
    - 20-21: SERVICIO SOLICITADO y A21:H21 (unos 48 pt).
    - 22-23: DIAGNÓSTICO y A23:H23 (unos 96 pt).
    - 25-26: CONDICIONES y A26:H26 con `Locked = False` (unos 48 pt).
    - 27: espacio para firmar (unos 40 pt).
    - 28: nombres en A28:B28, D28:E28 y G28:H28, con borde superior.
    - 29: pies Cliente, Recibió y Entregó.
    - Valores con ShrinkToFit. Los textos largos en celdas combinadas con WrapText, alineadas arriba.
  - `Public Sub DesignOrderTemplate(Optional ByVal silent As Boolean = False)`, con el mismo patrón que `DesignTemplate`: confirmación si no es `silent`, lienzo limpio, fuente y colores de la marca, colores de la empresa (título y línea con el principal), `ConfigurePage` con el área de impresión como parámetro (la cotización le pasa `PRINT_AREA`), sin cuadrícula y la hoja protegida al final.
  - `EnsureWorkbookStructure` crea la hoja y llama `DesignOrderTemplate True` si no existe. `ProtectAll` agrega `SHEET_ORDER_TEMPLATE` a las hojas visibles protegidas.
  - Criterios: CUANDO se arma el build o se abre un libro de una versión anterior, DEBE existir `PlantillaOrden` protegida, en carta y a una página. Solo A26:H26 DEBE poder editarse con la hoja protegida. CUANDO se rediseña con `DesignOrderTemplate`, el texto de las condiciones DEBE volver al de `TPL_ORD_TERMS_DEFAULT`.

- [ ] **T3 · Llenado, logo y PDF de la orden**
  - `Public Sub FillOrderTemplate(ByVal o As ServiceOrder)`: limpia solo las celdas de datos, nunca las condiciones. Escribe empresa, folio, estado, cliente, vehículo, entrada y salida, servicio, diagnóstico y los nombres de las firmas, todo con `SafeCellText`. Las fechas 0 quedan en blanco. Formatos: fecha `dd mmm yyyy`, hora `hh:mm`, kilometraje `#,##0 "km"` (0 = vacío), año vacío si es 0. Luego aplica los colores de la empresa y llama `SyncOrderLogo`. Maneja errores igual que `FillTemplate`.
  - `Private Sub SyncOrderLogo(ByVal ws As Worksheet)`: borra el `COMPANY_LOGO_NAME` de `PlantillaOrden`. Si `Plantilla` tiene logo, lo copia, lo nombra `COMPANY_LOGO_NAME` y lo acomoda en A1:D1 conservando la proporción, como `FitLogos`. Restaura la hoja activa. Si el pegado falla, sigue sin logo y sin mostrar error.
  - `Public Function ExportOrderPdf(ByVal o As ServiceOrder) As String`: hace lo mismo que `ExportPdf` sobre `PlantillaOrden`, con `PdfFolder()` y `SanitizeFileName(o.Number & "_" & o.ClientName) & ".pdf"`. Pregunta si el archivo ya existe y regresa la ruta, o "" si no se generó.
  - Criterios: CUANDO se exporta una orden, el sistema DEBE generar `[Folio]_[Cliente].pdf` de una página carta con los datos de la orden y el logo de la empresa, si hay. CUANDO un campo está vacío, su celda DEBE quedar en blanco, sin «0», «00:00» ni «0 km». Las condiciones editadas a mano NO DEBEN borrarse al llenar otra orden.

- [ ] **T4 · Botón «Generar PDF» en `frmOrder`**
  - Mover `ConfirmCompanyData` de `frmQuote` a `modTemplate` como `Public Function ConfirmCompanyData() As Boolean`, sin cambiar lo que hace, y hacer que `frmQuote` llame a esa.
  - `frmOrder`: `btnPdf` (`UI_BTN_PDF`, `kindDark`) en la barra inferior, a la derecha. «Guardar» se recorre a su izquierda y el crédito de innobytes se angosta para que nada se encime. El orden del tabulador sigue siendo coherente.
  - `btnPdf_Click`: `SyncOrderFromForm`. Si hay cambios o la orden no se ha guardado, pregunta `MSG_SAVE_ORDER_BEFORE_PDF` (nuevo, «La orden tiene cambios sin guardar…»; Sí guarda, No sigue, Cancelar sale). Luego `ConfirmCompanyData`, `FillOrderTemplate mOrder`, `ExportOrderPdf mOrder` y `MSG_PDF_CREATED`. Errores con `ShowError "btnPdf_Click"`.
  - README: `modTemplate` en el mapa de módulos y `DesignOrderTemplate` en la tabla de macros de administrador. Agregar a la verificación manual el PDF de la orden y el renglón del vehículo.
  - Criterios: CUANDO el usuario pulsa «Generar PDF» en la orden, el sistema DEBE ofrecer guardar los cambios pendientes, avisar si faltan datos de la empresa y abrir el PDF generado. CUANDO la orden no pasa la validación, NO DEBE generarse ningún PDF. El PDF de la cotización DEBE seguir funcionando igual.
  - Verificación manual: build sin diálogos de VBA; `test.ps1` en verde; PDF de una orden completa y de una casi vacía; PDF de una cotización con orden (renglón 7 lleno) y sin orden (igual que hoy); un libro de la versión anterior abre y crea `PlantillaOrden`.
