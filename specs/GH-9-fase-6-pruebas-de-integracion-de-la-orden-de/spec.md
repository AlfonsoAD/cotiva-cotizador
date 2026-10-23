# GH-9 · Fase 6 · Pruebas de integración de la orden de servicio

## Brief

**Objetivo:** comprobar de punta a punta que la orden de servicio (#4 a #8) funciona junta, automatizando en `test.ps1` lo que corre en Excel oculto y dejando constancia de lo que solo se ve a mano.

**Entra**
- Módulo nuevo `src/tests/modTestIntegration.bas` con `Test_Integracion_*` que encadenan pasos que hoy se prueban sueltos (P1 a P8, `QuoteFromOrder`, `CreateQuoteForOrder`): orden nueva → guardar → cargar; cliente que regresa; orden → cotización → Plantilla con vehículo; cotización vieja sin orden; ajuste a una hoja; diagnóstico largo en la PlantillaOrden; hojas protegidas tras los flujos.
- Verificación manual de lo que no corre oculto (`frmOrder`, los dos PDF, protección vista en Excel, `build.ps1` con y sin logo en `assets/`), con resultado anotado en `### Verificación manual` de Tasks.
- Corrección de defectos chicos que salgan (pocas líneas, dentro de lo que tocaron #4 a #8), cada uno con su prueba.
- Tres arreglos que ya salieron de una revisión manual previa: pie de los PDF (`TPL_PAGE` = `"Página &P"`; Excel en español guarda `&N` como `&F` y sale "Página 1 de Cotiva-Cotizador.xlsm"), crédito de innobytes cortado en `frmOrder` y `APP_VERSION` de 8.0.0 a 9.0.0 (`BD_Ordenes` y migración cambian la estructura del libro).
- Diagnóstico que crece (T9): en la PlantillaOrden el diagnóstico ocupa de 1 a 8 filas reservadas (5 renglones del texto por fila, alto medido con `AutoFit`) y la orden pasa a una 2ª página si no cabe; `frmOrder` limita el diagnóstico a 40 renglones. La cotización sigue en una hoja.

**No entra**
- Entrega al cliente (migración, contraseña, firma, OneDrive): su propia issue.
- Manejar formularios desde las pruebas; meter `build.ps1` dentro de `test.ps1`.

**Decisiones nuevas**
- [N] Módulo de integración propio. Descartado: repartir los casos en `modTestData`/`modTestOrderQuote`, porque se pierde la lectura del flujo completo.
- [N] El ajuste de página se prueba por configuración: tras `FillTemplate` la Plantilla sigue con `Zoom = False`, `FitToPagesWide = 1`, `FitToPagesTall = 1`; tras `FillOrderTemplate` la PlantillaOrden queda con `FitToPagesWide = 1` y `FitToPagesTall = False` (las páginas que pida el alto); las dos con `PrintArea` no vacío. Descartado: contar páginas del PDF (`ExportSheetPdf` abre el archivo y puede mostrar `MsgBox`) o `PageSetup.Pages.Count` (depende de la impresora instalada).
- [N] Diagnóstico en filas reservadas que se ocultan si sobran, como los conceptos de la Plantilla, y tope de 40 renglones en `frmOrder` (con 1000 caracteres, 40 renglones caben en dos hojas). Descartado: una sola celda con alto calculado (Excel no pasa de 409 pt por fila ni corta una fila entre páginas) e insertar filas al llenar (las celdas de condiciones y firmas cambiarían de lugar en cada orden).
- [N] Las pruebas no llaman `ExportPdf`/`ExportOrderPdf`; el PDF real se revisa a mano.
- [N] Pie solo con "Página &P": los dos PDF caben en una hoja, así que el total sobra. Descartado: buscar el código de total que acepte cada idioma de Excel.
- [N] Crédito de `frmOrder` a 380 pt como en `frmQuote`, arriba de la barra inferior; para hacerle lugar las pestañas bajan de 342 a 324 pt de alto (su contenido termina en 274). Descartado: dejarlo entre los botones (solo caben 186 pt) o partirlo en dos renglones.
- [N] Un defecto que no cabe en pocas líneas no se arregla aquí: se abre issue nueva y se anota en la verificación manual. Descartado: crecer esta tarea.

**Riesgos**
- `AutoFit` mide en pantalla y el PDF puede salir un poco más alto: la columna de medición queda al 95 % del ancho del bloque para que sobre espacio. Las condiciones y firmas bajan de la fila 26-29 a la 33-36: un libro armado antes de T9 se reconstruye con `build.ps1` (la orden aún no se entrega a clientes).
- Probar el build sin logo obliga a mover los archivos de `assets/`; hay que regresarlos antes del commit.

**Tamaño:** ≈3 h, carril light; código de pruebas, los tres arreglos, el diagnóstico que crece (≈1 h) y los defectos chicos que salgan.

## Discovery

<!-- Discovery completo; lo escribe bflow al cerrar el gate discovery. No lo edites. -->

## Requirements

<!-- Criterios EARS numerados R1..Rn, cada uno [D] (del discovery) o [N] (nuevo). -->

## Design

<!-- Estructuras, firmas, decisiones tomadas y descartadas, superficie de seguridad. -->

## Tasks

Todas las pruebas van en `src/tests/modTestIntegration.bas` (encabezado con el propósito del módulo, `Option Explicit`, datos de `modTestHelpers`). El runner deja las `BD_*` como recién abiertas antes de cada prueba. Toda prueba que llame `FillTemplate`/`FillOrderTemplate` revisa además `AssertEqual 0, LastErrorNumber()` (en modo prueba `ShowError` registra en vez de mostrar).

- [x] **T1 · Orden nueva, guardar y cargar.** `Test_Integracion_OrdenNuevaGuardarCargar`: toma `GetNextOrderNumber`, guarda `NewSampleOrder` con ese folio, la carga con `LoadOrder` y compara folio, cliente, vehículo y diagnóstico; `OrderExists` es True para ese folio y `GetNextOrderNumber` ya regresa otro.
  - C1: CUANDO se guarda una orden con el folio sugerido, el sistema DEBE cargarla con los mismos datos y sugerir un folio nuevo para la siguiente.
- [x] **T2 · Cliente que regresa.** `Test_Integracion_ClienteQueRegresa`: guarda una orden de "Cliente Prueba" con teléfonos, correo y dirección; `GetClientContact("Cliente Prueba", ...)` regresa True y esos cuatro datos.
  - C2: CUANDO un cliente ya tiene una orden guardada, el sistema DEBE regresar su contacto para llenar la orden siguiente.
- [x] **T3 · Pasar a cotización.** `Test_Integracion_OrdenACotizacionYPlantilla`: guarda una orden sin cotización, llama `CreateQuoteForOrder`, carga cotización y orden (ligadas en ambos sentidos), llena la Plantilla con `FillTemplate` y revisa que `CELL_VEHICLE` traiga el vehículo y que `OrderBannerText` no esté vacío.
  - C3: CUANDO una orden pasa a cotización, el sistema DEBE ligar ambos folios y mostrar el vehículo en la Plantilla de la cotización.
- [x] **T4 · Cotización vieja sin orden.** `Test_Integracion_CotizacionViejaSinOrden`: `WriteLegacyQuoteRow`, `LoadQuote`, `FillTemplate`; `OrderNumber` vacío, `CELL_VEHICLE` vacía, `OrderBannerText` vacío y sin error registrado.
  - C4: CUANDO se carga una cotización de una versión anterior, el sistema DEBE llenarla sin vehículo y sin error.
- [x] **T5 · Una sola hoja y diagnóstico largo.** `Test_Integracion_PlantillasUnaHoja`: tras `FillTemplate` (cotización con el tope de conceptos) y `FillOrderTemplate` (orden completa), las dos hojas conservan `Zoom = False`, `FitToPagesWide = 1`, `FitToPagesTall = 1` y `PrintArea` no vacío. `Test_Integracion_DiagnosticoLargoEnPlantilla`: orden con diagnóstico de 1000 caracteres con saltos de línea; `FillOrderTemplate` deja el texto completo en `CELL_ORD_DIAGNOSIS`.
  - C5 (ajustado en T9): CUANDO se llena la Plantilla, el sistema DEBE conservar el ajuste a una página; CUANDO se llena la PlantillaOrden, DEBE ajustarla a una página de ancho y dejar que el alto ocupe las páginas que necesite.
  - C6 (ajustado en T9): CUANDO el diagnóstico mide hasta 1000 caracteres, el sistema DEBE escribirlo completo en las filas del diagnóstico de la PlantillaOrden.
- [x] **T6 · Hojas protegidas.** `Test_Integracion_HojasProtegidasTrasFlujo`: después de guardar una orden, pasarla a cotización y llenar las dos plantillas, `Plantilla`, `PlantillaOrden` y las `BD_*` existentes tienen `ProtectContents = True`; tras `ProtectAll` el libro tiene `ProtectStructure = True`.
  - C7: CUANDO termina cualquier flujo de la orden, el sistema DEBE dejar protegidas las plantillas y las hojas de datos.
- [x] **T7 · Arreglos de la revisión previa.**
  - Pie: en `src/modules/modStrings.bas`, `TPL_PAGE = "Página &P"`. Prueba `Test_Integracion_PiePaginaSinTotal`: en `Plantilla` y `PlantillaOrden`, `PageSetup.LeftFooter` termina en `TPL_PAGE` y no contiene `&N` ni `&F`.
  - Crédito: en `src/forms/frmOrder.vba`, `mpgMain` pasa de alto 342 a 324 y `lblInnobytes` va en `FORM_MARGIN, 430, 380, 12` (termina antes de la línea de la barra en 450 y no toca botones ni pestañas). Sin prueba automática (formulario): se revisa en la verificación manual.
  - Versión: en `src/modules/modConfig.bas`, `APP_VERSION = "9.0.0"`. Prueba `Test_Integracion_VersionApp`: `AssertEqual "9.0.0", APP_VERSION`. Si algún texto o prueba compara contra "8.0.0" como versión actual, se ajusta; el comentario de `modTestHelpers` que llama v8.0.0 al libro anterior se queda.
  - C10: CUANDO se exporta cualquiera de los dos PDF, el pie DEBE mostrar solo "Página N", sin el nombre del archivo.
  - C11: CUANDO se abre `frmOrder`, el crédito de innobytes DEBE verse completo sin encimarse con otros controles.
  - C12: CUANDO se construye el libro con la orden de servicio, el sistema DEBE reportar la versión 9.0.0 (hoja Inicio).
- [x] **T8 · Verificación manual y cierre.** Corre `build\test.ps1` (todo en verde) y la lista de abajo; marca cada punto y anota en una línea lo que viste. Un defecto chico se corrige con su prueba en el módulo que corresponda; uno grande se abre como issue nueva y se anota aquí con su número.
  - C8: CUANDO se corre `build.ps1` con los logos en `assets/` y sin ellos, el sistema DEBE compilar sin diálogos de VBA y generar `dist/Cotiva-Cotizador.xlsm`.
  - C9 (ajustado en T9): CUANDO se exporta el PDF de una cotización, DEBE salir en una sola hoja carta y legible. CUANDO se exporta el PDF de una orden, DEBE salir en una hoja si el diagnóstico cabe y, si no, en dos, con el diagnóstico completo, legible y numeradas "Página 1" y "Página 2".
- [x] **T9 · Diagnóstico que crece y orden en dos hojas.** Sustituye el alto fijo de 200 pt de `A23` (arreglo anterior): con muchos renglones cortos el texto pasa de 13 renglones y se corta sin aviso.
  - `src/modules/modConfig.bas`: `ORD_DIAG_FIRST_ROW As Long = 23`, `ORD_DIAG_ROWS As Long = 8`, `ORD_DIAG_LINES_PER_ROW As Long = 5`, `ORD_DIAG_MAX_LINES As Long = ORD_DIAG_ROWS * ORD_DIAG_LINES_PER_ROW` (= 40), `ORD_DIAG_MIN_HEIGHT As Double = 96`, `ORD_MEASURE_COLUMN As String = "J"`. `CELL_ORD_DIAGNOSIS` sigue en `"A23"` (primera fila). Lo de abajo baja 7 filas: `CELL_ORD_TERMS = "A33"`, firmas `"A35"`/`"D35"`/`"G35"`, `ORDER_PRINT_AREA = "$A$1:$H$36"`.
  - `src/modules/modUtils.bas`, junto a `DiagnosisLines`:
    - `Public Function TextLineCount(ByVal text As String) As Long`: une `vbCrLf`/`vbCr` en `vbLf`, quita los saltos del final y cuenta renglones, vacíos de en medio incluidos; `""` regresa 0.
    - `Public Function DiagnosisChunks(ByVal diagnosisText As String) As Collection`: el mismo texto normalizado en grupos de `ORD_DIAG_LINES_PER_ROW` renglones unidos con `vbLf`; el grupo `ORD_DIAG_ROWS` se lleva todo lo que sobre (datos que no pasaron por `frmOrder`); texto vacío regresa una colección vacía. Unir los grupos con `vbLf` da el texto normalizado.
  - `src/modules/modTemplate.bas`:
    - `ConfigurePage` recibe `ByVal oneTallPage As Boolean`: True deja `FitToPagesTall = 1` (lo pasa `DesignTemplate`); False pone `FitToPagesTall = False` (lo pasa `DesignOrderTemplate`). `FitToPagesWide = 1` en los dos.
    - `DesignOrderTemplate`: filas `ORD_DIAG_FIRST_ROW` a `ORD_DIAG_FIRST_ROW + ORD_DIAG_ROWS - 1` (23-30) con `AddOrderTextBlock`; la 23 a `ORD_DIAG_MIN_HEIGHT` y las 24-30 ocultas. Espacio en 31, condiciones en 32-33, firmas en 34-36 (las mismas alturas y formatos de hoy, siete filas abajo). Columna `ORD_MEASURE_COLUMN` (fuera del área de impresión) con letra de 10 pt, `WrapText` y ancho del 95 % de `Range("A23:H23").Width` (convertir puntos a `ColumnWidth` con la razón `Width / ColumnWidth` de la propia columna).
    - `Private Sub FillOrderDiagnosis(ByVal ws As Worksheet, ByVal diagnosisText As String)`, que `FillOrderTemplate` llama en lugar de `PutOrderText ... CELL_ORD_DIAGNOSIS`: muestra y limpia las 8 filas; escribe cada grupo de `DiagnosisChunks` con `SafeCellText` en la columna A de su fila; para medir copia el grupo a `ORD_MEASURE_COLUMN` de esa fila, hace `Rows(r).AutoFit` y borra la copia; la fila 23 nunca queda por debajo de `ORD_DIAG_MIN_HEIGHT` y se ve aunque no haya diagnóstico; las filas sin grupo quedan ocultas.
    - `ClearOrderContent` limpia también las 8 filas del diagnóstico.
  - `src/forms/frmOrder.vba`, `SyncOrderFromForm`: si `TextLineCount(txtDiagnosis.Text) > ORD_DIAG_MAX_LINES`, `WarnOnPage` en la pestaña de `txtDiagnosis` con `Fmt(MSG_DIAGNOSIS_TOO_MANY_LINES, ORD_DIAG_MAX_LINES, <renglones>)` y no guarda. En `src/modules/modStrings.bas`: `MSG_DIAGNOSIS_TOO_MANY_LINES = "El diagnóstico admite hasta {0} renglones y tiene {1}. Junta o resume renglones."`
  - `README.md`, sección de la orden: el diagnóstico admite 1000 caracteres y 40 renglones; si no cabe en la hoja, el PDF de la orden sale en dos.
  - Pruebas en `src/tests/modTestOrderQuote.bas`: `Test_TextLineCount_CuentaRenglones` (`""`→0, `"a"`→1, `"a" & vbCrLf & "b"`→2, `"a" & vbLf & vbLf & "b"`→3, `"a" & vbLf`→1); `Test_DiagnosisChunks_CincoPorFila` (12 renglones → grupos de 5, 5 y 2; unidos dan el texto); `Test_DiagnosisChunks_UltimaFilaTomaElResto` (45 renglones → 8 grupos, el último con 10); `Test_DiagnosisChunks_Vacio` (0 grupos).
  - Pruebas en `src/tests/modTestIntegration.bas` (con un auxiliar privado que une con `vbLf` el texto de las filas 23-30 no vacías):
    - Se borra `Test_Integracion_DiagnosticoLargoCabeEnAlto` (probaba el alto fijo).
    - `Test_Integracion_DiagnosticoLargoEnPlantilla` compara el texto unido de las filas en vez de `CELL_ORD_DIAGNOSIS`.
    - `Test_Integracion_PlantillasUnaHoja` pasa a `Test_Integracion_AjusteDePagina`: la Plantilla como hoy; la PlantillaOrden con `Zoom = False`, `FitToPagesWide = 1`, `FitToPagesTall = False` y `PrintArea` igual a `ORDER_PRINT_AREA`.
    - `Test_Integracion_DiagnosticoRenglonesCortosCrece`: 40 renglones cortos; filas 23-30 visibles, cada una con alto ≥ 5 × 12 pt, texto unido igual al original y sin error registrado.
    - `Test_Integracion_DiagnosticoCortoOcultaSobrantes`: llena con 40 renglones y luego con 2; filas 24-30 ocultas y vacías, `A23` con los 2 renglones y alto `ORD_DIAG_MIN_HEIGHT`; con diagnóstico vacío, `A23` vacía y visible.
    - `Test_Integracion_CondicionesYFirmasEnSuLugar`: `CELL_ORD_TERMS` desbloqueada y con texto; debajo de `CELL_ORD_SIGN_CLIENT` dice `TPL_CLIENT`; la columna `ORD_MEASURE_COLUMN` queda vacía tras `FillOrderTemplate`.
  - C13: CUANDO el diagnóstico no cabe en la primera hoja, el sistema DEBE pasar el resto de la orden a una segunda hoja sin cortar texto.
  - C14: CUANDO el diagnóstico tiene pocos renglones o ninguno, el sistema DEBE mostrar un solo bloque de al menos 96 pt y ocultar las filas que sobran.
  - C15: CUANDO se intenta guardar en `frmOrder` un diagnóstico de más de 40 renglones, el sistema DEBE avisar en su pestaña y no guardar.
  - C16: CUANDO se llena la PlantillaOrden, la cotización DEBE seguir saliendo en una hoja (C5).

### Verificación manual

Ya revisado en una copia del `.xlsm` (revisión previa, antes de T7):
- [x] PDF de la orden completa y de una casi vacía: una hoja; solo falló el pie (arreglo de T7).
- [x] PDF de la cotización con orden ligada y sin ella: una hoja; mismo pie.
- [x] Cotización de 17 conceptos sin IVA: una hoja.
- [x] Formularios sin controles encimados; crédito de innobytes cortado en `frmOrder` (arreglo de T7).
- [x] Hojas Inicio y Mi empresa.

Verificado tras T7 (sesión principal y humano):
- [x] Tras T7: pie "Página 1" en los dos PDF, crédito completo en `frmOrder` y versión 9.0.0 en Inicio.
- [x] Clics y escritura en los formularios: validaciones, Guardar/Cargar y avisos.
- [x] `build.ps1` con los logos de `assets/`: compila, sin diálogos.
- [x] `build.ps1` sin logos (movidos fuera de `assets/` y regresados después): compila, sin diálogos; plantillas sin logo.
- [x] En `frmOrder`, al escribir un cliente que ya tiene orden se llenan solos teléfonos, correo y dirección.
- [x] Orden nueva desde el formulario: guardar, cerrar y volver a cargarla por folio.
- [x] "Orden a cotización" abre la cotización con la franja de la orden; "ir a la orden" regresa.
- [x] Cotización vieja (sin orden) abre sin franja y su PDF no muestra renglón de vehículo.
- Defecto corregido: un diagnóstico de 939 caracteres en 12 renglones solo mostraba 6.5 renglones (alto fijo de 96 pt); `A23` sube a 200 pt, con la prueba `Test_Integracion_DiagnosticoLargoCabeEnAlto`. PDF repetido con la corrección.
- [x] PDF de la orden con diagnóstico de 1000 caracteres: una hoja, texto completo visible.
- [x] (PDF generados desde dist/ por la sesión principal) Orden con 40 renglones cortos: 2 hojas completa, "Página 1" y "Página 2"; con 1000 caracteres: 1 hoja; sin diagnóstico: 1 hoja; cotización ligada: 1 hoja con vehículo.
- Defectos corregidos: hueco tras el renglón 5 (el mínimo de 96 pt solo aplica con una fila de diagnóstico) e interlineado apretado entre filas (cada fila medida suma `ORD_DIAG_ROW_PAD` = 6 pt). PDF por repetir.
- [x] PDF de la cotización ligada: una hoja, con vehículo.
- [x] En Excel, la protección impide escribir a mano en `Plantilla` y `PlantillaOrden` fuera de las celdas editables; las `BD_*` no se ven.

Pendiente tras T9 (reemplaza la revisión del diagnóstico de 1000 caracteres de arriba):
- [x] `build.ps1` compila sin diálogos con la nueva PlantillaOrden (corrido por el implementer: genera dist/Cotiva-Cotizador.xlsm).
- [x] PDF de una orden con 40 renglones cortos: dos hojas, diagnóstico completo, "Página 1" y "Página 2", condiciones y firmas sin cortar.
- [x] PDF de una orden con diagnóstico de 1000 caracteres en pocos renglones y de una sin diagnóstico: una hoja, como antes.
- [ ] En `frmOrder`, un diagnóstico de 41 renglones avisa en su pestaña y no se guarda.
- [x] PDF de la cotización ligada: sigue en una hoja.
