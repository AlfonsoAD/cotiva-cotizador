# GH-9 · Fase 6 · Pruebas de integración de la orden de servicio

## Brief

**Objetivo:** comprobar de punta a punta que la orden de servicio (#4 a #8) funciona junta, automatizando en `test.ps1` lo que corre en Excel oculto y dejando constancia de lo que solo se ve a mano.

**Entra**
- Módulo nuevo `src/tests/modTestIntegration.bas` con `Test_Integracion_*` que encadenan pasos que hoy se prueban sueltos (P1 a P8, `QuoteFromOrder`, `CreateQuoteForOrder`): orden nueva → guardar → cargar; cliente que regresa; orden → cotización → Plantilla con vehículo; cotización vieja sin orden; ajuste a una hoja; diagnóstico largo en la PlantillaOrden; hojas protegidas tras los flujos.
- Verificación manual de lo que no corre oculto (`frmOrder`, los dos PDF, protección vista en Excel, `build.ps1` con y sin logo en `assets/`), con resultado anotado en `### Verificación manual` de Tasks.
- Corrección de defectos chicos que salgan (pocas líneas, dentro de lo que tocaron #4 a #8), cada uno con su prueba.
- Tres arreglos que ya salieron de una revisión manual previa: pie de los PDF (`TPL_PAGE` = `"Página &P"`; Excel en español guarda `&N` como `&F` y sale "Página 1 de Cotiva-Cotizador.xlsm"), crédito de innobytes cortado en `frmOrder` y `APP_VERSION` de 8.0.0 a 9.0.0 (`BD_Ordenes` y migración cambian la estructura del libro).

**No entra**
- Entrega al cliente (migración, contraseña, firma, OneDrive): su propia issue.
- Manejar formularios desde las pruebas; meter `build.ps1` dentro de `test.ps1`.

**Decisiones nuevas**
- [N] Módulo de integración propio. Descartado: repartir los casos en `modTestData`/`modTestOrderQuote`, porque se pierde la lectura del flujo completo.
- [N] "Una sola hoja" se prueba por configuración: tras `FillTemplate`/`FillOrderTemplate` el `PageSetup` sigue con `Zoom = False`, `FitToPagesWide = 1`, `FitToPagesTall = 1` y `PrintArea` no vacío. Descartado: contar páginas del PDF (`ExportSheetPdf` abre el archivo y puede mostrar `MsgBox`) o `PageSetup.Pages.Count` (depende de la impresora instalada).
- [N] Las pruebas no llaman `ExportPdf`/`ExportOrderPdf`; el PDF real se revisa a mano.
- [N] Pie solo con "Página &P": los dos PDF caben en una hoja, así que el total sobra. Descartado: buscar el código de total que acepte cada idioma de Excel.
- [N] Crédito de `frmOrder` a 380 pt como en `frmQuote`, arriba de la barra inferior; para hacerle lugar las pestañas bajan de 342 a 324 pt de alto (su contenido termina en 274). Descartado: dejarlo entre los botones (solo caben 186 pt) o partirlo en dos renglones.
- [N] Un defecto que no cabe en pocas líneas no se arregla aquí: se abre issue nueva y se anota en la verificación manual. Descartado: crecer esta tarea.

**Riesgos**
- El diagnóstico admite 1000 caracteres (`frmOrder`) y en la PlantillaOrden vive en `A23:H23` con alto fijo de 96 pt: un texto largo probablemente se corta en el PDF. Si pasa, aplica la regla anterior.
- Probar el build sin logo obliga a mover los archivos de `assets/`; hay que regresarlos antes del commit.

**Tamaño:** ≈2 h, carril light; código de pruebas más los tres arreglos y los defectos chicos que salgan.

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
  - C5: CUANDO se llena cualquiera de las dos plantillas, el sistema DEBE conservar el ajuste a una página.
  - C6: CUANDO el diagnóstico mide hasta 1000 caracteres, el sistema DEBE escribirlo completo en la PlantillaOrden.
- [x] **T6 · Hojas protegidas.** `Test_Integracion_HojasProtegidasTrasFlujo`: después de guardar una orden, pasarla a cotización y llenar las dos plantillas, `Plantilla`, `PlantillaOrden` y las `BD_*` existentes tienen `ProtectContents = True`; tras `ProtectAll` el libro tiene `ProtectStructure = True`.
  - C7: CUANDO termina cualquier flujo de la orden, el sistema DEBE dejar protegidas las plantillas y las hojas de datos.
- [x] **T7 · Arreglos de la revisión previa.**
  - Pie: en `src/modules/modStrings.bas`, `TPL_PAGE = "Página &P"`. Prueba `Test_Integracion_PiePaginaSinTotal`: en `Plantilla` y `PlantillaOrden`, `PageSetup.LeftFooter` termina en `TPL_PAGE` y no contiene `&N` ni `&F`.
  - Crédito: en `src/forms/frmOrder.vba`, `mpgMain` pasa de alto 342 a 324 y `lblInnobytes` va en `FORM_MARGIN, 430, 380, 12` (termina antes de la línea de la barra en 450 y no toca botones ni pestañas). Sin prueba automática (formulario): se revisa en la verificación manual.
  - Versión: en `src/modules/modConfig.bas`, `APP_VERSION = "9.0.0"`. Prueba `Test_Integracion_VersionApp`: `AssertEqual "9.0.0", APP_VERSION`. Si algún texto o prueba compara contra "8.0.0" como versión actual, se ajusta; el comentario de `modTestHelpers` que llama v8.0.0 al libro anterior se queda.
  - C10: CUANDO se exporta cualquiera de los dos PDF, el pie DEBE mostrar solo "Página N", sin el nombre del archivo.
  - C11: CUANDO se abre `frmOrder`, el crédito de innobytes DEBE verse completo sin encimarse con otros controles.
  - C12: CUANDO se construye el libro con la orden de servicio, el sistema DEBE reportar la versión 9.0.0 (hoja Inicio).
- [ ] **T8 · Verificación manual y cierre.** Corre `build\test.ps1` (todo en verde) y la lista de abajo; marca cada punto y anota en una línea lo que viste. Un defecto chico se corrige con su prueba en el módulo que corresponda; uno grande se abre como issue nueva y se anota aquí con su número.
  - C8: CUANDO se corre `build.ps1` con los logos en `assets/` y sin ellos, el sistema DEBE compilar sin diálogos de VBA y generar `dist/Cotiva-Cotizador.xlsm`.
  - C9: CUANDO se exportan el PDF de una orden con diagnóstico largo y el de su cotización, cada uno DEBE salir en una sola hoja carta y legible.

### Verificación manual

Ya revisado en una copia del `.xlsm` (revisión previa, antes de T7):
- [x] PDF de la orden completa y de una casi vacía: una hoja; solo falló el pie (arreglo de T7).
- [x] PDF de la cotización con orden ligada y sin ella: una hoja; mismo pie.
- [x] Cotización de 17 conceptos sin IVA: una hoja.
- [x] Formularios sin controles encimados; crédito de innobytes cortado en `frmOrder` (arreglo de T7).
- [x] Hojas Inicio y Mi empresa.

Pendiente:
- [ ] Tras T7: pie "Página 1" en los dos PDF, crédito completo en `frmOrder` y versión 9.0.0 en Inicio.
- [ ] Clics y escritura en los formularios: validaciones, Guardar/Cargar y avisos.
- [ ] `build.ps1` con los logos de `assets/`: compila, sin diálogos.
- [ ] `build.ps1` sin logos (movidos fuera de `assets/` y regresados después): compila, sin diálogos; plantillas sin logo.
- [ ] En `frmOrder`, al escribir un cliente que ya tiene orden se llenan solos teléfonos, correo y dirección.
- [ ] Orden nueva desde el formulario: guardar, cerrar y volver a cargarla por folio.
- [ ] "Orden a cotización" abre la cotización con la franja de la orden; "ir a la orden" regresa.
- [ ] Cotización vieja (sin orden) abre sin franja y su PDF no muestra renglón de vehículo.
- Defecto corregido: un diagnóstico de 939 caracteres en 12 renglones solo mostraba 6.5 renglones (alto fijo de 96 pt); `A23` sube a 200 pt, con la prueba `Test_Integracion_DiagnosticoLargoCabeEnAlto`. Falta repetir el PDF con la corrección.
- [ ] PDF de la orden con diagnóstico de 1000 caracteres: una hoja, texto completo visible.
- [ ] PDF de la cotización ligada: una hoja, con vehículo.
- [ ] En Excel, la protección impide escribir a mano en `Plantilla` y `PlantillaOrden` fuera de las celdas editables; las `BD_*` no se ven.
