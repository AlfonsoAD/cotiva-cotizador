# GH-7 · Fase 4 · Pasar la orden a cotización

## Brief

**Objetivo:** con un botón «Pasar a cotización» en `frmOrder`, la orden queda guardada y ligada a una cotización cuyos conceptos son los renglones del diagnóstico, y esa cotización se abre en `frmQuote` con una franja que muestra la orden y el vehículo.

**Entra**
- Botón «Pasar a cotización» en la franja fija de `frmOrder`, junto al estado. Guarda la orden con `SaveCurrentOrder`, que valida y confirma. Si la orden no es válida, no pasa nada más.
- Si la orden no tiene cotización ligada, `modData.CreateQuoteForOrder` crea una con el folio del sistema, el cliente, `OrderNumber` y un concepto por renglón del diagnóstico (cantidad 1, precio 0). Luego la guarda y escribe `QuoteNumber` en la orden. Si falla el segundo guardado, borra la cotización recién creada.
- Si la orden ya tiene cotización, se abre esa sin crear otra. Si su folio ya no existe, se avisa y se crea una nueva que reemplaza a `QuoteNumber`.
- `frmQuote` gana el método público `OpenQuote(folio)`, que carga como «Cargar», y una franja de solo lectura: «Orden OS-0001 · Marca Modelo 2018 · Placas ABC123». La franja también aparece al cargar desde el historial y desaparece en «Nueva cotización».
- Funciones puras con pruebas en `modUtils`: `DiagnosisLines`, `QuoteFromOrder` y `OrderBannerText`.

**No entra:** cambiar el estado de la orden, editar la orden desde `frmQuote`, desligar o volver a ligar cotizaciones a mano, llevar los precios de un catálogo y el PDF de la orden (#8).

**Decisiones nuevas**
- [N] Si el diagnóstico tiene más de `MAX_ITEMS` (17) renglones, solo los primeros 17 pasan a conceptos y se avisa cuántos quedaron fuera; el diagnóstico de la orden no cambia. Descartado: juntar el sobrante en el último concepto, porque la plantilla no tiene espacio para una descripción tan larga.
- [N] Un renglón de más de 255 caracteres se corta a 255, lo mismo que admite el campo Descripción de `frmQuote` (nueva constante `MAX_ITEM_DESCRIPTION`). Descartado: rechazar el paso.
- [N] El «precio vacío» se guarda como 0 y en la lista se ve como $0.00, porque `UnitPrice` es `Double`. Descartado: mostrar en blanco los precios en 0, porque cambiaría cómo se ven los conceptos que de verdad valen $0.
- [N] Todo el armado de la liga vive en `modData.CreateQuoteForOrder`, con su reversión, para que se pueda probar sin formularios. Descartado: hacer los dos guardados desde `frmOrder`.
- [N] `frmQuote` recibe un folio ya guardado (`OpenQuote`), no un `Quote` en memoria: los caminos de «Cargar» y de pasar la orden quedan iguales. Si `frmQuote` ya estaba abierto con cambios, pregunta lo mismo que «Cargar».
- [N] La franja es una etiqueta en el renglón del título «Datos de la cotización», alineada a la derecha. No mueve el resto del diseño. Si la orden ligada no existe, solo dice «Orden <folio>».
- [N] Las funciones puras van en una sección nueva de `modUtils`, donde ya está `NewQuoteItem`. Descartado: crear un módulo nuevo para tres funciones.

**Riesgos**
- Pasar la orden la guarda hasta tres veces y, si `SAVE_WORKBOOK_ON_SAVE` está activo, también guarda el libro tres veces. Con un libro grande puede tardar un poco.
- La reversión cuando falla `SaveOrder` no tiene prueba automática, porque no hay forma limpia de hacer que falle: se revisa leyendo el código.
- La franja, el botón y el cierre de una ventana al abrir la otra se prueban a mano.

**Tamaño:** mediano, de 2 a 3 h. Son 5 tareas, 15 pruebas automáticas nuevas y una verificación manual.

## Discovery

﻿# Discovery GH-7 · Pasar la orden a cotización

## Alcance (del issue)
- En frmOrder, una acción "Pasar a cotización" que crea un Quote con el cliente y el OrderNumber y guarda el folio de la cotización en la orden (QuoteNumber).
- Cada renglón del diagnóstico se vuelve un concepto con cantidad 1 y precio vacío.
- Abre frmQuote con la cotización cargada.
- frmQuote tiene una franja de solo lectura con el vehículo y el folio de la orden; si la cotización no tiene orden ligada, la franja no aparece.

## Decisiones
1. Orden con QuoteNumber ya puesto: se abre esa cotización existente en frmQuote, sin crear otra (respeta el vínculo 1 a 0..1 de GH-4).
2. Guardado al pasar: se valida y guarda la orden; si no es válida, no se pasa. Se crea la cotización con folio del sistema (GetNextQuoteNumber), se guarda (OrdenFolio = OrderNumber) y luego se escribe QuoteNumber en la orden y se guarda. El vínculo nunca queda a medias.
3. Diagnóstico → conceptos: se parte por salto de línea; se recortan espacios y se omiten las líneas vacías. Texto sin saltos = un concepto. Diagnóstico vacío = cotización sin conceptos. Cantidad 1, precio vacío. El diagnóstico de la orden no cambia. No se quitan viñetas ni numeración.
4. Franja en frmQuote: "Orden <folio> · <Marca> <Modelo> <Año> · Placas <placas>", omitiendo los campos vacíos. Solo visual, no cambia la validación al guardar. Se lee de la orden ligada (OrdenFolio), también al cargar con Cargar una cotización que tenga orden. Sin orden ligada, no aparece.
5. frmOrder se cierra al abrir frmQuote (una ventana a la vez, como el selector de GH-6); como la orden ya se guardó, no se avisa de cambios sin guardar.
6. QuoteNumber que apunta a un folio que no existe en BD_Cotizaciones: se avisa que no se encontró y se crea una cotización nueva con el diagnóstico, reemplazando QuoteNumber en la orden.

## Restricciones
- Genérico para cualquier taller: ningún texto ni dato propio de un cliente en el código.
- modData es el único que toca las hojas BD_*; las clases no acceden a hojas.
- La lógica de armar el Quote desde la orden (partir el diagnóstico) debe ser probable con Test_* sin formularios.
- Fuera de alcance: cambiar el estado de la orden al pasar a cotización; editar la orden desde frmQuote.

## Requirements

**Acción en `frmOrder`**
- R1 [D] CUANDO el usuario pulsa «Pasar a cotización», el sistema DEBE validar y guardar la orden con las mismas reglas y la misma confirmación de sobrescritura que «Guardar». Si la validación falla o el usuario no confirma, el sistema NO DEBE crear, abrir ni ligar ninguna cotización.
- R2 [D] CUANDO la orden guardada no tiene `QuoteNumber`, el sistema DEBE crear una cotización con folio `GetNextQuoteNumber()`, `ClientName` = cliente de la orden, `OrderNumber` = folio de la orden y fecha de hoy. Luego DEBE guardarla y después escribir su folio en `QuoteNumber` de la orden y guardar la orden.
- R3 [D] CUANDO la orden tiene un `QuoteNumber` que existe en `BD_Cotizaciones`, el sistema DEBE abrir esa cotización sin crear otra ni modificarla.
- R4 [D] CUANDO la orden tiene un `QuoteNumber` que no existe en `BD_Cotizaciones`, el sistema DEBE avisar (`MSG_LINKED_QUOTE_MISSING`) y luego hacer lo mismo que en R2, reemplazando `QuoteNumber`.
- R5 [N] SI falla el guardado de la orden en R2 o R4, el sistema DEBE borrar de `BD_Cotizaciones` y `BD_Detalle` la cotización recién creada y devolver `QuoteNumber` a su valor anterior en memoria. Ningún lado DEBE quedar apuntando a un folio que el otro no tiene.
- R6 [D] CUANDO la cotización quedó creada o encontrada, el sistema DEBE abrir `frmQuote` con ella cargada y cerrar `frmOrder` sin preguntar por cambios sin guardar.

**Diagnóstico a conceptos**
- R7 [D] El sistema DEBE partir el diagnóstico en saltos de línea (`vbCrLf`, `vbLf` o `vbCr`), recortar los espacios de cada renglón y omitir los renglones vacíos. NO DEBE quitar viñetas ni numeración.
- R8 [D] Cada renglón DEBE volverse un concepto con `Quantity` = 1, `UnitPrice` = 0 y `Description` = el renglón. Un texto sin saltos DEBE dar un solo concepto y un diagnóstico vacío, ninguno.
- R9 [N] SI hay más de `MAX_ITEMS` renglones, el sistema DEBE pasar solo los primeros `MAX_ITEMS` y, al terminar, avisar con `MSG_DIAGNOSIS_TRUNCATED` cuántos renglones hay y cuántos admite la cotización.
- R10 [N] SI un renglón mide más de `MAX_ITEM_DESCRIPTION` (255) caracteres, el sistema DEBE dejar solo los primeros 255.
- R11 [D] El sistema NO DEBE modificar `Diagnosis` de la orden.

**Franja en `frmQuote`**
- R12 [D] MIENTRAS la cotización en pantalla tenga `OrderNumber`, `frmQuote` DEBE mostrar una franja de solo lectura con el texto de `OrderBannerText`. Si no lo tiene, la franja NO DEBE verse.
- R13 [D] `OrderBannerText` DEBE devolver «Orden <folio>», seguido de « · <Marca> <Modelo> <Año>» si alguno de esos tres tiene valor (con un espacio entre los que tengan valor; año 0 = sin año) y de « · Placas <placas>» si hay placas. Sin folio devuelve "".
- R14 [N] SI la orden ligada no existe en `BD_Ordenes`, la franja DEBE decir solo «Orden <folio>».
- R15 [D] La franja DEBE actualizarse al abrir la cotización desde la orden, al cargarla con «Cargar» y al pulsar «Nueva cotización», que la oculta. NO DEBE cambiar la validación al guardar.
- R16 [D] CUANDO se guarda una cotización con orden ligada, `OrderNumber` DEBE conservarse.

**Apertura de `frmQuote`**
- R17 [N] CUANDO `frmQuote` ya tenía cambios sin guardar y se le pide abrir una cotización, DEBE preguntar con `MSG_UNSAVED_LOAD`. Si el usuario dice No, conserva lo que tenía; la liga creada en R2 queda guardada.
- R18 [N] Ningún texto nuevo DEBE nombrar a un taller o cliente en particular: todos van en `modStrings`.

## Design

### Módulos y archivos
| Archivo | Cambio |
|---|---|
| `src/modules/modConfig.bas` | `Public Const MAX_ITEM_DESCRIPTION As Long = 255` junto a `MAX_ITEMS` |
| `src/modules/modUtils.bas` | Sección nueva `'--- Orden a cotización ---` con `DiagnosisLines`, `QuoteFromOrder` y `OrderBannerText` (puras: sin hojas ni formularios) |
| `src/modules/modData.bas` | `CreateQuoteForOrder` y `DeleteQuote` (privada) |
| `src/modules/modApp.bas` | `ShowQuoteFormWith` |
| `src/modules/modStrings.bas` | Textos nuevos (abajo) |
| `src/forms/frmOrder.vba` | Botón `btnToQuote` y `btnToQuote_Click` |
| `src/forms/frmQuote.vba` | `lblOrderBanner`, `RefreshOrderBanner`, `OpenQuote` público y `ShowSavedQuote` privado, que comparten «Cargar» y `OpenQuote` |
| `src/tests/modTestOrderQuote.bas` | Pruebas nuevas (T1) |
| `README.md` | Cómo pasar una orden a cotización |

### `modUtils` (firmas exactas)
```vb
' Renglones del diagnóstico: parte en vbCrLf/vbLf/vbCr, recorta y omite vacíos. Colección de String (vacía si no hay).
Public Function DiagnosisLines(ByVal diagnosisText As String) As Collection

' Cotización nueva (sin guardar) para la orden: Number = quoteNumber, ClientName = o.ClientName,
' OrderNumber = o.Number, QuoteDate = Date, TaxRate por omisión. Un concepto NewQuoteItem(1, renglón, 0)
' por cada uno de los primeros MAX_ITEMS renglones, cortados a MAX_ITEM_DESCRIPTION. No toca la orden.
Public Function QuoteFromOrder(ByVal o As ServiceOrder, ByVal quoteNumber As String) As Quote

' Texto de la franja. o puede ser Nothing (orden no encontrada). orderNumber vacío → "".
Public Function OrderBannerText(ByVal orderNumber As String, ByVal o As ServiceOrder) As String
```
- `DiagnosisLines`: `Replace(vbCrLf → vbLf)`, luego `Replace(vbCr → vbLf)`, `Split` por `vbLf` y `Trim$` de cada renglón. Solo se agregan los que tienen `Len > 0`.
- `OrderBannerText`: el separador es `" · "` (un espacio a cada lado; no se usa `JoinNonEmpty`, que pone tres). Vehículo = `Trim$(Make)`, `Trim$(Model)` y `CStr(Year)` si `Year > 0`, unidos con un espacio y sin los vacíos. Placas = `UCase$(Trim$(Plates))`. Plantillas `UI_ORDER_BANNER` («Orden {0}») y `UI_ORDER_BANNER_PLATES` («Placas {0}»).

### `modData.CreateQuoteForOrder`
```vb
' Crea, guarda y liga una cotización nueva a la orden (que ya debe estar guardada).
' Regresa el folio, o "" si algo falló (ShowError ya avisó). Si falla SaveOrder,
' borra la cotización nueva y devuelve o.QuoteNumber a su valor anterior.
Public Function CreateQuoteForOrder(ByVal o As ServiceOrder) As String
```
Pasos: `o Is Nothing` → "". `q = QuoteFromOrder(o, GetNextQuoteNumber())`; si `SaveQuote(q)` falla → `DeleteQuote q.Number` y "". `previous = o.QuoteNumber`; `o.QuoteNumber = q.Number`; si `SaveOrder(o)` falla → `o.QuoteNumber = previous`, `DeleteQuote q.Number` y "". `DeleteQuote` (privada) llama `DeleteRowsByNumber` sobre `BD_Cotizaciones` y `BD_Detalle`. No decide si reabrir o reemplazar: eso lo hace `frmOrder`.

### `modApp`
```vb
' Abre frmQuote sin modal con la cotización guardada de ese folio.
Public Sub ShowQuoteFormWith(ByVal quoteNumber As String)
    frmQuote.OpenQuote quoteNumber
    frmQuote.Show vbModeless
End Sub
```

### `frmQuote`
- `Private lblOrderBanner As MSForms.Label` = `AddLabel(Me, "c_OrderBanner", "", 330, 61, 312, 14, 8, True, mSecondary, fmTextAlignRight)` en `BuildLayout`, después de `c_Sec1` (que ocupa x 18 a 318), con `Visible = False`. No es tabulable porque es etiqueta.
- `Private Sub RefreshOrderBanner()`: si `Len(Trim$(mQuote.OrderNumber)) = 0`, oculta la franja. Si no, `Caption = OrderBannerText(mQuote.OrderNumber, LoadOrder(mQuote.OrderNumber))` y `Visible = True`. Se llama al final de `StartNewQuote` y de `ShowSavedQuote`.
- `Private Function ShowSavedQuote(ByVal quoteNumber As String) As Boolean`: es el cuerpo actual de `btnLoad_Click` desde `LoadQuote` (no encontrado → `MSG_QUOTE_NOT_FOUND`; con cambios → `MSG_UNSAVED_LOAD`; llena campos, `FillTemplate`, `mIsSaved = True`, `mOverwriteConfirmed = False`, `mIsDirty = False`) más `RefreshOrderBanner`. Regresa True si cargó. `btnLoad_Click` conserva la validación del buscador vacío y luego llama a esta función.
- `Public Function OpenQuote(ByVal quoteNumber As String) As Boolean`: `LoadHistoryList` (la cotización nueva aún no está en el combo) y luego `ShowSavedQuote(UCase$(Trim$(quoteNumber)))`, con `On Error` → `ShowError "OpenQuote"`.
- `SyncQuoteFromForm` y `SaveCurrentQuote` no cambian: `mQuote.OrderNumber` viaja en el objeto (R16).

### `frmOrder`
- `Private WithEvents btnToQuote As MSForms.CommandButton` = `AddButton(Me, "c_ToQuote", UI_BTN_ORDER_TO_QUOTE, kindSecondary, rightEdge - 130, 72, 130, 22)`, en la franja de folio y estado, alineado con el borde derecho de las pestañas. `ControlTipText = TIP_ORDER_TO_QUOTE`. Tab: cboStatus 0, mpgMain 1, btnSave 2, btnToQuote 3, btnNew 4, btnCompany 5, cboHistory 6, btnLoad 7, btnSwitch 8.
- `btnToQuote_Click`, en este orden:
  1. `If Not SaveCurrentOrder() Then Exit Sub` (R1).
  2. `linked = UCase$(Trim$(mOrder.QuoteNumber))`. Si `Len(linked) > 0 And QuoteExists(linked)`, entonces `quoteNumber = linked` (R3).
  3. Si no: cuando `Len(linked) > 0`, se muestra `MsgBox Fmt(MSG_LINKED_QUOTE_MISSING, linked), vbExclamation, TTL_REVIEW` (R4). Luego `quoteNumber = CreateQuoteForOrder(mOrder)`; si es "", `Exit Sub`. Si `DiagnosisLines(mOrder.Diagnosis).Count > MAX_ITEMS`, se muestra `MsgBox Fmt(MSG_DIAGNOSIS_TRUNCATED, n, MAX_ITEMS), vbInformation` (R9).
  4. `mIsDirty = False`, `ShowQuoteFormWith quoteNumber`, `Unload Me` (R6). `On Error` → `ShowError "btnToQuote_Click"`.

### Textos nuevos (`modStrings`, Windows-1252)
| Constante | Texto |
|---|---|
| `UI_BTN_ORDER_TO_QUOTE` | Pasar a cotización |
| `TIP_ORDER_TO_QUOTE` | Guarda la orden y abre su cotización con el diagnóstico como conceptos |
| `UI_ORDER_BANNER` | Orden {0} |
| `UI_ORDER_BANNER_PLATES` | Placas {0} |
| `MSG_LINKED_QUOTE_MISSING` | No se encontró la cotización {0} ligada a esta orden.\nSe creará una nueva con el diagnóstico. |
| `MSG_DIAGNOSIS_TRUNCATED` | El diagnóstico tiene {0} renglones y la cotización admite {1}.\nLos demás quedan solo en la orden. |

### Decisiones descartadas
- Pasar a `frmQuote` un `Quote` armado en memoria y sin guardar: entonces la orden apuntaría a un folio que no existe hasta que alguien guardara la cotización.
- Bloquear el botón si la orden ya tiene cotización: el discovery pide reabrirla.
- Volver a armar los conceptos de una cotización ya ligada cuando cambia el diagnóstico: borraría los precios que el usuario ya puso.
- Usar `JoinNonEmpty` en la franja: su separador de tres espacios no cabe en 312 pt.

### Superficie de seguridad
- Escritura en hojas: solo `modData` (`SaveQuote`, `SaveOrder`, `DeleteQuote`). Las descripciones pasan por `SafeCellText` dentro de `SaveQuote`, así que un renglón que empieza con `-` o `=` no se vuelve fórmula.
- Validación: la orden pasa por `SyncOrderFromForm` antes de todo. La cotización con 0 conceptos se guarda directo con `SaveQuote` (no pasa por `SyncQuoteFromForm`), pero para volver a guardarla desde `frmQuote` hacen falta conceptos.
- Errores: `CreateQuoteForOrder` revierte la liga a medias. Los handlers de los formularios terminan en `ShowError`.
- Datos sensibles: ninguno nuevo. La franja muestra placas, que ya están en la orden; el VIN y el teléfono no aparecen.

## Tasks

- [ ] **T1 · Contrato: funciones puras y pruebas** (R7-R11, R13, R14, R18): `MAX_ITEM_DESCRIPTION` en `modConfig`; `DiagnosisLines`, `QuoteFromOrder` y `OrderBannerText` en `modUtils` con las firmas de Design, ya implementadas; `UI_ORDER_BANNER` y `UI_ORDER_BANNER_PLATES` en `modStrings`. `src/tests/modTestOrderQuote.bas` con todas las pruebas de abajo; las de `CreateQuoteForOrder` pueden fallar hasta T2. Pruebas:
  - `Test_DiagnosisLines_PartePorSaltos`: «  Balatas \r\n\r\nDisco\nAceite\rFiltro » → «Balatas», «Disco», «Aceite», «Filtro».
  - `Test_DiagnosisLines_SinSaltosUnRenglon`: «Afinación mayor» → 1 renglón.
  - `Test_DiagnosisLines_VacioSinRenglones`: «» y «  \r\n  » → `Count = 0`.
  - `Test_DiagnosisLines_ConservaVinetas`: «- Balatas» y «1. Disco» quedan tal cual.
  - `Test_QuoteFromOrder_Encabezado`: `NewSampleOrder("OS-0001")` con diagnóstico de 2 renglones y «COT-1005» → Number, ClientName, OrderNumber = «OS-0001», `ItemCount = 2`, `Quantity = 1` y `UnitPrice = 0` en los dos, y `Total = 0`.
  - `Test_QuoteFromOrder_TopeMaxItems`: 20 renglones → `ItemCount = MAX_ITEMS` y el último es el renglón 17.
  - `Test_QuoteFromOrder_DescripcionLarga`: un renglón de 300 caracteres → `Len(Description) = 255`.
  - `Test_QuoteFromOrder_NoCambiaDiagnostico`: `o.Diagnosis` y `o.QuoteNumber` siguen igual después de la llamada.
  - `Test_OrderBannerText_Completo`: muestra con Make «Marca X», Model «Modelo Y», 2018, placas «abc123d» → «Orden OS-0001 · Marca X Modelo Y 2018 · Placas ABC123D».
  - `Test_OrderBannerText_OmiteVacios`: sin modelo ni año → «Orden OS-0001 · Marca X · Placas ABC123D»; sin vehículo ni placas → «Orden OS-0001».
  - `Test_OrderBannerText_OrdenNoEncontrada`: `o = Nothing` → «Orden OS-0009»; folio "" → "".
  - `Test_CreateQuoteForOrder_LigaAmbosLados`: orden guardada sin `QuoteNumber` → regresa «COT-1001»; `LoadQuote` tiene `OrderNumber = "OS-0001"` y un concepto; `LoadOrder` tiene `QuoteNumber = "COT-1001"`.
  - `Test_CreateQuoteForOrder_ReemplazaFolioPerdido`: orden con `QuoteNumber = "COT-1050"`, que no existe → regresa «COT-1001» y la orden guardada apunta a «COT-1001».
  - `Test_CreateQuoteForOrder_DiagnosticoVacio`: diagnóstico «» → cotización guardada con `ItemCount = 0`.
  - `Test_CreateQuoteForOrder_NothingRegresaVacio`: `CreateQuoteForOrder(Nothing) = ""` y `QuoteExists("COT-1001") = False`.
- [ ] **T2 · Liga en `modData`** (R2, R4, R5): `CreateQuoteForOrder` y `DeleteQuote` privada, con la reversión de Design. Las 15 pruebas de T1 pasan en `test.ps1`.
- [ ] **T3 · `frmQuote`: franja y `OpenQuote`** (R12, R14-R17): `lblOrderBanner`, `RefreshOrderBanner` en `StartNewQuote` y `ShowSavedQuote`, `btnLoad_Click` refactorizado sobre `ShowSavedQuote`, `OpenQuote` público y `modApp.ShowQuoteFormWith`. «Cargar» funciona igual que antes y además muestra u oculta la franja.
- [ ] **T4 · `frmOrder`: «Pasar a cotización»** (R1-R4, R6, R9, R18): `btnToQuote` con su posición, tooltip y orden de tabulación, `btnToQuote_Click` en el orden de Design y los textos `UI_BTN_ORDER_TO_QUOTE`, `TIP_ORDER_TO_QUOTE`, `MSG_LINKED_QUOTE_MISSING` y `MSG_DIAGNOSIS_TRUNCATED`.
- [ ] **T5 · Documentación y verificación manual** (R1, R3, R4, R6, R9, R12, R15, R17): una sección en `README.md` sobre cómo pasar la orden a cotización (qué se copia, qué pasa si ya tiene cotización y el tope de 17 conceptos). Prueba manual en Excel después del build:
  - Una orden nueva válida con diagnóstico de 3 renglones abre `frmQuote` con 3 conceptos en $0.00, la franja completa y `frmOrder` cerrado.
  - Una orden inválida (sin kilometraje) avisa y no abre nada.
  - Pasar de nuevo la misma orden abre la misma cotización, con los precios que ya se le pusieron.
  - Si se borra a mano la fila de la cotización en `BD_Cotizaciones`, aparece el aviso y se crea un folio nuevo.
  - Un diagnóstico de 20 renglones da el aviso de tope.
  - Al cargar desde el historial una cotización con orden aparece la franja; con una sin orden no aparece, y «Nueva cotización» la oculta.
  - La franja no tapa el título de la sección, con logo y sin logo.
