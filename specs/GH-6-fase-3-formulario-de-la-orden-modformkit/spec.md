# GH-6 · Fase 3 · Formulario de la orden (modFormKit + frmOrder)

## Brief

**Objetivo:** que el taller capture, guarde y cargue órdenes de servicio en un formulario `frmOrder` que cabe en 1366×768, con los helpers de maquetado compartidos con `frmQuote` en `modFormKit`.

**Entra**
- `modFormKit`: los helpers de maquetado de `frmQuote` (`Place`, `StyleInput`, `AddTextBox`, `AddComboBox`, `AddButton`, `AddLabel`, `AddBand`, `PlaceHeaderLogo`, `FocusOn`, `Warn`) pasan a un módulo público que recibe el contenedor (formulario o pestaña). `frmQuote` se ve y se comporta igual.
- `frmOrder` (660×496 pt, como `frmQuote`) con `MultiPage` de dos pestañas: «Cliente y vehículo» y «Entrada, servicio y entrega». Folio y estado quedan fuera de las pestañas, visibles siempre.
- Validadores puros con pruebas en `modUtils`: `IsValidDate`, `IsValidTime`, `IsValidVehicleYear`, `IsValidMileage`, `IsValidVin`. `frmQuote` cambia su `IsDate` por `IsValidDate`.
- `modData.ListOrderStaff`: los nombres distintos de «Recibió» y «Entregó» ya guardados, para los dos combos (también se puede escribir un nombre nuevo).
- Cliente con autocompletado y contacto con `GetClientContact` (solo llena campos vacíos).
- Guardar, Nueva orden, Cargar del historial (Enter = Cargar) y aviso al cerrar con cambios, todo igual que en la cotización.
- Selector entre formularios: la etiqueta del módulo actual y, junto a ella, un botón «Orden de servicio» o «Cotización» que pregunta si hay cambios, cierra este formulario y abre el otro.
- Build: `lib.ps1` crea `frmOrder` y el logo `imgLogo` va a los dos formularios.

**No entra:** PDF de la orden (Fase 5, #8), pasar la orden a cotización (Fase 4, #7; `QuoteNumber` se conserva tal como se cargó), botón propio en Inicio, catálogo de personal y reglas de transición de estados.

**Decisiones nuevas**
- [N] El selector es la etiqueta del módulo más un botón secundario que lleva al otro formulario. Descartado: un control de dos segmentos, porque el segmento activo sería un segundo botón primario en la vista o un control hecho a mano.
- [N] El aviso del VIN es una etiqueta junto al campo («El VIN suele tener 17 caracteres; lleva N») y no interrumpe. Descartado: un MsgBox Sí/No al guardar, que se vuelve un obstáculo con cada carro viejo.
- [N] Son obligatorios el cliente y la fecha de entrada, además del kilometraje y la gasolina. Fecha prometida, fecha y hora de salida, hora de entrada y año pueden ir vacíos, pero si se escriben deben ser válidos. Descartado: dejar el cliente opcional, porque una orden sin cliente no se encuentra en el historial.
- [N] `modFormKit` guarda la fuente y los colores que le pasa una vez `BeginFormKit` desde `BuildLayout`. Descartado: pasarlos en cada llamada (demasiados parámetros) o leerlos de «Mi empresa» en cada control (rompe la lectura única de GH-18).
- [N] Selector: el aviso con cambios sin guardar tiene el mismo flujo que Cargar (Sí/No, `TTL_REVIEW`), con un texto nuevo, `MSG_UNSAVED_SWITCH`, porque «¿Cargar … de todos modos?» no tiene sentido aquí.
- [N] Una orden nueva arranca con la fecha y la hora de entrada actuales y en estado «Recibida».

**Riesgos**
- `Enter`, `Exit` y `AfterUpdate` no llegan a los controles declarados con `WithEvents`. Por eso el contacto se llena en `Click` y en Enter/Tab del combo de cliente: se prueba a mano.
- Al pasar a `modFormKit` se puede mover un píxel o un color de `frmQuote`: se compara a mano contra el build anterior.
- El logo de respaldo de `PlaceHeaderLogo` se llamaba `c_Logo`, el mismo nombre que el botón «Mi logo». Pasa a llamarse `c_HeaderLogo`.

**Tamaño:** grande, unas 4 h. Son 6 tareas, 9 pruebas automáticas nuevas y una verificación manual de los dos formularios.

## Discovery

# Discovery GH-6 · Formulario de la orden (modFormKit + frmOrder)

## Alcance
- Refactor previo: extraer de frmQuote los helpers privados (Place, StyleInput, AddTextBox, AddComboBox, AddButton, AddLabel, AddBand; Warn si aplica) a un módulo modFormKit que recibe el formulario como parámetro. frmQuote debe verse y comportarse igual.
- frmOrder nuevo, construido por código como frmQuote, con colores según GH-18 (ContrastTextColor, lectura única en BuildLayout).

## Decisiones
1. Pantalla: MultiPage SIEMPRE, con dos pestañas: "Cliente y vehículo" y "Entrada, servicio y entrega". Debe caber en 1366×768 y servir a cualquier taller.
2. Fechas: crear IsValidDate en modUtils (función pura, con prueba automática). frmQuote deja su IsDate inline y usa IsValidDate; frmOrder también.
3. ComboBox "quién recibió / entregó": la lista son los valores distintos ya guardados en las órdenes (BD_Ordenes, vía modData); se puede escribir un nombre nuevo. Sin hoja ni configuración extra.
4. Selector "Orden de servicio | Cotización": cierra un formulario y abre el otro; si hay cambios sin guardar, pregunta antes con el mismo aviso que al cargar (MSG_UNSAVED_LOAD). Si cancela, se queda donde está.
5. Autocompletado del cliente: igual que en la cotización (ClientExists/ListClients) y llenado de contacto con GetClientContact (reemplazo de LastOrderForClient desde GH-4; BD_Clientes con la orden más reciente de respaldo). Solo llena campos vacíos.
6. Kilometraje y gasolina obligatorios: bloquean el guardado (Warn + foco en el campo), como los obligatorios de la cotización.

## Controles y validaciones
- ComboBox para gasolina (R, ¼, ½, ¾, Lleno, desde modConfig), estado (modConfig) y quién recibió/entregó.
- Hora formato hh:mm; año entre 1950 y año actual + 1; kilometraje solo números; placas en mayúsculas; VIN de 17 caracteres solo avisa, no bloquea.
- Validadores puros (hora, año, kilometraje, VIN) en modUtils con pruebas automáticas, siguiendo el precedente IsValidRfc.
- Botón Cargar (historial) funciona igual que en la cotización (incluye Enter en el buscador = Cargar), usando ListOrders/LoadOrder/SaveOrder de modData.

## Restricciones
- Genérico para cualquier taller: nada propio de un cliente en el código.
- modData es el único que toca hojas BD_*.
- Formularios, apariencia y navegación siguen siendo prueba manual.

## Requirements

**Kit de maquetado y cotización**
- **R1** [D] CUANDO un formulario construye su interfaz, el sistema DEBE usar los helpers públicos de `modFormKit`, que reciben el contenedor (formulario o pestaña de `MultiPage`). En `frmQuote` no DEBE quedar ninguna copia privada de esos helpers.
- **R2** [D] CUANDO se abre `frmQuote` después del refactor, el sistema DEBE mostrar los mismos controles, posiciones, colores, fuente, orden de tabulador y comportamiento que antes. La única diferencia es el botón del selector (R18).
- **R3** [D] CUANDO `frmQuote` valida la fecha, el sistema DEBE usar `IsValidDate`, que acepta y rechaza los mismos textos que el `IsDate` actual con un texto no vacío.

**Formulario de la orden: estructura**
- **R4** [D] CUANDO se abre `frmOrder`, el sistema DEBE construirlo por código, con un área interior de 660×496 pt como máximo, que cabe en 1366×768. Lleva un `MultiPage` con las pestañas «Cliente y vehículo» y «Entrada, servicio y entrega», y arriba de las pestañas el folio (bloqueado) y el estado.
- **R5** [D] CUANDO `frmOrder` pinta barra, botones y textos, el sistema DEBE leer `PrimaryColor()` y `SecondaryColor()` una sola vez en `BuildLayout` y usar `ContrastTextColor` para el texto sobre los colores de la empresa (regla de GH-18 y `assets/README.md`).
- **R6** [D] CUANDO se muestran los combos de gasolina y de estado, el sistema DEBE llenarlos con `Split(FUEL_LEVELS, "|")` y `Split(ORDER_STATUSES, "|")` y solo dejar elegir de la lista. [N] SI una orden cargada trae un valor que no está en la lista, el sistema DEBE agregarlo al combo y mostrarlo, para no perderlo.
- **R7** [D] CUANDO se muestran los combos «Recibió» y «Entregó», el sistema DEBE ofrecer la lista de `ListOrderStaff()` y permitir escribir un nombre que no esté en ella.
- **R8** [D] CUANDO se llama `ListOrderStaff`, el sistema DEBE regresar un arreglo 0-based con los valores distintos no vacíos de las columnas ReceivedBy (19) y DeliveredBy (23) de BD_Ordenes, sin repetir aunque cambien mayúsculas o espacios, en orden alfabético. Si no hay ninguno, DEBE regresar `Empty`.

**Cliente**
- **R9** [D] CUANDO el usuario escribe en el combo de cliente, el sistema DEBE autocompletar con `ListClients()` igual que en la cotización.
- **R10** [D] CUANDO el usuario elige de la lista o confirma con Enter/Tab un cliente que existe (`ClientExists`), el sistema DEBE llamar `GetClientContact` y llenar solo los campos de teléfono 1, teléfono 2, correo y dirección que estén vacíos.

**Validaciones**
- **R11** [D] CUANDO se guarda, el sistema DEBE rechazar con `MSG_DATE_INVALID` toda fecha (entrada, prometida, salida) escrita que no pase `IsValidDate`.
- **R12** [D] CUANDO se guarda, el sistema DEBE rechazar con `MSG_TIME_INVALID` toda hora (entrada, salida) escrita que no pase `IsValidTime`. Una hora válida es `H:MM` o `HH:MM`, de 00:00 a 23:59.
- **R13** [D] CUANDO se guarda con año escrito, el sistema DEBE rechazarlo con `MSG_YEAR_INVALID` si no pasa `IsValidVehicleYear`: 4 dígitos, entre 1950 y el año actual + 1.
- **R14** [D] CUANDO el usuario teclea en el kilometraje, el sistema DEBE ignorar todo lo que no sea dígito o retroceso. CUANDO se guarda, el kilometraje DEBE pasar `IsValidMileage`: de 1 a 9 dígitos.
- **R15** [D] CUANDO el usuario teclea o pega en placas o VIN, el sistema DEBE dejar el texto en mayúsculas.
- **R16** [D] MIENTRAS el VIN tenga texto y no pase `IsValidVin` (17 caracteres alfanuméricos), el sistema DEBE mostrar junto al campo un aviso con el largo actual. El aviso NO DEBE impedir el guardado.
- **R17** [D] CUANDO se guarda sin kilometraje o sin gasolina, el sistema DEBE bloquear el guardado con `Warn`, cambiar a la pestaña del campo y dejar el foco en él. [N] Lo mismo con el cliente vacío (`MSG_CLIENT_EMPTY`) y con la fecha de entrada vacía. Toda advertencia de R11-R14 también cambia a la pestaña del campo.

**Navegación, guardado e historial**
- **R18** [D] CUANDO se pulsa el botón del selector («Orden de servicio» en `frmQuote`, «Cotización» en `frmOrder`) con cambios sin guardar, el sistema DEBE preguntar con `MSG_UNSAVED_SWITCH` (Sí/No). Con No, se queda en el formulario actual sin cambios. Con Sí, o si no hay cambios, DEBE abrir el otro formulario sin modal y cerrar el actual sin volver a preguntar.
- **R19** [D] CUANDO se pulsa Cargar, o Enter en el buscador del historial, el sistema DEBE seguir el flujo de la cotización: folio vacío → `MSG_HISTORY_EMPTY`; no existe → `MSG_QUOTE_NOT_FOUND`; cambios sin guardar → `MSG_UNSAVED_LOAD`. Luego llena el formulario con `LoadOrder` sin marcarlo como cambiado. El historial muestra Folio, Cliente, Placas y Estado (`ListOrders`).
- **R20** [N] CUANDO se guarda una orden nueva cuyo folio ya existe, el sistema DEBE asignar `GetNextOrderNumber()` y avisar con `MSG_NUMBER_REASSIGNED`. CUANDO se guarda por primera vez una orden cargada, DEBE confirmar la sobrescritura con `MSG_CONFIRM_OVERWRITE` una sola vez. Después de guardar, DEBE recargar el historial, los clientes y `ListOrderStaff`. `QuoteNumber` se guarda tal como se cargó.
- **R21** [N] CUANDO se pulsa «Nueva orden» o se cierra `frmOrder` con cambios sin guardar, el sistema DEBE preguntar antes (`MSG_UNSAVED_NEW_ORDER` / `MSG_UNSAVED_CLOSE_ORDER`). Una orden nueva arranca con `GetNextOrderNumber()`, estado `ORDER_STATUS_RECEIVED` y la fecha y la hora de entrada actuales.

**Build y generalidad**
- **R22** [N] CUANDO se corre `build.ps1` o `test.ps1`, el sistema DEBE crear `frmOrder` desde `src/forms/frmOrder.vba`. CUANDO exista `assets/logo-form.jpg`, el logo `imgLogo` DEBE quedar en `frmQuote` y en `frmOrder`.
- **R23** [D] El sistema NO DEBE tener en código ningún nombre, color ni texto propio de un taller. Todo texto visible nuevo vive en `modStrings`.

## Design

### Módulos y archivos
| Archivo | Cambio |
|---|---|
| `src/modules/modFormKit.bas` | Nuevo: helpers de maquetado públicos |
| `src/forms/frmQuote.vba` | Usa `modFormKit`; `IsValidDate`; botón del selector |
| `src/forms/frmOrder.vba` | Nuevo |
| `src/modules/modUtils.bas` | Validadores puros |
| `src/modules/modConfig.bas` | `MIN_VEHICLE_YEAR` |
| `src/modules/modData.bas` | `ListOrderStaff` |
| `src/modules/modApp.bas` | `ShowQuoteForm`, `ShowOrderForm` |
| `src/modules/modStrings.bas` | Textos `UI_ORD_*`, `MSG_*`, `TTL_*` nuevos |
| `src/build/modBuild.bas` | `AddFormLogo` recibe el nombre del formulario y se llama para los dos |
| `build/lib.ps1` | Crea los formularios en un ciclo: `frmQuote`, `frmOrder` |
| `src/tests/modTestUtils.bas`, `modTestData.bas` | Pruebas nuevas |
| `README.md`, `.claude/skills/vba-cotizador/SKILL.md` | Estructura (`modFormKit`, `frmOrder`) y la regla «los helpers viven en `modFormKit`» |

### `modFormKit` (firmas exactas)
```vba
Public Enum ButtonKind
    kindPrimary = 0      ' Color principal: acción principal (uno por vista)
    kindDark = 1         ' Color secundario
    kindSecondary = 2    ' Blanco
End Enum
Public Const FORM_MARGIN As Single = 18

Public Sub BeginFormKit(ByVal fontName As String, ByVal primaryColor As Long, ByVal secondaryColor As Long)
Public Sub Place(ByVal ctl As Object, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
Public Sub StyleInput(ByVal ctl As Object)
Public Function AddTextBox(ByVal parent As Object, ByVal ctlName As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single) As MSForms.TextBox
Public Function AddComboBox(ByVal parent As Object, ByVal ctlName As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single) As MSForms.ComboBox
Public Function AddButton(ByVal parent As Object, ByVal ctlName As String, ByVal buttonText As String, ByVal kind As ButtonKind, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single) As MSForms.CommandButton
Public Function AddLabel(ByVal parent As Object, ByVal ctlName As String, ByVal labelText As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, ByVal fontSize As Single, ByVal isBold As Boolean, ByVal textColor As Long, Optional ByVal align As fmTextAlign = fmTextAlignLeft) As MSForms.Label
Public Sub AddBand(ByVal parent As Object, ByVal ctlName As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, ByVal fillColor As Long, ByVal toBack As Boolean)
Public Function PlaceHeaderLogo(ByVal frm As Object) As Boolean
Public Sub FocusOn(ByVal ctl As Object)
Public Sub Warn(ByVal message As String, ByVal ctl As Object)
```
- `parent` es cualquier objeto con `.Controls.Add`: el `UserForm` (`Me`) o una `Page` del `MultiPage`.
- El cuerpo de cada helper es el actual de `frmQuote`, cambiando `Me.Controls` por `parent.Controls` y `mFont`/`mPrimary`/`mSecondary` por las variables privadas del módulo que llena `BeginFormKit`. Cada formulario sigue guardando sus propios `mPrimary`/`mSecondary` para los repintados que hace después (por ejemplo `ExitEditMode`). La constante privada `MARGIN` de `frmQuote` puede quedarse o cambiarse por `FORM_MARGIN`.
- `PlaceHeaderLogo` busca el `Image` del formulario (el `imgLogo` del build) igual que hoy. El respaldo desde archivo se llama `c_HeaderLogo` (antes `c_Logo`, que choca con el botón «Mi logo»).
- `BeginFormKit` se llama al inicio de cada `BuildLayout`. Solo hay un formulario abierto a la vez (R18), así que el estado del módulo no se mezcla.
- `FocusOn` y `Warn` no cambian. Antes de `Warn`, `frmOrder` cambia `mpgMain.Value` a la pestaña del control (ver `WarnOnPage`).

### Validadores (`modUtils`, sección `'--- Validación de campos ---`)
```vba
Public Function IsValidDate(ByVal dateText As String) As Boolean            ' Len(Trim)>0 And IsDate(Trim)
Public Function IsValidTime(ByVal timeText As String) As Boolean            ' "H:MM" o "HH:MM", 0-23 : 0-59
Public Function IsValidVehicleYear(ByVal yearText As String, Optional ByVal referenceDate As Date = 0) As Boolean
    ' 4 dígitos, MIN_VEHICLE_YEAR <= año <= Year(ref)+1; ref = Date si referenceDate = 0 (para probar con fecha fija)
Public Function IsValidMileage(ByVal mileageText As String) As Boolean      ' 1 a 9 dígitos [0-9], sin separadores
Public Function IsValidVin(ByVal vinText As String) As Boolean              ' Trim, 17 caracteres [A-Z0-9] sin importar mayúsculas
' modConfig, sección Orden de servicio:
Public Const MIN_VEHICLE_YEAR As Long = 1950
```
Todas usan `Trim$`, no tocan hojas ni formularios y regresan False con texto vacío. Si un campo es opcional, quien llama decide si lo vacío se permite.

### `modData.ListOrderStaff`
```vba
' Nombres distintos de Recibió/Entregó en BD_Ordenes, orden alfabético; Empty si no hay.
Public Function ListOrderStaff() As Variant
```
Lee `ReadData(SHEET_ORDERS, ORDER_COLUMNS)` una sola vez, junta las columnas 19 y 23, quita espacios a los lados, descarta las cadenas vacías y omite repetidos con `StrComp(..., vbTextCompare)`. Ordena por inserción, como `ListClients`, y conserva la primera forma escrita.

### `modApp`
```vba
Public Sub ShowQuoteForm()   ' frmQuote.Show vbModeless
Public Sub ShowOrderForm()   ' frmOrder.Show vbModeless
```
`OpenQuoteTool` sigue llamando `PrepareWorkbook` y después `ShowQuoteForm`. El selector no vuelve a preparar el libro, que ya quedó listo.

### Selector (los dos formularios)
- En la barra superior: la etiqueta `c_Module` muestra el módulo actual (`UI_MODULE` = «Cotizador» en `frmQuote`, `UI_ORD_MODULE` = «Orden de servicio» en `frmOrder`). Junto a ella va `btnSwitch` (`c_Switch`, `kindSecondary`, 100×22, y=13) con `UI_BTN_TO_ORDER` = «Orden de servicio» o `UI_BTN_TO_QUOTE` = «Cotización». En `frmQuote` va en x=268 con logo y x=230 sin logo. Esos huecos están libres hoy: la etiqueta termina en x=262 y el historial empieza en x=374.
- `btnSwitch_Click`:
  ```vba
  If mIsDirty Then If MsgBox(Fmt(MSG_UNSAVED_SWITCH, <destino>), vbYesNo + vbExclamation, TTL_REVIEW) = vbNo Then Exit Sub
  mIsDirty = False          ' Para que QueryClose no vuelva a preguntar
  ShowOrderForm             ' (o ShowQuoteForm)
  Unload Me
  ```
- `MSG_UNSAVED_SWITCH = "Hay cambios sin guardar que se perderán.\n¿Ir a {0} de todos modos?"`. `{0}` es el nombre del otro formulario en minúsculas («la orden de servicio» o «la cotización»), con `UI_ORD_SWITCH_TARGET` y `UI_QUOTE_SWITCH_TARGET`.

### `frmOrder`: estado y controles
Estado: `mOrder As ServiceOrder`, `mIsSaved`, `mIsDirty`, `mSuppressEvents`, `mOverwriteConfirmed`, `mFont`, `mPrimary`, `mSecondary`. Tienen el mismo significado que en `frmQuote`.

Medidas: `FORM_WIDTH = 660`, `FORM_HEIGHT = 496`, barra superior de 48 pt igual que `frmQuote` (logo o marca, `c_Module`, `btnSwitch`, `cboHistory` de 4 columnas «70;150;70;80 pt», `btnLoad`) y pie desde y=450.

| Zona | Controles (nombre de variable · nombre del control) |
|---|---|
| Franja fija (y≈58-100) | `txtNumber`·c_Number (bloqueado), `cboStatus`·c_Status (lista) |
| `mpgMain`·c_Pages (y≈106, alto≈336) | Pestaña 0 `UI_ORD_TAB_CLIENT`, pestaña 1 `UI_ORD_TAB_SERVICE` |
| Pestaña 0, sección «CLIENTE» | `cboClient`·c_Client, `txtPhone1`·c_Phone1, `txtPhone2`·c_Phone2, `txtEmail`·c_Email, `txtAddress`·c_Address |
| Pestaña 0, sección «VEHÍCULO» | `txtMake`·c_Make, `txtModel`·c_Model, `txtEngine`·c_Engine, `txtYear`·c_Year (MaxLength 4), `txtColor`·c_Color, `txtPlates`·c_Plates (MaxLength 10), `txtMileage`·c_Mileage (MaxLength 9, a la derecha), `cboFuel`·c_Fuel (lista), `txtVin`·c_Vin (MaxLength 20), `lblVinHint`·c_VinHint |
| Pestaña 1, sección «ENTRADA» | `txtInDate`·c_InDate, `txtInTime`·c_InTime (MaxLength 5), `cboReceivedBy`·c_ReceivedBy, `txtPromisedDate`·c_PromisedDate |
| Pestaña 1, sección «SERVICIO» | `txtRequested`·c_Requested, `txtDiagnosis`·c_Diagnosis (MultiLine, EnterKeyBehavior, barra vertical, MaxLength 1000, ≈64 pt de alto) |
| Pestaña 1, sección «ENTREGA» | `txtOutDate`·c_OutDate, `txtOutTime`·c_OutTime, `cboDeliveredBy`·c_DeliveredBy |
| Pie | `lblInnobytes`, `btnNew`·c_New (`UI_BTN_NEW_ORDER`), `btnCompany`·c_Company, `btnSave`·c_Save (`kindPrimary`, el único primario) |

- Rejilla del skill: margen 18 pt, controles de 22 pt, etiqueta de 8 pt en `TextGray()` sobre cada campo y título de sección en mayúsculas, 8 pt, negritas y `mPrimary`. Las coordenadas exactas dentro de cada pestaña las ajusta quien implemente, sin encimar nada y sin pasar de 496 pt.
- Con eventos (`WithEvents`) van todos los campos editables (`_Change` → `MarkDirty`), más `cboClient_Click`/`_KeyDown`, `txtMileage_KeyPress`, `txtYear_KeyPress` (solo dígitos), `txtPlates_KeyPress`/`txtVin_KeyPress` (pasar `KeyAscii` a mayúscula), `txtPlates_Change`/`txtVin_Change` (`UCase$` del texto pegado, protegido con `mSuppressEvents`), `txtVin_Change` → `UpdateVinHint`, `cboHistory_KeyDown`, botones y `lblInnobytes_Click`.
- `TabIndex` en el orden de captura dentro de cada pestaña.
- Procedimientos privados: `BuildLayout`, `BuildClientPage`, `BuildServicePage`, `StartNewOrder`, `ShowOrder` (llena controles desde `mOrder`), `SyncOrderFromForm() As Boolean`, `SaveCurrentOrder() As Boolean`, `FillClientContact`, `UpdateVinHint`, `SelectOrAdd(cbo, value)` (R6), `LoadHistoryList`, `LoadClientList`, `LoadStaffLists`, `WarnOnPage(pageIndex, message, ctl)`, `MarkDirty`.

### Reglas de `SyncOrderFromForm` (en este orden; la primera que falla avisa y sale)
1. Cliente vacío → `MSG_CLIENT_EMPTY` (pestaña 0).
2. Año escrito y no `IsValidVehicleYear` → `Fmt(MSG_YEAR_INVALID, Year(Date)+1)` (0).
3. Kilometraje no `IsValidMileage` → `MSG_MILEAGE_REQUIRED` (0).
4. Gasolina con `ListIndex < 0` → `MSG_FUEL_REQUIRED` (0).
5. Fecha de entrada no `IsValidDate` → `MSG_DATE_INVALID` (1). Prometida y salida, solo si están escritas.
6. Horas de entrada y salida escritas y no `IsValidTime` → `MSG_TIME_INVALID` (1).
7. Copia a `mOrder`: texto con `Trim$`, `Year` con `CInt` (0 si va vacío), `Mileage` con `CLng`, fechas con `CDate` (0 si van vacías), horas con `TimeValue` (0 si van vacías), placas y VIN con `UCase$`. `Number` y `QuoteNumber` no se tocan.

El VIN no forma parte de la validación (R16).

### Textos nuevos (`modStrings`)
`UI_ORD_CAPTION` «Cotiva · Orden de servicio», `UI_ORD_MODULE`, `UI_BTN_TO_ORDER`, `UI_BTN_TO_QUOTE`, `UI_ORD_SWITCH_TARGET`, `UI_QUOTE_SWITCH_TARGET`, `UI_ORD_TAB_CLIENT`, `UI_ORD_TAB_SERVICE`, `UI_ORD_SEC_CLIENT`, `UI_ORD_SEC_VEHICLE`, `UI_ORD_SEC_IN`, `UI_ORD_SEC_SERVICE`, `UI_ORD_SEC_OUT`, una `UI_ORD_LBL_*` por campo («Teléfono 1», «Kilometraje *», «Gasolina *», «Hora (hh:mm)»…), `UI_BTN_NEW_ORDER` «Nueva orden», `TIP_ORD_HISTORY`, `TIP_STAFF` («Elige o escribe un nombre»), `MSG_TIME_INVALID`, `MSG_YEAR_INVALID` («El año debe estar entre 1950 y {0}.»), `MSG_MILEAGE_REQUIRED`, `MSG_FUEL_REQUIRED`, `MSG_VIN_HINT` («El VIN suele tener 17 caracteres; lleva {0}.»), `MSG_ORDER_SAVED`, `MSG_UNSAVED_SWITCH`, `MSG_UNSAVED_NEW_ORDER`, `MSG_UNSAVED_CLOSE_ORDER`, `TTL_NEW_ORDER`, `TTL_CLOSE_ORDER`. Se reutilizan `MSG_CLIENT_EMPTY`, `MSG_DATE_INVALID`, `MSG_HISTORY_EMPTY`, `MSG_QUOTE_NOT_FOUND` («No existe el folio {0}.»), `MSG_UNSAVED_LOAD`, `MSG_CONFIRM_OVERWRITE`, `MSG_NUMBER_REASSIGNED`, `TTL_REVIEW` y `TTL_OVERWRITE`.

### Build
- `lib.ps1`: el bloque «Creando formulario» pasa a un ciclo sobre `@('frmQuote','frmOrder')` y lee `src\forms\<nombre>.vba`.
- `modBuild`: queda `AddFormLogo(ByVal formName As String, ByVal imagePath As String)` y se llama para `frmQuote` y para `frmOrder`.
- `export.ps1` ya exporta cualquier formulario: no cambia.

### Decisiones descartadas
- Un `frmOrder` sin `MultiPage`, de unos 720 pt de alto: no cabe en 1366×768 (decisión 1 del discovery).
- Una hoja `BD_Personal` o una lista en «Mi empresa» para quien recibe o entrega: más configuración sin necesidad (decisión 3).
- Llenar el contacto en `cboClient_Change`: con autocompletado, un nombre a medias («Juan») llenaría los datos de otro cliente.
- Un `IsValidDate` estricto `dd/mm/aaaa`: cambiaría lo que acepta hoy la cotización (R2/R3). Puede ser otra tarea.

### Superficie de seguridad
- **Validación:** todo lo que entra a la hoja pasa por `SaveOrder`, que ya aplica `SafeCellText` y formatos de texto a teléfonos, placas y VIN. El formulario valida tipos y rangos antes de llamar y convierte con `CInt`/`CLng` solo después de pasar el validador, así que no hay desbordes.
- **Errores:** cada `_Click` con acción usa `On Error GoTo Fail` + `ShowError "<proc>"`, y `mSuppressEvents = False` en `Fail` donde se apaga (igual que `btnLoad_Click`).
- **Authz:** el libro sigue protegido como hoy (`PrepareWorkbook`/`ProtectAll`); `frmOrder` solo escribe a través de `modData`.
- **Datos sensibles:** teléfonos, correo y dirección del cliente se guardan como hoy (GH-4). No se agregan logs ni se exporta nada.

## Tasks

- [ ] **T1 · Contrato y funciones puras** (R1, R3, R8, R11-R16): `modFormKit` con las firmas de Design (incluidos `ButtonKind` y `FORM_MARGIN`) ya implementado. También quedan funcionando los validadores de `modUtils`, `MIN_VEHICLE_YEAR`, `modData.ListOrderStaff` y `modApp.ShowQuoteForm`, que `OpenQuoteTool` ya usa. `frmQuote` todavía no cambia. Como `ShowOrderForm` necesita a `frmOrder` para compilar, se agrega en T3. Para que no choquen nombres con los helpers privados de `frmQuote` que siguen ahí, T1 y T2 pueden ir en el mismo commit si hace falta. Pruebas, todas pasando en `test.ps1`:
  - `modTestUtils`: `Test_IsValidDate_AceptaFechas`, `Test_IsValidDate_RechazaTextoYVacio`, `Test_IsValidTime_AceptaHoras` («9:05», «09:05», «00:00», «23:59»), `Test_IsValidTime_RechazaFueraDeFormato` («24:00», «12:60», «1205», «12:5», «ab:cd», «»), `Test_IsValidVehicleYear_Limites` (con fecha de referencia fija: 1950 y ref+1 sí; 1949, ref+2, «95», «20a0» y «» no), `Test_IsValidMileage` («0» y «123456789» sí; «12,000», «1.5», «-5», «1234567890» y «» no), `Test_IsValidVin` (17 alfanuméricos, también en minúsculas, sí; 16, 18, con espacio interno y «» no).
  - `modTestData`: `Test_ListOrderStaff_DistintosOrdenados` (dos órdenes con «Ana», « ana », «Beto» en las dos columnas → «Ana», «Beto») y `Test_ListOrderStaff_SinOrdenesEmpty`.
- [ ] **T2 · Refactor de `frmQuote` a `modFormKit`** (R1, R2, R3): `frmQuote` llama `BeginFormKit` en `BuildLayout`, borra sus helpers privados y el `Enum`, pasa `Me` como `parent` y usa `IsValidDate` en `SyncQuoteFromForm` y `UpdateValidUntil`. El respaldo del logo se llama `c_HeaderLogo`. Verificación manual: build y comparación visual contra el `.xlsm` anterior, con logo y sin logo (nuevo, agregar, editar, eliminar, guardar, cargar, PDF).
- [ ] **T3 · `frmOrder` en el build, con estructura y selector** (R4, R5, R18, R22, R23): `src/forms/frmOrder.vba` con barra superior, franja de folio y estado, `MultiPage` de dos pestañas con todos los controles de Design y pie. `modApp.ShowOrderForm` y el botón del selector en los dos formularios, con `MSG_UNSAVED_SWITCH`. `lib.ps1` crea los dos formularios y `AddFormLogo(formName, ...)` pone el logo en los dos. Textos `UI_ORD_*`. Al final se puede ir y volver entre los dos formularios, y `frmOrder` muestra una orden nueva vacía (todavía sin guardar).
- [ ] **T4 · Captura y validación de la orden** (R6, R7, R9-R17, R21): combos de estado, gasolina y personal (`SelectOrAdd`, `LoadStaffLists`), autocompletado y `FillClientContact`, filtros de teclado, mayúsculas, aviso del VIN, `StartNewOrder` (fecha y hora actuales), «Nueva orden», `QueryClose` con aviso y `SyncOrderFromForm` con las reglas en orden y `WarnOnPage`.
- [ ] **T5 · Guardar y cargar** (R19, R20): `SaveCurrentOrder` (folio reasignado, sobrescritura confirmada una vez, recarga de historial, clientes y personal; `MSG_ORDER_SAVED`), `btnLoad_Click` y Enter en el buscador, `ShowOrder` sin marcar cambios y conservando `QuoteNumber`.
- [ ] **T6 · Documentación y verificación manual** (R2, R4, R22, R23): `README.md` (estructura con `modFormKit` y `frmOrder`; cómo se abre la orden desde el cotizador) y `SKILL.md` (los helpers viven en `modFormKit`). Prueba manual en Excel a 1366×768 y escala 100 %: el formulario cabe, se cambia de pestaña, guardar y cargar una orden con todos los campos, kilometraje o gasolina faltantes bloquean y cambian de pestaña, el VIN de 15 caracteres solo avisa, cliente existente llena el contacto vacío y no pisa lo escrito, selector con y sin cambios, colores de «Mi empresa» con un primario claro (texto en tinta) y con uno oscuro.
