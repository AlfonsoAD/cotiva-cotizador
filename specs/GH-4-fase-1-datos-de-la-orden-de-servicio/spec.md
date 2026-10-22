# GH-4 · Fase 1 · Datos de la orden de servicio (ServiceOrder, BD_Ordenes, folios OS-)

## Brief

**Objetivo:** dar al cotizador la capa de datos de la orden de servicio (clase `ServiceOrder`, hoja `BD_Ordenes`, folios `OS-0001`) y ligarla con la cotización y con el contacto del cliente, sin formulario.

**Entra**
- `ServiceOrder.cls` (26 propiedades) y `Quote.OrderNumber`.
- `modConfig`: `SHEET_ORDERS`, `ORDER_PREFIX`, base y relleno de folios, los seis estados y los niveles de combustible.
- `modData`: generador de folios genérico; `GetNextOrderNumber`, `SaveOrder`, `LoadOrder`, `ListOrders`, `OrderExists`, `GetClientContact`; columna I `OrdenFolio` en `BD_Cotizaciones`; `BD_Clientes` pasa de 2 a 6 columnas (teléfonos, correo, dirección).
- Hoja `BD_Ordenes` creada en `modApp.EnsureWorkbookStructure`, protegida y oculta en `modSecurity`.
- README: estructura y "Migrar datos".

**No entra**
- Formulario de la orden (`frmOrder`, #6) y validaciones de captura (rango de año, VIN de 17 caracteres, hora hh:mm).
- Datos fiscales del cliente y formulario del catálogo (#12).
- Sincronizar el vínculo en ambos sentidos: `SaveOrder` no escribe en `BD_Cotizaciones` (ver [N]).

**Decisiones nuevas**
- [N] `GetNextQuoteNumber()` conserva su firma y llama a un generador privado `NextFolio(hoja, prefijo, base, dígitos)`. Se agrega `GetNextOrderNumber()`. `frmQuote` no cambia. Descartado: cambiar la firma pública y tocar los dos llamadores de `frmQuote`.
- [N] Cada parte del vínculo orden–cotización se guarda por separado. Mantener ambos lados coherentes le toca al formulario (#6). Descartado: que `SaveOrder` reescriba la columna I de la cotización, porque acopla las dos funciones y puede pisar un vínculo anterior.
- [N] `SaveOrder` asigna el siguiente folio si la orden llega sin `Number`. Descartado: rechazar la orden y regresar False.
- [N] Fechas y horas vacías se guardan como 0 en la clase y como celda vacía en la hoja. `Year` = 0 significa "sin año".
- [N] Las celdas de teléfono, placas y VIN se escriben con formato de texto (`@`) para que Excel no las convierta en número (no se pierden ceros a la izquierda).
- [N] `GetClientContact` regresa el contacto en parámetros `ByRef` y un Boolean (el mismo patrón que `TryParseNumber`). Descartado: regresar un arreglo Variant.

**Riesgos**
- Libros viejos: las columnas nuevas y la hoja `BD_Ordenes` se crean al abrir el libro (`EnsureDataHeaders`). Una cotización vieja sin columna I carga con `OrderNumber` vacío.
- Si se olvida agregar `BD_Ordenes` a `modSecurity`, la hoja queda visible y sin protección.
- `CLIENT_COLUMNS` pasa a 6: `ClientExists` y `ListClients` leen más columnas, pero el resultado es el mismo.

**Tamaño:** M, ≈2 h. Una clase nueva, unas 250 líneas en `modData` y cambios de una línea en `modConfig`, `modApp`, `modSecurity` y `Quote`.

## Discovery

# Discovery GH-4 · Datos de la orden de servicio

## Alcance
- Solo capa de datos: `ServiceOrder.cls`, `modData`, `modConfig`, `Quote.OrderNumber`. El formulario `frmOrder` es la Fase 3 (#6) y queda fuera.
- Incluye lo necesario para que la hoja nueva exista y quede protegida: crear `BD_Ordenes` en `modApp.EnsureWorkbookStructure`/`AddSheet`, agregarla a las listas de `modSecurity` (`ProtectAll`, `AdminUnlock`, `HIDE_DATA_SHEETS`) y revisar `modBuild`.
- Actualizar los llamadores de `GetNextQuoteNumber` en `frmQuote` (`StartNewQuote`, `SaveCurrentQuote`) si cambia la firma.
- Actualizar el README (estructura y sección "Migrar datos": `BD_Ordenes` y las columnas nuevas).

## Folios
- Generalizar el generador de folios para recibir prefijo, base y relleno.
- Orden: `OS-0001` (prefijo `OS-`, base 0, relleno a 4 dígitos). `ORDER_PREFIX = "OS-"` en `modConfig`.
- Cotización: sin cambios, `COT-` con base 1000 y sin relleno (`COT-1001`). Los folios existentes no cambian de numeración.

## Estados
- Seis estados, en `modConfig`: Recibida, En diagnóstico, En reparación, Lista, Entregada, Cancelada.
- Una orden nueva nace en Recibida. Sin transiciones forzadas: se puede pasar a cualquier estado.

## Relación orden–cotización
- 1 a 0..1 en ambos sentidos: la orden guarda a lo más un `QuoteNumber` y la cotización a lo más un `OrdenFolio` (columna I de `BD_Cotizaciones`, propiedad `Quote.OrderNumber`).
- Ambas pueden existir solas. Las cotizaciones viejas sin columna I cargan con `OrderNumber` vacío.
- `SaveQuote`/`LoadQuote` y `QUOTE_COLUMNS` pasan de 8 a 9 columnas.

## Tipos
- `Year` Integer, `Mileage` Long (km), `FuelLevel` texto de la lista (R, ¼, ½, ¾, Lleno), `Vin` texto.
- Placas y VIN se guardan en mayúsculas. El resto de las validaciones (rango de año, VIN de 17, hora hh:mm) son del formulario (#6).

## Clientes (cambio respecto al plan)
- Cotizaciones y órdenes comparten `BD_Clientes`.
- `BD_Clientes` gana `Telefono1`, `Telefono2`, `Correo` y `Direccion` (hoy solo `Cliente | FechaAlta`), con `EnsureDataHeaders`.
- Al guardar una orden: si el cliente es nuevo se da de alta (como `AddClient` en la cotización); si existe, se actualizan sus datos de contacto con los capturados (solo los que vengan con valor).
- En lugar de `LastOrderForClient`, una función que lee el contacto del cliente desde `BD_Clientes` (p. ej. `GetClientContact`). El empate de nombre es el mismo de `ClientExists` (sin distinguir mayúsculas, con `Trim`).
- Si el cliente no tiene contacto en `BD_Clientes` pero sí órdenes previas, se toma de la orden con `InDate` más reciente (en empate, la última fila).
- La orden guarda además su propia copia del contacto, porque es un documento histórico.
- Los datos fiscales (RFC y CSF) y el formulario del catálogo quedan en la issue nueva #12.

## Errores y restricciones
- `modData` sigue siendo el único módulo que toca hojas `BD_*`; lee rangos completos a arreglos.
- Texto del usuario por `SafeCellText`; patrón de errores de la skill `vba-cotizador`.
- Libros de versiones anteriores sin `BD_Ordenes` ni las columnas nuevas deben abrir y migrarse solos (`EnsureDataHeaders`).
- Sin datos de un taller concreto en el código (regla de rama base genérica).
- Verificación manual: compilar, `build.ps1`, y probar guardar/cargar orden, cotización vieja y cliente que regresa.

## Requirements

**Folios**
- R1 [D] CUANDO se pide el siguiente folio de orden, el sistema DEBE regresar `OS-` seguido del número más alto de `BD_Ordenes` + 1, con relleno de ceros a 4 dígitos (`OS-0001` en una hoja vacía y `OS-10000` después de `OS-9999`).
- R2 [D] CUANDO se pide el siguiente folio de cotización, el sistema DEBE seguir regresando `COT-` + (máximo + 1), con base 1000 y sin relleno (`COT-1001` en una hoja vacía). Los folios existentes no cambian.
- R3 [N] CUANDO un folio guardado no tiene el prefijo esperado o sus dígitos no son válidos, el generador DEBE ignorarlo, igual que hoy con `QuoteNumberValue`.

**Orden de servicio**
- R4 [D] CUANDO se crea una `ServiceOrder`, `Status` DEBE valer `Recibida` y `InDate` la fecha de hoy. Los demás campos quedan vacíos o en 0.
- R5 [D] CUANDO se llama `SaveOrder`, el sistema DEBE reemplazar la fila anterior del mismo folio (si existía) por una fila nueva de 26 columnas en `BD_Ordenes`. Todo texto del usuario pasa por `SafeCellText`, y `Plates` y `Vin` se guardan con `Trim` y en mayúsculas. Regresa True si todo salió bien.
- R6 [N] CUANDO `SaveOrder` recibe una orden sin `Number`, el sistema DEBE asignarle `GetNextOrderNumber()` antes de guardar.
- R7 [D] CUANDO falla `SaveOrder`, el sistema DEBE mostrar el error con `ShowError "SaveOrder <folio>"`, restaurar `ScreenUpdating` y regresar False.
- R8 [D] CUANDO se llama `LoadOrder` con un folio que existe, el sistema DEBE regresar la orden con todos sus campos. Una fecha u hora que no es fecha queda en 0, un año o kilometraje que no es número válido queda en 0, y un estado vacío queda en `Recibida`. CUANDO el folio no existe, regresa Nothing.
- R9 [D] `OrderExists` DEBE regresar True solo si el folio está en `BD_Ordenes`, sin distinguir mayúsculas.
- R10 [N] `ListOrders` DEBE regresar una tabla (Folio, Cliente, Placas, Estado) de la más reciente a la más antigua (de la última fila hacia arriba), o Empty si no hay órdenes.

**Cotización**
- R11 [D] CUANDO se guarda una cotización, el sistema DEBE escribir `Quote.OrderNumber` en la columna I (`OrdenFolio`) de `BD_Cotizaciones`.
- R12 [D] CUANDO se carga una cotización sin valor en la columna I (incluidas las de versiones anteriores), `OrderNumber` DEBE quedar vacío y el resto de la carga debe funcionar igual que hoy.
- R13 [N] CUANDO se vuelve a guardar una cotización cargada, el sistema DEBE conservar su `OrdenFolio`.

**Clientes**
- R14 [D] CUANDO se guarda una orden con un cliente que no está en `BD_Clientes`, el sistema DEBE darlo de alta con su fecha y su contacto (Telefono1, Telefono2, Correo, Direccion).
- R15 [D] CUANDO se guarda una orden con un cliente que ya existe, el sistema DEBE sobrescribir solo los datos de contacto que vengan con valor y dejar los demás como estaban.
- R16 [D] CUANDO se pide `GetClientContact(cliente)`, el sistema DEBE buscar el cliente sin distinguir mayúsculas y con `Trim`, y regresar su contacto de `BD_Clientes`. SI los cuatro campos del catálogo están vacíos, DEBE tomar el contacto de la orden de ese cliente con `InDate` más reciente (en empate, la última fila). Regresa False si no encuentra ningún dato.
- R17 [D] La orden DEBE guardar su propia copia del contacto, aunque después cambie el catálogo.

**Estructura y migración**
- R18 [D] CUANDO se abre un libro sin `BD_Ordenes`, el sistema DEBE crear la hoja y escribir sus 26 encabezados. CUANDO faltan la columna I de `BD_Cotizaciones` o las columnas C:F de `BD_Clientes`, DEBE agregar esos encabezados sin tocar los datos.
- R19 [D] `ProtectAll` DEBE proteger `BD_Ordenes` y ocultarla (VeryHidden si `HIDE_DATA_SHEETS`), y `AdminUnlock` DEBE mostrarla y desprotegerla.
- R20 [D] `modData` DEBE seguir siendo el único módulo que lee o escribe hojas `BD_*`, y las lecturas DEBEN traer el rango completo a un arreglo. El código no lleva datos de ningún taller concreto.

## Design

### `modConfig` (constantes nuevas)
```vba
Public Const SHEET_ORDERS As String = "BD_Ordenes"
Public Const QUOTE_NUMBER_DIGITS As Long = 0            ' Sin relleno: COT-1001
Public Const ORDER_PREFIX As String = "OS-"
Public Const ORDER_BASE_NUMBER As Long = 0              ' El primer folio será OS-0001
Public Const ORDER_NUMBER_DIGITS As Long = 4
' --- Estados de la orden (sin transiciones forzadas) ---
Public Const ORDER_STATUS_RECEIVED As String = "Recibida"
Public Const ORDER_STATUS_DIAGNOSIS As String = "En diagnóstico"
Public Const ORDER_STATUS_REPAIR As String = "En reparación"
Public Const ORDER_STATUS_READY As String = "Lista"
Public Const ORDER_STATUS_DELIVERED As String = "Entregada"
Public Const ORDER_STATUS_CANCELLED As String = "Cancelada"
Public Const ORDER_STATUSES As String = "Recibida|En diagnóstico|En reparación|Lista|Entregada|Cancelada"
Public Const FUEL_LEVELS As String = "R|¼|½|¾|Lleno"     ' ¼ ½ ¾ existen en Windows-1252
```
Las listas son cadenas con `|` porque VBA no admite constantes de arreglo. El formulario (#6) hace `Split(ORDER_STATUSES, "|")`.

### `ServiceOrder.cls`
Se basa en `Quote.cls`: cabecera, `Option Explicit` y campos `Public`. No toca hojas.

| Propiedad | Tipo | Columna `BD_Ordenes` |
|---|---|---|
| Number | String | A Folio |
| Status | String | B Estado |
| ClientName, Phone1, Phone2, Email, Address | String | C Cliente, D Telefono1, E Telefono2, F Correo, G Direccion |
| Make, Model, Engine | String | H Marca, I Modelo, J Motor |
| Year | Integer | K Anio |
| Color, Plates | String | L Color, M Placas |
| Mileage | Long (km) | N Kilometraje |
| FuelLevel | String (de `FUEL_LEVELS`) | O Combustible |
| Vin | String | P VIN |
| InDate, InTime | Date | Q FechaEntrada, R HoraEntrada |
| ReceivedBy | String | S RecibioPor |
| PromisedDate, OutDate, OutTime | Date | T FechaPromesa, U FechaSalida, V HoraSalida |
| DeliveredBy | String | W EntregoPor |
| Requested, Diagnosis | String | X Solicitado, Y Diagnostico |
| QuoteNumber | String | Z CotizacionFolio |

`Class_Initialize`: `Status = ORDER_STATUS_RECEIVED`, `InDate = Date`. Fecha u hora en 0 significa "sin valor".

### `Quote.cls`
`Public OrderNumber As String  ' Folio de la orden ligada, ej. "OS-0001" (puede ir vacío)`.

### `modData`
Constantes privadas: `QUOTE_COLUMNS = 9`, `CLIENT_COLUMNS = 6`, `ORDER_COLUMNS = 26`. Se actualiza la cabecera de estructura del módulo.

Firmas públicas nuevas o que cambian:
```vba
Public Function GetNextQuoteNumber() As String          ' misma firma; delega en NextFolio
Public Function GetNextOrderNumber() As String
Public Function OrderExists(ByVal orderNumber As String) As Boolean
Public Function SaveOrder(ByVal o As ServiceOrder) As Boolean
Public Function LoadOrder(ByVal orderNumber As String) As ServiceOrder
Public Function ListOrders() As Variant                 ' (0..n-1, 0..3): Folio, Cliente, Placas, Estado
Public Function GetClientContact(ByVal clientText As String, ByRef phone1 As String, _
    ByRef phone2 As String, ByRef email As String, ByRef address As String) As Boolean
```
Funciones privadas:
```vba
Private Function NextFolio(ByVal ws As Worksheet, ByVal prefix As String, _
    ByVal baseNumber As Long, ByVal digits As Long) As String
Private Function FolioValue(ByVal folio As String, ByVal prefix As String) As Long  ' sustituye a QuoteNumberValue
Private Function FindOrderRow(ByVal orderNumber As String) As Long
Private Sub UpsertClientContact(ByVal clientText As String, ByVal phone1 As String, _
    ByVal phone2 As String, ByVal email As String, ByVal address As String)
Private Function CellDate(ByVal v As Variant) As Date     ' IsDate -> CDate; si no, 0
Private Function DateOrBlank(ByVal d As Date) As Variant  ' 0 -> Empty
```
- `NextFolio`: `highest = baseNumber` y se recorre la columna A con `FolioValue`. Con `digits > 0` regresa `prefix & Format$(highest + 1, String$(digits, "0"))`; si no, `prefix & (highest + 1)`.
- `SaveOrder`: sigue el patrón de `SaveQuote` (`On Error GoTo Fail`, `DeleteRowsByNumber`, una sola escritura `Resize(1, ORDER_COLUMNS)`). Antes de escribir pone `NumberFormat = "@"` en D:E, M y P. Fechas Q, T y U con `dd/mm/yyyy`, horas R y V con `hh:mm`. Después llama `UpsertClientContact` y guarda el libro si `SAVE_WORKBOOK_ON_SAVE`. No toca `BD_Cotizaciones`.
- `LoadOrder`: lee la fila con un solo `Range.Value` a arreglo. Año: `TryParseNumber`, y si queda entre 0 y 9999 se usa `CInt`, si no 0. Kilometraje: `TryParseNumber`, y si queda entre 0 y 2147483647 se usa `CLng`, si no 0.
- `UpsertClientContact`: sale si el nombre queda vacío. Llama `AddClient` (alta con fecha) y busca la fila con el mismo empate que `ClientExists`. Escribe con `SafeCellText`, en C:F, solo los valores con `Len(Trim$) > 0`. Teléfonos con formato `@`.
- `GetClientContact`: primero busca en `ReadData(BD_Clientes, 6)`. Si los cuatro campos del catálogo están vacíos, recorre `ReadData(BD_Ordenes, ORDER_COLUMNS)` y se queda con la fila de ese cliente con mayor `CellDate(col Q)`; con `>=`, en empate gana la última fila.
- `SaveQuote` agrega `q.OrderNumber` (con `UCase$(Trim$)`) como noveno valor. `LoadQuote` lee `CStr(ws.Cells(r, 9).Value)` con `Trim$`.
- `EnsureDataHeaders`: agrega `"OrdenFolio"` a `BD_Cotizaciones` y `"Telefono1","Telefono2","Correo","Direccion"` a `BD_Clientes`. `If SheetExists(SHEET_ORDERS)` escribe los 26 encabezados de la tabla.

### `modApp` y `modSecurity`
- `EnsureWorkbookStructure`: `If Not SheetExists(SHEET_ORDERS) Then AddSheet SHEET_ORDERS` después de `SHEET_ITEMS`. Ambas ramas ya llaman `EnsureDataHeaders`. `modBuild` no cambia, porque usa `EnsureWorkbookStructure`.
- `ProtectAll`: `Array(SHEET_QUOTES, SHEET_ITEMS, SHEET_CLIENTS, SHEET_ORDERS)`. `AdminUnlock`: se agrega `SHEET_ORDERS` a su arreglo.
- `build.ps1` ya importa todos los `.cls` de `src/classes`, así que no cambia.

### Seguridad
- Sin authz nueva. Las hojas BD_ siguen protegidas con `APP_PASSWORD` y `UserInterfaceOnly`.
- Validación: el texto libre pasa por `SafeCellText` (inyección de fórmulas). Los folios se normalizan con `UCase$(Trim$)`. Lo que se lee de la hoja se valida al cargar, porque alguien pudo editarla a mano.
- Datos sensibles: teléfonos, correo, dirección y VIN quedan en hojas ocultas y protegidas del mismo libro. No salen del libro ni se registran en otro lado.
- Errores: patrón `Fail`/`Cleanup`/`ShowError` en las funciones que escriben.

### Descartado
- `LastOrderForClient` del tracker: se reemplaza por `GetClientContact` sobre `BD_Clientes`, con la orden más reciente como respaldo (discovery).
- Validar en la capa de datos que `Status` y `FuelLevel` pertenezcan a sus listas: eso le toca al formulario (#6).

## Tasks

No hay pruebas automáticas. Cada tarea se verifica con Depurar › Compilar, `build/build.ps1` y las pruebas manuales P1..P8.

- [x] **T1 · Contrato** (R4, R11, R20): constantes de `modConfig`; `ServiceOrder.cls` completa; `Quote.OrderNumber`; en `modData`, las firmas públicas nuevas como esqueletos que compilan (regresan False, Nothing o Empty) y las constantes `ORDER_COLUMNS`, `QUOTE_COLUMNS`, `CLIENT_COLUMNS`. Pruebas manuales definidas:
  - P1 `FolioOrdenVacio`: libro nuevo, `GetNextOrderNumber` = `OS-0001`.
  - P2 `FolioCotizacionIntacto`: un libro con `COT-1007` sigue dando `COT-1008`.
  - P3 `GuardarCargarOrden`: guardar una orden con los 26 campos y cargarla igual (placas y VIN en mayúsculas, teléfono con 0 inicial).
  - P4 `ReemplazarOrden`: guardar dos veces el mismo folio deja una sola fila.
  - P5 `CotizacionVieja`: una cotización sin columna I carga y se vuelve a guardar sin errores.
  - P6 `CotizacionConOrden`: `OrderNumber` se guarda en I y se conserva al guardar otra vez.
  - P7 `ClienteRegresa`: el contacto sale de `BD_Clientes`; si el catálogo está vacío, sale de la orden más reciente; un campo vacío no borra el dato guardado.
  - P8 `LibroAnterior`: un libro v8.0.0 abre, crea `BD_Ordenes`, agrega los encabezados I y C:F, y la hoja queda VeryHidden y protegida.
- [x] **T2 · Folios genéricos** (R1, R2, R3): `NextFolio` y `FolioValue`; `GetNextQuoteNumber` delega en ellas; `GetNextOrderNumber`. Verifica P1 y P2.
- [x] **T3 · Estructura y protección** (R18, R19): `EnsureDataHeaders` (I, C:F y los 26 de `BD_Ordenes`); `modApp.EnsureWorkbookStructure`; las listas de `modSecurity`. Verifica P8.
- [x] **T4 · Columna OrdenFolio** (R11, R12, R13): `SaveQuote` y `LoadQuote` con 9 columnas. Verifica P5 y P6.
- [x] **T5 · CRUD de órdenes** (R5, R6, R7, R8, R9, R10, R17): `FindOrderRow`, `OrderExists`, `SaveOrder` (sin llamar todavía a `UpsertClientContact`), `LoadOrder`, `ListOrders`, `CellDate`, `DateOrBlank`. Verifica P3 y P4.
- [x] **T6 · Contacto del cliente** (R14, R15, R16): `UpsertClientContact` (se llama desde `SaveOrder`) y `GetClientContact`. Verifica P7.
- [x] **T7 · Documentación** (R18): README, con `ServiceOrder` en classes, `BD_Ordenes` y las columnas nuevas en "Migrar datos"; cabecera de `modData`. Repaso completo P1..P8.
