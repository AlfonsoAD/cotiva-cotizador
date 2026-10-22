# GH-14 · Pruebas automáticas de VBA con Excel (test.ps1 + paso test de bflow)

## Brief

**Objetivo:** que `bflow check` compile y pruebe el VBA dentro de Excel con `build/test.ps1` y falle si una prueba sale mal.

**Entra**
- `src/tests/`: asserts propios, runner `RunAllTests` y la primera tanda de pruebas: P1..P8 de GH-4, funciones puras de `modUtils` y `TryParseHexColor`.
- `build/lib.ps1` (armado común), `build.ps1` reescrito sobre él con el mismo resultado, y `build/test.ps1`: Excel oculto, límite de 5 min y `dist/test-results.txt`.
- Modo prueba en `modUtils.ShowError`. Paso `test` en `bflow.yaml` para todos los carriles. README y CLAUDE.md.

**No entra:** formularios, PDF, apariencia, CI. Tampoco los `MsgBox` que no pasan por `ShowError` (contraseña, diseño, PDF): las pruebas no llegan a ellos.

**Decisiones nuevas**
- [N1] El libro de pruebas no se guarda en disco: queda como `Libro1` en memoria. Así `SAVE_WORKBOOK_ON_SAVE` no lo guarda después de cada `SaveOrder` y no hay archivo temporal que borrar. Descartado: guardarlo en `%TEMP%` y borrarlo al final.
- [N2] `test.ps1` lanza un proceso hijo (él mismo con `-Worker`) que maneja Excel. El padre espera con límite de tiempo y mata solo al hijo y al Excel cuyo PID dejó el hijo. Descartado: un solo proceso, porque `Application.Run` bloquea el hilo y no se puede cortar.
- [N3] Si falta `modSecrets.bas`, se importa `modSecrets.example.bas` directo (ya declara `VB_Name = "modSecrets"`), sin crear archivo temporal. Descartado: copiarlo a `%TEMP%`.
- [N4] Además del runner, los helpers de `src/tests/` pueden escribir en `BD_*` para armar filas de versiones anteriores (P5, P8), que `modData` ya no sabe escribir. Descartado: agregar a `modData` código solo para pruebas.
- [N5] Si no encuentra ninguna `Test_*`, sale con 1. Descartado: contarlo como éxito.
- [N6] El paso no usa `needs` ni `optional` de bflow: no debe saltarse en silencio (discovery D1). Descartado: `optional: true`.
- [N7] Una falla de assert se anota y la prueba sigue; un error inesperado corta esa prueba y cuenta como falla. Descartado: abortar con `Err.Raise` en el primer assert.

**Riesgos**
- Que un error dentro de `Application.Run` no regrese al runner y abra el diálogo de depuración. El límite de tiempo lo convierte en "tiempo agotado". T3 lo comprueba con una prueba que truena a propósito.
- Mientras corre, `AccessVBOM` queda encendido (unos 1-2 min). Se restaura en el `finally` del padre aunque mate al hijo.
- `TryParseNumber` y `FormatMoney` dependen de la configuración regional. Las pruebas usan `DecimalSeparator()` o entradas enteras.
- Agrega 1-2 min a cada `bflow check` y exige Excel en la máquina que lo corre.

**Tamaño:** M. 5 tareas: 3 archivos `.bas` de prueba más el runner, 2 scripts nuevos, 1 script refactorizado y 3 líneas en `modUtils`.

## Discovery

﻿# Discovery GH-14 — Pruebas automáticas de VBA con Excel

## Alcance
- Pruebas en VBA sin dependencias: módulos `src/tests/modTest*.bas` con `Public Sub Test_*`, más un módulo propio de asserts (`AssertEqual`, `AssertTrue`, `AssertFails`). Sin Rubberduck.
- Los módulos de `src/tests/` solo se importan al libro de pruebas. Nunca entran al `.xlsm` que se entrega (`dist/Cotiva-Cotizador.xlsm`).
- `RunAllTests` encuentra los `Test_*` leyendo el `VBProject` y limpia las hojas `BD_*` antes de cada prueba.
- Primera tanda de pruebas: P1..P8 de GH-4 (folios `OS-`/`COT-`, guardar y cargar una orden, reemplazo, cotización vieja, `OrdenFolio`, contacto del cliente, migración de un libro anterior) y las funciones puras de `modUtils` (`AmountToWords`, `Round2`, `TryParseNumber`, `IsValidRfc`, `SanitizeFileName`, etc.) y `modTheme.TryParseHexColor`.
- Para P8, el runner arma por código un libro con la estructura anterior. No se sube ningún `.xlsm` como fixture.
- Fuera de alcance: formularios, PDF y apariencia (siguen manuales), y CI en GitHub Actions (los runners no traen Office).

## Decisiones
1. El paso `test` va en `check.steps` de `bflow.yaml`, después de `lint`, y corre en todos los carriles (full, light, hotfix). Si la máquina no tiene Excel, `test.ps1` falla con un mensaje claro; no se salta en silencio.
2. Limpieza de `BD_*`: el runner de `src/tests/` limpia las hojas directamente. Queda exento de la regla "solo modData toca BD_*" porque nunca viaja al libro entregado. No se agrega a `modData` código solo para pruebas. Las pruebas en sí ejercitan `modData` para leer y escribir.
3. La lógica de armado (importar módulos, clases y formulario; `AccessVBOM` con restauración en `finally`; cierre de Excel) se mueve a un script común `build/lib.ps1`, que usan `build.ps1` y `test.ps1`. `build.ps1` debe seguir produciendo exactamente lo mismo.
4. `build/test.ps1`: arma un libro temporal con Excel oculto, importa `src/tests/` y, si falta `src/modules/modSecrets.bas`, usa uno temporal a partir de `modSecrets.example.bas` sin dejarlo en `src/`. Llama por COM a cada prueba (`ListTests`, `BeginTest`, `EndTest`, `Report`), escribe `dist/test-results.txt` (`ok N / falla M` y el detalle de cada falla), también lo muestra en consola y sale con 1 si algo falla.
5. Límite de 5 minutos: si se pasa (por ejemplo, porque un error de compilación abrió un diálogo modal), mata el proceso de Excel que lanzó y lo reporta como "tiempo agotado / posible error de compilación". Siempre cierra Excel y restaura `AccessVBOM`.
6. Modo prueba en `modUtils`: una bandera hace que `ShowError` guarde el último error (dónde, número, descripción) en una variable que las pruebas y `AssertFails` pueden leer, sin abrir `MsgBox`. Fuera de modo prueba, el comportamiento no cambia.

## Restricciones
- Archivos de `src/` en Windows-1252 con CRLF, `Option Explicit` y las reglas de `build/lint.ps1` (`src/tests/` incluido).
- No se edita ni se sube ningún `.xlsm`. `dist/` sigue ignorado.
- Solo hay que matar la instancia de Excel que lanzó el script, nunca otros Excel abiertos por el usuario.

## Requirements

- **R1** [D] CUANDO corre el runner (`RunAllTests` desde VBA, o `ListTests` y `BeginTest`/`EndTest` por prueba desde `test.ps1`), el sistema DEBE ejecutar, en orden de módulo y de línea, cada `Public Sub Test_*()` sin parámetros de los módulos estándar cuyo nombre empiece con `modTest`, leyéndolos del `VBProject`.
- **R2** [D] CUANDO un `AssertEqual`, `AssertTrue` o `AssertFails` no se cumple, el sistema DEBE anotar la falla con `<módulo>.<prueba>`, la etiqueta y lo esperado contra lo obtenido, y seguir con la prueba.
- **R3** [N] CUANDO una prueba lanza un error que nadie maneja, el sistema DEBE contarla como falla con el número y la descripción del error, y seguir con la siguiente prueba.
- **R4** [D] ANTES de cada prueba, el sistema DEBE dejar las hojas `BD_*` solo con encabezados, recrear la estructura que falte (`EnsureWorkbookStructure`) y volver a proteger el libro (`ProtectAll`), como un libro recién abierto.
- **R5** [D] CUANDO corre `build.ps1`, el sistema NO DEBE importar nada de `src/tests/`. El `.xlsm` entregado no contiene `modTest*`, `modAssert` ni `modTestRunner`.
- **R6** [D] CUANDO corre `test.ps1`, el sistema DEBE armar con Excel oculto un libro con el mismo código que `build.ps1` (y `BuildWorkbook`) más `src/tests/*.bas`, y correr cada prueba en él por COM (`ListTests`, `BeginTest`, `EndTest`, `Report`).
- **R7** [D] CUANDO falta `src/modules/modSecrets.bas`, `test.ps1` DEBE usar `modSecrets.example.bas` sin crear ni dejar archivos en `src/`.
- **R8** [D] CUANDO terminan las pruebas, `test.ps1` DEBE escribir `dist/test-results.txt` y mostrar en consola la línea `ok N / falla M` seguida de una línea por falla. DEBE salir con 0 si M = 0 y N > 0, y con 1 en cualquier otro caso.
- **R9** [N] CUANDO `RunAllTests` no encuentra ninguna prueba, `test.ps1` DEBE salir con 1 y decir "no se encontraron pruebas".
- **R10** [D] CUANDO pasan 300 s (configurable con `-TimeoutSeconds`) sin terminar, `test.ps1` DEBE matar el proceso hijo y la instancia de Excel que lanzó (por PID), nunca otra. DEBE escribir y mostrar "tiempo agotado (N s): posible error de compilación; corre build.ps1 para verlo" y salir con 1.
- **R11** [D] SIEMPRE, también con error o tiempo agotado, `test.ps1` DEBE cerrar Excel sin guardar y dejar `AccessVBOM` como estaba.
- **R12** [D] CUANDO la máquina no tiene Excel de escritorio, `test.ps1` DEBE salir con 1 y decir que las pruebas de VBA necesitan Excel. No se salta en silencio.
- **R13** [D] MIENTRAS el modo prueba está activo, `ShowError` DEBE guardar dónde ocurrió el error, su número y su descripción, sin mostrar `MsgBox`. Fuera de modo prueba su comportamiento no cambia.
- **R14** [D] `build.ps1` y `test.ps1` DEBEN armar el libro con las mismas funciones de `build/lib.ps1`. `build.ps1` DEBE seguir generando el mismo `.xlsm`: mismos componentes, hojas y protección, y los mismos parámetros `-Output` y `-Hidden`.
- **R15** [D] `bflow check` DEBE correr el paso `test` (`build/test.ps1`) después de `lint`, en todos los carriles y sin marcarlo como opcional.
- **R16** [D] Las pruebas P1..P8 de GH-4 DEBEN quedar automatizadas con los resultados que esa spec define.
- **R17** [D] DEBE haber pruebas de las funciones puras `AmountToWords`, `Round2`, `TryParseNumber`, `IsValidRfc`, `SanitizeFileName`, `SafeCellText`, `TaxRateOf`/`TaxOptionOf` y `modTheme.TryParseHexColor`.
- **R18** [D] Para P8, la estructura de un libro anterior DEBE armarse por código en el libro de pruebas. No se sube ningún `.xlsm` como fixture (archivo de datos fijo para una prueba).
- **R19** [D] Los archivos de `src/tests/` DEBEN cumplir `build/lint.ps1`: Windows-1252, CRLF y `Option Explicit`.
- **R20** [N] README y CLAUDE.md DEBEN explicar cómo correr `test.ps1`, cómo agregar una prueba y qué sigue siendo manual.

## Design

### Archivos

| Archivo | Nuevo/cambia | Qué hace |
|---|---|---|
| `src/modules/modUtils.bas` | cambia | Modo prueba y último error (R13) |
| `src/tests/modAssert.bas` | nuevo | Asserts; anota fallas en el runner (R2) |
| `src/tests/modTestRunner.bas` | nuevo | `RunAllTests`, búsqueda de `Test_*`, limpieza de `BD_*` (R1, R3, R4, R9) |
| `src/tests/modTestHelpers.bas` | nuevo | Datos de ejemplo y estructura de versiones anteriores (R16, R18, [N4]) |
| `src/tests/modTestUtils.bas` | nuevo | Pruebas de funciones puras (R17) |
| `src/tests/modTestData.bas` | nuevo | P1..P8 de GH-4 (R16) |
| `build/lib.ps1` | nuevo | Armado común (R14) |
| `build/build.ps1` | cambia | Usa `lib.ps1`; mismo resultado (R5, R14) |
| `build/test.ps1` | nuevo | Padre con límite de tiempo e hijo `-Worker` que corre Excel (R6-R12) |
| `bflow.yaml` | cambia | Paso `test` (R15) |
| `README.md`, `CLAUDE.md` | cambian | Sección "Pruebas automáticas"; mapa y comandos (R20) |

### `modUtils`: modo prueba (R13)

```vba
Private mTestMode As Boolean
Private mLastErrorWhere As String, mLastErrorNumber As Long, mLastErrorDescription As String

Public Sub SetTestMode(ByVal enabled As Boolean)
Public Function IsTestMode() As Boolean
Public Sub ClearLastError()                        ' Vacía los tres valores
Public Function LastErrorWhere() As String
Public Function LastErrorNumber() As Long          ' 0 = no hubo error desde ClearLastError
Public Function LastErrorDescription() As String
```

`ShowError` lee `Err` primero, como hoy. Si `mTestMode` está activo, guarda los tres valores, deja `ScreenUpdating = True` y sale sin `MsgBox`. Si no, hace exactamente lo de hoy. El código de producción no llama `SetTestMode`, y una variable `Private` empieza en `False`: el libro entregado nunca entra en modo prueba.

### `modAssert` (R2)

```vba
Public Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal label As String)
Public Sub AssertTrue(ByVal condition As Boolean, Optional ByVal label As String)
Public Sub AssertFails(ByVal macroName As String, Optional ByVal label As String, _
                       Optional ByVal arg1 As Variant, Optional ByVal arg2 As Variant, Optional ByVal arg3 As Variant)
```

- `AssertEqual` falla si `VarType` es distinto (los numéricos se comparan entre sí: `Long` contra `Double` vale) o si los valores no son iguales. El mensaje es `<label>: esperado <expected>, obtenido <actual>`, con los textos entre comillas.
- `AssertFails` llama `ClearLastError` y corre `Application.Run "'<libro>'!" & macroName` con los argumentos que traiga (`IsMissing` decide con cuántos, de 0 a 3). Pasa si `Application.Run` lanza un error o si después `LastErrorNumber <> 0`. Al terminar vuelve a llamar `ClearLastError`.
- Las fallas se anotan con `modTestRunner.RecordFailure message`. Nunca hacen `Err.Raise` ([N7]).

### `modTestRunner` (R1, R3, R4, R9)

```vba
Public Function RunAllTests() As String            ' Reporte: 1a línea "ok N / falla M", luego "FALLA <mod>.<prueba>: <detalle>"
Public Sub RecordFailure(ByVal message As String)  ' La usa modAssert
Private Function FindTests() As Collection         ' "<módulo>.<prueba>", en orden
Private Sub ResetDataSheets()
```

- `FindTests` recorre `ThisWorkbook.VBProject.VBComponents` con enlace tardío, sin referencia a VBIDE (`Type = 1` es módulo estándar). Toma los módulos con `Name Like "modTest*"` y, en su `CodeModule`, las líneas que cumplen `Trim$(line) Like "Public Sub Test_*()*"`. El nombre es lo que va entre `Sub ` y `(`.
- Por cada prueba: `SetTestMode True`, `ResetDataSheets`, `ClearLastError`, falla actual = vacía. Después, `On Error GoTo` hacia un manejador local, `Application.Run "'" & ThisWorkbook.Name & "'!" & nombre` y `On Error GoTo 0`. Con error, anota `error <n>: <descripción>` (R3). La prueba pasa si no anotó nada. Al final restablece `Application.ScreenUpdating = True` y `EnableEvents`.
- `ResetDataSheets`: `TryUnprotectWorkbook`. Para cada hoja de `SHEET_QUOTES`, `SHEET_ITEMS`, `SHEET_CLIENTS` y `SHEET_ORDERS` que exista: `UnlockSheet` y borra las filas 2 en adelante (`ws.Rows("2:" & ws.Rows.Count).Delete`). Después `EnsureWorkbookStructure` (recrea lo que una prueba como P8 haya quitado) y `ProtectAll`. Esta escritura directa en `BD_*` es la excepción del discovery (D2): `src/tests/` nunca viaja al libro entregado.
- Sin pruebas: devuelve `ok 0 / falla 0` y la línea `no se encontraron pruebas` (R9).

### `modTestHelpers` (R16, R18, [N4])

```vba
Public Function NewSampleOrder(ByVal orderNumber As String) As ServiceOrder   ' Los 26 campos llenos; placas/VIN en minúsculas; Phone1 = "0123456789"; InTime = TimeSerial(0,0,0)
Public Function NewSampleQuote(ByVal quoteNumber As String) As Quote          ' Cliente, 16 % y un concepto
Public Sub WriteLegacyQuoteRow(ByVal quoteNumber As String)                  ' Fila A:H en BD_Cotizaciones (sin I) y su fila en BD_Detalle
Public Sub ClearClientContacts()                                             ' Vacía C:F de BD_Clientes (catálogo sin contacto)
Public Sub MakeLegacyWorkbook()                                              ' v8.0.0: borra BD_Ordenes, vacía I1 de BD_Cotizaciones y C1:F1 de BD_Clientes
```

Los helpers que escriben llaman `TryUnprotectWorkbook` y `UnlockSheet`. `MakeLegacyWorkbook` borra la hoja con `Application.DisplayAlerts = False` y lo restaura después. No se usa un segundo libro: `modData` trabaja sobre `ThisWorkbook` (R18).

### Pruebas (nombres del contrato)

`modTestUtils` (R17, R13):
- `Test_AmountToWords_MilCuarentaYCuatro`: 1044 → `(MIL CUARENTA Y CUATRO PESOS 00/100 M.N.)`
- `Test_AmountToWords_UnPeso`: 1 → `(UN PESO 00/100 M.N.)`
- `Test_AmountToWords_SoloCentavos`: 0.5 → `(CERO PESOS 50/100 M.N.)`
- `Test_AmountToWords_UnMillon`: 1000000 → `(UN MILLÓN DE PESOS 00/100 M.N.)`
- `Test_AmountToWords_VeintiunMil`: 21000 → `(VEINTIÚN MIL PESOS 00/100 M.N.)`
- `Test_Round2_RedondeaComoExcel`: 0.125 → 0.13; -0.125 → -0.13; 2 → 2
- `Test_TryParseNumber_Validos`: `"12"` → 12; `" 7 "` → 7; `"1" & DecimalSeparator() & "5"` → 1.5
- `Test_TryParseNumber_Invalidos`: `""` y `"abc"` → False
- `Test_IsValidRfc`: `ABC010203XY1` y `ABCD010203XY1` → True; `abcd010203xy1` → True; `AB010203` → False
- `Test_SanitizeFileName`: `a/b:c*?` → `a_b_c__`; `"  x   y  "` → `x y`; 200 caracteres → 120
- `Test_SafeCellText`: `=1+1` → `'=1+1`; `-5` → `'-5`; `hola` → `hola`; `""` → `""`
- `Test_TaxRateOf_TaxOptionOf`: taxStandard/taxBorder/taxExempt ↔ 0.16/0.08/0, de ida y de vuelta
- `Test_TryParseHexColor`: `#CF4A06` y `cf4a06` → True y `RGB(207, 74, 6)`; `#12345` y `GG0000` → False
- `Test_ShowError_ModoPruebaRegistra`: `Err.Raise 5` dentro de un manejador y `ShowError "X"` → `LastErrorWhere = "X"`, `LastErrorNumber = 5`, sin `MsgBox`

`modTestData` (R16, R18):
- `Test_P1_FolioOrdenVacio`: `GetNextOrderNumber` = `OS-0001`
- `Test_P2_FolioCotizacionIntacto`: vacío → `COT-1001`; después de `SaveQuote` de `COT-1007` → `COT-1008`
- `Test_P3_GuardarCargarOrden`: `SaveOrder(NewSampleOrder("OS-0001"))` y `LoadOrder` regresa los 26 campos iguales, con placas y VIN en mayúsculas, `Phone1 = "0123456789"` y `InTime` 00:00 conservado
- `Test_P4_ReemplazarOrden`: dos `SaveOrder` del mismo folio, con `Status` distinto, dejan `ListOrders` con 1 fila y el segundo `Status`
- `Test_P5_CotizacionVieja`: `WriteLegacyQuoteRow "COT-1001"`; `LoadQuote` no es `Nothing` y `OrderNumber = ""`; `SaveQuote` regresa True y `LastErrorNumber = 0`
- `Test_P6_CotizacionConOrden`: `OrderNumber = "os-0001"` se guarda y se carga como `OS-0001`; un segundo `SaveQuote` lo conserva
- `Test_P7_ContactoDelCatalogo`: después de `SaveOrder`, `GetClientContact` regresa los datos de la orden desde `BD_Clientes`
- `Test_P7_ContactoDeOrdenReciente`: dos órdenes del mismo cliente con `InDate` distinta, `ClearClientContacts`, y `GetClientContact` regresa los de la más reciente
- `Test_P7_CampoVacioNoBorra`: una segunda orden con `Email = ""` no borra el correo guardado en el catálogo
- `Test_P8_LibroAnterior`: `MakeLegacyWorkbook` y luego `PrepareWorkbook` = True; existe `BD_Ordenes` con `Visible = xlSheetVeryHidden` y `ProtectContents`; `I1 = "OrdenFolio"`; `C1:F1 = Telefono1, Telefono2, Correo, Direccion`; `ThisWorkbook.ProtectStructure`
- `Test_SaveOrder_NothingFalla`: `AssertFails "modData.SaveOrder", "SaveOrder Nothing", Nothing`

### `build/lib.ps1` (R14)

Se carga con dot-source (`. (Join-Path $PSScriptRoot 'lib.ps1')`), así que sus funciones quedan disponibles en el script que lo carga:

```powershell
function Enable-VbaAccess            # -> valor anterior de AccessVBOM ($null si no existía)
function Restore-VbaAccess { param($Previous) }
function New-ExcelApp { param([bool]$Visible) }             # DisplayAlerts=$false, EnableEvents=$false
function Get-ExcelProcessId { param($Excel) }               # PID desde $Excel.Hwnd (user32!GetWindowThreadProcessId vía Add-Type)
function Import-CotizadorCode { param($Workbook, [string]$Src, [string]$SecretsFile, [switch]$IncludeTests) }
function Close-ExcelApp { param($Excel, $Workbook) }        # Close($false), Quit, ReleaseComObject; tolera $null
```

`Import-CotizadorCode` importa, en este orden: `src/modules/*.bas` (sin `modSecrets.bas` ni `modSecrets.example.bas`), `$SecretsFile`, `src/classes/*.cls`, `src/build/modBuild.bas` y el formulario `frmQuote` desde `frmQuote.vba` (1252). Con `-IncludeTests`, al final agrega `src/tests/*.bas`. Escribe en consola las mismas líneas `  + archivo` que hoy. `build.ps1` llama `Import-CotizadorCode` sin `-IncludeTests`, con `-SecretsFile` apuntando a `modSecrets.bas`, que sigue siendo obligatorio. Su `finally` usa `Close-ExcelApp` y `Restore-VbaAccess`. Cambia el orden de importación de `modSecrets` (va después de los demás módulos), pero no el resultado: sigue habiendo un componente por archivo.

### `build/test.ps1` (R6-R12)

```powershell
param(
    [int]$TimeoutSeconds = 300,
    [string]$Results = "",       # por omisión dist\test-results.txt
    [switch]$Visible,            # para depurar: Excel visible
    [switch]$Worker,             # interno: lo pone el padre
    [string]$PidFile = ""        # interno
)
```

**Padre** (sin `-Worker`):
1. Si `[type]::GetTypeFromProgID('Excel.Application')` es `$null`, escribe "las pruebas de VBA necesitan Excel de escritorio; no está instalado" y sale con 1 (R12).
2. `Enable-VbaAccess`. Crea `PidFile` en `%TEMP%` y borra el `Results` anterior.
3. `Start-Process powershell -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File',<este script>,'-Worker','-PidFile',…,'-Results',…[, '-Visible'] -NoNewWindow -PassThru`. Luego `WaitForExit($TimeoutSeconds * 1000)`.
4. Si se pasa del tiempo: `Stop-Process` del hijo y del PID que esté en `PidFile`, solo si ese proceso existe y se llama `EXCEL`. Escribe en `Results` y en consola el mensaje de R10 y el código de salida es 1.
5. Si no, el código de salida es el del hijo. Si `Results` no existe, escribe "el proceso de pruebas terminó sin resultados" y sale con 1.
6. En el `finally`: `Restore-VbaAccess` y borrar `PidFile` (R11).

**Hijo** (`-Worker`): `New-ExcelApp` (oculto salvo `-Visible`), escribe `Get-ExcelProcessId` en `PidFile` y crea el libro con `Workbooks.Add()`, sin `SaveAs` ([N1]). Sigue con `Import-CotizadorCode -IncludeTests -SecretsFile <modSecrets.bas o modSecrets.example.bas>` (R7, [N3]), `Run("'<libro>'!modBuild.BuildWorkbook", assets)` y `Run("…ListTests")`; por cada prueba, `BeginTest`, `Run` de la prueba con `try/catch` y `EndTest`; al final `$report = Run("…Report", ok, falla, fallas)`. Escribe `$report` en `Results` (UTF-8) y en consola: verde si todo pasó, rojo si no. Sale con 0 solo si la primera línea es `ok N / falla 0` con N > 0 (R8, R9). Una excepción de COM o de PowerShell se escribe como `error: <mensaje>` y sale con 1. Su `finally` hace `Close-ExcelApp`. No toca `AccessVBOM`: de eso se encarga el padre.

### `bflow.yaml` (R15)

```yaml
    - name: test
      run: powershell -NoProfile -ExecutionPolicy Bypass -File build/test.ps1
```

Va después de `lint`, con un comentario de que requiere Excel de escritorio y tarda 1-2 minutos. Sin `needs` ni `optional` ([N6]).

### Decisiones descartadas
- Rubberduck: necesita la interfaz del VBE y no corre sin supervisión (tarea).
- Encontrar las pruebas con una lista fija en el runner: se olvidaría actualizarla. Se leen del `VBProject` (D).
- Detectar el PID de Excel comparando procesos antes y después: puede confundirse con un Excel del usuario que se abre al mismo tiempo. Se usa `Hwnd`.
- Guardar el último error también fuera de modo prueba: no hace falta, y se mantiene igual el comportamiento de producción (D6).

### Superficie de seguridad
- **AccessVBOM**: confiar en el modelo de objetos de VBA queda encendido solo mientras corre el padre. Lo restaura su `finally`, que también corre con tiempo agotado y aunque haya matado al hijo. Si el usuario interrumpe con Ctrl+C también corre, salvo que mate a la consola.
- **Procesos**: solo se mata el PID que dejó el hijo, después de comprobar que se llama `EXCEL`. Nunca `Stop-Process -Name EXCEL`.
- **Secretos**: el libro de pruebas vive solo en memoria y se cierra sin guardar. `APP_PASSWORD` no aparece en `test-results.txt` ni en la consola. Con `modSecrets.example.bas` la clave es la de ejemplo, que es pública a propósito.
- **Datos**: las pruebas usan datos inventados (sin clientes reales). `dist/test-results.txt` sigue ignorado por git.
- **Errores**: toda falla sale con 1 y un mensaje legible: sin Excel, tiempo agotado, excepción de COM, cero pruebas. Nunca se salta en silencio.
- **Libro entregado**: no cambia. `src/tests/` no se importa en `build.ps1` (R5), y el modo prueba no se puede encender desde la interfaz.

### Desviaciones (implementación)
- Runner: un error sin manejar dentro de una prueba abre el diálogo de depuración si la prueba se llama con `Application.Run` desde VBA. Por eso `test.ps1` llama cada prueba por COM (donde el error regresa a PowerShell sin diálogo) con `ListTests`, `BeginTest`, `EndTest` y `Report`; `RunAllTests` queda para la ventana Inmediato. VBA reinicia las variables de módulo tras ese error, así que el conteo y las fallas viajan en PowerShell. Consecuencia: las fallas de assert anotadas antes del error en esa misma prueba se pierden, y el error se informa como HRESULT (`800A9C68`), no como el número de VBA.
- `modData.CellDate`: Excel entrega una hora sola (`hh:mm`) como `Double`, no `Date`, y `IsDate` la descartaba: `LoadOrder` perdía `OutTime`/`InTime` distintos de 00:00. `CellDate` acepta ahora un `Double` en rango de fecha. Lo encontró P3.

## Tasks

Cada tarea se verifica con `build/lint.ps1 -Base auto-repair-shop`, `build/build.ps1` y, desde T4, con `build/test.ps1`.

- [x] **T1 · Contrato** (R1, R2, R13, R16, R17, R18, R19): firmas públicas como esqueletos que compilan: el modo prueba de `modUtils` (`SetTestMode`, `IsTestMode`, `ClearLastError`, `LastErrorWhere`, `LastErrorNumber`, `LastErrorDescription`; `ShowError` todavía igual), `modAssert` (`AssertEqual`, `AssertTrue`, `AssertFails`), `modTestRunner` (`RunAllTests`, `RecordFailure`) y `modTestHelpers` con sus cinco procedimientos. Pruebas completas con los nombres y lo esperado de Design: las 14 `Test_*` de `modTestUtils` y las 11 de `modTestData`. Las de GH-4 y las funciones puras prueban código que ya existe. `build.ps1` no cambia, así que el `.xlsm` sigue igual (R5).
- [x] **T2 · Armado común** (R5, R14): `build/lib.ps1` con las seis funciones; `build.ps1` reescrito sobre ellas, con los mismos parámetros y mensajes. Para verificarlo, compara el `.xlsm` antes y después: mismos componentes VBA, mismas hojas, `BD_*` VeryHidden y protegidas, sin `modBuild`.
- [x] **T3 · Modo prueba, asserts y runner** (R1, R2, R3, R4, R9, R13): `ShowError` con modo prueba; cuerpo de `modAssert`, `modTestRunner` y `modTestHelpers`. Para verificarlo mientras no existe `test.ps1`: arma con `build.ps1 -Output` en una copia, importa `src/tests/` a mano en el editor de VBA y corre `RunAllTests` desde la ventana Inmediato. Comprueba además, con una `Test_` temporal que haga `Err.Raise` (no se sube), que el error regresa al runner y cuenta como falla (R3). Si abre el diálogo de depuración, reporta NEEDS_DECISION.
- [x] **T4 · `build/test.ps1`** (R6, R7, R8, R9, R10, R11, R12): padre con límite de tiempo e hijo `-Worker`. Verifica: todo en verde sale con 0; una falla temporal en una prueba sale con 1 y la detalla; quitando por un momento `modSecrets.bas` usa el de ejemplo; un error de compilación temporal en `src/tests/`, con `-TimeoutSeconds 60`, mata solo ese Excel y deja otro Excel abierto intacto; `AccessVBOM` queda como estaba en los tres casos.
- [x] **T5 · bflow y documentación** (R15, R20): paso `test` en `bflow.yaml` después de `lint`. En README, la sección "Pruebas automáticas": cómo correr, `-Visible`, `-TimeoutSeconds`, cómo agregar `Test_*` y qué sigue siendo manual. En CLAUDE.md, `src/tests/`, `build/lib.ps1` y `build/test.ps1` en el mapa, y en Comandos el reemplazo de "No hay tests automáticos". `bflow check GH-14` pasa con `lint` y `test`.
