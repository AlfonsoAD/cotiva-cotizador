# GH-20 · Seguimiento de GH-18: colores de empresa en portada y Mi empresa, y ajustes de pruebas

## Brief

**Objetivo:** que la portada (Inicio) y la hoja "Mi empresa" tomen los colores capturados en B11/B12 y se repinten solas al cambiarlos, y cerrar los pendientes de pruebas del review de GH-18.

**Entra**
- `modHome`: regla, botón "Abrir cotizador", viñetas y textos del diseño usan `PrimaryColor()`/`SecondaryColor()`; el texto del botón usa `ContrastTextColor(PrimaryColor())` en vez de `vbWhite`.
- `modCompany`: título A1 en `PrimaryColor()` y texto de las celdas de captura B4:B12 en `SecondaryColor()`.
- Repintado sin rediseñar: `RefreshHomeColors` (nuevo) y `RefreshColorSwatches` (ampliado) se llaman desde `PrepareWorkbook`, es decir, cada vez que se pulsa "Abrir cotizador".
- Pruebas: B12 inválido, limpieza de B11/B12 aun con error, pruebas del repintado de portada y Mi empresa; comentario de `FORM_LOGO_FILE` alineado.

**No entra**
- `modTheme` (incluido `BrandInkLight()`), `modTemplate` y `frmQuote`, que ya quedaron en GH-18.
- El crédito de Cotiva e innobytes: los textos de respaldo de B2 ("Cotiva") y B15 ("innobytes"), B16 y la versión siguen con sus colores actuales.
- Repintar en el momento en que se edita B11/B12 (evento `Worksheet_Change`).

**Decisiones nuevas**
- [N] Disparador del repintado: `PrepareWorkbook` (botón "Abrir cotizador") y `OpenCompanySheet`. Descartado: un evento `Worksheet_Change` en "Mi empresa", porque el build no importa código de hojas, y dejarlo solo en la macro `DesignHome`, porque así no se cumple "sin tocar código ni correr macros de administrador".
- [N] El repintado solo cambia color (fondos de formas, color de letra). No borra ni rehace nada, así que no toca logos ni lo capturado. Descartado: llamar a `DesignHome True` en cada apertura, que es más lento y borra formas.
- [N] `RefreshColorSwatches` conserva su nombre y además pinta A1 y B4:B12, así sus tres llamadas actuales quedan igual. Descartado: renombrarla.
- [N] Limpieza con error: `On Error GoTo Cleanup` en cada prueba de color. Se guarda el error, se limpian B11 y B12 con un helper nuevo `ClearCompanyColors` y se vuelve a lanzar el error para que el runner lo registre. Descartado: dejar la celda sucia y que la siguiente prueba la sobrescriba, porque esconde fallas que dependen del orden.

**Riesgos**
- Un color secundario muy claro deja poco legibles los textos de la portada y de B4:B12, que van sobre fondo blanco. Es lo que pide la tarea (el secundario se usa como color de texto) y el respaldo a Cotiva solo aplica a valores inválidos. Lo verifica una prueba manual.
- Sin protección `UserInterfaceOnly` (se pierde al reabrir el libro), pintar una hoja protegida falla: el repintado desprotege antes y `ProtectAll` vuelve a proteger después.

**Tamaño:** S. Son 4 archivos de `src/modules`, 2 de `src/tests` y el README.

## Discovery

<!-- Discovery completo; lo escribe bflow al cerrar el gate discovery. No lo edites. -->

## Requirements

Carril light: los criterios van dentro de cada tarea en Tasks.

## Design

Carril light: las firmas y los nombres van dentro de cada tarea en Tasks.

## Tasks

- [x] **T1 · Portada con colores de empresa** (`src/modules/modHome.bas`)
  - Nuevo `Private Sub PaintHomeColors(ByVal ws As Worksheet)`:
    - Forma `HOME_RULE_NAME`: `.Fill.ForeColor.RGB = PrimaryColor()`.
    - Forma `HOME_BUTTON_NAME`: relleno `PrimaryColor()` y letra `ContrastTextColor(PrimaryColor())`.
    - B4 (título) y B10:B12 (ventajas): letra `SecondaryColor()`. Después, el primer carácter de cada viñeta (`HOME_BULLET`, filas `HOME_FIRST_FEATURE_ROW`..+2) en `PrimaryColor()`.
    - Si una forma no existe, se omite sin error (usa `ShapeExists`).
  - Nuevo `Public Sub RefreshHomeColors()`: si no existe `SHEET_HOME`, sale. Si existe, `UnlockSheet` y `PaintHomeColors`; no vuelve a proteger, eso lo hace quien la llama.
  - `DesignHome` deja de usar `BrandOrange()`/`BrandInk()` en las líneas 50, 69, 116 y 188, y llama a `PaintHomeColors ws` antes de `FitHomeLogos`. La letra base de A1:H40 (línea 50) pasa a `SecondaryColor()`, pero los textos de respaldo de B2 y B15 (crédito) se quedan en `BrandInk()` de forma explícita. B5, B8, B16 y C15 conservan `TextGray()`.
  - Criterios:
    - CUANDO B11 tiene un color válido y se corre `RefreshHomeColors`, la regla y el botón DEBEN tener ese relleno y el texto del botón DEBE ser `ContrastTextColor(PrimaryColor())`.
    - CUANDO B12 tiene un color válido, B4 y B10:B12 DEBEN tener esa letra, salvo el carácter de la viñeta, que DEBE ser `PrimaryColor()`.
    - CUANDO B11/B12 están vacías o no son válidas, DEBEN usarse `BrandOrange()`/`BrandInk()`.
    - B2, B15, B16 y C15 NO DEBEN cambiar de color.

- [x] **T2 · Hoja "Mi empresa" con colores de empresa** (`src/modules/modCompany.bas`)
  - `DesignCompanySheet`: la letra base de A1:D20 (línea 113) pasa a `SecondaryColor()` y el título A1 (línea 123) a `PrimaryColor()`. Las etiquetas de A4:A12, A2 y las notas conservan `TextGray()`.
  - `RefreshColorSwatches`: además de C11/C12, pinta A1 con `PrimaryColor()` y B4:B12 con `SecondaryColor()`. Actualiza su comentario.
  - Criterio: CUANDO cambian B11/B12 y se corre `RefreshColorSwatches`, A1 y B4:B12 DEBEN tomar los colores nuevos sin borrar lo capturado.

- [x] **T3 · Repintado al abrir el cotizador** (`src/modules/modApp.bas`)
  - En `PrepareWorkbook`, después de `EnsureWorkbookStructure` y antes de `ProtectAll`: `RefreshHomeColors`, después `UnlockSheet` de `SHEET_COMPANY` (si existe) y `RefreshColorSwatches`.
  - Criterio: CUANDO el usuario cambia B11/B12 y pulsa "Abrir cotizador", la portada y "Mi empresa" DEBEN quedar con los colores nuevos y protegidas, sin correr macros de administrador.

- [x] **T4 · Pruebas de color y limpieza con error** (`src/tests/modTestCompany.bas`, `src/tests/modTestHelpers.bas`)
  - `modTestHelpers`: nuevo `Public Sub ClearCompanyColors()`, que deja B11 y B12 en `""` usando `SetCompanyCell`.
  - Nueva `Test_SecondaryColor_InvalidaUsaCotiva`: `"#12G45Z"` y `"12345"` en B12 esperan `BrandInk()`, igual que la prueba de B11.
  - Nuevas pruebas:
    - `Test_RefreshHomeColors_UsaColorDeEmpresa`: B11 `"#0F7F3A"`, B12 `"#1A2B3C"`. Se verifica el relleno de la regla y del botón, el texto del botón (`ContrastTextColor`), la letra de B4 y que el primer carácter de B10 sea el principal.
    - `Test_RefreshColorSwatches_PintaMiEmpresa`: se verifica A1 y B4.
    - `Test_RefreshHomeColors_VaciasUsaCotiva`: con B11/B12 vacías, el botón usa `BrandOrange()`.
  - Todas las pruebas de este módulo que escriben B11/B12 siguen este patrón:
    ```vba
    Dim errNumber As Long, errText As String
    On Error GoTo Cleanup
    ' ... casos y asserts ...
    Cleanup:
        errNumber = Err.Number: errText = Err.Description
        ClearCompanyColors
        If errNumber <> 0 Then Err.Raise errNumber, , errText
    ```
    Las pruebas que repintan terminan con `RefreshHomeColors`, `RefreshColorSwatches` y `ProtectAll` en `Cleanup` para que el libro de pruebas vuelva a los colores de Cotiva.
  - Criterio: CUANDO una prueba de color lanza un error a la mitad, B11 y B12 DEBEN quedar vacías y el runner DEBE registrar el error.
  - `build\test.ps1` en verde.

- [x] **T5 · Comentario de `FORM_LOGO_FILE` y documentación** (`src/modules/modConfig.bas`, `README.md`)
  - En la línea 110, el comentario empieza en la misma columna que los de `PDF_FOLDER`, `SAVE_WORKBOOK_ON_SAVE` y `HIDE_DATA_SHEETS`. Si no cabe, va en la línea de arriba.
  - README, sección de "Mi empresa" y colores: los colores se aplican también a la portada y a "Mi empresa" al pulsar "Abrir cotizador". Agrega una entrada al changelog para consumidores, como en GH-18.
  - Prueba manual (en el PR): build, cambiar B11 a un color claro (`#F5D000`) y B12 a uno oscuro, pulsar "Abrir cotizador" y revisar la portada (botón con texto oscuro legible, crédito sin cambios) y "Mi empresa". Después, vaciar B11/B12 y comprobar que vuelven los colores de Cotiva.
  - Criterio: ningún nombre, logo ni color de un taller concreto en el código; `grep` de `BrandOrange\(\)|BrandInk\(\)` en `modHome`/`modCompany` solo encuentra los créditos de B2/B15 y los respaldos de `PrimaryColor`/`SecondaryColor`.
