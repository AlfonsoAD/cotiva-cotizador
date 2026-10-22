# GH-18 · Fase 2 · Colores y logo del formulario configurables desde Mi empresa y assets/

## Brief

**Objetivo:** que `frmQuote` tome sus colores de "Mi empresa" (B11 principal, B12 secundario) y que su logo tenga un solo nombre, `logo-form.jpg`, en el build y en el respaldo junto al libro, sin datos de un taller en el código. Rehace #5 en el carril light, con sus decisiones ya ratificadas.

**Entra**
- Los 20 usos de `BrandOrange()`/`BrandInk()` en `src/forms/frmQuote.vba` pasan a `PrimaryColor()`/`SecondaryColor()` de `modCompany` (ya existen y caen a los colores de Cotiva).
- `FORM_LOGO_FILE` (`modConfig`) pasa de `logo_cotizador.jpg` a `logo-form.jpg`, y `modBuild` usa la constante en lugar del texto fijo (hoy el build busca un nombre y el respaldo otro).
- `assets/README.md`: nombre, solo JPG, respaldo junto al libro, renombrar `logo_cotizador.jpg`, build sin logo y la regla de colores para formularios nuevos.
- Pruebas de #5 reaprovechadas de `feature/GH-5-fase-2-colores-y-logo-del-formulario`: `src/tests/modTestCompany.bas` y `SetCompanyCell` en `modTestHelpers.bas`.

**No entra**
- `modHome`, `modCompany` y `modTemplate` (colores de Cotiva directo en la portada): issue aparte. `modTheme` no cambia y `BrandInkLight()` sigue fijo.
- `frmOrder` (Fase 3, #6): solo la regla escrita en el README.
- El crédito de Cotiva e innobytes. `AddFormLogo` por dentro: su `Debug.Print` no es línea nueva, así que el lint no lo marca.

**Decisiones (ratificadas en #5)**
- [N] El texto sobre fondo de color de empresa (botones primario y oscuro, `btnAdd`, la palabra de marca `c_Brand` en la barra sin logo) usa `ContrastTextColor(fondo)` en vez de `vbWhite`, como la Plantilla. Descartado: blanco fijo, ilegible con un color claro.
- [N] Los colores se leen una vez en `BuildLayout` a `mPrimary`/`mSecondary`. Descartado: llamar `PrimaryColor()` en cada sitio (hasta 20 lecturas de hoja y `btnAdd` podría cambiar a mitad de la sesión).
- [N] `modBuild` usa `FORM_LOGO_FILE`. Descartado: dos textos con el mismo nombre que vuelven a separarse.
- [N] Cada prueba fija B11/B12 al empezar y los deja vacíos al terminar, porque `ResetDataSheets` no limpia "Mi empresa". Descartado: cambiar el runner.
- [N] Nuevo en GH-18: la prueba del logo se llama `Test_FormLogoFile_NombreUnico` (sin `T1_`) y entra en la misma tarea que el cambio de la constante, así nunca queda en rojo. Descartado: dejarla en rojo hasta otra tarea, como pedía el carril full.

**Riesgos**
- `BrandInkLight()` (texto `c_Module` de la barra sin logo) puede leerse mal con un secundario claro. Lo decidió el humano; se mira en la prueba manual.
- Quien tenga `logo_cotizador.jpg` junto al libro deja de verlo hasta renombrarlo. Va en el README.
- La apariencia del formulario solo se prueba a mano.

**Tamaño:** S (≈1 h). Toca `frmQuote.vba`, `modConfig.bas`, `modBuild.bas`, `assets/README.md` y dos archivos de `src/tests/`.

## Discovery

<!-- Discovery completo; lo escribe bflow al cerrar el gate discovery. No lo edites. -->

## Requirements

Carril light: los criterios van en cada tarea.

## Design

Carril light: el diseño va en cada tarea.

## Tasks

- [x] **T1 · Nombre único del logo y pruebas**
  - Traer de la rama de #5 `src/tests/modTestCompany.bas` y `SetCompanyCell(cellAddress, value)` de `src/tests/modTestHelpers.bas` (`TryUnprotectWorkbook`, `UnlockSheet` de `SHEET_COMPANY`, escribe, `ProtectAll`). Renombrar `Test_T1_FormLogoFile_NombreUnico` a `Test_FormLogoFile_NombreUnico`.
  - `src/modules/modConfig.bas` (l.110): `Public Const FORM_LOGO_FILE As String = "logo-form.jpg"`, con comentario: nombre del logo del formulario en `assets/` y, de respaldo, junto al libro.
  - `src/build/modBuild.bas` (l.37-38): `assetsFolder & "\" & FORM_LOGO_FILE` en el `Dir$` y en `AddFormLogo`. `AddFormLogo`, `PlaceHeaderLogo` y `build.ps1` no cambian.
  - Criterios: `Test_FormLogoFile_NombreUnico` pasa (`FORM_LOGO_FILE` = `"logo-form.jpg"`). Las de color pasan: B11/B12 vacía o inválida (`#12G45Z`, `12345`) regresan `BrandOrange()`/`BrandInk()`; válida con o sin `#`, en mayúsculas o minúsculas (`#1a2b3c`, `1A2B3C`, `0F7F3A`) regresa ese RGB. Sin `assets/logo-form.jpg` el build termina igual que hoy. Los valores de prueba son genéricos. `build\test.ps1` en verde.
- [x] **T2 · Colores de empresa en `frmQuote`**
  - `Private mPrimary As Long` y `Private mSecondary As Long` junto a `mFont`; al inicio de `BuildLayout`, antes de `PlaceHeaderLogo`, `mPrimary = PrimaryColor()` y `mSecondary = SecondaryColor()`.
  - `BrandOrange()` → `mPrimary` en l.103, 125, 153, 216, 336, 595. `BrandInk()` → `mSecondary` en l.106, 168, 179-182, 200, 212, 215, 301, 339, 343, 613.
  - Texto sobre fondo de color: en `AddButton`, `kindPrimary` → `ContrastTextColor(mPrimary)` y `kindDark` → `ContrastTextColor(mSecondary)`; en l.595 y 613 también `btnAdd.ForeColor = ContrastTextColor(<fondo nuevo>)`; `c_Brand` (l.107) usa `ContrastTextColor(mSecondary)`.
  - Se quedan igual: el `vbWhite` de la barra con logo y de fondos, `TextGray()`, `Neutral300()`, `BrandInkLight()` de `c_Module`, `mFont` y el crédito del pie.
  - Criterios: una búsqueda en `frmQuote.vba` no encuentra `BrandOrange(` ni `BrandInk(`; `modTheme.bas` sin cambios; ningún nombre, logo ni color de un taller en `src/`. Prueba manual tras el build: con B11/B12 vacíos los colores son los de hoy; con un color claro y uno oscuro cambian secciones, total, botones y textos, y el texto de los botones se lee; sin logo, la barra va en el secundario; con logo, barra blanca y línea en el primario; un `logo-form.jpg` junto a un libro sin imagen incrustada se usa de respaldo.
- [x] **T3 · Documentación**
  - `assets/README.md`: la fila de `logo-form.jpg` dice que es el único nombre aceptado y que debe ser JPG (el control Image no carga PNG), de preferencia con fondo blanco. Párrafos: (1) el build lo incrusta; sin el archivo, el formulario usa la barra del color secundario de "Mi empresa" (azul tinta de Cotiva si está vacío) con texto; (2) un `logo-form.jpg` junto al `.xlsm` se usa si el libro no trae imagen incrustada; antes se llamaba `logo_cotizador.jpg` y hay que renombrarlo; (3) los colores de este formulario y de los que se agreguen salen de B11/B12 con los de Cotiva de respaldo: en el código se usan `PrimaryColor()`/`SecondaryColor()`, nunca `BrandOrange()`/`BrandInk()` directo.
  - Criterio: el README cubre nombre, formato, respaldo, renombrado, qué pasa sin logo y la regla para formularios nuevos.
