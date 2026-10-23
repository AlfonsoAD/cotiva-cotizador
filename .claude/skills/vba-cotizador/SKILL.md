---
name: "vba-cotizador"
description: "Reglas de arquitectura, maquetado y clean code para el cotizador Excel/VBA de Cotiva (src/, modData, frmQuote, plantillas). Úsala al escribir o revisar código VBA de este repo."
---

# Cotizador VBA: reglas del repo

## Flujo de trabajo
- El código vive como texto en `src/`. Nunca edites el `.xlsm` a mano: compila con `build/build.ps1` y exporta con `export.ps1`.
- `modSecrets.bas` no se sube.
- La rama base de cada producto (`dev` para el cotizador, `auto-repair-shop` para talleres) es genérica: sirve a cualquier cliente de ese producto. Lo de un cliente (nombre, logos, colores, datos fiscales, precios, catálogos) es configuración ("Mi empresa", `assets/`, `BD_*`), no código. Lo que de verdad solo aplica a un cliente va en `cliente/<nombre>`, que sale de la base y nunca se mergea de vuelta.

## Arquitectura (una responsabilidad por módulo)
- `modConfig`: constantes (celdas `CELL_*`, filas, prefijos de folio). Nada de celdas mágicas en el código.
- `modStrings`: TODO texto visible (`UI_*`, `MSG_*`, `TPL_*`). Se formatea con `Fmt(...)`.
- `modTheme`: colores y fuente como funciones (`BrandOrange()`); nunca `RGB()` suelto en otros módulos.
- `modData`: ÚNICO módulo que lee o escribe hojas `BD_*`. Lee rangos completos a arreglos, no celda por celda. Mantén sus firmas estables (se podrá migrar a SQLite o API).
- `modTemplate`: diseña, llena y exporta plantillas. Clases (`Quote`, `QuoteItem`) = modelo sin acceso a hojas.
- Los formularios solo orquestan: validan, arman el modelo y llaman a `modData` o `modTemplate`.

## Clean code
- `Option Explicit` siempre. Código e identificadores en inglés; comentarios y textos en español.
- Nombres: `PascalCase` para procedimientos y propiedades, `camelCase` para locales, `mVariable` para módulo, `UPPER_SNAKE` para constantes.
- Funciones cortas que hacen una cosa. `ByVal` por defecto; `ByRef` solo en el patrón `TryParseX(... , ByRef out) As Boolean`.
- Sin `Select`, `Activate` ni `ActiveSheet`: referencias explícitas a `Worksheet` y `Range`.
- Manejo de errores con un solo patrón:
  ```vba
  On Error GoTo Fail
  Application.ScreenUpdating = False: UnlockSheet ws
  ' ...
  Cleanup:
      LockSheet ws: Application.ScreenUpdating = True
      Exit Sub
  Fail:
      ShowError "ProcName": Resume Cleanup
  ```
- El texto del usuario pasa por `SafeCellText` antes de escribirse en una celda. Montos con `Round2`.
- Sin código muerto ni `Debug.Print` olvidados. Comenta el porqué, no el qué.

## Maquetado de formularios
- Los formularios se construyen por código con los helpers públicos de `modFormKit` (`BeginFormKit` al inicio de `BuildLayout`, `Place`, `AddTextBox`, `AddLabel`, `AddButton`, `AddBand`; reciben el formulario o la pestaña como `parent`); nunca se copian como privados dentro de un formulario; nada arrastrado en el diseñador.
- Rejilla: margen 18 pt, controles de 22 pt de alto, etiqueta de 8 pt en `TextGray()` arriba de cada campo, título de sección en mayúsculas, 8 pt, negritas y color de marca.
- Un solo botón primario por vista. Define el `TabIndex` en el orden de captura.
- Revisa que la ventana quepa en 1366×768; si no, usa `MultiPage`.

## Plantillas PDF
- Las crea `Design*Template` desde código; los datos se escriben en celdas de `modConfig`.
- Carta, una página, hojas protegidas; solo las celdas editables quedan con `Locked = False`.

## Antes de dar algo por terminado
- Compila (Depurar › Compilar) sin errores, ejecuta `build.ps1` y prueba: nuevo, guardar, cargar, PDF y datos de versiones anteriores.