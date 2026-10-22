# GH-22 · Parpadeo al cambiar de hoja y al abrir Mi empresa o el cotizador

## Brief

**Objetivo:** que al abrir "Mi empresa" o "Abrir cotizador" la pantalla se repinte una sola vez, al final, en vez de mostrar cada paso (desproteger, pintar, proteger, ocultar hojas), y averiguar si el parpadeo con las pestañas viene del código o de la portada.

**Entra**
- `ProtectAll` (`modSecurity`) apaga `Application.ScreenUpdating` mientras trabaja y al salir lo deja como estaba, también cuando hay error. Así quedan cubiertos todos los que la llaman: `OpenCompanySheet`, `PrepareWorkbook`, `AdminLock`, `modBuild` y las pruebas.
- `ProtectAll` solo cambia `.Visible` de una hoja de datos si todavía no está en `xlSheetVeryHidden`. Así no redibuja las pestañas cada vez que se llama.
- `OpenCompanySheet` (`modCompany`) apaga `ScreenUpdating` antes de desproteger y lo vuelve a encender en un bloque `Cleanup`, por donde pasan tanto la salida normal como la de error (igual que en `PrepareWorkbook`).
- Pruebas automáticas que revisan cómo queda `ScreenUpdating` después de cada rutina.
- Diagnóstico manual del parpadeo con las pestañas: se compara con un build de `v8.0.0` (anterior a GH-20) y con un libro sin macros. El resultado se anota en el PR.

**No entra**
- Código nuevo para el clic en las pestañas: el repo no tiene eventos de hoja y el build no importa código de hojas.
- Achicar las imágenes o las formas de la portada. Si el diagnóstico muestra que el parpadeo con las pestañas viene de ahí, se abre otra tarea con la evidencia.
- `DesignHome`, `DesignTemplate` y `modData`, que ya apagan y encienden `ScreenUpdating`.
- Cuánto tarda en aparecer `frmQuote` (`frmQuote.Show`).

**Decisiones nuevas**
- [N] `ScreenUpdating` se apaga dentro de `ProtectAll`. Al entrar se guarda el valor que traía y al salir se restaura ese mismo valor, no `True`. Así, si `PrepareWorkbook` o `DesignHome` la llaman con la pantalla ya apagada, no la enciende antes de tiempo. Descartado: apagarlo solo en `OpenCompanySheet`, porque deja fuera a `AdminLock` y al build; y restaurar siempre a `True`, porque encendería la pantalla a la mitad de la rutina que la llamó.
- [N] Si `ProtectAll` falla, restaura `ScreenUpdating` y vuelve a lanzar el mismo error (`Err.Raise` con su número, origen y descripción) para que lo reporte quien la llamó. Descartado: atrapar el error y regresar `False`, porque hoy `False` solo significa "clave incorrecta" y quien llama ya muestra ese mensaje.
- [N] El parpadeo con las pestañas se diagnostica primero y no se corrige a ciegas. Descartado: comprimir o quitar los logos de la portada sin saber si son la causa.

**Riesgos**
- Con Excel oculto (`test.ps1`), puede que `Application.ScreenUpdating` no conserve el valor asignado. Si así pasa y las pruebas de `ScreenUpdating` de T3 no son confiables, se quitan, se hacen a mano y el motivo se explica en el PR.
- Si el parpadeo con las pestañas ya existía en `v8.0.0`, o también aparece en un libro sin macros, el criterio se cumple como "no más que en un libro sin macros" y las pestañas no necesitan código.

**Tamaño:** XS. Son 2 archivos de `src/modules` y 1 de `src/tests`, más el diagnóstico manual.

## Discovery

<!-- Discovery completo; lo escribe bflow al cerrar el gate discovery. No lo edites. -->

## Requirements

Carril light: los criterios van dentro de cada tarea en Tasks.

## Design

Carril light: las firmas y los nombres van dentro de cada tarea en Tasks.

## Tasks

- [x] **T1 · `ProtectAll` sin redibujos intermedios** (`src/modules/modSecurity.bas`)
  - Al inicio, `Dim wasUpdating As Boolean`, `wasUpdating = Application.ScreenUpdating`, después `Application.ScreenUpdating = False` y `On Error GoTo Fail`. La revisión de `TryUnprotectWorkbook` y su `MsgBox` no cambian, pero ahora salen por `Cleanup` en vez de `Exit Function`.
  - En las hojas de datos, se oculta solo si hace falta: `If HIDE_DATA_SHEETS And ws.Visible <> xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden`.
  - `Cleanup:` pone `Application.ScreenUpdating = wasUpdating` y sale. `Fail:` guarda `Err.Number`, `Err.Source` y `Err.Description`, restaura `ScreenUpdating` y vuelve a lanzar el error con esos tres valores.
  - Agrega al comentario de la función que deja `ScreenUpdating` como lo traía quien la llama.
  - Criterios:
    - CUANDO se llama `ProtectAll` con `ScreenUpdating` en `False`, al terminar DEBE seguir en `False`. CUANDO se llama con `True`, al terminar DEBE estar en `True`.
    - CUANDO la clave no coincide, `ProtectAll` DEBE regresar `False`, mostrar el mismo mensaje de hoy y dejar `ScreenUpdating` como estaba.
    - CUANDO ocurre un error dentro de `ProtectAll`, DEBE restaurar `ScreenUpdating`, y el error DEBE llegar con el mismo número a quien la llamó.
    - CUANDO una hoja de datos ya está en `xlSheetVeryHidden`, `ProtectAll` NO DEBE volver a asignarle `.Visible`. Al terminar, las 7 hojas y la estructura DEBEN quedar protegidas igual que hoy.

- [x] **T2 · `OpenCompanySheet` con la pantalla apagada** (`src/modules/modCompany.bas`)
  - Pone `Application.ScreenUpdating = False` después de `SheetExists` y antes de `UnlockSheet`. `Activate` y el `Select` de `COMPANY_NAME_CELL` ocurren todavía con la pantalla apagada.
  - Nuevo bloque `Cleanup:` con `Application.ScreenUpdating = True` y `Exit Sub`. Todas las salidas normales van a `Cleanup`, incluida la de `ProtectAll()` en `False`. `Fail:` llama a `ShowError "OpenCompanySheet"` y hace `Resume Cleanup`.
  - El archivo conserva su codificación Windows-1252.
  - Criterios:
    - CUANDO el usuario pulsa "Mi empresa", la hoja DEBE aparecer ya pintada, protegida y con `COMPANY_NAME_CELL` seleccionada, sin que se vean los pasos intermedios.
    - CUANDO `OpenCompanySheet` termina, por la salida normal o por la de error, `ScreenUpdating` DEBE quedar en `True`.

- [x] **T3 · Pruebas** (`src/tests/modTestCompany.bas`)
  - `Test_ProtectAll_ConservaScreenUpdating`: con `ScreenUpdating = False` llama a `ProtectAll` y espera `False`; con `True`, espera `True`. Al final lo deja en `True` aunque falle una aserción (mismo patrón `On Error GoTo Cleanup` del módulo).
  - `Test_ProtectAll_NoReasignaHojasOcultas`: llama dos veces seguidas a `ProtectAll` y revisa que las hojas de datos sigan en `xlSheetVeryHidden` (`HIDE_DATA_SHEETS` es `True` en `modConfig`) y protegidas.
  - `Test_OpenCompanySheet_DejaPantallaEncendida`: con `ScreenUpdating = False` llama a `OpenCompanySheet` y espera `True`, con `SHEET_COMPANY` como hoja activa y protegida.
  - Criterio: `build\test.ps1` DEBE pasar con las pruebas nuevas y con las que ya existen (las de color de GH-20 llaman a `ProtectAll`).

- [ ] **T4 · Diagnóstico manual y verificación** (sin código; el resultado va en el PR)
  - Con el build de esta rama, pulsa 3 veces "Mi empresa" (desde `frmQuote`) y 3 veces "Abrir cotizador" (desde la portada), y anota si parpadea.
  - Cambia de hoja con las pestañas (Inicio, Plantilla, Mi empresa) en tres libros: este build, un build de `v8.0.0` y un libro de Excel sin macros con una imagen. Anota en cuáles parpadea.
  - Si este build parpadea con las pestañas y `v8.0.0` no, registra en el PR la diferencia de formas e imágenes de la portada (nombres en `modConfig`, `HOME_*_LOGO_NAME`) y abre una tarea nueva. No se corrige en esta.
  - Criterios:
    - CUANDO se pulsan los botones "Mi empresa" o "Abrir cotizador", NO DEBE verse cómo se repintan, paso a paso, la hoja ni las pestañas.
    - CUANDO se cambia de hoja con las pestañas, el parpadeo NO DEBE ser mayor que en un libro de Excel sin macros, o la diferencia DEBE quedar documentada en el PR con su tarea de seguimiento.
