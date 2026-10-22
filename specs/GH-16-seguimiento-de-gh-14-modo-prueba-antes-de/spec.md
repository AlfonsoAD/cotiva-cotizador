# GH-16 · Seguimiento de GH-14: modo prueba antes de limpiar hojas y texto viejo de RunAllTests

## Brief

**Objetivo:** que `BeginTest` encienda el modo prueba antes de limpiar las hojas (así una falla al limpiar ya no abre un `MsgBox` que cuelga la corrida), que `test.ps1` marque cada prueba en rojo con una línea que bflow reconoce y funcione con `TEMP` con espacios, y quitar el texto que todavía describe el flujo viejo de `RunAllTests`.

**Entra**
- `src/tests/modTestRunner.bas` `BeginTest`: `SetTestMode True` pasa antes de `ResetDataSheets`.
- `build/test.ps1`, salida: por cada prueba que falla, una línea `--- FAIL: <modulo>.<prueba>` (formato que reconoce bflow; hoy solo sale `FALLA ...`, bflow no ve las pruebas en rojo y eso bloquea el contrato de GH-5).
- `build/test.ps1`, padre: entrecomillar las rutas que pasa al hijo `-Worker`, para que corra con `TEMP`/`TMP` con espacios (p. ej. `C:\Users\Alfonso Andrade\...`).
- `build/test.ps1` (bloque `.DESCRIPTION`, línea 7): describir el flujo real, una llamada por COM a cada prueba con `ListTests`/`BeginTest`/`EndTest`/`Report`.
- `specs/GH-14-…/spec.md`: R1, R6, la secuencia de la línea 136 y la descripción del hijo `-Worker` (línea 218) dicen lo que hace hoy el código.
- `src/modules/modUtils.bas` `SetTestMode`: un solo comentario.
- Formato: línea en blanco antes del comentario de `FindTests` (`modTestRunner.bas`) y antes de `## Flujo de trabajo` (`README.md`).

**No entra**
- Cambiar `RunAllTests`, `EndTest` ni `ResetDataSheets`. `EndTest` ya apaga el modo prueba siempre, falle o no `BeginTest`.
- Reescribir el resto de las menciones a `RunAllTests` en la spec de GH-14 (líneas 8, 38, 47, 66, 87, 129, T1, T3): ahí el nombre es correcto, porque `RunAllTests` sigue existiendo para la ventana Inmediato.
- Cambiar el formato `FALLA <modulo>.<prueba>: <detalle>` que arma `RecordFailure`, o la primera línea `ok N / falla M`.
- Cambios al `.xlsm` entregado: `src/tests/` no viaja en él y el comentario de `modUtils` no cambia comportamiento.

**Decisiones nuevas**
- [N1] Sin prueba automática del caso "falla `ResetDataSheets`". Para provocarla habría que meter en el runner un gancho que la fuerce, y no vale la pena para un cambio de orden de dos líneas. Se verifica leyendo el código y con `test.ps1` en verde. Descartado: un gancho `mForceResetFailure` en `modTestRunner` con su `Test_`.
- [N2] En la spec de GH-14 solo se corrige lo que describe el flujo de `test.ps1` (R1, R6, línea 136, línea 218). R1 pasa a nombrar a los dos caminos (`RunAllTests` y `ListTests`/`BeginTest` por COM). Descartado: quitar `RunAllTests` de toda la spec, porque la función sigue existiendo.

- [N3] La línea `--- FAIL: <modulo>.<prueba>` la agrega `test.ps1`, una por cada prueba que falla, justo antes de sus líneas `FALLA ...`. Sale en consola y en `dist/test-results.txt`. Descartado: cambiar el prefijo de `RecordFailure` a `--- FAIL:`, porque daría una línea por assert en lugar de una por prueba y cambiaría también la salida de `RunAllTests`.
- [N4] El padre arma `-ArgumentList` con cada ruta entre comillas dobles (`$PSCommandPath`, `-PidFile`, `-Results`). Windows PowerShell 5.1 une el arreglo con espacios sin entrecomillar, y por eso el hijo recibía un pedazo de la ruta como `-TimeoutSeconds`. Descartado: cambiar `Start-Process` por `ProcessStartInfo`, que es más código para el mismo efecto.

**Riesgos**
- Con el modo prueba encendido, un `ShowError` dentro de `EnsureWorkbookStructure` o `ProtectAll` se registra en lugar de abrir `MsgBox`. Es justo lo que se busca, y `ClearLastError` corre después, así que la prueba arranca sin error registrado.

- `--- FAIL:` solo aparece cuando hay fallas: una corrida en verde no cambia su salida.
- Una ruta que termine en `\` justo antes de la comilla de cierre rompería el entrecomillado. Las tres son rutas de archivo y no terminan en `\`.

**Tamaño:** S. Cuatro archivos de código o docs y una spec previa; dos arreglos chicos en `test.ps1`.

## Discovery

<!-- Discovery completo; lo escribe bflow al cerrar el gate discovery. No lo edites. -->

## Requirements

<!-- Criterios EARS numerados R1..Rn, cada uno [D] (del discovery) o [N] (nuevo). -->

## Design

<!-- Estructuras, firmas, decisiones tomadas y descartadas, superficie de seguridad. -->

## Tasks

- [ ] **T1 · Orden en `BeginTest`.** En `src/tests/modTestRunner.bas`, `BeginTest` queda `On Error GoTo Fail` → `SetTestMode True` → `ResetDataSheets` → `ClearLastError`. El comentario de arriba dice que activa el modo prueba y después deja el libro como recién abierto. Agrega una línea en blanco entre el `End Function` de `Report` y el comentario de `FindTests`.
  - Criterio: CUANDO `ResetDataSheets` lanza un error dentro de `BeginTest`, el modo prueba DEBE estar ya encendido (`IsTestMode() = True`) y la falla DEBE quedar anotada con `RecordFailure`, sin `MsgBox`.
  - Criterio: CUANDO `BeginTest` termina sin error, `IsTestMode()` DEBE ser True y no DEBE haber error registrado (`LastErrorNumber = 0`).
  - Criterio: `EndTest` no cambia y sigue apagando el modo prueba.
- [ ] **T2 · Comentario de `SetTestMode`.** En `src/modules/modUtils.bas`, encima de `SetTestMode` queda un solo comentario: `' Enciende o apaga el modo prueba; solo lo usa el runner de pruebas (modTestRunner): el libro entregado nunca entra en modo prueba.`
  - Criterio: no hay dos líneas de comentario seguidas que repitan el mismo propósito. La codificación del archivo sigue en Windows-1252 con CRLF.
- [ ] **T3 · Textos del flujo de `test.ps1`.**
  - `build/test.ps1` `.DESCRIPTION`: cambia "corre modTestRunner.RunAllTests" por "llama por COM a cada prueba (modTestRunner.ListTests, BeginTest, EndTest y Report)".
  - `specs/GH-14-pruebas-automaticas-de-vba-con-excel-test-ps1/spec.md`:
    - R1 empieza "CUANDO corre el runner (`RunAllTests` desde VBA, o `ListTests` y `BeginTest`/`EndTest` por prueba desde `test.ps1`)".
    - R6 termina "y correr cada prueba en él por COM (`ListTests`, `BeginTest`, `EndTest`, `Report`)".
    - La línea 136 dice el orden nuevo: `SetTestMode True`, `ResetDataSheets`, `ClearLastError`.
    - La línea 218 cambia `$report = Run("'<libro>'!modTestRunner.RunAllTests")` por: `Run("…ListTests")`; por cada prueba, `BeginTest`, `Run` de la prueba con `try/catch` y `EndTest`; al final `$report = Run("…Report", ok, falla, fallas)`.
  - `README.md`: línea en blanco antes de `## Flujo de trabajo`.
  - Criterio: `grep -n RunAllTests build/test.ps1` no regresa nada. En la spec de GH-14, ninguna línea dice que `test.ps1` corre `RunAllTests`.
- [ ] **T4 · `--- FAIL:` por prueba en `test.ps1`** ([N3]). En el ciclo del hijo `-Worker`, cuando una prueba falla (`$recorded` o `$failure`), se agrega `"--- FAIL: $name"` a `$failureLines` antes de sus líneas `FALLA ...`. Se actualizan el comentario de `Report` en `modTestRunner.bas` y el párrafo de salida del README (línea 61): `test.ps1` antepone `--- FAIL: <modulo>.<prueba>` a las fallas de cada prueba.
  - Criterio: CUANDO una prueba falla, la consola y `dist/test-results.txt` DEBEN tener una línea que empieza exactamente con `--- FAIL: <modulo>.<prueba>`, una sola por prueba aunque tenga varios asserts en rojo, seguida de sus líneas `FALLA ...`, y `test.ps1` DEBE salir con 1.
  - Criterio: CUANDO todas pasan, la salida DEBE ser la de hoy (sin `--- FAIL:`) y la primera línea sigue siendo `ok N / falla 0`.
- [ ] **T5 · Rutas con espacios al lanzar el hijo** ([N4]). En el padre de `build/test.ps1`, `$childArgs` pasa `$PSCommandPath`, `$pidFilePath` y `$Results` entre comillas dobles (`'"' + $ruta + '"'`). El resto de los argumentos no cambia.
  - Criterio: CUANDO `TEMP` y `TMP` apuntan a una carpeta con espacios, `test.ps1` DEBE correr completo y salir con 0, sin error de conversión en `-TimeoutSeconds`.
  - Criterio: CUANDO `-Results` es una ruta con espacios, el archivo DEBE quedar en esa ruta.
- [ ] **T6 · Verificación.**
  - `powershell -ExecutionPolicy Bypass -File build\test.ps1` sale con 0 y `dist/test-results.txt` reporta el mismo número de pruebas que antes, con `falla 0` ([N1]).
  - Con `$env:TEMP` y `$env:TMP` en una carpeta con espacios (p. ej. `C:\Users\Alfonso Andrade\AppData\Local\Temp` en forma larga, no la corta `ALFONS~1`) y `-Results` en una ruta con espacios, sale con 0.
  - Con una `Public Sub Test_` temporal que falla dos asserts (no se sube), la salida trae una sola línea `--- FAIL: <modulo>.Test_...` seguida de dos `FALLA ...`, y sale con 1. La prueba temporal se borra antes del commit.
  - Sin `.xlsm` en el repo.
