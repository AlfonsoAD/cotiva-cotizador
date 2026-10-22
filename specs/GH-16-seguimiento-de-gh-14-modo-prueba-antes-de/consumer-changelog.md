# Changelog para consumidores · GH-16

## Cambios que rompen compatibilidad

Ninguno.

## Cambios

### `build\test.ps1`: línea `--- FAIL: <modulo>.<prueba>` por prueba en rojo (aditivo)

- Antes: por cada falla solo salían las líneas `FALLA <modulo>.<prueba>: <detalle>`, y la primera línea `ok N / falla M`.
- Después: justo antes de las `FALLA` de cada prueba que falla sale una línea `--- FAIL: <modulo>.<prueba>`. Una prueba en verde no agrega nada.
- Consumidor: si lees la salida de `test.ps1` o `dist\test-results.txt`, puedes usar esa línea para marcar la prueba en rojo. Las líneas `ok N / falla M` y `FALLA ...` no cambian.
- Códigos de salida: sin cambios (0 si todo pasa, 1 si falla).

### `build\test.ps1`: rutas con espacios (corrección)

- Antes: con `TEMP` o `TMP` con espacios (por ejemplo `C:\Users\Nombre Apellido\...`) el proceso hijo podía fallar al leer sus argumentos.
- Después: las rutas que pasa el proceso padre al hijo van entre comillas; funciona con rutas con espacios.
- Consumidor: no hay que cambiar nada.

### Modo prueba en `BeginTest` (corrección interna de `src\tests\`)

- El modo prueba se enciende antes de limpiar las hojas. Si la limpieza falla, la prueba ya no abre un `MsgBox` que cuelga la corrida; el error queda registrado como falla de la prueba.
- Solo afecta al libro de pruebas; el `.xlsm` entregado no cambia.
