# Changelog para consumidores: GH-14 (pruebas automáticas de VBA)

Rama base: `auto-repair-shop`. Cambia lo que consumen otros equipos: la salida de `bflow check` y un
comando nuevo. El `.xlsm` entregado no cambia.

## Cambios que rompen compatibilidad

### 1. `bflow check` ahora ejecuta el paso `test` y necesita Excel de escritorio

- **Antes:** `bflow check` corría solo `lint`. Podía correr en cualquier máquina Windows o en un runner sin Office.
- **Después:** `bflow check` corre `lint` y después `test` (`powershell -NoProfile -ExecutionPolicy Bypass -File build/test.ps1`).
  El paso no es opcional: si falla, el check falla.
- **Efecto:** el check tarda 1-2 minutos más y exige Excel de escritorio en la máquina donde corre.
  En un runner sin Excel el paso `test` falla con el mensaje "las pruebas de VBA necesitan Excel de escritorio";
  no se salta en silencio.
- **Qué debes hacer:** correr `bflow check` en una máquina con Excel. Si el entorno no lo tiene, pide
  una máquina con Excel de escritorio; no quites el paso del `bflow.yaml` sin una decisión de la persona
  responsable del flujo.

### 2. Si una prueba falla, el check falla

- `test.ps1` sale con 1 si falla una prueba, si no hay ninguna prueba, si no hay Excel o si se pasa el
  límite de tiempo.
- **Qué debes hacer:** revisar `dist/test-results.txt` (o la consola): la primera línea es `ok N / falla M`
  y después una línea por falla.

## Comandos y archivos nuevos

### `build/test.ps1` (nuevo)

```powershell
powershell -ExecutionPolicy Bypass -File build\test.ps1 [-TimeoutSeconds 300] [-Visible] [-Results <ruta>]
```

- Arma en memoria un libro con el mismo código que `build.ps1` más `src/tests/`, corre las pruebas y no guarda
  el libro en disco.
- Escribe `dist/test-results.txt` (ignorado por Git).
- Códigos de salida: `0` = hubo al menos una prueba y ninguna falló; `1` = cualquier otro caso.
- Si falta `src/modules/modSecrets.bas`, usa `modSecrets.example.bas` sin crear archivos en `src/`.
- Activa temporalmente "Confiar en el acceso al modelo de objetos de proyectos de VBA" y lo restaura al terminar.

### `build/lib.ps1` (nuevo, uso interno)

Funciones comunes de `build.ps1` y `test.ps1`. No es una interfaz para consumidores.

## Lo que no cambia

- `build/build.ps1`: mismos parámetros (`-Output`, `-Hidden`), mismo `.xlsm` y misma protección.
- El `.xlsm` entregado no contiene módulos de prueba (`modTest*`, `modAssert`).
- Formularios, PDF y apariencia siguen siendo prueba manual; las pruebas no los cubren.
