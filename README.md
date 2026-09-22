# Cotiva · Cotizador para Excel

Herramienta de cotizaciones en Excel (VBA) de **Cotiva**, un producto de [innobytes](https://innobytes.tech).
Genera cotizaciones con folio automático, catálogo de clientes, IVA 16% / 8% / sin IVA, vigencia,
datos y colores de la empresa emisora, y exporta a PDF.

El código fuente vive en `src/` como texto. El archivo `.xlsm` **no se edita a mano**: se genera con
`build/build.ps1`.

## Estructura

```
src/
  modules/     Módulos estándar (.bas)
    modApp         Punto de entrada (OpenQuoteTool) y preparación del libro
    modConfig      Constantes: hojas, celdas, límites, datos de innobytes
    modStrings     Todos los textos visibles (español)
    modTheme       Colores de marca, tipografía, utilidades de color
    modUtils       Validaciones, formatos, importe con letra, errores
    modData        Capa de datos (lo único a cambiar para migrar a SQLite/API)
    modCompany     Hoja "Mi empresa": datos del emisor, vigencia, colores
    modTemplate    Plantilla: diseño, llenado, logos y PDF
    modHome        Hoja "Inicio": portada, botón y logos de Cotiva e innobytes
    modSecurity    Protección del libro y macros de administrador
    modSecrets.example.bas   Plantilla de la clave (modSecrets.bas NO se sube)
  classes/     Quote, QuoteItem
  forms/       frmQuote.vba (solo código; la interfaz se construye por código)
  build/       modBuild.bas (arma el libro; se elimina del .xlsm final)
build/
  build.ps1    Genera dist/Cotiva-Cotizador.xlsm
  export.ps1   Exporta el código de un .xlsm de vuelta a src/
assets/        Logos opcionales para el build (ver assets/README.md)
dist/          Salida del build (ignorada por Git)
```

## Primer uso

1. Copia `src/modules/modSecrets.example.bas` como `src/modules/modSecrets.bas` y pon tu clave.
2. (Opcional) Pon los logos en `assets/` según `assets/README.md`.
3. Genera el libro:

   ```powershell
   powershell -ExecutionPolicy Bypass -File build\build.ps1
   ```

4. Abre `dist/Cotiva-Cotizador.xlsm`, habilita macros y presiona **Abrir cotizador**.

> El build abre Excel visible. Si aparece una ventana de error de VBA, significa que hay un error de
> compilación: toma captura, cierra Excel y corrige en `src/`.

## Flujo de trabajo

**Opción A (recomendada): editar en `src/`** (VS Code / Claude Code) → `build.ps1` → probar en Excel → commit.

**Opción B: editar en el editor de VBA** → `export.ps1` → revisar el diff → commit.

`.gitattributes` convierte los archivos de código a UTF-8 en el repositorio y a Windows-1252 con CRLF
en tu carpeta de trabajo, que es lo que el editor de VBA necesita para importar acentos correctamente.

## Distribuir

**Al cliente se le entrega un solo archivo: `dist/Cotiva-Cotizador.xlsm`.** No necesita PowerShell,
ni la carpeta `src/`, ni los logos: el código, las hojas y las imágenes ya están dentro del `.xlsm`.

`build.ps1` es una herramienta de desarrollo: se ejecuta cuando **tú** cambias el código en `src/`,
para regenerar el archivo. El cliente nunca lo corre.

> El `.xlsm` que entregas es también el que guarda las cotizaciones de ese cliente. Al publicar una
> versión nueva, cada quien conserva su archivo hasta que migre sus datos (ver más abajo); no
> sobrescribas el archivo de un cliente con el `.xlsm` recién construido.

## Antes de distribuir (pasos manuales)

1. **Contraseña del proyecto VBA**: Alt+F11 → Herramientas → Propiedades de VBAProject → Protección.
   Guarda, cierra y vuelve a abrir.
2. **Firma digital** con un certificado de firma de código. Sin firma, Windows bloquea las macros de
   archivos descargados de internet y el cliente tendría que desbloquear el archivo en sus propiedades.
3. Sube el `.xlsm` como adjunto de un *Release* de GitHub, no al repositorio.

## Seguridad: qué protege y qué no

| Medida | Protege contra |
|---|---|
| Hojas protegidas y ocultas (`modSecurity`) | Cambios accidentales del usuario común |
| Contraseña del proyecto VBA | Que se lea o modifique el código (se puede quitar con herramientas) |
| Firma digital | Bloqueo de macros y alteración del archivo firmado |
| Funciones en el servidor de Cotiva (futuro) | Piratería real: lo que no está en el archivo no se puede copiar |

La protección de Excel **no es cifrado**. `modSecrets.bas` evita que la clave llegue al historial de
Git, pero dentro del `.xlsm` solo la protege la contraseña del proyecto VBA.

## Macros de administrador (Alt+F8)

| Macro | Uso |
|---|---|
| `AdminUnlock` | Pide la clave y deja todas las hojas visibles y editables |
| `AdminLock` | Vuelve a proteger todo |
| `DesignTemplate` | Rediseña la Plantilla (conserva logos) |
| `DesignHome` | Rediseña la pantalla de Inicio (conserva logos) |
| `InsertCotivaLogo` | Cambia el logo de Cotiva del pie de la cotización |
| `InsertHomeCotivaLogo` | Cambia el logo de Cotiva de la pantalla de Inicio |
| `InsertInnobytesLogo` | Cambia el logo de innobytes de la pantalla de Inicio |

**Cambiar la clave** de un libro ya protegido: `AdminUnlock` con la clave vieja → cambia
`APP_PASSWORD` → `AdminLock`. Si cambias la constante sin desproteger antes, el libro queda bloqueado.

## Migrar datos de una versión anterior

Las hojas de datos conservan los mismos nombres y columnas. En el archivo viejo y en el nuevo ejecuta
`AdminUnlock`, copia las filas de `BD_Cotizaciones`, `BD_Detalle` y `BD_Clientes` (sin encabezados)
y ejecuta `AdminLock` en el nuevo.
