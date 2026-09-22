# Logos para el build (opcionales)

| Archivo | Uso | Formato |
|---|---|---|
| `logo-cotiva.png` | Logo de Cotiva: arriba en la pantalla de **Inicio** y al pie de la cotización | PNG horizontal, fondo transparente |
| `logo-innobytes.png` | Logo de **innobytes** al pie de la pantalla de Inicio | PNG horizontal, fondo transparente |
| `logo-form.jpg` | Logo en la barra superior del formulario | **JPG** (los formularios de VBA no aceptan PNG), fondo blanco |

Si un archivo no existe, el build lo omite y se usa el texto de respaldo: la Plantilla escribe
«Cotiva», la pantalla de Inicio escribe «Cotiva» o «innobytes», y el formulario usa la barra azul
tinta con texto. El logo de cada empresa que cotiza lo pone el usuario con el botón «Mi logo».

Los logos quedan guardados **dentro** del `.xlsm`: la carpeta `assets/` solo se necesita al construir.
En un libro ya generado se cambian con las macros `InsertCotivaLogo`, `InsertHomeCotivaLogo` e
`InsertInnobytesLogo` (Alt+F8).
