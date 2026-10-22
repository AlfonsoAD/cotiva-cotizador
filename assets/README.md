# Logos para el build (opcionales)

| Archivo | Uso | Formato |
|---|---|---|
| `logo-cotiva.png` | Logo de Cotiva: arriba en la pantalla de **Inicio** y al pie de la cotización | PNG horizontal, fondo transparente |
| `logo-innobytes.png` | Logo de **innobytes** al pie de la pantalla de Inicio | PNG horizontal, fondo transparente |
| `logo-form.jpg` | Logo en la barra superior del formulario; es el único nombre aceptado | **JPG** (el control Image de VBA no carga PNG), de preferencia con fondo blanco |

Si un archivo no existe, el build lo omite y se usa el texto de respaldo: la Plantilla escribe
«Cotiva», la pantalla de Inicio escribe «Cotiva» o «innobytes», y el formulario usa la barra del color secundario de «Mi empresa» (azul tinta de Cotiva si está vacío)
con texto. El logo de cada empresa que cotiza lo pone el usuario con el botón «Mi logo».

Los logos quedan guardados **dentro** del `.xlsm`: la carpeta `assets/` solo se necesita al construir.
En un libro ya generado se cambian con las macros `InsertCotivaLogo`, `InsertHomeCotivaLogo` e
`InsertInnobytesLogo` (Alt+F8).

Un `logo-form.jpg` junto al `.xlsm` se usa si el libro no trae imagen incrustada. Antes se llamaba
`logo_cotizador.jpg`: quien lo tenga junto a su libro debe renombrarlo, o el formulario deja de verlo.

Los colores del formulario, y de los que se agreguen, salen de «Mi empresa» (B11 principal, B12
secundario) con los de Cotiva de respaldo. En el código se usan `PrimaryColor()` y `SecondaryColor()`
de `modCompany`, nunca `BrandOrange()` ni `BrandInk()` directo.
