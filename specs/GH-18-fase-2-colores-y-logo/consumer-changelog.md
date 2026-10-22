# Changelog para consumidores · GH-18 · Colores y logo del formulario

## Cambios que rompen compatibilidad

### 1. El logo del formulario se llama `logo-form.jpg` (antes `logo_cotizador.jpg`)

- **Antes:** el respaldo del logo del formulario se buscaba junto al libro como `logo_cotizador.jpg`. El build tomaba `assets/logo-form.jpg`.
- **Después:** tanto el build (`assets/logo-form.jpg`) como el respaldo junto al libro (`logo-form.jpg`) usan el mismo nombre. El nombre viejo se ignora.
- **Qué debe hacer el consumidor:** si tiene un `logo_cotizador.jpg` junto a su `.xlsm`, renómbralo a `logo-form.jpg`. Si no lo renombra, el formulario deja de mostrar ese logo y muestra la barra del color secundario con texto.
- **Formato:** solo JPG. El control Image de VBA no carga PNG.

Ejemplo:

```
Antes:   Cotizador/Cotiva-Cotizador.xlsm
         Cotizador/logo_cotizador.jpg     (ya no se usa)

Después: Cotizador/Cotiva-Cotizador.xlsm
         Cotizador/logo-form.jpg          (se usa de respaldo)
```

### 2. Cambia el comportamiento de los colores de `frmQuote`

- **Antes:** la barra, los botones, los totales y los encabezados del formulario usaban siempre los colores de Cotiva (naranja y azul tinta).
- **Después:** esos mismos elementos toman los colores de «Mi empresa»: B11 (principal) y B12 (secundario). Si B11 o B12 están vacías o no son un hexadecimal válido de 6 dígitos, se usan los colores de Cotiva.
- **Texto sobre color:** el texto de los botones y de la barra sin logo se elige para que se lea sobre el fondo (blanco o tinta), así que un color claro en B11/B12 cambia el color del texto.
- **Qué debe hacer el consumidor:** revise B11 y B12 en «Mi empresa». Si los dejó vacíos, el formulario se ve igual que antes. Si tiene colores configurados, el formulario ya los usa. Los colores se leen al abrir el formulario: si cambia B11 o B12 con el formulario abierto, cierre y vuelva a abrirlo para verlos.

## Sin cambio para el consumidor

- Los colores de la Plantilla y de la portada de Inicio (`modHome`) siguen con los colores de Cotiva.
- El crédito de Cotiva e innobytes no cambia.
- No cambian endpoints, campos, enums, errores ni permisos.
