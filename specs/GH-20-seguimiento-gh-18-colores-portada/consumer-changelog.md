# Changelog para consumidores · GH-20 · Colores de empresa en portada y «Mi empresa»

## Cambios que rompen compatibilidad

Ninguno. No cambian endpoints, campos, enums, errores ni permisos. Si no hay colores capturados en «Mi empresa», la portada y la hoja se ven como antes.

## Cambios visibles

### 1. La portada (Inicio) usa los colores de «Mi empresa»

- **Antes:** la regla, el botón «Abrir cotizador», las viñetas y el texto de la portada usaban siempre los colores de Cotiva (naranja y azul tinta). El texto del botón era blanco.
- **Después:**
  - La regla, el botón y las viñetas usan el color principal (B11 de «Mi empresa»).
  - El texto del título y de las ventajas (B4 y B10:B12 de la portada) usa el color secundario (B12).
  - El texto del botón se elige para leerse sobre el color principal: blanco o tinta, según el fondo.
  - Si B11 o B12 están vacías o no son un hexadecimal válido de 6 dígitos, se usan los colores de Cotiva.
- **Cuándo se actualiza:** la portada se repinta al pulsar «Abrir cotizador». Si cambia B11 o B12 y solo abre «Mi empresa», la portada conserva los colores anteriores hasta pulsar «Abrir cotizador».
- **Qué debe hacer el consumidor:** nada, salvo revisar que el color secundario sea legible sobre blanco, porque el texto de la portada ahora lo usa.

### 2. La hoja «Mi empresa» usa los colores de la empresa

- **Antes:** el título (A1) era naranja de Cotiva y el texto capturado (B4:B12) era azul tinta.
- **Después:** el título (A1) usa el color principal y el texto capturado (B4:B12, incluidos B11 y B12) usa el color secundario. Las etiquetas y las notas siguen en gris.
- **Cuándo se actualiza:** al abrir «Mi empresa» y al pulsar «Abrir cotizador».
- **Qué debe hacer el consumidor:** nada. Lo capturado no cambia, solo su color.

## Sin cambio para el consumidor

- El crédito de Cotiva e innobytes en la portada conserva sus colores.
- Los logos, los textos capturados y las hojas BD_* no cambian.
- El formulario de cotización (`frmQuote`) no cambia respecto a GH-18.
