# Changelog para consumidores · GH-8 · PDF de la orden y vehículo en el PDF de la cotización

## Cambios que rompen compatibilidad

Ninguno. Las cotizaciones y las órdenes guardadas siguen abriendo igual, las hojas `BD_*` no cambian de formato y los PDF de cotización sin orden ligada salen igual que antes.

## Cambios nuevos

### 1. Botón «Generar PDF» en la orden de servicio

- **Antes:** la orden de servicio no tenía forma de imprimirse; solo se guardaba y se capturaba en pantalla.
- **Después:** la barra de la orden tiene el botón «Generar PDF». Antes de exportar:
  - valida la orden con las mismas reglas que «Guardar»; si no es válida, no genera nada;
  - si hay cambios sin guardar, pregunta si guardar: **Sí** guarda y exporta, **No** exporta con los datos de pantalla (sin guardarlos), **Cancelar** no hace nada;
  - si faltan datos de la empresa, avisa y pregunta si continuar, igual que en la cotización.
- **Qué debe hacer el consumidor:** nada. Si el usuario cancela o la exportación falla, no aparece archivo y no hay mensaje de éxito.

### 2. PDF de la orden de servicio

- Hoja nueva `PlantillaOrden` en el libro (visible y protegida, como `Plantilla`). Se crea sola al abrir libros anteriores.
- Formato carta, una página. Contenido: encabezado con el logo de la empresa, datos del cliente, vehículo, entrada y salida, servicio solicitado, diagnóstico, condiciones y tres firmas (Cliente, Recibió, Entregó).
- Nombre del archivo: `[Folio]_[Cliente].pdf`, por ejemplo `OS-0001_Juan Pérez.pdf`. Se guarda en la carpeta `Cotizaciones_PDF`, junto a los PDF de cotización.
- Límites del diseño: el servicio solicitado y el diagnóstico caben en un recuadro de alto fijo. Un texto más largo se corta en el PDF; el dato completo sigue en la orden.
- **Qué debe hacer el consumidor:** si imprime o archiva órdenes, el PDF nuevo sustituye a la impresión manual. Si el texto largo debe verse completo, resumirlo en la orden antes de generar el PDF.

### 3. Renglón «Vehículo» en el PDF de la cotización

- **Antes:** el renglón 7 del PDF de la cotización era un espacio vacío de 10 pt.
- **Después:** si la cotización está ligada a una orden con vehículo o placas, el renglón 7 muestra la etiqueta «Vehículo» y, a su lado, el valor, por ejemplo `Lincoln MKX 2010 · Placas ABC123`. El renglón mide 18 pt.
- Si la orden tiene solo placas, el valor es `Placas ABC123`. Si la orden no tiene ni vehículo ni placas, o la cotización no tiene orden, el renglón queda vacío y en 10 pt, como antes.
- Las filas de conceptos y totales no se mueven.
- **Qué debe hacer el consumidor:** nada.

### 4. Macro nueva `DesignOrderTemplate`

- Rediseña la hoja `PlantillaOrden`. Conserva el logo de la Plantilla.
- **Qué debe hacer el consumidor:** si el taller editó a mano las condiciones de la orden, la macro las reemplaza por el texto original y lo advierte antes de hacerlo. Guardar el texto propio antes de correrla.

## Sin cambio para el consumidor

- Los campos de la orden y de la cotización, los formatos de folio (`COT-`, `OS-`) y las hojas `BD_*` no cambian.
- El PDF de una cotización sin orden ligada no cambia.
- La hoja `Plantilla` y la macro `DesignTemplate` siguen igual.
