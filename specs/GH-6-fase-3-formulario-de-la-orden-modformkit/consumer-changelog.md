# Changelog para consumidores · GH-6 · Formulario de la orden de servicio

## Cambios que rompen compatibilidad

Ninguno. Las cotizaciones guardadas y los archivos del libro siguen abriendo igual.

## Cambios de comportamiento

### 1. Nuevo formulario «Orden de servicio»

- **Antes:** el libro solo tenía el formulario de cotización.
- **Después:** existe un formulario de orden de servicio que cabe en una pantalla de 1366×768. Tiene dos pestañas: «Cliente y vehículo» y «Entrada, servicio y entrega». El folio y el estado están siempre visibles.
- **Qué debe hacer el consumidor:** nada. Para abrirlo, use el botón «Orden de servicio» del selector (ver punto 2).
- **Obligatorios:** cliente, fecha de entrada, kilometraje y nivel de gasolina. Fecha prometida, fecha y hora de salida, hora de entrada y año pueden ir vacíos, pero si se escriben deben ser válidos.
- **Avisos que no bloquean:** el VIN que no tiene 17 caracteres se marca con una etiqueta junto al campo; se puede guardar igual.

### 2. Selector entre cotización y orden en `frmQuote` y en la orden

- **Antes:** la cotización no tenía forma de llegar a otro formulario.
- **Después:** junto a la etiqueta del formulario actual hay un botón que lleva al otro: «Orden de servicio» desde la cotización y «Cotización» desde la orden.
- Si el formulario actual tiene cambios sin guardar, aparece un aviso de Sí/No. Con «Sí» se descartan los cambios y se abre el otro formulario. Con «No» se queda donde está.
- **Qué debe hacer el consumidor:** guarde antes de cambiar de formulario si no quiere perder cambios.

### 3. Nombres de «Recibió» y «Entregó» sugeridos del historial

- **Antes:** los campos de personal no sugerían nombres.
- **Después:** las listas de «Recibió» y «Entregó» de la orden muestran los nombres ya guardados en órdenes anteriores, sin repetidos y en orden alfabético. También se puede escribir un nombre nuevo.
- **Qué debe hacer el consumidor:** nada. Si la hoja `BD_Ordenes` no existe, el libro la crea al preparar el libro.

### 4. Validación de fechas, horas, año y kilometraje

- Las fechas se validan con el mismo criterio en la cotización y en la orden.
- Las horas deben ir en formato `hh:mm`, de 00:00 a 23:59.
- El año del vehículo debe estar entre 1950 y el año actual + 1.
- El kilometraje solo admite números enteros.
- Las placas se guardan en mayúsculas.
