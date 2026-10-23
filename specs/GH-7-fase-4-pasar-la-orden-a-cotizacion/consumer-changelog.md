# Changelog para consumidores · GH-7 · Pasar la orden a cotización

## Cambios que rompen compatibilidad

Ninguno. Las cotizaciones y las órdenes guardadas siguen abriendo igual, y el libro no cambia de formato.

## Cambios nuevos

### 1. Botón «Pasar a cotización» en la orden de servicio

- **Antes:** la orden de servicio solo tenía «Guardar» y «Cotización» (regresa a `frmQuote`) en la barra superior.
- **Después:** la franja fija de la orden tiene el botón «Pasar a cotización», junto al estado. Guarda la orden con las mismas validaciones y confirmación que «Guardar», y abre su cotización en `frmQuote`.
- **Qué debe hacer el consumidor:** nada. Si la orden no es válida, el botón no crea nada y muestra el mismo aviso que «Guardar».

### 2. La cotización de una orden trae los renglones del diagnóstico

- Al pasar una orden por primera vez se crea una cotización con:
  - el cliente de la orden;
  - el folio de la orden en la cotización (`OrderNumber`), y el folio de la cotización en la orden (`QuoteNumber`);
  - un concepto por cada renglón del diagnóstico, con cantidad 1 y precio $0.00. Solo faltan los precios.
- Un renglón de más de 255 caracteres se corta a 255 (el máximo del campo Descripción).
- Si el diagnóstico tiene más de 17 renglones, solo los primeros 17 pasan a conceptos y se avisa cuántos quedaron fuera. El diagnóstico de la orden no cambia.
- **Qué debe hacer el consumidor:** revisar los precios de los conceptos antes de imprimir o enviar la cotización. Los renglones que no entraron a la cotización siguen en el diagnóstico de la orden.

### 3. Pasar una orden que ya tiene cotización

- Si la orden ya tiene cotización, se abre esa misma y no se crea otra. El diagnóstico no se vuelve a copiar, así que los renglones nuevos del diagnóstico no aparecen como conceptos.
- Si el folio de la cotización ligada ya no existe, se avisa y se crea una cotización nueva, que reemplaza el folio en la orden.

### 4. Franja de la orden en `frmQuote`

- Al cargar una cotización ligada a una orden, aparece a la derecha del título «Datos de la cotización» la franja de solo lectura, por ejemplo: «Orden OS-0001 · Toyota Corolla 2018 · Placas ABC123».
- Si la orden ligada no existe, la franja dice solo «Orden OS-0001».
- La franja no aparece en una cotización sin orden y desaparece al empezar una «Nueva cotización».
- Cargar una cotización (desde «Cargar» o el historial) pregunta antes de descartar cambios sin guardar, igual que antes.

### 5. Nuevo método público `frmQuote.OpenQuote(folio)`

- Abre una cotización ya guardada por su folio, como «Cargar». Regresa `True` si se abrió y `False` si no se encontró o si el usuario canceló por cambios sin guardar.
- **Qué debe hacer el consumidor:** si llama a `OpenQuote` desde código propio, debe mostrar el formulario después de llamarlo.

## Sin cambio para el consumidor

- Los campos, los formatos de folio (`COT-`, `OS-`) y las hojas `BD_*` no cambian.
- La cotización creada desde una orden se guarda y se imprime con la plantilla de siempre.
- Los textos y el diseño de las pantallas de cotización que no tocan la orden no cambian.
