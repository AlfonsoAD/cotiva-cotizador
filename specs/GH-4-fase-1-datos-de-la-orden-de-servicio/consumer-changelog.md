# Changelog para consumidores · GH-4 Datos de la orden de servicio

Este cambio es aditivo: no quita ni renombra columnas, hojas ni folios existentes. No hay endpoints ni API que cambien. Lo que sigue describe la estructura del libro de datos (`.xlsm`) para quien lo lea desde fuera del módulo `modData`.

## Compatibilidad

- Ningún cambio rompe compatibilidad. Las columnas existentes conservan su letra y su nombre.
- Los libros anteriores siguen abriendo; al abrirlos se agregan los encabezados y la hoja nuevos sin tocar los datos.

## Cambios

### 1. `BD_Cotizaciones`: columna nueva I `OrdenFolio`

- Antes: `A:H` (Folio, Fecha, Cliente, Subtotal, IVA, Total, TasaIVA, Vigencia).
- Después: `A:I`. La columna I guarda el folio de la orden ligada (ej. `OS-0001`).
- Las cotizaciones anteriores tienen la celda I vacía. Es válido.
- Consumidor: si lees por nombre de encabezado, no necesitas cambios. Si lees por posición, la columna I es nueva y puede venir vacía.

### 2. `BD_Clientes`: columnas C:F de contacto

- Antes: `A:B` (Cliente, FechaAlta).
- Después: `A:F` (Cliente, FechaAlta, Telefono1, Telefono2, Correo, Direccion).
- Los clientes existentes quedan con C:F vacías hasta que se guarde una orden o se capture el contacto.
- Un campo vacío no borra un dato ya guardado.
- Consumidor: si lees por nombre de encabezado, no necesitas cambios.

### 3. `BD_Ordenes`: hoja nueva (oculta y protegida)

- Una fila por orden de servicio. Fila 1 = encabezados. Columnas `A:Z`:

| Col | Encabezado | Col | Encabezado |
|---|---|---|---|
| A | Number (folio, `OS-0001`) | N | Mileage (km) |
| B | Status | O | FuelLevel |
| C | ClientName | P | Vin |
| D | Phone1 | Q | InDate |
| E | Phone2 | R | InTime |
| F | Email | S | ReceivedBy |
| G | Address | T | PromisedDate |
| H | Make | U | OutDate |
| I | Model | V | OutTime |
| J | Engine | W | DeliveredBy |
| K | Year | X | Requested |
| L | Color | Y | Diagnosis |
| M | Plates | Z | QuoteNumber (cotización ligada) |

- Estados válidos (`Status`): `Recibida`, `En diagnóstico`, `En reparación`, `Lista`, `Entregada`, `Cancelada`.
- Niveles de combustible (`FuelLevel`): `R`, `¼`, `½`, `¾`, `Lleno`.
- Placas y VIN se guardan en mayúsculas. Teléfonos, placas y VIN son texto (conservan el 0 inicial).
- `QuoteNumber` se guarda sin espacios y en mayúsculas (ej. `COT-1001`).
- Fechas y horas vacías se guardan como celda vacía. `Year` = 0 significa "sin año".
- Consumidor: la hoja está oculta y protegida. Si necesitas leerla desde otra herramienta, hazlo por encabezado. Un folio se repite como máximo una vez; al guardar de nuevo, la fila anterior se reemplaza.

### 4. Folios

- Nuevo prefijo `OS-` con 4 dígitos (`OS-0001`, `OS-0002`...).
- Los folios `COT-` no cambian de numeración (`COT-1001`, `COT-1002`...).

## Qué tiene que hacer el consumidor

- Si tu lectura del libro usa nombres de encabezado: nada.
- Si tu lectura usa posiciones fijas en `BD_Clientes` o `BD_Cotizaciones`: revisa que no asuma que el número de columnas es 2 u 8. Las columnas nuevas van al final.
- Si quieres leer órdenes, lee `BD_Ordenes` por encabezado. No asumas que hay una orden por cotización: una cotización guarda su `OrdenFolio` y una orden guarda su `QuoteNumber`, pero cada lado se escribe por separado.
