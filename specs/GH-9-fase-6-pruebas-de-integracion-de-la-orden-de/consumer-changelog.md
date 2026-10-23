# Changelog para consumidores · GH-9 · Pruebas de integración de la orden de servicio

## Cambios que rompen compatibilidad

Ninguno. Los datos guardados (`BD_*`, órdenes y cotizaciones) no cambian de formato ni de estructura para quien los usa.

## Cambios nuevos

### 1. Pie de página de los PDF sin total de páginas

- **Antes:** el pie decía «Página 1 de Cotiva-Cotizador.xlsm» (en Excel en español, el total salía como nombre del archivo).
- **Después:** el pie dice solo «Página N», con el número de la página. Aplica a los PDF de cotización y de orden.
- **Qué debe hacer el consumidor:** nada, salvo si su archivado o búsqueda depende del texto del pie.

### 2. Diagnóstico largo en el PDF de la orden

- **Antes:** el bloque del diagnóstico medía 96 pt; un texto largo se cortaba en el PDF.
- **Después:** el bloque mide 200 pt, para que quepa el diagnóstico completo (hasta 1000 caracteres). La hoja sigue ajustada a una sola página carta, así que el PDF se reduce un poco cuando el texto es largo.
- **Qué debe hacer el consumidor:** nada. Los PDF de órdenes con diagnósticos cortos se ven casi igual, con letra un poco más pequeña si el texto es largo.

### 3. Crédito de innobytes en el formulario de la orden

- **Antes:** el crédito quedaba junto a los botones de la barra inferior y se cortaba.
- **Después:** el crédito tiene un renglón propio, arriba de la barra de botones. Los botones no cambian de lugar.
- **Qué debe hacer el consumidor:** nada.

### 4. Versión de la aplicación 9.0.0

- **Antes:** la hoja Inicio mostraba la versión 8.0.0.
- **Después:** la hoja Inicio muestra la versión 9.0.0, por el cambio de estructura de `BD_Ordenes`.
- **Qué debe hacer el consumidor:** si valida la versión del libro contra un número fijo, debe aceptar 9.0.0.

## Sin cambio para el consumidor

- Formato de los datos guardados, nombres de hojas, nombres de archivo de los PDF y permisos de las hojas protegidas.
