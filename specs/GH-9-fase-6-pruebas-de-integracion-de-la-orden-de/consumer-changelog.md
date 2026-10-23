# Changelog para consumidores · GH-9 · Pruebas de integración de la orden de servicio

## Cambios que rompen compatibilidad

Ninguno. Los datos guardados (`BD_*`, órdenes y cotizaciones) no cambian de formato ni de estructura para quien los usa.

## Cambios nuevos

### 1. Pie de página de los PDF sin total de páginas

- **Antes:** el pie decía «Página 1 de Cotiva-Cotizador.xlsm» (en Excel en español, el total salía como nombre del archivo).
- **Después:** el pie dice solo «Página N», con el número de la página. Aplica a los PDF de cotización y de orden.
- **Qué debe hacer el consumidor:** nada, salvo si su archivado o búsqueda depende del texto del pie.

### 2. Diagnóstico largo en el PDF de la orden

- **Antes:** el bloque del diagnóstico tenía un alto fijo; un texto largo se cortaba en el PDF.
- **Después:** el diagnóstico admite hasta 1000 caracteres y 40 renglones. El bloque crece con el texto y, si no cabe en una hoja carta, el PDF de la orden sale en dos hojas («Página 1» y «Página 2»). Los PDF con diagnóstico corto siguen en una hoja.
- **Al guardar:** un diagnóstico de 41 renglones o más no se guarda; el formulario avisa en la pestaña de servicio hasta que se acorte.
- **Qué debe hacer el consumidor:** si imprime o archiva la orden, debe aceptar que puede tener dos páginas. Si envía una orden con más de 40 renglones de diagnóstico, debe acortarlo antes de guardar.

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
