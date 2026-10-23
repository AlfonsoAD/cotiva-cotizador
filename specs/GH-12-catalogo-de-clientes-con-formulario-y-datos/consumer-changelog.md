# Changelog para consumidores · GH-12 · Catálogo de clientes con datos fiscales

## Cambios que rompen compatibilidad

Ninguno. Los libros anteriores siguen funcionando: al abrirlos por primera vez se agregan las columnas y hojas nuevas, vacías.

## Cambios

### 1. `BD_Clientes` tiene 17 columnas (A:Q)

- **Antes:** A:F (`Cliente`, `FechaAlta`, `Telefono1`, `Telefono2`, `Correo`, `Direccion`).
- **Después:** A:Q. A:F no cambian de nombre ni de posición. G:Q son nuevas y opcionales:

| Columna | Encabezado | Contenido |
|---|---|---|
| G | `RFC` | RFC (12 o 13 caracteres, en mayúsculas). Solo avisa si el formato no cuadra. |
| H | `RazonSocial` | Nombre o denominación, como en la CSF |
| I | `RegimenFiscal` | Clave del catálogo del SAT, ej. `612` |
| J | `CPFiscal` | Código postal fiscal, 5 dígitos (texto, conserva el cero inicial) |
| K | `UsoCFDI` | Clave del catálogo del SAT, ej. `G03` |
| L | `Calle` | Domicilio fiscal |
| M | `NumExterior` | Domicilio fiscal |
| N | `NumInterior` | Domicilio fiscal |
| O | `Colonia` | Domicilio fiscal |
| P | `Municipio` | Municipio o alcaldía |
| Q | `Estado` | Estado |

- **Qué debe hacer el consumidor:** nada, si solo lee A:F. Si tiene integraciones que leen `BD_Clientes` por número de columna, no cambian. Si leen por encabezado, las columnas nuevas aparecen al final.
- **Formato:** I:K y M:N van con formato texto, para que las claves y números con cero a la izquierda (`01000`, `007`) no se pierdan al capturarlos.

### 2. Hojas ocultas nuevas `BD_RegimenFiscal` y `BD_UsoCFDI`

- Cada una tiene `Clave` (A) y `Descripcion` (B). Se siembran con los catálogos del SAT (CFDI 4.0): 19 regímenes fiscales y 24 usos de CFDI.
- Están ocultas y protegidas, igual que las demás `BD_*`.
- **Si la hoja ya tiene filas, no se reescribe.** Para actualizar un catálogo cuando cambie el SAT, edite esa hoja (con la macro de administrador `AdminUnlock` si está protegida). Una clave guardada que ya no esté en el catálogo sigue apareciendo en el combo.
- **Qué debe hacer el consumidor:** nada, salvo que tenga código propio que recorra las hojas ocultas.

### 3. Botón «Clientes» en la pantalla de Inicio

- Abre el nuevo formulario de clientes con la macro `OpenClientsTool`.
- En libros ya entregados, el botón aparece la primera vez que se abre el cotizador. Si ya existe un botón con ese nombre (`btnClientes`), no se duplica.

### 4. Botón «Clientes…» en la cotización y en la orden

- Junto al combo de cliente de la cotización y de la orden. Abre el mismo formulario y, al elegir un cliente, lo pone en el combo.

### 5. Formulario de clientes (`frmClients`)

- Buscar por nombre, teléfono, correo, RFC o razón social; dar de alta, editar y consultar.
- Dos pestañas: **Contacto** (teléfonos, correo, dirección) y **Datos fiscales** (CSF y domicilio fiscal). Todos los campos son opcionales.
- Avisos que no bloquean el guardado: RFC con formato inválido, código postal fiscal que no tiene 5 dígitos, y RFC que ya pertenece a otro cliente.
- Confirma antes de descartar cambios sin guardar.

## Cambia el comportamiento

### 6. En la orden, elegir un cliente del catálogo reemplaza los 4 campos de contacto

- **Antes:** al escribir el nombre, solo se llenaban los campos de contacto que estaban vacíos.
- **Después:** al elegir un cliente del catálogo, los campos Teléfono 1, Teléfono 2, Correo y Dirección se reemplazan con los del catálogo, incluso si quedan vacíos. Un teléfono escrito a mano se pierde si el cliente del catálogo no tiene teléfono.
- **Qué debe hacer el consumidor:** si quiere conservar un dato escrito a mano, cópielo antes de elegir el cliente del catálogo.

### 7. Las hojas de catálogo y `BD_Clientes` entran en la protección del libro

- `ProtectAll` y `AdminUnlock` incluyen `BD_RegimenFiscal` y `BD_UsoCFDI`.

## Sin cambio para el consumidor

- La cotización, la orden y los PDF no cambian en su salida.
- La estructura de `BD_Cotizaciones`, `BD_Detalle` y `BD_Ordenes` no cambia.
- No hay endpoints, permisos ni errores nuevos.
- No hay timbrado ni emisión de CFDI: esta versión solo guarda los datos fiscales.
