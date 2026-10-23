# GH-12 · Catálogo de clientes con formulario y datos fiscales opcionales (CSF)

## Brief

**Objetivo**: un formulario `frmClients` para buscar, dar de alta, editar y consultar clientes de `BD_Clientes`, con sus datos fiscales de la CSF (Constancia de Situación Fiscal), todos opcionales.

**Entra**
- 11 columnas fiscales nuevas en `BD_Clientes` (G:Q), agregadas por `EnsureDataHeaders`; los libros viejos siguen funcionando con ellas vacías.
- Hojas ocultas `BD_RegimenFiscal` y `BD_UsoCFDI` (clave + descripción), sembradas con los catálogos del SAT vigentes.
- Clase `Client` y funciones de `modData` para leer, guardar, buscar y detectar RFC repetido.
- `frmClients` (lista con buscador + pestañas Contacto / Datos fiscales), botón "Clientes…" en `frmQuote` y `frmOrder`, y botón "Clientes" en Inicio.
- Avisos que no bloquean: RFC con formato inválido, CP fiscal distinto de 5 dígitos, RFC que ya tiene otro cliente.
- Pruebas automáticas de la capa de datos; README.

**No entra**: timbrado o emisión de CFDI, eliminar o renombrar clientes, validar contra el SAT en línea, catálogo de estados o colonias (son texto libre).

**Decisiones nuevas**
- [N] Clase `Client` (modelo, como `ServiceOrder`) para mover los 17 campos. Descartado: parámetros `ByRef` sueltos, como `GetClientContact` (17 parámetros).
- [N] Las semillas de los catálogos son constantes en `modConfig` y se escriben al crear la hoja (en el build y en libros viejos, desde `EnsureWorkbookStructure`). Si la hoja ya existe con datos, no se toca: ahí la actualiza el administrador. Descartado: un CSV en `assets/` que solo leería el build, porque los libros ya entregados nunca tendrían los catálogos.
- [N] `frmClients` siempre se abre modal. Descartado: sin modal desde Inicio, porque un formulario ya abierto sin modal truena al volver a abrirlo modal desde la cotización.
- [N] En el formulario, guardar escribe los 15 campos editables tal como están (un campo vacío borra el dato). `SaveOrder` sigue usando `UpsertClientContact`, que no pisa con vacíos. Descartado: no pisar con vacíos también desde el formulario, porque no habría manera de borrar un dato.
- [N] Al elegir un cliente desde la orden se reemplazan los 4 campos de contacto con los del catálogo. Descartado: llenar solo los vacíos (como hoy al escribir el nombre), porque dejaría el teléfono del cliente anterior.
- [N] En libros ya entregados, el botón de Inicio lo agrega `PrepareWorkbook` cuando falta (la primera vez que se abre el cotizador). Descartado: pedir que se corra `DesignHome`.
- [N] `APP_VERSION` pasa a 9.1.0.

**Riesgos**
- El CP y las claves con cero a la izquierda ("01000") se pierden si la celda no tiene formato texto: van con `NumberFormat = "@"`.
- El catálogo del SAT cambia: la semilla queda vieja con el tiempo, pero la hoja se puede actualizar sin tocar código. Una clave guardada que ya no esté en el catálogo se sigue mostrando.
- Achicar el combo de cliente en `frmQuote` y `frmOrder` para que quepa el botón.

**Tamaño**: L (un formulario nuevo de ~600 líneas, capa de datos, 2 hojas, 3 integraciones, pruebas).

## Discovery

# Discovery GH-12 · Catálogo de clientes con datos fiscales opcionales

Rama base: auto-repair-shop. Dependencias #4 y #6 cerradas; BD_Clientes ya tiene A:F (Cliente, FechaAlta, Telefono1, Telefono2, Correo, Direccion).

## Decisiones
1. Catálogos del SAT (Régimen fiscal y Uso de CFDI): hojas ocultas `BD_RegimenFiscal` y `BD_UsoCFDI` (clave + descripción), sembradas por el build con los catálogos vigentes del SAT y actualizables sin tocar código; modData es quien las lee. Los combos muestran clave + descripción y guardan la clave.
2. Llave del cliente: sigue siendo el nombre (como hoy en ClientExists y los combos). El RFC es opcional; si ya lo tiene otro cliente, solo se avisa.
3. Integración con cotización y orden: botón "Clientes…" junto al combo de cliente en frmQuote y frmOrder; abre frmClients posicionado en el cliente actual. Al cerrar, el cliente elegido o dado de alta queda en el combo y se recarga su contacto. También hay botón de acceso desde la hoja de Inicio (modHome).
4. Alcance del formulario: buscar, dar de alta, editar contacto y datos fiscales, y consultar. No se eliminan ni se renombran clientes (el nombre se usa como referencia de texto en cotizaciones y órdenes guardadas).
5. Maquetado (1366×768): lista con buscador a la izquierda; a la derecha un MultiPage con páginas "Contacto" y "Datos fiscales". Construido por código con modFormKit, sin nada en el diseñador.
6. Validación: todo opcional y solo avisa (no bloquea): RFC con formato inválido (reusar IsValidRfc), CP fiscal que no sea de 5 dígitos, RFC duplicado con otro cliente.

## Datos fiscales (columnas nuevas en BD_Clientes, agregadas con EnsureDataHeaders)
RFC, RazonSocial, RegimenFiscal (clave), CPFiscal, UsoCFDI (clave), Calle, NumExterior, NumInterior, Colonia, Municipio (o alcaldía), Estado. Libros viejos siguen funcionando con esas columnas vacías.

## Restricciones
- Genérico para cualquier taller (y llevable a dev): nada propio de un cliente en el código.
- Textos de UI como constantes en modStrings; solo modData toca hojas BD_*; build/modBuild.bas debe crear frmClients.
- Pruebas automáticas en src/tests para la capa de datos (encabezados, alta/edición de datos fiscales, lectura de catálogos, avisos de validación); el formulario se verifica a mano.

## Fuera de alcance
Timbrado o emisión de CFDI; eliminar o renombrar clientes; validación contra el SAT en línea.

## Requirements

**Datos y estructura**
- R1 [D] CUANDO corre `EnsureDataHeaders`, el sistema DEBE dejar en `BD_Clientes` G1:Q1 los encabezados `RFC`, `RazonSocial`, `RegimenFiscal`, `CPFiscal`, `UsoCFDI`, `Calle`, `NumExterior`, `NumInterior`, `Colonia`, `Municipio`, `Estado`, sin tocar A1:F1 ni las filas de datos.
- R2 [D] CUANDO se abre un libro de una versión anterior (sin G1:Q1 ni hojas de catálogo), `PrepareWorkbook` DEBE regresar True, agregar los encabezados y las hojas, y los clientes existentes DEBEN leerse con sus datos de contacto intactos y los fiscales vacíos.
- R3 [D] CUANDO `EnsureWorkbookStructure` no encuentra `BD_RegimenFiscal` o `BD_UsoCFDI`, el sistema DEBE crearla con encabezados `Clave | Descripcion` y sembrarla con el catálogo del SAT de `modConfig`.
- R4 [N] CUANDO la hoja de catálogo ya existe y tiene al menos una fila de datos, el sistema NO DEBE reescribirla; SI existe sin filas de datos, DEBE sembrarla.
- R5 [D] `ProtectAll` DEBE dejar `BD_RegimenFiscal` y `BD_UsoCFDI` protegidas y, con `HIDE_DATA_SHEETS`, `xlSheetVeryHidden`, igual que las demás `BD_*`; `AdminUnlock` DEBE mostrarlas.
- R6 [D] `ListTaxRegimes` y `ListCfdiUses` DEBEN regresar las filas no vacías de su hoja como arreglo `(0 To n-1, 0 To 1)` (clave, descripción) en el orden de la hoja, o Empty si no hay.
- R7 [N] `LoadClient(nombre)` DEBE regresar un `Client` con los 17 campos de la fila cuyo nombre coincida (sin distinguir mayúsculas ni espacios de los extremos), o Nothing si no existe.
- R8 [N] CUANDO se llama `SaveClient` con un nombre que no existe, el sistema DEBE agregar la fila con `FechaAlta` = hoy; CUANDO el nombre ya existe, DEBE reescribir C:Q de esa fila (vacíos incluidos) y conservar A:B. Regresa False sin escribir si el nombre está vacío.
- R9 [N] `SaveClient` DEBE guardar el RFC sin espacios de los extremos y en mayúsculas, y RFC, CPFiscal, RegimenFiscal, UsoCFDI, NumExterior, NumInterior y teléfonos como texto (un "01000" se lee "01000"); los textos libres pasan por `SafeCellText`.
- R10 [N] `UpsertClientContact` (lo usa `SaveOrder`) NO DEBE borrar ni cambiar las columnas G:Q.
- R11 [D] `FindClientByRfc(rfc, nombreExcluido)` DEBE regresar el nombre del primer otro cliente con ese RFC (sin distinguir mayúsculas), o "" si no hay o si el RFC está vacío.
- R12 [N] `SearchClients(texto)` DEBE regresar, en orden alfabético, los nombres cuyo nombre, RFC, razón social, teléfonos o correo contengan el texto (sin distinguir mayúsculas); con texto vacío, lo mismo que `ListClients`; Empty si nada coincide.
- R13 [D] `IsValidPostalCode(texto)` DEBE regresar True solo con exactamente 5 dígitos.

**Formulario**
- R14 [D] `frmClients` DEBE construirse por código con `modFormKit` (nada en el diseñador), caber en 1366×768 y tener: buscador y lista de clientes a la izquierda; a la derecha el nombre y un MultiPage con las páginas "Contacto" y "Datos fiscales".
- R15 [D] CUANDO el usuario escribe en el buscador, la lista DEBE filtrarse con `SearchClients`; CUANDO elige un cliente de la lista, el formulario DEBE mostrar sus datos.
- R16 [D] CUANDO el usuario pulsa "Nuevo", el formulario DEBE limpiar los campos y habilitar el nombre; CUANDO muestra un cliente guardado, el nombre DEBE quedar bloqueado (no se renombra) y no DEBE haber forma de eliminarlo.
- R17 [N] CUANDO el usuario guarda un cliente nuevo cuyo nombre ya existe, el sistema NO DEBE crear otro: DEBE avisar y mostrar el existente. CUANDO el nombre está vacío, DEBE avisar y no guardar.
- R18 [D] Los combos de régimen fiscal y uso de CFDI DEBEN mostrar "clave - descripción", guardar solo la clave y permitir dejarlos vacíos. CUANDO la clave guardada no está en el catálogo, DEBEN mostrarla de todos modos.
- R19 [D] CUANDO el RFC no está vacío y no pasa `IsValidRfc`, o el CP fiscal no está vacío y no pasa `IsValidPostalCode`, el formulario DEBE mostrar un aviso junto al campo sin impedir guardar.
- R20 [D] CUANDO se guarda un RFC que ya tiene otro cliente, el sistema DEBE guardar y avisar con el nombre de ese cliente.
- R21 [N] CUANDO hay cambios sin guardar y el usuario cierra, elige otro cliente o pulsa "Nuevo", el formulario DEBE pedir confirmación.

**Accesos**
- R22 [D] Inicio DEBE tener un botón "Clientes" que prepare el libro (`PrepareWorkbook`) y abra `frmClients`; en un libro que no lo tenga, `PrepareWorkbook` DEBE agregarlo.
- R23 [D] `frmQuote` y `frmOrder` DEBEN tener un botón "Clientes…" junto al combo de cliente que abra `frmClients` posicionado en el cliente escrito (si existe).
- R24 [D] CUANDO el usuario pulsa "Usar este cliente", el formulario DEBE cerrarse y el formulario que lo abrió DEBE recargar su lista de clientes y poner ese nombre en el combo; en `frmOrder` DEBE además reemplazar teléfonos, correo y dirección con `GetClientContact`. CUANDO se cierra sin elegir, el combo DEBE conservar lo que tenía y solo recargar la lista.
- R25 [D] Los textos de la interfaz nuevos DEBEN ser constantes de `modStrings`; nada propio de un taller en el código.

## Design

### Estructura de `BD_Clientes` (A:Q, 17 columnas)

| Col | Encabezado | Campo de `Client` | Formato |
|---|---|---|---|
| A | Cliente | `Name` | texto (llave, no cambia) |
| B | FechaAlta | `CreatedOn` (Date) | dd/mm/yyyy |
| C:F | Telefono1, Telefono2, Correo, Direccion | `Phone1`, `Phone2`, `Email`, `Address` | C:D texto |
| G | RFC | `Rfc` | texto, mayúsculas |
| H | RazonSocial | `LegalName` | |
| I | RegimenFiscal | `TaxRegime` (clave, p. ej. "612") | texto |
| J | CPFiscal | `TaxZip` | texto |
| K | UsoCFDI | `CfdiUse` (clave, p. ej. "G03") | texto |
| L | Calle | `Street` | |
| M | NumExterior | `ExtNumber` | texto |
| N | NumInterior | `IntNumber` | texto |
| O | Colonia | `Neighborhood` | |
| P | Municipio | `Municipality` (municipio o alcaldía) | |
| Q | Estado | `State` | |

`CLIENT_COLUMNS` pasa de 6 a 17. El comentario de cabecera de `modData` se actualiza.

### Hojas de catálogo
- `modConfig`: `SHEET_TAX_REGIMES = "BD_RegimenFiscal"`, `SHEET_CFDI_USES = "BD_UsoCFDI"`; A:B `Clave | Descripcion`, claves con formato texto.
- Semillas en `modConfig` como constantes con el formato de `FUEL_LEVELS`: entradas separadas por `|`, clave y descripción por `=`. Se pueden partir en varias constantes concatenadas si una línea pasa de ~1000 caracteres.
  - `SAT_TAX_REGIMES_SEED` (c_RegimenFiscal, CFDI 4.0): 601 General de Ley Personas Morales; 603 Personas Morales con Fines no Lucrativos; 605 Sueldos y Salarios e Ingresos Asimilados a Salarios; 606 Arrendamiento; 607 Régimen de Enajenación o Adquisición de Bienes; 608 Demás ingresos; 610 Residentes en el Extranjero sin Establecimiento Permanente en México; 611 Ingresos por Dividendos (socios y accionistas); 612 Personas Físicas con Actividades Empresariales y Profesionales; 614 Ingresos por intereses; 615 Régimen de los ingresos por obtención de premios; 616 Sin obligaciones fiscales; 620 Sociedades Cooperativas de Producción que optan por diferir sus ingresos; 621 Incorporación Fiscal; 622 Actividades Agrícolas, Ganaderas, Silvícolas y Pesqueras; 623 Opcional para Grupos de Sociedades; 624 Coordinados; 625 Régimen de las Actividades Empresariales con ingresos a través de Plataformas Tecnológicas; 626 Régimen Simplificado de Confianza.
  - `SAT_CFDI_USES_SEED` (c_UsoCFDI, CFDI 4.0): G01 Adquisición de mercancías; G02 Devoluciones, descuentos o bonificaciones; G03 Gastos en general; I01 Construcciones; I02 Mobiliario y equipo de oficina por inversiones; I03 Equipo de transporte; I04 Equipo de cómputo y accesorios; I05 Dados, troqueles, moldes, matrices y herramental; I06 Comunicaciones telefónicas; I07 Comunicaciones satelitales; I08 Otra maquinaria y equipo; D01 Honorarios médicos, dentales y gastos hospitalarios; D02 Gastos médicos por incapacidad o discapacidad; D03 Gastos funerales; D04 Donativos; D05 Intereses reales efectivamente pagados por créditos hipotecarios (casa habitación); D06 Aportaciones voluntarias al SAR; D07 Primas por seguros de gastos médicos; D08 Gastos de transportación escolar obligatoria; D09 Depósitos en cuentas para el ahorro, primas que tengan como base planes de pensiones; D10 Pagos por servicios educativos (colegiaturas); S01 Sin efectos fiscales; CP01 Pagos; CN01 Nómina.
- `modApp.EnsureWorkbookStructure`: si falta la hoja la crea (`AddSheet`); después llama `EnsureSatCatalogs` (modData), que pone encabezados y siembra las que no tienen filas de datos (R3, R4). El build pasa por aquí, así que también siembra.
- `modSecurity.ProtectAll` y `AdminUnlock`: las dos hojas se suman a las listas de hojas `BD_*`. `modTestRunner.ResetDataSheets` NO las vacía.

### Contrato público

```vba
' src/classes/Client.cls — solo modelo, sin hojas. Campos Public como ServiceOrder.
Public Name As String, CreatedOn As Date   ' 0 = sin fecha
Public Phone1 As String, Phone2 As String, Email As String, Address As String
Public Rfc As String, LegalName As String, TaxRegime As String, TaxZip As String, CfdiUse As String
Public Street As String, ExtNumber As String, IntNumber As String
Public Neighborhood As String, Municipality As String, State As String

' modData
Public Sub EnsureSatCatalogs()
Public Function ListTaxRegimes() As Variant
Public Function ListCfdiUses() As Variant
Public Function LoadClient(ByVal clientText As String) As Client
Public Function SaveClient(ByVal c As Client) As Boolean
Public Function FindClientByRfc(ByVal rfc As String, ByVal exceptClient As String) As String
Public Function SearchClients(ByVal searchText As String) As Variant

' modUtils
Public Function IsValidPostalCode(ByVal zipText As String) As Boolean

' modApp
Public Sub OpenClientsTool()                ' Macro del botón de Inicio
Public Function PickClient(ByVal currentClient As String, _
        Optional ByVal pickMode As Boolean = True) As String   ' "" si no eligió

' modHome
Public Sub EnsureHomeClientsButton()   ' Lo agrega si falta; lo llama PrepareWorkbook

' frmClients
Public Sub Start(ByVal currentClient As String, ByVal pickMode As Boolean)
Public Property Get ChosenClient() As String
```

- `SaveClient`: `Nothing` → error (como `SaveOrder`, prueba con `AssertFails`). Busca la fila por nombre como `ClientExists`; si no está, la agrega al final escribiendo A:B como `AddClient`. Pone `NumberFormat = "@"` en C:D, G, I:K, M:N de la fila antes de escribir.
- `SearchClients` reutiliza el ordenamiento por inserción de `ListClients` (extraerlo a una `Private Function SortedNames` en vez de duplicarlo).
- `ListTaxRegimes`/`ListCfdiUses` comparten una `Private Function ReadCatalog(sheetName)`; si la hoja no existe regresan Empty.
- `modConfig.APP_VERSION = "9.1.0"`.

### Flujo de apertura
- `OpenClientsTool`: `If Not PrepareWorkbook() Then Exit Sub` → `PickClient "", False` (resultado ignorado).
- `PickClient(current, pickMode)`: `frmClients.Start current, pickMode` (posiciona la lista y muestra el cliente si existe; si no, queda en alta nueva con ese nombre escrito) → `frmClients.Show vbModal` → lee `ChosenClient` → `Unload frmClients`. `frmQuote` y `frmOrder` la llaman con `pickMode = True`.
- Dentro de `frmClients`, "Usar este cliente" (solo visible con `pickMode`) y "Cerrar" hacen `Me.Hide`; la X se intercepta en `QueryClose` (`Cancel = True` + confirmación de cambios + `Me.Hide`) para que el llamador pueda leer `ChosenClient` antes de descargarlo.
- "Usar este cliente" con cambios sin guardar: pregunta si guarda; si el cliente no está guardado (alta sin guardar), guarda primero o no permite usarlo.

### Maquetado de `frmClients` (760×496, como el alto de `frmQuote`)
- Barra superior y logo con el patrón de `frmOrder` (`PlaceHeaderLogo`; `modBuild.AddFormLogo "frmClients"` y `lib.ps1` crea el formulario: se suma `'frmClients'` a la lista `@('frmQuote', 'frmOrder')`).
- Izquierda (x = 18, ancho 230): `txtSearch` y `lstClients` (ListBox, alto ~330).
- Derecha (x = 260): etiqueta y `txtName` (MaxLength 100); `mpgClient` (`Forms.MultiPage.1`) con:
  - "Contacto": `txtPhone1`, `txtPhone2` (MaxLength 20), `txtEmail` (100), `txtAddress` (200).
  - "Datos fiscales": `txtRfc` (13, mayúsculas al escribir, `lblRfcHint`), `txtLegalName` (200), `cboTaxRegime` y `cboCfdiUse` (`fmStyleDropDownList`, `ColumnCount = 2`, `BoundColumn = 1`, primera fila vacía = sin dato), `txtTaxZip` (5, solo dígitos al teclear, `lblZipHint`), `txtStreet` (100), `txtExtNumber` (20), `txtIntNumber` (20), `txtNeighborhood` (100), `txtMunicipality` (100), `txtState` (50).
- Pie: `btnNew` "Nuevo", `btnSave` "Guardar" (primario), `btnUse` "Usar este cliente" (oscuro, solo `pickMode`), `btnClose` "Cerrar".
- Estado del formulario: `mIsDirty`, `mSuppressEvents`, `mIsSaved` como en `frmOrder`.
- `frmQuote`: `cboClient` pasa de ancho 400 a 316 y `btnClients` (secundario, 78×22) va en x = 438, y = 94. `frmOrder` (página Cliente): `cboClient` de 280 a 200 y `btnClients` (74×22) en x = 218, y = 38. Ambos con `ControlTipText = TIP_CLIENTS`.
- Inicio: `EnsureHomeClientsButton` agrega `HOME_CLIENTS_BUTTON_NAME = "btnClientes"` (rectángulo redondeado 160×46 en B7, a la derecha del de abrir, x = B7.Left + 212), `OnAction = "OpenClientsTool"`, color secundario; `DesignHome` lo llama tras `AddOpenButton` y `PaintHomeColors` lo repinta.

### Textos (`modStrings`)
Prefijo `UI_CLI_*` para etiquetas y botones, `MSG_CLI_*` para avisos, `TTL_CLIENTS`, `TIP_CLIENTS`, `HOME_CLIENTS_BUTTON`. Avisos: RFC con formato inválido, CP de 5 dígitos, RFC repetido (con `Fmt` y el nombre del otro cliente), nombre vacío, nombre ya existente, cambios sin guardar.

### Descartado
- Identificar al cliente por RFC: es opcional y varios clientes pueden no tenerlo; el nombre ya es la referencia en cotizaciones y órdenes (discovery 2).
- Secciones en una sola página en vez de MultiPage: no caben 15 campos más la lista en 496 de alto.
- Validar `UsoCFDI` contra el régimen (matriz de compatibilidad del SAT): eso es del timbrado, fuera de alcance.

### Seguridad
- Authz: sin cambios. Las hojas nuevas quedan protegidas y muy ocultas con la misma contraseña (`ProtectAll`); el usuario solo las ve con `AdminUnlock`.
- Validación de entrada: `MaxLength` en cada caja; RFC y CP se avisan, no se bloquean; todo texto libre va por `SafeCellText` (evita que "=..." se escriba como fórmula).
- Datos sensibles: RFC y domicilio fiscal son datos personales de clientes; se quedan dentro del libro, no se imprimen en la cotización ni en la orden en esta tarea, y no se escriben en `Debug.Print` ni en mensajes de error.
- Errores: cada evento de botón con `On Error GoTo Fail` → `ShowError "<procedimiento>"`, como `frmOrder`; si `SaveClient` regresa False el formulario avisa y conserva lo capturado.

## Tasks

- [x] T1 · Contrato. `src/classes/Client.cls` completa; en `modData`, `modUtils`, `modApp`, `modHome` las firmas públicas del Design con cuerpo mínimo que compila (regresan Empty / Nothing / False / ""); constantes `SHEET_TAX_REGIMES`, `SHEET_CFDI_USES`, `HOME_CLIENTS_BUTTON_NAME` en `modConfig`. Nombres de las pruebas que T2 y T3 escriben en `src/tests/modTestClients.bas` (T1 no agrega pruebas que fallen): `Test_Clientes_EncabezadosFiscales`, `Test_Clientes_LibroAnteriorSinFiscales`, `Test_Catalogos_SembradosYOcultos`, `Test_Catalogos_NoSeReescriben`, `Test_Catalogos_ListaClaveDescripcion`, `Test_SaveClient_AltaNueva`, `Test_SaveClient_EditaYBorraVacios`, `Test_SaveClient_CeroALaIzquierda`, `Test_SaveClient_RfcMayusculas`, `Test_SaveClient_NombreVacioFalse`, `Test_SaveClient_NothingFalla`, `Test_LoadClient_NoExisteNothing`, `Test_SaveOrder_ConservaFiscales`, `Test_FindClientByRfc_OtroCliente`, `Test_SearchClients_PorRfcYTelefono`, `Test_IsValidPostalCode`. (R1-R13)
- [x] T2 · Columnas y catálogos. `EnsureDataHeaders` con G:Q; `CLIENT_COLUMNS = 17`; semillas en `modConfig`; `EnsureSatCatalogs` + creación de hojas en `EnsureWorkbookStructure`; `ProtectAll`/`AdminUnlock`; `ListTaxRegimes`, `ListCfdiUses`. `modTestHelpers.MakeLegacyWorkbook` también borra G1:Q1 y las dos hojas de catálogo (y su comentario dice v9.0.0). Pasan las pruebas de encabezados, libro anterior y catálogos; siguen pasando `Test_P8_LibroAnterior` y las demás. (R1-R6)
- [x] T3 · Capa de clientes. `LoadClient`, `SaveClient`, `FindClientByRfc`, `SearchClients` (con `SortedNames` compartida con `ListClients`), `IsValidPostalCode`; revisar que `UpsertClientContact` no toque G:Q. Pasan todas las pruebas de T1. (R7-R13)
- [x] T4 · `frmClients`. `src/forms/frmClients.vba` con el maquetado, búsqueda, alta, edición, combos, avisos y confirmación de cambios; textos en `modStrings`; `lib.ps1` crea el formulario y `modBuild` le pone el logo. `PickClient` y `OpenClientsTool` en `modApp`. Verificación manual: build, abrirlo desde la ventana Inmediato con `OpenClientsTool`, alta con datos fiscales, CP "01000", RFC inválido y repetido, cerrar con cambios. (R14-R21, R25)
- [ ] T5 · Accesos. Botón "Clientes…" en `frmQuote` y `frmOrder` con el flujo de R24; `EnsureHomeClientsButton` en `PrepareWorkbook` y `DesignHome`, repintado en `PaintHomeColors`. Prueba automática `Test_Inicio_BotonClientes` (tras `PrepareWorkbook` existe la forma `HOME_CLIENTS_BUTTON_NAME` con `OnAction = "OpenClientsTool"`). Verificación manual: desde la orden elegir otro cliente reemplaza el contacto; cerrar sin elegir conserva el combo; libro de versión anterior gana el botón en Inicio. (R22-R24)
- [ ] T6 · Cierre. `APP_VERSION = "9.1.0"`; README: estructura de `BD_Clientes` A:Q, hojas de catálogo y cómo actualizarlas con `AdminUnlock`, uso de `frmClients`; `bflow check` en verde. (R2, R5, R25)
