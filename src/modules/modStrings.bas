Attribute VB_Name = "modStrings"
'==============================================================================
' modStrings · Textos que ve el usuario (interfaz, mensajes y PDF)
'
' Todo el texto visible está aquí, en español. Para una versión en otro idioma
' basta con traducir este módulo.
'
' Marcadores dentro de los textos (los reemplaza Fmt en modUtils):
'   {0}, {1}...  -> valores que se insertan
'   \n           -> salto de línea
'==============================================================================
Option Explicit

' --- Formulario: títulos y etiquetas ---
Public Const UI_FORM_CAPTION As String = "Cotiva · Cotizador"
Public Const UI_BRAND As String = "Cotiva"
Public Const UI_MODULE As String = "Cotizador"
Public Const UI_SECTION_QUOTE As String = "DATOS DE LA COTIZACIÓN"
Public Const UI_SECTION_ITEM As String = "AGREGAR CONCEPTO"
Public Const UI_LBL_NUMBER As String = "Folio"
Public Const UI_LBL_CLIENT As String = "Cliente (se guarda en tu catálogo)"
Public Const UI_LBL_DATE As String = "Fecha (dd/mm/aaaa)"
Public Const UI_LBL_QTY As String = "Cantidad"
Public Const UI_LBL_DESCRIPTION As String = "Descripción"
Public Const UI_LBL_PRICE As String = "Precio unitario"
Public Const UI_LBL_AMOUNT As String = "Importe"
Public Const UI_COL_QTY As String = "Cant."
Public Const UI_COL_DESCRIPTION As String = "Descripción"
Public Const UI_COL_PRICE As String = "Precio unit."
Public Const UI_COL_AMOUNT As String = "Importe"
Public Const UI_LBL_SUBTOTAL As String = "Subtotal"
Public Const UI_LBL_TAX As String = "IVA"
Public Const UI_LBL_TOTAL As String = "Total"
Public Const UI_VALID_UNTIL As String = "Vigente hasta {0}"

' --- Formulario: botones ---
Public Const UI_BTN_LOAD As String = "Cargar"
Public Const UI_BTN_ADD As String = "Agregar"
Public Const UI_BTN_UPDATE As String = "Actualizar"
Public Const UI_BTN_EDIT As String = "Editar"
Public Const UI_BTN_CANCEL As String = "Cancelar"
Public Const UI_BTN_DELETE As String = "Eliminar"
Public Const UI_BTN_NEW As String = "Nueva cotización"
Public Const UI_BTN_LOGO As String = "Mi logo"
Public Const UI_BTN_COMPANY As String = "Mi empresa"
Public Const UI_BTN_PREVIEW As String = "Vista previa"
Public Const UI_BTN_SAVE As String = "Guardar"
Public Const UI_BTN_PDF As String = "Generar PDF"

' --- Formulario: ayudas (tooltips) ---
Public Const TIP_HISTORY As String = "Escribe o elige un folio guardado"
Public Const TIP_CLIENT As String = "Escribe el nombre: si ya existe se autocompleta"
Public Const TIP_TAX As String = "Tasa de IVA de esta cotización"
Public Const TIP_LOGO As String = "Poner o cambiar el logo de tu empresa en la cotización"
Public Const TIP_COMPANY As String = "Nombre, RFC, dirección, teléfono, correo, vigencia y colores"
Public Const TIP_INNOBYTES As String = "Visitar {0}"

' --- Opciones de IVA (mismo orden que el Enum TaxOption) ---
Public Const UI_TAX_16 As String = "16%"
Public Const UI_TAX_8 As String = "8%"
Public Const UI_TAX_EXEMPT As String = "Sin IVA"

' --- Títulos de ventanas de mensaje ---
Public Const TTL_REVIEW As String = "Revisa los datos"
Public Const TTL_ERROR As String = "Error"
Public Const TTL_CLOSE As String = "Cerrar cotizador"
Public Const TTL_NEW As String = "Nueva cotización"
Public Const TTL_DELETE As String = "Eliminar concepto"
Public Const TTL_OVERWRITE As String = "Sobrescribir"
Public Const TTL_PDF As String = "Generar PDF"
Public Const TTL_PDF_EXISTS As String = "PDF existente"
Public Const TTL_PDF_ERROR As String = "Error al exportar"
Public Const TTL_COMPANY As String = "Datos de tu empresa"
Public Const TTL_WRONG_PASSWORD As String = "Clave incorrecta"
Public Const TTL_MISSING_SHEETS As String = "Faltan hojas"
Public Const TTL_DESIGN As String = "Diseño Cotiva"
Public Const TTL_ADMIN As String = "Administrador"

' --- Validaciones ---
Public Const MSG_QTY_NOT_NUMBER As String = "La cantidad debe ser un número."
Public Const MSG_QTY_NOT_POSITIVE As String = "La cantidad debe ser mayor a 0."
Public Const MSG_DESCRIPTION_EMPTY As String = "Escribe una descripción para el concepto."
Public Const MSG_PRICE_NOT_NUMBER As String = "El precio unitario debe ser un número."
Public Const MSG_PRICE_NEGATIVE As String = "El precio unitario no puede ser negativo."
Public Const MSG_MAX_ITEMS As String = "La plantilla admite máximo {0} conceptos."
Public Const MSG_FINISH_EDIT As String = "Termina la edición del concepto (Actualizar o Cancelar)."
Public Const MSG_CLIENT_EMPTY As String = "Escribe el nombre del cliente."
Public Const MSG_DATE_INVALID As String = "La fecha no es válida. Usa el formato dd/mm/aaaa."
Public Const MSG_NO_ITEMS As String = "Agrega al menos un concepto a la cotización."
Public Const MSG_SELECT_ITEM_EDIT As String = "Selecciona primero un concepto de la lista."
Public Const MSG_SELECT_ITEM_DELETE As String = "Selecciona el concepto que deseas eliminar."
Public Const MSG_CONFIRM_DELETE As String = "¿Eliminar este concepto?\n\n{0}"

' --- Guardado, historial y cierre ---
Public Const MSG_UNSAVED_CLOSE As String = "Hay cambios sin guardar. ¿Cerrar de todos modos?"
Public Const MSG_UNSAVED_NEW As String = "Hay cambios sin guardar que se perderán.\n¿Empezar una cotización nueva?"
Public Const MSG_UNSAVED_LOAD As String = "Hay cambios sin guardar que se perderán.\n¿Cargar {0} de todos modos?"
Public Const MSG_CONFIRM_OVERWRITE As String = "El folio {0} ya existe.\n¿Sobrescribirlo con los cambios actuales?"
Public Const MSG_NUMBER_REASSIGNED As String = "El folio ya estaba ocupado. Se asignó {0}."
Public Const MSG_SAVED As String = "Cotización {0} guardada correctamente."
Public Const MSG_HISTORY_EMPTY As String = "Elige o escribe un folio del historial."
Public Const MSG_QUOTE_NOT_FOUND As String = "No existe el folio {0}."

' --- PDF ---
Public Const MSG_SAVE_BEFORE_PDF As String = "La cotización tiene cambios sin guardar.\n¿Guardarla antes de generar el PDF?"
Public Const MSG_PDF_CREATED As String = "PDF generado en:\n{0}"
Public Const MSG_PDF_EXISTS As String = "Ya existe este PDF:\n{0}\n\n¿Deseas reemplazarlo?"
Public Const MSG_PDF_FAILED As String = "No se pudo generar el PDF.\nSi ya existe, revisa que no esté abierto en otro programa.\n\nDetalle: {0}"
Public Const MSG_COMPANY_NAME_MISSING As String = "No has capturado el nombre de tu empresa."
Public Const MSG_COMPANY_RFC_INVALID As String = "El RFC {0} no tiene un formato válido."
Public Const MSG_COMPANY_CONTINUE As String = "{0}\nPuedes corregirlo con el botón «Mi empresa».\n\n¿Generar el PDF de todos modos?"

' --- Seguridad y administración ---
Public Const MSG_WRONG_PASSWORD As String = "La clave de modSecrets no coincide con la que protege este libro.\nPon la clave correcta en APP_PASSWORD y vuelve a intentar."
Public Const MSG_ADMIN_PROMPT As String = "Clave de administrador:"
Public Const MSG_ADMIN_WRONG As String = "Clave incorrecta."
Public Const MSG_ADMIN_UNLOCKED As String = "Hojas visibles y editables.\nAl terminar ejecuta AdminLock."
Public Const MSG_ADMIN_LOCKED As String = "Libro protegido."
Public Const MSG_ERROR As String = "Ocurrió un error en: {0}\n\nError {1}: {2}"

' --- Diseño y logos ---
Public Const MSG_DESIGN_CONFIRM As String = "Esto rediseña la hoja Plantilla desde cero (los logos se conservan).\n¿Continuar?"
Public Const MSG_DESIGN_DONE As String = "Plantilla lista. Revisa el resultado con Archivo > Imprimir."
Public Const MSG_DESIGN_ASK_LOGO As String = "Plantilla lista. ¿Quieres insertar el logo de tu empresa ahora?"
Public Const MSG_COTIVA_LOGO_DONE As String = "Logo de Cotiva colocado."
Public Const DLG_IMAGE_FILTER As String = "Imágenes (*.png;*.jpg;*.jpeg;*.gif;*.bmp),*.png;*.jpg;*.jpeg;*.gif;*.bmp"
Public Const DLG_PICK_COMPANY_LOGO As String = "Elige el logo de tu empresa"
Public Const DLG_PICK_COTIVA_LOGO As String = "Elige el logo de Cotiva"
Public Const MSG_DESIGN_HOME_CONFIRM As String = "Esto rediseña la hoja Inicio desde cero (los logos se conservan).\n¿Continuar?"
Public Const MSG_DESIGN_HOME_DONE As String = "Pantalla de inicio lista."
Public Const MSG_HOME_LOGO_DONE As String = "Logo colocado en la pantalla de inicio."
Public Const DLG_PICK_COTIVA_HOME_LOGO As String = "Elige el logo de Cotiva para la pantalla de inicio"
Public Const DLG_PICK_INNOBYTES_LOGO As String = "Elige el logo de innobytes"

' --- Plantilla (PDF) ---
Public Const TPL_TITLE As String = "COTIZACIÓN"
Public Const TPL_NUMBER As String = "Folio"
Public Const TPL_DATE As String = "Fecha"
Public Const TPL_VALID_UNTIL As String = "Vigencia"
Public Const TPL_CLIENT As String = "Cliente"
Public Const TPL_COL_QTY As String = "Cant."
Public Const TPL_COL_DESCRIPTION As String = "Descripción"
Public Const TPL_COL_PRICE As String = "Precio unitario"
Public Const TPL_COL_AMOUNT As String = "Importe"
Public Const TPL_SUBTOTAL As String = "Subtotal"
Public Const TPL_TAX As String = "IVA ({0})"
Public Const TPL_TAX_EXEMPT As String = "Sin IVA"
Public Const TPL_TOTAL As String = "Total"
Public Const TPL_AMOUNT_WORDS As String = "Importe con letra"
Public Const TPL_TERMS_TITLE As String = "Condiciones"
Public Const TPL_TERMS_DEFAULT As String = "Precios en pesos mexicanos (MXN). Esta cotización es válida hasta la fecha de vigencia indicada. Tiempos de entrega sujetos a confirmación al aprobar la cotización."
Public Const TPL_MADE_WITH As String = "Generado con"
Public Const TPL_INNOBYTES As String = "Cotiva es un producto de innobytes"
Public Const TPL_PAGE As String = "Página &P de &N"

' --- Hoja "Mi empresa" ---
Public Const CMP_TITLE As String = "Datos de tu empresa"
Public Const CMP_SUBTITLE As String = "Aparecen en cada cotización. Escribe solo en las celdas blancas."
Public Const CMP_LBL_NAME As String = "Nombre o razón social"
Public Const CMP_LBL_RFC As String = "RFC"
Public Const CMP_LBL_ADDRESS As String = "Dirección"
Public Const CMP_LBL_PHONE As String = "Teléfono"
Public Const CMP_LBL_EMAIL As String = "Correo electrónico"
Public Const CMP_LBL_WEB As String = "Sitio web (opcional)"
Public Const CMP_LBL_VALIDITY As String = "Vigencia de las cotizaciones (días)"
Public Const CMP_LBL_PRIMARY As String = "Color principal (título, línea y total)"
Public Const CMP_LBL_SECONDARY As String = "Color secundario (encabezado de la tabla)"
Public Const CMP_ERR_VALIDITY As String = "Escribe un número de días entre 1 y 365."
Public Const CMP_ERR_RFC As String = "El RFC tiene 12 caracteres (persona moral) o 13 (persona física)."
Public Const CMP_ERR_COLOR As String = "Escribe el color en formato hexadecimal, por ejemplo #CF4A06."
Public Const CMP_NOTE_COLORS As String = "Colores en formato hexadecimal (#RRGGBB). Tómalos de tu manual de marca; si el código no es válido se usan los colores de Cotiva."
Public Const CMP_NOTE_LOGO As String = "El logo se cambia con el botón «Mi logo» del cotizador."

' --- Hoja "Inicio" ---
Public Const HOME_TITLE As String = "Cotizador"
Public Const HOME_SUBTITLE As String = "Crea, guarda y exporta cotizaciones en PDF con la imagen de tu empresa."
Public Const HOME_BUTTON As String = "Abrir cotizador"
Public Const HOME_FIRST_TIME As String = "Primera vez: abre el cotizador y usa «Mi empresa» y «Mi logo» para personalizar tus cotizaciones."
Public Const HOME_BULLET As String = "•"
Public Const HOME_FEATURE_1 As String = "Folio automático, catálogo de clientes e IVA 16% / 8% / sin IVA."
Public Const HOME_FEATURE_2 As String = "Tu logo, tus colores y tus datos en cada cotización."
Public Const HOME_FEATURE_3 As String = "PDF listo para enviar, guardado junto al archivo."
Public Const HOME_MAKER As String = "innobytes"
Public Const HOME_VERSION As String = "Versión {0}"

' --- Importe con letra ---
Public Const WORDS_CURRENCY_ONE As String = "PESO"
Public Const WORDS_CURRENCY_MANY As String = "PESOS"
Public Const WORDS_SUFFIX As String = "M.N."
