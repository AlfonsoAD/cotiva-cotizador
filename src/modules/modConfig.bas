Attribute VB_Name = "modConfig"
'==============================================================================
' modConfig · Configuración general del cotizador
'
' Aquí viven todas las constantes: nombres de hojas, celdas, límites y datos
' de innobytes. Si algo "de configuración" cambia, se cambia aquí y en ningún
' otro lugar. La clave de protección NO está aquí: vive en modSecrets, que no
' se sube a Git.
'==============================================================================
Option Explicit

' --- Aplicación ---
Public Const APP_VERSION As String = "8.0.0"

' --- Nombres de hojas (se conservan en español por compatibilidad con los datos) ---
Public Const SHEET_HOME As String = "Inicio"
Public Const SHEET_TEMPLATE As String = "Plantilla"
Public Const SHEET_QUOTES As String = "BD_Cotizaciones"
Public Const SHEET_ITEMS As String = "BD_Detalle"
Public Const SHEET_CLIENTS As String = "BD_Clientes"
Public Const SHEET_COMPANY As String = "Mi empresa"

' --- Folios ---
Public Const QUOTE_PREFIX As String = "COT-"
Public Const QUOTE_BASE_NUMBER As Long = 1000          ' El primer folio será COT-1001

' --- Tasas de IVA ---
' El orden coincide con la lista del formulario (16% / 8% / Sin IVA).
Public Enum TaxOption
    taxStandard = 0        ' 16 %
    taxBorder = 1          ' 8 % (región fronteriza)
    taxExempt = 2          ' Sin IVA
End Enum
Public Const DEFAULT_TAX_OPTION As Long = 0            ' Tasa con la que arranca cada cotización

' --- Plantilla: filas y límites ---
Public Const FIRST_ITEM_ROW As Long = 9
Public Const LAST_ITEM_ROW As Long = 25
Public Const MAX_ITEMS As Long = LAST_ITEM_ROW - FIRST_ITEM_ROW + 1   ' = 17 conceptos
Public Const MIN_VISIBLE_ITEM_ROWS As Long = 8          ' Filas visibles aunque haya menos conceptos
Public Const MIN_ITEM_ROW_HEIGHT As Double = 20         ' Alto mínimo de cada fila de concepto
Public Const PRINT_AREA As String = "$A$1:$D$34"

' --- Plantilla: celdas de datos ---
Public Const CELL_QUOTE_NUMBER As String = "D2"
Public Const CELL_DATE As String = "D3"
Public Const CELL_VALID_UNTIL As String = "D4"
Public Const CELL_CLIENT As String = "B6"
Public Const CELL_COMPANY_NAME As String = "A2"
Public Const CELL_COMPANY_CONTACT As String = "A3"
Public Const CELL_COMPANY_ADDRESS As String = "A4"
Public Const CELL_SUBTOTAL As String = "D26"
Public Const CELL_TAX_LABEL As String = "C27"
Public Const CELL_TAX As String = "D27"
Public Const CELL_TOTAL As String = "D28"
Public Const CELL_AMOUNT_WORDS As String = "A27"

' --- Logos ---
Public Const COMPANY_LOGO_NAME As String = "LogoCotizacion"   ' Logo de la empresa que cotiza
Public Const COTIVA_LOGO_NAME As String = "LogoCotiva"        ' Logo fijo de Cotiva (abajo)
Public Const COTIVA_LOGO_HEIGHT As Double = 40                ' Alto del logo de Cotiva (puntos)

' --- Hoja "Inicio": formas y logos de la portada ---
' Los nombres identifican cada imagen dentro de la hoja: DesignHome borra las
' demás formas y vuelve a dibujarlas, pero estas dos las conserva.
Public Const HOME_COTIVA_LOGO_NAME As String = "LogoCotivaInicio"
Public Const HOME_INNOBYTES_LOGO_NAME As String = "LogoInnobytesInicio"
Public Const HOME_COTIVA_LOGO_HEIGHT As Double = 44           ' Alto del logo de Cotiva (puntos)
Public Const HOME_INNOBYTES_LOGO_HEIGHT As Double = 48        ' Alto del logo de innobytes (pie)
Public Const HOME_BUTTON_NAME As String = "btnAbrirCotizador"
Public Const HOME_RULE_NAME As String = "reglaInicio"
Public Const HOME_FIRST_FEATURE_ROW As Long = 10              ' Primera fila de la lista de ventajas

' --- Hoja "Mi empresa": celdas que captura el usuario ---
Public Const COMPANY_NAME_CELL As String = "B4"
Public Const COMPANY_RFC_CELL As String = "B5"
Public Const COMPANY_ADDRESS_CELL As String = "B6"
Public Const COMPANY_PHONE_CELL As String = "B7"
Public Const COMPANY_EMAIL_CELL As String = "B8"
Public Const COMPANY_WEB_CELL As String = "B9"
Public Const COMPANY_VALIDITY_CELL As String = "B10"
Public Const COMPANY_PRIMARY_COLOR_CELL As String = "B11"
Public Const COMPANY_SECONDARY_COLOR_CELL As String = "B12"

' --- Valores por defecto ---
Public Const DEFAULT_VALIDITY_DAYS As Long = 15
Public Const DEFAULT_PRIMARY_HEX As String = "#CF4A06"        ' Naranja Cotiva
Public Const DEFAULT_SECONDARY_HEX As String = "#1A2638"      ' Azul tinta Cotiva

' --- Comportamiento ---
Public Const PDF_FOLDER As String = "Cotizaciones_PDF"        ' Subcarpeta para los PDF
Public Const SAVE_WORKBOOK_ON_SAVE As Boolean = True          ' Guarda el .xlsm tras cada cotización
Public Const HIDE_DATA_SHEETS As Boolean = True               ' Hojas BD_ invisibles para el usuario
Public Const FORM_LOGO_FILE As String = "logo_cotizador.jpg"  ' Respaldo si el formulario no trae imagen

' --- innobytes (pie del PDF y del formulario) ---
' Lo que quede vacío simplemente no se muestra.
' Si los cambias, ejecuta DesignTemplate para actualizar la Plantilla.
Public Const INNOBYTES_WEB As String = "innobytes.tech"
Public Const INNOBYTES_EMAIL As String = ""
Public Const INNOBYTES_PHONE As String = ""
