Attribute VB_Name = "modUtils"
'==============================================================================
' modUtils · Utilidades generales
'
' Casi todo aquí son funciones "puras": reciben datos y regresan un resultado
' sin tocar hojas ni el formulario. Así son fáciles de probar y reutilizar.
'==============================================================================
Option Explicit

' Modo prueba (ver SetTestMode): guarda el ultimo error en vez de mostrar MsgBox.
Private mTestMode As Boolean
Private mLastErrorWhere As String
Private mLastErrorNumber As Long
Private mLastErrorDescription As String

'--- Números -------------------------------------------------------------------

' Intenta convertir texto a número.
'   value: (salida) el número convertido.
' Regresa True si la conversión fue exitosa.
Public Function TryParseNumber(ByVal inputText As String, ByRef numberValue As Double) As Boolean
    inputText = Trim$(inputText)
    If Len(inputText) = 0 Then Exit Function
    If Not IsNumeric(inputText) Then Exit Function

    On Error GoTo Fail                          ' Por si el número es gigantesco
    numberValue = CDbl(inputText)
    TryParseNumber = True
    Exit Function
Fail:
    TryParseNumber = False
End Function

' Redondea a 2 decimales como Excel (0.125 -> 0.13).
' El Round de VBA usa "redondeo bancario" (0.125 -> 0.12), por eso no se usa.
Public Function Round2(ByVal numberValue As Double) As Double
    Round2 = Application.WorksheetFunction.Round(numberValue, 2)
End Function

' Separador decimal de la configuración regional ("." o ",").
Public Function DecimalSeparator() As String
    DecimalSeparator = Mid$(CStr(1.5), 2, 1)
End Function

'--- Formatos ------------------------------------------------------------------

' 5 -> "5"  ·  2.5 -> "2.50"
Public Function FormatQuantity(ByVal qtyValue As Double) As String
    If qtyValue = Int(qtyValue) Then
        FormatQuantity = Format$(qtyValue, "#,##0")
    Else
        FormatQuantity = Format$(qtyValue, "#,##0.00")
    End If
End Function

' 1234.5 -> "$1,234.50"
Public Function FormatMoney(ByVal numberValue As Double) As String
    FormatMoney = Format$(numberValue, "$#,##0.00")
End Function

' 0.16 -> "16%"
Public Function FormatRate(ByVal rate As Double) As String
    FormatRate = Format$(rate * 100, "0") & "%"
End Function

' Reemplaza {0}, {1}... por los valores dados y \n por un salto de línea.
' Ejemplo: Fmt("Folio {0} guardado", "COT-1001")
Public Function Fmt(ByVal template As String, ParamArray args() As Variant) As String
    Dim i As Long
    For i = LBound(args) To UBound(args)
        template = Replace(template, "{" & i & "}", CStr(args(i)))
    Next i
    Fmt = Replace(template, "\n", vbCrLf)
End Function

' Une los textos que no estén vacíos con " · ".
Public Function JoinNonEmpty(ParamArray parts() As Variant) As String
    Dim part As Variant, result As String
    For Each part In parts
        If Len(CStr(part)) > 0 Then
            If Len(result) > 0 Then result = result & "   ·   "
            result = result & CStr(part)
        End If
    Next part
    JoinNonEmpty = result
End Function

'--- Seguridad de textos -------------------------------------------------------

' Evita que un texto que empieza con = + - @ se interprete como fórmula
' (inyección de fórmulas). Le antepone un apóstrofo para guardarlo como texto.
Public Function SafeCellText(ByVal inputText As String) As String
    If Len(inputText) > 0 Then
        If InStr("=+-@", Left$(inputText, 1)) > 0 Then inputText = "'" & inputText
    End If
    SafeCellText = inputText
End Function

' Quita caracteres prohibidos en nombres de archivo de Windows.
Public Function SanitizeFileName(ByVal inputText As String) As String
    Dim invalidChars As Variant, ch As Variant
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In invalidChars
        inputText = Replace(inputText, ch, "_")
    Next ch

    inputText = Trim$(inputText)
    Do While InStr(inputText, "  ") > 0
        inputText = Replace(inputText, "  ", " ")
    Loop
    If Len(inputText) > 120 Then inputText = Left$(inputText, 120)        ' Evita rutas demasiado largas
    SanitizeFileName = inputText
End Function

' Revisa el formato del RFC: 12 caracteres (persona moral) o 13 (persona física).
Public Function IsValidRfc(ByVal rfc As String) As Boolean
    rfc = UCase$(Trim$(rfc))
    Select Case Len(rfc)
        Case 12
            IsValidRfc = rfc Like "[A-ZÑ&][A-ZÑ&][A-ZÑ&]######[A-Z0-9][A-Z0-9][A-Z0-9]"
        Case 13
            IsValidRfc = rfc Like "[A-ZÑ&][A-ZÑ&][A-ZÑ&][A-ZÑ&]######[A-Z0-9][A-Z0-9][A-Z0-9]"
    End Select
End Function

' True solo con exactamente 5 digitos (codigo postal).
Public Function IsValidPostalCode(ByVal zipText As String) As Boolean
    IsValidPostalCode = (Len(zipText) = 5) And (zipText Like "#####")
End Function

'--- IVA -----------------------------------------------------------------------

' Tasa que corresponde a una opción de la lista (TaxOption).
Public Function TaxRateOf(ByVal opt As Long) As Double
    Select Case opt
        Case taxBorder: TaxRateOf = 0.08
        Case taxExempt: TaxRateOf = 0
        Case Else:      TaxRateOf = 0.16
    End Select
End Function

' Opción de la lista (TaxOption) que corresponde a una tasa.
Public Function TaxOptionOf(ByVal rate As Double) As Long
    If Abs(rate - 0.08) < 0.0001 Then
        TaxOptionOf = taxBorder
    ElseIf rate = 0 Then
        TaxOptionOf = taxExempt
    Else
        TaxOptionOf = taxStandard
    End If
End Function

'--- Objetos -------------------------------------------------------------------

' Crea un concepto ya lleno (las clases de VBA no aceptan parámetros al crearse).
Public Function NewQuoteItem(ByVal qtyValue As Double, ByVal descText As String, _
                             ByVal priceValue As Double) As QuoteItem
    Dim newItem As QuoteItem
    Set newItem = New QuoteItem
    newItem.Quantity = qtyValue
    newItem.Description = descText
    newItem.UnitPrice = priceValue
    Set NewQuoteItem = newItem
End Function

'--- Orden a cotización ---------------------------------------------------------

' Renglones del diagnóstico: parte en vbCrLf/vbLf/vbCr, recorta y omite vacíos.
Public Function DiagnosisLines(ByVal diagnosisText As String) As Collection
    Dim result As New Collection, parts() As String, i As Long, lineText As String

    parts = Split(Replace(Replace(diagnosisText, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For i = LBound(parts) To UBound(parts)
        lineText = Trim$(parts(i))
        If Len(lineText) > 0 Then result.Add lineText
    Next i
    Set DiagnosisLines = result
End Function

' Renglones de un texto (vacíos de en medio incluidos, sin los saltos del final).
Public Function TextLineCount(ByVal text As String) As Long
    Dim normalized As String
    normalized = NormalizeLineBreaks(text)
    Do While Right$(normalized, 1) = vbLf
        normalized = Left$(normalized, Len(normalized) - 1)
    Loop
    If Len(normalized) = 0 Then Exit Function
    TextLineCount = UBound(Split(normalized, vbLf)) + 1
End Function

Private Function NormalizeLineBreaks(ByVal text As String) As String
    NormalizeLineBreaks = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
End Function

' Diagnóstico en grupos de ORD_DIAG_LINES_PER_ROW renglones unidos con vbLf, uno por fila de la orden.
' El último grupo permitido se lleva lo que sobre. Unidos con vbLf dan el texto normalizado.
Public Function DiagnosisChunks(ByVal diagnosisText As String) As Collection
    Dim result As New Collection, parts() As String, n As Long
    Dim g As Long, groupCount As Long, first As Long, last As Long, i As Long, chunk As String

    Set DiagnosisChunks = result
    n = TextLineCount(diagnosisText)
    If n = 0 Then Exit Function

    parts = Split(NormalizeLineBreaks(diagnosisText), vbLf)
    groupCount = (n + ORD_DIAG_LINES_PER_ROW - 1) \ ORD_DIAG_LINES_PER_ROW
    If groupCount > ORD_DIAG_ROWS Then groupCount = ORD_DIAG_ROWS
    For g = 1 To groupCount
        first = (g - 1) * ORD_DIAG_LINES_PER_ROW
        last = first + ORD_DIAG_LINES_PER_ROW - 1
        If g = groupCount And g = ORD_DIAG_ROWS Then last = n - 1
        If last > n - 1 Then last = n - 1
        chunk = parts(first)
        For i = first + 1 To last
            chunk = chunk & vbLf & parts(i)
        Next i
        result.Add chunk
    Next g
End Function

' Cotización nueva (sin guardar) con un concepto por renglón del diagnóstico, hasta MAX_ITEMS.
Public Function QuoteFromOrder(ByVal o As ServiceOrder, ByVal quoteNumber As String) As Quote
    Dim q As Quote, rows As Collection, i As Long

    Set q = New Quote
    q.Number = quoteNumber
    q.ClientName = o.ClientName
    q.OrderNumber = o.Number
    q.QuoteDate = Date
    Set rows = DiagnosisLines(o.Diagnosis)
    For i = 1 To rows.Count
        If i > MAX_ITEMS Then Exit For
        q.AddItem NewQuoteItem(1, Left$(CStr(rows(i)), MAX_ITEM_DESCRIPTION), 0)
    Next i
    Set QuoteFromOrder = q
End Function

' Texto de la franja de frmQuote. o puede ser Nothing; sin folio regresa "".
Public Function OrderBannerText(ByVal orderNumber As String, ByVal o As ServiceOrder) As String
    Dim vehicle As String

    If Len(Trim$(orderNumber)) = 0 Then Exit Function
    OrderBannerText = Fmt(UI_ORDER_BANNER, Trim$(orderNumber))
    vehicle = VehicleText(o)
    If Len(vehicle) > 0 Then OrderBannerText = OrderBannerText & " · " & vehicle
End Function

' "Marca Modelo Año · Placas X" sin las partes vacías. o puede ser Nothing; sin datos regresa "".
Public Function VehicleText(ByVal o As ServiceOrder) As String
    Dim vehicle As String, plates As String, part As Variant

    If o Is Nothing Then Exit Function
    For Each part In Array(Trim$(o.Make), Trim$(o.Model), IIf(o.Year > 0, CStr(o.Year), ""))
        If Len(part) > 0 Then vehicle = vehicle & IIf(Len(vehicle) > 0, " ", "") & part
    Next part
    plates = UCase$(Trim$(o.Plates))
    VehicleText = vehicle
    If Len(plates) > 0 Then
        VehicleText = VehicleText & IIf(Len(vehicle) > 0, " · ", "") & Fmt(UI_ORDER_BANNER_PLATES, plates)
    End If
End Function

'--- Libro ---------------------------------------------------------------------

' True si existe una hoja con ese nombre en este libro.
Public Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

' Mensaje de error estándar. Llámalo desde el bloque de errores de una rutina.
'   where: nombre de la acción que falló.
Public Sub ShowError(ByVal where As String)
    Dim errNumber As Long, errText As String
    errNumber = Err.Number                      ' Se lee ANTES de hacer otra cosa
    errText = Err.Description
    Application.ScreenUpdating = True
    If mTestMode Then
        mLastErrorWhere = where
        mLastErrorNumber = errNumber
        mLastErrorDescription = errText
        Exit Sub
    End If
    MsgBox Fmt(MSG_ERROR, where, errNumber, errText), vbCritical, TTL_ERROR
End Sub

'--- Modo prueba -----------------------------------------------------------------

' Enciende o apaga el modo prueba; solo lo usa el runner de pruebas (modTestRunner): el libro entregado nunca entra en modo prueba.
Public Sub SetTestMode(ByVal enabled As Boolean)
    mTestMode = enabled
End Sub

' True mientras el modo prueba esta encendido.
Public Function IsTestMode() As Boolean
    IsTestMode = mTestMode
End Function

' Vacia donde, numero y descripcion del ultimo error registrado.
Public Sub ClearLastError()
    mLastErrorWhere = vbNullString
    mLastErrorNumber = 0
    mLastErrorDescription = vbNullString
End Sub

' Lugar (el texto de ShowError) del ultimo error registrado en modo prueba.
Public Function LastErrorWhere() As String
    LastErrorWhere = mLastErrorWhere
End Function

' 0 = no hubo error desde ClearLastError.
Public Function LastErrorNumber() As Long
    LastErrorNumber = mLastErrorNumber
End Function

' Descripcion del ultimo error registrado en modo prueba.
Public Function LastErrorDescription() As String
    LastErrorDescription = mLastErrorDescription
End Function

'--- Importe con letra (formato usual en México) -------------------------------

' 1044.00 -> "(MIL CUARENTA Y CUATRO PESOS 00/100 M.N.)"
' Soporta hasta 999,999,999.99
Public Function AmountToWords(ByVal amountValue As Double) As String
    Dim whole As Long, cents As Long, words As String, currencyName As String

    amountValue = Round2(amountValue)
    whole = Fix(amountValue)
    cents = CLng((amountValue - whole) * 100)

    If whole = 0 Then words = "CERO" Else words = IntegerToWords(whole)
    If whole = 1 Then currencyName = WORDS_CURRENCY_ONE Else currencyName = WORDS_CURRENCY_MANY
    ' "UN MILLÓN DE PESOS", "DOS MILLONES DE PESOS"
    If whole >= 1000000 And (whole Mod 1000000) = 0 Then currencyName = "DE " & currencyName

    AmountToWords = "(" & words & " " & currencyName & " " & Format$(cents, "00") & "/100 " & WORDS_SUFFIX & ")"
End Function

' Entero en palabras: separa millones, miles y el resto.
Private Function IntegerToWords(ByVal n As Long) As String
    Dim millions As Long, thousands As Long, rest As Long, result As String

    millions = n \ 1000000
    thousands = (n \ 1000) Mod 1000
    rest = n Mod 1000

    If millions = 1 Then
        result = "UN MILLÓN"
    ElseIf millions > 1 Then
        result = HundredsToWords(millions, True) & " MILLONES"
    End If

    If thousands = 1 Then
        result = result & " MIL"                    ' "MIL", no "UN MIL"
    ElseIf thousands > 1 Then
        result = result & " " & HundredsToWords(thousands, True) & " MIL"
    End If

    If rest > 0 Then result = result & " " & HundredsToWords(rest, True)
    IntegerToWords = Trim$(result)
End Function

' 1..999 en palabras. shorten=True convierte "UNO" en "UN" (antes de MIL o PESOS).
Private Function HundredsToWords(ByVal n As Long, ByVal shorten As Boolean) As String
    Dim hundreds As Long, rest As Long, result As String

    If n = 100 Then
        HundredsToWords = "CIEN"
        Exit Function
    End If

    hundreds = n \ 100
    rest = n Mod 100
    result = Choose(hundreds + 1, "", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", _
                    "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS")
    If rest > 0 Then result = result & " " & TensToWords(rest, shorten)
    HundredsToWords = Trim$(result)
End Function

' 1..99 en palabras.
Private Function TensToWords(ByVal n As Long, ByVal shorten As Boolean) As String
    Dim units As Variant, result As String
    units = Array("", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE")

    Select Case n
        Case 1 To 9
            result = units(n)
        Case 10 To 19
            result = Array("DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", _
                           "DIECISÉIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE")(n - 10)
        Case 20
            result = "VEINTE"
        Case 21 To 29
            result = Array("", "VEINTIUNO", "VEINTIDÓS", "VEINTITRÉS", "VEINTICUATRO", "VEINTICINCO", _
                           "VEINTISÉIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE")(n - 20)
        Case Else
            result = Array("TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", _
                           "OCHENTA", "NOVENTA")(n \ 10 - 3)
            If n Mod 10 > 0 Then result = result & " Y " & units(n Mod 10)
    End Select

    If shorten Then
        If result = "VEINTIUNO" Then
            result = "VEINTIÚN"
        ElseIf Right$(result, 3) = "UNO" Then
            result = Left$(result, Len(result) - 1)  ' "UNO" -> "UN", "Y UNO" -> "Y UN"
        End If
    End If
    TensToWords = result
End Function

'--- Validacion de campos ------------------------------------------------------

Public Function IsValidDate(ByVal dateText As String) As Boolean
    dateText = Trim$(dateText)
    IsValidDate = (Len(dateText) > 0 And IsDate(dateText))
End Function

' "H:MM" o "HH:MM", de 00:00 a 23:59.
Public Function IsValidTime(ByVal timeText As String) As Boolean
    Dim p As Long, hh As String, mm As String
    timeText = Trim$(timeText)
    p = InStr(timeText, ":")
    If p < 2 Then Exit Function
    hh = Left$(timeText, p - 1)
    mm = Mid$(timeText, p + 1)
    If Len(hh) > 2 Or Len(mm) <> 2 Then Exit Function
    If Not (hh Like "#" Or hh Like "##") Then Exit Function
    If Not mm Like "##" Then Exit Function
    IsValidTime = (CLng(hh) <= 23 And CLng(mm) <= 59)
End Function

' 4 digitos entre MIN_VEHICLE_YEAR y el ano de referencia + 1.
Public Function IsValidVehicleYear(ByVal yearText As String, Optional ByVal referenceDate As Date = 0) As Boolean
    Dim y As Long
    yearText = Trim$(yearText)
    If Not yearText Like "####" Then Exit Function
    If referenceDate = 0 Then referenceDate = Date
    y = CLng(yearText)
    IsValidVehicleYear = (y >= MIN_VEHICLE_YEAR And y <= Year(referenceDate) + 1)
End Function

' De 1 a 9 digitos, sin separadores.
Public Function IsValidMileage(ByVal mileageText As String) As Boolean
    Dim i As Long
    mileageText = Trim$(mileageText)
    If Len(mileageText) < 1 Or Len(mileageText) > 9 Then Exit Function
    For i = 1 To Len(mileageText)
        If Not Mid$(mileageText, i, 1) Like "[0-9]" Then Exit Function
    Next i
    IsValidMileage = True
End Function

' 17 caracteres alfanumericos, sin importar mayusculas.
Public Function IsValidVin(ByVal vinText As String) As Boolean
    Dim i As Long
    vinText = Trim$(vinText)
    If Len(vinText) <> 17 Then Exit Function
    For i = 1 To 17
        If Not Mid$(vinText, i, 1) Like "[A-Za-z0-9]" Then Exit Function
    Next i
    IsValidVin = True
End Function
