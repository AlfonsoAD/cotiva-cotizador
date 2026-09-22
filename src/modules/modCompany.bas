Attribute VB_Name = "modCompany"
'==============================================================================
' modCompany · Datos de la empresa que emite las cotizaciones (hoja "Mi empresa")
'
' El usuario captura aquí nombre, RFC, dirección, teléfono, correo, sitio web,
' días de vigencia y colores. Todo lo demás del cotizador lee estos datos con
' las funciones de este módulo.
'==============================================================================
Option Explicit

'--- Lectura de datos ----------------------------------------------------------

' Lee un dato de la hoja Mi empresa ("" si la hoja no existe o la celda está vacía).
'   cellAddress: una de las constantes COMPANY_..._CELL de modConfig.
Public Function CompanyValue(ByVal cellAddress As String) As String
    If Not SheetExists(SHEET_COMPANY) Then Exit Function
    CompanyValue = Trim$(CStr(ThisWorkbook.Worksheets(SHEET_COMPANY).Range(cellAddress).Value))
    If cellAddress = COMPANY_RFC_CELL Then CompanyValue = UCase$(CompanyValue)
End Function

' Días de vigencia configurados (1 a 365); si no hay dato válido, el valor por defecto.
Public Function ValidityDays() As Long
    Dim days As Double
    If TryParseNumber(CompanyValue(COMPANY_VALIDITY_CELL), days) Then
        If days >= 1 And days <= 365 Then
            ValidityDays = CLng(days)
            Exit Function
        End If
    End If
    ValidityDays = DEFAULT_VALIDITY_DAYS
End Function

' Color principal de la empresa; si no hay uno válido, el naranja de Cotiva.
Public Function PrimaryColor() As Long
    Dim c As Long
    If TryParseHexColor(CompanyValue(COMPANY_PRIMARY_COLOR_CELL), c) Then
        PrimaryColor = c
    Else
        PrimaryColor = BrandOrange()
    End If
End Function

' Color secundario de la empresa; si no hay uno válido, el azul tinta de Cotiva.
Public Function SecondaryColor() As Long
    Dim c As Long
    If TryParseHexColor(CompanyValue(COMPANY_SECONDARY_COLOR_CELL), c) Then
        SecondaryColor = c
    Else
        SecondaryColor = BrandInk()
    End If
End Function

' Primera línea bajo el nombre: "RFC ... · Tel. ... · correo"
Public Function CompanyContactLine() As String
    Dim rfc As String, phone As String
    rfc = CompanyValue(COMPANY_RFC_CELL)
    If Len(rfc) > 0 Then rfc = "RFC " & rfc
    phone = CompanyValue(COMPANY_PHONE_CELL)
    If Len(phone) > 0 Then phone = "Tel. " & phone
    CompanyContactLine = JoinNonEmpty(rfc, phone, CompanyValue(COMPANY_EMAIL_CELL))
End Function

' Segunda línea: "dirección · sitio web"
Public Function CompanyAddressLine() As String
    CompanyAddressLine = JoinNonEmpty(CompanyValue(COMPANY_ADDRESS_CELL), CompanyValue(COMPANY_WEB_CELL))
End Function

' Créditos de innobytes: "Cotiva es un producto de innobytes · web · correo · tel."
Public Function InnobytesCredit() As String
    InnobytesCredit = JoinNonEmpty(TPL_INNOBYTES, INNOBYTES_WEB, INNOBYTES_EMAIL, INNOBYTES_PHONE)
End Function

'--- Hoja Mi empresa -----------------------------------------------------------

' Crea la hoja la primera vez. El libro debe estar desprotegido.
Public Sub EnsureCompanySheet()
    Dim ws As Worksheet

    If SheetExists(SHEET_COMPANY) Then Exit Sub

    If SheetExists(SHEET_TEMPLATE) Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(SHEET_TEMPLATE))
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    End If
    ws.Name = SHEET_COMPANY
    DesignCompanySheet ws
    ws.Activate
    ActiveWindow.DisplayGridlines = False
End Sub

' Si la hoja se creó con una versión anterior (sin colores), agrega los campos
' nuevos conservando lo capturado.
Public Sub UpgradeCompanySheet()
    Dim ws As Worksheet
    If Not SheetExists(SHEET_COMPANY) Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SHEET_COMPANY)
    If Len(Trim$(CStr(ws.Range("A12").Value))) > 0 Then Exit Sub   ' Ya está al día
    UnlockSheet ws
    DesignCompanySheet ws
End Sub

' Arma el formato: etiquetas a la izquierda, celdas blancas desbloqueadas a la
' derecha (lo único editable con la hoja protegida).
' Se puede ejecutar varias veces: NO borra lo que el usuario ya capturó.
Private Sub DesignCompanySheet(ByVal ws As Worksheet)
    Dim labels As Variant, i As Long, r As Long

    With ws.Range("A1:D20").Font
        .Name = BrandFont()
        .Size = 10
        .Color = BrandInk()
    End With
    ws.Columns("A").ColumnWidth = 40
    ws.Columns("B").ColumnWidth = 60
    ws.Columns("C").ColumnWidth = 5

    With ws.Range("A1")
        .Value = CMP_TITLE
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = BrandOrange()
    End With
    With ws.Range("A2")
        .Value = CMP_SUBTITLE
        .Font.Size = 9
        .Font.Color = TextGray()
    End With
    ws.Rows(1).RowHeight = 28

    labels = Array(CMP_LBL_NAME, CMP_LBL_RFC, CMP_LBL_ADDRESS, CMP_LBL_PHONE, CMP_LBL_EMAIL, _
                   CMP_LBL_WEB, CMP_LBL_VALIDITY, CMP_LBL_PRIMARY, CMP_LBL_SECONDARY)
    ws.Range("A3:C14").Interior.Color = Neutral50()

    For i = 0 To UBound(labels)
        r = 4 + i                                   ' Filas 4 a 12
        ws.Rows(r).RowHeight = 24
        With ws.Cells(r, 1)
            .Value = labels(i)
            .Font.Color = TextGray()
            .VerticalAlignment = xlCenter
        End With
        With ws.Cells(r, 2)
            .Locked = False                         ' Editable con la hoja protegida
            .Interior.Color = vbWhite
            .VerticalAlignment = xlCenter
            .IndentLevel = 1
            .Borders.LineStyle = xlContinuous
            .Borders.Color = Neutral300()
        End With
    Next i

    ' Valores iniciales solo si la celda está vacía (no pisa lo capturado)
    SetIfEmpty ws.Range(COMPANY_VALIDITY_CELL), DEFAULT_VALIDITY_DAYS
    SetIfEmpty ws.Range(COMPANY_PRIMARY_COLOR_CELL), DEFAULT_PRIMARY_HEX
    SetIfEmpty ws.Range(COMPANY_SECONDARY_COLOR_CELL), DEFAULT_SECONDARY_HEX

    ' Validaciones de captura
    AddValidation ws.Range(COMPANY_VALIDITY_CELL), xlValidateWholeNumber, xlValidAlertStop, "1", "365", CMP_ERR_VALIDITY
    AddValidation ws.Range(COMPANY_RFC_CELL), xlValidateTextLength, xlValidAlertWarning, "12", "13", CMP_ERR_RFC
    AddValidation ws.Range(COMPANY_PRIMARY_COLOR_CELL & ":" & COMPANY_SECONDARY_COLOR_CELL), _
                  xlValidateTextLength, xlValidAlertStop, "6", "7", CMP_ERR_COLOR
    ws.Range(COMPANY_VALIDITY_CELL).HorizontalAlignment = xlLeft

    ' Muestras de color (columna C) y notas
    With ws.Range("C11:C12").Borders
        .LineStyle = xlContinuous
        .Color = Neutral300()
    End With
    RefreshColorSwatches

    ws.Range("A14:A16").ClearContents
    WriteNote ws.Range("A15"), CMP_NOTE_COLORS
    WriteNote ws.Range("A16"), CMP_NOTE_LOGO
End Sub

Private Sub SetIfEmpty(ByVal cell As Range, ByVal newValue As Variant)
    If Len(Trim$(CStr(cell.Value))) = 0 Then cell.Value = newValue
End Sub

' Reemplaza la validación de datos de un rango.
Private Sub AddValidation(ByVal target As Range, ByVal validationType As XlDVType, _
                          ByVal alertStyle As XlDVAlertStyle, ByVal minValue As String, _
                          ByVal maxValue As String, ByVal errorMessage As String)
    With target.Validation
        .Delete
        .Add Type:=validationType, AlertStyle:=alertStyle, Operator:=xlBetween, _
             Formula1:=minValue, Formula2:=maxValue
        .ErrorMessage = errorMessage
    End With
End Sub

Private Sub WriteNote(ByVal cell As Range, ByVal noteText As String)
    With cell
        .Value = noteText
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = TextGray()
    End With
End Sub

' Pinta las muestras de color junto a los códigos.
Public Sub RefreshColorSwatches()
    If Not SheetExists(SHEET_COMPANY) Then Exit Sub
    With ThisWorkbook.Worksheets(SHEET_COMPANY)
        .Range("C11").Interior.Color = PrimaryColor()
        .Range("C12").Interior.Color = SecondaryColor()
    End With
End Sub

' Muestra la hoja lista para capturar. La usa el botón "Mi empresa".
Public Sub OpenCompanySheet()
    If Not SheetExists(SHEET_COMPANY) Then Exit Sub
    RefreshColorSwatches
    ThisWorkbook.Worksheets(SHEET_COMPANY).Activate
    ThisWorkbook.Worksheets(SHEET_COMPANY).Range(COMPANY_NAME_CELL).Select
End Sub
