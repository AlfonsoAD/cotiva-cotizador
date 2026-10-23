Attribute VB_Name = "modTemplate"
'==============================================================================
' modTemplate · Hojas Plantilla y PlantillaOrden: diseño, llenado, colores, logos y PDF
'
' Mapa de la hoja:
'   A1:B1  logo de la empresa       C1:D1  "COTIZACIÓN"
'   A2:B2  nombre de la empresa     C2/D2  Folio
'   A3:B3  RFC · Tel. · correo      C3/D3  Fecha
'   A4:B4  dirección · web          C4/D4  Vigencia
'   Fila 5 línea de color           A6 / B6:D6  Cliente
'                                   A7 / B7:D7  Vehículo (solo con orden ligada)
'   Fila 8 encabezado de tabla      A9:D25 conceptos
'   A26:B28 importe con letra       C26:D28 totales
'   A30:D31 condiciones (editable)  B33:D33 "Generado con" + logo Cotiva
'   A34:D34 créditos de innobytes (con liga)
'
' PlantillaOrden (cuadrícula A-H: etiqueta/valor x3; C y F separan):
'   Fila 1 logo A1:D1 y título   Filas 2-4 empresa y folio/estado/prometida
'   Filas 6-18 cliente, vehículo, entrada y salida
'   Filas 20-23 servicio y diagnóstico   A26:H26 condiciones (editable)
'   Filas 28-29 firmas: Cliente, Recibió, Entregó
'==============================================================================
Option Explicit

Private Function TemplateSheet() As Worksheet
    Set TemplateSheet = ThisWorkbook.Worksheets(SHEET_TEMPLATE)
End Function

Private Function OrderTemplateSheet() As Worksheet
    Set OrderTemplateSheet = ThisWorkbook.Worksheets(SHEET_ORDER_TEMPLATE)
End Function

' True si en la hoja existe una imagen o forma con ese nombre.
' Publica porque modHome la usa para los logos de la portada.
Public Function ShapeExists(ByVal ws As Worksheet, ByVal shapeName As String) As Boolean
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0
    ShapeExists = Not shp Is Nothing
End Function

'--- Llenado -------------------------------------------------------------------

' Deja la Plantilla en blanco (versión pública: maneja la protección).
Public Sub ClearTemplate()
    Dim ws As Worksheet
    Set ws = TemplateSheet()
    UnlockSheet ws
    ClearTemplateContent ws
    LockSheet ws
End Sub

' Borra solo los datos (conserva el diseño) y muestra todas las filas de conceptos.
Private Sub ClearTemplateContent(ByVal ws As Worksheet)
    Dim cellAddress As Variant
    For Each cellAddress In Array(CELL_QUOTE_NUMBER, CELL_DATE, CELL_VALID_UNTIL, CELL_CLIENT, _
                              CELL_COMPANY_NAME, CELL_COMPANY_CONTACT, CELL_COMPANY_ADDRESS, _
                              CELL_AMOUNT_WORDS, CELL_SUBTOTAL, CELL_TAX, CELL_TOTAL)
        ws.Range(CStr(cellAddress)).MergeArea.ClearContents   ' MergeArea: funciona con celdas combinadas
    Next cellAddress
    ws.Range("A" & FIRST_ITEM_ROW & ":D" & LAST_ITEM_ROW).ClearContents
    ws.Rows(FIRST_ITEM_ROW & ":" & LAST_ITEM_ROW).Hidden = False
    ws.Range(CELL_VEHICLE_LABEL).ClearContents
    ws.Range(CELL_VEHICLE).MergeArea.ClearContents
    ws.Range(CELL_VEHICLE).EntireRow.RowHeight = VEHICLE_ROW_EMPTY_HEIGHT
End Sub

' Copia una cotización a la Plantilla. Oculta las filas de conceptos que sobran
' para que el PDF no tenga un hueco entre los conceptos y los totales.
Public Sub FillTemplate(ByVal q As Quote)
    Dim ws As Worksheet, i As Long, r As Long, visibleRows As Long, vehicle As String

    On Error GoTo Fail
    Set ws = TemplateSheet()
    Application.ScreenUpdating = False
    UnlockSheet ws
    ClearTemplateContent ws

    ' Encabezado
    ws.Range(CELL_QUOTE_NUMBER).Value = q.Number
    ws.Range(CELL_DATE).Value = q.QuoteDate
    ws.Range(CELL_VALID_UNTIL).Value = q.ValidUntil(ValidityDays())
    ws.Range(CELL_CLIENT).Value = SafeCellText(q.ClientName)

    ' Vehículo de la orden ligada (el renglón queda vacío si no hay)
    If Len(q.OrderNumber) > 0 Then vehicle = VehicleText(LoadOrder(q.OrderNumber))
    If Len(vehicle) > 0 Then
        ws.Range(CELL_VEHICLE_LABEL).Value = TPL_VEHICLE
        ws.Range(CELL_VEHICLE).Value = SafeCellText(vehicle)
        ws.Range(CELL_VEHICLE).EntireRow.RowHeight = VEHICLE_ROW_HEIGHT
    End If

    ' Empresa emisora
    ws.Range(CELL_COMPANY_NAME).Value = SafeCellText(CompanyValue(COMPANY_NAME_CELL))
    ws.Range(CELL_COMPANY_CONTACT).Value = SafeCellText(CompanyContactLine())
    ws.Range(CELL_COMPANY_ADDRESS).Value = SafeCellText(CompanyAddressLine())

    ' Conceptos
    For i = 1 To q.ItemCount
        r = FIRST_ITEM_ROW + i - 1
        If r > LAST_ITEM_ROW Then Exit For            ' Seguridad: no invadir los totales
        ws.Cells(r, 1).Resize(1, 4).Value = Array(q.Item(i).Quantity, _
            SafeCellText(q.Item(i).Description), q.Item(i).UnitPrice, q.Item(i).Amount)
    Next i

    ' Filas visibles: las usadas, pero nunca menos del mínimo
    visibleRows = q.ItemCount
    If visibleRows < MIN_VISIBLE_ITEM_ROWS Then visibleRows = MIN_VISIBLE_ITEM_ROWS
    For r = FIRST_ITEM_ROW To LAST_ITEM_ROW
        If r - FIRST_ITEM_ROW + 1 > visibleRows Then
            ws.Rows(r).Hidden = True                  ' Las filas ocultas no salen en el PDF
        Else
            ws.Rows(r).AutoFit                        ' Crece si la descripción ocupa 2+ líneas
            If ws.Rows(r).RowHeight < MIN_ITEM_ROW_HEIGHT Then ws.Rows(r).RowHeight = MIN_ITEM_ROW_HEIGHT
        End If
    Next r

    ' Totales
    ws.Range(CELL_SUBTOTAL).Value = q.Subtotal
    ws.Range(CELL_TAX).Value = q.Tax
    ws.Range(CELL_TOTAL).Value = q.Total
    If q.TaxRate = 0 Then
        ws.Range(CELL_TAX_LABEL).Value = TPL_TAX_EXEMPT
    Else
        ws.Range(CELL_TAX_LABEL).Value = Fmt(TPL_TAX, FormatRate(q.TaxRate))
    End If
    ws.Range(CELL_AMOUNT_WORDS).Value = AmountToWords(q.Total)

    ApplyCompanyColors ws
Cleanup:
    LockSheet ws
    Application.ScreenUpdating = True
    RefreshColorSwatches
    Exit Sub
Fail:
    ShowError "FillTemplate"
    Resume Cleanup
End Sub

' Pinta la Plantilla con los colores de la empresa.
'   Principal: título, línea bajo el encabezado y banda del Total.
'   Secundario: encabezado de la tabla de conceptos.
Private Sub ApplyCompanyColors(ByVal ws As Worksheet)
    Dim primary As Long, secondary As Long
    primary = PrimaryColor()
    secondary = SecondaryColor()

    ws.Range("C1").Font.Color = primary
    With ws.Range("A5:D5").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = primary
    End With
    ws.Range("C28:D28").Interior.Color = primary
    ws.Range("C28:D28").Font.Color = ContrastTextColor(primary)   ' Legible sobre cualquier color
    ws.Range("A8:D8").Interior.Color = secondary
    ws.Range("A8:D8").Font.Color = ContrastTextColor(secondary)
End Sub

'--- PDF -----------------------------------------------------------------------

' Carpeta de los PDF (se crea si no existe). Si el libro está en OneDrive
' (ruta "http...") o no se ha guardado, usa Documentos.
Private Function PdfFolder() As String
    Dim base As String, folder As String

    base = ThisWorkbook.Path
    If Len(base) = 0 Or LCase$(Left$(base, 4)) = "http" Then
        base = Environ$("USERPROFILE") & "\Documents"
    End If
    folder = base & Application.PathSeparator & PDF_FOLDER
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    PdfFolder = folder
End Function

' Exporta la Plantilla como [Folio]_[Cliente].pdf.
' Regresa la ruta del PDF, o "" si no se generó.
Public Function ExportPdf(ByVal q As Quote) As String
    ExportPdf = ExportSheetPdf(TemplateSheet(), q.Number & "_" & q.ClientName)
End Function

' Exporta la PlantillaOrden como [Folio]_[Cliente].pdf en la misma carpeta.
' Regresa la ruta del PDF, o "" si no se generó.
Public Function ExportOrderPdf(ByVal o As ServiceOrder) As String
    ExportOrderPdf = ExportSheetPdf(OrderTemplateSheet(), o.Number & "_" & o.ClientName)
End Function

' Pregunta si el archivo ya existe, exporta la hoja y abre el PDF.
Private Function ExportSheetPdf(ByVal ws As Worksheet, ByVal baseName As String) As String
    Dim filePath As String

    On Error GoTo Fail
    filePath = PdfFolder() & Application.PathSeparator & SanitizeFileName(baseName) & ".pdf"

    If Len(Dir$(filePath)) > 0 Then
        If MsgBox(Fmt(MSG_PDF_EXISTS, filePath), vbYesNo + vbQuestion, TTL_PDF_EXISTS) = vbNo Then Exit Function
    End If

    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
    ExportSheetPdf = filePath
    Exit Function
Fail:
    MsgBox Fmt(MSG_PDF_FAILED, Err.Description), vbCritical, TTL_PDF_ERROR
End Function

' Antes del PDF: avisa si falta el nombre de la empresa o el RFC es inválido.
' Regresa True si se puede continuar.
Public Function ConfirmCompanyData() As Boolean
    Dim warning As String, rfc As String

    rfc = CompanyValue(COMPANY_RFC_CELL)
    If Len(CompanyValue(COMPANY_NAME_CELL)) = 0 Then
        warning = MSG_COMPANY_NAME_MISSING
    ElseIf Len(rfc) > 0 And Not IsValidRfc(rfc) Then
        warning = Fmt(MSG_COMPANY_RFC_INVALID, rfc)
    End If

    If Len(warning) = 0 Then
        ConfirmCompanyData = True
    Else
        ConfirmCompanyData = (MsgBox(Fmt(MSG_COMPANY_CONTINUE, warning), _
                                     vbYesNo + vbExclamation, TTL_COMPANY) = vbYes)
    End If
End Function

'--- Logos ---------------------------------------------------------------------

' Pide una imagen y la coloca como logo de la empresa (arriba a la izquierda).
' Lo usa el botón "Mi logo". Ideal: PNG con fondo transparente.
Public Sub InsertCompanyLogo()
    Dim filePath As Variant
    filePath = Application.GetOpenFilename(DLG_IMAGE_FILTER, , DLG_PICK_COMPANY_LOGO)
    If VarType(filePath) = vbBoolean Then Exit Sub        ' El usuario canceló
    PlaceLogo CStr(filePath), COMPANY_LOGO_NAME
End Sub

' [Administrador] Pide la imagen del logo de Cotiva y la fija abajo de la cotización.
Public Sub InsertCotivaLogo()
    Dim filePath As Variant
    filePath = Application.GetOpenFilename(DLG_IMAGE_FILTER, , DLG_PICK_COTIVA_LOGO)
    If VarType(filePath) = vbBoolean Then Exit Sub
    PlaceLogo CStr(filePath), COTIVA_LOGO_NAME
    MsgBox MSG_COTIVA_LOGO_DONE, vbInformation
End Sub

' Inserta (o reemplaza) una imagen como logo y la acomoda en su zona.
'   logoName: COMPANY_LOGO_NAME o COTIVA_LOGO_NAME.
' LinkToFile:=False y SaveWithDocument:=True -> la imagen queda dentro del libro.
Public Sub PlaceLogo(ByVal filePath As String, ByVal logoName As String)
    Dim ws As Worksheet, shp As Shape

    Set ws = TemplateSheet()
    UnlockSheet ws
    If ShapeExists(ws, logoName) Then ws.Shapes(logoName).Delete
    Set shp = ws.Shapes.AddPicture(filePath, msoFalse, msoTrue, 0, 0, -1, -1)
    shp.Name = logoName
    If logoName = COTIVA_LOGO_NAME Then ws.Range("C33").ClearContents   ' Quita el texto de respaldo
    FitLogos ws
    LockSheet ws
End Sub

' Acomoda el logo de la empresa (si existe) dentro de la zona, conservando la proporción.
Private Sub FitCompanyLogo(ByVal ws As Worksheet, ByVal zone As Range)
    Dim shp As Shape

    If Not ShapeExists(ws, COMPANY_LOGO_NAME) Then Exit Sub
    Set shp = ws.Shapes(COMPANY_LOGO_NAME)
    shp.LockAspectRatio = msoTrue
    shp.Height = zone.Height - 8
    If shp.Width > zone.Width * 0.6 Then shp.Width = zone.Width * 0.6
    shp.Left = zone.Left + 2
    shp.Top = zone.Top + (zone.Height - shp.Height) / 2
    shp.Placement = xlMove
End Sub

' Acomoda los dos logos en sus zonas conservando la proporción.
Private Sub FitLogos(ByVal ws As Worksheet)
    Dim shp As Shape, zone As Range

    FitCompanyLogo ws, ws.Range("A1:B1")

    If ShapeExists(ws, COTIVA_LOGO_NAME) Then
        Set shp = ws.Shapes(COTIVA_LOGO_NAME)
        Set zone = ws.Range("C33:D33")
        shp.LockAspectRatio = msoTrue
        shp.Height = COTIVA_LOGO_HEIGHT
        If shp.Width > zone.Width - 6 Then shp.Width = zone.Width - 6
        shp.Left = zone.Left + 4
        shp.Top = zone.Top + (zone.Height - shp.Height) / 2
        shp.Placement = xlMove
    End If
End Sub

'--- Diseño --------------------------------------------------------------------

' Aplica un borde inferior a un rango.
Private Sub BottomBorder(ByVal target As Range, ByVal lineColor As Long, _
                         Optional ByVal lineWeight As XlBorderWeight = xlThin)
    With target.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = lineWeight
        .Color = lineColor
    End With
End Sub

' Rediseña la Plantilla desde cero con la identidad Cotiva. Conserva los logos.
'   silent: True = sin preguntas ni mensajes (lo usa el build y la creación automática).
Public Sub DesignTemplate(Optional ByVal silent As Boolean = False)
    Dim ws As Worksheet, fontName As String, r As Long

    If Not silent Then
        If MsgBox(Fmt(MSG_DESIGN_CONFIRM), vbYesNo + vbQuestion, TTL_DESIGN) = vbNo Then Exit Sub
    End If

    On Error GoTo Fail
    Set ws = TemplateSheet()
    fontName = BrandFont()
    Application.ScreenUpdating = False
    UnlockSheet ws

    ' --- 1. Lienzo limpio ---
    ws.Range("A1:J60").Clear                      ' Valores, formatos y combinaciones
    ws.Rows("1:60").Hidden = False
    ws.Rows("1:60").RowHeight = 15
    With ws.Range("A1:J60").Font
        .Name = fontName
        .Size = 10
        .Color = BrandInk()
    End With
    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 50
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 18

    ' --- 2. Encabezado: empresa (izquierda) / título, folio y fechas (derecha) ---
    ws.Rows(1).RowHeight = 46
    ws.Rows(2).RowHeight = 18
    ws.Rows(3).RowHeight = 14
    ws.Rows(4).RowHeight = 14
    With ws.Range("C1:D1")
        .Merge
        .Value = TPL_TITLE
        .Font.Size = 20
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
    End With
    ws.Range("C2").Value = TPL_NUMBER
    ws.Range("C3").Value = TPL_DATE
    ws.Range("C4").Value = TPL_VALID_UNTIL
    With ws.Range("C2:C4")
        .Font.Size = 9
        .Font.Color = TextGray()
        .HorizontalAlignment = xlRight
    End With
    With ws.Range("D2:D4")
        .HorizontalAlignment = xlRight
        .Font.Bold = True
    End With
    ws.Range(CELL_QUOTE_NUMBER).Font.Size = 11
    ws.Range("D3:D4").NumberFormat = "[$-080A]dd mmm yyyy"     ' 19 sep 2026

    With ws.Range("A2:B2")
        .Merge
        .Font.Size = 11
        .Font.Bold = True
        .VerticalAlignment = xlBottom
        .ShrinkToFit = True                       ' Textos largos se encogen, no se cortan
    End With
    With ws.Range("A3:B3")
        .Merge
        .Font.Size = 8
        .Font.Color = TextGray()
        .ShrinkToFit = True
    End With
    With ws.Range("A4:B4")
        .Merge
        .Font.Size = 8
        .Font.Color = TextGray()
        .ShrinkToFit = True
    End With

    ' --- 3. Línea de color y cliente ---
    ws.Rows(5).RowHeight = 10
    ws.Rows(6).RowHeight = 30
    With ws.Range("A6")
        .Value = TPL_CLIENT
        .Font.Size = 9
        .Font.Color = TextGray()
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("B6:D6")
        .Merge
        .Font.Size = 13
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(7).RowHeight = VEHICLE_ROW_EMPTY_HEIGHT
    With ws.Range(CELL_VEHICLE_LABEL)
        .Font.Size = 9
        .Font.Color = TextGray()
        .VerticalAlignment = xlCenter
    End With
    With ws.Range(CELL_VEHICLE).Resize(1, 3)
        .Merge
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
    End With

    ' --- 4. Tabla de conceptos ---
    ws.Rows(8).RowHeight = 24
    ws.Range("A8:D8").Value = Array(TPL_COL_QTY, TPL_COL_DESCRIPTION, TPL_COL_PRICE, TPL_COL_AMOUNT)
    With ws.Range("A8:D8")
        .Font.Bold = True
        .Font.Size = 9
        .VerticalAlignment = xlCenter
    End With
    For r = FIRST_ITEM_ROW To LAST_ITEM_ROW
        ws.Rows(r).RowHeight = MIN_ITEM_ROW_HEIGHT
        If (r - FIRST_ITEM_ROW) Mod 2 = 1 Then ws.Range("A" & r & ":D" & r).Interior.Color = Neutral50()
        BottomBorder ws.Range("A" & r & ":D" & r), Neutral200()
    Next r
    ws.Range("A8:A" & LAST_ITEM_ROW).HorizontalAlignment = xlCenter
    With ws.Range("B8:B" & LAST_ITEM_ROW)
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    With ws.Range("C8:D" & LAST_ITEM_ROW)
        .HorizontalAlignment = xlRight
        .IndentLevel = 1
    End With
    ws.Range("A" & FIRST_ITEM_ROW & ":D" & LAST_ITEM_ROW).VerticalAlignment = xlCenter
    ws.Range("B" & FIRST_ITEM_ROW & ":B" & LAST_ITEM_ROW).WrapText = True
    ws.Range("C" & FIRST_ITEM_ROW & ":D" & LAST_ITEM_ROW).NumberFormat = "$#,##0.00"

    ' --- 5. Totales ---
    ws.Rows(26).RowHeight = 22
    ws.Rows(27).RowHeight = 22
    ws.Rows(28).RowHeight = 28
    ws.Range("C26").Value = TPL_SUBTOTAL
    ws.Range(CELL_TAX_LABEL).Value = Fmt(TPL_TAX, FormatRate(0.16))
    ws.Range("C28").Value = TPL_TOTAL
    With ws.Range("C26:C27")
        .Font.Color = TextGray()
        .Font.Size = 9
    End With
    With ws.Range("C26:D28")
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Range("D26:D28").NumberFormat = "$#,##0.00"
    BottomBorder ws.Range("C27:D27"), Neutral200()
    With ws.Range("C28:D28").Font
        .Bold = True
        .Size = 12
    End With

    ' --- 6. Importe con letra ---
    With ws.Range("A26")
        .Value = TPL_AMOUNT_WORDS
        .Font.Size = 8
        .Font.Bold = True
        .Font.Color = TextGray()
        .VerticalAlignment = xlBottom
    End With
    With ws.Range("A27:B28")
        .Merge
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Size = 9
    End With

    ' --- 7. Condiciones (editables aunque la hoja esté protegida) ---
    ws.Rows(29).RowHeight = 18
    ws.Rows(30).RowHeight = 16
    ws.Rows(31).RowHeight = 40
    With ws.Range("A30")
        .Value = TPL_TERMS_TITLE
        .Font.Size = 9
        .Font.Bold = True
    End With
    BottomBorder ws.Range("A30:D30"), Neutral200()
    With ws.Range("A31:D31")
        .Merge
        .Locked = False
        .Value = TPL_TERMS_DEFAULT
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Size = 9
        .Font.Color = TextGray()
    End With

    ' --- 8. Firma de Cotiva e innobytes ---
    ws.Rows(32).RowHeight = 14
    ws.Rows(33).RowHeight = COTIVA_LOGO_HEIGHT + 12
    With ws.Range("B33")
        .Value = TPL_MADE_WITH
        .Font.Size = 9
        .Font.Color = TextGray()
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    If Not ShapeExists(ws, COTIVA_LOGO_NAME) Then
        With ws.Range("C33")                      ' Texto de respaldo hasta que haya logo
            .Value = UI_BRAND
            .Font.Bold = True
            .Font.Size = 14
            .VerticalAlignment = xlCenter
        End With
    End If

    ws.Rows(34).RowHeight = 16
    With ws.Range("A34:D34")
        .Merge
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
    End With
    If Len(INNOBYTES_WEB) > 0 Then                ' La liga funciona dentro del PDF
        ws.Hyperlinks.Add Anchor:=ws.Range("A34"), Address:="https://" & INNOBYTES_WEB, _
                          TextToDisplay:=InnobytesCredit()
    Else
        ws.Range("A34").Value = InnobytesCredit()
    End If
    With ws.Range("A34:D34").Font                 ' Después de la liga, que cambia el estilo
        .Name = fontName
        .Size = 8
        .Color = TextGray()
        .Underline = xlUnderlineStyleNone
    End With

    ' --- 9. Colores de la empresa (título, línea, total y tabla) ---
    ApplyCompanyColors ws

    ' --- 10. Página ---
    ConfigurePage ws, fontName, PRINT_AREA, True

    FitLogos ws
    LockSheet ws
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    Application.ScreenUpdating = True

    If silent Then Exit Sub
    If Not ShapeExists(ws, COMPANY_LOGO_NAME) Then
        If MsgBox(MSG_DESIGN_ASK_LOGO, vbYesNo + vbQuestion, TTL_DESIGN) = vbYes Then InsertCompanyLogo
    Else
        MsgBox MSG_DESIGN_DONE, vbInformation, TTL_DESIGN
    End If
    Exit Sub
Fail:
    ShowError "DesignTemplate"
    Application.ScreenUpdating = True
End Sub

' Tamaño carta, márgenes, 1 página de ancho (y de alto si oneTallPage) y pie con número de página.
Private Sub ConfigurePage(ByVal ws As Worksheet, ByVal fontName As String, ByVal printArea As String, _
                          ByVal oneTallPage As Boolean)
    ' PrintCommunication=False acelera PageSetup; puede no existir en Excel muy viejo
    On Error Resume Next
    Application.PrintCommunication = False
    On Error GoTo 0

    With ws.PageSetup
        .PrintArea = printArea
        .Orientation = xlPortrait
        .PaperSize = xlPaperLetter
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.8)
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(1.5)
        .HeaderMargin = Application.CentimetersToPoints(0.8)
        .FooterMargin = Application.CentimetersToPoints(0.8)
        .CenterHorizontally = True
        .PrintGridlines = False
        .Zoom = False                             ' Necesario para que funcione FitTo...
        .FitToPagesWide = 1
        If oneTallPage Then .FitToPagesTall = 1 Else .FitToPagesTall = False
        .LeftFooter = "&""" & fontName & ",Regular""&8&K62666D" & TPL_PAGE
        .RightFooter = ""
    End With

    On Error Resume Next
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub

'--- Orden de servicio -----------------------------------------------------------

' Celdas de datos de la PlantillaOrden (las condiciones no están: se editan a mano).
Private Function OrderDataCells() As Variant
    OrderDataCells = Array(CELL_ORD_COMPANY_NAME, CELL_ORD_COMPANY_CONTACT, CELL_ORD_COMPANY_ADDRESS, _
        CELL_ORD_NUMBER, CELL_ORD_STATUS, CELL_ORD_PROMISED_HEAD, CELL_ORD_CLIENT, CELL_ORD_PHONE1, _
        CELL_ORD_EMAIL, CELL_ORD_PHONE2, CELL_ORD_CLIENT_ADDRESS, CELL_ORD_MAKE, CELL_ORD_MODEL, _
        CELL_ORD_YEAR, CELL_ORD_PLATES, CELL_ORD_COLOR, CELL_ORD_ENGINE, CELL_ORD_VIN, CELL_ORD_MILEAGE, _
        CELL_ORD_FUEL, CELL_ORD_IN, CELL_ORD_RECEIVED_BY, CELL_ORD_PROMISED, CELL_ORD_OUT, _
        CELL_ORD_DELIVERED_BY, CELL_ORD_REQUESTED, CELL_ORD_DIAGNOSIS, CELL_ORD_SIGN_CLIENT, _
        CELL_ORD_SIGN_RECEIVED, CELL_ORD_SIGN_DELIVERED)
End Function

' Escribe un texto de usuario en una celda de la orden.
Private Sub PutOrderText(ByVal ws As Worksheet, ByVal cellAddress As String, ByVal value As String)
    ws.Range(cellAddress).Value = SafeCellText(value)
End Sub

' Fecha y hora en una sola celda; lo que valga 0 queda fuera y sin ambos la celda queda en blanco.
Private Sub PutOrderDateTime(ByVal ws As Worksheet, ByVal cellAddress As String, _
                             ByVal d As Date, ByVal t As Date)
    With ws.Range(cellAddress)
        If d > 0 And t > 0 Then
            .NumberFormat = "[$-080A]dd mmm yyyy hh:mm"
            .Value = CDbl(Int(d)) + (CDbl(t) - Int(CDbl(t)))
        ElseIf d > 0 Then
            .NumberFormat = "[$-080A]dd mmm yyyy"
            .Value = CDbl(Int(d))
        ElseIf t > 0 Then
            .NumberFormat = "hh:mm"
            .Value = CDbl(t) - Int(CDbl(t))
        End If
    End With
End Sub

' Copia una orden a la PlantillaOrden. Las condiciones no se tocan.
Public Sub FillOrderTemplate(ByVal o As ServiceOrder)
    Dim ws As Worksheet

    On Error GoTo Fail
    Set ws = OrderTemplateSheet()
    Application.ScreenUpdating = False
    UnlockSheet ws
    ClearOrderContent ws

    PutOrderText ws, CELL_ORD_COMPANY_NAME, CompanyValue(COMPANY_NAME_CELL)
    PutOrderText ws, CELL_ORD_COMPANY_CONTACT, CompanyContactLine()
    PutOrderText ws, CELL_ORD_COMPANY_ADDRESS, CompanyAddressLine()
    PutOrderText ws, CELL_ORD_NUMBER, o.Number
    PutOrderText ws, CELL_ORD_STATUS, o.Status
    PutOrderDateTime ws, CELL_ORD_PROMISED_HEAD, o.PromisedDate, 0

    PutOrderText ws, CELL_ORD_CLIENT, o.ClientName
    PutOrderText ws, CELL_ORD_PHONE1, o.Phone1
    PutOrderText ws, CELL_ORD_EMAIL, o.Email
    PutOrderText ws, CELL_ORD_PHONE2, o.Phone2
    PutOrderText ws, CELL_ORD_CLIENT_ADDRESS, o.Address

    PutOrderText ws, CELL_ORD_MAKE, o.Make
    PutOrderText ws, CELL_ORD_MODEL, o.Model
    If o.Year > 0 Then ws.Range(CELL_ORD_YEAR).Value = o.Year
    PutOrderText ws, CELL_ORD_PLATES, o.Plates
    PutOrderText ws, CELL_ORD_COLOR, o.Color
    PutOrderText ws, CELL_ORD_ENGINE, o.Engine
    PutOrderText ws, CELL_ORD_VIN, o.Vin
    If o.Mileage > 0 Then
        ws.Range(CELL_ORD_MILEAGE).NumberFormat = "#,##0 ""km"""
        ws.Range(CELL_ORD_MILEAGE).Value = o.Mileage
    End If
    PutOrderText ws, CELL_ORD_FUEL, o.FuelLevel

    PutOrderDateTime ws, CELL_ORD_IN, o.InDate, o.InTime
    PutOrderText ws, CELL_ORD_RECEIVED_BY, o.ReceivedBy
    PutOrderDateTime ws, CELL_ORD_PROMISED, o.PromisedDate, 0
    PutOrderDateTime ws, CELL_ORD_OUT, o.OutDate, o.OutTime
    PutOrderText ws, CELL_ORD_DELIVERED_BY, o.DeliveredBy

    PutOrderText ws, CELL_ORD_REQUESTED, o.Requested
    FillOrderDiagnosis ws, o.Diagnosis

    PutOrderText ws, CELL_ORD_SIGN_CLIENT, o.ClientName
    PutOrderText ws, CELL_ORD_SIGN_RECEIVED, o.ReceivedBy
    PutOrderText ws, CELL_ORD_SIGN_DELIVERED, o.DeliveredBy

    ApplyOrderColors ws
    SyncOrderLogo ws
Cleanup:
    LockSheet ws
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    ShowError "FillOrderTemplate"
    Resume Cleanup
End Sub

' Borra solo las celdas de datos y deja los formatos; nunca las condiciones.
Private Sub ClearOrderContent(ByVal ws As Worksheet)
    Dim cellAddress As Variant
    For Each cellAddress In OrderDataCells()
        ws.Range(CStr(cellAddress)).MergeArea.ClearContents
    Next cellAddress
    ws.Range("A" & ORD_DIAG_FIRST_ROW & ":H" & (ORD_DIAG_FIRST_ROW + ORD_DIAG_ROWS - 1)).ClearContents
End Sub

' El diagnóstico va en grupos de renglones, uno por fila combinada. Cada fila se mide en la
' columna auxiliar con AutoFit (las combinadas no se autoajustan); las filas sin grupo se ocultan.
Private Sub FillOrderDiagnosis(ByVal ws As Worksheet, ByVal diagnosisText As String)
    Dim chunks As Collection, i As Long, r As Long, lastRow As Long, chunk As String, fitted As Double

    lastRow = ORD_DIAG_FIRST_ROW + ORD_DIAG_ROWS - 1
    ws.Rows(ORD_DIAG_FIRST_ROW & ":" & lastRow).Hidden = False
    ws.Range("A" & ORD_DIAG_FIRST_ROW & ":H" & lastRow).ClearContents
    Set chunks = DiagnosisChunks(diagnosisText)

    For i = 1 To ORD_DIAG_ROWS
        r = ORD_DIAG_FIRST_ROW + i - 1
        If i <= chunks.Count Then
            chunk = SafeCellText(CStr(chunks(i)))
            ws.Range("A" & r).Value = chunk
            ws.Range(ORD_MEASURE_COLUMN & r).Value = chunk
            ws.Rows(r).AutoFit
            fitted = ws.Rows(r).RowHeight
            ws.Range(ORD_MEASURE_COLUMN & r).ClearContents
            ws.Rows(r).RowHeight = fitted           ' Si no, la fila se reajusta al vaciar la medición
        ElseIf i > 1 Then
            ws.Rows(r).Hidden = True
        End If
    Next i
    If ws.Rows(ORD_DIAG_FIRST_ROW).RowHeight < ORD_DIAG_MIN_HEIGHT Then ws.Rows(ORD_DIAG_FIRST_ROW).RowHeight = ORD_DIAG_MIN_HEIGHT
End Sub

' Pinta la PlantillaOrden con el color principal de la empresa (título y línea).
Private Sub ApplyOrderColors(ByVal ws As Worksheet)
    Dim primary As Long
    primary = PrimaryColor()
    ws.Range("E1").Font.Color = primary
    BottomBorder ws.Range("A5:H5"), primary, xlMedium
End Sub

' Copia el logo de la Plantilla a la PlantillaOrden, así "Mi logo" sigue siendo la única fuente.
' Si el portapapeles falla la orden sale sin logo y sin error.
Private Sub SyncOrderLogo(ByVal ws As Worksheet)
    Dim source As Worksheet, shapesBefore As Long

    If ShapeExists(ws, COMPANY_LOGO_NAME) Then ws.Shapes(COMPANY_LOGO_NAME).Delete
    Set source = TemplateSheet()
    If Not ShapeExists(source, COMPANY_LOGO_NAME) Then Exit Sub

    shapesBefore = ws.Shapes.Count
    On Error Resume Next
    source.Shapes(COMPANY_LOGO_NAME).Copy
    ws.Paste Destination:=ws.Range("A1")
    Application.CutCopyMode = False
    On Error GoTo 0

    If ws.Shapes.Count <> shapesBefore + 1 Then Exit Sub
    ws.Shapes(ws.Shapes.Count).Name = COMPANY_LOGO_NAME
    FitCompanyLogo ws, ws.Range("A1:D1")
End Sub

'--- Diseño de la orden ----------------------------------------------------------

' Etiqueta (9 pt, gris) y valor (negritas; combinado si abarca varias celdas).
Private Sub AddOrderField(ByVal ws As Worksheet, ByVal labelCell As String, _
                          ByVal caption As String, ByVal valueRange As String)
    With ws.Range(labelCell)
        .Value = caption
        .Font.Size = 9
        .Font.Color = TextGray()
        .VerticalAlignment = xlCenter
    End With
    With ws.Range(valueRange)
        If .Cells.Count > 1 Then .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
    End With
End Sub

' Título de sección con filete inferior.
Private Sub AddOrderSection(ByVal ws As Worksheet, ByVal sectionRow As Long, ByVal caption As String)
    ws.Rows(sectionRow).RowHeight = 18
    With ws.Range("A" & sectionRow)
        .Value = caption
        .Font.Size = 9
        .Font.Bold = True
        .VerticalAlignment = xlBottom
    End With
    BottomBorder ws.Range("A" & sectionRow & ":H" & sectionRow), Neutral200()
End Sub

' Texto largo en una fila combinada A:H, alineado arriba.
Private Sub AddOrderTextBlock(ByVal ws As Worksheet, ByVal blockRow As Long, ByVal blockHeight As Double)
    ws.Rows(blockRow).RowHeight = blockHeight
    With ws.Range("A" & blockRow & ":H" & blockRow)
        .Merge
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .Font.Size = 10
    End With
End Sub

' Columna de medición: mismo ancho de letra que el diagnóstico y 95 % de su ancho, para que el
' alto medido en pantalla deje holgura al imprimir.
Private Sub SetupMeasureColumn(ByVal ws As Worksheet)
    Dim blockWidth As Double, col As Range
    blockWidth = ws.Range("A" & ORD_DIAG_FIRST_ROW & ":H" & ORD_DIAG_FIRST_ROW).Width
    Set col = ws.Columns(ORD_MEASURE_COLUMN)
    col.ColumnWidth = 10
    col.ColumnWidth = 0.95 * blockWidth * col.ColumnWidth / col.Width
    With ws.Range(ORD_MEASURE_COLUMN & ORD_DIAG_FIRST_ROW & ":" & ORD_MEASURE_COLUMN & (ORD_DIAG_FIRST_ROW + ORD_DIAG_ROWS - 1))
        .WrapText = True
        .Font.Size = 10
        .VerticalAlignment = xlTop
    End With
End Sub

' Redibuja la PlantillaOrden desde cero con la identidad Cotiva. Conserva el logo.
'   silent: True = sin preguntas ni mensajes (lo usa el build y la creación automática).
Public Sub DesignOrderTemplate(Optional ByVal silent As Boolean = False)
    Dim ws As Worksheet, fontName As String, r As Variant

    If Not silent Then
        If MsgBox(Fmt(MSG_ORDER_DESIGN_CONFIRM), vbYesNo + vbQuestion, TTL_DESIGN) = vbNo Then Exit Sub
    End If

    On Error GoTo Fail
    Set ws = OrderTemplateSheet()
    fontName = BrandFont()
    Application.ScreenUpdating = False
    UnlockSheet ws

    ' --- 1. Lienzo limpio y cuadrícula de 8 columnas ---
    ws.Range("A1:J40").Clear
    ws.Rows("1:40").Hidden = False
    ws.Rows("1:40").RowHeight = 20
    With ws.Range("A1:J40").Font
        .Name = fontName
        .Size = 10
        .Color = BrandInk()
    End With
    For Each r In Array("A", "D", "G")
        ws.Columns(CStr(r)).ColumnWidth = 11
    Next r
    For Each r In Array("B", "E", "H")
        ws.Columns(CStr(r)).ColumnWidth = 19
    Next r
    ws.Columns("C").ColumnWidth = 2
    ws.Columns("F").ColumnWidth = 2

    ' --- 2. Encabezado: logo y empresa (izquierda), título y folio (derecha) ---
    ws.Rows(1).RowHeight = 46
    ws.Rows(2).RowHeight = 18
    ws.Rows(3).RowHeight = 14
    ws.Rows(4).RowHeight = 14
    ws.Rows(5).RowHeight = 10
    With ws.Range("E1:H1")
        .Merge
        .Value = TPL_ORD_TITLE
        .Font.Size = 20
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
    End With
    ws.Range("A1:D1").Merge
    With ws.Range("A2:D2")
        .Merge
        .Font.Size = 11
        .Font.Bold = True
        .VerticalAlignment = xlBottom
        .ShrinkToFit = True
    End With
    For Each r In Array("A3:D3", "A4:D4")
        With ws.Range(CStr(r))
            .Merge
            .Font.Size = 8
            .Font.Color = TextGray()
            .ShrinkToFit = True
        End With
    Next r
    ws.Range("G2").Value = TPL_NUMBER
    ws.Range("G3").Value = TPL_ORD_STATUS
    ws.Range("G4").Value = TPL_ORD_PROMISED
    With ws.Range("G2:G4")
        .Font.Size = 9
        .Font.Color = TextGray()
        .HorizontalAlignment = xlRight
    End With
    With ws.Range("H2:H4")
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .ShrinkToFit = True
    End With
    ws.Range(CELL_ORD_NUMBER).Font.Size = 11
    ws.Range(CELL_ORD_PROMISED_HEAD).NumberFormat = "[$-080A]dd mmm yyyy"

    ' --- 3. Cliente ---
    AddOrderSection ws, 6, TPL_ORD_SEC_CLIENT
    AddOrderField ws, "A7", TPL_ORD_NAME, "B7:E7"
    AddOrderField ws, "G7", TPL_ORD_PHONE1, "H7"
    AddOrderField ws, "A8", TPL_ORD_EMAIL, "B8:E8"
    AddOrderField ws, "G8", TPL_ORD_PHONE2, "H8"
    AddOrderField ws, "A9", TPL_ORD_ADDRESS, "B9:H9"
    ws.Rows(10).RowHeight = 8

    ' --- 4. Vehículo ---
    AddOrderSection ws, 11, TPL_ORD_SEC_VEHICLE
    AddOrderField ws, "A12", TPL_ORD_MAKE, "B12"
    AddOrderField ws, "D12", TPL_ORD_MODEL, "E12"
    AddOrderField ws, "G12", TPL_ORD_YEAR, "H12"
    AddOrderField ws, "A13", TPL_ORD_PLATES, "B13"
    AddOrderField ws, "D13", TPL_ORD_COLOR, "E13"
    AddOrderField ws, "G13", TPL_ORD_ENGINE, "H13"
    AddOrderField ws, "A14", TPL_ORD_VIN, "B14"
    AddOrderField ws, "D14", TPL_ORD_MILEAGE, "E14"
    AddOrderField ws, "G14", TPL_ORD_FUEL, "H14"
    ws.Rows(15).RowHeight = 8

    ' --- 5. Entrada y salida ---
    AddOrderSection ws, 16, TPL_ORD_SEC_VISIT
    AddOrderField ws, "A17", TPL_ORD_IN, "B17"
    AddOrderField ws, "D17", TPL_ORD_RECEIVED_BY, "E17"
    AddOrderField ws, "G17", TPL_ORD_PROMISED, "H17"
    AddOrderField ws, "A18", TPL_ORD_OUT, "B18"
    AddOrderField ws, "D18", TPL_ORD_DELIVERED_BY, "E18"
    ws.Rows(19).RowHeight = 8

    ' Los datos de texto se guardan como texto (teléfonos, placas y VIN no son números)
    For Each r In Array(CELL_ORD_PHONE1, CELL_ORD_PHONE2, CELL_ORD_MODEL, CELL_ORD_PLATES, _
                        CELL_ORD_COLOR, CELL_ORD_ENGINE, CELL_ORD_VIN, CELL_ORD_FUEL)
        ws.Range(CStr(r)).NumberFormat = "@"
    Next r
    ws.Range(CELL_ORD_YEAR).NumberFormat = "0"
    ws.Range(CELL_ORD_YEAR).HorizontalAlignment = xlLeft

    ' --- 6. Servicio solicitado y diagnóstico ---
    AddOrderSection ws, 20, TPL_ORD_SEC_REQUESTED
    AddOrderTextBlock ws, 21, 48
    AddOrderSection ws, 22, TPL_ORD_SEC_DIAGNOSIS
    For r = ORD_DIAG_FIRST_ROW To ORD_DIAG_FIRST_ROW + ORD_DIAG_ROWS - 1
        AddOrderTextBlock ws, CLng(r), ORD_DIAG_MIN_HEIGHT
        If r > ORD_DIAG_FIRST_ROW Then ws.Rows(CLng(r)).Hidden = True
    Next r
    ws.Rows(ORD_DIAG_FIRST_ROW + ORD_DIAG_ROWS).RowHeight = 8
    SetupMeasureColumn ws

    ' --- 7. Condiciones (editables aunque la hoja esté protegida) ---
    AddOrderSection ws, 32, TPL_ORD_SEC_TERMS
    AddOrderTextBlock ws, 33, 48
    With ws.Range("A33:H33")
        .Locked = False
        .Value = TPL_ORD_TERMS_DEFAULT
        .Font.Size = 9
        .Font.Color = TextGray()
    End With

    ' --- 8. Firmas ---
    ws.Rows(34).RowHeight = 40
    ws.Rows(35).RowHeight = 16
    ws.Rows(36).RowHeight = 14
    For Each r In Array("A", "D", "G")
        With ws.Range(CStr(r) & "35:" & Chr$(Asc(CStr(r)) + 1) & "35")
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .ShrinkToFit = True
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeTop).Color = TextGray()
        End With
        With ws.Range(CStr(r) & "36:" & Chr$(Asc(CStr(r)) + 1) & "36")
            .Merge
            .HorizontalAlignment = xlCenter
            .Font.Size = 8
            .Font.Color = TextGray()
        End With
    Next r
    ws.Range("A36").Value = TPL_CLIENT
    ws.Range("D36").Value = TPL_ORD_RECEIVED_BY
    ws.Range("G36").Value = TPL_ORD_DELIVERED_BY

    ' --- 9. Colores de la empresa y página ---
    ApplyOrderColors ws
    ConfigurePage ws, fontName, ORDER_PRINT_AREA, False

    LockSheet ws
    Application.Goto ws.Range("A1")           ' Para apagar la cuadrícula de la ventana
    ActiveWindow.DisplayGridlines = False
    Application.ScreenUpdating = True

    If Not silent Then MsgBox MSG_ORDER_DESIGN_DONE, vbInformation, TTL_DESIGN
    Exit Sub
Fail:
    ShowError "DesignOrderTemplate"
    Application.ScreenUpdating = True
End Sub
