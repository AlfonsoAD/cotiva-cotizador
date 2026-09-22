Attribute VB_Name = "modTemplate"
'==============================================================================
' modTemplate · Hoja Plantilla: diseño, llenado, colores, logos y PDF
'
' Mapa de la hoja:
'   A1:B1  logo de la empresa       C1:D1  "COTIZACIÓN"
'   A2:B2  nombre de la empresa     C2/D2  Folio
'   A3:B3  RFC · Tel. · correo      C3/D3  Fecha
'   A4:B4  dirección · web          C4/D4  Vigencia
'   Fila 5 línea de color           A6 / B6:D6  Cliente
'   Fila 8 encabezado de tabla      A9:D25 conceptos
'   A26:B28 importe con letra       C26:D28 totales
'   A30:D31 condiciones (editable)  B33:D33 "Generado con" + logo Cotiva
'   A34:D34 créditos de innobytes (con liga)
'==============================================================================
Option Explicit

Private Function TemplateSheet() As Worksheet
    Set TemplateSheet = ThisWorkbook.Worksheets(SHEET_TEMPLATE)
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
End Sub

' Copia una cotización a la Plantilla. Oculta las filas de conceptos que sobran
' para que el PDF no tenga un hueco entre los conceptos y los totales.
Public Sub FillTemplate(ByVal q As Quote)
    Dim ws As Worksheet, i As Long, r As Long, visibleRows As Long

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
    Dim filePath As String

    On Error GoTo Fail
    filePath = PdfFolder() & Application.PathSeparator & _
           SanitizeFileName(q.Number & "_" & q.ClientName) & ".pdf"

    If Len(Dir$(filePath)) > 0 Then
        If MsgBox(Fmt(MSG_PDF_EXISTS, filePath), vbYesNo + vbQuestion, TTL_PDF_EXISTS) = vbNo Then Exit Function
    End If

    TemplateSheet().ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
    ExportPdf = filePath
    Exit Function
Fail:
    MsgBox Fmt(MSG_PDF_FAILED, Err.Description), vbCritical, TTL_PDF_ERROR
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

' Acomoda los dos logos en sus zonas conservando la proporción.
Private Sub FitLogos(ByVal ws As Worksheet)
    Dim shp As Shape, zone As Range

    If ShapeExists(ws, COMPANY_LOGO_NAME) Then
        Set shp = ws.Shapes(COMPANY_LOGO_NAME)
        Set zone = ws.Range("A1:B1")
        shp.LockAspectRatio = msoTrue
        shp.Height = zone.Height - 8
        If shp.Width > zone.Width * 0.6 Then shp.Width = zone.Width * 0.6
        shp.Left = zone.Left + 2
        shp.Top = zone.Top + (zone.Height - shp.Height) / 2
        shp.Placement = xlMove
    End If

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
    ws.Rows(7).RowHeight = 10

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
    ConfigurePage ws, fontName

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

' Tamaño carta, márgenes, 1 página y pie con número de página.
Private Sub ConfigurePage(ByVal ws As Worksheet, ByVal fontName As String)
    ' PrintCommunication=False acelera PageSetup; puede no existir en Excel muy viejo
    On Error Resume Next
    Application.PrintCommunication = False
    On Error GoTo 0

    With ws.PageSetup
        .PrintArea = PRINT_AREA
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
        .FitToPagesTall = 1
        .LeftFooter = "&""" & fontName & ",Regular""&8&K62666D" & TPL_PAGE
        .RightFooter = ""
    End With

    On Error Resume Next
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub
