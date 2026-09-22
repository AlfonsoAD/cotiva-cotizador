Attribute VB_Name = "modHome"
'==============================================================================
' modHome — Hoja "Inicio": portada con el botón y la marca (Cotiva · innobytes)
'
' Mapa de la hoja:
'   B2       logo de Cotiva (imagen) o el texto "Cotiva"
'   B3       regla naranja
'   B4 / B5  título y subtítulo
'   B7       botón "Abrir cotizador"
'   B8       nota de primer uso
'   B10:B12  lo que hace el cotizador
'   B14      línea separadora
'   B15      logo de innobytes (imagen) o el texto "innobytes"   C15  versión
'   B16      créditos de innobytes (con liga a innobytes.tech)
'
' DesignHome vive aquí y no en modBuild para poder rediseñar la portada en un
' libro ya distribuido, igual que DesignTemplate con la Plantilla.
'==============================================================================
Option Explicit

Private Function HomeSheet() As Worksheet
    Set HomeSheet = ThisWorkbook.Worksheets(SHEET_HOME)
End Function

'--- Diseño --------------------------------------------------------------------

' [Administrador] Rediseña la hoja Inicio desde cero. Conserva los logos.
'   silent: True = sin preguntas ni mensajes (lo usa el build).
Public Sub DesignHome(Optional ByVal silent As Boolean = False)
    Dim ws As Worksheet, fontName As String, features As Variant, i As Long, r As Long

    If Not SheetExists(SHEET_HOME) Then Exit Sub
    If Not silent Then
        If MsgBox(Fmt(MSG_DESIGN_HOME_CONFIRM), vbYesNo + vbQuestion, TTL_DESIGN) = vbNo Then Exit Sub
    End If

    On Error GoTo Fail
    Set ws = HomeSheet()
    fontName = BrandFont()
    Application.ScreenUpdating = False
    UnlockSheet ws

    ' --- 1. Lienzo limpio (las imágenes de los logos se conservan) ---
    RemoveDesignShapes ws
    ws.Range("A1:H40").Clear                      ' Valores, formatos, ligas y combinaciones
    ws.Rows("1:40").RowHeight = 15
    With ws.Range("A1:H40").Font
        .Name = fontName
        .Size = 10
        .Color = BrandInk()
    End With
    ws.Columns("A").ColumnWidth = 3               ' Margen izquierdo
    ws.Columns("B").ColumnWidth = 58
    ws.Columns("C").ColumnWidth = 32

    ' --- 2. Logo de Cotiva y regla de color ---
    ws.Rows(1).RowHeight = 18
    ws.Rows(2).RowHeight = HOME_COTIVA_LOGO_HEIGHT + 8
    If Not ShapeExists(ws, HOME_COTIVA_LOGO_NAME) Then
        With ws.Range("B2")                       ' Texto de respaldo hasta que haya logo
            .Value = UI_BRAND
            .Font.Size = 22
            .Font.Bold = True
            .VerticalAlignment = xlCenter
        End With
    End If

    ws.Rows(3).RowHeight = 14
    AddBar ws, HOME_RULE_NAME, ws.Range("B3"), 64, 4, BrandOrange()

    ' --- 3. Título y subtítulo ---
    ws.Rows(4).RowHeight = 38
    With ws.Range("B4:C4")
        .Merge
        .Value = HOME_TITLE
        .Font.Size = 26
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(5).RowHeight = 20
    With ws.Range("B5:C5")
        .Merge
        .Value = HOME_SUBTITLE
        .Font.Size = 11
        .Font.Color = TextGray()
        .VerticalAlignment = xlCenter
    End With

    ' --- 4. Botón que abre el cotizador ---
    ws.Rows(6).RowHeight = 16
    ws.Rows(7).RowHeight = 54
    AddOpenButton ws, fontName

    ws.Rows(8).RowHeight = 18
    With ws.Range("B8:C8")
        .Merge
        .Value = HOME_FIRST_TIME
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = TextGray()
        .VerticalAlignment = xlCenter
    End With

    ' --- 5. Qué hace el cotizador ---
    ws.Rows(9).RowHeight = 14
    features = Array(HOME_FEATURE_1, HOME_FEATURE_2, HOME_FEATURE_3)
    For i = LBound(features) To UBound(features)
        r = HOME_FIRST_FEATURE_ROW + i
        ws.Rows(r).RowHeight = 19
        With ws.Range("B" & r & ":C" & r)
            .Merge
            .Value = HOME_BULLET & "   " & CStr(features(i))
            .Font.Size = 10
            .VerticalAlignment = xlCenter
        End With
        ws.Range("B" & r).Characters(1, Len(HOME_BULLET)).Font.Color = BrandOrange()
    Next i

    ' --- 6. Separador ---
    ws.Rows(13).RowHeight = 20
    ws.Rows(14).RowHeight = 6
    With ws.Range("B14:C14").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = Neutral200()
    End With

    ' --- 7. Pie: innobytes y versión ---
    ws.Rows(15).RowHeight = HOME_INNOBYTES_LOGO_HEIGHT + 14
    If Not ShapeExists(ws, HOME_INNOBYTES_LOGO_NAME) Then
        With ws.Range("B15")                      ' Texto de respaldo hasta que haya logo
            .Value = HOME_MAKER
            .Font.Size = 13
            .Font.Bold = True
            .VerticalAlignment = xlCenter
        End With
    End If
    With ws.Range("C15")
        .Value = Fmt(HOME_VERSION, APP_VERSION)
        .Font.Size = 9
        .Font.Color = TextGray()
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With

    ws.Rows(16).RowHeight = 16
    With ws.Range("B16:C16")
        .Merge
        .VerticalAlignment = xlTop
    End With
    If Len(INNOBYTES_WEB) > 0 Then
        ws.Hyperlinks.Add Anchor:=ws.Range("B16"), Address:="https://" & INNOBYTES_WEB, _
                          TextToDisplay:=InnobytesCredit()
    Else
        ws.Range("B16").Value = InnobytesCredit()
    End If
    With ws.Range("B16:C16").Font                 ' Después de la liga, que cambia el estilo
        .Name = fontName
        .Size = 9
        .Color = TextGray()
        .Underline = xlUnderlineStyleNone
    End With

    ' --- 8. Vista limpia ---
    FitHomeLogos ws
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False

Cleanup:
    LockSheet ws
    Application.ScreenUpdating = True
    If Not silent Then MsgBox MSG_DESIGN_HOME_DONE, vbInformation, TTL_DESIGN
    Exit Sub
Fail:
    ShowError "DesignHome"
    Resume Cleanup
End Sub

' Botón (rectángulo redondeado) con la macro OpenQuoteTool.
Private Sub AddOpenButton(ByVal ws As Worksheet, ByVal fontName As String)
    Dim btn As Shape

    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("B7").Left, _
                                 ws.Range("B7").Top + 4, 200, 46)
    With btn
        .Name = HOME_BUTTON_NAME
        .Fill.ForeColor.RGB = BrandOrange()
        .Line.Visible = msoFalse
        .Placement = xlMove
        .OnAction = "OpenQuoteTool"
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.Text = HOME_BUTTON
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Size = 13
            .TextRange.Font.Name = fontName
            .TextRange.Font.Fill.ForeColor.RGB = vbWhite
        End With
    End With
End Sub

' Barra de color (la regla bajo el logo): un rectángulo sin borde.
Private Sub AddBar(ByVal ws As Worksheet, ByVal barName As String, ByVal anchor As Range, _
                   ByVal barWidth As Double, ByVal barHeight As Double, ByVal barColor As Long)
    Dim bar As Shape

    Set bar = ws.Shapes.AddShape(msoShapeRectangle, anchor.Left, anchor.Top + 4, barWidth, barHeight)
    With bar
        .Name = barName
        .Fill.ForeColor.RGB = barColor
        .Line.Visible = msoFalse
        .Placement = xlMove
    End With
End Sub

' Borra las formas del diseño (botón y regla) y conserva los logos.
Private Sub RemoveDesignShapes(ByVal ws As Worksheet)
    Dim i As Long, shapeName As String

    For i = ws.Shapes.Count To 1 Step -1
        shapeName = ws.Shapes(i).Name
        If shapeName <> HOME_COTIVA_LOGO_NAME And shapeName <> HOME_INNOBYTES_LOGO_NAME Then
            ws.Shapes(i).Delete
        End If
    Next i
End Sub

'--- Logos ---------------------------------------------------------------------

' [Administrador] Cambia el logo de Cotiva de la pantalla de inicio.
Public Sub InsertHomeCotivaLogo()
    PickHomeLogo HOME_COTIVA_LOGO_NAME, DLG_PICK_COTIVA_HOME_LOGO
End Sub

' [Administrador] Cambia el logo de innobytes del pie de la pantalla de inicio.
Public Sub InsertInnobytesLogo()
    PickHomeLogo HOME_INNOBYTES_LOGO_NAME, DLG_PICK_INNOBYTES_LOGO
End Sub

' Pide una imagen y la coloca como logo de la portada.
Private Sub PickHomeLogo(ByVal logoName As String, ByVal dialogTitle As String)
    Dim filePath As Variant

    filePath = Application.GetOpenFilename(DLG_IMAGE_FILTER, , dialogTitle)
    If VarType(filePath) = vbBoolean Then Exit Sub       ' El usuario canceló
    PlaceHomeLogo CStr(filePath), logoName
    MsgBox MSG_HOME_LOGO_DONE, vbInformation, TTL_DESIGN
End Sub

' Inserta (o reemplaza) una imagen en la portada y la acomoda en su zona.
'   logoName: HOME_COTIVA_LOGO_NAME o HOME_INNOBYTES_LOGO_NAME.
' LinkToFile:=False y SaveWithDocument:=True -> la imagen queda dentro del libro.
Public Sub PlaceHomeLogo(ByVal filePath As String, ByVal logoName As String)
    Dim ws As Worksheet, shp As Shape

    If Not SheetExists(SHEET_HOME) Then Exit Sub
    Set ws = HomeSheet()
    UnlockSheet ws
    If ShapeExists(ws, logoName) Then ws.Shapes(logoName).Delete
    Set shp = ws.Shapes.AddPicture(filePath, msoFalse, msoTrue, 0, 0, -1, -1)
    shp.Name = logoName

    ' Con imagen ya no hace falta el texto de respaldo
    If logoName = HOME_COTIVA_LOGO_NAME Then ws.Range("B2").ClearContents
    If logoName = HOME_INNOBYTES_LOGO_NAME Then ws.Range("B15").ClearContents

    FitHomeLogos ws
    LockSheet ws
End Sub

' Acomoda los logos en sus zonas conservando la proporción.
Private Sub FitHomeLogos(ByVal ws As Worksheet)
    FitLogoInZone ws, HOME_COTIVA_LOGO_NAME, ws.Range("B2"), HOME_COTIVA_LOGO_HEIGHT
    FitLogoInZone ws, HOME_INNOBYTES_LOGO_NAME, ws.Range("B15"), HOME_INNOBYTES_LOGO_HEIGHT
End Sub

' Escala un logo a la altura pedida y lo centra verticalmente en su celda.
Private Sub FitLogoInZone(ByVal ws As Worksheet, ByVal logoName As String, _
                          ByVal zone As Range, ByVal logoHeight As Double)
    Dim shp As Shape

    If Not ShapeExists(ws, logoName) Then Exit Sub
    Set shp = ws.Shapes(logoName)
    shp.LockAspectRatio = msoTrue
    shp.Height = logoHeight
    If shp.Width > zone.Width - 4 Then shp.Width = zone.Width - 4   ' No invadir la otra columna
    shp.Left = zone.Left
    shp.Top = zone.Top + (zone.Height - shp.Height) / 2
    shp.Placement = xlMove
End Sub
