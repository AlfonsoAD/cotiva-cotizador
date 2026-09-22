'==============================================================================
' frmQuote · Formulario del cotizador
'
' Todo el diseño se construye por código en BuildLayout: el formulario en el
' editor está vacío (salvo el logo opcional imgLogo que agrega el build).
' Ventajas: el diseño vive en texto (se versiona en Git) y no hay controles
' "sueltos" que se puedan desacomodar.
'
' Este archivo contiene SOLO el código del formulario; build.ps1 crea el
' formulario y le inserta este código.
'==============================================================================
Option Explicit

'--- Estado de la cotización ---
Private mQuote As Quote                 ' Cotización en pantalla
Private mEditIndex As Long              ' 0 = no se edita; >0 = posición del concepto en edición
Private mIsSaved As Boolean             ' True si el folio ya existe en la BD
Private mIsDirty As Boolean             ' True si hay cambios sin guardar
Private mSuppressEvents As Boolean      ' True mientras el CÓDIGO llena campos (no cuenta como cambio)
Private mOverwriteConfirmed As Boolean  ' True si ya se confirmó sobrescribir este folio
Private mFont As String                 ' Fuente de marca

'--- Controles sin eventos ---
Private txtNumber As MSForms.TextBox
Private txtDescription As MSForms.TextBox
Private lblAmount As MSForms.Label
Private lblSubtotal As MSForms.Label
Private lblTax As MSForms.Label
Private lblTotal As MSForms.Label
Private lblValidUntil As MSForms.Label

'--- Controles con eventos (WithEvents permite recibir sus eventos) ---
Private WithEvents cboClient As MSForms.ComboBox
Private WithEvents txtDate As MSForms.TextBox
Private WithEvents txtQty As MSForms.TextBox
Private WithEvents txtPrice As MSForms.TextBox
Private WithEvents lstItems As MSForms.ListBox
Private WithEvents cboTax As MSForms.ComboBox
Private WithEvents cboHistory As MSForms.ComboBox
Private WithEvents btnLoad As MSForms.CommandButton
Private WithEvents btnAdd As MSForms.CommandButton
Private WithEvents btnEdit As MSForms.CommandButton
Private WithEvents btnDelete As MSForms.CommandButton
Private WithEvents btnNew As MSForms.CommandButton
Private WithEvents btnLogo As MSForms.CommandButton
Private WithEvents btnCompany As MSForms.CommandButton
Private WithEvents btnPreview As MSForms.CommandButton
Private WithEvents btnSave As MSForms.CommandButton
Private WithEvents btnPdf As MSForms.CommandButton
Private WithEvents lblInnobytes As MSForms.Label

'--- Medidas del diseño (puntos) ---
Private Const FORM_WIDTH As Single = 660
Private Const FORM_HEIGHT As Single = 496
Private Const MARGIN As Single = 18

' Estilos de botón
Private Enum ButtonKind
    kindPrimary = 0         ' Naranja: acción principal
    kindDark = 1            ' Tinta
    kindSecondary = 2       ' Blanco
End Enum


'==============================================================================
' ARRANQUE Y CIERRE
'==============================================================================

Private Sub UserForm_Initialize()
    BuildLayout
    LoadHistoryList
    LoadClientList
    StartNewQuote
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mIsDirty Then
        If MsgBox(MSG_UNSAVED_CLOSE, vbYesNo + vbExclamation, TTL_CLOSE) = vbNo Then Cancel = True
    End If
End Sub


'==============================================================================
' DISEÑO (se construye por código)
'==============================================================================

Private Sub BuildLayout()
    Dim rightEdge As Single, hasLogo As Boolean

    mFont = BrandFont()
    rightEdge = FORM_WIDTH - MARGIN

    ' --- Ventana (la diferencia Width-InsideWidth son los bordes de Windows) ---
    Me.Caption = UI_FORM_CAPTION
    Me.BackColor = vbWhite
    Me.Width = FORM_WIDTH + (Me.Width - Me.InsideWidth)
    Me.Height = FORM_HEIGHT + (Me.Height - Me.InsideHeight)

    ' --- Barra superior: clara si hay logo (JPG con fondo blanco), tinta si no ---
    hasLogo = PlaceHeaderLogo()
    If hasLogo Then
        AddBand "c_Bar", 0, 0, FORM_WIDTH, 48, vbWhite, True
        AddBand "c_BarLine", 0, 48, FORM_WIDTH, 2, BrandOrange(), False
        AddLabel "c_Module", UI_MODULE, 142, 17, 120, 18, 11, True, TextGray()
    Else
        AddBand "c_Bar", 0, 0, FORM_WIDTH, 48, BrandInk(), True
        AddLabel "c_Brand", UI_BRAND, MARGIN, 12, 80, 24, 16, True, vbWhite
        AddLabel "c_Module", UI_MODULE, 104, 17, 120, 18, 11, False, BrandInkLight()
    End If

    Set cboHistory = AddComboBox("c_History", rightEdge - 72 - 6 - 190, 13, 190, 22)
    With cboHistory
        .ColumnCount = 2                          ' Folio | Cliente
        .ColumnWidths = "70 pt;150 pt"
        .ListWidth = 230
        .BoundColumn = 1
        .TextColumn = 1
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryComplete
        .ControlTipText = TIP_HISTORY
    End With
    Set btnLoad = AddButton("c_Load", UI_BTN_LOAD, kindSecondary, rightEdge - 72, 13, 72, 22)

    ' --- Sección 1: datos de la cotización ---
    AddLabel "c_Sec1", UI_SECTION_QUOTE, MARGIN, 62, 300, 12, 8, True, BrandOrange()
    AddLabel "c_LNumber", UI_LBL_NUMBER, MARGIN, 80, 90, 12, 8, False, TextGray()
    AddLabel "c_LClient", UI_LBL_CLIENT, 116, 80, 250, 12, 8, False, TextGray()
    AddLabel "c_LDate", UI_LBL_DATE, 524, 80, 118, 12, 8, False, TextGray()

    Set txtNumber = AddTextBox("c_Number", MARGIN, 94, 90, 22)
    With txtNumber
        .Locked = True                            ' El folio no se edita a mano
        .TabStop = False
        .BackColor = Neutral50()
        .Font.Bold = True
    End With

    Set cboClient = AddComboBox("c_Client", 116, 94, 400, 22)
    With cboClient
        .Style = fmStyleDropDownCombo             ' Se puede escribir un cliente nuevo
        .MatchEntry = fmMatchEntryComplete        ' Completa mientras escribes
        .MatchRequired = False
        .MaxLength = 100
        .ListRows = 10
        .ControlTipText = TIP_CLIENT
    End With

    Set txtDate = AddTextBox("c_Date", 524, 94, 118, 22)
    txtDate.MaxLength = 10
    Set lblValidUntil = AddLabel("c_ValidUntil", "", 524, 118, 118, 12, 8, False, TextGray())

    ' --- Sección 2: captura de conceptos ---
    AddLabel "c_Sec2", UI_SECTION_ITEM, MARGIN, 132, 300, 12, 8, True, BrandOrange()
    AddLabel "c_LQty", UI_LBL_QTY, MARGIN, 150, 64, 12, 8, False, TextGray()
    AddLabel "c_LDesc", UI_LBL_DESCRIPTION, 90, 150, 200, 12, 8, False, TextGray()
    AddLabel "c_LPrice", UI_LBL_PRICE, 388, 150, 90, 12, 8, False, TextGray()
    AddLabel "c_LAmount", UI_LBL_AMOUNT, 486, 150, 76, 12, 8, False, TextGray()

    Set txtQty = AddTextBox("c_Qty", MARGIN, 164, 64, 22)
    txtQty.TextAlign = fmTextAlignRight
    txtQty.MaxLength = 12
    Set txtDescription = AddTextBox("c_Desc", 90, 164, 290, 22)
    txtDescription.MaxLength = 255
    Set txtPrice = AddTextBox("c_Price", 388, 164, 90, 22)
    txtPrice.TextAlign = fmTextAlignRight
    txtPrice.MaxLength = 15

    Set lblAmount = AddLabel("c_Amount", "", 486, 164, 76, 22, 10, True, BrandInk(), fmTextAlignRight)
    With lblAmount
        .BackStyle = fmBackStyleOpaque
        .BackColor = Neutral50()
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = Neutral200()
    End With
    Set btnAdd = AddButton("c_Add", UI_BTN_ADD, kindDark, 570, 164, 72, 22)

    ' --- Lista de conceptos con encabezado de columnas ---
    AddBand "c_ListHead", MARGIN, 198, FORM_WIDTH - 2 * MARGIN, 20, Neutral100(), False
    AddLabel "c_C1", UI_COL_QTY, 22, 202, 50, 12, 8, True, BrandInk()
    AddLabel "c_C2", UI_COL_DESCRIPTION, 78, 202, 200, 12, 8, True, BrandInk()
    AddLabel "c_C3", UI_COL_PRICE, 408, 202, 100, 12, 8, True, BrandInk()
    AddLabel "c_C4", UI_COL_AMOUNT, 518, 202, 100, 12, 8, True, BrandInk()

    Set lstItems = Me.Controls.Add("Forms.ListBox.1", "c_Items", True)
    Place lstItems, MARGIN, 218, FORM_WIDTH - 2 * MARGIN, 146
    StyleInput lstItems
    lstItems.ColumnCount = 4
    lstItems.ColumnWidths = "56 pt;330 pt;110 pt;110 pt"

    Set btnEdit = AddButton("c_Edit", UI_BTN_EDIT, kindSecondary, MARGIN, 372, 84, 24)
    Set btnDelete = AddButton("c_Delete", UI_BTN_DELETE, kindSecondary, MARGIN + 90, 372, 84, 24)

    ' --- Créditos de innobytes (clic = abrir el sitio) ---
    Set lblInnobytes = AddLabel("c_Innobytes", InnobytesCredit(), MARGIN, 418, 380, 12, 8, False, TextGray())
    lblInnobytes.ControlTipText = Fmt(TIP_INNOBYTES, INNOBYTES_WEB)
    lblInnobytes.MousePointer = fmMousePointerUpArrow

    ' --- Totales ---
    AddLabel "c_LSub", UI_LBL_SUBTOTAL, 436, 374, 100, 14, 9, False, TextGray(), fmTextAlignRight
    Set lblSubtotal = AddLabel("c_Subtotal", "", 540, 373, 102, 16, 10, False, BrandInk(), fmTextAlignRight)

    AddLabel "c_LTax", UI_LBL_TAX, 400, 392, 56, 14, 9, False, TextGray(), fmTextAlignRight
    Set cboTax = AddComboBox("c_TaxRate", 462, 389, 70, 19)
    With cboTax
        .Style = fmStyleDropDownList              ' Solo se puede elegir, no escribir
        .AddItem UI_TAX_16                        ' taxStandard
        .AddItem UI_TAX_8                         ' taxBorder
        .AddItem UI_TAX_EXEMPT                    ' taxExempt
        .Font.Size = 9
        .ControlTipText = TIP_TAX
    End With
    Set lblTax = AddLabel("c_Tax", "", 540, 391, 102, 16, 10, False, BrandInk(), fmTextAlignRight)

    AddBand "c_TotalLine", 436, 412, 206, 1, Neutral200(), False
    AddLabel "c_LTotal", UI_LBL_TOTAL, 436, 422, 100, 16, 11, True, BrandInk(), fmTextAlignRight
    Set lblTotal = AddLabel("c_Total", "", 520, 418, 122, 22, 15, True, BrandOrange(), fmTextAlignRight)

    ' --- Barra inferior de acciones ---
    AddBand "c_Footer", 0, 450, FORM_WIDTH, FORM_HEIGHT - 450, Neutral50(), True
    AddBand "c_FooterLine", 0, 450, FORM_WIDTH, 1, Neutral200(), False
    Set btnNew = AddButton("c_New", UI_BTN_NEW, kindSecondary, MARGIN, 460, 120, 26)
    Set btnLogo = AddButton("c_Logo", UI_BTN_LOGO, kindSecondary, MARGIN + 126, 460, 84, 26)
    btnLogo.ControlTipText = TIP_LOGO
    Set btnCompany = AddButton("c_Company", UI_BTN_COMPANY, kindSecondary, MARGIN + 216, 460, 90, 26)
    btnCompany.ControlTipText = TIP_COMPANY
    Set btnPreview = AddButton("c_Preview", UI_BTN_PREVIEW, kindSecondary, 330, 460, 100, 26)
    Set btnSave = AddButton("c_Save", UI_BTN_SAVE, kindPrimary, 436, 460, 100, 26)
    Set btnPdf = AddButton("c_Pdf", UI_BTN_PDF, kindDark, 542, 460, 100, 26)

    ' --- Orden del tabulador ---
    cboClient.TabIndex = 0
    txtDate.TabIndex = 1
    txtQty.TabIndex = 2
    txtDescription.TabIndex = 3
    txtPrice.TabIndex = 4
    btnAdd.TabIndex = 5
    lstItems.TabIndex = 6
    btnEdit.TabIndex = 7
    btnDelete.TabIndex = 8
    cboTax.TabIndex = 9
    btnPreview.TabIndex = 10
    btnSave.TabIndex = 11
    btnPdf.TabIndex = 12
    btnNew.TabIndex = 13
    cboHistory.TabIndex = 14
    btnLoad.TabIndex = 15
End Sub

' Busca un control Image en el formulario (lo agrega el build) y lo acomoda.
' Regresa True si hay logo.
Private Function PlaceHeaderLogo() As Boolean
    Dim ctl As Object, img As Object, filePath As String

    For Each ctl In Me.Controls
        If TypeName(ctl) = "Image" Then
            Set img = ctl
            Exit For
        End If
    Next ctl

    ' Respaldo: archivo JPG junto al libro
    If img Is Nothing Then
        If Len(ThisWorkbook.Path) = 0 Then Exit Function
        filePath = ThisWorkbook.Path & Application.PathSeparator & FORM_LOGO_FILE
        If Len(Dir$(filePath)) = 0 Then Exit Function
        On Error GoTo NoLogo
        Set img = Me.Controls.Add("Forms.Image.1", "c_Logo", True)
        img.Picture = LoadPicture(filePath)
        On Error GoTo 0
    End If

    img.AutoSize = False                          ' Si no, vuelve al tamaño de la imagen
    Place img, MARGIN, 7, 112, 34
    img.PictureSizeMode = fmPictureSizeModeZoom   ' Encoge sin deformar
    img.PictureAlignment = fmPictureAlignmentTopLeft
    img.BorderStyle = fmBorderStyleNone
    img.BackStyle = fmBackStyleTransparent
    img.ZOrder fmZOrderFront
    PlaceHeaderLogo = True
    Exit Function
NoLogo:
    PlaceHeaderLogo = False
End Function

'--- Fábricas de controles -----------------------------------------------------

Private Sub Place(ByVal ctl As Object, ByVal x As Single, ByVal y As Single, _
                  ByVal w As Single, ByVal h As Single)
    ctl.Left = x
    ctl.Top = y
    ctl.Width = w
    ctl.Height = h
End Sub

' Estilo plano para cajas de texto, listas y combos: borde gris fino.
Private Sub StyleInput(ByVal ctl As Object)
    ctl.SpecialEffect = fmSpecialEffectFlat       ' Primero quitar el relieve 3D...
    ctl.BorderStyle = fmBorderStyleSingle         ' ...luego el borde sencillo
    ctl.BorderColor = Neutral300()
    ctl.BackColor = vbWhite
    ctl.ForeColor = BrandInk()
    ctl.Font.Name = mFont
    ctl.Font.Size = 10
End Sub

Private Function AddTextBox(ByVal ctlName As String, ByVal x As Single, ByVal y As Single, _
                            ByVal w As Single, ByVal h As Single) As MSForms.TextBox
    Dim ctl As MSForms.TextBox
    Set ctl = Me.Controls.Add("Forms.TextBox.1", ctlName, True)
    Place ctl, x, y, w, h
    StyleInput ctl
    Set AddTextBox = ctl
End Function

Private Function AddComboBox(ByVal ctlName As String, ByVal x As Single, ByVal y As Single, _
                             ByVal w As Single, ByVal h As Single) As MSForms.ComboBox
    Dim ctl As MSForms.ComboBox
    Set ctl = Me.Controls.Add("Forms.ComboBox.1", ctlName, True)
    Place ctl, x, y, w, h
    StyleInput ctl
    Set AddComboBox = ctl
End Function

Private Function AddButton(ByVal ctlName As String, ByVal buttonText As String, ByVal kind As ButtonKind, _
                           ByVal x As Single, ByVal y As Single, _
                           ByVal w As Single, ByVal h As Single) As MSForms.CommandButton
    Dim ctl As MSForms.CommandButton
    Set ctl = Me.Controls.Add("Forms.CommandButton.1", ctlName, True)
    Place ctl, x, y, w, h
    ctl.Caption = buttonText
    ctl.Font.Name = mFont
    ctl.Font.Size = 9
    ctl.Font.Bold = True
    Select Case kind
        Case kindPrimary
            ctl.BackColor = BrandOrange()
            ctl.ForeColor = vbWhite
        Case kindDark
            ctl.BackColor = BrandInk()
            ctl.ForeColor = vbWhite
        Case Else
            ctl.BackColor = vbWhite
            ctl.ForeColor = BrandInk()
    End Select
    Set AddButton = ctl
End Function

' Etiqueta de texto.
Private Function AddLabel(ByVal ctlName As String, ByVal labelText As String, _
                          ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, _
                          ByVal fontSize As Single, ByVal isBold As Boolean, ByVal textColor As Long, _
                          Optional ByVal align As fmTextAlign = fmTextAlignLeft) As MSForms.Label
    Dim ctl As MSForms.Label
    Set ctl = Me.Controls.Add("Forms.Label.1", ctlName, True)
    Place ctl, x, y, w, h
    With ctl
        .Caption = labelText
        .BackStyle = fmBackStyleTransparent
        .WordWrap = False
        .TextAlign = align
        .Font.Name = mFont
        .Font.Size = fontSize
        .Font.Bold = isBold
        .ForeColor = textColor
    End With
    Set AddLabel = ctl
End Function

' Rectángulo de color (bandas y líneas).
'   toBack: True lo manda detrás de los demás controles.
Private Sub AddBand(ByVal ctlName As String, ByVal x As Single, ByVal y As Single, _
                    ByVal w As Single, ByVal h As Single, ByVal fillColor As Long, ByVal toBack As Boolean)
    Dim ctl As MSForms.Label
    Set ctl = Me.Controls.Add("Forms.Label.1", ctlName, True)
    Place ctl, x, y, w, h
    ctl.Caption = ""
    ctl.BackStyle = fmBackStyleOpaque
    ctl.BackColor = fillColor
    If toBack Then ctl.ZOrder fmZOrderBack
End Sub


'==============================================================================
' ESTADO DEL FORMULARIO
'==============================================================================

' Deja el formulario listo para una cotización nueva.
Private Sub StartNewQuote()
    mSuppressEvents = True

    Set mQuote = New Quote
    mQuote.Number = GetNextQuoteNumber()
    mIsSaved = False
    mOverwriteConfirmed = False

    txtNumber.Text = mQuote.Number
    txtDate.Text = Format$(mQuote.QuoteDate, "dd/mm/yyyy")
    cboClient.Text = ""
    cboTax.ListIndex = DEFAULT_TAX_OPTION
    cboHistory.ListIndex = -1
    cboHistory.Text = ""

    ExitEditMode
    ClearItemFields
    RefreshItems
    UpdateValidUntil

    mSuppressEvents = False
    mIsDirty = False
    FocusOn cboClient
End Sub

Private Sub FocusOn(ByVal ctl As Object)
    On Error Resume Next                          ' Falla si el formulario aún no es visible
    ctl.SetFocus
    On Error GoTo 0
End Sub

' Muestra un aviso y lleva el cursor al control con problema.
Private Sub Warn(ByVal message As String, ByVal ctl As Object)
    MsgBox message, vbExclamation, TTL_REVIEW
    FocusOn ctl
    On Error Resume Next
    ctl.SelStart = 0
    ctl.SelLength = Len(ctl.Text)
    On Error GoTo 0
End Sub

' Cambio hecho por el usuario (no por el código).
Private Sub MarkDirty()
    If Not mSuppressEvents Then mIsDirty = True
End Sub

Private Sub cboClient_Change()
    MarkDirty
End Sub

Private Sub txtDate_Change()
    UpdateValidUntil                              ' Siempre, aunque lo cambie el código
    MarkDirty
End Sub

Private Sub UpdateValidUntil()
    If IsDate(txtDate.Text) Then
        lblValidUntil.Caption = Fmt(UI_VALID_UNTIL, Format$(CDate(txtDate.Text) + ValidityDays(), "dd/mm/yyyy"))
    Else
        lblValidUntil.Caption = ""
    End If
End Sub

' Cambiar la tasa recalcula los totales al instante.
Private Sub cboTax_Change()
    If mQuote Is Nothing Then Exit Sub
    mQuote.TaxRate = TaxRateOf(cboTax.ListIndex)
    RefreshItems
    MarkDirty
End Sub


'==============================================================================
' CAPTURA DE CONCEPTOS
'==============================================================================

Private Sub txtQty_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    FilterNumericKey txtQty.Text, KeyAscii
End Sub

Private Sub txtPrice_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    FilterNumericKey txtPrice.Text, KeyAscii
End Sub

' Solo deja pasar dígitos, retroceso y UN separador decimal.
Private Sub FilterNumericKey(ByVal currentText As String, ByVal KeyAscii As MSForms.ReturnInteger)
    Dim key As String
    If KeyAscii = 8 Then Exit Sub                 ' Retroceso
    key = Chr$(KeyAscii)
    If key Like "[0-9]" Then Exit Sub
    If key = DecimalSeparator() And InStr(currentText, DecimalSeparator()) = 0 Then Exit Sub
    KeyAscii = 0                                  ' Cualquier otra tecla se ignora
End Sub

Private Sub txtQty_Change()
    UpdateLineAmount
End Sub

Private Sub txtPrice_Change()
    UpdateLineAmount
End Sub

' Enter en el precio = Agregar.
Private Sub txtPrice_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        btnAdd_Click
    End If
End Sub

Private Sub UpdateLineAmount()
    Dim qty As Double, price As Double
    If TryParseNumber(txtQty.Text, qty) And TryParseNumber(txtPrice.Text, price) Then
        lblAmount.Caption = FormatMoney(Round2(qty * price)) & " "
    Else
        lblAmount.Caption = FormatMoney(0) & " "
    End If
End Sub

Private Sub ClearItemFields()
    txtQty.Text = ""
    txtDescription.Text = ""
    txtPrice.Text = ""
    lblAmount.Caption = FormatMoney(0) & " "
End Sub

' Valida los campos del concepto. Si todo está bien, regresa un QuoteItem; si no, Nothing.
Private Function ReadItemFields() As QuoteItem
    Dim qty As Double, price As Double, descText As String

    If Not TryParseNumber(txtQty.Text, qty) Then
        Warn MSG_QTY_NOT_NUMBER, txtQty: Exit Function
    End If
    If qty <= 0 Then
        Warn MSG_QTY_NOT_POSITIVE, txtQty: Exit Function
    End If

    descText = Trim$(txtDescription.Text)
    If Len(descText) = 0 Then
        Warn MSG_DESCRIPTION_EMPTY, txtDescription: Exit Function
    End If

    If Not TryParseNumber(txtPrice.Text, price) Then
        Warn MSG_PRICE_NOT_NUMBER, txtPrice: Exit Function
    End If
    If price < 0 Then
        Warn MSG_PRICE_NEGATIVE, txtPrice: Exit Function
    End If

    Set ReadItemFields = NewQuoteItem(qty, descText, price)
End Function

Private Sub btnAdd_Click()
    Dim newItem As QuoteItem
    On Error GoTo Fail

    Set newItem = ReadItemFields()
    If newItem Is Nothing Then Exit Sub

    If mEditIndex > 0 Then
        mQuote.ReplaceItem mEditIndex, newItem
        ExitEditMode
    Else
        If mQuote.ItemCount >= MAX_ITEMS Then
            MsgBox Fmt(MSG_MAX_ITEMS, MAX_ITEMS), vbExclamation, TTL_REVIEW
            Exit Sub
        End If
        mQuote.AddItem newItem
    End If

    MarkDirty
    RefreshItems
    ClearItemFields
    FocusOn txtQty
    Exit Sub
Fail:
    ShowError "btnAdd_Click"
End Sub


'==============================================================================
' EDITAR / ELIMINAR CONCEPTOS
'==============================================================================

' Primer clic: carga el concepto seleccionado para editarlo.
' Segundo clic (dice "Cancelar"): descarta la edición.
Private Sub btnEdit_Click()
    Dim currentItem As QuoteItem

    If mEditIndex > 0 Then
        ExitEditMode
        ClearItemFields
        Exit Sub
    End If

    If lstItems.ListIndex < 0 Then
        MsgBox MSG_SELECT_ITEM_EDIT, vbInformation, TTL_REVIEW
        Exit Sub
    End If

    mEditIndex = lstItems.ListIndex + 1           ' ListBox empieza en 0, la colección en 1
    Set currentItem = mQuote.Item(mEditIndex)
    txtQty.Text = CStr(currentItem.Quantity)
    txtDescription.Text = currentItem.Description
    txtPrice.Text = CStr(currentItem.UnitPrice)

    btnAdd.Caption = UI_BTN_UPDATE
    btnAdd.BackColor = BrandOrange()              ' Resalta que estás editando
    btnEdit.Caption = UI_BTN_CANCEL
    btnDelete.Enabled = False
    lstItems.Enabled = False
    FocusOn txtQty
End Sub

Private Sub lstItems_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If mEditIndex = 0 Then btnEdit_Click
End Sub

Private Sub lstItems_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyDelete And mEditIndex = 0 Then btnDelete_Click
End Sub

Private Sub ExitEditMode()
    mEditIndex = 0
    btnAdd.Caption = UI_BTN_ADD
    btnAdd.BackColor = BrandInk()
    btnEdit.Caption = UI_BTN_EDIT
    btnDelete.Enabled = True
    lstItems.Enabled = True
End Sub

Private Sub btnDelete_Click()
    Dim index As Long
    On Error GoTo Fail

    index = lstItems.ListIndex
    If index < 0 Then
        MsgBox MSG_SELECT_ITEM_DELETE, vbInformation, TTL_REVIEW
        Exit Sub
    End If

    If MsgBox(Fmt(MSG_CONFIRM_DELETE, mQuote.Item(index + 1).Description), _
              vbYesNo + vbQuestion, TTL_DELETE) = vbNo Then Exit Sub

    mQuote.RemoveItem index + 1
    MarkDirty
    RefreshItems

    ' Deja seleccionado el siguiente para borrar varios seguidos
    If lstItems.ListCount > 0 Then
        If index > lstItems.ListCount - 1 Then index = lstItems.ListCount - 1
        lstItems.ListIndex = index
    End If
    Exit Sub
Fail:
    ShowError "btnDelete_Click"
End Sub

' Redibuja la lista y los totales a partir de mQuote.
Private Sub RefreshItems()
    Dim i As Long, r As Long

    If mQuote Is Nothing Then Exit Sub
    lstItems.Clear
    For i = 1 To mQuote.ItemCount
        With mQuote.Item(i)
            lstItems.AddItem FormatQuantity(.Quantity)
            r = lstItems.ListCount - 1
            lstItems.List(r, 1) = .Description
            lstItems.List(r, 2) = FormatMoney(.UnitPrice)
            lstItems.List(r, 3) = FormatMoney(.Amount)
        End With
    Next i

    lblSubtotal.Caption = FormatMoney(mQuote.Subtotal)
    lblTax.Caption = FormatMoney(mQuote.Tax)
    lblTotal.Caption = FormatMoney(mQuote.Total)
End Sub


'==============================================================================
' VALIDACIÓN Y SINCRONIZACIÓN
'==============================================================================

' Pasa los datos del encabezado del formulario a mQuote, validándolos.
' Regresa False (y avisa) si algo falta o es inválido.
Private Function SyncQuoteFromForm() As Boolean
    Dim clientText As String

    If mEditIndex > 0 Then
        MsgBox MSG_FINISH_EDIT, vbExclamation, TTL_REVIEW
        Exit Function
    End If

    clientText = Trim$(cboClient.Text)
    If Len(clientText) = 0 Then
        Warn MSG_CLIENT_EMPTY, cboClient: Exit Function
    End If
    If Not IsDate(txtDate.Text) Then
        Warn MSG_DATE_INVALID, txtDate: Exit Function
    End If
    If mQuote.ItemCount = 0 Then
        Warn MSG_NO_ITEMS, txtQty: Exit Function
    End If

    mQuote.Number = txtNumber.Text
    mQuote.ClientName = clientText
    mQuote.QuoteDate = CDate(txtDate.Text)
    mQuote.TaxRate = TaxRateOf(cboTax.ListIndex)
    SyncQuoteFromForm = True
End Function


'==============================================================================
' ACCIONES PRINCIPALES
'==============================================================================

Private Sub btnPreview_Click()
    On Error GoTo Fail
    If Not SyncQuoteFromForm() Then Exit Sub
    FillTemplate mQuote
    ThisWorkbook.Worksheets(SHEET_TEMPLATE).Activate
    Exit Sub
Fail:
    ShowError "btnPreview_Click"
End Sub

Private Sub btnSave_Click()
    On Error GoTo Fail
    If SaveCurrentQuote() Then
        MsgBox Fmt(MSG_SAVED, mQuote.Number), vbInformation
    End If
    Exit Sub
Fail:
    ShowError "btnSave_Click"
End Sub

' Guarda la cotización en pantalla. Regresa True si se guardó.
' Solo pregunta "¿sobrescribir?" la primera vez para una cotización cargada del historial.
Private Function SaveCurrentQuote() As Boolean
    If Not SyncQuoteFromForm() Then Exit Function

    If QuoteExists(mQuote.Number) Then
        If mIsSaved Then
            If Not mOverwriteConfirmed Then
                If MsgBox(Fmt(MSG_CONFIRM_OVERWRITE, mQuote.Number), _
                          vbYesNo + vbQuestion, TTL_OVERWRITE) = vbNo Then Exit Function
                mOverwriteConfirmed = True
            End If
        Else
            ' Cotización nueva cuyo folio ya ocupó otra: se asigna el siguiente libre
            mQuote.Number = GetNextQuoteNumber()
            txtNumber.Text = mQuote.Number
            MsgBox Fmt(MSG_NUMBER_REASSIGNED, mQuote.Number), vbInformation
        End If
    End If

    If SaveQuote(mQuote) Then
        FillTemplate mQuote
        LoadHistoryList
        LoadClientList                            ' Por si fue un cliente nuevo
        mIsSaved = True
        mOverwriteConfirmed = True                ' Ya es "nuestra": no vuelve a preguntar
        mIsDirty = False                          ' Al final, para que nada lo vuelva a marcar
        SaveCurrentQuote = True
    End If
End Function

Private Sub btnPdf_Click()
    Dim filePath As String, answer As VbMsgBoxResult
    On Error GoTo Fail

    If Not SyncQuoteFromForm() Then Exit Sub

    If mIsDirty Or Not mIsSaved Then
        answer = MsgBox(Fmt(MSG_SAVE_BEFORE_PDF), vbYesNoCancel + vbQuestion, TTL_PDF)
        If answer = vbCancel Then Exit Sub
        If answer = vbYes Then
            If Not SaveCurrentQuote() Then Exit Sub
        End If
    End If

    If Not ConfirmCompanyData() Then Exit Sub
    FillTemplate mQuote
    filePath = ExportPdf(mQuote)
    If Len(filePath) > 0 Then MsgBox Fmt(MSG_PDF_CREATED, filePath), vbInformation
    Exit Sub
Fail:
    ShowError "btnPdf_Click"
End Sub

' Antes del PDF: avisa si falta el nombre de la empresa o el RFC es inválido.
' Regresa True si se puede continuar.
Private Function ConfirmCompanyData() As Boolean
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

Private Sub btnNew_Click()
    On Error GoTo Fail
    If mIsDirty Then
        If MsgBox(Fmt(MSG_UNSAVED_NEW), vbYesNo + vbExclamation, TTL_NEW) = vbNo Then Exit Sub
    End If
    StartNewQuote
    ClearTemplate
    Exit Sub
Fail:
    ShowError "btnNew_Click"
End Sub

Private Sub btnLogo_Click()
    On Error GoTo Fail
    InsertCompanyLogo
    Exit Sub
Fail:
    ShowError "btnLogo_Click"
End Sub

Private Sub btnCompany_Click()
    On Error GoTo Fail
    OpenCompanySheet
    Exit Sub
Fail:
    ShowError "btnCompany_Click"
End Sub

Private Sub lblInnobytes_Click()
    If Len(INNOBYTES_WEB) = 0 Then Exit Sub
    On Error Resume Next                          ' Sin navegador o sin internet: no truena
    ThisWorkbook.FollowHyperlink "https://" & INNOBYTES_WEB
    On Error GoTo 0
End Sub


'==============================================================================
' HISTORIAL Y CLIENTES
'==============================================================================

Private Sub LoadHistoryList()
    Dim data As Variant
    cboHistory.Clear
    data = ListQuotes()
    If Not IsEmpty(data) Then cboHistory.List = data
End Sub

' Recarga el catálogo de clientes conservando lo escrito (no cuenta como cambio).
Private Sub LoadClientList()
    Dim data As Variant, typed As String, wasSuppressed As Boolean

    wasSuppressed = mSuppressEvents
    mSuppressEvents = True

    typed = cboClient.Text
    cboClient.Clear
    data = ListClients()
    If Not IsEmpty(data) Then cboClient.List = data
    cboClient.Text = typed

    mSuppressEvents = wasSuppressed
End Sub

Private Sub btnLoad_Click()
    Dim quoteNumber As String, loaded As Quote
    On Error GoTo Fail

    quoteNumber = UCase$(Trim$(cboHistory.Text))
    If Len(quoteNumber) = 0 Then
        MsgBox MSG_HISTORY_EMPTY, vbInformation, TTL_REVIEW
        Exit Sub
    End If

    Set loaded = LoadQuote(quoteNumber)
    If loaded Is Nothing Then
        MsgBox Fmt(MSG_QUOTE_NOT_FOUND, quoteNumber), vbExclamation, TTL_REVIEW
        Exit Sub
    End If

    If mIsDirty Then
        If MsgBox(Fmt(MSG_UNSAVED_LOAD, quoteNumber), vbYesNo + vbExclamation, TTL_REVIEW) = vbNo Then Exit Sub
    End If

    mSuppressEvents = True                        ' Llenar campos desde la BD no es un cambio
    Set mQuote = loaded
    ExitEditMode
    ClearItemFields
    txtNumber.Text = mQuote.Number
    txtDate.Text = Format$(mQuote.QuoteDate, "dd/mm/yyyy")
    cboClient.Text = mQuote.ClientName
    cboTax.ListIndex = TaxOptionOf(mQuote.TaxRate)
    RefreshItems
    FillTemplate mQuote
    mSuppressEvents = False

    mIsSaved = True
    mOverwriteConfirmed = False                   ' Preguntará una vez antes de sobrescribir
    mIsDirty = False
    Exit Sub
Fail:
    mSuppressEvents = False
    ShowError "btnLoad_Click"
End Sub

' Enter en el buscador = Cargar.
Private Sub cboHistory_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        btnLoad_Click
    End If
End Sub
