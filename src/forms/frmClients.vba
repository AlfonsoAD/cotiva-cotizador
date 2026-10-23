'==============================================================================
' frmClients · Catálogo de clientes: búsqueda, alta, edición y datos fiscales
'
' Igual que frmQuote y frmOrder: todo el diseño se construye por código en
' BuildLayout y este archivo contiene SOLO el código del formulario (build.ps1
' crea el formulario y le inserta este código).
'
' Siempre se abre modal desde modApp.PickClient. Los botones "Usar este cliente"
' y "Cerrar" y la X solo ocultan el formulario: quien lo abrió lee ChosenClient
' y después lo descarga. Todos los datos fiscales son opcionales y los avisos
' no impiden guardar. El nombre es la llave del cliente: no se renombra y no
' hay forma de eliminarlo.
'==============================================================================
Option Explicit

'--- Estado ---
Private mCurrent As String              ' Nombre del cliente guardado en pantalla; "" = alta nueva
Private mIsDirty As Boolean             ' True si hay cambios sin guardar
Private mSuppressEvents As Boolean      ' True mientras el CÓDIGO llena campos (no cuenta como cambio)
Private mPickMode As Boolean            ' True si lo abrió una cotización u orden para elegir un cliente
Private mChosen As String               ' Cliente elegido con "Usar este cliente"
Private mFont As String                 ' Fuente de marca
Private mPrimary As Long                ' Color principal de la empresa
Private mSecondary As Long              ' Color secundario de la empresa

'--- Controles sin eventos ---
Private mpgClient As MSForms.MultiPage
Private lblRfcHint As MSForms.Label
Private lblZipHint As MSForms.Label

'--- Controles con eventos ---
Private WithEvents txtSearch As MSForms.TextBox
Private WithEvents lstClients As MSForms.ListBox
Private WithEvents txtName As MSForms.TextBox
Private WithEvents txtPhone1 As MSForms.TextBox
Private WithEvents txtPhone2 As MSForms.TextBox
Private WithEvents txtEmail As MSForms.TextBox
Private WithEvents txtAddress As MSForms.TextBox
Private WithEvents txtRfc As MSForms.TextBox
Private WithEvents txtLegalName As MSForms.TextBox
Private WithEvents cboTaxRegime As MSForms.ComboBox
Private WithEvents txtTaxZip As MSForms.TextBox
Private WithEvents cboCfdiUse As MSForms.ComboBox
Private WithEvents txtStreet As MSForms.TextBox
Private WithEvents txtExtNumber As MSForms.TextBox
Private WithEvents txtIntNumber As MSForms.TextBox
Private WithEvents txtNeighborhood As MSForms.TextBox
Private WithEvents txtMunicipality As MSForms.TextBox
Private WithEvents txtState As MSForms.TextBox
Private WithEvents btnNew As MSForms.CommandButton
Private WithEvents btnSave As MSForms.CommandButton
Private WithEvents btnUse As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton

'--- Medidas del diseño (puntos) ---
Private Const FORM_WIDTH As Single = 760
Private Const FORM_HEIGHT As Single = 496
Private Const LIST_WIDTH As Single = 230
Private Const DETAIL_X As Single = 260
Private Const PAGE_CONTACT As Long = 0
Private Const PAGE_FISCAL As Long = 1


'==============================================================================
' ARRANQUE Y CIERRE
'==============================================================================

Private Sub UserForm_Initialize()
    BuildLayout
    LoadClientList vbNullString
    LoadCatalogCombos
    StartNewClient vbNullString
End Sub

' Prepara el formulario antes de mostrarlo. Si currentClient existe, lo muestra;
' si no, queda en alta nueva con ese nombre escrito.
Public Sub Start(ByVal currentClient As String, ByVal pickMode As Boolean)
    Dim c As Client

    mPickMode = pickMode
    mChosen = vbNullString
    btnUse.Visible = pickMode

    mSuppressEvents = True
    txtSearch.Text = vbNullString
    mSuppressEvents = False
    LoadClientList vbNullString

    currentClient = Trim$(currentClient)
    If Len(currentClient) > 0 Then Set c = LoadClient(currentClient)
    If c Is Nothing Then
        StartNewClient currentClient
    Else
        ShowClient c
    End If
End Sub

' Cliente elegido con "Usar este cliente"; vacio si se cerro sin elegir. Se lee antes de descargar el formulario.
Public Property Get ChosenClient() As String
    ChosenClient = mChosen
End Property

' La X de la ventana oculta el formulario (quien lo abrió lo descarga).
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        CloseWithoutChoosing
    End If
End Sub

Private Sub CloseWithoutChoosing()
    If Not ConfirmDiscard() Then Exit Sub
    mChosen = vbNullString
    mIsDirty = False
    Me.Hide
End Sub


'==============================================================================
' DISEÑO (se construye por código)
'==============================================================================

Private Sub BuildLayout()
    Dim rightEdge As Single, hasLogo As Boolean

    mFont = BrandFont()
    rightEdge = FORM_WIDTH - FORM_MARGIN

    ' --- Ventana (la diferencia Width-InsideWidth son los bordes de Windows) ---
    Me.Caption = UI_CLI_CAPTION
    Me.BackColor = vbWhite
    Me.Width = FORM_WIDTH + (Me.Width - Me.InsideWidth)
    Me.Height = FORM_HEIGHT + (Me.Height - Me.InsideHeight)

    ' --- Barra superior: clara si hay logo (JPG con fondo blanco), tinta si no ---
    mPrimary = PrimaryColor()
    mSecondary = SecondaryColor()
    BeginFormKit mFont, mPrimary, mSecondary
    hasLogo = PlaceHeaderLogo(Me)
    If hasLogo Then
        AddBand Me, "c_Bar", 0, 0, FORM_WIDTH, 48, vbWhite, True
        AddBand Me, "c_BarLine", 0, 48, FORM_WIDTH, 2, mPrimary, False
        AddLabel Me, "c_Module", UI_CLI_MODULE, 142, 17, 160, 18, 11, True, TextGray()
    Else
        AddBand Me, "c_Bar", 0, 0, FORM_WIDTH, 48, mSecondary, True
        AddLabel Me, "c_Brand", UI_BRAND, FORM_MARGIN, 12, 80, 24, 16, True, ContrastTextColor(mSecondary)
        AddLabel Me, "c_Module", UI_CLI_MODULE, 104, 17, 160, 18, 11, False, BrandInkLight()
    End If

    ' --- Izquierda: buscador y lista ---
    AddLabel Me, "c_LSearch", UI_CLI_LBL_SEARCH, FORM_MARGIN, 58, LIST_WIDTH, 12, 8, False, TextGray()
    Set txtSearch = AddTextBox(Me, "c_Search", FORM_MARGIN, 72, LIST_WIDTH, 22)
    txtSearch.MaxLength = 100
    txtSearch.ControlTipText = TIP_CLI_SEARCH
    Set lstClients = Me.Controls.Add("Forms.ListBox.1", "c_List", True)
    Place lstClients, FORM_MARGIN, 100, LIST_WIDTH, 340
    StyleInput lstClients

    ' --- Derecha: nombre y pestañas ---
    AddLabel Me, "c_LName", UI_CLI_LBL_NAME, DETAIL_X, 58, 300, 12, 8, False, TextGray()
    Set txtName = AddTextBox(Me, "c_Name", DETAIL_X, 72, rightEdge - DETAIL_X, 22)
    txtName.MaxLength = 100
    txtName.ControlTipText = TIP_CLI_NAME

    Set mpgClient = Me.Controls.Add("Forms.MultiPage.1", "c_Pages", True)
    Place mpgClient, DETAIL_X, 102, rightEdge - DETAIL_X, 338
    mpgClient.Font.Name = mFont
    mpgClient.Font.Size = 9
    mpgClient.Pages(PAGE_CONTACT).Caption = UI_CLI_TAB_CONTACT
    mpgClient.Pages(PAGE_FISCAL).Caption = UI_CLI_TAB_FISCAL
    BuildContactPage mpgClient.Pages(PAGE_CONTACT)
    BuildFiscalPage mpgClient.Pages(PAGE_FISCAL)

    ' --- Barra inferior de acciones ---
    AddBand Me, "c_Footer", 0, 450, FORM_WIDTH, FORM_HEIGHT - 450, Neutral50(), True
    AddBand Me, "c_FooterLine", 0, 450, FORM_WIDTH, 1, Neutral200(), False
    Set btnNew = AddButton(Me, "c_New", UI_CLI_BTN_NEW, kindSecondary, FORM_MARGIN, 460, 100, 26)
    Set btnSave = AddButton(Me, "c_Save", UI_CLI_BTN_SAVE, kindPrimary, rightEdge - 330, 460, 100, 26)
    Set btnUse = AddButton(Me, "c_Use", UI_CLI_BTN_USE, kindDark, rightEdge - 224, 460, 130, 26)
    Set btnClose = AddButton(Me, "c_Close", UI_CLI_BTN_CLOSE, kindSecondary, rightEdge - 88, 460, 88, 26)
    btnUse.Visible = False
    btnUse.ControlTipText = TIP_CLI_USE

    ' --- Orden del tabulador ---
    txtSearch.TabIndex = 0
    lstClients.TabIndex = 1
    txtName.TabIndex = 2
    mpgClient.TabIndex = 3
    btnSave.TabIndex = 4
    btnUse.TabIndex = 5
    btnNew.TabIndex = 6
    btnClose.TabIndex = 7
End Sub

' Pestaña "Contacto".
Private Sub BuildContactPage(ByVal pg As Object)
    AddFieldLabel pg, "c_LPhone1", UI_CLI_LBL_PHONE1, 12, 8, 220
    AddFieldLabel pg, "c_LPhone2", UI_CLI_LBL_PHONE2, 244, 8, 220
    Set txtPhone1 = AddTextBox(pg, "c_Phone1", 12, 22, 220, 22)
    txtPhone1.MaxLength = 20
    Set txtPhone2 = AddTextBox(pg, "c_Phone2", 244, 22, 220, 22)
    txtPhone2.MaxLength = 20

    AddFieldLabel pg, "c_LEmail", UI_CLI_LBL_EMAIL, 12, 52, 452
    Set txtEmail = AddTextBox(pg, "c_Email", 12, 66, 452, 22)
    txtEmail.MaxLength = 100

    AddFieldLabel pg, "c_LAddress", UI_CLI_LBL_ADDRESS, 12, 96, 452
    Set txtAddress = AddTextBox(pg, "c_Address", 12, 110, 452, 22)
    txtAddress.MaxLength = 200

    txtPhone1.TabIndex = 0
    txtPhone2.TabIndex = 1
    txtEmail.TabIndex = 2
    txtAddress.TabIndex = 3
End Sub

' Pestaña "Datos fiscales" (los de la Constancia de Situación Fiscal).
Private Sub BuildFiscalPage(ByVal pg As Object)
    AddFieldLabel pg, "c_LRfc", UI_CLI_LBL_RFC, 12, 8, 150
    AddFieldLabel pg, "c_LLegalName", UI_CLI_LBL_LEGAL_NAME, 170, 8, 294
    Set txtRfc = AddTextBox(pg, "c_Rfc", 12, 22, 150, 22)
    txtRfc.MaxLength = 13
    Set txtLegalName = AddTextBox(pg, "c_LegalName", 170, 22, 294, 22)
    txtLegalName.MaxLength = 200
    Set lblRfcHint = AddLabel(pg, "c_RfcHint", "", 12, 46, 452, 12, 8, False, TextGray())

    AddFieldLabel pg, "c_LTaxRegime", UI_CLI_LBL_TAX_REGIME, 12, 64, 340
    AddFieldLabel pg, "c_LTaxZip", UI_CLI_LBL_TAX_ZIP, 360, 64, 104
    Set cboTaxRegime = AddCatalogCombo(pg, "c_TaxRegime", 12, 78, 340)
    Set txtTaxZip = AddTextBox(pg, "c_TaxZip", 360, 78, 104, 22)
    txtTaxZip.MaxLength = 5
    Set lblZipHint = AddLabel(pg, "c_ZipHint", "", 12, 102, 452, 12, 8, False, TextGray())

    AddFieldLabel pg, "c_LCfdiUse", UI_CLI_LBL_CFDI_USE, 12, 120, 452
    Set cboCfdiUse = AddCatalogCombo(pg, "c_CfdiUse", 12, 134, 452)

    AddLabel pg, "c_SecAddress", UI_CLI_SEC_FISCAL_ADDRESS, 12, 168, 300, 12, 8, True, mPrimary
    AddFieldLabel pg, "c_LStreet", UI_CLI_LBL_STREET, 12, 184, 290
    AddFieldLabel pg, "c_LExt", UI_CLI_LBL_EXT_NUMBER, 310, 184, 70
    AddFieldLabel pg, "c_LInt", UI_CLI_LBL_INT_NUMBER, 388, 184, 76
    Set txtStreet = AddTextBox(pg, "c_Street", 12, 198, 290, 22)
    txtStreet.MaxLength = 100
    Set txtExtNumber = AddTextBox(pg, "c_Ext", 310, 198, 70, 22)
    txtExtNumber.MaxLength = 20
    Set txtIntNumber = AddTextBox(pg, "c_Int", 388, 198, 76, 22)
    txtIntNumber.MaxLength = 20

    AddFieldLabel pg, "c_LNeighborhood", UI_CLI_LBL_NEIGHBORHOOD, 12, 228, 220
    AddFieldLabel pg, "c_LMunicipality", UI_CLI_LBL_MUNICIPALITY, 244, 228, 220
    Set txtNeighborhood = AddTextBox(pg, "c_Neighborhood", 12, 242, 220, 22)
    txtNeighborhood.MaxLength = 100
    Set txtMunicipality = AddTextBox(pg, "c_Municipality", 244, 242, 220, 22)
    txtMunicipality.MaxLength = 100

    AddFieldLabel pg, "c_LState", UI_CLI_LBL_STATE, 12, 272, 220
    Set txtState = AddTextBox(pg, "c_State", 12, 286, 220, 22)
    txtState.MaxLength = 50

    txtRfc.TabIndex = 0
    txtLegalName.TabIndex = 1
    cboTaxRegime.TabIndex = 2
    txtTaxZip.TabIndex = 3
    cboCfdiUse.TabIndex = 4
    txtStreet.TabIndex = 5
    txtExtNumber.TabIndex = 6
    txtIntNumber.TabIndex = 7
    txtNeighborhood.TabIndex = 8
    txtMunicipality.TabIndex = 9
    txtState.TabIndex = 10
End Sub

' Combo de catálogo del SAT: columna 0 = clave (oculta), columna 1 = "clave - descripción".
Private Function AddCatalogCombo(ByVal pg As Object, ByVal ctlName As String, ByVal x As Single, _
        ByVal y As Single, ByVal w As Single) As MSForms.ComboBox
    Dim cbo As MSForms.ComboBox

    Set cbo = AddComboBox(pg, ctlName, x, y, w, 22)
    With cbo
        .Style = fmStyleDropDownList              ' Solo se elige de la lista (o se deja vacío)
        .ColumnCount = 2
        .ColumnWidths = "0 pt;" & CStr(w - 20) & " pt"
        .BoundColumn = 1
        .TextColumn = 2
        .ListRows = 12
        .ListWidth = w
    End With
    Set AddCatalogCombo = cbo
End Function

' Etiqueta pequeña sobre un campo.
Private Sub AddFieldLabel(ByVal pg As Object, ByVal ctlName As String, ByVal labelText As String, _
                          ByVal x As Single, ByVal y As Single, ByVal w As Single)
    AddLabel pg, ctlName, labelText, x, y, w, 12, 8, False, TextGray()
End Sub

' Llena los combos de régimen fiscal y uso de CFDI con los catálogos de BD_*.
Private Sub LoadCatalogCombos()
    FillCatalogCombo cboTaxRegime, ListTaxRegimes()
    FillCatalogCombo cboCfdiUse, ListCfdiUses()
End Sub

' La primera fila va vacía: es "sin dato".
Private Sub FillCatalogCombo(ByVal cbo As MSForms.ComboBox, ByVal catalog As Variant)
    Dim i As Long, rows As Long

    cbo.Clear
    cbo.AddItem vbNullString
    If IsEmpty(catalog) Then Exit Sub
    rows = UBound(catalog, 1) - LBound(catalog, 1) + 1
    For i = 0 To rows - 1
        cbo.AddItem CStr(catalog(LBound(catalog, 1) + i, 0))
        cbo.List(i + 1, 1) = CStr(catalog(LBound(catalog, 1) + i, 0)) & " - " & _
            CStr(catalog(LBound(catalog, 1) + i, 1))
    Next i
End Sub

' Elige la clave en un combo de catálogo; si ya no está en el catálogo la agrega
' para no perderla.
Private Sub SelectCatalogKey(ByVal cbo As MSForms.ComboBox, ByVal key As String)
    Dim i As Long

    key = Trim$(key)
    If Len(key) = 0 Then
        cbo.ListIndex = 0
        Exit Sub
    End If
    For i = 1 To cbo.ListCount - 1
        If StrComp(CStr(cbo.List(i, 0)), key, vbTextCompare) = 0 Then
            cbo.ListIndex = i
            Exit Sub
        End If
    Next i
    cbo.AddItem key
    cbo.List(cbo.ListCount - 1, 1) = key
    cbo.ListIndex = cbo.ListCount - 1
End Sub

' Clave elegida en un combo de catálogo ("" si no hay).
Private Function SelectedCatalogKey(ByVal cbo As MSForms.ComboBox) As String
    If cbo.ListIndex > 0 Then SelectedCatalogKey = CStr(cbo.List(cbo.ListIndex, 0))
End Function


'==============================================================================
' ESTADO DEL FORMULARIO
'==============================================================================

' Cambio hecho por el usuario (no por el código).
Private Sub MarkDirty()
    If Not mSuppressEvents Then mIsDirty = True
End Sub

' Pide confirmación si hay cambios sin guardar. True = se puede continuar.
Private Function ConfirmDiscard() As Boolean
    If Not mIsDirty Then
        ConfirmDiscard = True
        Exit Function
    End If
    ConfirmDiscard = (MsgBox(MSG_CLI_UNSAVED, vbYesNo + vbExclamation, TTL_CLI_UNSAVED) = vbYes)
End Function

' Deja el formulario en alta nueva, con el nombre escrito si se indica.
Private Sub StartNewClient(ByVal clientName As String)
    mSuppressEvents = True
    ClearFields
    txtName.Text = clientName
    txtName.Locked = False
    txtName.BackColor = vbWhite
    mCurrent = vbNullString
    lstClients.ListIndex = -1
    mpgClient.Value = PAGE_CONTACT
    mSuppressEvents = False
    mIsDirty = False
    FocusOn txtName
End Sub

Private Sub ClearFields()
    txtName.Text = vbNullString
    txtPhone1.Text = vbNullString
    txtPhone2.Text = vbNullString
    txtEmail.Text = vbNullString
    txtAddress.Text = vbNullString
    txtRfc.Text = vbNullString
    txtLegalName.Text = vbNullString
    cboTaxRegime.ListIndex = 0
    txtTaxZip.Text = vbNullString
    cboCfdiUse.ListIndex = 0
    txtStreet.Text = vbNullString
    txtExtNumber.Text = vbNullString
    txtIntNumber.Text = vbNullString
    txtNeighborhood.Text = vbNullString
    txtMunicipality.Text = vbNullString
    txtState.Text = vbNullString
    UpdateHints
End Sub

' Muestra un cliente guardado (no cuenta como cambio). El nombre queda bloqueado.
Private Sub ShowClient(ByVal c As Client)
    mSuppressEvents = True
    txtName.Text = c.Name
    txtName.Locked = True
    txtName.BackColor = Neutral50()
    txtPhone1.Text = c.Phone1
    txtPhone2.Text = c.Phone2
    txtEmail.Text = c.Email
    txtAddress.Text = c.Address
    txtRfc.Text = c.Rfc
    txtLegalName.Text = c.LegalName
    SelectCatalogKey cboTaxRegime, c.TaxRegime
    txtTaxZip.Text = c.TaxZip
    SelectCatalogKey cboCfdiUse, c.CfdiUse
    txtStreet.Text = c.Street
    txtExtNumber.Text = c.ExtNumber
    txtIntNumber.Text = c.IntNumber
    txtNeighborhood.Text = c.Neighborhood
    txtMunicipality.Text = c.Municipality
    txtState.Text = c.State
    mCurrent = c.Name
    SelectInList c.Name
    UpdateHints
    mSuppressEvents = False
    mIsDirty = False
End Sub

' Pasa los campos a un Client.
Private Function ReadClientFromForm() As Client
    Dim c As Client

    Set c = New Client
    c.Name = Trim$(txtName.Text)
    c.Phone1 = Trim$(txtPhone1.Text)
    c.Phone2 = Trim$(txtPhone2.Text)
    c.Email = Trim$(txtEmail.Text)
    c.Address = Trim$(txtAddress.Text)
    c.Rfc = UCase$(Trim$(txtRfc.Text))
    c.LegalName = Trim$(txtLegalName.Text)
    c.TaxRegime = SelectedCatalogKey(cboTaxRegime)
    c.TaxZip = Trim$(txtTaxZip.Text)
    c.CfdiUse = SelectedCatalogKey(cboCfdiUse)
    c.Street = Trim$(txtStreet.Text)
    c.ExtNumber = Trim$(txtExtNumber.Text)
    c.IntNumber = Trim$(txtIntNumber.Text)
    c.Neighborhood = Trim$(txtNeighborhood.Text)
    c.Municipality = Trim$(txtMunicipality.Text)
    c.State = Trim$(txtState.Text)
    Set ReadClientFromForm = c
End Function

' Recarga la lista con el texto del buscador (o con el que se indique).
Private Sub LoadClientList(ByVal filterText As String)
    Dim data As Variant, wasSuppressed As Boolean

    wasSuppressed = mSuppressEvents
    mSuppressEvents = True
    lstClients.Clear
    data = SearchClients(filterText)
    If Not IsEmpty(data) Then lstClients.List = data
    SelectInList mCurrent
    mSuppressEvents = wasSuppressed
End Sub

' Marca el cliente en la lista, o quita la selección si no está.
Private Sub SelectInList(ByVal clientName As String)
    Dim i As Long, wasSuppressed As Boolean

    wasSuppressed = mSuppressEvents
    mSuppressEvents = True
    lstClients.ListIndex = -1
    If Len(clientName) > 0 Then
        For i = 0 To lstClients.ListCount - 1
            If StrComp(CStr(lstClients.List(i)), clientName, vbTextCompare) = 0 Then
                lstClients.ListIndex = i
                Exit For
            End If
        Next i
    End If
    mSuppressEvents = wasSuppressed
End Sub


'==============================================================================
' EVENTOS DE LOS CAMPOS
'==============================================================================

Private Sub txtSearch_Change()
    If mSuppressEvents Then Exit Sub
    LoadClientList Trim$(txtSearch.Text)
End Sub

Private Sub lstClients_Click()
    Dim picked As String, c As Client
    On Error GoTo Fail

    If mSuppressEvents Then Exit Sub
    If lstClients.ListIndex < 0 Then Exit Sub
    picked = CStr(lstClients.List(lstClients.ListIndex))
    If StrComp(picked, mCurrent, vbTextCompare) = 0 Then Exit Sub

    If Not ConfirmDiscard() Then
        SelectInList mCurrent                     ' Se queda con el cliente que estaba
        Exit Sub
    End If
    Set c = LoadClient(picked)
    If c Is Nothing Then Exit Sub
    ShowClient c
    Exit Sub
Fail:
    mSuppressEvents = False
    ShowError "lstClients_Click"
End Sub

Private Sub txtName_Change()
    MarkDirty
End Sub

Private Sub txtPhone1_Change()
    MarkDirty
End Sub

Private Sub txtPhone2_Change()
    MarkDirty
End Sub

Private Sub txtEmail_Change()
    MarkDirty
End Sub

Private Sub txtAddress_Change()
    MarkDirty
End Sub

Private Sub txtLegalName_Change()
    MarkDirty
End Sub

Private Sub cboTaxRegime_Change()
    MarkDirty
End Sub

Private Sub cboCfdiUse_Change()
    MarkDirty
End Sub

Private Sub txtStreet_Change()
    MarkDirty
End Sub

Private Sub txtExtNumber_Change()
    MarkDirty
End Sub

Private Sub txtIntNumber_Change()
    MarkDirty
End Sub

Private Sub txtNeighborhood_Change()
    MarkDirty
End Sub

Private Sub txtMunicipality_Change()
    MarkDirty
End Sub

Private Sub txtState_Change()
    MarkDirty
End Sub

Private Sub txtRfc_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

' Cubre el texto pegado, que no pasa por KeyPress.
Private Sub txtRfc_Change()
    Dim pos As Long, wasSuppressed As Boolean

    If txtRfc.Text <> UCase$(txtRfc.Text) Then
        pos = txtRfc.SelStart
        wasSuppressed = mSuppressEvents
        mSuppressEvents = True                    ' Evita que el cambio de texto recurse
        txtRfc.Text = UCase$(txtRfc.Text)
        mSuppressEvents = wasSuppressed
        txtRfc.SelStart = pos
    End If
    UpdateHints
    MarkDirty
End Sub

Private Sub txtTaxZip_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 8 Then Exit Sub                 ' Retroceso
    If Not Chr$(KeyAscii) Like "[0-9]" Then KeyAscii = 0
End Sub

Private Sub txtTaxZip_Change()
    UpdateHints
    MarkDirty
End Sub

' Avisos junto al RFC y al CP fiscal; nunca bloquean el guardado.
Private Sub UpdateHints()
    If Len(Trim$(txtRfc.Text)) > 0 And Not IsValidRfc(txtRfc.Text) Then
        lblRfcHint.Caption = MSG_CLI_RFC_HINT
    Else
        lblRfcHint.Caption = vbNullString
    End If
    If Len(Trim$(txtTaxZip.Text)) > 0 And Not IsValidPostalCode(Trim$(txtTaxZip.Text)) Then
        lblZipHint.Caption = MSG_CLI_ZIP_HINT
    Else
        lblZipHint.Caption = vbNullString
    End If
End Sub


'==============================================================================
' ACCIONES
'==============================================================================

Private Sub btnNew_Click()
    On Error GoTo Fail
    If Not ConfirmDiscard() Then Exit Sub
    StartNewClient vbNullString
    Exit Sub
Fail:
    mSuppressEvents = False
    ShowError "btnNew_Click"
End Sub

Private Sub btnSave_Click()
    On Error GoTo Fail
    SaveCurrentClient True
    Exit Sub
Fail:
    mSuppressEvents = False
    ShowError "btnSave_Click"
End Sub

' Guarda el cliente en pantalla. Regresa True si quedó guardado. Un nombre que ya
' existe en una alta nueva no se duplica: se muestra el existente.
Private Function SaveCurrentClient(ByVal announce As Boolean) As Boolean
    Dim c As Client, existing As Client, otherName As String, message As String

    Set c = ReadClientFromForm()
    If Len(c.Name) = 0 Then
        Warn MSG_CLI_NAME_EMPTY, txtName
        Exit Function
    End If

    If Len(mCurrent) = 0 Then
        Set existing = LoadClient(c.Name)
        If Not existing Is Nothing Then
            MsgBox Fmt(MSG_CLI_NAME_EXISTS, existing.Name), vbExclamation, TTL_CLIENTS
            ShowClient existing
            Exit Function
        End If
    End If

    otherName = FindClientByRfc(c.Rfc, c.Name)
    If Not SaveClient(c) Then
        MsgBox MSG_CLI_SAVE_FAILED, vbExclamation, TTL_CLIENTS
        Exit Function
    End If

    mSuppressEvents = True
    mCurrent = c.Name
    txtName.Locked = True
    txtName.BackColor = Neutral50()
    txtRfc.Text = c.Rfc
    mSuppressEvents = False
    mIsDirty = False
    LoadClientList Trim$(txtSearch.Text)
    SelectInList mCurrent

    If announce Or Len(otherName) > 0 Then
        message = MSG_CLI_SAVED
        If Len(otherName) > 0 Then message = message & vbLf & Fmt(MSG_CLI_RFC_DUPLICATE, otherName)
        MsgBox message, IIf(Len(otherName) > 0, vbExclamation, vbInformation), TTL_CLIENTS
    End If
    SaveCurrentClient = True
End Function

' Elige el cliente en pantalla para quien abrió el formulario.
Private Sub btnUse_Click()
    Dim answer As VbMsgBoxResult
    On Error GoTo Fail

    If Len(mCurrent) = 0 Or mIsDirty Then
        If Len(Trim$(txtName.Text)) = 0 Then
            Warn MSG_CLI_NAME_EMPTY, txtName
            Exit Sub
        End If
        answer = MsgBox(MSG_CLI_USE_UNSAVED, vbYesNo + vbQuestion, TTL_CLIENTS)
        If answer = vbYes Then
            If Not SaveCurrentClient(False) Then Exit Sub
        ElseIf Len(mCurrent) = 0 Then
            Exit Sub                              ' Un cliente sin guardar no se puede usar
        End If
    End If

    mChosen = mCurrent
    mIsDirty = False
    Me.Hide
    Exit Sub
Fail:
    mSuppressEvents = False
    ShowError "btnUse_Click"
End Sub

Private Sub btnClose_Click()
    On Error GoTo Fail
    CloseWithoutChoosing
    Exit Sub
Fail:
    ShowError "btnClose_Click"
End Sub
