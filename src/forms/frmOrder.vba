'==============================================================================
' frmOrder · Formulario de la orden de servicio
'
' Igual que frmQuote: todo el diseño se construye por código en BuildLayout y
' este archivo contiene SOLO el código del formulario (build.ps1 crea el
' formulario y le inserta este código).
'
' Dos pestañas (MultiPage) para que quepa en 1366x768; el folio y el estado
' quedan arriba de las pestañas, siempre visibles.
'==============================================================================
Option Explicit

'--- Estado de la orden ---
Private mOrder As ServiceOrder          ' Orden en pantalla
Private mIsSaved As Boolean             ' True si el folio ya existe en la BD
Private mIsDirty As Boolean             ' True si hay cambios sin guardar
Private mSuppressEvents As Boolean      ' True mientras el CÓDIGO llena campos (no cuenta como cambio)
Private mOverwriteConfirmed As Boolean  ' True si ya se confirmó sobrescribir este folio
Private mFont As String                 ' Fuente de marca
Private mPrimary As Long                ' Color principal de la empresa
Private mSecondary As Long              ' Color secundario de la empresa

'--- Controles sin eventos ---
Private txtNumber As MSForms.TextBox
Private mpgMain As MSForms.MultiPage
Private lblVinHint As MSForms.Label

'--- Controles con eventos (WithEvents permite recibir sus eventos) ---
Private WithEvents cboStatus As MSForms.ComboBox
Private WithEvents cboClient As MSForms.ComboBox
Private WithEvents txtPhone1 As MSForms.TextBox
Private WithEvents txtPhone2 As MSForms.TextBox
Private WithEvents txtEmail As MSForms.TextBox
Private WithEvents txtAddress As MSForms.TextBox
Private WithEvents txtMake As MSForms.TextBox
Private WithEvents txtModel As MSForms.TextBox
Private WithEvents txtEngine As MSForms.TextBox
Private WithEvents txtYear As MSForms.TextBox
Private WithEvents txtColor As MSForms.TextBox
Private WithEvents txtPlates As MSForms.TextBox
Private WithEvents txtMileage As MSForms.TextBox
Private WithEvents cboFuel As MSForms.ComboBox
Private WithEvents txtVin As MSForms.TextBox
Private WithEvents txtInDate As MSForms.TextBox
Private WithEvents txtInTime As MSForms.TextBox
Private WithEvents cboReceivedBy As MSForms.ComboBox
Private WithEvents txtPromisedDate As MSForms.TextBox
Private WithEvents txtRequested As MSForms.TextBox
Private WithEvents txtDiagnosis As MSForms.TextBox
Private WithEvents txtOutDate As MSForms.TextBox
Private WithEvents txtOutTime As MSForms.TextBox
Private WithEvents cboDeliveredBy As MSForms.ComboBox
Private WithEvents cboHistory As MSForms.ComboBox
Private WithEvents btnLoad As MSForms.CommandButton
Private WithEvents btnSwitch As MSForms.CommandButton
Private WithEvents btnClients As MSForms.CommandButton
Private WithEvents btnNew As MSForms.CommandButton
Private WithEvents btnCompany As MSForms.CommandButton
Private WithEvents btnSave As MSForms.CommandButton
Private WithEvents btnPdf As MSForms.CommandButton
Private WithEvents btnToQuote As MSForms.CommandButton
Private WithEvents lblInnobytes As MSForms.Label

'--- Medidas del diseño (puntos) ---
Private Const FORM_WIDTH As Single = 660
Private Const FORM_HEIGHT As Single = 496
Private Const PAGE_CLIENT As Long = 0
Private Const PAGE_SERVICE As Long = 1


'==============================================================================
' ARRANQUE Y CIERRE
'==============================================================================

Private Sub UserForm_Initialize()
    BuildLayout
    LoadHistoryList
    LoadClientList
    LoadStaffLists
    StartNewOrder
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mIsDirty Then
        If MsgBox(MSG_UNSAVED_CLOSE_ORDER, vbYesNo + vbExclamation, TTL_CLOSE_ORDER) = vbNo Then Cancel = True
    End If
End Sub


'==============================================================================
' DISEÑO (se construye por código)
'==============================================================================

Private Sub BuildLayout()
    Dim rightEdge As Single, hasLogo As Boolean

    mFont = BrandFont()
    rightEdge = FORM_WIDTH - FORM_MARGIN

    ' --- Ventana (la diferencia Width-InsideWidth son los bordes de Windows) ---
    Me.Caption = UI_ORD_CAPTION
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
        AddLabel Me, "c_Module", UI_ORD_MODULE, 142, 17, 120, 18, 11, True, TextGray()
        Set btnSwitch = AddButton(Me, "c_Switch", UI_BTN_TO_QUOTE, kindSecondary, 268, 13, 100, 22)
    Else
        AddBand Me, "c_Bar", 0, 0, FORM_WIDTH, 48, mSecondary, True
        AddLabel Me, "c_Brand", UI_BRAND, FORM_MARGIN, 12, 80, 24, 16, True, ContrastTextColor(mSecondary)
        AddLabel Me, "c_Module", UI_ORD_MODULE, 104, 17, 120, 18, 11, False, BrandInkLight()
        Set btnSwitch = AddButton(Me, "c_Switch", UI_BTN_TO_QUOTE, kindSecondary, 230, 13, 100, 22)
    End If

    Set cboHistory = AddComboBox(Me, "c_History", rightEdge - 72 - 6 - 190, 13, 190, 22)
    With cboHistory
        .ColumnCount = 4                          ' Folio | Cliente | Placas | Estado
        .ColumnWidths = "56 pt;100 pt;56 pt;58 pt"
        .ListWidth = 270
        .BoundColumn = 1
        .TextColumn = 1
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryComplete
        .ControlTipText = TIP_ORD_HISTORY
    End With
    Set btnLoad = AddButton(Me, "c_Load", UI_BTN_LOAD, kindSecondary, rightEdge - 72, 13, 72, 22)

    ' --- Franja fija: folio y estado ---
    AddLabel Me, "c_LNumber", UI_LBL_NUMBER, FORM_MARGIN, 58, 90, 12, 8, False, TextGray()
    Set txtNumber = AddTextBox(Me, "c_Number", FORM_MARGIN, 72, 90, 22)
    With txtNumber
        .Locked = True                            ' El folio no se edita a mano
        .TabStop = False
        .BackColor = Neutral50()
        .Font.Bold = True
    End With
    AddLabel Me, "c_LStatus", UI_ORD_LBL_STATUS, 116, 58, 150, 12, 8, False, TextGray()
    Set cboStatus = AddComboBox(Me, "c_Status", 116, 72, 150, 22)
    cboStatus.Style = fmStyleDropDownList         ' Solo se puede elegir de la lista
    FillList cboStatus, ORDER_STATUSES

    Set btnToQuote = AddButton(Me, "c_ToQuote", UI_BTN_ORDER_TO_QUOTE, kindSecondary, rightEdge - 130, 72, 130, 22)
    btnToQuote.ControlTipText = TIP_ORDER_TO_QUOTE

    ' --- Pestañas ---
    Set mpgMain = Me.Controls.Add("Forms.MultiPage.1", "c_Pages", True)
    Place mpgMain, FORM_MARGIN, 102, FORM_WIDTH - 2 * FORM_MARGIN, 324
    mpgMain.Font.Name = mFont
    mpgMain.Font.Size = 9
    mpgMain.Pages(PAGE_CLIENT).Caption = UI_ORD_TAB_CLIENT
    mpgMain.Pages(PAGE_SERVICE).Caption = UI_ORD_TAB_SERVICE
    BuildClientPage mpgMain.Pages(PAGE_CLIENT)
    BuildServicePage mpgMain.Pages(PAGE_SERVICE)

    ' --- Barra inferior de acciones ---
    AddBand Me, "c_Footer", 0, 450, FORM_WIDTH, FORM_HEIGHT - 450, Neutral50(), True
    AddBand Me, "c_FooterLine", 0, 450, FORM_WIDTH, 1, Neutral200(), False
    Set btnNew = AddButton(Me, "c_New", UI_BTN_NEW_ORDER, kindSecondary, FORM_MARGIN, 460, 120, 26)
    Set btnCompany = AddButton(Me, "c_Company", UI_BTN_COMPANY, kindSecondary, FORM_MARGIN + 126, 460, 90, 26)
    btnCompany.ControlTipText = TIP_COMPANY
    Set btnSave = AddButton(Me, "c_Save", UI_BTN_SAVE, kindPrimary, rightEdge - 206, 460, 100, 26)
    Set btnPdf = AddButton(Me, "c_Pdf", UI_BTN_PDF, kindDark, rightEdge - 100, 460, 100, 26)

    ' --- Créditos de innobytes (clic = abrir el sitio) ---
    Set lblInnobytes = AddLabel(Me, "c_Innobytes", InnobytesCredit(), FORM_MARGIN, 430, 380, 12, 8, False, TextGray())
    lblInnobytes.ControlTipText = Fmt(TIP_INNOBYTES, INNOBYTES_WEB)
    lblInnobytes.MousePointer = fmMousePointerUpArrow

    ' --- Orden del tabulador (el de cada pestaña va en su propia Build...Page) ---
    cboStatus.TabIndex = 0
    mpgMain.TabIndex = 1
    btnSave.TabIndex = 2
    btnPdf.TabIndex = 3
    btnToQuote.TabIndex = 4
    btnNew.TabIndex = 5
    btnCompany.TabIndex = 6
    cboHistory.TabIndex = 7
    btnLoad.TabIndex = 8
    btnSwitch.TabIndex = 9
End Sub

' Pestaña 0: cliente y vehículo.
Private Sub BuildClientPage(ByVal pg As Object)
    Dim ctlTop As Single

    AddSection pg, "c_SecClient", UI_ORD_SEC_CLIENT, 8
    AddFieldLabel pg, "c_LClient", UI_ORD_LBL_CLIENT, 12, 24, 280
    AddFieldLabel pg, "c_LPhone1", UI_ORD_LBL_PHONE1, 300, 24, 140
    AddFieldLabel pg, "c_LPhone2", UI_ORD_LBL_PHONE2, 448, 24, 140
    Set cboClient = AddComboBox(pg, "c_Client", 12, 38, 200, 22)
    With cboClient
        .Style = fmStyleDropDownCombo             ' Se puede escribir un cliente nuevo
        .MatchEntry = fmMatchEntryComplete        ' Completa mientras escribes
        .MatchRequired = False
        .MaxLength = 100
        .ListRows = 10
        .ControlTipText = TIP_CLIENT
    End With
    Set btnClients = AddButton(pg, "c_Clients", UI_BTN_CLIENTS, kindSecondary, 218, 38, 74, 22)
    btnClients.ControlTipText = TIP_CLIENTS
    Set txtPhone1 = AddTextBox(pg, "c_Phone1", 300, 38, 140, 22)
    txtPhone1.MaxLength = 20
    Set txtPhone2 = AddTextBox(pg, "c_Phone2", 448, 38, 140, 22)
    txtPhone2.MaxLength = 20

    AddFieldLabel pg, "c_LEmail", UI_ORD_LBL_EMAIL, 12, 68, 280
    AddFieldLabel pg, "c_LAddress", UI_ORD_LBL_ADDRESS, 300, 68, 288
    Set txtEmail = AddTextBox(pg, "c_Email", 12, 82, 280, 22)
    txtEmail.MaxLength = 100
    Set txtAddress = AddTextBox(pg, "c_Address", 300, 82, 288, 22)
    txtAddress.MaxLength = 200

    AddSection pg, "c_SecVehicle", UI_ORD_SEC_VEHICLE, 118
    AddFieldLabel pg, "c_LMake", UI_ORD_LBL_MAKE, 12, 134, 130
    AddFieldLabel pg, "c_LModel", UI_ORD_LBL_MODEL, 150, 134, 130
    AddFieldLabel pg, "c_LEngine", UI_ORD_LBL_ENGINE, 288, 134, 130
    AddFieldLabel pg, "c_LYear", UI_ORD_LBL_YEAR, 426, 134, 60
    AddFieldLabel pg, "c_LColor", UI_ORD_LBL_COLOR, 494, 134, 94
    ctlTop = 148
    Set txtMake = AddTextBox(pg, "c_Make", 12, ctlTop, 130, 22)
    txtMake.MaxLength = 50
    Set txtModel = AddTextBox(pg, "c_Model", 150, ctlTop, 130, 22)
    txtModel.MaxLength = 50
    Set txtEngine = AddTextBox(pg, "c_Engine", 288, ctlTop, 130, 22)
    txtEngine.MaxLength = 50
    Set txtYear = AddTextBox(pg, "c_Year", 426, ctlTop, 60, 22)
    txtYear.MaxLength = 4
    Set txtColor = AddTextBox(pg, "c_Color", 494, ctlTop, 94, 22)
    txtColor.MaxLength = 30

    AddFieldLabel pg, "c_LPlates", UI_ORD_LBL_PLATES, 12, 178, 100
    AddFieldLabel pg, "c_LMileage", UI_ORD_LBL_MILEAGE, 120, 178, 100
    AddFieldLabel pg, "c_LFuel", UI_ORD_LBL_FUEL, 228, 178, 80
    AddFieldLabel pg, "c_LVin", UI_ORD_LBL_VIN, 316, 178, 272
    ctlTop = 192
    Set txtPlates = AddTextBox(pg, "c_Plates", 12, ctlTop, 100, 22)
    txtPlates.MaxLength = 10
    Set txtMileage = AddTextBox(pg, "c_Mileage", 120, ctlTop, 100, 22)
    txtMileage.MaxLength = 9
    txtMileage.TextAlign = fmTextAlignRight
    Set cboFuel = AddComboBox(pg, "c_Fuel", 228, ctlTop, 80, 22)
    cboFuel.Style = fmStyleDropDownList
    FillList cboFuel, FUEL_LEVELS
    Set txtVin = AddTextBox(pg, "c_Vin", 316, ctlTop, 272, 22)
    txtVin.MaxLength = 20
    Set lblVinHint = AddLabel(pg, "c_VinHint", "", 316, 216, 272, 12, 8, False, TextGray())

    cboClient.TabIndex = 0
    txtPhone1.TabIndex = 1
    txtPhone2.TabIndex = 2
    txtEmail.TabIndex = 3
    txtAddress.TabIndex = 4
    txtMake.TabIndex = 5
    txtModel.TabIndex = 6
    txtEngine.TabIndex = 7
    txtYear.TabIndex = 8
    txtColor.TabIndex = 9
    txtPlates.TabIndex = 10
    txtMileage.TabIndex = 11
    cboFuel.TabIndex = 12
    txtVin.TabIndex = 13
    btnClients.TabIndex = 14
End Sub

' Pestaña 1: entrada, servicio y entrega.
Private Sub BuildServicePage(ByVal pg As Object)
    AddSection pg, "c_SecIn", UI_ORD_SEC_IN, 8
    AddFieldLabel pg, "c_LInDate", UI_ORD_LBL_IN_DATE, 12, 24, 190
    AddFieldLabel pg, "c_LInTime", UI_ORD_LBL_TIME, 210, 24, 80
    AddFieldLabel pg, "c_LReceivedBy", UI_ORD_LBL_RECEIVED_BY, 298, 24, 150
    AddFieldLabel pg, "c_LPromised", UI_ORD_LBL_PROMISED, 456, 24, 132
    Set txtInDate = AddTextBox(pg, "c_InDate", 12, 38, 190, 22)
    txtInDate.MaxLength = 10
    Set txtInTime = AddTextBox(pg, "c_InTime", 210, 38, 80, 22)
    txtInTime.MaxLength = 5
    Set cboReceivedBy = AddComboBox(pg, "c_ReceivedBy", 298, 38, 150, 22)
    SetupStaffCombo cboReceivedBy
    Set txtPromisedDate = AddTextBox(pg, "c_PromisedDate", 456, 38, 132, 22)
    txtPromisedDate.MaxLength = 10

    AddSection pg, "c_SecService", UI_ORD_SEC_SERVICE, 70
    AddFieldLabel pg, "c_LRequested", UI_ORD_LBL_REQUESTED, 12, 86, 576
    Set txtRequested = AddTextBox(pg, "c_Requested", 12, 100, 576, 22)
    txtRequested.MaxLength = 255
    AddFieldLabel pg, "c_LDiagnosis", UI_ORD_LBL_DIAGNOSIS, 12, 128, 576
    Set txtDiagnosis = AddTextBox(pg, "c_Diagnosis", 12, 142, 576, 64)
    With txtDiagnosis
        .MultiLine = True
        .EnterKeyBehavior = True
        .ScrollBars = fmScrollBarsVertical
        .MaxLength = 1000
    End With

    AddSection pg, "c_SecOut", UI_ORD_SEC_OUT, 222
    AddFieldLabel pg, "c_LOutDate", UI_ORD_LBL_OUT_DATE, 12, 238, 190
    AddFieldLabel pg, "c_LOutTime", UI_ORD_LBL_TIME, 210, 238, 80
    AddFieldLabel pg, "c_LDeliveredBy", UI_ORD_LBL_DELIVERED_BY, 298, 238, 150
    Set txtOutDate = AddTextBox(pg, "c_OutDate", 12, 252, 190, 22)
    txtOutDate.MaxLength = 10
    Set txtOutTime = AddTextBox(pg, "c_OutTime", 210, 252, 80, 22)
    txtOutTime.MaxLength = 5
    Set cboDeliveredBy = AddComboBox(pg, "c_DeliveredBy", 298, 252, 150, 22)
    SetupStaffCombo cboDeliveredBy

    txtInDate.TabIndex = 0
    txtInTime.TabIndex = 1
    cboReceivedBy.TabIndex = 2
    txtPromisedDate.TabIndex = 3
    txtRequested.TabIndex = 4
    txtDiagnosis.TabIndex = 5
    txtOutDate.TabIndex = 6
    txtOutTime.TabIndex = 7
    cboDeliveredBy.TabIndex = 8
End Sub

' Título de sección: mayúsculas pequeñas en el color principal.
Private Sub AddSection(ByVal pg As Object, ByVal ctlName As String, ByVal sectionText As String, ByVal y As Single)
    AddLabel pg, ctlName, sectionText, 12, y, 300, 12, 8, True, mPrimary
End Sub

' Etiqueta pequeña sobre un campo.
Private Sub AddFieldLabel(ByVal pg As Object, ByVal ctlName As String, ByVal labelText As String, _
                          ByVal x As Single, ByVal y As Single, ByVal w As Single)
    AddLabel pg, ctlName, labelText, x, y, w, 12, 8, False, TextGray()
End Sub

' Combo de "Recibió" / "Entregó": se elige de lo ya guardado o se escribe un nombre nuevo.
Private Sub SetupStaffCombo(ByVal cbo As MSForms.ComboBox)
    With cbo
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryComplete
        .MatchRequired = False
        .MaxLength = 60
        .ListRows = 10
        .ControlTipText = TIP_STAFF
    End With
End Sub

' Llena un combo de lista fija con valores separados por "|".
Private Sub FillList(ByVal cbo As MSForms.ComboBox, ByVal values As String)
    Dim item As Variant
    For Each item In Split(values, "|")
        cbo.AddItem CStr(item)
    Next item
End Sub


'==============================================================================
' ESTADO DEL FORMULARIO
'==============================================================================

' Deja el formulario listo para una orden nueva.
Private Sub StartNewOrder()
    mSuppressEvents = True

    Set mOrder = New ServiceOrder
    mOrder.Number = GetNextOrderNumber()
    mOrder.Status = ORDER_STATUS_RECEIVED
    mOrder.InDate = Date
    mOrder.InTime = TimeValue(Format$(Time, "hh:nn"))
    mIsSaved = False
    mOverwriteConfirmed = False

    ShowOrder mOrder
    txtMileage.Text = ""                          ' Kilometraje y gasolina se capturan siempre
    cboFuel.ListIndex = -1
    cboHistory.ListIndex = -1
    cboHistory.Text = ""
    mpgMain.Value = PAGE_CLIENT

    mSuppressEvents = False
    mIsDirty = False
    FocusOn cboClient
End Sub

' Cambio hecho por el usuario (no por el código).
Private Sub MarkDirty()
    If Not mSuppressEvents Then mIsDirty = True
End Sub

' Llena los controles con los datos de una orden (no cuenta como cambio).
Private Sub ShowOrder(ByVal o As ServiceOrder)
    Dim wasSuppressed As Boolean

    wasSuppressed = mSuppressEvents
    mSuppressEvents = True

    txtNumber.Text = o.Number
    SelectOrAdd cboStatus, o.Status
    cboClient.Text = o.ClientName
    txtPhone1.Text = o.Phone1
    txtPhone2.Text = o.Phone2
    txtEmail.Text = o.Email
    txtAddress.Text = o.Address
    txtMake.Text = o.Make
    txtModel.Text = o.Model
    txtEngine.Text = o.Engine
    If o.Year > 0 Then txtYear.Text = CStr(o.Year) Else txtYear.Text = ""
    txtColor.Text = o.Color
    txtPlates.Text = o.Plates
    txtMileage.Text = CStr(o.Mileage)
    SelectOrAdd cboFuel, o.FuelLevel
    txtVin.Text = o.Vin
    txtInDate.Text = DateText(o.InDate)
    txtInTime.Text = TimeText(o.InTime)
    cboReceivedBy.Text = o.ReceivedBy
    txtPromisedDate.Text = DateText(o.PromisedDate)
    txtRequested.Text = o.Requested
    txtDiagnosis.Text = o.Diagnosis
    txtOutDate.Text = DateText(o.OutDate)
    txtOutTime.Text = TimeText(o.OutTime)
    cboDeliveredBy.Text = o.DeliveredBy
    UpdateVinHint

    mSuppressEvents = wasSuppressed
End Sub

Private Function DateText(ByVal d As Date) As String
    If d <> 0 Then DateText = Format$(d, "dd/mm/yyyy")
End Function

Private Function TimeText(ByVal t As Date) As String
    If t <> 0 Then TimeText = Format$(t, "hh:nn")
End Function

' Elige el valor en un combo de lista; si no está, lo agrega para no perderlo.
Private Sub SelectOrAdd(ByVal cbo As MSForms.ComboBox, ByVal value As String)
    Dim i As Long

    value = Trim$(value)
    If Len(value) = 0 Then
        cbo.ListIndex = -1
        Exit Sub
    End If
    For i = 0 To cbo.ListCount - 1
        If StrComp(cbo.List(i), value, vbTextCompare) = 0 Then
            cbo.ListIndex = i
            Exit Sub
        End If
    Next i
    cbo.AddItem value
    cbo.ListIndex = cbo.ListCount - 1
End Sub

Private Sub cboStatus_Change()
    MarkDirty
End Sub

Private Sub cboFuel_Change()
    MarkDirty
End Sub

Private Sub cboReceivedBy_Change()
    MarkDirty
End Sub

Private Sub cboDeliveredBy_Change()
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

Private Sub txtMake_Change()
    MarkDirty
End Sub

Private Sub txtModel_Change()
    MarkDirty
End Sub

Private Sub txtEngine_Change()
    MarkDirty
End Sub

Private Sub txtYear_Change()
    MarkDirty
End Sub

Private Sub txtColor_Change()
    MarkDirty
End Sub

Private Sub txtMileage_Change()
    MarkDirty
End Sub

Private Sub txtInDate_Change()
    MarkDirty
End Sub

Private Sub txtInTime_Change()
    MarkDirty
End Sub

Private Sub txtPromisedDate_Change()
    MarkDirty
End Sub

Private Sub txtRequested_Change()
    MarkDirty
End Sub

Private Sub txtDiagnosis_Change()
    MarkDirty
End Sub

Private Sub txtOutDate_Change()
    MarkDirty
End Sub

Private Sub txtOutTime_Change()
    MarkDirty
End Sub


'==============================================================================
' CLIENTE
'==============================================================================

Private Sub cboClient_Change()
    MarkDirty
End Sub

' Elegir de la lista llena el contacto.
Private Sub cboClient_Click()
    FillClientContact
End Sub

' Enter o Tab confirman el nombre escrito; si ya existe, llena el contacto.
Private Sub cboClient_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then FillClientContact
End Sub

' Llena teléfonos, correo y dirección del cliente, SOLO en los campos vacíos.
Private Sub FillClientContact()
    Dim clientText As String, phone1 As String, phone2 As String, email As String, address As String

    If mSuppressEvents Then Exit Sub
    clientText = Trim$(cboClient.Text)
    If Len(clientText) = 0 Then Exit Sub
    If Not ClientExists(clientText) Then Exit Sub
    If Not GetClientContact(clientText, phone1, phone2, email, address) Then Exit Sub

    FillIfEmpty txtPhone1, phone1
    FillIfEmpty txtPhone2, phone2
    FillIfEmpty txtEmail, email
    FillIfEmpty txtAddress, address
End Sub

' Abre el catálogo de clientes. Al elegir uno se reemplaza el contacto con el del catálogo;
' sin elegir, el combo conserva lo escrito.
Private Sub btnClients_Click()
    Dim chosen As String, phone1 As String, phone2 As String, email As String, address As String
    On Error GoTo Fail

    chosen = PickClient(Trim$(cboClient.Text), True)
    LoadClientList
    If Len(chosen) = 0 Then Exit Sub

    cboClient.Text = chosen
    GetClientContact chosen, phone1, phone2, email, address
    txtPhone1.Text = phone1
    txtPhone2.Text = phone2
    txtEmail.Text = email
    txtAddress.Text = address
    Exit Sub
Fail:
    ShowError "btnClients_Click"
End Sub

Private Sub FillIfEmpty(ByVal txt As MSForms.TextBox, ByVal value As String)
    If Len(Trim$(txt.Text)) = 0 And Len(value) > 0 Then txt.Text = value
End Sub


'==============================================================================
' VEHÍCULO: FILTROS DE TECLADO Y AVISO DEL VIN
'==============================================================================

Private Sub txtMileage_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    FilterDigitKey KeyAscii
End Sub

Private Sub txtYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    FilterDigitKey KeyAscii
End Sub

' Solo deja pasar dígitos y retroceso.
Private Sub FilterDigitKey(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 8 Then Exit Sub                 ' Retroceso
    If Not Chr$(KeyAscii) Like "[0-9]" Then KeyAscii = 0
End Sub

Private Sub txtPlates_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtVin_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

' Cubre el texto pegado, que no pasa por KeyPress.
Private Sub txtPlates_Change()
    UppercaseText txtPlates
    MarkDirty
End Sub

Private Sub txtVin_Change()
    UppercaseText txtVin
    UpdateVinHint
    MarkDirty
End Sub

Private Sub UppercaseText(ByVal txt As MSForms.TextBox)
    Dim pos As Long, wasSuppressed As Boolean

    If txt.Text = UCase$(txt.Text) Then Exit Sub
    pos = txt.SelStart
    wasSuppressed = mSuppressEvents
    mSuppressEvents = True                        ' Evita que el cambio de texto recurse
    txt.Text = UCase$(txt.Text)
    mSuppressEvents = wasSuppressed
    txt.SelStart = pos
End Sub

' Aviso junto al VIN con el largo actual; nunca bloquea el guardado.
Private Sub UpdateVinHint()
    Dim vinText As String

    vinText = Trim$(txtVin.Text)
    If Len(vinText) = 0 Or IsValidVin(vinText) Then
        lblVinHint.Caption = ""
    Else
        lblVinHint.Caption = Fmt(MSG_VIN_HINT, Len(vinText))
    End If
End Sub


'==============================================================================
' VALIDACIÓN Y SINCRONIZACIÓN
'==============================================================================

' Cambia a la pestaña del campo y avisa.
Private Sub WarnOnPage(ByVal pageIndex As Long, ByVal message As String, ByVal ctl As Object)
    mpgMain.Value = pageIndex
    Warn message, ctl
End Sub

' Pasa los datos del formulario a mOrder, validándolos en el orden del formulario.
' Regresa False (y avisa) si algo falta o es inválido.
Private Function SyncOrderFromForm() As Boolean
    Dim clientText As String

    clientText = Trim$(cboClient.Text)
    If Len(clientText) = 0 Then
        WarnOnPage PAGE_CLIENT, MSG_CLIENT_EMPTY, cboClient: Exit Function
    End If
    If Len(Trim$(txtYear.Text)) > 0 And Not IsValidVehicleYear(txtYear.Text) Then
        WarnOnPage PAGE_CLIENT, Fmt(MSG_YEAR_INVALID, Year(Date) + 1), txtYear: Exit Function
    End If
    If Not IsValidMileage(txtMileage.Text) Then
        WarnOnPage PAGE_CLIENT, MSG_MILEAGE_REQUIRED, txtMileage: Exit Function
    End If
    If cboFuel.ListIndex < 0 Then
        WarnOnPage PAGE_CLIENT, MSG_FUEL_REQUIRED, cboFuel: Exit Function
    End If
    If Not IsValidDate(txtInDate.Text) Then
        WarnOnPage PAGE_SERVICE, MSG_DATE_INVALID, txtInDate: Exit Function
    End If
    If Len(Trim$(txtPromisedDate.Text)) > 0 And Not IsValidDate(txtPromisedDate.Text) Then
        WarnOnPage PAGE_SERVICE, MSG_DATE_INVALID, txtPromisedDate: Exit Function
    End If
    If Len(Trim$(txtOutDate.Text)) > 0 And Not IsValidDate(txtOutDate.Text) Then
        WarnOnPage PAGE_SERVICE, MSG_DATE_INVALID, txtOutDate: Exit Function
    End If
    If Len(Trim$(txtInTime.Text)) > 0 And Not IsValidTime(txtInTime.Text) Then
        WarnOnPage PAGE_SERVICE, MSG_TIME_INVALID, txtInTime: Exit Function
    End If
    If Len(Trim$(txtOutTime.Text)) > 0 And Not IsValidTime(txtOutTime.Text) Then
        WarnOnPage PAGE_SERVICE, MSG_TIME_INVALID, txtOutTime: Exit Function
    End If

    If TextLineCount(txtDiagnosis.Text) > ORD_DIAG_MAX_LINES Then
        WarnOnPage PAGE_SERVICE, Fmt(MSG_DIAGNOSIS_TOO_MANY_LINES, ORD_DIAG_MAX_LINES, TextLineCount(txtDiagnosis.Text)), txtDiagnosis: Exit Function
    End If

    ' Number y QuoteNumber no se tocan: el folio lo asigna el sistema y la cotización ligada se conserva
    With mOrder
        .Status = Trim$(cboStatus.Text)
        .ClientName = clientText
        .Phone1 = Trim$(txtPhone1.Text)
        .Phone2 = Trim$(txtPhone2.Text)
        .Email = Trim$(txtEmail.Text)
        .Address = Trim$(txtAddress.Text)
        .Make = Trim$(txtMake.Text)
        .Model = Trim$(txtModel.Text)
        .Engine = Trim$(txtEngine.Text)
        If Len(Trim$(txtYear.Text)) > 0 Then .Year = CInt(Trim$(txtYear.Text)) Else .Year = 0
        .Color = Trim$(txtColor.Text)
        .Plates = UCase$(Trim$(txtPlates.Text))
        .Mileage = CLng(Trim$(txtMileage.Text))
        .FuelLevel = Trim$(cboFuel.Text)
        .Vin = UCase$(Trim$(txtVin.Text))
        .InDate = CDate(Trim$(txtInDate.Text))
        .InTime = ParseTime(txtInTime.Text)
        .ReceivedBy = Trim$(cboReceivedBy.Text)
        .PromisedDate = ParseDate(txtPromisedDate.Text)
        .OutDate = ParseDate(txtOutDate.Text)
        .OutTime = ParseTime(txtOutTime.Text)
        .DeliveredBy = Trim$(cboDeliveredBy.Text)
        .Requested = Trim$(txtRequested.Text)
        .Diagnosis = Trim$(txtDiagnosis.Text)
    End With
    SyncOrderFromForm = True
End Function

' Fecha ya validada, o 0 si el campo va vacío.
Private Function ParseDate(ByVal dateText As String) As Date
    dateText = Trim$(dateText)
    If Len(dateText) > 0 Then ParseDate = CDate(dateText)
End Function

' Hora ya validada, o 0 si el campo va vacío.
Private Function ParseTime(ByVal timeText As String) As Date
    timeText = Trim$(timeText)
    If Len(timeText) > 0 Then ParseTime = TimeValue(timeText)
End Function


'==============================================================================
' ACCIONES PRINCIPALES
'==============================================================================

Private Sub btnSave_Click()
    On Error GoTo Fail
    If SaveCurrentOrder() Then
        MsgBox Fmt(MSG_ORDER_SAVED, mOrder.Number), vbInformation
    End If
    Exit Sub
Fail:
    ShowError "btnSave_Click"
End Sub

' Guarda la orden en pantalla. Regresa True si se guardó.
' Solo pregunta "¿sobrescribir?" la primera vez para una orden cargada del historial.
Private Function SaveCurrentOrder() As Boolean
    If Not SyncOrderFromForm() Then Exit Function

    If OrderExists(mOrder.Number) Then
        If mIsSaved Then
            If Not mOverwriteConfirmed Then
                If MsgBox(Fmt(MSG_CONFIRM_OVERWRITE, mOrder.Number), _
                          vbYesNo + vbQuestion, TTL_OVERWRITE) = vbNo Then Exit Function
                mOverwriteConfirmed = True
            End If
        Else
            ' Orden nueva cuyo folio ya ocupó otra: se asigna el siguiente libre
            mOrder.Number = GetNextOrderNumber()
            txtNumber.Text = mOrder.Number
            MsgBox Fmt(MSG_NUMBER_REASSIGNED, mOrder.Number), vbInformation
        End If
    End If

    If SaveOrder(mOrder) Then
        LoadHistoryList
        LoadClientList                            ' Por si fue un cliente nuevo
        LoadStaffLists                            ' Por si hay un nombre nuevo
        mIsSaved = True
        mOverwriteConfirmed = True                ' Ya es "nuestra": no vuelve a preguntar
        mIsDirty = False                          ' Al final, para que nada lo vuelva a marcar
        SaveCurrentOrder = True
    End If
End Function

Private Sub btnPdf_Click()
    Dim filePath As String, answer As VbMsgBoxResult
    On Error GoTo Fail

    If Not SyncOrderFromForm() Then Exit Sub

    If mIsDirty Or Not mIsSaved Then
        answer = MsgBox(Fmt(MSG_SAVE_ORDER_BEFORE_PDF), vbYesNoCancel + vbQuestion, TTL_PDF)
        If answer = vbCancel Then Exit Sub
        If answer = vbYes Then
            If Not SaveCurrentOrder() Then Exit Sub
        End If
    End If

    If Not ConfirmCompanyData() Then Exit Sub
    FillOrderTemplate mOrder
    filePath = ExportOrderPdf(mOrder)
    If Len(filePath) > 0 Then MsgBox Fmt(MSG_PDF_CREATED, filePath), vbInformation
    Exit Sub
Fail:
    ShowError "btnPdf_Click"
End Sub

Private Sub btnNew_Click()
    On Error GoTo Fail
    If mIsDirty Then
        If MsgBox(MSG_UNSAVED_NEW_ORDER, vbYesNo + vbExclamation, TTL_NEW_ORDER) = vbNo Then Exit Sub
    End If
    StartNewOrder
    Exit Sub
Fail:
    mSuppressEvents = False
    ShowError "btnNew_Click"
End Sub

' Va a la cotización; si hay cambios, pregunta antes.
Private Sub btnSwitch_Click()
    On Error GoTo Fail
    If mIsDirty Then
        If MsgBox(Fmt(MSG_UNSAVED_SWITCH, UI_QUOTE_SWITCH_TARGET), vbYesNo + vbExclamation, TTL_REVIEW) = vbNo Then Exit Sub
    End If
    mIsDirty = False                              ' Para que QueryClose no vuelva a preguntar
    ShowQuoteForm
    Unload Me
    Exit Sub
Fail:
    ShowError "btnSwitch_Click"
End Sub

' Guarda la orden y abre su cotizacion (la ligada, o una nueva con el diagnostico).
Private Sub btnToQuote_Click()
    Dim linked As String, quoteNumber As String, lineCount As Long
    On Error GoTo Fail

    If Not SaveCurrentOrder() Then Exit Sub

    linked = UCase$(Trim$(mOrder.QuoteNumber))
    If Len(linked) > 0 And QuoteExists(linked) Then
        quoteNumber = linked
    Else
        If Len(linked) > 0 Then
            MsgBox Fmt(MSG_LINKED_QUOTE_MISSING, linked), vbExclamation, TTL_REVIEW
        End If
        quoteNumber = CreateQuoteForOrder(mOrder)
        If Len(quoteNumber) = 0 Then Exit Sub
        lineCount = DiagnosisLines(mOrder.Diagnosis).Count
        If lineCount > MAX_ITEMS Then
            MsgBox Fmt(MSG_DIAGNOSIS_TRUNCATED, lineCount, MAX_ITEMS), vbInformation
        End If
    End If

    mIsDirty = False                              ' Para que QueryClose no pregunte
    ShowQuoteFormWith quoteNumber
    Unload Me
    Exit Sub
Fail:
    ShowError "btnToQuote_Click"
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
' HISTORIAL, CLIENTES Y PERSONAL
'==============================================================================

Private Sub LoadHistoryList()
    Dim data As Variant
    cboHistory.Clear
    data = ListOrders()
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

' Recarga los nombres de "Recibió" y "Entregó" conservando lo escrito.
Private Sub LoadStaffLists()
    Dim data As Variant, wasSuppressed As Boolean

    wasSuppressed = mSuppressEvents
    mSuppressEvents = True

    data = ListOrderStaff()
    ReloadStaffCombo cboReceivedBy, data
    ReloadStaffCombo cboDeliveredBy, data

    mSuppressEvents = wasSuppressed
End Sub

Private Sub ReloadStaffCombo(ByVal cbo As MSForms.ComboBox, ByVal names As Variant)
    Dim typed As String, i As Long

    typed = cbo.Text
    cbo.Clear
    If Not IsEmpty(names) Then
        For i = LBound(names) To UBound(names)
            cbo.AddItem CStr(names(i))
        Next i
    End If
    cbo.Text = typed
End Sub

Private Sub btnLoad_Click()
    Dim orderNumber As String, loaded As ServiceOrder
    On Error GoTo Fail

    orderNumber = UCase$(Trim$(cboHistory.Text))
    If Len(orderNumber) = 0 Then
        MsgBox MSG_HISTORY_EMPTY, vbInformation, TTL_REVIEW
        Exit Sub
    End If

    Set loaded = LoadOrder(orderNumber)
    If loaded Is Nothing Then
        MsgBox Fmt(MSG_QUOTE_NOT_FOUND, orderNumber), vbExclamation, TTL_REVIEW
        Exit Sub
    End If

    If mIsDirty Then
        If MsgBox(Fmt(MSG_UNSAVED_LOAD, orderNumber), vbYesNo + vbExclamation, TTL_REVIEW) = vbNo Then Exit Sub
    End If

    mSuppressEvents = True                        ' Llenar campos desde la BD no es un cambio
    Set mOrder = loaded
    ShowOrder mOrder
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
