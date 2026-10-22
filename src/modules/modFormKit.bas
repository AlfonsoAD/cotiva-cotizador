Attribute VB_Name = "modFormKit"
'==============================================================================
' modFormKit - Helpers de maquetado compartidos por los formularios
'
' Reciben el contenedor (UserForm o Page de un MultiPage) en "parent".
' BeginFormKit guarda la fuente y los colores una vez por BuildLayout.
'==============================================================================
Option Explicit

Public Enum ButtonKind
    kindPrimary = 0         ' Naranja: acción principal
    kindDark = 1            ' Tinta
    kindSecondary = 2       ' Blanco
End Enum

Public Const FORM_MARGIN As Single = 18

Private mKitFont As String
Private mKitPrimary As Long
Private mKitSecondary As Long

Public Sub BeginFormKit(ByVal fontName As String, ByVal primaryColor As Long, ByVal secondaryColor As Long)
    mKitFont = fontName
    mKitPrimary = primaryColor
    mKitSecondary = secondaryColor
End Sub

' Busca un control Image en el formulario (lo agrega el build) y lo acomoda.
' Regresa True si hay logo.
Public Function PlaceHeaderLogo(ByVal frm As Object) As Boolean
    Dim ctl As Object, img As Object, filePath As String

    For Each ctl In frm.Controls
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
        Set img = frm.Controls.Add("Forms.Image.1", "c_HeaderLogo", True)
        img.Picture = LoadPicture(filePath)
        On Error GoTo 0
    End If

    img.AutoSize = False                          ' Si no, vuelve al tamaño de la imagen
    Place img, FORM_MARGIN, 7, 112, 34
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

Public Sub Place(ByVal ctl As Object, ByVal x As Single, ByVal y As Single, _
                  ByVal w As Single, ByVal h As Single)
    ctl.Left = x
    ctl.Top = y
    ctl.Width = w
    ctl.Height = h
End Sub

' Estilo plano para cajas de texto, listas y combos: borde gris fino.
Public Sub StyleInput(ByVal ctl As Object)
    ctl.SpecialEffect = fmSpecialEffectFlat       ' Primero quitar el relieve 3D...
    ctl.BorderStyle = fmBorderStyleSingle         ' ...luego el borde sencillo
    ctl.BorderColor = Neutral300()
    ctl.BackColor = vbWhite
    ctl.ForeColor = mKitSecondary
    ctl.Font.Name = mKitFont
    ctl.Font.Size = 10
End Sub

Public Function AddTextBox(ByVal parent As Object, ByVal ctlName As String, ByVal x As Single, ByVal y As Single, _
                            ByVal w As Single, ByVal h As Single) As MSForms.TextBox
    Dim ctl As MSForms.TextBox
    Set ctl = parent.Controls.Add("Forms.TextBox.1", ctlName, True)
    Place ctl, x, y, w, h
    StyleInput ctl
    Set AddTextBox = ctl
End Function

Public Function AddComboBox(ByVal parent As Object, ByVal ctlName As String, ByVal x As Single, ByVal y As Single, _
                             ByVal w As Single, ByVal h As Single) As MSForms.ComboBox
    Dim ctl As MSForms.ComboBox
    Set ctl = parent.Controls.Add("Forms.ComboBox.1", ctlName, True)
    Place ctl, x, y, w, h
    StyleInput ctl
    Set AddComboBox = ctl
End Function

Public Function AddButton(ByVal parent As Object, ByVal ctlName As String, ByVal buttonText As String, ByVal kind As ButtonKind, _
                           ByVal x As Single, ByVal y As Single, _
                           ByVal w As Single, ByVal h As Single) As MSForms.CommandButton
    Dim ctl As MSForms.CommandButton
    Set ctl = parent.Controls.Add("Forms.CommandButton.1", ctlName, True)
    Place ctl, x, y, w, h
    ctl.Caption = buttonText
    ctl.Font.Name = mKitFont
    ctl.Font.Size = 9
    ctl.Font.Bold = True
    Select Case kind
        Case kindPrimary
            ctl.BackColor = mKitPrimary
            ctl.ForeColor = ContrastTextColor(mKitPrimary)
        Case kindDark
            ctl.BackColor = mKitSecondary
            ctl.ForeColor = ContrastTextColor(mKitSecondary)
        Case Else
            ctl.BackColor = vbWhite
            ctl.ForeColor = mKitSecondary
    End Select
    Set AddButton = ctl
End Function

' Etiqueta de texto.
Public Function AddLabel(ByVal parent As Object, ByVal ctlName As String, ByVal labelText As String, _
                          ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, _
                          ByVal fontSize As Single, ByVal isBold As Boolean, ByVal textColor As Long, _
                          Optional ByVal align As fmTextAlign = fmTextAlignLeft) As MSForms.Label
    Dim ctl As MSForms.Label
    Set ctl = parent.Controls.Add("Forms.Label.1", ctlName, True)
    Place ctl, x, y, w, h
    With ctl
        .Caption = labelText
        .BackStyle = fmBackStyleTransparent
        .WordWrap = False
        .TextAlign = align
        .Font.Name = mKitFont
        .Font.Size = fontSize
        .Font.Bold = isBold
        .ForeColor = textColor
    End With
    Set AddLabel = ctl
End Function

' Rectángulo de color (bandas y líneas).
'   toBack: True lo manda detrás de los demás controles.
Public Sub AddBand(ByVal parent As Object, ByVal ctlName As String, ByVal x As Single, ByVal y As Single, _
                    ByVal w As Single, ByVal h As Single, ByVal fillColor As Long, ByVal toBack As Boolean)
    Dim ctl As MSForms.Label
    Set ctl = parent.Controls.Add("Forms.Label.1", ctlName, True)
    Place ctl, x, y, w, h
    ctl.Caption = ""
    ctl.BackStyle = fmBackStyleOpaque
    ctl.BackColor = fillColor
    If toBack Then ctl.ZOrder fmZOrderBack
End Sub


Public Sub FocusOn(ByVal ctl As Object)
    On Error Resume Next                          ' Falla si el formulario aún no es visible
    ctl.SetFocus
    On Error GoTo 0
End Sub

' Muestra un aviso y lleva el cursor al control con problema.
Public Sub Warn(ByVal message As String, ByVal ctl As Object)
    MsgBox message, vbExclamation, TTL_REVIEW
    FocusOn ctl
    On Error Resume Next
    ctl.SelStart = 0
    ctl.SelLength = Len(ctl.Text)
    On Error GoTo 0
End Sub

