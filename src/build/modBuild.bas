Attribute VB_Name = "modBuild"
'==============================================================================
' modBuild — Solo para el build (build.ps1)
'
' Arma un libro nuevo: hojas del cotizador, portada con el botón, logos y
' protección. build.ps1 lo ejecuta y después lo ELIMINA del archivo final,
' así el .xlsm distribuido no contiene código de construcción.
'
' El diseño de la portada NO vive aquí, sino en modHome.DesignHome, para poder
' rediseñarla también en un libro ya distribuido.
'==============================================================================
Option Explicit

' Punto de entrada que llama build.ps1.
'   assetsFolder: carpeta "assets" del repositorio (logos opcionales).
Public Sub BuildWorkbook(ByVal assetsFolder As String)
    Dim home As Worksheet, i As Long

    Application.DisplayAlerts = False

    ' 1) Hoja Inicio: la primera hoja del libro nuevo; se borran las demás
    Set home = ThisWorkbook.Worksheets(1)
    home.Name = SHEET_HOME
    For i = ThisWorkbook.Worksheets.Count To 2 Step -1
        ThisWorkbook.Worksheets(i).Delete
    Next i

    ' 2) Hojas del cotizador (Plantilla, BD_, Mi empresa)
    EnsureWorkbookStructure

    ' 3) Logos opcionales desde assets/ (van antes del diseño: si la imagen
    '    existe, la portada y la Plantilla no dibujan el texto de respaldo)
    ImportLogo assetsFolder, "logo-cotiva.png", COTIVA_LOGO_NAME, False
    ImportLogo assetsFolder, "logo-cotiva.png", HOME_COTIVA_LOGO_NAME, True
    ImportLogo assetsFolder, "logo-innobytes.png", HOME_INNOBYTES_LOGO_NAME, True

    If Len(Dir$(assetsFolder & "\logo-form.jpg")) > 0 Then
        AddFormLogo assetsFolder & "\logo-form.jpg"
    End If

    ' 4) Portada
    DesignHome True

    ' 5) Protección y hoja inicial
    ProtectAll
    home.Activate
    Application.DisplayAlerts = True
End Sub

' Coloca un logo si el archivo existe en assets/ (si no, se omite sin error).
'   onHome: True = va en la hoja Inicio; False = va en la Plantilla.
Private Sub ImportLogo(ByVal assetsFolder As String, ByVal fileName As String, _
                       ByVal logoName As String, ByVal onHome As Boolean)
    Dim logoPath As String

    logoPath = assetsFolder & "\" & fileName
    If Len(Dir$(logoPath)) = 0 Then Exit Sub

    If onHome Then
        PlaceHomeLogo logoPath, logoName
    Else
        PlaceLogo logoPath, logoName
    End If
End Sub

' Agrega el logo al diseño del formulario como control Image.
' La imagen queda guardada dentro del archivo. Debe ser JPG, BMP o GIF
' (los formularios de VBA no aceptan PNG). Requiere acceso al modelo de
' objetos de VBA (build.ps1 lo activa).
Private Sub AddFormLogo(ByVal imagePath As String)
    Dim formDesigner As Object, img As Object

    On Error GoTo Fail
    Set formDesigner = ThisWorkbook.VBProject.VBComponents("frmQuote").Designer
    Set img = formDesigner.Controls.Add("Forms.Image.1", "imgLogo")
    img.Picture = LoadPicture(imagePath)
    img.AutoSize = False
    img.PictureSizeMode = 3                        ' fmPictureSizeModeZoom
    img.BorderStyle = 0                            ' fmBorderStyleNone
    img.BackStyle = 0                              ' fmBackStyleTransparent
    Exit Sub
Fail:
    Debug.Print "AddFormLogo: " & Err.Description  ' Sin logo, el formulario usa la barra con texto
End Sub
