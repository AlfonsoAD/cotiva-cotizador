Attribute VB_Name = "modTheme"
'==============================================================================
' modTheme · Identidad visual de Cotiva
'
' Paleta aprobada: naranja #CF4A06 (primario) + azul tinta #1A2638 (acento).
' Son funciones y no constantes porque RGB() no se puede usar en un Const.
'==============================================================================
Option Explicit

Public Function BrandOrange() As Long         ' cotiva-600
    BrandOrange = RGB(207, 74, 6)
End Function

Public Function BrandInk() As Long            ' tinta-950
    BrandInk = RGB(26, 38, 56)
End Function

Public Function BrandInkLight() As Long       ' tinta-300
    BrandInkLight = RGB(170, 193, 228)
End Function

Public Function Neutral50() As Long
    Neutral50 = RGB(246, 247, 248)
End Function

Public Function Neutral100() As Long
    Neutral100 = RGB(235, 237, 239)
End Function

Public Function Neutral200() As Long
    Neutral200 = RGB(214, 217, 221)
End Function

Public Function Neutral300() As Long
    Neutral300 = RGB(188, 192, 198)
End Function

Public Function TextGray() As Long            ' neutral-700, cumple contraste AA
    TextGray = RGB(98, 102, 109)
End Function

' Devuelve "Archivo" si la fuente de marca está instalada en Windows; si no, "Segoe UI".
' Static guarda el resultado para no buscar en disco cada vez.
Public Function BrandFont() As String
    Static cachedFont As String
    If Len(cachedFont) = 0 Then
        cachedFont = "Segoe UI"
        If FontFileExists(Environ$("WINDIR") & "\Fonts\Archivo*.ttf") Then cachedFont = "Archivo"
        If FontFileExists(Environ$("LOCALAPPDATA") & "\Microsoft\Windows\Fonts\Archivo*.ttf") Then cachedFont = "Archivo"
    End If
    BrandFont = cachedFont
End Function

' True si existe algún archivo que coincida con el patrón (sin tronar si la ruta no existe).
Private Function FontFileExists(ByVal pattern As String) As Boolean
    On Error Resume Next
    FontFileExists = (Len(Dir$(pattern)) > 0)
    On Error GoTo 0
End Function

' Convierte "#CF4A06" (o "CF4A06") a un color de VBA.
'   color: (salida) el color convertido.
' Regresa False si el texto no es un color hexadecimal válido.
Public Function TryParseHexColor(ByVal hexText As String, ByRef parsedColor As Long) As Boolean
    hexText = UCase$(Replace(Trim$(hexText), "#", ""))
    If Len(hexText) <> 6 Then Exit Function
    If Not hexText Like "[0-9A-F][0-9A-F][0-9A-F][0-9A-F][0-9A-F][0-9A-F]" Then Exit Function

    parsedColor = RGB(CLng("&H" & Mid$(hexText, 1, 2)), _
                CLng("&H" & Mid$(hexText, 3, 2)), _
                CLng("&H" & Mid$(hexText, 5, 2)))
    TryParseHexColor = True
End Function

' Color de texto legible sobre un fondo: tinta si el fondo es claro, blanco si es oscuro.
' Usa la luminosidad percibida (los ojos ven el verde más brillante que el azul).
Public Function ContrastTextColor(ByVal background As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = background Mod 256
    g = (background \ 256) Mod 256
    b = background \ 65536

    If 0.299 * r + 0.587 * g + 0.114 * b > 160 Then
        ContrastTextColor = BrandInk()
    Else
        ContrastTextColor = vbWhite
    End If
End Function
