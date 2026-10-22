Attribute VB_Name = "modTestCompany"
'==============================================================================
' modTestCompany - Colores de empresa y nombre del logo del formulario
'
' Las pruebas que escriben B11/B12 limpian en Cleanup aunque falle un assert,
' y vuelven a lanzar el error para que el runner lo registre.
'==============================================================================
Option Explicit

Public Sub Test_PrimaryColor_VaciaUsaCotiva()
    Dim errNumber As Long, errText As String
    On Error GoTo Cleanup
    SetCompanyCell COMPANY_PRIMARY_COLOR_CELL, ""
    AssertEqual BrandOrange(), PrimaryColor(), "B11 vacia"
Cleanup:
    errNumber = Err.Number: errText = Err.Description
    ClearCompanyColors
    If errNumber <> 0 Then Err.Raise errNumber, , errText
End Sub

Public Sub Test_PrimaryColor_InvalidaUsaCotiva()
    Dim errNumber As Long, errText As String
    On Error GoTo Cleanup
    SetCompanyCell COMPANY_PRIMARY_COLOR_CELL, "#12G45Z"
    AssertEqual BrandOrange(), PrimaryColor(), "hex con letra invalida"
    SetCompanyCell COMPANY_PRIMARY_COLOR_CELL, "12345"
    AssertEqual BrandOrange(), PrimaryColor(), "hex de 5 digitos"
Cleanup:
    errNumber = Err.Number: errText = Err.Description
    ClearCompanyColors
    If errNumber <> 0 Then Err.Raise errNumber, , errText
End Sub

Public Sub Test_PrimaryColor_Valida()
    Dim errNumber As Long, errText As String
    On Error GoTo Cleanup
    SetCompanyCell COMPANY_PRIMARY_COLOR_CELL, "#1a2b3c"
    AssertEqual RGB(&H1A, &H2B, &H3C), PrimaryColor(), "con # y minusculas"
    SetCompanyCell COMPANY_PRIMARY_COLOR_CELL, "1A2B3C"
    AssertEqual RGB(&H1A, &H2B, &H3C), PrimaryColor(), "sin # y mayusculas"
Cleanup:
    errNumber = Err.Number: errText = Err.Description
    ClearCompanyColors
    If errNumber <> 0 Then Err.Raise errNumber, , errText
End Sub

Public Sub Test_SecondaryColor_VaciaUsaCotiva()
    Dim errNumber As Long, errText As String
    On Error GoTo Cleanup
    SetCompanyCell COMPANY_SECONDARY_COLOR_CELL, ""
    AssertEqual BrandInk(), SecondaryColor(), "B12 vacia"
Cleanup:
    errNumber = Err.Number: errText = Err.Description
    ClearCompanyColors
    If errNumber <> 0 Then Err.Raise errNumber, , errText
End Sub

Public Sub Test_SecondaryColor_InvalidaUsaCotiva()
    Dim errNumber As Long, errText As String
    On Error GoTo Cleanup
    SetCompanyCell COMPANY_SECONDARY_COLOR_CELL, "#12G45Z"
    AssertEqual BrandInk(), SecondaryColor(), "hex con letra invalida"
    SetCompanyCell COMPANY_SECONDARY_COLOR_CELL, "12345"
    AssertEqual BrandInk(), SecondaryColor(), "hex de 5 digitos"
Cleanup:
    errNumber = Err.Number: errText = Err.Description
    ClearCompanyColors
    If errNumber <> 0 Then Err.Raise errNumber, , errText
End Sub

Public Sub Test_SecondaryColor_Valida()
    Dim errNumber As Long, errText As String
    On Error GoTo Cleanup
    SetCompanyCell COMPANY_SECONDARY_COLOR_CELL, "0F7F3A"
    AssertEqual RGB(&H0F, &H7F, &H3A), SecondaryColor(), "0F7F3A"
Cleanup:
    errNumber = Err.Number: errText = Err.Description
    ClearCompanyColors
    If errNumber <> 0 Then Err.Raise errNumber, , errText
End Sub

Public Sub Test_RefreshHomeColors_UsaColorDeEmpresa()
    Dim errNumber As Long, errText As String, ws As Worksheet
    On Error GoTo Cleanup
    SetCompanyCell COMPANY_PRIMARY_COLOR_CELL, "#0F7F3A"
    SetCompanyCell COMPANY_SECONDARY_COLOR_CELL, "#1A2B3C"
    RefreshHomeColors
    Set ws = ThisWorkbook.Worksheets(SHEET_HOME)
    AssertEqual RGB(&H0F, &H7F, &H3A), ws.Shapes(HOME_RULE_NAME).Fill.ForeColor.RGB, "relleno de la regla"
    AssertEqual RGB(&H0F, &H7F, &H3A), ws.Shapes(HOME_BUTTON_NAME).Fill.ForeColor.RGB, "relleno del boton"
    AssertEqual ContrastTextColor(RGB(&H0F, &H7F, &H3A)), _
                ws.Shapes(HOME_BUTTON_NAME).TextFrame2.TextRange.Font.Fill.ForeColor.RGB, "texto del boton"
    AssertEqual RGB(&H1A, &H2B, &H3C), ws.Range("B4").Font.Color, "letra del titulo"
    AssertEqual RGB(&H0F, &H7F, &H3A), _
                ws.Range("B" & HOME_FIRST_FEATURE_ROW).Characters(1, Len(HOME_BULLET)).Font.Color, "vineta"
Cleanup:
    errNumber = Err.Number: errText = Err.Description
    ClearCompanyColors
    RefreshHomeColors
    RefreshColorSwatches
    ProtectAll
    If errNumber <> 0 Then Err.Raise errNumber, , errText
End Sub

Public Sub Test_RefreshHomeColors_VaciasUsaCotiva()
    Dim errNumber As Long, errText As String, ws As Worksheet
    On Error GoTo Cleanup
    ClearCompanyColors
    RefreshHomeColors
    Set ws = ThisWorkbook.Worksheets(SHEET_HOME)
    AssertEqual BrandOrange(), ws.Shapes(HOME_BUTTON_NAME).Fill.ForeColor.RGB, "boton con el naranja de Cotiva"
Cleanup:
    errNumber = Err.Number: errText = Err.Description
    ClearCompanyColors
    RefreshHomeColors
    RefreshColorSwatches
    ProtectAll
    If errNumber <> 0 Then Err.Raise errNumber, , errText
End Sub

Public Sub Test_RefreshColorSwatches_PintaMiEmpresa()
    Dim errNumber As Long, errText As String, ws As Worksheet
    On Error GoTo Cleanup
    SetCompanyCell COMPANY_PRIMARY_COLOR_CELL, "#0F7F3A"
    SetCompanyCell COMPANY_SECONDARY_COLOR_CELL, "#1A2B3C"
    UnlockSheet ThisWorkbook.Worksheets(SHEET_COMPANY)
    RefreshColorSwatches
    Set ws = ThisWorkbook.Worksheets(SHEET_COMPANY)
    AssertEqual RGB(&H0F, &H7F, &H3A), ws.Range("A1").Font.Color, "titulo A1"
    AssertEqual RGB(&H1A, &H2B, &H3C), ws.Range("B4").Font.Color, "texto de captura B4"
Cleanup:
    errNumber = Err.Number: errText = Err.Description
    ClearCompanyColors
    UnlockSheet ThisWorkbook.Worksheets(SHEET_COMPANY)
    RefreshColorSwatches
    ProtectAll
    If errNumber <> 0 Then Err.Raise errNumber, , errText
End Sub

Public Sub Test_ProtectAll_ConservaScreenUpdating()
    Dim errNumber As Long, errText As String
    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    AssertTrue ProtectAll(), "ProtectAll regresa True"
    AssertEqual False, Application.ScreenUpdating, "sigue apagada"
    Application.ScreenUpdating = True
    AssertTrue ProtectAll(), "ProtectAll regresa True"
    AssertEqual True, Application.ScreenUpdating, "sigue encendida"
Cleanup:
    errNumber = Err.Number: errText = Err.Description
    Application.ScreenUpdating = True
    If errNumber <> 0 Then Err.Raise errNumber, , errText
End Sub

Public Sub Test_ProtectAll_NoReasignaHojasOcultas()
    Dim sheetName As Variant, ws As Worksheet
    AssertTrue ProtectAll(), "primera llamada"
    AssertTrue ProtectAll(), "segunda llamada"
    For Each sheetName In Array(SHEET_QUOTES, SHEET_ITEMS, SHEET_CLIENTS, SHEET_ORDERS)
        If SheetExists(CStr(sheetName)) Then
            Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
            AssertEqual xlSheetVeryHidden, ws.Visible, "oculta: " & sheetName
            AssertTrue ws.ProtectContents, "protegida: " & sheetName
        End If
    Next sheetName
End Sub

Public Sub Test_OpenCompanySheet_DejaPantallaEncendida()
    Dim errNumber As Long, errText As String
    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    OpenCompanySheet
    AssertEqual True, Application.ScreenUpdating, "pantalla encendida"
    AssertEqual SHEET_COMPANY, ActiveSheet.Name, "hoja activa"
    AssertTrue ThisWorkbook.Worksheets(SHEET_COMPANY).ProtectContents, "hoja protegida"
Cleanup:
    errNumber = Err.Number: errText = Err.Description
    Application.ScreenUpdating = True
    If errNumber <> 0 Then Err.Raise errNumber, , errText
End Sub

Public Sub Test_FormLogoFile_NombreUnico()
    AssertEqual "logo-form.jpg", FORM_LOGO_FILE, "nombre del logo del formulario"
End Sub
