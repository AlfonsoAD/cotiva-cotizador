Attribute VB_Name = "modTestHome"
'==============================================================================
' modTestHome - Boton "Clientes" de la hoja Inicio
'==============================================================================
Option Explicit

Public Sub Test_Inicio_BotonClientes()
    Dim ws As Worksheet

    TryUnprotectWorkbook
    Set ws = ThisWorkbook.Worksheets(SHEET_HOME)
    UnlockSheet ws
    If ShapeExists(ws, HOME_CLIENTS_BUTTON_NAME) Then ws.Shapes(HOME_CLIENTS_BUTTON_NAME).Delete

    AssertTrue PrepareWorkbook(), "PrepareWorkbook"
    AssertTrue ShapeExists(ws, HOME_CLIENTS_BUTTON_NAME), "el boton existe"
    If ShapeExists(ws, HOME_CLIENTS_BUTTON_NAME) Then
        AssertEqual "OpenClientsTool", ws.Shapes(HOME_CLIENTS_BUTTON_NAME).OnAction, "macro del boton"
    End If
    AssertTrue ShapeExists(ws, HOME_BUTTON_NAME), "el boton de abrir sigue"
End Sub
