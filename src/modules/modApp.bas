Attribute VB_Name = "modApp"
'==============================================================================
' modApp · Punto de entrada del cotizador
'
' OpenQuoteTool es la macro del botón "Abrir cotizador". Antes de mostrar el
' formulario deja el libro listo: crea las hojas que falten, actualiza las de
' versiones anteriores y vuelve a activar la protección.
'==============================================================================
Option Explicit

' Macro del botón de la hoja Inicio.
Public Sub OpenQuoteTool()
    If Not PrepareWorkbook() Then Exit Sub
    frmQuote.Show vbModeless          ' vbModeless: se puede ver la Plantilla con el formulario abierto
End Sub

' Deja el libro listo para trabajar. Regresa False si algo impide abrir el cotizador.
Public Function PrepareWorkbook() As Boolean
    On Error GoTo Fail

    If Not TryUnprotectWorkbook() Then
        MsgBox Fmt(MSG_WRONG_PASSWORD), vbCritical, TTL_WRONG_PASSWORD
        Exit Function
    End If

    Application.ScreenUpdating = False
    EnsureWorkbookStructure
    PrepareWorkbook = ProtectAll()

Cleanup:
    Application.ScreenUpdating = True
    Exit Function
Fail:
    ShowError "PrepareWorkbook"
    Resume Cleanup
End Function

' Crea las hojas que falten y actualiza las de versiones anteriores.
' El libro debe estar desprotegido (lo hace PrepareWorkbook).
Public Sub EnsureWorkbookStructure()
    Dim previousSheet As Object
    Set previousSheet = ActiveSheet

    If Not SheetExists(SHEET_TEMPLATE) Then
        AddSheet SHEET_TEMPLATE
        DesignTemplate True
    End If
    If Not SheetExists(SHEET_QUOTES) Then AddSheet SHEET_QUOTES
    If Not SheetExists(SHEET_ITEMS) Then AddSheet SHEET_ITEMS

    If Not SheetExists(SHEET_CLIENTS) Then
        AddSheet SHEET_CLIENTS
        EnsureDataHeaders
        SeedClientsFromQuotes                      ' Migra clientes de cotizaciones existentes
    Else
        EnsureDataHeaders
    End If

    EnsureCompanySheet
    UpgradeCompanySheet

    If Not previousSheet Is Nothing Then previousSheet.Activate
End Sub

' Agrega una hoja al final del libro con el nombre indicado.
Private Function AddSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName
    Set AddSheet = ws
End Function
