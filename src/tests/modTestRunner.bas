Attribute VB_Name = "modTestRunner"
'==============================================================================
' modTestRunner - Corre cada Public Sub Test_*() de los modulos modTest*
'
' Es la excepcion a "solo modData toca BD_*": vive solo en el libro de pruebas y
' deja las hojas de datos como en un libro recien abierto antes de cada prueba.
'
' Dos formas de correr:
'  - RunAllTests (ventana Inmediato): corre todo desde VBA. Un error que nadie
'    maneja dentro de una prueba abre el dialogo de depuracion de Excel.
'  - ListTests / BeginTest / EndTest / Report: las usa build\test.ps1, que llama
'    cada prueba por COM. Un error sin manejar regresa a PowerShell, pero VBA
'    reinicia las variables de modulo: por eso el conteo y las fallas ya
'    anotadas viajan en PowerShell y no se guardan aqui entre pruebas.
'==============================================================================
Option Explicit

Private mCurrentName As String
Private mFailureLines As Collection

' Corre todas las pruebas desde VBA y regresa el reporte (ver Report).
Public Function RunAllTests() As String
    Dim testName As Variant, failureText As String, allFailures As String
    Dim passed As Long, failed As Long

    For Each testName In FindTests()
        BeginTest CStr(testName)
        On Error Resume Next
        Application.Run "'" & ThisWorkbook.Name & "'!" & CStr(testName)
        If Err.Number <> 0 Then RecordFailure "error " & Err.Number & ": " & Err.Description
        On Error GoTo 0
        failureText = EndTest()
        If Len(failureText) = 0 Then
            passed = passed + 1
        Else
            failed = failed + 1
            If Len(allFailures) > 0 Then allFailures = allFailures & vbLf
            allFailures = allFailures & failureText
        End If
    Next testName
    RunAllTests = Report(passed, failed, allFailures)
End Function

' "<modulo>.<prueba>" separados por "|", en orden de modulo y de linea; vacio si no hay.
Public Function ListTests() As String
    Dim testName As Variant
    For Each testName In FindTests()
        If Len(ListTests) > 0 Then ListTests = ListTests & "|"
        ListTests = ListTests & CStr(testName)
    Next testName
End Function

' Activa el modo prueba y despues deja el libro como recien abierto. Si eso falla, la prueba ya cuenta como falla.
Public Sub BeginTest(ByVal testName As String)
    mCurrentName = testName
    Set mFailureLines = New Collection

    On Error GoTo Fail
    SetTestMode True
    ResetDataSheets
    ClearLastError
    Exit Sub
Fail:
    RecordFailure "error " & Err.Number & ": " & Err.Description
End Sub

' Termina la prueba en curso. Regresa sus fallas (una linea "FALLA ..." cada una) o "" si paso.
Public Function EndTest() As String
    Dim line As Variant

    If Not mFailureLines Is Nothing Then
        For Each line In mFailureLines
            If Len(EndTest) > 0 Then EndTest = EndTest & vbLf
            EndTest = EndTest & CStr(line)
        Next line
    End If
    Set mFailureLines = Nothing
    SetTestMode False
    On Error Resume Next
    Application.ScreenUpdating = True
    On Error GoTo 0
End Function

' La usa modAssert para anotar una falla de la prueba en curso.
Public Sub RecordFailure(ByVal message As String)
    If mFailureLines Is Nothing Then Set mFailureLines = New Collection
    mFailureLines.Add "FALLA " & mCurrentName & ": " & message
End Sub

' Primera linea "ok N / falla M", luego una linea "FALLA <modulo>.<prueba>: <detalle>" por falla
' (test.ps1 antepone "--- FAIL: <modulo>.<prueba>" a las fallas de cada prueba).
' Sin ninguna prueba agrega la linea "no se encontraron pruebas".
Public Function Report(ByVal passed As Long, ByVal failed As Long, ByVal failureText As String) As String
    Report = "ok " & passed & " / falla " & failed
    If passed + failed = 0 Then Report = Report & vbLf & "no se encontraron pruebas"
    If Len(failureText) > 0 Then Report = Report & vbLf & failureText
End Function

' "<modulo>.<prueba>" de cada Public Sub Test_*() sin parametros, en orden de modulo y de linea.
' Enlace tardio: no hace falta la referencia a la biblioteca de extensibilidad de VBA.
Private Function FindTests() As Collection
    Dim result As New Collection, comp As Object, code As Object
    Dim i As Long, line As String, procName As String

    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 1 Then                        ' 1 = modulo estandar
            If comp.Name Like "modTest*" Then
                Set code = comp.CodeModule
                For i = 1 To code.CountOfLines
                    line = Trim$(code.Lines(i, 1))
                    If line Like "Public Sub Test_*()*" Then
                        procName = Mid$(line, Len("Public Sub ") + 1)
                        procName = Left$(procName, InStr(procName, "(") - 1)
                        result.Add comp.Name & "." & procName
                    End If
                Next i
            End If
        End If
    Next comp
    Set FindTests = result
End Function

' Hojas BD_* solo con encabezados, estructura completa y libro protegido.
Private Sub ResetDataSheets()
    Dim sheetName As Variant, ws As Worksheet

    TryUnprotectWorkbook
    For Each sheetName In Array(SHEET_QUOTES, SHEET_ITEMS, SHEET_CLIENTS, SHEET_ORDERS)
        If SheetExists(CStr(sheetName)) Then
            Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
            UnlockSheet ws
            ws.Rows("2:" & ws.Rows.Count).Delete
        End If
    Next sheetName
    EnsureWorkbookStructure
    If Not ProtectAll() Then Err.Raise vbObjectError + 1, , "ProtectAll fallo al preparar la prueba"
End Sub
