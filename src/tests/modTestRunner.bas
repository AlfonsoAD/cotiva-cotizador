Attribute VB_Name = "modTestRunner"
'==============================================================================
' modTestRunner - Corre cada Public Sub Test_*() de los modulos modTest*
'
' Es la excepcion a "solo modData toca BD_*": vive solo en el libro de pruebas y
' deja las hojas de datos como en un libro recien abierto antes de cada prueba.
'==============================================================================
Option Explicit

Private mCurrentName As String
Private mCurrentFailed As Boolean
Private mFailureLines As Collection
Private mEventsBefore As Boolean

' Reporte: primera linea "ok N / falla M", luego "FALLA <modulo>.<prueba>: <detalle>".
Public Function RunAllTests() As String
    Dim tests As Collection, testName As Variant, line As Variant
    Dim passed As Long, failed As Long, report As String

    Set mFailureLines = New Collection
    Set tests = FindTests()
    If tests.Count = 0 Then
        RunAllTests = "ok 0 / falla 0" & vbLf & "no se encontraron pruebas"
        Exit Function
    End If

    mEventsBefore = Application.EnableEvents
    For Each testName In tests
        If RunOne(CStr(testName)) Then passed = passed + 1 Else failed = failed + 1
    Next testName
    SetTestMode False

    report = "ok " & passed & " / falla " & failed
    For Each line In mFailureLines
        report = report & vbLf & CStr(line)
    Next line
    RunAllTests = report
End Function

' La usa modAssert para anotar una falla de la prueba en curso.
Public Sub RecordFailure(ByVal message As String)
    If mFailureLines Is Nothing Then Set mFailureLines = New Collection
    mFailureLines.Add "FALLA " & mCurrentName & ": " & message
    mCurrentFailed = True
End Sub

' Corre una prueba. Un error que nadie maneja la cuenta como falla y no corta el resto.
Private Function RunOne(ByVal testName As String) As Boolean
    mCurrentName = testName
    mCurrentFailed = False

    On Error GoTo Fail
    ResetDataSheets
    SetTestMode True
    ClearLastError
    Application.Run "'" & ThisWorkbook.Name & "'!" & testName
    GoTo Done
Fail:
    RecordFailure "error " & Err.Number & ": " & Err.Description
    Resume Done
Done:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = mEventsBefore
    On Error GoTo 0
    RunOne = Not mCurrentFailed
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
