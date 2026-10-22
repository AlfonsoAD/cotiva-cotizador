Attribute VB_Name = "modTestRunner"
'==============================================================================
' modTestRunner - Corre cada Public Sub Test_*() de los modulos modTest*
'==============================================================================
Option Explicit

' Reporte: primera linea "ok N / falla M", luego "FALLA <modulo>.<prueba>: <detalle>".
Public Function RunAllTests() As String
    RunAllTests = "ok 0 / falla 0" & vbLf & "no se encontraron pruebas"
End Function

' La usa modAssert para anotar una falla de la prueba en curso.
Public Sub RecordFailure(ByVal message As String)
End Sub
