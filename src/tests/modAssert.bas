Attribute VB_Name = "modAssert"
'==============================================================================
' modAssert - Verificaciones para las pruebas (solo existe en el libro de pruebas)
'
' Una verificacion que no se cumple anota la falla en modTestRunner y deja que la
' prueba siga; nunca lanza error.
'==============================================================================
Option Explicit

' Falla si los valores difieren en tipo (los numericos se comparan entre si) o en valor.
Public Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal label As String)
    If Not AreEqual(expected, actual) Then
        modTestRunner.RecordFailure label & ": esperado " & Describe(expected) & ", obtenido " & Describe(actual)
    End If
End Sub

Public Sub AssertTrue(ByVal condition As Boolean, Optional ByVal label As String)
    If Not condition Then modTestRunner.RecordFailure label & ": esperado verdadero, obtenido falso"
End Sub

' Pasa si la macro lanza un error o deja un error registrado por ShowError.
Public Sub AssertFails(ByVal macroName As String, Optional ByVal label As String, _
                       Optional ByVal arg1 As Variant, Optional ByVal arg2 As Variant, Optional ByVal arg3 As Variant)
    Dim fullName As String, raised As Boolean

    fullName = "'" & ThisWorkbook.Name & "'!" & macroName
    ClearLastError

    On Error Resume Next
    If IsMissing(arg1) Then
        Application.Run fullName
    ElseIf IsMissing(arg2) Then
        Application.Run fullName, arg1
    ElseIf IsMissing(arg3) Then
        Application.Run fullName, arg1, arg2
    Else
        Application.Run fullName, arg1, arg2, arg3
    End If
    raised = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0

    If Not (raised Or LastErrorNumber() <> 0) Then
        modTestRunner.RecordFailure label & ": se esperaba una falla y la macro termino sin error"
    End If
    ClearLastError
End Sub

Private Function IsNumberType(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            IsNumberType = True
    End Select
End Function

Private Function AreEqual(ByVal expected As Variant, ByVal actual As Variant) As Boolean
    On Error GoTo Fail
    If IsObject(expected) Or IsObject(actual) Then
        If IsObject(expected) And IsObject(actual) Then AreEqual = (expected Is actual)
    ElseIf IsNumberType(expected) And IsNumberType(actual) Then
        AreEqual = (CDbl(expected) = CDbl(actual))
    ElseIf VarType(expected) <> VarType(actual) Then
        AreEqual = False
    ElseIf IsNull(expected) Or IsEmpty(expected) Then
        AreEqual = True
    Else
        AreEqual = (expected = actual)
    End If
    Exit Function
Fail:
    AreEqual = False
End Function

Private Function Describe(ByVal v As Variant) As String
    On Error GoTo Fail
    If IsObject(v) Then
        If v Is Nothing Then Describe = "Nothing" Else Describe = "<objeto>"
        Exit Function
    End If
    Select Case VarType(v)
        Case vbEmpty: Describe = "<vacio>"
        Case vbNull: Describe = "<nulo>"
        Case vbString: Describe = """" & v & """"
        Case vbDate: Describe = Format$(v, "yyyy-mm-dd hh:nn:ss")
        Case Else: Describe = CStr(v)
    End Select
    Exit Function
Fail:
    Describe = "<" & TypeName(v) & ">"
End Function
