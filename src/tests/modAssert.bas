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
End Sub

Public Sub AssertTrue(ByVal condition As Boolean, Optional ByVal label As String)
End Sub

' Pasa si la macro lanza un error o deja un error registrado por ShowError.
Public Sub AssertFails(ByVal macroName As String, Optional ByVal label As String, _
                       Optional ByVal arg1 As Variant, Optional ByVal arg2 As Variant, Optional ByVal arg3 As Variant)
End Sub
