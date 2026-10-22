Attribute VB_Name = "modTestHelpers"
'==============================================================================
' modTestHelpers - Datos de ejemplo y estructura de versiones anteriores
'==============================================================================
Option Explicit

' Los 26 campos llenos; placas y VIN en minusculas; Phone1 = "0123456789"; InTime = 00:00.
Public Function NewSampleOrder(ByVal orderNumber As String) As ServiceOrder
    Set NewSampleOrder = New ServiceOrder
End Function

' Cliente "Cliente Prueba", 16 % y un concepto.
Public Function NewSampleQuote(ByVal quoteNumber As String) As Quote
    Set NewSampleQuote = New Quote
End Function

' Fila A:H en BD_Cotizaciones (sin columna I) y su fila en BD_Detalle.
Public Sub WriteLegacyQuoteRow(ByVal quoteNumber As String)
End Sub

' Vacia C:F de BD_Clientes (catalogo sin contacto).
Public Sub ClearClientContacts()
End Sub

' Libro v8.0.0: sin BD_Ordenes, sin I1 en BD_Cotizaciones y sin C1:F1 en BD_Clientes.
Public Sub MakeLegacyWorkbook()
End Sub
