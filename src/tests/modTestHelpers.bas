Attribute VB_Name = "modTestHelpers"
'==============================================================================
' modTestHelpers - Datos de ejemplo y estructura de versiones anteriores
'
' Escribe directo en BD_*: modData ya no sabe producir filas de versiones
' anteriores y este modulo solo existe en el libro de pruebas.
'==============================================================================
Option Explicit

' Los 26 campos llenos; placas y VIN en minusculas; Phone1 = "0123456789"; InTime = 00:00.
Public Function NewSampleOrder(ByVal orderNumber As String) As ServiceOrder
    Dim o As ServiceOrder
    Set o = New ServiceOrder
    o.Number = orderNumber
    o.Status = ORDER_STATUS_DIAGNOSIS
    o.ClientName = "Cliente Prueba"
    o.Phone1 = "0123456789"
    o.Phone2 = "5551234567"
    o.Email = "prueba@example.com"
    o.Address = "Calle Falsa 123"
    o.Make = "Marca X"
    o.Model = "Modelo Y"
    o.Engine = "1.6L"
    o.Year = 2018
    o.Color = "Rojo"
    o.Plates = "abc123d"
    o.Mileage = 85000
    o.FuelLevel = "Lleno"
    o.Vin = "1hgcm82633a004352"
    o.InDate = DateSerial(2026, 3, 15)
    o.InTime = TimeSerial(0, 0, 0)
    o.ReceivedBy = "Recepcion Uno"
    o.PromisedDate = DateSerial(2026, 3, 20)
    o.OutDate = DateSerial(2026, 3, 19)
    o.OutTime = TimeSerial(14, 30, 0)
    o.DeliveredBy = "Entrega Dos"
    o.Requested = "Revision de frenos"
    o.Diagnosis = "Balatas gastadas"
    o.QuoteNumber = "COT-1001"
    Set NewSampleOrder = o
End Function

' Cliente "Cliente Prueba", 16 % y un concepto.
Public Function NewSampleQuote(ByVal quoteNumber As String) As Quote
    Dim q As Quote
    Set q = New Quote
    q.Number = quoteNumber
    q.QuoteDate = DateSerial(2026, 3, 15)
    q.ClientName = "Cliente Prueba"
    q.TaxRate = 0.16
    q.AddItem NewQuoteItem(1, "Servicio de prueba", 100)
    Set NewSampleQuote = q
End Function

' Fila A:H en BD_Cotizaciones (sin columna I) y su fila en BD_Detalle.
Public Sub WriteLegacyQuoteRow(ByVal quoteNumber As String)
    Dim wsQuotes As Worksheet, wsItems As Worksheet, r As Long

    TryUnprotectWorkbook
    Set wsQuotes = ThisWorkbook.Worksheets(SHEET_QUOTES)
    Set wsItems = ThisWorkbook.Worksheets(SHEET_ITEMS)
    UnlockSheet wsQuotes
    UnlockSheet wsItems

    r = wsQuotes.Cells(wsQuotes.Rows.Count, "A").End(xlUp).Row + 1
    wsQuotes.Cells(r, 1).Resize(1, 8).Value = Array(quoteNumber, DateSerial(2026, 3, 15), _
        "Cliente Prueba", 100, 16, 116, 0.16, DateSerial(2026, 4, 14))
    r = wsItems.Cells(wsItems.Rows.Count, "A").End(xlUp).Row + 1
    wsItems.Cells(r, 1).Resize(1, 5).Value = Array(quoteNumber, "Servicio de prueba", 1, 100, 100)
End Sub

' Vacia C:F de BD_Clientes (catalogo sin contacto).
Public Sub ClearClientContacts()
    Dim ws As Worksheet

    TryUnprotectWorkbook
    Set ws = ThisWorkbook.Worksheets(SHEET_CLIENTS)
    UnlockSheet ws
    ws.Range("C2:F" & ws.Rows.Count).ClearContents
End Sub

' Libro v8.0.0: sin BD_Ordenes, sin I1 en BD_Cotizaciones y sin C1:F1 en BD_Clientes.
Public Sub MakeLegacyWorkbook()
    Dim ws As Worksheet, alertsBefore As Boolean

    TryUnprotectWorkbook
    Set ws = ThisWorkbook.Worksheets(SHEET_QUOTES)
    UnlockSheet ws
    ws.Range("I1").ClearContents
    Set ws = ThisWorkbook.Worksheets(SHEET_CLIENTS)
    UnlockSheet ws
    ws.Range("C1:F1").ClearContents

    alertsBefore = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_ORDERS).Delete
    On Error GoTo 0
    Application.DisplayAlerts = alertsBefore
End Sub
