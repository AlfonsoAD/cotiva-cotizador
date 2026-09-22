Attribute VB_Name = "modData"
'==============================================================================
' modData · Capa de datos (hojas BD_Cotizaciones, BD_Detalle y BD_Clientes)
'
' Es el ÚNICO módulo que lee o escribe las hojas de datos. Si algún día migras
' a SQLite o a la API de Cotiva, reescribes solo este módulo respetando los
' nombres, parámetros y lo que regresa cada función.
'
' Rendimiento: las lecturas traen el rango completo a un arreglo en memoria
' (una sola operación) en lugar de leer celda por celda.
'
' Estructura (fila 1 = encabezados; se conservan en español por compatibilidad):
'   BD_Cotizaciones  A:H  Folio | Fecha | Cliente | Subtotal | IVA | Total | TasaIVA | Vigencia
'   BD_Detalle       A:E  Folio | Descripcion | Cantidad | PrecioUnitario | Total
'   BD_Clientes      A:B  Cliente | FechaAlta
'==============================================================================
Option Explicit

Private Const QUOTE_COLUMNS As Long = 8
Private Const ITEM_COLUMNS As Long = 5
Private Const CLIENT_COLUMNS As Long = 2

'--- Utilidades internas -------------------------------------------------------

' Última fila con datos en la columna A.
Private Function LastRow(ByVal ws As Worksheet) As Long
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
End Function

' Lee las filas de datos (desde la fila 2) a un arreglo 2D que empieza en 1.
' Regresa Empty si no hay datos. columnCount debe ser >= 2 para que Excel
' siempre regrese un arreglo (con una sola celda regresaría un valor suelto).
Private Function ReadData(ByVal ws As Worksheet, ByVal columnCount As Long) As Variant
    Dim last As Long
    last = LastRow(ws)
    If last < 2 Then
        ReadData = Empty
    Else
        ReadData = ws.Range(ws.Cells(2, 1), ws.Cells(last, columnCount)).Value
    End If
End Function

' Borra de una hoja todas las filas cuyo folio (columna A) coincida.
' Junta las filas con Union y las borra en una sola operación.
Private Sub DeleteRowsByNumber(ByVal ws As Worksheet, ByVal quoteNumber As String)
    Dim data As Variant, i As Long, target As Range

    data = ReadData(ws, 2)
    If IsEmpty(data) Then Exit Sub

    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, 1)), quoteNumber, vbTextCompare) = 0 Then
            If target Is Nothing Then
                Set target = ws.Rows(i + 1)
            Else
                Set target = Union(target, ws.Rows(i + 1))
            End If
        End If
    Next i
    If Not target Is Nothing Then target.Delete
End Sub

' "COT-1005" -> 1005 ; formato inválido -> 0
Private Function QuoteNumberValue(ByVal quoteNumber As String) As Long
    Dim digits As String

    quoteNumber = UCase$(Trim$(quoteNumber))
    If Left$(quoteNumber, Len(QUOTE_PREFIX)) <> QUOTE_PREFIX Then Exit Function

    digits = Mid$(quoteNumber, Len(QUOTE_PREFIX) + 1)
    If Len(digits) > 0 And Len(digits) <= 9 Then          ' Máximo 9 dígitos: cabe en un Long
        If digits Like String(Len(digits), "#") Then QuoteNumberValue = CLng(digits)
    End If
End Function

' Escribe en la fila 1 los encabezados que falten (sirve también para
' actualizar hojas creadas con versiones anteriores).
Private Sub EnsureHeaders(ByVal ws As Worksheet, ByVal headers As Variant)
    Dim i As Long
    For i = 0 To UBound(headers)
        If Len(Trim$(CStr(ws.Cells(1, i + 1).Value))) = 0 Then
            ws.Cells(1, i + 1).Value = headers(i)
            ws.Cells(1, i + 1).Font.Bold = True
        End If
    Next i
End Sub

'--- Estructura ----------------------------------------------------------------

' Deja los encabezados completos en las 3 hojas de datos.
' Desbloquea las hojas; ProtectAll las vuelve a proteger al final.
Public Sub EnsureDataHeaders()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_QUOTES)
    UnlockSheet ws
    EnsureHeaders ws, Array("Folio", "Fecha", "Cliente", "Subtotal", "IVA", "Total", "TasaIVA", "Vigencia")

    Set ws = ThisWorkbook.Worksheets(SHEET_ITEMS)
    UnlockSheet ws
    EnsureHeaders ws, Array("Folio", "Descripcion", "Cantidad", "PrecioUnitario", "Total")

    If SheetExists(SHEET_CLIENTS) Then
        Set ws = ThisWorkbook.Worksheets(SHEET_CLIENTS)
        UnlockSheet ws
        EnsureHeaders ws, Array("Cliente", "FechaAlta")
    End If
End Sub

'--- Cotizaciones --------------------------------------------------------------

' Siguiente folio libre: el número más alto guardado + 1 (no depende del orden).
Public Function GetNextQuoteNumber() As String
    Dim data As Variant, i As Long, n As Long, highest As Long

    highest = QUOTE_BASE_NUMBER
    data = ReadData(ThisWorkbook.Worksheets(SHEET_QUOTES), 2)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            n = QuoteNumberValue(CStr(data(i, 1)))
            If n > highest Then highest = n
        Next i
    End If
    GetNextQuoteNumber = QUOTE_PREFIX & (highest + 1)
End Function

' Fila donde está el folio en BD_Cotizaciones, o 0 si no existe.
Private Function FindQuoteRow(ByVal quoteNumber As String) As Long
    Dim data As Variant, i As Long

    data = ReadData(ThisWorkbook.Worksheets(SHEET_QUOTES), 2)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, 1)), quoteNumber, vbTextCompare) = 0 Then
            FindQuoteRow = i + 1
            Exit Function
        End If
    Next i
End Function

Public Function QuoteExists(ByVal quoteNumber As String) As Boolean
    QuoteExists = (FindQuoteRow(quoteNumber) > 0)
End Function

' Guarda (o reemplaza, si ya existía) una cotización completa.
' Regresa True si todo salió bien.
Public Function SaveQuote(ByVal q As Quote) As Boolean
    Dim wsQuotes As Worksheet, wsItems As Worksheet
    Dim r As Long, i As Long, itemRows() As Variant

    On Error GoTo Fail
    Application.ScreenUpdating = False

    Set wsQuotes = ThisWorkbook.Worksheets(SHEET_QUOTES)
    Set wsItems = ThisWorkbook.Worksheets(SHEET_ITEMS)

    ' 1) Si el folio ya existía, se borra su versión anterior
    DeleteRowsByNumber wsQuotes, q.Number
    DeleteRowsByNumber wsItems, q.Number

    ' 2) Encabezado en la primera fila libre
    r = LastRow(wsQuotes) + 1
    wsQuotes.Cells(r, 1).Resize(1, QUOTE_COLUMNS).Value = Array( _
        q.Number, q.QuoteDate, SafeCellText(q.ClientName), q.Subtotal, q.Tax, q.Total, _
        q.TaxRate, q.ValidUntil(ValidityDays()))
    wsQuotes.Cells(r, 2).NumberFormat = "dd/mm/yyyy"
    wsQuotes.Cells(r, 4).Resize(1, 3).NumberFormat = "$#,##0.00"
    wsQuotes.Cells(r, 8).NumberFormat = "dd/mm/yyyy"

    ' 3) Conceptos: se arma un arreglo y se escribe en bloque (una sola operación)
    If q.ItemCount > 0 Then
        ReDim itemRows(1 To q.ItemCount, 1 To ITEM_COLUMNS)
        For i = 1 To q.ItemCount
            itemRows(i, 1) = q.Number
            itemRows(i, 2) = SafeCellText(q.Item(i).Description)
            itemRows(i, 3) = q.Item(i).Quantity
            itemRows(i, 4) = q.Item(i).UnitPrice
            itemRows(i, 5) = q.Item(i).Amount
        Next i
        r = LastRow(wsItems) + 1
        wsItems.Cells(r, 1).Resize(q.ItemCount, ITEM_COLUMNS).Value = itemRows
        wsItems.Cells(r, 4).Resize(q.ItemCount, 2).NumberFormat = "$#,##0.00"
    End If

    ' 4) Cliente al catálogo (si es nuevo) y guardado del archivo
    AddClient q.ClientName
    If SAVE_WORKBOOK_ON_SAVE And Len(ThisWorkbook.Path) > 0 Then ThisWorkbook.Save

    SaveQuote = True
Cleanup:
    Application.ScreenUpdating = True
    Exit Function
Fail:
    ShowError "SaveQuote " & q.Number
    Resume Cleanup
End Function

' Carga una cotización guardada. Regresa Nothing si el folio no existe.
' Los datos leídos se validan: alguien pudo editar la hoja a mano.
Public Function LoadQuote(ByVal quoteNumber As String) As Quote
    Dim ws As Worksheet, r As Long, q As Quote
    Dim cellValue As Variant, rate As Double
    Dim data As Variant, i As Long, qty As Double, price As Double

    r = FindQuoteRow(quoteNumber)
    If r = 0 Then Exit Function

    Set ws = ThisWorkbook.Worksheets(SHEET_QUOTES)
    Set q = New Quote
    q.Number = CStr(ws.Cells(r, 1).Value)

    cellValue = ws.Cells(r, 2).Value
    If IsDate(cellValue) Then q.QuoteDate = CDate(cellValue) Else q.QuoteDate = Date

    q.ClientName = CStr(ws.Cells(r, 3).Value)

    ' Tasa: vacía = cotización de una versión anterior (16 %). Un 0 es "Sin IVA".
    If TryParseNumber(CStr(ws.Cells(r, 7).Value), rate) Then
        If rate >= 0 And rate < 1 Then q.TaxRate = rate Else q.TaxRate = 0.16
    Else
        q.TaxRate = 0.16
    End If

    ' Conceptos
    data = ReadData(ThisWorkbook.Worksheets(SHEET_ITEMS), ITEM_COLUMNS)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If StrComp(CStr(data(i, 1)), q.Number, vbTextCompare) = 0 Then
                If Not TryParseNumber(CStr(data(i, 3)), qty) Then qty = 0
                If Not TryParseNumber(CStr(data(i, 4)), price) Then price = 0
                q.AddItem NewQuoteItem(qty, CStr(data(i, 2)), price)
            End If
        Next i
    End If

    Set LoadQuote = q
End Function

' Tabla (Folio, Cliente) de la más reciente a la más antigua, o Empty si no hay.
Public Function ListQuotes() As Variant
    Dim data As Variant, i As Long, totalRows As Long, n As Long, result() As String

    data = ReadData(ThisWorkbook.Worksheets(SHEET_QUOTES), 3)
    If IsEmpty(data) Then
        ListQuotes = Empty
        Exit Function
    End If

    For i = 1 To UBound(data, 1)
        If Len(Trim$(CStr(data(i, 1)))) > 0 Then totalRows = totalRows + 1
    Next i
    If totalRows = 0 Then
        ListQuotes = Empty
        Exit Function
    End If

    ReDim result(0 To totalRows - 1, 0 To 1)
    For i = UBound(data, 1) To 1 Step -1                 ' De abajo hacia arriba = más reciente primero
        If Len(Trim$(CStr(data(i, 1)))) > 0 Then
            result(n, 0) = CStr(data(i, 1))
            result(n, 1) = CStr(data(i, 3))
            n = n + 1
        End If
    Next i
    ListQuotes = result
End Function

'--- Clientes ------------------------------------------------------------------

Public Function ClientExists(ByVal clientText As String) As Boolean
    Dim data As Variant, i As Long

    data = ReadData(ThisWorkbook.Worksheets(SHEET_CLIENTS), CLIENT_COLUMNS)
    If IsEmpty(data) Then Exit Function
    clientText = Trim$(clientText)
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(CStr(data(i, 1))), clientText, vbTextCompare) = 0 Then
            ClientExists = True
            Exit Function
        End If
    Next i
End Function

' Agrega el cliente al catálogo si todavía no existe.
Public Sub AddClient(ByVal clientText As String)
    Dim ws As Worksheet, r As Long

    clientText = Trim$(clientText)
    If Len(clientText) = 0 Then Exit Sub
    If ClientExists(clientText) Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SHEET_CLIENTS)
    r = LastRow(ws) + 1
    ws.Cells(r, 1).Value = SafeCellText(clientText)
    ws.Cells(r, 2).Value = Date
    ws.Cells(r, 2).NumberFormat = "dd/mm/yyyy"
End Sub

' Catálogo ordenado alfabéticamente (arreglo), o Empty si no hay clientes.
Public Function ListClients() As Variant
    Dim data As Variant, i As Long, j As Long, n As Long
    Dim names() As String, current As String

    data = ReadData(ThisWorkbook.Worksheets(SHEET_CLIENTS), CLIENT_COLUMNS)
    If IsEmpty(data) Then
        ListClients = Empty
        Exit Function
    End If

    ReDim names(0 To UBound(data, 1) - 1)
    For i = 1 To UBound(data, 1)
        current = Trim$(CStr(data(i, 1)))
        If Len(current) > 0 Then
            ' Ordenamiento por inserción: recorremos a la derecha los mayores
            j = n - 1
            Do While j >= 0
                If StrComp(names(j), current, vbTextCompare) <= 0 Then Exit Do
                names(j + 1) = names(j)
                j = j - 1
            Loop
            names(j + 1) = current
            n = n + 1
        End If
    Next i

    If n = 0 Then
        ListClients = Empty
    Else
        ReDim Preserve names(0 To n - 1)
        ListClients = names
    End If
End Function

' Llena el catálogo con los clientes de las cotizaciones ya guardadas
' (migración cuando BD_Clientes se crea en un libro con datos).
Public Sub SeedClientsFromQuotes()
    Dim data As Variant, i As Long

    data = ReadData(ThisWorkbook.Worksheets(SHEET_QUOTES), 3)
    If IsEmpty(data) Then Exit Sub
    For i = 1 To UBound(data, 1)
        AddClient CStr(data(i, 3))
    Next i
End Sub
