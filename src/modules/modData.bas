Attribute VB_Name = "modData"
'==============================================================================
' modData · Capa de datos (hojas BD_Cotizaciones, BD_Detalle, BD_Clientes y BD_Ordenes)
'
' Es el ÚNICO módulo que lee o escribe las hojas de datos. Si algún día migras
' a SQLite o a la API de Cotiva, reescribes solo este módulo respetando los
' nombres, parámetros y lo que regresa cada función.
'
' Rendimiento: las lecturas traen el rango completo a un arreglo en memoria
' (una sola operación) en lugar de leer celda por celda.
'
' Estructura (fila 1 = encabezados; se conservan en español por compatibilidad):
'   BD_Cotizaciones  A:I  Folio | Fecha | Cliente | Subtotal | IVA | Total | TasaIVA | Vigencia | OrdenFolio
'   BD_Detalle       A:E  Folio | Descripcion | Cantidad | PrecioUnitario | Total
'   BD_Clientes      A:F  Cliente | FechaAlta | Telefono1 | Telefono2 | Correo | Direccion
'   BD_Ordenes       A:Z  Ver ServiceOrder.cls (26 columnas)
'==============================================================================
Option Explicit

Private Const QUOTE_COLUMNS As Long = 9
Private Const ITEM_COLUMNS As Long = 5
Private Const CLIENT_COLUMNS As Long = 6
Private Const ORDER_COLUMNS As Long = 26

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

' "COT-1005" -> 1005 ; prefijo distinto o digitos invalidos -> 0
Private Function FolioValue(ByVal folio As String, ByVal prefix As String) As Long
    Dim digits As String

    folio = UCase$(Trim$(folio))
    If Left$(folio, Len(prefix)) <> prefix Then Exit Function

    digits = Mid$(folio, Len(prefix) + 1)
    If Len(digits) > 0 And Len(digits) <= 9 Then          ' Maximo 9 digitos: cabe en un Long
        If digits Like String(Len(digits), "#") Then FolioValue = CLng(digits)
    End If
End Function

' Siguiente folio libre de una hoja: el numero mas alto de la columna A + 1.
' Con digits > 0 rellena con ceros a ese ancho; con 0 no rellena.
Private Function NextFolio(ByVal ws As Worksheet, ByVal prefix As String, _
        ByVal baseNumber As Long, ByVal digits As Long) As String
    Dim data As Variant, i As Long, n As Long, highest As Long

    highest = baseNumber
    data = ReadData(ws, 2)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            n = FolioValue(CStr(data(i, 1)), prefix)
            If n > highest Then highest = n
        Next i
    End If
    If digits > 0 Then
        NextFolio = prefix & Format$(highest + 1, String$(digits, "0"))
    Else
        NextFolio = prefix & (highest + 1)
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

' Deja los encabezados completos en las hojas de datos.
' Desbloquea las hojas; ProtectAll las vuelve a proteger al final.
Public Sub EnsureDataHeaders()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_QUOTES)
    UnlockSheet ws
    EnsureHeaders ws, Array("Folio", "Fecha", "Cliente", "Subtotal", "IVA", "Total", "TasaIVA", "Vigencia", "OrdenFolio")

    Set ws = ThisWorkbook.Worksheets(SHEET_ITEMS)
    UnlockSheet ws
    EnsureHeaders ws, Array("Folio", "Descripcion", "Cantidad", "PrecioUnitario", "Total")

    If SheetExists(SHEET_CLIENTS) Then
        Set ws = ThisWorkbook.Worksheets(SHEET_CLIENTS)
        UnlockSheet ws
        EnsureHeaders ws, Array("Cliente", "FechaAlta", "Telefono1", "Telefono2", "Correo", "Direccion")
    End If

    If SheetExists(SHEET_ORDERS) Then
        Set ws = ThisWorkbook.Worksheets(SHEET_ORDERS)
        UnlockSheet ws
        EnsureHeaders ws, Array("Folio", "Estado", "Cliente", "Telefono1", "Telefono2", "Correo", _
            "Direccion", "Marca", "Modelo", "Motor", "Anio", "Color", "Placas", "Kilometraje", _
            "Combustible", "VIN", "FechaEntrada", "HoraEntrada", "RecibioPor", "FechaPromesa", _
            "FechaSalida", "HoraSalida", "EntregoPor", "Solicitado", "Diagnostico", "CotizacionFolio")
    End If
End Sub

'--- Cotizaciones --------------------------------------------------------------

' Siguiente folio libre: el número más alto guardado + 1 (no depende del orden).
Public Function GetNextQuoteNumber() As String
    GetNextQuoteNumber = NextFolio(ThisWorkbook.Worksheets(SHEET_QUOTES), QUOTE_PREFIX, _
        QUOTE_BASE_NUMBER, QUOTE_NUMBER_DIGITS)
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
        q.TaxRate, q.ValidUntil(ValidityDays()), UCase$(Trim$(q.OrderNumber)))
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
    q.OrderNumber = Trim$(CStr(ws.Cells(r, 9).Value))

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

' Escribe (sin pisar con vacios) el contacto del cliente en el catalogo;
' lo da de alta si no existe.
Private Sub UpsertClientContact(ByVal clientText As String, ByVal phone1 As String, _
        ByVal phone2 As String, ByVal email As String, ByVal address As String)
    Dim ws As Worksheet, data As Variant, i As Long, r As Long, values As Variant, c As Long

    clientText = Trim$(clientText)
    If Len(clientText) = 0 Then Exit Sub
    AddClient clientText

    Set ws = ThisWorkbook.Worksheets(SHEET_CLIENTS)
    data = ReadData(ws, CLIENT_COLUMNS)
    If IsEmpty(data) Then Exit Sub
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(CStr(data(i, 1))), clientText, vbTextCompare) = 0 Then
            r = i + 1
            Exit For
        End If
    Next i
    If r = 0 Then Exit Sub

    values = Array(phone1, phone2, email, address)
    ws.Range(ws.Cells(r, 3), ws.Cells(r, 4)).NumberFormat = "@"
    For c = 0 To 3
        If Len(Trim$(CStr(values(c)))) > 0 Then
            ' Telefonos (formato texto): sin apostrofo, un "+52..." se guarda tal cual
            If c < 2 Then
                ws.Cells(r, 3 + c).Value = Trim$(CStr(values(c)))
            Else
                ws.Cells(r, 3 + c).Value = SafeCellText(Trim$(CStr(values(c))))
            End If
        End If
    Next c
End Sub

' Contacto guardado del cliente; si el catalogo no tiene datos, el de su orden
' mas reciente. False si no hay ningun dato.
Public Function GetClientContact(ByVal clientText As String, ByRef phone1 As String, _
        ByRef phone2 As String, ByRef email As String, ByRef address As String) As Boolean
    Dim data As Variant, i As Long, best As Long, bestDate As Date, d As Date

    phone1 = vbNullString: phone2 = vbNullString: email = vbNullString: address = vbNullString
    clientText = Trim$(clientText)
    If Len(clientText) = 0 Then Exit Function

    data = ReadData(ThisWorkbook.Worksheets(SHEET_CLIENTS), CLIENT_COLUMNS)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If StrComp(Trim$(CStr(data(i, 1))), clientText, vbTextCompare) = 0 Then
                phone1 = CStr(data(i, 3)): phone2 = CStr(data(i, 4))
                email = CStr(data(i, 5)): address = CStr(data(i, 6))
                Exit For
            End If
        Next i
    End If
    If Len(phone1 & phone2 & email & address) > 0 Then
        GetClientContact = True
        Exit Function
    End If

    data = ReadData(ThisWorkbook.Worksheets(SHEET_ORDERS), ORDER_COLUMNS)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(CStr(data(i, 3))), clientText, vbTextCompare) = 0 Then
            d = CellDate(data(i, 17))
            If best = 0 Or d >= bestDate Then
                best = i
                bestDate = d
            End If
        End If
    Next i
    If best = 0 Then Exit Function

    phone1 = CStr(data(best, 4)): phone2 = CStr(data(best, 5))
    email = CStr(data(best, 6)): address = CStr(data(best, 7))
    GetClientContact = (Len(phone1 & phone2 & email & address) > 0)
End Function

'--- Ordenes de servicio -------------------------------------------------------

' Siguiente folio de orden libre (OS-0001, OS-0002...), con el mismo criterio que
' GetNextQuoteNumber pero con ORDER_PREFIX y relleno de ORDER_NUMBER_DIGITS.
Public Function GetNextOrderNumber() As String
    GetNextOrderNumber = NextFolio(ThisWorkbook.Worksheets(SHEET_ORDERS), ORDER_PREFIX, _
        ORDER_BASE_NUMBER, ORDER_NUMBER_DIGITS)
End Function

' Fecha de una celda; 0 si no es fecha.
Private Function CellDate(ByVal v As Variant) As Date
    If IsDate(v) Then
        CellDate = CDate(v)
    ElseIf VarType(v) = vbDouble Then
        ' Excel entrega como Double una hora sola (0.6041 = 14:30)
        If v >= 0 And v < 2958466 Then CellDate = CDate(v)
    End If
End Function

' Fecha para escribir en la hoja: 0 ("sin valor") se guarda como celda vacia.
Private Function DateOrBlank(ByVal d As Date) As Variant
    If d = 0 Then DateOrBlank = Empty Else DateOrBlank = d
End Function

' Hora para escribir en la hoja: con fecha asociada, 0 es medianoche valida (00:00);
' sin fecha ni hora la celda queda vacia.
Private Function TimeOrBlank(ByVal t As Date, ByVal d As Date) As Variant
    If t = 0 And d = 0 Then TimeOrBlank = Empty Else TimeOrBlank = t
End Function

' Fila donde esta el folio en BD_Ordenes, o 0 si no existe.
Private Function FindOrderRow(ByVal orderNumber As String) As Long
    Dim data As Variant, i As Long

    orderNumber = Trim$(orderNumber)
    data = ReadData(ThisWorkbook.Worksheets(SHEET_ORDERS), 2)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(CStr(data(i, 1))), orderNumber, vbTextCompare) = 0 Then
            FindOrderRow = i + 1
            Exit Function
        End If
    Next i
End Function

' True si el folio de la orden ya esta en BD_Ordenes (sin distinguir mayusculas).
Public Function OrderExists(ByVal orderNumber As String) As Boolean
    OrderExists = (FindOrderRow(orderNumber) > 0)
End Function

' Guarda (o reemplaza, si ya existia) una orden y pone su contacto en el catalogo
' de clientes. Regresa True si todo salio bien.
Public Function SaveOrder(ByVal o As ServiceOrder) As Boolean
    Dim ws As Worksheet, r As Long, label As String

    On Error GoTo Fail
    Application.ScreenUpdating = False

    If Len(Trim$(o.Number)) = 0 Then o.Number = GetNextOrderNumber()
    o.Number = UCase$(Trim$(o.Number))
    label = o.Number
    Set ws = ThisWorkbook.Worksheets(SHEET_ORDERS)

    ' Se escribe sobre la fila existente: si la escritura falla no se pierde la orden anterior
    r = FindOrderRow(o.Number)
    If r = 0 Then r = LastRow(ws) + 1

    ' Telefonos, placas y VIN como texto: conservan el 0 inicial y no se interpretan
    ws.Range(ws.Cells(r, 4), ws.Cells(r, 5)).NumberFormat = "@"
    ws.Cells(r, 13).NumberFormat = "@"
    ws.Cells(r, 16).NumberFormat = "@"
    ws.Cells(r, 1).Resize(1, ORDER_COLUMNS).Value = Array( _
        o.Number, SafeCellText(o.Status), SafeCellText(o.ClientName), Trim$(o.Phone1), _
        Trim$(o.Phone2), SafeCellText(o.Email), SafeCellText(o.Address), _
        SafeCellText(o.Make), SafeCellText(o.Model), SafeCellText(o.Engine), o.Year, _
        SafeCellText(o.Color), UCase$(Trim$(o.Plates)), o.Mileage, _
        SafeCellText(o.FuelLevel), UCase$(Trim$(o.Vin)), _
        DateOrBlank(o.InDate), TimeOrBlank(o.InTime, o.InDate), SafeCellText(o.ReceivedBy), _
        DateOrBlank(o.PromisedDate), DateOrBlank(o.OutDate), TimeOrBlank(o.OutTime, o.OutDate), _
        SafeCellText(o.DeliveredBy), SafeCellText(o.Requested), SafeCellText(o.Diagnosis), _
        SafeCellText(UCase$(Trim$(o.QuoteNumber))))
    ws.Cells(r, 17).NumberFormat = "dd/mm/yyyy"
    ws.Cells(r, 20).Resize(1, 2).NumberFormat = "dd/mm/yyyy"
    ws.Cells(r, 18).NumberFormat = "hh:mm"
    ws.Cells(r, 22).NumberFormat = "hh:mm"

    UpsertClientContact o.ClientName, o.Phone1, o.Phone2, o.Email, o.Address
    If SAVE_WORKBOOK_ON_SAVE And Len(ThisWorkbook.Path) > 0 Then ThisWorkbook.Save

    SaveOrder = True
Cleanup:
    Application.ScreenUpdating = True
    Exit Function
Fail:
    ShowError "SaveOrder " & label
    Resume Cleanup
End Function

' Carga una orden guardada. Regresa Nothing si el folio no existe.
' Los datos leidos se validan: alguien pudo editar la hoja a mano.
Public Function LoadOrder(ByVal orderNumber As String) As ServiceOrder
    Dim ws As Worksheet, r As Long, v As Variant, o As ServiceOrder, n As Double

    r = FindOrderRow(orderNumber)
    If r = 0 Then Exit Function

    Set ws = ThisWorkbook.Worksheets(SHEET_ORDERS)
    v = ws.Range(ws.Cells(r, 1), ws.Cells(r, ORDER_COLUMNS)).Value
    Set o = New ServiceOrder

    o.Number = Trim$(CStr(v(1, 1)))
    If Len(Trim$(CStr(v(1, 2)))) > 0 Then o.Status = Trim$(CStr(v(1, 2))) Else o.Status = ORDER_STATUS_RECEIVED
    o.ClientName = CStr(v(1, 3))
    o.Phone1 = CStr(v(1, 4))
    o.Phone2 = CStr(v(1, 5))
    o.Email = CStr(v(1, 6))
    o.Address = CStr(v(1, 7))
    o.Make = CStr(v(1, 8))
    o.Model = CStr(v(1, 9))
    o.Engine = CStr(v(1, 10))
    If TryParseNumber(CStr(v(1, 11)), n) Then
        If n >= 0 And n <= 9999 Then o.Year = CInt(n)
    End If
    o.Color = CStr(v(1, 12))
    o.Plates = CStr(v(1, 13))
    If TryParseNumber(CStr(v(1, 14)), n) Then
        If n >= 0 And n <= 2147483647# Then o.Mileage = CLng(n)
    End If
    o.FuelLevel = CStr(v(1, 15))
    o.Vin = CStr(v(1, 16))
    o.InDate = CellDate(v(1, 17))
    o.InTime = CellDate(v(1, 18))
    o.ReceivedBy = CStr(v(1, 19))
    o.PromisedDate = CellDate(v(1, 20))
    o.OutDate = CellDate(v(1, 21))
    o.OutTime = CellDate(v(1, 22))
    o.DeliveredBy = CStr(v(1, 23))
    o.Requested = CStr(v(1, 24))
    o.Diagnosis = CStr(v(1, 25))
    o.QuoteNumber = CStr(v(1, 26))

    Set LoadOrder = o
End Function

' Tabla (0..n-1, 0..3): Folio, Cliente, Placas, Estado, de la mas reciente a la
' mas antigua; Empty si no hay ordenes.
Public Function ListOrders() As Variant
    Dim data As Variant, i As Long, totalRows As Long, n As Long, result() As String

    data = ReadData(ThisWorkbook.Worksheets(SHEET_ORDERS), ORDER_COLUMNS)
    If IsEmpty(data) Then
        ListOrders = Empty
        Exit Function
    End If

    For i = 1 To UBound(data, 1)
        If Len(Trim$(CStr(data(i, 1)))) > 0 Then totalRows = totalRows + 1
    Next i
    If totalRows = 0 Then
        ListOrders = Empty
        Exit Function
    End If

    ReDim result(0 To totalRows - 1, 0 To 3)
    For i = UBound(data, 1) To 1 Step -1
        If Len(Trim$(CStr(data(i, 1)))) > 0 Then
            result(n, 0) = CStr(data(i, 1))
            result(n, 1) = CStr(data(i, 3))
            result(n, 2) = CStr(data(i, 13))
            result(n, 3) = CStr(data(i, 2))
            n = n + 1
        End If
    Next i
    ListOrders = result
End Function

' Nombres distintos de Recibio/Entrego en BD_Ordenes, orden alfabetico; Empty si no hay.
Public Function ListOrderStaff() As Variant
    Dim data As Variant, i As Long, j As Long, c As Long, n As Long, found As Boolean
    Dim names() As String, name As String, tmp As String

    data = ReadData(ThisWorkbook.Worksheets(SHEET_ORDERS), ORDER_COLUMNS)
    If IsEmpty(data) Then Exit Function

    ReDim names(0 To UBound(data, 1) * 2)
    For i = 1 To UBound(data, 1)
        For c = 19 To 23 Step 4                   ' Recibio y Entrego
            name = Trim$(CStr(data(i, c)))
            If Len(name) > 0 Then
                found = False
                For j = 0 To n - 1
                    If StrComp(names(j), name, vbTextCompare) = 0 Then found = True: Exit For
                Next j
                If Not found Then
                    names(n) = name
                    n = n + 1
                End If
            End If
        Next c
    Next i
    If n = 0 Then Exit Function

    ReDim Preserve names(0 To n - 1)
    For i = 0 To n - 2                            ' Ordenar (listas cortas)
        For j = i + 1 To n - 1
            If StrComp(names(j), names(i), vbTextCompare) < 0 Then
                tmp = names(i): names(i) = names(j): names(j) = tmp
            End If
        Next j
    Next i
    ListOrderStaff = names
End Function

'--- Orden a cotización ---------------------------------------------------------

' Crea, guarda y liga una cotización nueva a la orden (que ya debe estar guardada).
' Regresa el folio, o "" si algo falló.
Public Function CreateQuoteForOrder(ByVal o As ServiceOrder) As String
    Dim q As Quote
    Dim previous As String

    CreateQuoteForOrder = ""
    If o Is Nothing Then Exit Function

    Set q = QuoteFromOrder(o, GetNextQuoteNumber())
    If Not SaveQuote(q) Then
        DeleteQuote q.Number
        Exit Function
    End If

    previous = o.QuoteNumber
    o.QuoteNumber = q.Number
    If Not SaveOrder(o) Then
        o.QuoteNumber = previous
        DeleteQuote q.Number
        Exit Function
    End If
    CreateQuoteForOrder = q.Number
End Function

' Deshace una cotizacion a medio crear (encabezado y conceptos).
Private Sub DeleteQuote(ByVal quoteNumber As String)
    On Error Resume Next
    DeleteRowsByNumber ThisWorkbook.Worksheets(SHEET_QUOTES), quoteNumber
    DeleteRowsByNumber ThisWorkbook.Worksheets(SHEET_ITEMS), quoteNumber
End Sub

'--- Clientes con datos fiscales (contrato; el cuerpo lo completan T2 y T3) ------

Public Sub EnsureSatCatalogs()
End Sub

Public Function ListTaxRegimes() As Variant
    ListTaxRegimes = Empty
End Function

Public Function ListCfdiUses() As Variant
    ListCfdiUses = Empty
End Function

Public Function LoadClient(ByVal clientText As String) As Client
    Set LoadClient = Nothing
End Function

Public Function SaveClient(ByVal c As Client) As Boolean
    SaveClient = False
End Function

Public Function FindClientByRfc(ByVal rfc As String, ByVal exceptClient As String) As String
    FindClientByRfc = vbNullString
End Function

Public Function SearchClients(ByVal searchText As String) As Variant
    SearchClients = Empty
End Function
