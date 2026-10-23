Attribute VB_Name = "modTestOrderQuote"
'==============================================================================
' modTestOrderQuote - Diagnóstico a conceptos, franja de orden y liga orden-cotización
'==============================================================================
Option Explicit

Private Function OrderWithDiagnosis(ByVal diagnosisText As String) As ServiceOrder
    Dim o As ServiceOrder
    Set o = NewSampleOrder("OS-0001")
    o.Diagnosis = diagnosisText
    Set OrderWithDiagnosis = o
End Function

Private Function NumberedLines(ByVal n As Long) As String
    Dim i As Long, t As String
    For i = 1 To n
        If i > 1 Then t = t & vbLf
        t = t & "R" & i
    Next i
    NumberedLines = t
End Function

Private Function JoinChunks(ByVal c As Collection) As String
    Dim i As Long
    For i = 1 To c.Count
        If i > 1 Then JoinChunks = JoinChunks & vbLf
        JoinChunks = JoinChunks & c(i)
    Next i
End Function

Public Sub Test_TextLineCount_CuentaRenglones()
    AssertEqual 0, TextLineCount(""), "vacio"
    AssertEqual 1, TextLineCount("a"), "uno"
    AssertEqual 2, TextLineCount("a" & vbCrLf & "b"), "crlf"
    AssertEqual 3, TextLineCount("a" & vbLf & vbLf & "b"), "vacio de en medio"
    AssertEqual 1, TextLineCount("a" & vbLf), "salto final"
End Sub

Public Sub Test_DiagnosisChunks_CincoPorFila()
    Dim c As Collection, t As String
    t = NumberedLines(12)
    Set c = DiagnosisChunks(t)
    AssertEqual 3, c.Count, "tres grupos"
    If c.Count <> 3 Then Exit Sub
    AssertEqual 5, TextLineCount(c(1)), "grupo 1"
    AssertEqual 5, TextLineCount(c(2)), "grupo 2"
    AssertEqual 2, TextLineCount(c(3)), "grupo 3"
    AssertEqual t, JoinChunks(c), "unidos dan el texto"
End Sub

Public Sub Test_DiagnosisChunks_UltimaFilaTomaElResto()
    Dim c As Collection, t As String
    t = NumberedLines(45)
    Set c = DiagnosisChunks(t)
    AssertEqual ORD_DIAG_ROWS, c.Count, "ocho grupos"
    If c.Count <> ORD_DIAG_ROWS Then Exit Sub
    AssertEqual 10, TextLineCount(c(ORD_DIAG_ROWS)), "el ultimo toma el resto"
    AssertEqual t, JoinChunks(c), "unidos dan el texto"
End Sub

Public Sub Test_DiagnosisChunks_Vacio()
    AssertEqual 0, DiagnosisChunks("").Count, "sin grupos"
End Sub

Public Sub Test_DiagnosisLines_PartePorSaltos()
    Dim c As Collection
    Set c = DiagnosisLines("  Balatas " & vbCrLf & vbCrLf & "Disco" & vbLf & "Aceite" & vbCr & "Filtro ")
    AssertEqual 4, c.Count, "cuatro renglones"
    If c.Count <> 4 Then Exit Sub
    AssertEqual "Balatas", c(1), "renglon 1"
    AssertEqual "Disco", c(2), "renglon 2"
    AssertEqual "Aceite", c(3), "renglon 3"
    AssertEqual "Filtro", c(4), "renglon 4"
End Sub

Public Sub Test_DiagnosisLines_SinSaltosUnRenglon()
    Dim c As Collection
    Set c = DiagnosisLines("Afinación mayor")
    AssertEqual 1, c.Count, "un renglon"
    If c.Count = 1 Then AssertEqual "Afinación mayor", c(1), "texto"
End Sub

Public Sub Test_DiagnosisLines_VacioSinRenglones()
    AssertEqual 0, DiagnosisLines("").Count, "vacio"
    AssertEqual 0, DiagnosisLines("  " & vbCrLf & "  ").Count, "solo espacios y saltos"
End Sub

Public Sub Test_DiagnosisLines_ConservaVinetas()
    Dim c As Collection
    Set c = DiagnosisLines("- Balatas" & vbLf & "1. Disco")
    AssertEqual 2, c.Count, "dos renglones"
    If c.Count <> 2 Then Exit Sub
    AssertEqual "- Balatas", c(1), "vineta"
    AssertEqual "1. Disco", c(2), "numeracion"
End Sub

Public Sub Test_QuoteFromOrder_Encabezado()
    Dim q As Quote
    Set q = QuoteFromOrder(OrderWithDiagnosis("Balatas" & vbLf & "Disco"), "COT-1005")
    AssertEqual "COT-1005", q.Number, "Number"
    AssertEqual "Cliente Prueba", q.ClientName, "ClientName"
    AssertEqual "OS-0001", q.OrderNumber, "OrderNumber"
    AssertEqual 2, q.ItemCount, "ItemCount"
    If q.ItemCount <> 2 Then Exit Sub
    AssertEqual 1, q.Item(1).Quantity, "cantidad 1"
    AssertEqual 0, q.Item(1).UnitPrice, "precio 1"
    AssertEqual 1, q.Item(2).Quantity, "cantidad 2"
    AssertEqual 0, q.Item(2).UnitPrice, "precio 2"
    AssertEqual "Disco", q.Item(2).Description, "descripcion 2"
    AssertEqual 0, q.Total, "Total"
End Sub

Public Sub Test_QuoteFromOrder_TopeMaxItems()
    Dim q As Quote, text As String, i As Long
    For i = 1 To 20
        text = text & "Renglon " & i & vbLf
    Next i
    Set q = QuoteFromOrder(OrderWithDiagnosis(text), "COT-1005")
    AssertEqual MAX_ITEMS, q.ItemCount, "tope"
    If q.ItemCount = MAX_ITEMS Then AssertEqual "Renglon 17", q.Item(MAX_ITEMS).Description, "ultimo"
End Sub

Public Sub Test_QuoteFromOrder_DescripcionLarga()
    Dim q As Quote
    Set q = QuoteFromOrder(OrderWithDiagnosis(String$(300, "a")), "COT-1005")
    AssertEqual 1, q.ItemCount, "un concepto"
    If q.ItemCount = 1 Then AssertEqual 255, Len(q.Item(1).Description), "largo cortado"
End Sub

Public Sub Test_QuoteFromOrder_NoCambiaDiagnostico()
    Dim o As ServiceOrder, q As Quote
    Set o = OrderWithDiagnosis("  Balatas " & vbLf & "Disco")
    Set q = QuoteFromOrder(o, "COT-1005")
    AssertEqual "  Balatas " & vbLf & "Disco", o.Diagnosis, "Diagnosis"
    AssertEqual "COT-1001", o.QuoteNumber, "QuoteNumber"
End Sub

Public Sub Test_OrderBannerText_Completo()
    AssertEqual "Orden OS-0001 " & ChrW(183) & " Marca X Modelo Y 2018 " & ChrW(183) & " Placas ABC123D", _
        OrderBannerText("OS-0001", NewSampleOrder("OS-0001")), "completo"
End Sub

Public Sub Test_OrderBannerText_OmiteVacios()
    Dim o As ServiceOrder
    Set o = NewSampleOrder("OS-0001")
    o.Model = "": o.Year = 0
    AssertEqual "Orden OS-0001 " & ChrW(183) & " Marca X " & ChrW(183) & " Placas ABC123D", _
        OrderBannerText("OS-0001", o), "sin modelo ni año"
    o.Make = "": o.Plates = ""
    AssertEqual "Orden OS-0001", OrderBannerText("OS-0001", o), "sin vehiculo ni placas"
End Sub

Public Sub Test_OrderBannerText_OrdenNoEncontrada()
    AssertEqual "Orden OS-0009", OrderBannerText("OS-0009", Nothing), "orden Nothing"
    AssertEqual "", OrderBannerText("", NewSampleOrder("OS-0001")), "sin folio"
End Sub

Public Sub Test_CreateQuoteForOrder_LigaAmbosLados()
    Dim o As ServiceOrder, q As Quote
    Set o = NewSampleOrder("OS-0001")
    o.QuoteNumber = ""
    AssertTrue SaveOrder(o), "SaveOrder"
    AssertEqual "COT-1001", CreateQuoteForOrder(o), "folio creado"
    Set q = LoadQuote("COT-1001")
    AssertTrue Not q Is Nothing, "cotizacion guardada"
    If q Is Nothing Then Exit Sub
    AssertEqual "OS-0001", q.OrderNumber, "OrderNumber"
    AssertEqual 1, q.ItemCount, "un concepto"
    AssertEqual "COT-1001", LoadOrder("OS-0001").QuoteNumber, "orden apunta a la cotizacion"
End Sub

Public Sub Test_CreateQuoteForOrder_ReemplazaFolioPerdido()
    Dim o As ServiceOrder
    Set o = NewSampleOrder("OS-0001")
    o.QuoteNumber = "COT-1050"
    AssertTrue SaveOrder(o), "SaveOrder"
    AssertEqual "COT-1001", CreateQuoteForOrder(o), "folio nuevo"
    AssertEqual "COT-1001", LoadOrder("OS-0001").QuoteNumber, "orden reapuntada"
End Sub

Public Sub Test_CreateQuoteForOrder_DiagnosticoVacio()
    Dim o As ServiceOrder, q As Quote
    Set o = OrderWithDiagnosis("")
    o.QuoteNumber = ""
    AssertTrue SaveOrder(o), "SaveOrder"
    AssertEqual "COT-1001", CreateQuoteForOrder(o), "folio creado"
    Set q = LoadQuote("COT-1001")
    AssertTrue Not q Is Nothing, "cotizacion guardada"
    If Not q Is Nothing Then AssertEqual 0, q.ItemCount, "sin conceptos"
End Sub

Public Sub Test_CreateQuoteForOrder_NothingRegresaVacio()
    AssertEqual "", CreateQuoteForOrder(Nothing), "Nothing"
    AssertTrue Not QuoteExists("COT-1001"), "no se creo cotizacion"
End Sub
