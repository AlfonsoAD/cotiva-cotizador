Attribute VB_Name = "modTestIntegration"
'==============================================================================
' modTestIntegration - Flujos completos de la orden de servicio encadenados
'
' Cada prueba une pasos que otros modulos prueban sueltos. "Una sola hoja" se
' comprueba por configuracion de PageSetup, no contando paginas del PDF.
'==============================================================================
Option Explicit

Private Sub AssertPageFit(ByVal ws As Worksheet, ByVal oneTall As Boolean)
    With ws.PageSetup
        AssertTrue .Zoom = False, ws.Name & ": Zoom = False"
        AssertEqual 1, CLng(.FitToPagesWide), ws.Name & ": una pagina de ancho"
        If oneTall Then
            AssertEqual 1, CLng(.FitToPagesTall), ws.Name & ": una pagina de alto"
        Else
            AssertTrue .FitToPagesTall = False, ws.Name & ": alto libre"
        End If
        AssertTrue Len(.PrintArea) > 0, ws.Name & ": area de impresion"
    End With
End Sub

' Texto de las filas del diagnostico no vacias, unido con vbLf.
Private Function DiagnosisFromSheet(ByVal ws As Worksheet) As String
    Dim r As Long, v As String
    For r = ORD_DIAG_FIRST_ROW To ORD_DIAG_FIRST_ROW + ORD_DIAG_ROWS - 1
        v = Replace(CStr(ws.Range("A" & r).Value), vbCrLf, vbLf)
        If Len(v) > 0 Then
            If Len(DiagnosisFromSheet) > 0 Then DiagnosisFromSheet = DiagnosisFromSheet & vbLf
            DiagnosisFromSheet = DiagnosisFromSheet & v
        End If
    Next r
End Function

Private Function ShortLines(ByVal n As Long) As String
    Dim i As Long
    For i = 1 To n
        If i > 1 Then ShortLines = ShortLines & vbLf
        ShortLines = ShortLines & "Renglon " & i
    Next i
End Function

Public Sub Test_Integracion_OrdenNuevaGuardarCargar()
    Dim folio As String, o As ServiceOrder, l As ServiceOrder
    folio = GetNextOrderNumber()
    Set o = NewSampleOrder(folio)
    AssertTrue SaveOrder(o), "SaveOrder"
    AssertTrue OrderExists(folio), "OrderExists"
    AssertTrue GetNextOrderNumber() <> folio, "folio nuevo sugerido"
    Set l = LoadOrder(folio)
    AssertTrue Not l Is Nothing, "LoadOrder"
    If l Is Nothing Then Exit Sub
    AssertEqual folio, l.Number, "folio"
    AssertEqual o.ClientName, l.ClientName, "cliente"
    AssertEqual o.Make, l.Make, "marca"
    AssertEqual o.Model, l.Model, "modelo"
    AssertEqual o.Year, l.Year, "anio"
    AssertEqual UCase$(o.Plates), UCase$(l.Plates), "placas"
    AssertEqual o.Diagnosis, l.Diagnosis, "diagnostico"
End Sub

Public Sub Test_Integracion_ClienteQueRegresa()
    Dim p1 As String, p2 As String, em As String, ad As String
    Dim o As ServiceOrder
    Set o = NewSampleOrder("OS-0001")
    o.QuoteNumber = ""
    AssertTrue SaveOrder(o), "SaveOrder"
    AssertTrue GetClientContact("Cliente Prueba", p1, p2, em, ad), "contacto"
    AssertEqual o.Phone1, p1, "telefono 1"
    AssertEqual o.Phone2, p2, "telefono 2"
    AssertEqual o.Email, em, "correo"
    AssertEqual o.Address, ad, "direccion"
End Sub

Public Sub Test_Integracion_OrdenACotizacionYPlantilla()
    Dim o As ServiceOrder, q As Quote, folio As String
    Set o = NewSampleOrder("OS-0001")
    o.QuoteNumber = ""
    AssertTrue SaveOrder(o), "SaveOrder"
    folio = CreateQuoteForOrder(o)
    AssertTrue Len(folio) > 0, "folio de cotizacion"
    Set q = LoadQuote(folio)
    Set o = LoadOrder("OS-0001")
    AssertTrue Not q Is Nothing And Not o Is Nothing, "ambas cargan"
    If q Is Nothing Or o Is Nothing Then Exit Sub
    AssertEqual "OS-0001", q.OrderNumber, "cotizacion apunta a la orden"
    AssertEqual folio, o.QuoteNumber, "orden apunta a la cotizacion"
    FillTemplate q
    AssertEqual 0, LastErrorNumber(), "sin error"
    AssertTrue InStr(CStr(ThisWorkbook.Worksheets(SHEET_TEMPLATE).Range(CELL_VEHICLE).Value), o.Make) > 0, "vehiculo en la Plantilla"
    AssertTrue Len(OrderBannerText("OS-0001", o)) > 0, "franja de la orden"
End Sub

Public Sub Test_Integracion_CotizacionViejaSinOrden()
    Dim q As Quote
    WriteLegacyQuoteRow "COT-1001"
    Set q = LoadQuote("COT-1001")
    AssertTrue Not q Is Nothing, "LoadQuote"
    If q Is Nothing Then Exit Sub
    AssertEqual "", q.OrderNumber, "sin orden"
    FillTemplate q
    AssertEqual 0, LastErrorNumber(), "sin error"
    AssertEqual "", CStr(ThisWorkbook.Worksheets(SHEET_TEMPLATE).Range(CELL_VEHICLE).Value), "sin vehiculo"
    AssertEqual "", OrderBannerText(q.OrderNumber, Nothing), "sin franja"
End Sub

Public Sub Test_Integracion_AjusteDePagina()
    Dim q As Quote, i As Long
    Set q = NewSampleQuote("COT-1001")
    For i = 2 To MAX_ITEMS
        q.AddItem NewQuoteItem(1, "Concepto " & i, 10)
    Next i
    FillTemplate q
    AssertEqual 0, LastErrorNumber(), "FillTemplate sin error"
    AssertPageFit ThisWorkbook.Worksheets(SHEET_TEMPLATE), True
    FillOrderTemplate NewSampleOrder("OS-0001")
    AssertEqual 0, LastErrorNumber(), "FillOrderTemplate sin error"
    AssertPageFit ThisWorkbook.Worksheets(SHEET_ORDER_TEMPLATE), False
    AssertEqual ORDER_PRINT_AREA, CStr(ThisWorkbook.Worksheets(SHEET_ORDER_TEMPLATE).PageSetup.PrintArea), "area de la orden"
End Sub

Public Sub Test_Integracion_DiagnosticoLargoEnPlantilla()
    Dim o As ServiceOrder, s As String, i As Long
    Do While Len(s) < 1000
        i = i + 1
        s = s & "Linea " & i & ": revisar sistema de frenos y suspension" & vbLf
    Loop
    s = Left$(s, 1000)
    Set o = NewSampleOrder("OS-0001")
    o.Diagnosis = s
    FillOrderTemplate o
    AssertEqual 0, LastErrorNumber(), "sin error"
    AssertEqual s, DiagnosisFromSheet(ThisWorkbook.Worksheets(SHEET_ORDER_TEMPLATE)), "diagnostico completo"
End Sub

Public Sub Test_Integracion_DiagnosticoRenglonesCortosCrece()
    Dim o As ServiceOrder, ws As Worksheet, t As String, r As Long
    t = ShortLines(ORD_DIAG_MAX_LINES)
    Set o = NewSampleOrder("OS-0001")
    o.Diagnosis = t
    FillOrderTemplate o
    AssertEqual 0, LastErrorNumber(), "sin error"
    Set ws = ThisWorkbook.Worksheets(SHEET_ORDER_TEMPLATE)
    For r = ORD_DIAG_FIRST_ROW To ORD_DIAG_FIRST_ROW + ORD_DIAG_ROWS - 1
        AssertTrue Not ws.Rows(r).Hidden, "fila " & r & " visible"
        AssertTrue ws.Rows(r).RowHeight >= ORD_DIAG_LINES_PER_ROW * 12 + ORD_DIAG_ROW_PAD, "fila " & r & " alcanza para 5 renglones con margen"
    Next r
    AssertEqual t, DiagnosisFromSheet(ws), "texto completo"
End Sub

Public Sub Test_Integracion_DiagnosticoCortoOcultaSobrantes()
    Dim o As ServiceOrder, ws As Worksheet, r As Long
    Set o = NewSampleOrder("OS-0001")
    o.Diagnosis = ShortLines(ORD_DIAG_MAX_LINES)
    FillOrderTemplate o
    o.Diagnosis = ShortLines(2)
    FillOrderTemplate o
    AssertEqual 0, LastErrorNumber(), "sin error"
    Set ws = ThisWorkbook.Worksheets(SHEET_ORDER_TEMPLATE)
    For r = ORD_DIAG_FIRST_ROW + 1 To ORD_DIAG_FIRST_ROW + ORD_DIAG_ROWS - 1
        AssertTrue ws.Rows(r).Hidden, "fila " & r & " oculta"
        AssertEqual "", CStr(ws.Range("A" & r).Value), "fila " & r & " vacia"
    Next r
    AssertEqual ShortLines(2), DiagnosisFromSheet(ws), "dos renglones"
    AssertEqual ORD_DIAG_MIN_HEIGHT, CDbl(ws.Rows(ORD_DIAG_FIRST_ROW).RowHeight), "alto minimo"
    o.Diagnosis = ""
    FillOrderTemplate o
    AssertEqual "", CStr(ws.Range(CELL_ORD_DIAGNOSIS).Value), "sin diagnostico"
    AssertTrue Not ws.Rows(ORD_DIAG_FIRST_ROW).Hidden, "primera fila visible"
    AssertEqual ORD_DIAG_MIN_HEIGHT, CDbl(ws.Rows(ORD_DIAG_FIRST_ROW).RowHeight), "alto minimo sin diagnostico"
End Sub

Public Sub Test_Integracion_DiagnosticoDeVariasFilasSinMinimo()
    Dim o As ServiceOrder, ws As Worksheet
    Set o = NewSampleOrder("OS-0001")
    o.Diagnosis = ShortLines(ORD_DIAG_LINES_PER_ROW + 1)
    FillOrderTemplate o
    AssertEqual 0, LastErrorNumber(), "sin error"
    Set ws = ThisWorkbook.Worksheets(SHEET_ORDER_TEMPLATE)
    AssertTrue Not ws.Rows(ORD_DIAG_FIRST_ROW + 1).Hidden, "segunda fila visible"
    AssertTrue ws.Rows(ORD_DIAG_FIRST_ROW).RowHeight < ORD_DIAG_MIN_HEIGHT, "primera fila sin hueco del minimo"
End Sub

Public Sub Test_Integracion_CondicionesYFirmasEnSuLugar()
    Dim o As ServiceOrder, ws As Worksheet, r As Long
    Set o = NewSampleOrder("OS-0001")
    o.Diagnosis = ShortLines(ORD_DIAG_MAX_LINES)
    FillOrderTemplate o
    Set ws = ThisWorkbook.Worksheets(SHEET_ORDER_TEMPLATE)
    AssertTrue Not ws.Range(CELL_ORD_TERMS).Locked, "condiciones editables"
    AssertTrue Len(CStr(ws.Range(CELL_ORD_TERMS).Value)) > 0, "condiciones con texto"
    AssertEqual TPL_CLIENT, CStr(ws.Range(CELL_ORD_SIGN_CLIENT).Offset(1, 0).Value), "etiqueta de firma del cliente"
    For r = ORD_DIAG_FIRST_ROW To ORD_DIAG_FIRST_ROW + ORD_DIAG_ROWS - 1
        AssertEqual "", CStr(ws.Range(ORD_MEASURE_COLUMN & r).Value), "columna de medicion vacia, fila " & r
    Next r
End Sub

Public Sub Test_Integracion_HojasProtegidasTrasFlujo()
    Dim o As ServiceOrder, q As Quote, folio As String, n As Variant
    Set o = NewSampleOrder("OS-0001")
    o.QuoteNumber = ""
    AssertTrue SaveOrder(o), "SaveOrder"
    folio = CreateQuoteForOrder(o)
    Set q = LoadQuote(folio)
    If q Is Nothing Then AssertTrue False, "cotizacion": Exit Sub
    FillTemplate q
    FillOrderTemplate LoadOrder("OS-0001")
    AssertEqual 0, LastErrorNumber(), "sin error"
    AssertTrue ProtectAll(), "ProtectAll"
    For Each n In Array(SHEET_TEMPLATE, SHEET_ORDER_TEMPLATE, SHEET_QUOTES, SHEET_ITEMS, SHEET_CLIENTS, SHEET_ORDERS)
        If SheetExists(CStr(n)) Then
            AssertTrue ThisWorkbook.Worksheets(CStr(n)).ProtectContents, CStr(n) & " protegida"
        End If
    Next n
    AssertTrue ThisWorkbook.ProtectStructure, "estructura protegida"
End Sub

Public Sub Test_Integracion_PiePaginaSinTotal()
    Dim n As Variant, f As String
    FillTemplate NewSampleQuote("COT-1001")
    FillOrderTemplate NewSampleOrder("OS-0001")
    For Each n In Array(SHEET_TEMPLATE, SHEET_ORDER_TEMPLATE)
        f = ThisWorkbook.Worksheets(CStr(n)).PageSetup.LeftFooter
        AssertTrue Right$(f, Len(TPL_PAGE)) = TPL_PAGE, CStr(n) & ": termina en TPL_PAGE"
        AssertTrue InStr(f, "&N") = 0, CStr(n) & ": sin total"
        AssertTrue InStr(f, "&F") = 0, CStr(n) & ": sin nombre de archivo"
    Next n
End Sub

Public Sub Test_Integracion_VersionApp()
    AssertEqual "9.0.0", APP_VERSION
End Sub
