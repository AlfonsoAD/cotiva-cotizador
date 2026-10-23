Attribute VB_Name = "modTestData"
'==============================================================================
' modTestData - Pruebas de la capa de datos (BD_Cotizaciones, BD_Ordenes, BD_Clientes)
'==============================================================================
Option Explicit

Public Sub Test_P1_FolioOrdenVacio()
    AssertEqual "OS-0001", GetNextOrderNumber(), "primer folio de orden"
End Sub

Public Sub Test_P2_FolioCotizacionIntacto()
    AssertEqual "COT-1001", GetNextQuoteNumber(), "libro vacio"
    AssertTrue SaveQuote(NewSampleQuote("COT-1007")), "SaveQuote"
    AssertEqual "COT-1008", GetNextQuoteNumber(), "despues de COT-1007"
End Sub

Public Sub Test_P3_GuardarCargarOrden()
    Dim src As ServiceOrder, o As ServiceOrder
    Set src = NewSampleOrder("OS-0001")
    AssertTrue SaveOrder(src), "SaveOrder"
    Set o = LoadOrder("OS-0001")
    AssertTrue Not o Is Nothing, "LoadOrder existe"
    If o Is Nothing Then Exit Sub

    AssertEqual "OS-0001", o.Number, "Number"
    AssertEqual src.Status, o.Status, "Status"
    AssertEqual src.ClientName, o.ClientName, "ClientName"
    AssertEqual "0123456789", o.Phone1, "Phone1"
    AssertEqual src.Phone2, o.Phone2, "Phone2"
    AssertEqual src.Email, o.Email, "Email"
    AssertEqual src.Address, o.Address, "Address"
    AssertEqual src.Make, o.Make, "Make"
    AssertEqual src.Model, o.Model, "Model"
    AssertEqual src.Engine, o.Engine, "Engine"
    AssertEqual src.Year, o.Year, "Year"
    AssertEqual src.Color, o.Color, "Color"
    AssertEqual UCase$(src.Plates), o.Plates, "Plates en mayusculas"
    AssertEqual src.Mileage, o.Mileage, "Mileage"
    AssertEqual src.FuelLevel, o.FuelLevel, "FuelLevel"
    AssertEqual UCase$(src.Vin), o.Vin, "Vin en mayusculas"
    AssertEqual src.InDate, o.InDate, "InDate"
    AssertEqual src.InTime, o.InTime, "InTime 00:00 conservado"
    AssertEqual src.ReceivedBy, o.ReceivedBy, "ReceivedBy"
    AssertEqual src.PromisedDate, o.PromisedDate, "PromisedDate"
    AssertEqual src.OutDate, o.OutDate, "OutDate"
    AssertEqual src.OutTime, o.OutTime, "OutTime"
    AssertEqual src.DeliveredBy, o.DeliveredBy, "DeliveredBy"
    AssertEqual src.Requested, o.Requested, "Requested"
    AssertEqual src.Diagnosis, o.Diagnosis, "Diagnosis"
    AssertEqual src.QuoteNumber, o.QuoteNumber, "QuoteNumber"
End Sub

Public Sub Test_P4_ReemplazarOrden()
    Dim o As ServiceOrder, orderRows As Variant
    Set o = NewSampleOrder("OS-0001")
    o.Status = ORDER_STATUS_RECEIVED
    AssertTrue SaveOrder(o), "primer SaveOrder"
    o.Status = ORDER_STATUS_READY
    AssertTrue SaveOrder(o), "segundo SaveOrder"

    orderRows = ListOrders()
    AssertTrue Not IsEmpty(orderRows), "ListOrders con datos"
    If IsEmpty(orderRows) Then Exit Sub
    AssertEqual 0, UBound(orderRows, 1), "una sola fila"
    AssertEqual ORDER_STATUS_READY, orderRows(0, 3), "segundo Status"
End Sub

Public Sub Test_P5_CotizacionVieja()
    Dim q As Quote
    WriteLegacyQuoteRow "COT-1001"
    Set q = LoadQuote("COT-1001")
    AssertTrue Not q Is Nothing, "LoadQuote existe"
    If q Is Nothing Then Exit Sub
    AssertEqual "", q.OrderNumber, "OrderNumber vacio"
    ClearLastError
    AssertTrue SaveQuote(q), "SaveQuote"
    AssertEqual 0, LastErrorNumber(), "sin error"
End Sub

Public Sub Test_P6_CotizacionConOrden()
    Dim q As Quote
    Set q = NewSampleQuote("COT-1001")
    q.OrderNumber = "os-0001"
    AssertTrue SaveQuote(q), "primer SaveQuote"
    Set q = LoadQuote("COT-1001")
    AssertEqual "OS-0001", q.OrderNumber, "OrderNumber normalizado"
    AssertTrue SaveQuote(q), "segundo SaveQuote"
    Set q = LoadQuote("COT-1001")
    AssertEqual "OS-0001", q.OrderNumber, "OrderNumber conservado"
End Sub

Public Sub Test_P7_ContactoDelCatalogo()
    Dim o As ServiceOrder, p1 As String, p2 As String, em As String, ad As String
    Set o = NewSampleOrder("OS-0001")
    AssertTrue SaveOrder(o), "SaveOrder"
    AssertTrue GetClientContact(o.ClientName, p1, p2, em, ad), "hay contacto"
    AssertEqual o.Phone1, p1, "Phone1"
    AssertEqual o.Phone2, p2, "Phone2"
    AssertEqual o.Email, em, "Email"
    AssertEqual o.Address, ad, "Address"
End Sub

Public Sub Test_P7_ContactoDeOrdenReciente()
    Dim oNew As ServiceOrder, oOld As ServiceOrder
    Dim p1 As String, p2 As String, em As String, ad As String
    Set oNew = NewSampleOrder("OS-0001")
    oNew.InDate = DateSerial(2026, 5, 20)
    oNew.Phone1 = "5550000002"
    oNew.Email = "nuevo@example.com"
    Set oOld = NewSampleOrder("OS-0002")
    oOld.InDate = DateSerial(2026, 1, 10)
    oOld.Phone1 = "5550000001"
    oOld.Email = "viejo@example.com"
    AssertTrue SaveOrder(oNew), "SaveOrder reciente"
    AssertTrue SaveOrder(oOld), "SaveOrder antigua"

    ClearClientContacts
    AssertTrue GetClientContact(oNew.ClientName, p1, p2, em, ad), "hay contacto"
    AssertEqual "5550000002", p1, "Phone1 de la mas reciente"
    AssertEqual "nuevo@example.com", em, "Email de la mas reciente"
End Sub

Public Sub Test_P7_CampoVacioNoBorra()
    Dim o As ServiceOrder, savedEmail As String
    Dim p1 As String, p2 As String, em As String, ad As String
    Set o = NewSampleOrder("OS-0001")
    savedEmail = o.Email
    AssertTrue SaveOrder(o), "primer SaveOrder"
    Set o = NewSampleOrder("OS-0002")
    o.Email = ""
    AssertTrue SaveOrder(o), "segundo SaveOrder"
    AssertTrue GetClientContact(o.ClientName, p1, p2, em, ad), "hay contacto"
    AssertEqual savedEmail, em, "Email conservado"
End Sub

Public Sub Test_P8_LibroAnterior()
    Dim ws As Worksheet
    MakeLegacyWorkbook
    AssertTrue PrepareWorkbook(), "PrepareWorkbook"
    Set ws = ThisWorkbook.Worksheets(SHEET_ORDERS)
    AssertEqual CLng(xlSheetVeryHidden), CLng(ws.Visible), "BD_Ordenes VeryHidden"
    AssertTrue ws.ProtectContents, "BD_Ordenes protegida"
    AssertEqual "OrdenFolio", CStr(ThisWorkbook.Worksheets(SHEET_QUOTES).Range("I1").Value), "I1"
    With ThisWorkbook.Worksheets(SHEET_CLIENTS)
        AssertEqual "Telefono1", CStr(.Range("C1").Value), "C1"
        AssertEqual "Telefono2", CStr(.Range("D1").Value), "D1"
        AssertEqual "Correo", CStr(.Range("E1").Value), "E1"
        AssertEqual "Direccion", CStr(.Range("F1").Value), "F1"
    End With
    AssertTrue ThisWorkbook.ProtectStructure, "libro protegido"
End Sub

Public Sub Test_SaveOrder_NothingFalla()
    AssertFails "modData.SaveOrder", "SaveOrder Nothing", Nothing
End Sub

Public Sub Test_ListOrderStaff_DistintosOrdenados()
    Dim o As ServiceOrder, staff As Variant
    Set o = NewSampleOrder("OS-0001")
    o.ReceivedBy = "Ana"
    o.DeliveredBy = " ana "
    AssertTrue SaveOrder(o), "SaveOrder 1"
    Set o = NewSampleOrder("OS-0002")
    o.ReceivedBy = "Beto"
    o.DeliveredBy = "Beto"
    AssertTrue SaveOrder(o), "SaveOrder 2"

    staff = ListOrderStaff()
    AssertTrue Not IsEmpty(staff), "hay nombres"
    If IsEmpty(staff) Then Exit Sub
    AssertEqual 2, UBound(staff) - LBound(staff) + 1, "dos nombres"
    AssertEqual 0, LBound(staff), "0-based"
    AssertEqual "Ana", staff(0), "primero"
    AssertEqual "Beto", staff(1), "segundo"
End Sub

Public Sub Test_ListOrderStaff_SinOrdenesEmpty()
    AssertTrue IsEmpty(ListOrderStaff()), "sin ordenes"
End Sub
