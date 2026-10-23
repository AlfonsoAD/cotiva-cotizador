Attribute VB_Name = "modTestClients"
'==============================================================================
' modTestClients - Pruebas del catalogo de clientes (BD_Clientes A:Q) y de los
' catalogos del SAT (BD_RegimenFiscal, BD_UsoCFDI)
'==============================================================================
Option Explicit

Private Function NewFiscalClient(ByVal clientName As String) As Client
    Dim c As Client
    Set c = New Client
    c.Name = clientName
    c.Phone1 = "5551112222"
    c.Phone2 = "5553334444"
    c.Email = "contacto@example.com"
    c.Address = "Av. Siempre Viva 742"
    c.Rfc = "XAXX010101000"
    c.LegalName = "Razon Social de Prueba SA"
    c.TaxRegime = "601"
    c.TaxZip = "06600"
    c.CfdiUse = "G03"
    c.Street = "Reforma"
    c.ExtNumber = "100"
    c.IntNumber = "4B"
    c.Neighborhood = "Juarez"
    c.Municipality = "Cuauhtemoc"
    c.State = "Ciudad de Mexico"
    Set NewFiscalClient = c
End Function

' Fila de BD_Clientes con ese nombre (0 si no esta).
Private Function ClientRow(ByVal clientName As String) As Long
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(SHEET_CLIENTS)
    For r = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If StrComp(CStr(ws.Cells(r, 1).Value), clientName, vbTextCompare) = 0 Then
            ClientRow = r
            Exit Function
        End If
    Next r
End Function

' Filas de una tabla (0 To n-1, 0 To 1); 0 si es Empty.
Private Function ListLength(ByVal list As Variant) As Long
    If IsEmpty(list) Then Exit Function
    ListLength = UBound(list, 1) - LBound(list, 1) + 1
End Function

' Elementos de un arreglo de nombres; 0 si es Empty.
Private Function NamesLength(ByVal names As Variant) As Long
    If IsEmpty(names) Then Exit Function
    NamesLength = UBound(names) - LBound(names) + 1
End Function

Private Function IndexOfKey(ByVal list As Variant, ByVal key As String) As Long
    Dim i As Long
    IndexOfKey = -1
    If IsEmpty(list) Then Exit Function
    For i = 0 To UBound(list, 1)
        If CStr(list(i, 0)) = key Then
            IndexOfKey = i
            Exit Function
        End If
    Next i
End Function

Private Sub OpenSheetForWrite(ByVal sheetName As String)
    TryUnprotectWorkbook
    UnlockSheet ThisWorkbook.Worksheets(sheetName)
End Sub

'--- Estructura (R1-R6) --------------------------------------------------------

Public Sub Test_Clientes_EncabezadosFiscales()
    Dim ws As Worksheet, expected As Variant, i As Long

    OpenSheetForWrite SHEET_CLIENTS
    Set ws = ThisWorkbook.Worksheets(SHEET_CLIENTS)
    ws.Range("A2").Value = "Cliente Previo"
    ws.Range("C2").Value = "5550001111"
    ws.Range("G1:Q1").ClearContents

    EnsureDataHeaders

    expected = Array("RFC", "RazonSocial", "RegimenFiscal", "CPFiscal", "UsoCFDI", "Calle", _
        "NumExterior", "NumInterior", "Colonia", "Municipio", "Estado")
    For i = 0 To 10
        AssertEqual expected(i), CStr(ws.Cells(1, 7 + i).Value), "encabezado " & expected(i)
    Next i
    AssertEqual "Cliente", CStr(ws.Range("A1").Value), "A1 intacto"
    AssertEqual "Direccion", CStr(ws.Range("F1").Value), "F1 intacto"
    AssertEqual "Cliente Previo", CStr(ws.Range("A2").Value), "fila de datos A2"
    AssertEqual "5550001111", CStr(ws.Range("C2").Value), "fila de datos C2"
End Sub

Public Sub Test_Clientes_LibroAnteriorSinFiscales()
    Dim ws As Worksheet, c As Client, o As ServiceOrder

    Set o = NewSampleOrder("OS-0001")
    AssertTrue SaveOrder(o), "SaveOrder"
    MakeLegacyWorkbook
    AssertTrue PrepareWorkbook(), "PrepareWorkbook"

    Set ws = ThisWorkbook.Worksheets(SHEET_CLIENTS)
    AssertEqual "RFC", CStr(ws.Range("G1").Value), "G1"
    AssertEqual "Estado", CStr(ws.Range("Q1").Value), "Q1"
    AssertTrue SheetExists(SHEET_TAX_REGIMES), "hoja de regimenes"
    AssertTrue SheetExists(SHEET_CFDI_USES), "hoja de usos"

    Set c = LoadClient(o.ClientName)
    AssertTrue Not c Is Nothing, "LoadClient existe"
    If c Is Nothing Then Exit Sub
    AssertEqual o.Phone1, c.Phone1, "Phone1 intacto"
    AssertEqual o.Email, c.Email, "Email intacto"
    AssertEqual o.Address, c.Address, "Address intacto"
    AssertEqual "", c.Rfc & c.LegalName & c.TaxRegime & c.TaxZip & c.CfdiUse & c.Street & _
        c.ExtNumber & c.IntNumber & c.Neighborhood & c.Municipality & c.State, "fiscales vacios"
End Sub

Public Sub Test_Catalogos_SembradosYOcultos()
    Dim sheetName As Variant, ws As Worksheet

    AssertTrue PrepareWorkbook(), "PrepareWorkbook"
    For Each sheetName In Array(SHEET_TAX_REGIMES, SHEET_CFDI_USES)
        AssertTrue SheetExists(CStr(sheetName)), sheetName & " existe"
        If SheetExists(CStr(sheetName)) Then
            Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
            AssertEqual "Clave", CStr(ws.Range("A1").Value), sheetName & " A1"
            AssertEqual "Descripcion", CStr(ws.Range("B1").Value), sheetName & " B1"
            AssertTrue Len(CStr(ws.Range("A2").Value)) > 0, sheetName & " sembrada"
            AssertTrue ws.ProtectContents, sheetName & " protegida"
            AssertEqual CLng(xlSheetVeryHidden), CLng(ws.Visible), sheetName & " VeryHidden"
        End If
    Next sheetName
End Sub

Public Sub Test_Catalogos_NoSeReescriben()
    Dim regimes As Variant, ws As Worksheet

    AssertTrue PrepareWorkbook(), "PrepareWorkbook"
    If Not SheetExists(SHEET_TAX_REGIMES) Then
        AssertTrue False, "hoja de regimenes existe"
        Exit Sub
    End If

    OpenSheetForWrite SHEET_TAX_REGIMES
    Set ws = ThisWorkbook.Worksheets(SHEET_TAX_REGIMES)
    ws.Range("B2").Value = "SENTINELA"
    EnsureSatCatalogs
    AssertEqual "SENTINELA", CStr(ws.Range("B2").Value), "hoja con datos no se reescribe"

    OpenSheetForWrite SHEET_TAX_REGIMES
    ws.Rows("2:" & ws.Rows.Count).Delete
    EnsureSatCatalogs
    regimes = ListTaxRegimes()
    AssertEqual 19, ListLength(regimes), "hoja sin filas se siembra"
    AssertTrue CStr(ws.Range("B2").Value) <> "SENTINELA", "semilla repuesta"
End Sub

Public Sub Test_Catalogos_ListaClaveDescripcion()
    Dim regimes As Variant, uses As Variant, ws As Worksheet, k As Long

    AssertTrue PrepareWorkbook(), "PrepareWorkbook"
    regimes = ListTaxRegimes()
    uses = ListCfdiUses()
    AssertEqual 19, ListLength(regimes), "19 regimenes"
    AssertEqual 24, ListLength(uses), "24 usos de CFDI"
    If IsEmpty(regimes) Or IsEmpty(uses) Then Exit Sub

    AssertEqual 0, LBound(regimes, 1), "regimenes base 0"
    AssertEqual 1, UBound(regimes, 2), "dos columnas"
    AssertEqual "601", CStr(regimes(0, 0)), "primer regimen"
    AssertTrue Len(CStr(regimes(0, 1))) > 0, "descripcion del primer regimen"
    AssertEqual "626", CStr(regimes(18, 0)), "ultimo regimen"
    AssertEqual "G01", CStr(uses(0, 0)), "primer uso"
    AssertEqual "CN01", CStr(uses(23, 0)), "ultimo uso"
    AssertTrue IndexOfKey(uses, "G03") >= 0, "G03 en usos"
    AssertTrue IndexOfKey(regimes, "612") >= 0, "612 en regimenes"

    ' Una fila despues de un hueco se lee; el hueco no.
    OpenSheetForWrite SHEET_TAX_REGIMES
    Set ws = ThisWorkbook.Worksheets(SHEET_TAX_REGIMES)
    ws.Range("A24:B24").Value = Array("999", "Regimen de prueba")
    regimes = ListTaxRegimes()
    k = IndexOfKey(regimes, "999")
    AssertEqual 20, ListLength(regimes), "hueco omitido"
    If k >= 0 Then AssertEqual "Regimen de prueba", CStr(regimes(k, 1)), "descripcion de la fila nueva"
    ws.Range("A24:B24").ClearContents
End Sub

'--- Clientes (R7-R13) ---------------------------------------------------------

Public Sub Test_SaveClient_AltaNueva()
    Dim src As Client, c As Client

    Set src = NewFiscalClient("Cliente Fiscal")
    AssertTrue SaveClient(src), "SaveClient"
    Set c = LoadClient("Cliente Fiscal")
    AssertTrue Not c Is Nothing, "LoadClient existe"
    If c Is Nothing Then Exit Sub

    AssertEqual "Cliente Fiscal", c.Name, "Name"
    AssertEqual Date, c.CreatedOn, "CreatedOn = hoy"
    AssertEqual src.Phone1, c.Phone1, "Phone1"
    AssertEqual src.Phone2, c.Phone2, "Phone2"
    AssertEqual src.Email, c.Email, "Email"
    AssertEqual src.Address, c.Address, "Address"
    AssertEqual src.Rfc, c.Rfc, "Rfc"
    AssertEqual src.LegalName, c.LegalName, "LegalName"
    AssertEqual src.TaxRegime, c.TaxRegime, "TaxRegime"
    AssertEqual src.TaxZip, c.TaxZip, "TaxZip"
    AssertEqual src.CfdiUse, c.CfdiUse, "CfdiUse"
    AssertEqual src.Street, c.Street, "Street"
    AssertEqual src.ExtNumber, c.ExtNumber, "ExtNumber"
    AssertEqual src.IntNumber, c.IntNumber, "IntNumber"
    AssertEqual src.Neighborhood, c.Neighborhood, "Neighborhood"
    AssertEqual src.Municipality, c.Municipality, "Municipality"
    AssertEqual src.State, c.State, "State"
    AssertEqual 1, NamesLength(ListClients()), "una fila en el catalogo"
End Sub

Public Sub Test_SaveClient_EditaYBorraVacios()
    Dim c As Client, loaded As Client

    AssertTrue SaveClient(NewFiscalClient("Cliente Fiscal")), "alta"
    Set c = NewFiscalClient("cliente fiscal ")
    c.CreatedOn = DateSerial(2000, 1, 1)
    c.Phone1 = ""
    c.Rfc = ""
    c.Street = "Calle Nueva"
    AssertTrue SaveClient(c), "edicion"

    Set loaded = LoadClient("Cliente Fiscal")
    AssertTrue Not loaded Is Nothing, "LoadClient existe"
    If loaded Is Nothing Then Exit Sub
    AssertEqual 1, NamesLength(ListClients()), "no duplica la fila"
    AssertEqual Date, loaded.CreatedOn, "FechaAlta conservada"
    AssertEqual "Cliente Fiscal", loaded.Name, "nombre original conservado"
    AssertEqual "", loaded.Phone1, "telefono borrado"
    AssertEqual "", loaded.Rfc, "RFC borrado"
    AssertEqual "Calle Nueva", loaded.Street, "calle nueva"
    AssertEqual "5553334444", loaded.Phone2, "Phone2 intacto"
End Sub

Public Sub Test_SaveClient_CeroALaIzquierda()
    Dim c As Client, loaded As Client, ws As Worksheet, r As Long

    Set c = NewFiscalClient("Cliente Ceros")
    c.TaxZip = "01000"
    c.Phone1 = "0123456789"
    c.ExtNumber = "007"
    c.IntNumber = "01"
    AssertTrue SaveClient(c), "SaveClient"

    Set loaded = LoadClient("Cliente Ceros")
    AssertTrue Not loaded Is Nothing, "LoadClient existe"
    If loaded Is Nothing Then Exit Sub
    AssertEqual "01000", loaded.TaxZip, "CP"
    AssertEqual "0123456789", loaded.Phone1, "telefono"
    AssertEqual "007", loaded.ExtNumber, "NumExterior"
    AssertEqual "01", loaded.IntNumber, "NumInterior"

    Set ws = ThisWorkbook.Worksheets(SHEET_CLIENTS)
    r = ClientRow("Cliente Ceros")
    AssertTrue r > 0, "fila encontrada"
    If r = 0 Then Exit Sub
    AssertEqual "01000", CStr(ws.Cells(r, 10).Value), "celda CP"
    AssertEqual "@", CStr(ws.Cells(r, 10).NumberFormat), "formato texto del CP"
End Sub

Public Sub Test_SaveClient_RfcMayusculas()
    Dim c As Client, loaded As Client

    Set c = NewFiscalClient("Cliente Rfc")
    c.Rfc = "  xaxx010101abc "
    AssertTrue SaveClient(c), "SaveClient"
    Set loaded = LoadClient("Cliente Rfc")
    AssertTrue Not loaded Is Nothing, "LoadClient existe"
    If loaded Is Nothing Then Exit Sub
    AssertEqual "XAXX010101ABC", loaded.Rfc, "RFC recortado y en mayusculas"
End Sub

Public Sub Test_SaveClient_NombreVacioFalse()
    Dim c As Client

    Set c = NewFiscalClient("   ")
    AssertTrue Not SaveClient(c), "nombre vacio regresa False"
    AssertTrue IsEmpty(ListClients()), "no se escribio ninguna fila"
    AssertTrue SaveClient(NewFiscalClient("Con Nombre")), "con nombre regresa True"
    AssertEqual 1, NamesLength(ListClients()), "solo la fila valida"
End Sub

Public Sub Test_SaveClient_NothingFalla()
    AssertFails "modData.SaveClient", "SaveClient Nothing", Nothing
End Sub

Public Sub Test_LoadClient_NoExisteNothing()
    Dim c As Client

    AssertTrue LoadClient("Nadie") Is Nothing, "no existe"
    AssertTrue SaveClient(NewFiscalClient("Cliente Fiscal")), "alta"
    AssertTrue LoadClient("Otro") Is Nothing, "otro nombre no existe"
    Set c = LoadClient("  CLIENTE FISCAL ")
    AssertTrue Not c Is Nothing, "sin distinguir mayusculas ni extremos"
End Sub

Public Sub Test_SaveOrder_ConservaFiscales()
    Dim o As ServiceOrder, c As Client

    Set o = NewSampleOrder("OS-0001")
    AssertTrue SaveClient(NewFiscalClient(o.ClientName)), "alta fiscal"
    o.Phone1 = "5559998888"
    o.Email = ""
    AssertTrue SaveOrder(o), "SaveOrder"

    Set c = LoadClient(o.ClientName)
    AssertTrue Not c Is Nothing, "LoadClient existe"
    If c Is Nothing Then Exit Sub
    AssertEqual "5559998888", c.Phone1, "telefono actualizado por la orden"
    AssertEqual "contacto@example.com", c.Email, "correo vacio no pisa"
    AssertEqual "XAXX010101000", c.Rfc, "RFC conservado"
    AssertEqual "Razon Social de Prueba SA", c.LegalName, "razon social conservada"
    AssertEqual "601", c.TaxRegime, "regimen conservado"
    AssertEqual "06600", c.TaxZip, "CP conservado"
    AssertEqual "G03", c.CfdiUse, "uso conservado"
    AssertEqual "Ciudad de Mexico", c.State, "estado conservado"
End Sub

Public Sub Test_FindClientByRfc_OtroCliente()
    Dim c As Client

    AssertTrue SaveClient(NewFiscalClient("Cliente Uno")), "alta 1"
    Set c = NewFiscalClient("Cliente Dos")
    c.Rfc = "PEPJ800101AB1"
    AssertTrue SaveClient(c), "alta 2"

    AssertEqual "Cliente Uno", FindClientByRfc("xaxx010101000", "Cliente Dos"), "otro cliente, sin distinguir mayusculas"
    AssertEqual "", FindClientByRfc("XAXX010101000", "Cliente Uno"), "el mismo cliente no cuenta"
    AssertEqual "", FindClientByRfc("", "Cliente Dos"), "RFC vacio"
    AssertEqual "", FindClientByRfc("ZZZZ010101000", "Cliente Dos"), "RFC sin dueno"
End Sub

Public Sub Test_SearchClients_PorRfcYTelefono()
    Dim a As Client, b As Client, d As Client, found As Variant

    Set a = NewFiscalClient("Beta Taller")
    a.Rfc = "AAAA010101AA1": a.Phone1 = "5550000001": a.Phone2 = "": a.Email = "b@uno.mx": a.LegalName = "Uno SA"
    Set b = NewFiscalClient("Alfa Garage")
    b.Rfc = "BBBB020202BB2": b.Phone1 = "5550000002": b.Phone2 = "7771234567": b.Email = "a@dos.mx": b.LegalName = "Dos SA"
    Set d = NewFiscalClient("Gamma Auto")
    d.Rfc = "CCCC030303CC3": d.Phone1 = "5550000003": d.Phone2 = "": d.Email = "g@tres.mx": d.LegalName = "Tres SA"
    AssertTrue SaveClient(a), "alta a"
    AssertTrue SaveClient(b), "alta b"
    AssertTrue SaveClient(d), "alta d"

    found = SearchClients("bbbb02")
    AssertEqual 1, NamesLength(found), "por RFC"
    If NamesLength(found) = 1 Then AssertEqual "Alfa Garage", CStr(found(LBound(found))), "RFC -> Alfa"

    AssertEqual 1, NamesLength(SearchClients("7771234")), "por telefono 2"
    AssertEqual 1, NamesLength(SearchClients("TRES SA")), "por razon social"
    AssertEqual 1, NamesLength(SearchClients("@dos.mx")), "por correo"

    found = SearchClients("5550000")
    AssertEqual 3, NamesLength(found), "tres coinciden por telefono"
    If NamesLength(found) = 3 Then
        AssertEqual "Alfa Garage", CStr(found(LBound(found))), "orden alfabetico 1"
        AssertEqual "Beta Taller", CStr(found(LBound(found) + 1)), "orden alfabetico 2"
        AssertEqual "Gamma Auto", CStr(found(LBound(found) + 2)), "orden alfabetico 3"
    End If

    AssertEqual 3, NamesLength(SearchClients("")), "texto vacio = todos"
    AssertTrue IsEmpty(SearchClients("no-existe-xyz")), "sin coincidencias = Empty"
End Sub

Public Sub Test_IsValidPostalCode()
    AssertTrue IsValidPostalCode("01000"), "01000"
    AssertTrue IsValidPostalCode("06600"), "06600"
    AssertTrue Not IsValidPostalCode("1000"), "4 digitos"
    AssertTrue Not IsValidPostalCode("010000"), "6 digitos"
    AssertTrue Not IsValidPostalCode("0100a"), "letra"
    AssertTrue Not IsValidPostalCode(""), "vacio"
    AssertTrue Not IsValidPostalCode("01 00"), "espacio"
End Sub
