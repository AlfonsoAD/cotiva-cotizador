Attribute VB_Name = "modTestUtils"
'==============================================================================
' modTestUtils - Pruebas de las funciones puras
'==============================================================================
Option Explicit

Public Sub Test_AmountToWords_MilCuarentaYCuatro()
    AssertEqual "(MIL CUARENTA Y CUATRO PESOS 00/100 M.N.)", AmountToWords(1044), "1044"
End Sub

Public Sub Test_AmountToWords_UnPeso()
    AssertEqual "(UN PESO 00/100 M.N.)", AmountToWords(1), "1"
End Sub

Public Sub Test_AmountToWords_SoloCentavos()
    AssertEqual "(CERO PESOS 50/100 M.N.)", AmountToWords(0.5), "0.5"
End Sub

Public Sub Test_AmountToWords_UnMillon()
    AssertEqual "(UN MILL" & ChrW(211) & "N DE PESOS 00/100 M.N.)", AmountToWords(1000000), "1000000"
End Sub

Public Sub Test_AmountToWords_VeintiunMil()
    AssertEqual "(VEINTI" & ChrW(218) & "N MIL PESOS 00/100 M.N.)", AmountToWords(21000), "21000"
End Sub

Public Sub Test_Round2_RedondeaComoExcel()
    AssertEqual 0.13, Round2(0.125), "0.125"
    AssertEqual -0.13, Round2(-0.125), "-0.125"
    AssertEqual 2, Round2(2), "2"
End Sub

Public Sub Test_TryParseNumber_Validos()
    Dim n As Double
    AssertTrue TryParseNumber("12", n), "12 se reconoce"
    AssertEqual 12, n, "12"
    AssertTrue TryParseNumber(" 7 ", n), "7 con espacios se reconoce"
    AssertEqual 7, n, "7"
    AssertTrue TryParseNumber("1" & DecimalSeparator() & "5", n), "1.5 se reconoce"
    AssertEqual 1.5, n, "1.5"
End Sub

Public Sub Test_TryParseNumber_Invalidos()
    Dim n As Double
    AssertTrue Not TryParseNumber("", n), "vacio"
    AssertTrue Not TryParseNumber("abc", n), "abc"
End Sub

Public Sub Test_IsValidRfc()
    AssertTrue IsValidRfc("ABC010203XY1"), "moral"
    AssertTrue IsValidRfc("ABCD010203XY1"), "fisica"
    AssertTrue IsValidRfc("abcd010203xy1"), "minusculas"
    AssertTrue Not IsValidRfc("AB010203"), "corto"
End Sub

Public Sub Test_SanitizeFileName()
    AssertEqual "a_b_c__", SanitizeFileName("a/b:c*?"), "caracteres prohibidos"
    AssertEqual "x y", SanitizeFileName("  x   y  "), "espacios"
    AssertEqual 120, Len(SanitizeFileName(String$(200, "a"))), "largo maximo"
End Sub

Public Sub Test_SafeCellText()
    AssertEqual "'=1+1", SafeCellText("=1+1"), "formula"
    AssertEqual "'-5", SafeCellText("-5"), "negativo"
    AssertEqual "hola", SafeCellText("hola"), "texto"
    AssertEqual "", SafeCellText(""), "vacio"
End Sub

Public Sub Test_TaxRateOf_TaxOptionOf()
    AssertEqual 0.16, TaxRateOf(taxStandard), "tasa estandar"
    AssertEqual 0.08, TaxRateOf(taxBorder), "tasa frontera"
    AssertEqual 0, TaxRateOf(taxExempt), "tasa exenta"
    AssertEqual CLng(taxStandard), TaxOptionOf(0.16), "opcion estandar"
    AssertEqual CLng(taxBorder), TaxOptionOf(0.08), "opcion frontera"
    AssertEqual CLng(taxExempt), TaxOptionOf(0), "opcion exenta"
End Sub

Public Sub Test_TryParseHexColor()
    Dim c As Long
    AssertTrue TryParseHexColor("#CF4A06", c), "con almohadilla"
    AssertEqual RGB(207, 74, 6), c, "color con almohadilla"
    c = 0
    AssertTrue TryParseHexColor("cf4a06", c), "minusculas"
    AssertEqual RGB(207, 74, 6), c, "color en minusculas"
    AssertTrue Not TryParseHexColor("#12345", c), "5 digitos"
    AssertTrue Not TryParseHexColor("GG0000", c), "no hexadecimal"
End Sub

Public Sub Test_ShowError_ModoPruebaRegistra()
    SetTestMode True
    ClearLastError
    On Error GoTo Handler
    Err.Raise 5
    Exit Sub
Handler:
    ShowError "X"
    AssertEqual "X", LastErrorWhere(), "donde"
    AssertEqual 5, LastErrorNumber(), "numero"
    AssertTrue Len(LastErrorDescription()) > 0, "descripcion"
End Sub

Public Sub Test_IsValidDate_AceptaFechas()
    AssertTrue IsValidDate("15/03/2026"), "dd/mm/yyyy"
    AssertTrue IsValidDate(" 2026-03-15 "), "iso con espacios"
    AssertTrue IsValidDate("29/02/2024"), "bisiesto"
End Sub

Public Sub Test_IsValidDate_RechazaTextoYVacio()
    AssertTrue Not IsValidDate(""), "vacio"
    AssertTrue Not IsValidDate("   "), "solo espacios"
    AssertTrue Not IsValidDate("hola"), "texto"
    AssertTrue Not IsValidDate("31/02/2026"), "dia inexistente"
End Sub

Public Sub Test_IsValidTime_AceptaHoras()
    AssertTrue IsValidTime("9:05"), "H:MM"
    AssertTrue IsValidTime("09:05"), "HH:MM"
    AssertTrue IsValidTime("00:00"), "medianoche"
    AssertTrue IsValidTime("23:59"), "ultimo minuto"
End Sub

Public Sub Test_IsValidTime_RechazaFueraDeFormato()
    AssertTrue Not IsValidTime("24:00"), "24:00"
    AssertTrue Not IsValidTime("12:60"), "12:60"
    AssertTrue Not IsValidTime("1205"), "sin dos puntos"
    AssertTrue Not IsValidTime("12:5"), "un digito de minutos"
    AssertTrue Not IsValidTime("ab:cd"), "letras"
    AssertTrue Not IsValidTime(""), "vacio"
End Sub

Public Sub Test_IsValidVehicleYear_Limites()
    Dim ref As Date
    ref = DateSerial(2026, 6, 15)
    AssertTrue IsValidVehicleYear("1950", ref), "minimo"
    AssertTrue IsValidVehicleYear("2027", ref), "ref+1"
    AssertTrue Not IsValidVehicleYear("1949", ref), "bajo el minimo"
    AssertTrue Not IsValidVehicleYear("2028", ref), "ref+2"
    AssertTrue Not IsValidVehicleYear("95", ref), "dos digitos"
    AssertTrue Not IsValidVehicleYear("20a0", ref), "letra"
    AssertTrue Not IsValidVehicleYear("", ref), "vacio"
End Sub

Public Sub Test_IsValidMileage()
    AssertTrue IsValidMileage("0"), "cero"
    AssertTrue IsValidMileage("123456789"), "9 digitos"
    AssertTrue Not IsValidMileage("12,000"), "coma"
    AssertTrue Not IsValidMileage("1.5"), "punto"
    AssertTrue Not IsValidMileage("-5"), "negativo"
    AssertTrue Not IsValidMileage("1234567890"), "10 digitos"
    AssertTrue Not IsValidMileage(""), "vacio"
End Sub

Public Sub Test_IsValidVin()
    AssertTrue IsValidVin("1HGCM82633A004352"), "17 mayusculas"
    AssertTrue IsValidVin("1hgcm82633a004352"), "17 minusculas"
    AssertTrue Not IsValidVin("1HGCM82633A00435"), "16"
    AssertTrue Not IsValidVin("1HGCM82633A0043522"), "18"
    AssertTrue Not IsValidVin("1HGCM 82633A004352"), "espacio interno"
    AssertTrue Not IsValidVin(""), "vacio"
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
