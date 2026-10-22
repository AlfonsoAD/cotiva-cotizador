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
