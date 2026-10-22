Attribute VB_Name = "modTestCompany"
'==============================================================================
' modTestCompany - Colores de empresa y nombre del logo del formulario
'==============================================================================
Option Explicit

Public Sub Test_PrimaryColor_VaciaUsaCotiva()
    SetCompanyCell COMPANY_PRIMARY_COLOR_CELL, ""
    AssertEqual BrandOrange(), PrimaryColor(), "B11 vacia"
End Sub

Public Sub Test_PrimaryColor_InvalidaUsaCotiva()
    SetCompanyCell COMPANY_PRIMARY_COLOR_CELL, "#12G45Z"
    AssertEqual BrandOrange(), PrimaryColor(), "hex con letra invalida"
    SetCompanyCell COMPANY_PRIMARY_COLOR_CELL, "12345"
    AssertEqual BrandOrange(), PrimaryColor(), "hex de 5 digitos"
    SetCompanyCell COMPANY_PRIMARY_COLOR_CELL, ""
End Sub

Public Sub Test_PrimaryColor_Valida()
    SetCompanyCell COMPANY_PRIMARY_COLOR_CELL, "#1a2b3c"
    AssertEqual RGB(&H1A, &H2B, &H3C), PrimaryColor(), "con # y minusculas"
    SetCompanyCell COMPANY_PRIMARY_COLOR_CELL, "1A2B3C"
    AssertEqual RGB(&H1A, &H2B, &H3C), PrimaryColor(), "sin # y mayusculas"
    SetCompanyCell COMPANY_PRIMARY_COLOR_CELL, ""
End Sub

Public Sub Test_SecondaryColor_VaciaUsaCotiva()
    SetCompanyCell COMPANY_SECONDARY_COLOR_CELL, ""
    AssertEqual BrandInk(), SecondaryColor(), "B12 vacia"
End Sub

Public Sub Test_SecondaryColor_Valida()
    SetCompanyCell COMPANY_SECONDARY_COLOR_CELL, "0F7F3A"
    AssertEqual RGB(&H0F, &H7F, &H3A), SecondaryColor(), "0F7F3A"
    SetCompanyCell COMPANY_SECONDARY_COLOR_CELL, ""
End Sub

Public Sub Test_FormLogoFile_NombreUnico()
    AssertEqual "logo-form.jpg", FORM_LOGO_FILE, "nombre del logo del formulario"
End Sub
