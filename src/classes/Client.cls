VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' Client - Un cliente del catalogo: contacto y datos fiscales (CSF), todos opcionales
'
' Es solo el modelo; no toca hojas (eso es de modData, hoja BD_Clientes).
' En CreatedOn, 0 significa "sin fecha". TaxRegime y CfdiUse guardan la clave del SAT.
'==============================================================================
Option Explicit

Public Name As String  ' Llave del cliente; no cambia despues del alta
Public CreatedOn As Date
Public Phone1 As String
Public Phone2 As String
Public Email As String
Public Address As String
Public Rfc As String
Public LegalName As String  ' Razon social
Public TaxRegime As String  ' Clave del regimen fiscal, ej. "612"
Public TaxZip As String  ' Codigo postal fiscal
Public CfdiUse As String  ' Clave del uso de CFDI, ej. "G03"
Public Street As String
Public ExtNumber As String
Public IntNumber As String
Public Neighborhood As String  ' Colonia
Public Municipality As String  ' Municipio o alcaldia
Public State As String
