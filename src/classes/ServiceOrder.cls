VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServiceOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' ServiceOrder · Una orden de servicio: cliente, vehículo, recepción y entrega
'
' Es solo el modelo; no toca hojas (eso es de modData, hoja BD_Ordenes).
' En los campos Date, 0 significa "sin valor".
'==============================================================================
Option Explicit

Public Number As String  ' Folio, ej. "OS-0001"
Public Status As String  ' Uno de ORDER_STATUSES
Public ClientName As String
Public Phone1 As String
Public Phone2 As String
Public Email As String
Public Address As String
Public Make As String  ' Marca
Public Model As String
Public Engine As String  ' Motor
Public Year As Integer  ' 0 = sin año
Public Color As String
Public Plates As String  ' En mayúsculas al guardar
Public Mileage As Long  ' Kilometraje (km)
Public FuelLevel As String  ' Uno de FUEL_LEVELS
Public Vin As String  ' En mayúsculas al guardar
Public InDate As Date  ' 0 = sin valor
Public InTime As Date  ' 0 = sin valor
Public ReceivedBy As String
Public PromisedDate As Date  ' 0 = sin valor
Public OutDate As Date  ' 0 = sin valor
Public OutTime As Date  ' 0 = sin valor
Public DeliveredBy As String
Public Requested As String  ' Servicio solicitado
Public Diagnosis As String
Public QuoteNumber As String  ' Folio de la cotización ligada (puede ir vacío)

Private Sub Class_Initialize()
    Status = ORDER_STATUS_RECEIVED
    InDate = Date
End Sub
