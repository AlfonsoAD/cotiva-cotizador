VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QuoteItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' QuoteItem · Un concepto de la cotización
'
' Para crear uno ya lleno usa NewQuoteItem(cantidad, descripción, precio)
' de modUtils, porque las clases de VBA no reciben parámetros al crearse.
'==============================================================================
Option Explicit

Public Quantity As Double          ' Cantidad (> 0)
Public Description As String       ' Descripción del concepto
Public UnitPrice As Double         ' Precio unitario (>= 0)

' Importe = cantidad x precio, redondeado a 2 decimales.
' Es calculado: no se guarda, así nunca queda desincronizado.
Public Property Get Amount() As Double
    Amount = Round2(Quantity * UnitPrice)
End Property
