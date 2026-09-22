VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Quote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' Quote · Una cotización completa: encabezado, conceptos y totales
'
' Los totales se calculan al momento a partir de los conceptos; no se guardan
' en variables, así siempre coinciden con lo que hay en la lista.
'==============================================================================
Option Explicit

Public Number As String            ' Folio, ej. "COT-1001"
Public QuoteDate As Date           ' Fecha de emisión
Public ClientName As String        ' Nombre del cliente
Public TaxRate As Double           ' 0.16, 0.08 o 0

Private mItems As Collection       ' Conceptos (objetos QuoteItem)

Private Sub Class_Initialize()
    Set mItems = New Collection
    QuoteDate = Date
    TaxRate = TaxRateOf(DEFAULT_TAX_OPTION)
End Sub

'--- Conceptos -----------------------------------------------------------------

Public Property Get ItemCount() As Long
    ItemCount = mItems.Count
End Property

' Concepto en la posición indicada (empieza en 1).
Public Function Item(ByVal index As Long) As QuoteItem
    Set Item = mItems(index)
End Function

Public Sub AddItem(ByVal newItem As QuoteItem)
    mItems.Add newItem
End Sub

' Reemplaza el concepto de una posición conservando el orden de la lista.
Public Sub ReplaceItem(ByVal index As Long, ByVal newItem As QuoteItem)
    mItems.Add newItem, Before:=index          ' Se inserta el nuevo antes del viejo...
    mItems.Remove index + 1                    ' ...y se quita el viejo
End Sub

Public Sub RemoveItem(ByVal index As Long)
    mItems.Remove index
End Sub

'--- Totales -------------------------------------------------------------------

Public Property Get Subtotal() As Double
    Dim i As Long, sum As Double
    For i = 1 To mItems.Count
        sum = sum + mItems(i).Amount
    Next i
    Subtotal = Round2(sum)
End Property

Public Property Get Tax() As Double
    Tax = Round2(Subtotal * TaxRate)
End Property

Public Property Get Total() As Double
    Total = Round2(Subtotal + Tax)
End Property

' Fecha hasta la que es válida la cotización.
Public Function ValidUntil(ByVal validityDays As Long) As Date
    ValidUntil = QuoteDate + validityDays
End Function
