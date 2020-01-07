Attribute VB_Name = "Module1"
Option Explicit
Private price_ As Long
Public Property Let price(ByVal newPrice As Long)
  If newPrice >= 0 Then price_ = newPrice Else price_ = 0
End Property
Public Property Get price() As Long
  price = price_
End Property
Public Property Get TaxIncluded() As Currency
  Const TAX_RATE  As Currency = 0.1
  TaxIncluded = price_ * (1 + TAX_RATE)
End Property
