Attribute VB_Name = "M1"
Option Explicit

Public Const TAX_RATE As Currency = 0.08
Public Function GetTaxIncluded(ByVal price As Long) As Currency
  GetTaxIncluded = price * (1 + TAX_RATE)
End Function
Public Function GetTax(ByVal price As Long) As Currency
  GetTax = price * TAX_RATE
End Function

'Test to use Object Module
Sub MytestUseOM()
  Sheet1.FirstName = "Mike"
  Debug.Print Sheet1.Greet
  
End Sub
