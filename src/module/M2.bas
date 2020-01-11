Attribute VB_Name = "M2"
Option Explicit

Sub MySub()
 Dim myPrice As Long: myPrice = 500
 Debug.Print GetTaxIncluded(myPrice)
 Debug.Print GetTax(myPrice)
 
End Sub
