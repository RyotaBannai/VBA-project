VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private x As String 'Private

Sub MySub()
    Debug.Print (d1)
End Sub

Function d1() As Double
    d1 = CDbl(#12/31/1899 12:00:00 PM#)
End Function

Sub MySub2()
    Dim wb As Workbook
    Set wb = Workbooks.Add
    Debug.Print wb.Name
End Sub









