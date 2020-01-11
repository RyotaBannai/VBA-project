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
Private Enum myEnums
  hoge
  fuga
  piyo
End Enum

Private Enum e
  Todohuken = 1
  Kenchoshodaiti
  Populaiton
  Area
End Enum

Sub MySub()
    Debug.Print hoge, fuga, piyo
    Debug.Print e.Todohuken, e.Kenchoshodaiti
End Sub

Function d1() As Double
    d1 = CDbl(#12/31/1899 12:00:00 PM#)
End Function

Sub MySub2()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim rng As Range: Set rng = ActiveCell
    Debug.Print wb.Name
    Debug.Print rng.Address
End Sub

Sub MyGrade()
  Dim rank As Long: rank = 1
  Dim msg As String
  
  Select Case rank
   Case 4
    msg = "Great!"
  Case 3
    msg = "Average"
  Case 2
    msg = "Below Average"
  Case Else
    msg = "Fail"
  End Select
  Debug.Print msg
End Sub

Sub MyCounter()
 Dim i As Long
 For i = 1 To 5
  Debug.Print "Value of i:", i
Next i

For i = 10 To 0 Step -3
  Debug.Print "Value of i:", i
Next i
End Sub

Sub MyLoop()
  Dim cell As Range
  For Each cell In Sheet1.Range("A1:C2")
    'Debug.Print cell.Address
  Next cell
  
  Dim i As Long
  For i = 1 To Sheet1.Range("A1:C2").Count
    Debug.Print i, Sheet1.Range("A1:C2")(i).Address
  Next i
    
End Sub

Sub MyStop()
  Dim i As Long
  For i = 1 To 10
    'If i Mod 5 = 0 Then Stop
    Debug.Assert i Mod 5 <> 0
    Debug.Print "Value of i", i
  Next i
End Sub

Sub MyErrorHandler()

  On Error GoTo ErrorHandler
  Dim x As Long, y As Long
  x = 1
  Debug.Print x / y
  Exit Sub
  
ErrorHandler:
  Debug.Print "Error description: ", Err.Description
  y = 5
  Resume Afterhandling
  
Afterhandling:
  Debug.Print "After handling Error"
  
End Sub

Sub MyWithStatement()
With Sheet1.Range("A1")
  .Value = 1000
  .Interior.Color = RGB(255, 255, 0)
  .Font.Bold = True
End With
End Sub





