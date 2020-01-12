Attribute VB_Name = "TestOutExcelLibs"
Option Explicit

Sub MyExcelLib()
  'Dim v As variable: Set v =
  Debug.Print TypeName(Selection) ' TypeName
  Debug.Print Selection.Address
End Sub

'Change the background color of cells found with Intersect and Union methods
Sub FillColor()
  With Sheet1
    Dim rng1 As Range: Set rng1 = .Range("C1:E5")
    Dim rng2 As Range: Set rng2 = .Range("B2:F4")
    Dim rng3 As Range: Set rng3 = .Range("A3:G3")
  End With
  
  With Union(rng1, rng2, rng3)
  .Select
      Sheet1.Range(.Address).Interior.ColorIndex = 43 'Green
  End With
  
  With Intersect(rng1, rng2, rng3)
  .Select
      Sheet1.Range(.Address).Interior.ColorIndex = 37 'Blue
  End With
End Sub

'Ristrict user inputs in the value of cells
Sub MyInputMethod()
  With Sheet1
    .Range("A1").Value = Application.InputBox("Input hours you work in numbers:", Type:=1)
  End With
End Sub

'Test wait and ontime methods
Sub MyWait()
  Application.OnTime Now + TimeSerial(0, 0, 3), "ShowMessage" ' Set procedures in String
End Sub
Sub ShowMessage()
  MsgBox "It's time!"
  Application.Wait Now + TimeSerial(0, 0, 3) ' wait 3 minutes
  MsgBox "It's 3 minutes."
End Sub

Sub UseWSFun()
  With Sheet1
    Dim rng As Range: Set rng = .Range("A6:A10")
    Debug.Print WorksheetFunction.max(rng)
    Debug.Print WorksheetFunction.Min(rng)
  End With
End Sub

Sub MyEval()
  'You can use Evaluate to use worksheet function
  'as well as sell range, fuctions
  '[] is the short cut of evaulate
  Evaluate("B2").Value = 123
  [B3].Value = 456
  
  Debug.Print [Max(A6:A10)]
  Debug.Print [Fuga].Address ' this is awesome
  Debug.Print TypeName([Mike]) ' Mike's type is Rectangle
  
End Sub
