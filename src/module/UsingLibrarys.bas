Attribute VB_Name = "UsingLibrarys"
Option Explicit

Sub MySub()
  Dim text As String, words As Variant, w As Variant: text = "My name is Bob"
  words = Split(text)
  For Each w In words
    Debug.Print w
  Next
  
  'Debug.Print Join(words, ",")
  
  'When you want to check the data out,
  For Each w In Filter(words, "My")
    Debug.Print w
  Next
  
End Sub

Sub GenerateRnd()
  Dim x As Long: x = 1
  Dim y As Long: y = 6
  Randomize
  Dim i As Long
  For i = 1 To 5
    Debug.Print Int(Rnd * (y - x + 1) + x)
  Next i
  Dim s As String
  's = InputBox("Your name?")
  'Debug.Print s
End Sub

Sub CallByNameSamp1()

    Dim myRange As Range
    Dim myPrName(1 To 6) As String
    Dim i As Integer
    Dim myMsg As String

    myPrName(1) = "Address"
    myPrName(2) = "Top"
    myPrName(3) = "Left"
    myPrName(4) = "Value"
    myPrName(5) = "Formula"
    myPrName(6) = "Style"
    Set myRange = Selection

    For i = 1 To 6

        myMsg = myMsg & myPrName(i) & vbTab & _
            CallByName(myRange, myPrName(i), VbGet) & vbCr
    Next i

   MsgBox myMsg

End Sub

Sub MyCallByName()
Dim properties(1 To 3) As String
properties(1) = "FirstName"
properties(2) = "Gender"
properties(3) = "Age"

Dim p As Person: Set p = New Person
Dim i As Long
For i = LBound(properties) To UBound(properties)
  CallByName p, properties(i), VbLet, Sheet1.Cells(2, i).Value
Next i
Stop

  
End Sub


Function GetUserNameMac() As String
    Dim sMyScript As String

    sMyScript = "set userName to short user name of (system info)" & vbNewLine & "return userName"

    GetUserNameMac = MacScript(sMyScript)
    
End Function

Sub AutoNote()
Dim taskId As Long: taskId = Shell("/Applications/CotEditor.app/", vbNormalFocus)
'If you want to sendkeys things, then use Apple
End Sub
