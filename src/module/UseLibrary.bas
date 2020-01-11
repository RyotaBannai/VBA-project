Attribute VB_Name = "UseLibrary"
Option Explicit

Sub MySub()
  VBA.Interaction.Beep
  Debug.Print VBA.DateTime.DateSerial(2000, 10, 0)
  
End Sub

Sub MyDic()
  Dim dict As New Dictionary
  dict.Add "Apple", 50
  dict.Add "Orenge", 99
  dict.Item("Apple") = 111
  
  Dim vKeys
  Dim vKey
  vKeys = dict.Keys

  For Each vKey In vKeys
    ' if you use vKey as artument of Item, then Compile error: ByRef argument type mismatch
    ' So assign the key to variable first
    Dim k As String: k = vKey
    Debug.Print k & " = " & dict.Item(k)
  Next
  
  Dim i As Long
  For i = LBound(vKeys) To UBound(vKeys)
    Debug.Print vKeys(i), dict.Items(i)
  Next i
  
End Sub

Sub LateBinding()
  'Dim doc As Object
  'Set doc = GetObject(, "Word.Application")
  'Debug.Print doc.Name
End Sub
