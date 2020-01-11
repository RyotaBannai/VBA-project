Attribute VB_Name = "Module2"
Option Explicit

Sub MySub_()
    price = 100
    Debug.Print price, TaxIncluded
End Sub

Sub mySub()
  x = 123: Debug.Print "Variabel x ", x
  Debug.Print "Const MSG : ", msg
  Debug.Print "Enum e: ", e.Todohuken, e.Kenchoshozaiti
  
  Call VBAProject2.mySub
  Call VBAProject2.MySub2
  
  'To use Class from another project, you need Procedure in the same project that gest instance for you
  Dim c As Class1: Set c = CreateClass1("Bob")
  Debug.Print c.FirstName
  c.Greet
  
  End Sub
