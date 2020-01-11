VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private firstName_ As String

Public Property Let FirstName(ByVal newName As String)
  If firstName_ = "" Then firstName_ = newName
  Debug.Print firstName_
End Property

Public Function Greet()
  Dim msg As String
  msg = "Hey! " & firstName_ & ", how are you?"
  Greet = MsgBox(msg, vbYesNo)
End Function
