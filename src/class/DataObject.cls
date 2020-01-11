VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "DateObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public value As Date



Public Sub MsgDate()
  MsgBox value
End Sub
Public Function SetDate(ByVal y As Long, ByVal m As Long, d As Long) As Date
  value = DateSerial(y, m, d)
  SetDate = value
End Function
Public Property Get EndOfMonth() As Date
  'When you set 0 at DAY, then this returns end of last Month
  EndOfMonth = DateSerial(Year(value), Month(value) + 1, 0)
End Property

Private Sub Class_Initialize()
  value = Now
End Sub

Private Sub Class_Terminate()
  Debug.Print "Class Termination:" & value
End Sub



