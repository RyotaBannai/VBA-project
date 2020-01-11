VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetsCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private count_ As Long
Private Const MAX_COUNT As Long = 2

Public Event Over(ByVal value As Long, ByVal max As Long)

Public Property Let Count(ByVal newCount As Long)
  'Debug.Print "Property Count Called"
  count_ = newCount
  If count_ > MAX_COUNT Then RaiseEvent Over(count_, MAX_COUNT)
End Property
