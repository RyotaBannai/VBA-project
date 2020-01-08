VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private WithEvents c As SheetsCounter
Attribute c.VB_VarHelpID = -1

Private Sub Workbook_NewSheet(ByVal Sh As Object)
  'Debug.Print "Called"
  If c Is Nothing Then Set c = New SheetsCounter
  c.Count = Worksheets.Count
End Sub

Private Sub c_Over(ByVal value As Long, ByVal max As Long)
  'Debug.Print "Over Called"
  Dim msg As String
  msg = ""
  msg = msg & "Current Sheet count: " & value & vbNewLine
  msg = msg & "You shouldn't add more than " & max
  MsgBox msg
End Sub
