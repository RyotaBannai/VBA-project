Attribute VB_Name = "ControlingTableObj"
Option Explicit

Sub MyCtrlTable()
  Debug.Print Sheet2.ListObjects.Count '1
  MyCopyTable sheet:=Sheet2, inx:=1, cell:="B10"
End Sub

Sub MyCreateDelete()
  'MyAddTable "Mytable1", "B10"
  MyDeleteTable 1
End Sub

Sub MyCopyTable(ByRef sheet As Worksheet, ByVal inx As Long, ByVal orientation As String)
  With sheet
    Dim v As Variant
    v = .ListObjects(inx).Range.Value
    .Range(orientation).Resize(UBound(v), UBound(v, 2)).Value = v
  End With
End Sub

'If the range of cells is already table, then cause an error
'
'
Sub MyAddTable(ByVal name As String, ByVal orientation As String)
  With Sheet2
    With .ListObjects.Add(Source:=.Range(orientation).CurrentRegion, XlListObjectHasHeaders:=xlNo)
      .name = name
      .TableStyle = "TableStyleLight9"
    End With
  End With
End Sub

'If the range of cells is not table, then cause an error
'
'
Sub MyDeleteTable(ByVal index As Long)
  With Sheet2
    If .ListObjects.Count = 0 Then
      MsgBox "Table NOT Exists."
    Else
      With .ListObjects(index)
        .Unlist
      End With
    End If
  End With
End Sub

'Add new data into row
'you can ommit potision. if so, then add method inserts the data at the last
'
Sub MyInsertData()
  With Sheet2.ListObjects(1).ListRows
    Dim r As ListRow: Set r = .Add(Position:=2)
    r.Range.Value = Array("Ivy", "24", "Go", "Male")
  End With
End Sub
