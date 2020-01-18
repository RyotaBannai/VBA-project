Attribute VB_Name = "GettingCells"
Option Explicit

' Getting cells one by one w/ For Each In
' Can use index to get
' Such as .Cells(index1, index2), .Rows(index1), .Cells("1","A")

Sub MyGetCells()
  Dim r As Range
  Debug.Print "Showing cells in the area" & vbNewLine & "-----"
  With Sheet2.Range("A1:C2")
    For Each r In .Cells
      Debug.Print "Address: "; r.Address; ", Value: "; r.Value
    Next r
  End With
End Sub

Sub MyGetRows()
  Dim r As Range
  With Sheet2.Range("A1:C2")
    Debug.Print "Showing rows in the area" & vbNewLine & "-----"
    For Each r In .Rows
      Debug.Print r.Address
    Next r
  End With
End Sub

' Get the whole selected areas and read each cells
'
'
Sub MyUsufulMethods()
  With Sheet2.Range("B4")
    '
    ' Select an area contain B4 cell
    .CurrentRegion.Select
    
    Dim a As Range: Set a = Range(Selection.Address)
    Dim r As Range:
    
    For Each r In a.Cells: 'a.address $B$4:$D$7
      Debug.Print VarType(r)
      Debug.Print r.Address
      Debug.Print r.Value
    Next r
    
    '
    'Resize the area
    
    Dim b As Range: Set b = a.Resize(2, 3)
    Debug.Print b.Address ' $B$4:$D$5
    
  End With
End Sub

Sub MyGetEnd()
  '
  '
  ' Get the last row and column number
  ' You know the first number of the area,
  ' But you don't know the last number.
  
  With Sheet2
    Debug.Print .Cells(.Rows.Count, "B").End(xlUp).Row
    Debug.Print .Cells(4, .Columns.Count).End(xlToLeft).Column
  End With
End Sub

Sub MyGetSC()
  '
  'Get special cell, set type
  'xlCellTypeBlanks 'get blank cell
  'xlCellTypeFormulas ' get cell with formula inside
  
  With Sheet2.Range("B4:D7")
    Debug.Print .SpecialCells(xlCellTypeLastCell).Address
  End With
End Sub

Sub MyDel()
  '
  ' Delete an entire row even if you delete selected range
  
  With Sheet2
    '.Range("B6:D6").Delete
    .Rows("6").Delete
    
  End With
End Sub

Sub MyFindValues()
  With Sheet2.Range("B4").CurrentRegion
    Dim m As Range: Set m = .Find(What:="Ste", LookIn:=xlValues, _
      LookAt:=xlPart, MatchCase:=False, MatchByte:=True)
    If Not m Is Nothing Then
      Dim firstAdd As String: firstAdd = m.Address
      Do
        Debug.Print m.Value,
        '
        ' the way immediate window displays results are different from when
        ' you use "," or ";", or neither of them.
        '
        Set m = .FindNext(m)
      Loop While m.Address <> firstAdd
    End If
   End With
  
End Sub

Sub MyReplceValues()
  With Sheet2.Range("B4").CurrentRegion
    '.Replace What:="male", Replacement:="MALE"
    
    .Replace What:="male", Replacement:="Male"
  End With
End Sub
