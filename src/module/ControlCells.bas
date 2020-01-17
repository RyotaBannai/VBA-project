Attribute VB_Name = "ControlCells"
Option Explicit

Sub MyCrlCells()
  With Sheet2
    Debug.Print .Range("A1", "E3").Address
    Debug.Print .Range(.Range("A1"), .Range("E3")).Address
        
  End With
    
End Sub

Sub MySelectSheets()
  Sheet2.Activate
  Stop
  Sheet3.Activate
  Stop
  ' Selecting all sheets
  ' Select and activate are different concepts
  ' You can activate only one sheet
  ThisWorkBook.Sheets.Select
  Stop
  ' select 2 sheets from the left
  ThisWorkBook.Sheets(Array(1, 2)).Select
  
  
End Sub

Sub MyVisibleSheets()
  ' can't activate invisible sheets
  
  Worksheets("Sheet2").Visible = False
  Stop
  'Sheets collection's index start from 0
  ThisWorkBook.Sheets(1).Visible = True
End Sub

Sub MyProtectSheets()
  With Sheet1
    ' Protect sheets from editing cells
    
    .Protect Password:="ban", UserInterfaceOnly:=True
    Stop
    
    ' UserInterfaceOnly := Ture allows editing from Macro only
    
    .Range("A1").Value = "Editing from Macro only."
    .Unprotect "ban"
    End With
    
End Sub

Sub MyCheckTab()
 
 ' Determine if color index of 1st tab is set to none.
 ' the 1st tab is on the left
 
 If Worksheets(1).Tab.ColorIndex = xlColorIndexNone Then
  MsgBox "The color index is set to none for the 1st " & _
 "worksheet tab."
 Else
  MsgBox "The color index for the tab of the 1st worksheet " & _
 "is not set none."
 End If
 
End Sub



