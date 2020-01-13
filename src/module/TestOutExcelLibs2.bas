Attribute VB_Name = "TestOutExcelLibs2"
Option Explicit


Sub MyCorlWB()

' [_Default] is the default member, so you can omit it
' All return the same result

  Debug.Print Workbooks.Item(1).Name
  Debug.Print Workbooks.[_Default](1).Name
  Debug.Print Workbooks(1).Name

End Sub

Sub MyOpenRistrictedWB()

'Open Workbook locked read and write by password

Dim path As String: Let path = ThisWorkBook.path & "/RistrictRW.xlsx"
Dim pw As String: Let pw = "ban"

'W/o pass arguments
'Dim wb As Workbook: Set wb = Workbooks.Open(path)
'Debug.Print wb.Name

'W/ pass arguments
 With Workbooks.Open( _
    path, _
    Password:=pw, _
    WriteResPassword:=pw)
    
    Debug.Print .Name
  End With
End Sub

Sub MyUsingTWB()
'Check out workbook objects methods and properties, and get workbook infomation

  Dim twb As Workbook: Set twb = ThisWorkBook
  Debug.Print twb.Name
  
  Dim nl As String: nl = vbNewLine
  With twb
    Debug.Print .path; nl; .HasVBProject; nl; .Parent; nl; _
      .Password; nl; .Worksheets.Count; nl; _
      TypeName(.Worksheets)
      ' The Worksheets property doesn't return Worksheets collection, but Sheets collection
  
  End With
  
  ' This causes type mismatch error
  'Dim mysheets As Worksheets
  'Set mysheets = twb.Worksheets
  
End Sub

Sub MyPr()
  'PrintOut method doesn't allow mac users to preview. :(
  
  'ThisWorkBook.PrintOut From:=2, To:=4, Preview:=True
End Sub

Sub MyProtectWB()
  With ThisWorkBook
   .Protect "", Structure:=True
   Stop
   .Unprotect ""
  End With
End Sub

Sub MyGetEmbeddedProperties()
  On Error Resume Next
  Dim p As DocumentProperty
  For Each p In ThisWorkBook.BuiltinDocumentProperties
     Debug.Print p.Name; p.Value
  Next
End Sub

Sub MyCantCreateInstance()
  'Some class is not createble, in example, Workbook class
  
  'Dim wb As Workbook: Set wb = New Workbook
End Sub
