Attribute VB_Name = "UsingCollection_ErrObject"
Option Explicit

'   /////
'   Trying to use VBA Collection class to see how it works.
'   There are not many members.
'   Only purpose of using this is usingAssociative array
'   ////

Sub MyCollection()
  Dim Persons As Collection: Set Persons = New Collection
  With Persons:
    .Add "Bob", "m01"
    .Add "Tom", "m02"
    .Add "Ivy", "m03"
  
    Debug.Print .Count
    Debug.Print .Item(1)
    Debug.Print ("m02")
  End With
  
End Sub

'Made class that contains Collection with another class as a member
'allowing us to store a bunch of members' data into one class

Sub MyColClass()
  Dim myPersons As Persons: Set myPersons = New Persons
  myPersons.Add "00", "Mike", 20
  myPersons.Add "01", "Paul", 25
  Debug.Print myPersons.Items("00").Age
End Sub
