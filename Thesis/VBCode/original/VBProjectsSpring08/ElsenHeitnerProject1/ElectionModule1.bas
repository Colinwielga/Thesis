Attribute VB_Name = "Module1"
'PROJECT: Choose or Lose: Election Perfection
'FORM: Module1(ElectionMondule1.bas)
'AUTHOR:  Nick Elsen and Andrew Heitner
'DATE:  March 26, 2008
'PURPOSE:  This provides the public counters we need to count the users views and come to a final cantidate.

Public UsersName As String
Public CantidateCtr(1 To 4) As Integer
Public Cantidate(1 To 4) As String

'Sets all of the variables to a slot in the array or sets variables that will be used in multiple forms
Sub main()
UsersName = "notset"
Cantidate(1) = "Barack Obama"
Cantidate(2) = "Hillary Clinton"
Cantidate(3) = "John McCain"
Cantidate(4) = "Mike Huckabee"
frmIntro.Show
End Sub




