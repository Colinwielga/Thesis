Attribute VB_Name = "Module1"
Option Explicit
Public names(1 To 1) As String
Public ctr As Integer

Sub main()
'Open the file of picture names and put them in an array called names.
Open App.Path & "\PictureNames.txt" For Input As #1


ctr = 0

Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, zone(ctr)
Loop


End Sub

