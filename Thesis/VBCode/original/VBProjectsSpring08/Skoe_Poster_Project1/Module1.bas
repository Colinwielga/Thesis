Attribute VB_Name = "Module1"
'Super Smash Bros.
'Module 1
'Ryan Poster and Erik Skoe
'March 26th
Option Explicit
Public Ctr As Integer
Public Pics(1 To 20) As String
Public characters(1 To 20) As String
Public Char As Integer
Sub Main()

Open App.Path & "\characters.txt" For Input As #1   'This is going to happen automatically and it will load the names of the characters and their pictures into an array.

Do While Not EOF(1)
Ctr = Ctr + 1
Input #1, characters(Ctr), Pics(Ctr)
Loop
Close (1)
Opening.Show

End Sub
