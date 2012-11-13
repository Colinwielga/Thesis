Attribute VB_Name = "Module1"
Option Explicit

'The Artist's Multimedia Portfolio
'Module 1
'Ashley Thompson
'Friday March 20, 2009
'This module reads the primary file and places pictures and corresponding information into multiple arrays
'The ctr reads the number of objects in this file
'The frmMain is made visible to begin the program


Public Art(1 To 25) As String
Public Year(1 To 25) As Integer
Public Medium(1 To 25) As String
Public Draft(1 To 25) As String

Public ctr As Integer

Sub main()

Open App.Path & "\AshleyPaintings.txt" For Input As #1

ctr = 0

Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Art(ctr), Year(ctr), Medium(ctr), Draft(ctr)
    
Loop
Close #1

frmMain.Show

End Sub

