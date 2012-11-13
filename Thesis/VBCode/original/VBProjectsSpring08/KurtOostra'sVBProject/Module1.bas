Attribute VB_Name = "Module1"
'Project Name: Theater Lighting
'Form Name: Module1.bas
'Author: Kurt Oostra
'Date Written:3/11/08
'Objective: Load an array as soon as you hit play
'           this array can be used in any form
'this entire project is designed as a learning experience for beginners and people
'who are already knowledgeable in  theater. you can learn about mostly lights, but
'also venues at CSB/SJU where they would most likely be used.
Option Explicit
Public names(1 To 7) As String, watts(1 To 7) As Single, use(1 To 7) As String, picnames(1 To 7) As String
Public ctr As Integer
Sub main()
Open App.Path & "\lightinfo.txt" For Input As #1
    Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, names(ctr), watts(ctr), use(ctr)
Loop
Close #1
'opens the main menu
frmMainMenu.Show
End Sub

