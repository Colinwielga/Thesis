Attribute VB_Name = "Module1"
'SKI TRIP'
'MODULE'
'MAX TUSA'
'10-18'
'THIS IS THE MODULE THAT LOADS THE SKI RESORTS ARRAY'
Option Explicit

Public resorts(1 To 50) As String, skiruns(1 To 50) As Single, ctr As Integer, totalCost As Single
Public skiticketcost As Single, runningtotal As Single, totalairfarecost As Single
Public totalhotelcost As Single, place As Integer
Sub Main()


'import the file with the ski resorts and create an array'
Open App.Path & "\skiresorts.txt" For Input As #1

ctr = 0

Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, resorts(ctr)
    Input #1, skiruns(ctr)
Loop
Close #1

Title.Show

End Sub
