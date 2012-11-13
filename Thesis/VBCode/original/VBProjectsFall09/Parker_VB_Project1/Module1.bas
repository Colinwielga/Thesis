Attribute VB_Name = "Module1"
Option Explicit
Public UserName As String, price(1 To 6) As Single, ctr As Integer, priceOption(1 To 20) As Single, names(1 To 10) As String
Public autonames(1 To 30) As String, autoprices(1 To 30) As Single
Sub main()
'Open the file of picture names and put them in an array called names.
Open App.Path & "\picNames.txt" For Input As #1

ctr = 0

Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, autonames(ctr), autoprices(ctr)
Loop
Close #1

'display the startup form
frmFirst.Show

End Sub
