Attribute VB_Name = "Module1"
Public wolfpics(1 To 5) As String
Public ctr As Integer

Sub main()

'Open file with slide show pictures and put them in an array
Open App.Path & "\slideshow.txt" For Input As #1
'set variable
ctr = 0
'load the array
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, wolfpics(ctr)
Loop
Close #1

'display the first form
frmpart1.Show

End Sub


