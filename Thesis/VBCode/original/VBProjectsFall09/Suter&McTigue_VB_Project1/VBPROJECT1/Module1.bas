Attribute VB_Name = "Module1"
Public names(1 To 100) As String
Public ctr As Integer
Function getLoadedCount() As Double
    Dim programINI As String, countString As String, countNum As Double
    programINI = App.Path & App.EXEName & "_info.ini"
    
    ' checks for ini if not create the file and set the
    ' ini to the number of times this program has been loaded to 1
    If Len(Dir(programINI)) = 0 Then
        ' this is a if then to keep track of the number of times the program is opened
        Open programINI For Output As #1
        Print #1, "Times Loaded: 1"
        Close #1
        countNum = 1
    Else
        ' this then finds how many times the program has been opened
        If FileLen(programINI) <> 0 Then
            Open programINI For Input As #1
            Line Input #1, countString
            Close #1
        End If
        ' checks format
        If Len(countString) < 15 Then
            countNum = 1
        Else
            countNum = Val(Mid(countString, 14)) + 1
        End If
        Open programINI For Output As #1
        Print #1, "Times Loaded: " & countNum
        Close #1
    End If
    
    getLoadedCount = countNum
End Function


Sub Main()
'this opens an array before the project starts and grabs the names of the pictures from
'a text box in order to grab those pictures later
Open App.Path & "\Soccerpics.txt" For Input As #2

ctr = 0

Do While Not EOF(2)
    ctr = ctr + 1
    Input #2, names(ctr)
Loop
Close #2
'this brings you to the entry page
frmentryform.Show
End Sub


