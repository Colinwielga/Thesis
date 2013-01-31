Attribute VB_Name = "Module1"
Public remainder As Integer, Names(1 To 12) As String, Names2(1 To 12) As String, Names3(1 To 12) As String, zodiac(1 To 12) As String, num(1 To 12) As Integer
Public score As Integer

'remainder tells the program the number of the desired zodiac, and will be used in multiple forms.
'zodiac(1 to 12) are the names of the zodiacs
'num(1 to 12) are the numbers of the zodiacs.
'Names(1 to 12) are the names of the first set of pictures,and this array would be used in multiple forms.
'Names2(1 to 12) are the names of the second set of pictures, and will be used in multiple forms..
'Names3(1 to 12) are the names of the text files, and will be used in 2 forms.
'Targetzodiac is the name of the zodiac the users is looking for, targetnumber is its number.
'score is the score on trivia

Sub main()
Dim I As Integer
Open App.Path & "\picNames.txt" For Input As #1
For I = 1 To 12
    Input #1, Names(I)
Next I
Close #1
Open App.Path & "\Zodiac.txt" For Input As #2
I = 9999
    For I = 1 To 12                             'I know there will be only 12 lines of data, so I used the for-next statement instead of do-while-loop
    Input #2, zodiac(I), num(I)
    Next I
Close #2
Open App.Path & "\picNames2.txt" For Input As #3
    For I = 1 To 12
    Input #3, Names2(I)
    Next I
Close #3
Open App.Path & "\commentsNames.txt" For Input As #4
    For I = 1 To 12
    Input #4, Names3(I)
    Next I
Close #4
Home.Visible = True
End Sub
