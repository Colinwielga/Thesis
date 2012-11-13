Attribute VB_Name = "Module1"
Option Explicit
'Millionaire
'Code Module
'Angela Mathis and Katie Frazier
'2-10-2010
'The code module holds the public variables, and reads data into arrays.

'Declare variables as public and available to all forms in the project.
Public J As Integer, K As Integer
Public Question(1 To 20) As String, A(1 To 20) As String, B(1 To 20) As String
Public C(1 To 20) As String, D(1 To 20) As String, MoneyValues(1 To 20) As Single
Public Answer(1 To 20) As String, UserAnswer As String
Public CTR As Integer, CTR2 As Integer


Sub main()

'Initializes the counters and the variables used for the arrays to 0.
J = 0
K = 0
CTR = 0
CTR2 = 0

'Reads a data file that includes questions, multiple choice answers, and the correct answer.
    Open App.Path & "\Millionaire.txt" For Input As #1
        Do While Not EOF(1)
            CTR = CTR + 1
            Input #1, Question(CTR), A(CTR), B(CTR), C(CTR), D(CTR), Answer(CTR)
        Loop
    Close #1

'Reads a data file that includes the amount of money the user wins for each question.
    Open App.Path & "\MoneyValues.txt" For Input As #2
        Do While Not EOF(2)
            CTR2 = CTR2 + 1
            Input #2, MoneyValues(CTR2)
        Loop
    Close #2

'Brings the user to the Enter Game form.
    frmEnterGame.Show

End Sub




