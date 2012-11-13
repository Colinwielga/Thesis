VERSION 5.00
Begin VB.Form frmSimpsons 
   BackColor       =   &H00FF0000&
   Caption         =   "Characters"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   Picture         =   "simpsons.frx":0000
   ScaleHeight     =   8850
   ScaleWidth      =   10875
   Begin VB.PictureBox picresults 
      BackColor       =   &H0000FFFF&
      FillColor       =   &H000000C0&
      ForeColor       =   &H000000C0&
      Height          =   4455
      Left            =   120
      ScaleHeight     =   4395
      ScaleWidth      =   10635
      TabIndex        =   3
      Top             =   240
      Width           =   10695
   End
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H000000C0&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8040
      TabIndex        =   2
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton cmdtest 
      Caption         =   "Take the Test"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4560
      TabIndex        =   1
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton cmdreturn 
      BackColor       =   &H8000000A&
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1440
      TabIndex        =   0
      Top             =   5160
      Width           =   1815
   End
End
Attribute VB_Name = "frmSimpsons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'simpsons show test (final.vbp)
'main form (Simpsons.frm)
'Jim Berg
'October 30, 2005
'this form contains the quiz
'it will ask the user the questions and display their results
'then it will store the score in a file to be accessed at another time
Dim question(1 To 10) As String, answer(1 To 10) As String, CTR As Integer, score As Integer
Dim k As Integer, j As Integer, results As String, N(1 To 10) As Integer, number As Integer

Private Sub cmdclear_Click()
    picresults.Cls
End Sub

Private Sub cmdreturn_Click()
    frmSimpsons.Hide
    frmmain.Show

End Sub

Private Sub cmdTest_Click()
    MsgBox "To get credit for answering a question you must spell the answer correctly. Also you must capitalize the first letter of your answer and any words such as names or titles that should be capitalized. Click Okay to begin the test", , "Instructions"
    score = 0
    CTR = 0
    'puts the test questions into an array
    Open App.Path & "\Simpsons2.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, question(CTR), answer(CTR)
    Loop
    'makes the quiz ask the questions in a random order
    For j = 1 To 10
Redo:
        number = (Int(Rnd * CTR + 1))
            For k = 1 To j
                If (number) = (N(k)) Then
                    GoTo Redo
                End If
            Next k
            N(j) = number
    
    theiranswer = InputBox(question(N(j)), "Simpsons")
     'decides whether the answer is correct or not
        If theiranswer = answer(N(j)) Then
            results = "Correct"
            score = score + 1
        Else
            results = "Incorrect"
        End If
    'displays the question, their answer, and the correct answer
    picresults.Print j; ".)", question(N(j))
    picresults.Print "Your Answer:", theiranswer, results, "Correct Answer", answer(N(j))
        Next j
    picresults.Print "Your Score: "; score
    'enter the users name
    player = InputBox("Please Enter Your Name")
        Close #1
    'store the name and score in a file
    Open App.Path & "\highscore2.txt" For Append As #1
        Print #1, player
        Print #1, score
    Close #1
    MsgBox score, , "Number of correct answers"
           
      
End Sub

Private Sub Form_Load()
    'allowing me to randomize the questions
    Randomize
End Sub
