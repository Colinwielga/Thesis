VERSION 5.00
Begin VB.Form frmtest 
   BackColor       =   &H00800000&
   Caption         =   "Garrett Sohn"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   1215
      Left            =   8040
      TabIndex        =   3
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdplay 
      Caption         =   "Play"
      Height          =   1095
      Left            =   8040
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "Back to Main menu"
      Height          =   1095
      Left            =   8040
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H000000FF&
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3555
      ScaleWidth      =   7755
      TabIndex        =   0
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "frmtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'test form (test.frm)
'Garrett Sohn
'March 24, 2006
'this gives the user multiple questions based on what has they have seen in the other forms.  It allows them to input an answer and know if they are correct or not.
Option Explicit
Private Sub cmdclear_Click()
    picresults.Cls
End Sub

Private Sub cmdmain_Click()
    frmtest.Hide
    frmmadness.Show
End Sub

Private Sub cmdplay_Click()
Dim answer(1 To 8) As String, Pos As Integer, question(1 To 8) As String, score As Integer, player As String
Dim m As Integer, B As Integer, N(1 To 8) As Integer, results As String, number As Integer, theiranswer As String
MsgBox "To get the right answer you have to spell the player's name and team right with the correct capitalization. Click Okay to begin the test", , "Instructions"
score = 0
Pos = 0
'puts the test questions into an array
Open App.Path & "\questions.txt" For Input As #1
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, question(Pos), answer(Pos)
    Loop
    For B = 1 To 8
Redo:
        number = (Int(Rnd * Pos + 1))
            
            For m = 1 To B
                If (number) = (N(m)) Then
                    GoTo Redo
                End If
            Next m
            N(B) = number
    
    theiranswer = InputBox(question(N(B)), "NCAA")
     'this decides if the answer is right or wrong
        If theiranswer = answer(N(B)) Then
            results = "That is correct"
            score = score + 1
        Else
            results = "That is not right"
        End If
    'this displays the question, what they answered, and what the correct answer is
    picresults.Print B; ".)", question(N(B))
    picresults.Print "Your Answer:", theiranswer, results, "Correct Answer", answer(N(B))
        Next B
    picresults.Print "Your Score: "; score
    'this puts in the user name
    player = InputBox("Please Enter Your Name")
        Close #1
    'store the name and score in a file
    MsgBox score, , "Number of correct answers"
           
End Sub

