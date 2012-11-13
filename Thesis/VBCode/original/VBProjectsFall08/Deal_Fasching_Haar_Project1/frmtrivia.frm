VERSION 5.00
Begin VB.Form FrmTrivia 
   BackColor       =   &H80000008&
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11415
   ForeColor       =   &H80000006&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picresults 
      Height          =   6495
      Left            =   4320
      ScaleHeight     =   6435
      ScaleWidth      =   6795
      TabIndex        =   4
      Top             =   1680
      Width           =   6855
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear Results"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5760
      MaskColor       =   &H000000FF&
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdbacktoeasthigh 
      Caption         =   "Back to East High"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8520
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton cmdhsm2 
      Caption         =   "Trivia for HSM 2"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdhsm1 
      BackColor       =   &H80000004&
      Caption         =   "Trivia for HSM 1"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   5850
      Left            =   0
      Picture         =   "frmtrivia.frx":0000
      Top             =   1920
      Width           =   4320
   End
End
Attribute VB_Name = "FrmTrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Project Name: High School Musical
' Form name: Name that Tune
' Author: Laura Deal, Megan Haar, Kirsten Fasching
' Date Written: 10/28/08
'Objective: This program quizes the users of their high school musical knowledge.
 'objective cont' There are two quizes for the first two high school musical movies.
 ' objective cont' A counter is used to display the end score of the trivia quiz.

Private Sub cmdbacktoeasthigh_Click()
'brings the user back to the buttons page where they can choose another activity or leave
frmauthors.Hide
frmbuttons.Show
Frmcharacter.Hide
frmtitle.Hide
frmnamethattune.Hide
FrmTrivia.Hide
frmquiz.Hide
frmtitle.Hide
End Sub

Private Sub cmdclear_Click()
'clears the results box
picresults.Cls
End Sub

Private Sub cmdhsm1_Click()
Dim question(1 To 10) As String
Dim answer(1 To 10) As String
Dim inputanswer As String
Dim pos As Integer
Dim ctr As Integer
Dim score As Integer
Dim N As Single

ctr = 0
score = 0

picresults.Cls

'open HSM1 file
Open App.Path & "\hsm1trivia.txt" For Input As #1

Do Until EOF(1)
    ctr = ctr + 1
    Input #1, question(ctr), answer(ctr)
Loop
Close #1

'start quiz for HSM1
For pos = 1 To ctr
    inputanswer = InputBox(question(pos), "Question")
        If LCase(inputanswer) = LCase(answer(pos)) Then
            score = score + 1
        Else
            MsgBox "Sorry, the right answer is: " & answer(pos) & ".", , "OOPS!"
        End If
Next pos

'print results for HSM1 quiz
Select Case score
Case Is >= 9
    picresults.Print "You got "; score; "out of 10 correct! You're AmAzInG!"
Case Is >= 7
    picresults.Print "You got "; score; "out of 10 correct! Good Job!"
Case Is >= 5
    picresults.Print "You got "; score; "out of 10 correct! Keep Trying!"
Case Is >= 3
    picresults.Print "You got "; score; "out of 10 correct! Maybe you'll do better next time.."
Case Is >= 0
    picresults.Print "You got "; score; "out of 10 correct! Go watch HSM1 and try again!"
End Select

'print the questions and answers so the user can 'study' for next time
picresults.Print "Questions and Answers are as follows:"

For N = 1 To ctr
picresults.Print
picresults.Print N; question(N)
picresults.Print answer(N)
Next N

End Sub

Private Sub cmdhsm2_Click()
Dim question(1 To 10) As String
Dim answer(1 To 10) As String
Dim inputanswer As String
Dim pos As Integer
Dim ctr As Integer
Dim score As Integer
Dim N As Single

ctr = 0
score = 0

picresults.Cls

'open HSM2 file
Open App.Path & "\hsm2trivia.txt" For Input As #1

Do Until EOF(1)
    ctr = ctr + 1
    Input #1, question(ctr), answer(ctr)
Loop
Close #1

'start quiz for HSM2
For pos = 1 To ctr
    inputanswer = InputBox(question(pos), "Question")
        If LCase(inputanswer) = LCase(answer(pos)) Then
            score = score + 1
        Else
            MsgBox "Sorry, the right asnwer is: " & answer(pos) & ".", , "OOPS!"
        End If
Next pos

'print results
Select Case score
Case Is >= 9
    picresults.Print "You got "; score; "out of 10 correct! You're AmAzInG!"
Case Is >= 7
    picresults.Print "You got "; score; "out of 10 correct! Good Job!"
Case Is >= 5
    picresults.Print "You got "; score; "out of 10 correct! Keep Trying!"
Case Is >= 3
    picresults.Print "You got "; score; "out of 10 correct! Maybe you'll do better next time.."
Case Is >= 0
    picresults.Print "You got "; score; "out of 10 correct! Go watch HSM2 and try again!"
End Select

'print questions and answers so the user can 'study' for next time
picresults.Print "Questions and Answers are as follows:"

For N = 1 To ctr
picresults.Print
picresults.Print N; question(N)
picresults.Print answer(N)
Next N

End Sub


