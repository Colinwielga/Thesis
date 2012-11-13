VERSION 5.00
Begin VB.Form frmGame 
   BackColor       =   &H80000005&
   Caption         =   "Game                                                                                                               "
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H80000005&
      Caption         =   "Submit Answer"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5520
      Width           =   2775
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H80000009&
      Caption         =   "Quit "
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6480
      Width           =   2775
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H80000009&
      Caption         =   "Back To Main Menu"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6480
      Width           =   2775
   End
   Begin VB.PictureBox pic6 
      Height          =   1575
      Left            =   14040
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   18
      Top             =   9360
      Width           =   1335
   End
   Begin VB.PictureBox pic5 
      Height          =   1575
      Left            =   14040
      Picture         =   "Form1.frx":7FC2
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   17
      Top             =   7800
      Width           =   1335
   End
   Begin VB.PictureBox pic4 
      Height          =   1455
      Left            =   14040
      Picture         =   "Form1.frx":F8AC
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   16
      Top             =   6360
      Width           =   1095
   End
   Begin VB.PictureBox pic3 
      Height          =   1815
      Left            =   14040
      Picture         =   "Form1.frx":1690E
      ScaleHeight     =   1755
      ScaleWidth      =   1155
      TabIndex        =   15
      Top             =   4560
      Width           =   1215
   End
   Begin VB.PictureBox pic2 
      Height          =   1575
      Left            =   14040
      Picture         =   "Form1.frx":23D30
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   14
      Top             =   3000
      Width           =   1335
   End
   Begin VB.PictureBox Picture8 
      Height          =   1575
      Left            =   14040
      ScaleHeight     =   1515
      ScaleWidth      =   1155
      TabIndex        =   12
      Top             =   1440
      Width           =   1215
      Begin VB.PictureBox picTim2 
         Height          =   1695
         Left            =   0
         Picture         =   "Form1.frx":2A3D2
         ScaleHeight     =   1635
         ScaleWidth      =   2235
         TabIndex        =   13
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H80000004&
      Caption         =   "Play !!"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5520
      Width           =   2775
   End
   Begin VB.TextBox txtAnswer 
      Height          =   1695
      Left            =   3720
      TabIndex        =   10
      Top             =   3360
      Width           =   7455
   End
   Begin VB.PictureBox picQuestion 
      Height          =   1575
      Left            =   2880
      ScaleHeight     =   1515
      ScaleWidth      =   8955
      TabIndex        =   9
      Top             =   1320
      Width           =   9015
   End
   Begin VB.PictureBox Picture7 
      Height          =   1695
      Left            =   0
      Picture         =   "Form1.frx":31794
      ScaleHeight     =   1635
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   5760
      Width           =   1335
   End
   Begin VB.PictureBox Picture6 
      Height          =   1455
      Left            =   0
      Picture         =   "Form1.frx":3DEE6
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   7
      Top             =   7440
      Width           =   1455
   End
   Begin VB.PictureBox Picture5 
      Height          =   1215
      Left            =   0
      Picture         =   "Form1.frx":465BC
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   6
      Top             =   4560
      Width           =   1335
   End
   Begin VB.PictureBox Picture4 
      Height          =   2055
      Left            =   0
      Picture         =   "Form1.frx":4D05E
      ScaleHeight     =   1995
      ScaleWidth      =   1395
      TabIndex        =   5
      Top             =   8880
      Width           =   1455
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   0
      Picture         =   "Form1.frx":56520
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   0
      Picture         =   "Form1.frx":5CBC2
      ScaleHeight     =   1635
      ScaleWidth      =   1275
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.PictureBox picScore 
      Height          =   1335
      Left            =   5640
      ScaleHeight     =   1275
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   9600
      Width           =   3735
   End
   Begin VB.Label lblScore 
      Caption         =   "                Score"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   9240
      Width           =   4815
   End
   Begin VB.Label lblHeader 
      Caption         =   "  Pop Culture Trivia "
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   735
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   9735
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nextquestion As Integer
Dim counter As Integer, Question(1 To 20) As String, Questions As String
Dim Answer(1 To 20) As String, Answers As String, Names As String, score As Single
Dim pos As Integer
    
'Pop Culture Trivia. (PopCultureTrivia.vbp)
'Form Name: frmGame
'Author: Megan Zinken
'Date Written: November 2nd, 2006
'Form Objective: The purpose of this form is to allow the user to play
                 '"Pop Culture Trivia." When the play button is clicked the
                 'First question is displayed and the user answers via text box
                 'It allows the user to keep score as well as go back to the first form
                 



Private Sub cmdBack_Click()
    'This button allows the user to go back to the first form
    'frmGame becomes invisible
    
    frmGame.Visible = False
    frmPopCultureTrivia.Visible = True
    
End Sub

Private Sub cmdExit_Click()
    End
    
End Sub

Private Sub cmdPlay_Click()
   'This button allows the user to play "Pop Culture Trivia" by displaying the questions.
   'The questions and answers are filed into an array
   
    counter = 0
    nextquestion = nextquestion + 1
    Names = txtAnswer.Text
    Open App.Path & "\Game.txt" For Input As #1
    Do Until EOF(1)
        Input #1, Questions, Answers
        counter = counter + 1
        Question(counter) = Questions
        Answer(counter) = Answers
Loop
Close #1
If nextquestion < counter Then
picQuestion.Cls
picQuestion.Print Question(nextquestion)
Else
MsgBox "End of Game"
End If




End Sub


'This button checks the answer inputed by the user with the answer filed into an array
'If the user gets the correct answer then a message box "Congrats" is displayed
'If the user gets the question wrong a message box displays the correct answer

Private Sub cmdSubmit_Click()
Dim found As Boolean
Answers = txtAnswer.Text
found = False
If Answers = Answer(nextquestion) Then
    found = True
    score = score + 100
    picScore.Cls
    picScore.Print score
    MsgBox ("Congrats!")
    End If
If found = False Then
    MsgBox ("I'm Sorry That is Incorrect, The Correct Answer is " & Answer(nextquestion))
End If
End Sub

