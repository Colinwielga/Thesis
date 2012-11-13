VERSION 5.00
Begin VB.Form frmHerMattson3 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Question #1"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   Picture         =   "frmThree.frx":0000
   ScaleHeight     =   5895
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FF0000&
      Caption         =   "Peter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H000000FF&
      Caption         =   "Mark"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFF00&
      Caption         =   "James"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtQuestionOne 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Text            =   "Which Book Is A Gospel?"
      Top             =   360
      Width           =   5655
   End
End
Attribute VB_Name = "frmHerMattson3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AmazingQuiz
'frmHerMattson3
'Ee Her and Jennifer Mattson
'Written on 3/19/06
'This is the first question to the quiz.

Dim RightAnswer As Boolean
Dim Answer As String

Private Sub Form_Load()
'The timer starts and is counted in seconds.
frmHerMattson1.tmtTest.Enabled = True
frmHerMattson1.tmtTest.Interval = 1000
run = True
End Sub

Private Sub Option1_Click()
'This is a wrong answer, so the mxgbox will appear with a wrong answer and show the correct answer.
    RightAnswer = False
    Answer = MsgBox("Wrong Answer, the correct Answer was Mark", , "Wrong")
    frmHerMattson3.Hide
    frmHerMattson4.Show
End Sub

Private Sub Option2_Click()
'This is a right answer, so the mxgbox will appear with a congratulations and will add number to counter.
    RightAnswer = True
    Counter = Counter + 1
    Answer = MsgBox("Congratulations! You got the answer right!", , "Right Answer")
    frmHerMattson3.Hide
    frmHerMattson4.Show
End Sub

Private Sub Option3_Click()
    RightAnswer = False
    Answer = MsgBox("Wrong Answer, the correct Answer was Mark", , "Wrong")
    frmHerMattson3.Hide
    frmHerMattson4.Show
End Sub
