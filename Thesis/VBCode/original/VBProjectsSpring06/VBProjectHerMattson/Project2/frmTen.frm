VERSION 5.00
Begin VB.Form frmHerMattson10 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Question #8"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFF80&
      Caption         =   "When you don't know the size of your file"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FF80FF&
      Caption         =   "When you know your file size"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080FF80&
      Caption         =   "Never"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   360
      TabIndex        =   0
      Text            =   "When should you use Do Loops?"
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "frmHerMattson10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AmazingQuiz
'frmHerMattson10
'Ee Her and Jennifer Mattson
'Written on 3/18/06
'This is question eight in our quiz.

Dim RightAnswer As Boolean
Dim Answer As String
Private Sub Option1_Click()
    RightAnswer = False
    Answer = MsgBox("Wrong Answer, the correct Answer was When you don't know the size of your file", , "Wrong")
    frmHerMattson10.Hide
    frmHerMattson1.Show
    frmHerMattson1.tmtTest.Enabled = False
    run = False
End Sub

Private Sub Option2_Click()
    RightAnswer = False
    Answer = MsgBox("Wrong Answer, the correct Answer was When you don't know the size of your file", , "Wrong")
    frmHerMattson10.Hide
    frmHerMattson1.Show
    frmHerMattson1.tmtTest.Enabled = False
    run = False
End Sub

Private Sub Option3_Click()
    RightAnswer = True
    Counter = Counter + 1
    Answer = MsgBox("Congratulations! You got the answer right!", , "Right Answer")
    frmHerMattson10.Hide
    frmHerMattson1.Show
    frmHerMattson1.tmtTest.Enabled = False
    run = False
End Sub
