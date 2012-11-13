VERSION 5.00
Begin VB.Form frmHerMattson6 
   BackColor       =   &H0080C0FF&
   Caption         =   "Question #4"
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option3 
      BackColor       =   &H008080FF&
      Caption         =   "Physiological "
      BeginProperty Font 
         Name            =   "Informal Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Self  Actualization"
      BeginProperty Font 
         Name            =   "Informal Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Safety"
      BeginProperty Font 
         Name            =   "Informal Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtQuestionFour 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Informal Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Text            =   "What is the most basic of Maslow's Hierarchy of Needs?"
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "frmHerMattson6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AmazingQuiz
'frmHerMattson6
'Ee Her and Jennifer Mattson
'This is the fourth quiz question.
'Written on 3/19/06

Dim RightAnswer As Boolean
Dim Answer As String

Private Sub Option1_Click()
    RightAnswer = False
    Answer = MsgBox("Wrong Answer, the correct Answer was Physiological", , "Wrong")
    frmHerMattson6.Hide
    frmHerMattson7.Show
End Sub

Private Sub Option2_Click()
    RightAnswer = False
    Answer = MsgBox("Wrong Answer, the correct Answer was Physiological", , "Wrong")
    frmHerMattson6.Hide
    frmHerMattson7.Show
End Sub

Private Sub Option3_Click()
    RightAnswer = True
    Counter = Counter + 1
    Answer = MsgBox("Congratulations! You got the answer right!", , "Right Answer")
    frmHerMattson6.Hide
    frmHerMattson7.Show
End Sub
