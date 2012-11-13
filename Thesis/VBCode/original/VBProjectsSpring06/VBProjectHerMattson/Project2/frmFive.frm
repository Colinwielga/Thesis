VERSION 5.00
Begin VB.Form frmHerMattson5 
   Caption         =   "Question #3"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   Picture         =   "frmFive.frx":0000
   ScaleHeight     =   4305
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option3 
      Caption         =   "Dorthy Day"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Margaret Thatcher"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Eleanor Roosevelt"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtPeaceStudies 
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   240
      TabIndex        =   0
      Text            =   "Which woman helped draft the Universal Declaration of Human Rights?"
      Top             =   360
      Width           =   7575
   End
End
Attribute VB_Name = "frmHerMattson5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AmazingQuiz
'frmHerMattson5
'Ee Her and Jennifer Mattson
'This is the third quiz question.
'Written on 3/19/06

Dim RightAnswer As Boolean
Dim Answer As String
Private Sub Option1_Click()
    RightAnswer = True
    Counter = Counter + 1
    Answer = MsgBox("Congratulations! You got the answer right!", , "Right Answer")
    frmHerMattson5.Hide
    frmHerMattson6.Show
End Sub

Private Sub Option2_Click()
    RightAnswer = False
    Answer = MsgBox("Wrong Answer, the correct Answer was Eleanor Roosevelt", , "Wrong")
    frmHerMattson5.Hide
    frmHerMattson6.Show
End Sub

Private Sub Option3_Click()
    RightAnswer = False
    Answer = MsgBox("Wrong Answer, the correct Answer was Eleanor Roosevelt", , "Wrong")
    frmHerMattson5.Hide
    frmHerMattson6.Show
End Sub
