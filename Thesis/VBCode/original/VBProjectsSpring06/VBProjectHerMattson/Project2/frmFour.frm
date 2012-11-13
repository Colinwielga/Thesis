VERSION 5.00
Begin VB.Form frmHerMattson4 
   Caption         =   "Question #2"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   Picture         =   "frmFour.frx":0000
   ScaleHeight     =   4785
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option3 
      Caption         =   "Mitch Albom"
      Height          =   735
      Left            =   2640
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Arthur Miller"
      Height          =   735
      Left            =   1320
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Leif Erikson"
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Text            =   "Who wrote the following book?"
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "frmHerMattson4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AmazingQuiz
'frmHerMattson4
'Ee Her and Jennifer Mattson
'This is the second quiz question.
'Written on 3/19/06

Dim RightAnswer As Boolean
Dim Answer As String

Private Sub Option1_Click()
    RightAnswer = False
    Answer = MsgBox("Wrong Answer, the correct Answer was Mitch Albom", , "Wrong")
    frmHerMattson4.Hide
    frmHerMattson5.Show
End Sub

Private Sub Option2_Click()
    RightAnswer = False
    Answer = MsgBox("Wrong Answer, the correct Answer was Mitch Albom", , "Wrong")
    frmHerMattson4.Hide
    frmHerMattson5.Show
End Sub

Private Sub Option3_Click()
    RightAnswer = True
    Counter = Counter + 1
    Answer = MsgBox("Congratulations! You got the answer right!", , "Right Answer")
    frmHerMattson4.Hide
    frmHerMattson5.Show
End Sub
