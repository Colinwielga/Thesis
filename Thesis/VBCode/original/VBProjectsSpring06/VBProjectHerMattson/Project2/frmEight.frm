VERSION 5.00
Begin VB.Form frmHerMattson8 
   BackColor       =   &H0000FFFF&
   Caption         =   "Question #6"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8625
   FillColor       =   &H00C0C0FF&
   BeginProperty Font 
      Name            =   "Lithos Pro Regular"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0080FFFF&
   LinkTopic       =   "Form1"
   Picture         =   "frmEight.frx":0000
   ScaleHeight     =   5985
   ScaleWidth      =   8625
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FF8080&
      Caption         =   "Pike"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      TabIndex        =   3
      Top             =   3000
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0000FF00&
      Caption         =   "Perk"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF00FF&
      Caption         =   "Peak"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "What is the name of the coffee shop in Clemens?"
      Top             =   480
      Width           =   7935
   End
End
Attribute VB_Name = "frmHerMattson8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AmazingQuiz
'frmHerMattson8
'Ee Her and Jennifer Mattson
'This is the sixth quiz question.
'Written on 3/19/06

Dim RightAnswer As Boolean
Dim Answer As String

Private Sub Option1_Click()
    RightAnswer = False
    Answer = MsgBox("Wrong Answer, the correct Answer was Perk", , "Wrong")
    frmHerMattson8.Hide
    frmHerMattson9.Show
End Sub

Private Sub Option2_Click()
    RightAnswer = True
    Counter = Counter + 1
    Answer = MsgBox("Congratulations! You got the answer right!", , "Right Answer")
    frmHerMattson8.Hide
    frmHerMattson9.Show
End Sub

Private Sub Option3_Click()
    RightAnswer = False
    Answer = MsgBox("Wrong Answer, the correct Answer was Perk", , "Wrong")
    frmHerMattson8.Hide
    frmHerMattson9.Show
End Sub
