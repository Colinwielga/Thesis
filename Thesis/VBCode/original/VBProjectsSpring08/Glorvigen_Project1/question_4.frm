VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00FFFF80&
   Caption         =   "Question 4"
   ClientHeight    =   4005
   ClientLeft      =   3555
   ClientTop       =   3870
   ClientWidth     =   7515
   LinkTopic       =   "Form12"
   ScaleHeight     =   4005
   ScaleWidth      =   7515
   Begin VB.CommandButton cmdsmall 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Smallmouth Bass"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5280
      Picture         =   "question_4.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdwalleye 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Walleye"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      Picture         =   "question_4.frx":068A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Which is the Minnesota state fish?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   6735
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Minnesota Fisher
'Question 4
'Eric Glorvigen
'March 12
'question 4 of five, asks fourth questions tallies ctr, and
'remembers wrong answers

Private Sub cmdwalleye_Click()
    'this button is the correct anwser and tallie to question ctr
            wrongfour = False
            Form12.Hide
            Form13.Show
            questionctr = questionctr + 1

End Sub

Private Sub cmdsmall_Click()
    'this is the wrong answer, set boolean to true
            wrongfour = True
            Form12.Hide
            Form13.Show
            
End Sub



