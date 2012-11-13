VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFF80&
   Caption         =   "Question 1"
   ClientHeight    =   3345
   ClientLeft      =   3555
   ClientTop       =   4110
   ClientWidth     =   8295
   LinkTopic       =   "Form4"
   ScaleHeight     =   3345
   ScaleWidth      =   8295
   Visible         =   0   'False
   Begin VB.CommandButton cmdcrappie 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Crappie"
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
      Left            =   600
      Picture         =   "question_one.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdbluegill 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bluegill"
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
      Left            =   5760
      Picture         =   "question_one.frx":075E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Which fish is more likely to be caught at dusk?"
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
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesota Fisher
'Question 1
'Eric Glorvigen
'March 12
'this is part one of the quiz and navigates to part 2


Private Sub cmdbluegill_Click()
    'this is the wrong answer and makes the boolean false
    'navigates to question two
            Form4.Hide
            Form10.Show
            wrongone = True
End Sub

Private Sub cmdcrappie_Click()
    'this is the correct answer and adds one to the question ctr
    'navigates to question two
        questionctr = questionctr + 1
        Form4.Hide
        Form10.Show
        wrongone = False
    
End Sub


