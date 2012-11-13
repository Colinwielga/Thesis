VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00FFFF80&
   Caption         =   "Question 2"
   ClientHeight    =   3570
   ClientLeft      =   3780
   ClientTop       =   3870
   ClientWidth     =   7605
   LinkTopic       =   "Form10"
   ScaleHeight     =   3570
   ScaleWidth      =   7605
   Visible         =   0   'False
   Begin VB.CommandButton cmdmusky 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Musky"
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
      Picture         =   "question_2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdnorthern 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Northern Pike"
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
      Left            =   5400
      Picture         =   "question_2.frx":05F7
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Which fish has six or more pores on the bottom jaw?"
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
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesota Fisher
'Question 2
'Eric Glorvigen
'March 12
'Question two of five, tallies correct answers, and remembers
' wrong anwsers

Private Sub cmdnorthern_Click()
    'this button is incorrect
    'moves to third question
            Form10.Hide
            Form11.Show
            wrongtwo = True
 
End Sub

Private Sub cmdmusky_Click()
    'this button is correct and tallys
    'correct answers to the ctr
        wrongtwo = False
        questionctr = questionctr + 1
        Form10.Hide
        Form11.Show
    
End Sub


