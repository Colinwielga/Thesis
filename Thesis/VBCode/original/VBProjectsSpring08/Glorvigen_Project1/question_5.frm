VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H00FFFF80&
   Caption         =   "Question 5"
   ClientHeight    =   3540
   ClientLeft      =   3090
   ClientTop       =   3630
   ClientWidth     =   7470
   LinkTopic       =   "Form13"
   ScaleHeight     =   3540
   ScaleWidth      =   7470
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
      Picture         =   "question_5.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdsteelhead 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Steelhead"
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
      Left            =   5040
      Picture         =   "question_5.frx":05F7
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Which fish fights harder pound for pound?"
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
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Minnesota Fisher
'Question 5
'Eric Glorvigen
'March 12
'question five of five, navigate back to fish id page



Private Sub cmdsteelhead_Click()
    'button is correct, tallies to global ctr
        wrongfive = False
        Form13.Hide
        Form3.Show
        questionctr = questionctr + 1
        
  End Sub

Private Sub cmdmusky_Click()
    'button is incorrect, sets boolean to true
        wrongfive = True
        Form13.Hide
        Form3.Show
    
End Sub

