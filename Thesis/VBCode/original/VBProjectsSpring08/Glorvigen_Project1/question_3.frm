VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00FFFF80&
   Caption         =   "Question 3"
   ClientHeight    =   3915
   ClientLeft      =   3555
   ClientTop       =   3870
   ClientWidth     =   7650
   LinkTopic       =   "Form11"
   ScaleHeight     =   3915
   ScaleWidth      =   7650
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
      Left            =   5520
      Picture         =   "question_3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
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
      Picture         =   "question_3.frx":062C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Which fish is most often caught?"
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
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   6735
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesota Fisher
'Question 3
'Eric Glorvigen
'March 12
'question 3 of 5, tallies correct answers, remembers wrong answers

Private Sub cmdnorthern_Click()
    'this answer is correct, tallies to global ctr
            wrongthree = False
            Form11.Hide
            Form12.Show
            questionctr = questionctr + 1
        
    
    
End Sub

Private Sub cmdmusky_Click()
    'this button is wrong, converts boolean to true
            wrongthree = True
            Form11.Hide
            Form12.Show
        
End Sub




