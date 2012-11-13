VERSION 5.00
Begin VB.Form frmBennieTrivia 
   Caption         =   "Bennie Trivia"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   Picture         =   "frmBennieTrivia.frx":0000
   ScaleHeight     =   9615
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Back to Menu"
      Height          =   975
      Left            =   10320
      TabIndex        =   2
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play Now!"
      Height          =   1335
      Left            =   9120
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblBennieTrivia 
      Caption         =   "Bennie Trivia!"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmBennieTrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Fun with CSB/SJU History!
'frmBennieTrivia
'Audrey Gabe
'Written 3/16/09
'This section is Bennie trivia.  The user gets to answer a few questions and can also see if their answer is correct.


Private Sub cmdMenu_Click()
frmBennieTrivia.Hide
frmMenu.Show
End Sub

Private Sub cmdPlay_Click()
Dim K As String
Dim L As String

K = InputBox("What year was the College of Saint Benedict established?", , "Enter answer here") 'Inputbox gives user a question and allows user to type in an answer
If K = "1913" Then 'Checks to see if answer matches what I have said it should be
    MsgBox "Great!", , "Your answer is correct!" 'What user sees if answer is correct
    Else
        MsgBox "Sorry, the answer is 1913", , "Incorrect" 'What user sees if answer is incorrect
End If
K = InputBox("Who was the first CSB president?", , "Enter answer here")
If K = "Mother Cecilia Kapsner" Then
    MsgBox "Awesome!", , "Your answer is correct!"
    Else
        MsgBox "Sorry, the answer is Mother Cecilia Kapsner", , "Incorrect"
End If
K = InputBox("What is the CSB athletic mascot name?", , "Enter answer here")
If K = "Blazer" Then
    MsgBox "Splendid!", , "Your answer is correct!"
    Else
        MsgBox "Sorry, the answer is Blazer", , "Incorrect"
End If
K = InputBox("What year did Bennies start taking classes at St. John's University?", , "Enter answer here")
If K = "1963" Then
    MsgBox "Great!", , "Your answer is correct!"
    Else
        MsgBox "Sorry, the answer is 1963", , "Incorrect"
End If

frmBennieTrivia.Hide
frmMenu.Show
        


End Sub


