VERSION 5.00
Begin VB.Form frmFinalQuestion 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17310
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   17310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWager 
      BackColor       =   &H80000013&
      Caption         =   "Click here to enter your final wager."
      Height          =   1215
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000013&
      Caption         =   "Quit"
      Height          =   1215
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdBacktoBeginning 
      BackColor       =   &H80000013&
      Caption         =   "Click here to start the game over."
      Height          =   1215
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdClickAnswer 
      BackColor       =   &H80000013&
      Caption         =   "Click here after you have entered your answer."
      Enabled         =   0   'False
      Height          =   1215
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdShowPicture 
      BackColor       =   &H80000013&
      Caption         =   "Click here to show the picture hint."
      Enabled         =   0   'False
      Height          =   1215
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtCricketAnswer 
      Height          =   1215
      Left            =   11400
      TabIndex        =   2
      Top             =   3960
      Width           =   5175
   End
   Begin VB.PictureBox picResults3 
      BackColor       =   &H80000012&
      Height          =   7095
      Left            =   480
      ScaleHeight     =   7035
      ScaleWidth      =   5475
      TabIndex        =   1
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label lblCricketLabel 
      BackColor       =   &H8000000E&
      Caption         =   "What sport used the term ""home run"" long before baseball?"
      Height          =   255
      Left            =   11400
      TabIndex        =   0
      Top             =   3480
      Width           =   4215
   End
End
Attribute VB_Name = "frmFinalQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is the form for the Final Question. After answering the question, the user has the option to start the game over or quit.
Option Explicit
Dim Wager As Single
Private Sub cmdBacktoBeginning_Click()          'click here after the question has been answered to start the game over again.
frmMain.Show
frmFinalQuestion.Hide
picResults3.Picture = LoadPicture("")
txtCricketAnswer.Text = ""
runningTotal = 0
CTR = 0
cmdShowPicture.Enabled = False
cmdClickAnswer.Enabled = False
cmdWager.Enabled = True
End Sub

Private Sub cmdQuit_Click()                     'quit
End
End Sub

Private Sub cmdShowPicture_Click()              'click here to show the picture in the picture box
picResults3.Picture = LoadPicture(App.Path & "\images\" & "cricket.jpg")
End Sub

Private Sub cmdClickAnswer_Click()              'click here after the user has entered their answer into the text box. Also used to
Dim Answer As String                            'show if the user got the question right or wrong, how many points the user finished
Answer = txtCricketAnswer.Text                  'the game with, and tell the user the options of either quitting or starting over.
                If Answer = "cricket" Then
                    runningTotal = runningTotal + Wager
                    MsgBox ("That's right! You win " & Wager & " points!")
                Else:
                    runningTotal = runningTotal - Wager
                    MsgBox ("I'm sorry, the answer is Cricket.")
                End If
cmdShowPicture.Enabled = False
cmdClickAnswer.Enabled = False
MsgBox ("Congradulations! You finished the game with " & runningTotal & " points! If you would like to play again, please click the appropriate button.")
End Sub


Private Sub cmdWager_Click()                    'click here to enter a wager for the question.

Wager = InputBox("Please enter a wager that is less than or equal to your total. If you have less than 200 points, then the most you may wager is 200. Your points thus far are: " & runningTotal & ".")
If runningTotal <= 200 Then
    Do While Wager > 200
        Wager = InputBox("Since you have less than 200 points, the most you may wager is 200.")
    Loop
ElseIf runningTotal > 200 Then
            Do While Wager > runningTotal
                Wager = InputBox("Please enter a wager that is less than or equal to your total.")
            Loop
Else: Wager = InputBox("I'm sorry. Please enter a wager that is less than or equal to your total. If you have less than 200 points, then the most you may wager is 200.")
End If
cmdShowPicture.Enabled = True
cmdClickAnswer.Enabled = True
cmdWager.Enabled = False
End Sub
