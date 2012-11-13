VERSION 5.00
Begin VB.Form frmGame 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Pictitle 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   4515
      TabIndex        =   10
      Top             =   120
      Width           =   4575
   End
   Begin VB.PictureBox picresults6 
      Height          =   495
      Left            =   4080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Cmdback 
      Caption         =   "Go Back"
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdguess 
      Caption         =   "Start"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   1335
   End
   Begin VB.PictureBox Picresults5 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   3915
      TabIndex        =   4
      Top             =   3120
      Width           =   3975
   End
   Begin VB.PictureBox picresults4 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   3915
      TabIndex        =   3
      Top             =   2520
      Width           =   3975
   End
   Begin VB.PictureBox picresults3 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   3915
      TabIndex        =   2
      Top             =   1920
      Width           =   3975
   End
   Begin VB.PictureBox picresults2 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   3915
      TabIndex        =   1
      Top             =   1320
      Width           =   3975
   End
   Begin VB.PictureBox picresults1 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label lblstrikes 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Number of strikes against you:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   3855
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmdback_Click()
frmGame.Hide                'switches back to welcome form
frmWelcome.Show
picresults1.Cls             'clears all picture boxes
picresults2.Cls
picresults3.Cls
picresults4.Cls
Picresults5.Cls
picresults6.Cls
Pictitle.Cls
cmdguess.Visible = True     'allows user to make guesses again
End Sub

Private Sub cmdguess_Click()
Dim Guess As String     'delcaring variables used in the command button
Dim Pos As Integer
Dim Strikes As Integer
Dim Correct As Integer
Dim Found As Boolean
Dim result1 As Boolean
Dim result2 As Boolean
Dim result3 As Boolean
Dim result4 As Boolean
Dim result5 As Boolean
Dim Continue As Boolean

If Question = 1 Then 'Displaying correct question
    Pictitle.Print "Name a good activity for a hot summer day."
ElseIf Question = 2 Then
    Pictitle.Print "Name a sporting event you might attend.";
Else
    Pictitle.Print "Name a musical instrument a jazz musician might play."
End If

Found = False
Pos = 0
Correct = 0
Strikes = 0

Do Until Strikes = 3 Or Correct = 5 'establishing rules for the game
    Guess = InputBox("Enter a guess in all lower case letters. Do not repeat guesses.") 'Geting the user to input their guess.
    For Pos = 1 To CTR
        If Guess = Answer(Pos) Then     'Checking to see if the guess is correct
            Found = True
            Continue = True
        End If
        
        If Found = True And Pos = 1 Then    'Determining which answer was guessed if any
            result1 = True
        End If
            
        If Found = True And Pos = 2 Then    'Determining which answer was guessed if any
            result2 = True
        End If
            
        If Found = True And Pos = 3 Then    'Determining which answer was guessed if any
            result3 = True
        End If
            
        If Found = True And Pos = 4 Then    'Determining which answer was guessed if any
            result4 = True
        End If
            
        If Found = True And Pos = 5 Then    'Determining which answer was guessed if any
            result5 = True
        End If
        
        If result1 = True Then
            picresults1.Print Answer(Pos), Answernum(Pos)   'Printing correct answer if any
            result1 = False
        End If
        
        If result2 = True Then
            picresults2.Print Answer(Pos), Answernum(Pos)   'Printing correct answer if any
            result2 = False
        End If
        
        If result3 = True Then
            picresults3.Print Answer(Pos), Answernum(Pos)   'Printing correct answer if any
            result3 = False
        End If
        
        If result4 = True Then
            picresults4.Print Answer(Pos), Answernum(Pos)   'Printing correct answer if any
            result4 = False
        End If
        
        If result5 = True Then
            Picresults5.Print Answer(Pos), Answernum(Pos)   'Printing correct answer if any
            result5 = False
        End If
        Found = False       'reseting loop
    Next Pos                'checking next answer
        If Continue = False Then            'counting either strikes of correct answers
            Strikes = Strikes + 1
            picresults6.Cls
            picresults6.Print Strikes
        Else
            Correct = Correct + 1
        End If
    Continue = False                     'reseting strikes and correct answers counter
Loop                            'Getting next input
        
    If Strikes = 3 Then         'user loses game
        MsgBox ("Sorry you are out of strikes. Game over. Press Go Back to play again.")
        cmdguess.Visible = False    'Stops user from input more guesses after game is over.
    End If
        
    If Correct = 5 Then         'user wins game
        MsgBox ("You win! Press Go Back to play again.")
        cmdguess.Visible = False    'Stops user from input more guesses after game is over.
    End If

End Sub

Private Sub CmdQuit_Click()
End                 'quit program
End Sub

