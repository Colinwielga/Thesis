VERSION 5.00
Begin VB.Form frmWin 
   BackColor       =   &H00800080&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOutput 
      BackColor       =   &H80000015&
      Caption         =   "Output Score onto File"
      Height          =   800
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6720
      Width           =   2500
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000015&
      Caption         =   "Quit"
      Height          =   800
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7920
      UseMaskColor    =   -1  'True
      Width           =   2500
   End
   Begin VB.CommandButton cmdRetry 
      BackColor       =   &H80000015&
      Caption         =   "Replay?"
      Height          =   800
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7920
      UseMaskColor    =   -1  'True
      Width           =   2500
   End
   Begin VB.CommandButton cmdScore 
      BackColor       =   &H80000015&
      Caption         =   "See Final Score!"
      Height          =   800
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      UseMaskColor    =   -1  'True
      Width           =   2500
   End
   Begin VB.PictureBox picScore 
      BackColor       =   &H00800080&
      ForeColor       =   &H8000000F&
      Height          =   5175
      Left            =   8520
      ScaleHeight     =   5115
      ScaleWidth      =   5355
      TabIndex        =   1
      Top             =   1320
      Width           =   5415
   End
   Begin VB.Label lblWin 
      BackColor       =   &H00800080&
      Caption         =   "Good Job!  You won!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   90
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   8295
   End
End
Attribute VB_Name = "frmWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name:  Super Awesome Cave Adventure Game
'Form Name:  frmWin
'Author:  Peter Woodruff
'Date Written:  3-15-09
'Purpose:  This is the screen the user see if he/she wins.  It totals the user's points.


Option Explicit
Dim I As Integer
Dim Score As Single
Dim Item(1 To 24) As String
Dim Price(1 To 100) As Integer

Private Sub cmdOutput_Click()

    'Write total score and date to Scores.txt
    
    Open App.Path & "\Scores.txt" For Append As #1
        
            Print #1, Score, Date
        
    Close 1
    
    'Tell user the score was saved and hide button so he/she does not save score too much
    MsgBox "Score and date saved in Scores.txt", , ""
    
    cmdOutput.Visible = False
    
    
    
End Sub

Private Sub cmdQuit_Click()

    'End
    End
    
End Sub

Private Sub cmdRetry_Click()
    
    'Restarts the game
    frmWin.Visible = False
    frmTitle.Visible = True
    
End Sub

Private Sub cmdScore_Click()

    'Clear score
    Score = 0
    picScore.Cls
    
    'Counts user's points
    picScore.Print Tab(20); "Points"
    picScore.Print ""
    
    
    I = 0

        Open App.Path & "\Store.txt" For Input As #1
        
            Do Until EOF(1)
                I = I + 1
                Input #1, Item(I), Price(I)
            Loop
    
        Close 1
        
    If Sword = True Then
        picScore.Print Item(1); Tab(20); "10"
        Score = Score + 10
    End If
    
    If Light = True Then
        picScore.Print Item(3); Tab(20); "10"
        Score = Score + 10
    End If
    
    picScore.Print "Health:  " & Life
    picScore.Print Tab(10); "UnSquared..."
    picScore.Print Tab(15); "="; Tab(20); Sqr(Life)
    Score = Score + Sqr(Life)
    
    
    If Secret = True Then
        picScore.Print "Secret Room"; Tab(20); "15"
        Score = Score + 15
    End If
    
    
    picScore.Print "Coins:  " & Coins
    picScore.Print Tab(10); "Times 3.14159265 Rounded to the Nearest Integer..."
    picScore.Print Tab(15); "="; Tab(20); Int(Coins * 3.14159265)
    Score = Score + Int(Coins * 3.14159265)
    
    picScore.Print
    picScore.Print Tab(0); "Dragon Reformed"; Tab(20); "35"
    Score = Score + 35
        
    picScore.Print "***************************************"
    
    picScore.Print "Total"
    picScore.Print ""
    picScore.Print Score; " Points"
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdSurvey_Click()

    'Go to survey
    frmWin.Visible = False
    frmSurvey.Visible = True
    
End Sub

