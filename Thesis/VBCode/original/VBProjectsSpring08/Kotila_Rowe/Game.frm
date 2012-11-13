VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   ClientHeight    =   12015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17955
   LinkTopic       =   "Form1"
   ScaleHeight     =   12015
   ScaleWidth      =   17955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdend 
      BackColor       =   &H000000C0&
      Caption         =   "Ready to leave?"
      Height          =   1215
      Left            =   15960
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   10200
      Width           =   1695
   End
   Begin VB.PictureBox picRunning 
      Height          =   2535
      Left            =   15960
      ScaleHeight     =   2475
      ScaleWidth      =   1635
      TabIndex        =   28
      Top             =   7440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H0000C000&
      Caption         =   "Go"
      Height          =   1215
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   10200
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   10095
      Left            =   0
      Picture         =   "Game.frx":0000
      ScaleHeight     =   10035
      ScaleWidth      =   15675
      TabIndex        =   0
      Top             =   0
      Width           =   15735
      Begin VB.CommandButton cmdTarget 
         BackColor       =   &H000000FF&
         Height          =   495
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdJeuneLune 
         BackColor       =   &H00FFFF00&
         Height          =   375
         Left            =   7320
         MaskColor       =   &H00FFFF00&
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdTIR 
         BackColor       =   &H00FFFF00&
         Height          =   375
         Left            =   14880
         MaskColor       =   &H00FFFF00&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   6600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdElliotPark 
         BackColor       =   &H0000FF00&
         Height          =   495
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   6840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdCurriePark 
         BackColor       =   &H0000FF00&
         Height          =   495
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   6600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdGoldMedalPark 
         BackColor       =   &H0000FF00&
         Height          =   495
         Left            =   12360
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdWestRiverPark 
         BackColor       =   &H0000FF00&
         Height          =   495
         Left            =   11760
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdPWP 
         BackColor       =   &H0000FF00&
         Height          =   495
         Left            =   12480
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdGuthrie 
         BackColor       =   &H00FFFF00&
         Height          =   375
         Left            =   11760
         MaskColor       =   &H00FFFF00&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2880
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdLoringPark 
         BackColor       =   &H0000FF00&
         Height          =   495
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   6720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdHHH 
         BackColor       =   &H000000FF&
         Height          =   495
         Left            =   10920
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5040
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdNicollet 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdPalomino 
         BackColor       =   &H00FF00FF&
         Height          =   375
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdIchiban 
         BackColor       =   &H00FF00FF&
         Height          =   375
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdZelo 
         BackColor       =   &H00FF00FF&
         Height          =   375
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdCafeHavana 
         BackColor       =   &H00FF00FF&
         Height          =   375
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdOldSpaghetti 
         BackColor       =   &H00FF00FF&
         Height          =   375
         Left            =   11880
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdMacysParking 
         BackColor       =   &H000080FF&
         Height          =   375
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdMillParking 
         BackColor       =   &H000080FF&
         Height          =   375
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmd4thStParking 
         BackColor       =   &H000080FF&
         Height          =   375
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmd7thStParking 
         BackColor       =   &H000080FF&
         Height          =   375
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2760
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdPlazaParking 
         BackColor       =   &H000080FF&
         Height          =   375
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdPantages 
         BackColor       =   &H00FFFF00&
         Height          =   375
         Left            =   5280
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3000
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdOrpheum 
         BackColor       =   &H00FFFF00&
         Height          =   375
         Left            =   4440
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdMB 
         BackColor       =   &H00FFFF00&
         Height          =   375
         Left            =   14040
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdSpooner 
         BackColor       =   &H00FF00FF&
         Height          =   375
         Left            =   11280
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2640
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Label lblScore 
      BackColor       =   &H80000007&
      Caption         =   "Your Score Is:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   15840
      TabIndex        =   30
      Top             =   6960
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Minneapolis Travel Guide
'Form1- Game
'Kayla Kotila and Chris Rowe
'March 30
'Page for game



Option Explicit
Dim Score As Integer
Dim Answer As String


Private Sub cmdend_Click()
Form1.Hide
frm2.Show
End Sub

Private Sub cmdGo_Click()
'play the game


Dim EndGame As Label
Dim Answer As String
Score = 0
picRunning.Cls

picRunning.Visible = True
picRunning.Print "Your score is"

cmdHHH.Visible = True
Answer = InputBox("Whats this place called?", "Question 1")
If Answer = "Metrodome" Then
    Score = Score + 1
End If

picRunning.Print Score

cmdHHH.Visible = False
cmdTIR.Visible = True

Answer = InputBox("Whats this place called?", "Question 2")
If Answer = "Theater in the Round" Then
    Score = Score + 1
End If
picRunning.Print Score

cmdTIR.Visible = False
cmdPalomino.Visible = True

Answer = InputBox("Whats this place called?", "Question 3")
If Answer = "Palomino" Then
Score = Score + 1
End If
picRunning.Print Score

cmdPalomino.Visible = False
cmdGuthrie.Visible = True


Answer = InputBox("Whats this place called?", "Question 4")
If Answer = "Guthrie Theater" Then
Score = Score + 1
End If
picRunning.Print Score

cmdGuthrie.Visible = False
cmdLoringPark.Visible = True


Answer = InputBox("Whats this place called?", "Question 5")
If Answer = "Loring Park" Then
Score = Score + 1
End If
picRunning.Print Score

cmdLoringPark.Visible = False
cmdIchiban.Visible = True


Answer = InputBox("Whats this place called?", "Question 6")
If Answer = "Ichibans" Then
Score = Score + 1
End If
picRunning.Print Score

cmdIchiban.Visible = False
cmdPantages.Visible = True

Answer = InputBox("Whats this place called?", "Question 7")
If Answer = "Pantages Theater" Then
Score = Score + 1
End If
picRunning.Print Score

cmdPantages.Visible = False
cmdGoldMedalPark.Visible = True

Answer = InputBox("Whats this place called?", "Question 8")
If Answer = "Gold Medal Park" Then
Score = Score + 1
End If
picRunning.Print Score

cmdGoldMedalPark.Visible = False
cmdSpooner.Visible = True

Answer = InputBox("Whats this place called?", "Question 9")
If Answer = "Spoonriver" Then
Score = Score + 1
End If
picRunning.Print Score

cmdSpooner.Visible = False
cmdOrpheum.Visible = True

Answer = InputBox("Whats this place called?", "Question 10")
If Answer = "Orpheum" Then
Score = Score + 1
End If
picRunning.Print Score

cmdOrpheum.Visible = False
'cmdCafeHavana.Visible = True

'Answer = InputBox("Whats this place called?", "Question 11")
'If Answer = "Cafe Havana" Then
'Score = Score + 1
'End If
'picRunning.Print Score

'cmdCafeHavana.Visible = False
'cmdNicollet.Visible = True

'Answer = InputBox("Whats this place called?", "Question 12")
'If Answer = "Nicollet mall" Then
'Score = Score + 1
'End If
'picRunning.Print Score

'cmdNicollet.Visible = False
'cmdElliotPark.Visible = True

'Answer = InputBox("Whats this place called?", "Question 13")
'If Answer = "Elliot Park" Then
'Score = Score + 1
'End If
'picRunning.Print Score

'cmdElliotPark.Visible = False
'cmdMB.Visible = True

'picRunning.Cls
'picRunning.Print "Your score is"

'Answer = InputBox("Whats this place called?", "Question 14")
'If Answer = "Mixed Blood Theater" Then
'Score = Score + 1
'End If
'picRunning.Print Score

'cmdMB.Visible = False
'cmdLoringPark.Visible = True

'Answer = InputBox("Whats this place called?", "Question 15")
'If Answer = "Loring Park" Then
'Score = Score + 1
'End If
'picRunning.Print Score

'cmdLoringPark.Visible = False
'cmdTarget.Visible = True

'picRunning.Cls
'picRunning.Print "Your score is"

'Answer = InputBox("Whats this place called?", "Question 16")
'If Answer = "Target Center" Then
'Score = Score + 1
'End If
'picRunning.Print Score

'cmdTarget.Visible = False
'cmdCurriePark.Visible = True

'picRunning.Cls
'picRunning.Print "Your score is"

'Answer = InputBox("Whats this place called?", "Question 17")
'If Answer = "Currie Park" Then
'Score = Score + 1
'End If
'picRunning.Print Score

'cmdCurriePark.Visible = False
'cmdJeuneLune.Visible = True

'Answer = InputBox("Whats this place called?", "Question 18")
'If Answer = "Theatre de la Jeune Lune" Then
'Score = Score + 1
'End If
'picRunning.Print Score

'cmdJeuneLune.Visible = False
'cmdZelo.Visible = True

'Answer = InputBox("Whats this place called?", "Question 19")
'If Answer = "Zelo" Then
'Score = Score + 1
'End If
'picRunning.Print Score

'cmdZelo.Visible = False
'cmdWestRiverPark.Visible = True

'picRunning.Cls
'picRunning.Print "Your score is"

'Answer = InputBox("Whats this place called?", "Question 20")
'If Answer = "West River Park" Then
'Score = Score + 1
'End If
'picRunning.Print Score

'cmdWestRiverPark.Visible = False
'cmdOldSpaghetti.Visible = True

'picRunning.Cls
'picRunning.Print "Your score is"

'Answer = InputBox("Whats this place called?", "Question 21")
'If Answer = "Old Spaghetti Factory" Then
'Score = Score + 1
'End If
'picRunning.Print Score

'cmdOldSpaghetti.Visible = False
'cmdPWP.Visible = True

'Answer = InputBox("Whats this place called?", "Question 22")
'If Answer = "Philip W. Pillsbury Park" Then
'Score = Score + 1
'End If



Select Case Score
Case Is = 10
    MsgBox "WOW! you got a perfect score! 10 points!", , "Congratulations!"
Case Is > 7
    MsgBox "Not too bad! you got " & Score & " points!", , "you sure know your city"
Case Is > 5
    MsgBox "meh, I've seen better...but I've seen worse too. yours score was " & Score, , "Mediocre"
Case Is > 3
    MsgBox "You're not too good at this...go back and learn a little more about this city and try to get better than a " & Score, , "not so good"
Case Is >= 0
    MsgBox "were you trying to fail?...you got " & Score & " points", , "Yikes"
End Select

cmdPWP.Visible = False
cmdGo.Enabled = True

End Sub

