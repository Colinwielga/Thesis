VERSION 5.00
Begin VB.Form Pitching 
   BackColor       =   &H00400000&
   Caption         =   "Form2"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9720
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   48.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   8340
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox piclogo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   600
      Picture         =   "pitchingstats.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdteamera 
      Caption         =   "Average Team ERA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdgoto 
      Caption         =   "Go to Batting Stats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdera 
      Caption         =   "Sort by ERA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdpitcher 
      Caption         =   "Find the pitcher's ERA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtpitching 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   720
      ScaleHeight     =   3675
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   3840
      Width           =   8175
   End
   Begin VB.Label lblMNTWINS 
      BackColor       =   &H00400000&
      Caption         =   "Minnesota Twins Pitchers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   600
      TabIndex        =   10
      Top             =   480
      Width           =   6135
   End
   Begin VB.Label lbldesigner 
      Caption         =   "Designed:  Jason Pfeilsticker"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   9
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label lblpitching 
      Caption         =   "Enter the name of the pitcher you would like to find"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "Pitching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form deals with pitching statistics.
'A person will be able to find a players pitching stats by entering the name in a text box.
'a person will also be able to sort by ERA and find a team ERA for the season.

Private Sub cmdera_Click()

Dim Comp As Integer
Dim Pass As Integer
Dim tempERA As Single
Dim tempPitchers As String
Dim tempWins As Integer
Dim tempLosses As Integer
Dim tempSaves As Integer
Dim tempStrikeouts As Integer
Dim tempInnings As Single


'clears anything in the picture box
picResults.Cls

'this will sort by lowest ERA
'have to sort other stats so that they go with the correct pitcher
For Pass = 1 To CTR2 - 1
    For Comp = 1 To CTR2 - Pass
        If ERA(Comp) > ERA(Comp + 1) Then
            tempERA = ERA(Comp)
            ERA(Comp) = ERA(Comp + 1)
            ERA(Comp + 1) = tempERA
            tempPitchers = Pitchers(Comp)
            Pitchers(Comp) = Pitchers(Comp + 1)
            Pitchers(Comp + 1) = tempPitchers
            tempWins = Wins(Comp)
            Wins(Comp) = Wins(Comp + 1)
            Wins(Comp + 1) = tempWins
            tempLosses = Losses(Comp)
            Losses(Comp) = Losses(Comp + 1)
            Losses(Comp + 1) = tempLosses
            tempSaves = Saves(Comp)
            Saves(Comp) = Saves(Comp + 1)
            Saves(Comp + 1) = temSaves
            tempStrikeouts = Strikeouts(Comp)
            Strikeouts(Comp) = Strikeouts(Comp + 1)
            Strikeouts(Comp + 1) = tempStrikeouts
            tempInnings = Innings(Comp)
            Innings(Comp) = Innings(Comp + 1)
            Innings(Comp + 1) = tempInnings
        End If
    Next Comp
Next Pass

'print headers
picResults.Print "Pitchers"; Tab(20); "ERA"; Tab(30); "Wins"; Tab(45); "Losses"; Tab(60); "Saves"; Tab(75); "Strikeouts"; Tab(90); "Innings"
'run a loop to print out the names of players and their ERA
For J = 1 To CTR2
             picResults.Print Pitchers(J); Tab(20); ERA(J); Tab(30); Wins(J); Tab(45); Losses(J); Tab(60); Saves(J); Tab(75); Strikeouts(J); Tab(90); Innings(J)
Next J
End Sub

Private Sub cmdgoto_Click()
'this will go to the batting stats form
PositionPlayers.Show
Pitching.Hide
End Sub

Private Sub cmdk_Click()

End Sub

Private Sub cmdpitcher_Click()
'clears anything in the picture box
picResults.Cls

'this will find the player's stats when a user enters the name in the text box
I = 0
NotFound = True
Pitcher = txtpitching.Text
'goes through a loop to find the player
Do While NotFound And I < CTR2
    I = I + 1
    If Pitcher = Pitchers(I) Then NotFound = False
Loop
If NotFound Then
 'message box statement for player not found
        MsgBox "The name of the person entered in not on the Minnesota Twins roster.", , "Error"
    Else
        'print headers
        picResults.Print "Pitcher"; Tab(20); "ERA"
        'print statement for found player
        picResults.Print Pitchers(I); Tab(20); ERA(I)
       
End If

End Sub

Private Sub cmdQuit_Click()
End
'ends the program
End Sub

Private Sub cmdteamera_Click()
'dim variables to find team ERA

Dim Sum As Single
Dim K As Integer
Dim TeamERA As Single
K = 0
'clears anything in the picture box
picResults.Cls

'for/next to find the ERA of all the players combined
For K = 1 To CTR2
    Sum = Sum + ERA(K)
Next K
'this will find the team average
TeamERA = Sum / CTR2
picResults.Print "The team average for ERA is "; FormatNumber(TeamERA, 2); "."
End Sub

Private Sub Command1_Click()
For J = 1 To CTR
             picResults.Print Pitchers(J); Tab(20); ERA(J); Tab(30); Wins(J); Tab(45); Losses(J); Tab(60)
Next J
End Sub
