VERSION 5.00
Begin VB.Form frmLeague 
   BackColor       =   &H00004000&
   Caption         =   "Standings in the NHL"
   ClientHeight    =   11715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13860
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11715
   ScaleWidth      =   13860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTop 
      BackColor       =   &H00000080&
      Caption         =   "Who are the top 10 teams in the league?"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9000
      Width           =   4935
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00000080&
      Caption         =   "Back to the rink"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   10800
      Width           =   4575
   End
   Begin VB.CommandButton cmdEastern 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   9360
      Picture         =   "frmLeague.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   4455
   End
   Begin VB.CommandButton cmdLeague 
      BackColor       =   &H00000080&
      Caption         =   "Who is the best team in the league?"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   4935
   End
   Begin VB.CommandButton cmdWestern 
      Height          =   3495
      Left            =   0
      Picture         =   "frmLeague.frx":6637
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7200
      Width           =   4215
   End
   Begin VB.PictureBox picLeague 
      BackColor       =   &H0080FFFF&
      Height          =   7095
      Left            =   7200
      ScaleHeight     =   7035
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
   Begin VB.Image Image1 
      Height          =   6135
      Left            =   720
      Picture         =   "frmLeague.frx":A77E
      Top             =   360
      Width           =   5505
   End
   Begin VB.Label lblLegend 
      BackColor       =   &H00000080&
      Caption         =   "GP=Games Played     W=Wins     L=Losses     OTL=Over Time Losses"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   10800
      Width           =   5295
   End
End
Attribute VB_Name = "frmLeague"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesota Wild Visual Basic Project
'League Form
'Authors: Patrick Johnson
'17 Feb 2010
'The purpose of this form is so the user can interact with
'the Western and Eastern conference standings,
'along with the league standings

Dim CTR As Integer, Team(1 To 100) As String, GamesPlayed(1 To 100) As Integer
Dim Wins(1 To 100) As Integer, Losses(1 To 100) As Integer, OverTimeLosses(1 To 100) As Integer

Private Sub cmdBack_Click()
'show the main form and hide the other forms
frmLeague.Hide
frmMain.Show
frmRoster.Hide
frmShot.Hide
frmWelcome.Hide
frmShop.Hide
End Sub

Private Sub cmdEastern_Click()
'This button prints the Eastern Conference standings

'clear previous standings
picLeague.Cls

'initialize CTR to zero, to be used for position in the array
CTR = 0

'prepare file to be read
Open App.Path & "\EasternConferenceRecord.txt" For Input As #1

'print header
picLeague.Print "Team"; Tab(20); "GP"; Tab(40); "W"; Tab(60); "L"; Tab(80); "OTL"
picLeague.Print "********************************************************************************************************"

'begin loop statement
Do While Not EOF(1)
    'increment CTR each time through the loop
    'to move to next position in the array
    CTR = CTR + 1
    
    'read data from file and print into arrays
    Input #1, Team(CTR), GamesPlayed(CTR), Wins(CTR), Losses(CTR), OverTimeLosses(CTR)
    picLeague.Print Team(CTR); Tab(20); GamesPlayed(CTR); Tab(40); Wins(CTR); Tab(60); Losses(CTR); Tab(80); OverTimeLosses(CTR)
Loop
Close #1
End Sub

Private Sub CmdLeague_Click()
'this button shows the best team in the NHL

'declare variables
Dim MostWins As Integer, BestTeam As String, WinsSearch As Integer, TeamSearch As String
Dim Blank1 As String, Blank2 As String, Blank3 As String
'declared Blank variables to account for the arrays not used in this button

'prepare file to be read
Open App.Path & "\NHLTeams.txt" For Input As #1

'initialize most wins
MostWins = 0

'begin loop statement
Do While Not EOF(1)
    'get data from file
    Input #1, TeamSearch, Blank1, WinsSearch, Blank2, Blank3
    
    'ask which team has most wins and remember the team name and how many wins
    If WinsSearch > MostWins Then
        MostWins = WinsSearch
        BestTeam = TeamSearch
    End If
Loop
Close #1

'show computation in a message box
MsgBox ("The best team in the NHL is " & BestTeam & " with " & MostWins & " wins")

End Sub

Private Sub cmdTop_Click()
'this button shows the top 10 teams in the league
'declare variables
Dim J As Integer, Pass As Integer, Pos As Integer, TempTeam As String, TempWins As String, TempGamesPlayed As String
Dim TempLosses As String, TempOverTimeLosses As String

'clear previous standings
picLeague.Cls

'initialize CTR to zero, to be used for position in the array
 CTR = 0
 
'prepare file to be read
Open App.Path & "\NHLTeams.txt" For Input As #1

'begin loop statement
Do While Not EOF(1)
    'increment CTR
    CTR = CTR + 1
    'get data from file
    Input #1, Team(CTR), GamesPlayed(CTR), Wins(CTR), Losses(CTR), OverTimeLosses(CTR)
Loop
Close #1

'print header
picLeague.Print "Team"; Tab(20); "GP"; Tab(40); "W"; Tab(60); "L"; Tab(80); "OTL"
picLeague.Print "********************************************************************************************************"

'use Bubble Sort to sort the Team and Wins arrays into desired order
'this will sort top 10 teams by number of wins
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If Wins(Pos) < Wins(Pos + 1) Then
            TempWins = Wins(Pos)
            Wins(Pos) = Wins(Pos + 1)
            Wins(Pos + 1) = TempWins
            TempTeam = Team(Pos)
            Team(Pos) = Team(Pos + 1)
            Team(Pos + 1) = TempTeam
            TempGamesPlayed = GamesPlayed(Pos)
            GamesPlayed(Pos) = GamesPlayed(Pos + 1)
            GamesPlayed(Pos + 1) = TempGamesPlayed
            TempLosses = Losses(Pos)
            Losses(Pos) = Losses(Pos + 1)
            Losses(Pos + 1) = TempLosses
            TempOverTimeLosses = OverTimeLosses(Pos)
            OverTimeLosses(Pos) = OverTimeLosses(Pos + 1)
            OverTimeLosses(Pos + 1) = TempOverTimeLosses
        End If  'end If/Then statement
    Next Pos    'end For/Next statement
Next Pass       'end For/Pass statement

'print sorted list for top 10 teams
For J = 1 To 10
    picLeague.Print Team(J); Tab(20); GamesPlayed(J); Tab(40); Wins(J); Tab(60); Losses(J); Tab(80); OverTimeLosses(J)
Next J
End Sub

Private Sub cmdWestern_Click()
'This button prints Western Conference standings

'clear previous standings
picLeague.Cls

'initialize CTR to zero, to be used for position in the array
CTR = 0

'prepare file to be read
Open App.Path & "\WesternConferenceRecord.txt" For Input As #1

'print header
picLeague.Print "Team"; Tab(20); "GP"; Tab(40); "W"; Tab(60); "L"; Tab(80); "OTL"
picLeague.Print "********************************************************************************************************"


'begin loop statement
Do While Not EOF(1)
    'increment CTR each time through the loop
    'to move to the next position in the array
    CTR = CTR + 1
    
    'read data from file and print into arrays
    Input #1, Team(CTR), GamesPlayed(CTR), Wins(CTR), Losses(CTR), OverTimeLosses(CTR)
    picLeague.Print Team(CTR); Tab(20); GamesPlayed(CTR); Tab(40); Wins(CTR); Tab(60); Losses(CTR); Tab(80); OverTimeLosses(CTR)
Loop
Close #1

End Sub

