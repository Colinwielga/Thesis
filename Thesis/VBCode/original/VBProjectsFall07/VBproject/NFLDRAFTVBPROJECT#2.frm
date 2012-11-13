VERSION 5.00
Begin VB.Form PreDraft 
   BackColor       =   &H0000C000&
   Caption         =   "Pre Draft"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMel 
      Height          =   5295
      Left            =   7200
      Picture         =   "NFLDRAFTVBPROJECT#2.frx":0000
      ScaleHeight     =   5235
      ScaleWidth      =   6675
      TabIndex        =   9
      Top             =   960
      Width           =   6735
   End
   Begin VB.CommandButton cmdSortform 
      Caption         =   "Move On To The Draft"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   8
      Top             =   5880
      Width           =   2175
   End
   Begin VB.PictureBox picResult3 
      Height          =   495
      Left            =   7200
      ScaleHeight     =   435
      ScaleWidth      =   6675
      TabIndex        =   7
      Top             =   6480
      Width           =   6735
   End
   Begin VB.PictureBox picResult2 
      Height          =   735
      Left            =   7200
      ScaleHeight     =   675
      ScaleWidth      =   6675
      TabIndex        =   6
      Top             =   120
      Width           =   6735
   End
   Begin VB.PictureBox picResult1 
      Height          =   6855
      Left            =   2640
      ScaleHeight     =   6795
      ScaleWidth      =   4275
      TabIndex        =   5
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   4
      Top             =   7200
      Width           =   4335
   End
   Begin VB.CommandButton cmdPlayerSearch 
      Caption         =   "Search For Desired Player"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display Draft Order According To Last Year's Rankings"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdRead2 
      Caption         =   "Load Players Available In The Draft"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdRead1 
      Caption         =   "Load Each Team And Their Ranking"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "PreDraft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NFL Draft by Justin Buysse and Pete Larson. (NFLPOSTDRAFT.vbp)
'November 6th, 2007
'Project Objective: Our project will attempt to simulate the first round of the
    'NFL draft.  Each of the 32 NFL teams is represented and there are 32 available
    'players.  The assumption is made that each team will select only once.  There
    'is also a post draft feature that allows for a negotiation between Owner/GM
    'and Agent/Player with the Agent/Player ultimately winning the negotiation.
    'Overall, the project allows for experimentation within the boundaries of
    'our draft assumptions in order to have some fun in being different teams
    'and selecting different players.
'Form Objective: This form loads the available teams and players.  It allows the user
    'to search for different players to see if they applied for our draft.  There are
    'different players that play different positions that are from different schools.
    'There is also one fun fact associated with each player.  This form has added
    'commentary from ESPN Draft Analyst Mel Kiper Jr.
Option Explicit
Private Sub cmdDisplay_Click()
'This button takes the teams and their ranking from the array and sorts them in reverse order according to those rankings.
'This means that the team ranked #32 will pick #1, #31 will pick #2, #30 will pick #3 etc..
    cmdDisplay.Enabled = False
    cmdPlayerSearch.Enabled = True
    cmdSortform.Enabled = True
    Dim Pass As Integer
    Dim Pos As Integer
    Dim TempTeam As String
    Dim TempRank As Integer
    For Pass = 1 To (CTR_Teams - 1)
        For Pos = 1 To (CTR_Teams - Pass)
            If Rank(Pos) < Rank(Pos + 1) Then
                TempRank = Rank(Pos)
                Rank(Pos) = Rank(Pos + 1)
                Rank(Pos + 1) = TempRank
                
                TempTeam = Team(Pos)
                Team(Pos) = Team(Pos + 1)
                Team(Pos + 1) = TempTeam
            End If
        Next Pos
    Next Pass
        picResult1.Print "Pick"; Tab(10); "Team Name"; Tab(27); "Last Year's Final Ranking"
        picResult1.Print "***************************************************************"
    For Pos = 1 To CTR_Teams
        picResult1.Print (Pos); ":"; Team(Pos), Rank(Pos)
    Next Pos
End Sub
Private Sub cmdPlayerSearch_Click()
'This button allows the user to search for a player and see if they are included in our data file with the top 32 players.
'If the player being searched for is not in our data file, a message will be printed that they are not in our draft.
'The search for the player is case sensitive.  This is useful in that only the spelling needs to be correct in order to successfully find the desired player.
'This button also displays the searched for player's position, school, and one fun fact.
    Dim SPlayer As String
    SPlayer = InputBox("Please search for a player", "Player Search", , 1000, 1000)
    Dim Pos As Integer
    Dim Found As Boolean
    Pos = 0
    Found = False
    Do While (Found = False And Pos < CTR_Players)
        Pos = Pos + 1
        If LCase(Player(Pos)) = LCase(SPlayer) Then
            Found = True
        End If
    Loop
    picResult2.Cls
    picResult3.Cls
    If Found = True Then
        picResult2.Print Player(Pos); " has entered the draft "; "and is projected to be drafted #"; (Pos); "in this year's draft"
        picResult2.Print Player(Pos); " is a "; Position(Pos); " from " & School(Pos)
        picResult2.Print Player(Pos); " " & Fact(Pos)
        picResult3.Print "According to ESPN analyst Mel Kiper Jr."
    Else
        picResult2.Print SPlayer; " is not in the draft"
        picResult3.Print "According to ESPN analyst Mel Kiper Jr."
    End If
End Sub
Private Sub cmdQuit_Click()
'This button is available in case the user would like to end the program at anytime.
    End
End Sub
Private Sub cmdRead1_Click()
'This button will read all 32 NFL teams and the ranking that we gave them into an array.  We receive all of this information from a data file.
'When this button is hit, the next button that will be available to be selected will be the one we want to be available.
    cmdRead1.Enabled = False
    cmdDisplay.Enabled = False
    cmdRead2.Enabled = True
    cmdPlayerSearch.Enabled = False
    cmdSortform.Enabled = False
    CTR_Teams = 0
    Open App.Path & "\TeamsRanks.txt" For Input As #1
    Do Until EOF(1)
        CTR_Teams = CTR_Teams + 1
        Input #1, Team(CTR_Teams), Rank(CTR_Teams)
    Loop
    Close #1
End Sub
Private Sub cmdRead2_Click()
'This button will read the 32 players and their position, school, plus one fun fact
'if they are eligible in the first round of our draft into an array from a data file.
'When this button is hit, only the correct button in the next step will be available.
    cmdDisplay.Enabled = True
    cmdRead2.Enabled = False
    cmdPlayerSearch.Enabled = False
    cmdSortform.Enabled = False
    CTR_Players = 0
    Open App.Path & "\NFLdraft.txt" For Input As #2
    Do Until EOF(2)
        CTR_Players = CTR_Players + 1
        Input #2, Player(CTR_Players), Position(CTR_Players), School(CTR_Players), Fact(CTR_Players)
        Selected(CTR_Players) = "NONE"
    Loop
    Close #2
End Sub
Private Sub cmdSortform_Click()
'This button will hide the Pre-Draft form and move our project on to second step which is the Draft form.
    Draft.Show
    PreDraft.Hide
End Sub

