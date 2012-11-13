VERSION 5.00
Begin VB.Form frmStatfinder 
   BackColor       =   &H000000FF&
   Caption         =   "Player Finder"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   Picture         =   "Player Finder.frx":0000
   ScaleHeight     =   9660
   ScaleWidth      =   13170
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picStatSwitch 
      BackColor       =   &H8000000E&
      Height          =   252
      Left            =   5160
      ScaleHeight     =   195
      ScaleWidth      =   3075
      TabIndex        =   14
      Top             =   3000
      Width           =   3132
   End
   Begin VB.CommandButton cmdSwitchconf 
      Caption         =   "Work with a Player's MIAC Conference Stats"
      Height          =   732
      Left            =   10800
      TabIndex        =   12
      Top             =   2520
      Width           =   1692
   End
   Begin VB.CommandButton cmdSwitchOverall 
      Caption         =   "Work with a Player's Overall Stats"
      Height          =   732
      Left            =   8760
      TabIndex        =   11
      Top             =   2520
      Width           =   1692
   End
   Begin VB.CommandButton cmdFindname 
      Caption         =   "Player Identifier"
      Height          =   732
      Left            =   600
      TabIndex        =   10
      Top             =   6480
      Width           =   1692
   End
   Begin VB.TextBox txtPlayerNumb 
      Height          =   288
      Left            =   600
      TabIndex        =   8
      Top             =   6000
      Width           =   732
   End
   Begin VB.PictureBox picStatfinder 
      BackColor       =   &H8000000E&
      Height          =   2412
      Left            =   120
      ScaleHeight     =   2355
      ScaleWidth      =   12795
      TabIndex        =   7
      Top             =   3360
      Width           =   12855
   End
   Begin VB.PictureBox picLogo 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   4440
      Picture         =   "Player Finder.frx":11E218
      ScaleHeight     =   2115
      ScaleWidth      =   4515
      TabIndex        =   4
      Top             =   0
      Width           =   4575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   732
      Left            =   10800
      TabIndex        =   3
      Top             =   6480
      Width           =   1692
   End
   Begin VB.CommandButton cmdStatSorter 
      Caption         =   "Go to Stat Sorter"
      Height          =   732
      Left            =   8760
      TabIndex        =   2
      Top             =   6480
      Width           =   1692
   End
   Begin VB.CommandButton cmdPlayerAvg 
      Caption         =   "Get a Player's Season Averages and Career Totals"
      Height          =   732
      Left            =   2640
      TabIndex        =   1
      Top             =   2520
      Width           =   2052
   End
   Begin VB.CommandButton cmdSearchPlay 
      Caption         =   "Search for Player Stats"
      Height          =   732
      Left            =   600
      TabIndex        =   0
      Top             =   2520
      Width           =   1692
   End
   Begin VB.Label lblconfrecord 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "MIAC Record:11-3-2"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   16
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lblrecord 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Overall 2005-06 record:16-7-3"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   15
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label lblStatSwitch 
      Caption         =   "You are currently working with:"
      Height          =   255
      Left            =   5160
      TabIndex        =   13
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label lblnumbersearch 
      Caption         =   "Don't know a player's name? Type in his number and click the ""Player Identifier"" to find out"
      Height          =   372
      Left            =   1440
      TabIndex        =   9
      Top             =   6000
      Width           =   3612
   End
   Begin VB.Label lblTitle2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "2005-2006 St. John's Hockey"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   6
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Statistical Searcher and Sorter"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9840
      TabIndex        =   5
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmStatfinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Project Name: SJU Hockey Statistical Searcher and Sorter
'Form Name: frmStatfinder
'Author: Jeff Brown
'Date Written: 3/10/2006
'Objective: This is a program that will allow the user to search for individual stats for this hockey season
'Or to search for a particular player's career numbers and season averages
'The program also lets the user go to another form and sort 2005-06 stats by particular categories

Dim StatSwitch As String

Private Sub cmdSearchPlay_Click()
    'declare variables
    Dim GamePlay, Goal, Assist, Points, PlusMinus, PIM, GWG, PPG, SHG, Shots, Shutout, ShotAgainst, Saves, Wins, Losses, Mins, Ties, GoalAgainst As Integer
    Dim GAA, Savepct, ShotPct, PtsAGame As Single
    Dim PlayerName As String
    
    'clear picture box
    picStatfinder.Cls
    
    'Get player last name from user with input box
    PlayerName = InputBox("Enter a Johnnie Hockey Player Last Name or Roster Number:", "Enter Player Name")
    
    'Display error message if user chooses to search for a player but does not indicate which stats to search through
    If cmdSwitchOverall.Enabled = True And cmdSwitchConf.Enabled = True Then
        MsgBox "You must select either 'Overall Stats' or 'MIAC Stats' before searching for a player.", , "Search Error"
    End If
    
    'Use if-then statements to determine which player name user wants and output that name
    If cmdSwitchOverall.Enabled = False Then
    'Indicates that Overall statistics have been selected
        If PlayerName = "Speidel" Or PlayerName = "speidel" Or PlayerName = "1" Then
            picStatfinder.Print "#1 Nate Speidel, Freshman Goaltender"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print "Nate Speidel has not participated in any games this season"
        ElseIf PlayerName = "Swan" Or PlayerName = "swan" Or PlayerName = "2" Then
            'assign values to variables
            GamePlay = 25
            Goal = 5
            Assist = 9
            'calculate points and shot percentage
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 48
            ShotPct = Goal / Shots
            PlusMinus = 12
            PPG = 1
            GWG = 1
            SHG = 0
            PIM = 12
            'output information
            picStatfinder.Print "#2 Jordan Swan, Sophomore Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Meinz" Or PlayerName = "meinz" Or PlayerName = "3" Then
            GamePlay = 17
            Goal = 0
            Assist = 4
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 19
            ShotPct = Goal / Shots
            PlusMinus = 7
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 10
            picStatfinder.Print "#3 Nate Meinz, Sophomore Defenseman"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Czech" Or PlayerName = "czech" Or PlayerName = "4" Then
            GamePlay = 24
            Goal = 0
            Assist = 6
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 14
            ShotPct = Goal / Shots
            PlusMinus = 11
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 26
            picStatfinder.Print "#4 Matt Czech, Senior Defenseman"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Dorr" Or PlayerName = "dorr" Or PlayerName = "5" Then
            GamePlay = 2
            Goal = 0
            Assist = 0
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 0
            ShotPct = 0
            PlusMinus = 0
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 0
            picStatfinder.Print "#5 Sam Dorr, Sophomore Defenseman"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Mercado" Or PlayerName = "mercado" Or PlayerName = "6" Then
            GamePlay = 22
            Goal = 3
            Assist = 7
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 39
            ShotPct = Goal / Shots
            PlusMinus = 18
            PPG = 0
            GWG = 1
            SHG = 0
            PIM = 18
            picStatfinder.Print "#6 Dustin Mercado, Junior Defenseman"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Bjorklund" Or PlayerName = "bjorklund" Or PlayerName = "7" Then
            GamePlay = 26
            Goal = 11
            Assist = 16
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 57
            ShotPct = Goal / Shots
            PlusMinus = 9
            PPG = 4
            GWG = 0
            SHG = 2
            PIM = 24
            picStatfinder.Print "#7 Alternate Captain Scott Bjorklund, Senior Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Eagles" Or PlayerName = "Eagles" Or PlayerName = "8" Then
            GamePlay = 26
            Goal = 6
            Assist = 13
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 29
            ShotPct = Goal / Shots
            PlusMinus = 12
            PPG = 3
            GWG = 2
            SHG = 0
            PIM = 14
            picStatfinder.Print "#8 Pat Eagles, Sophomore Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Getchell" Or PlayerName = "getchell" Or PlayerName = "9" Then
            GamePlay = 25
            Goal = 7
            Assist = 12
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 65
            ShotPct = Goal / Shots
            PlusMinus = 2
            PPG = 4
            GWG = 0
            SHG = 0
            PIM = 10
            picStatfinder.Print "#9 Alternate Captain Aaron Getchell, Senior Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Freeman" Or PlayerName = "freeman" Or PlayerName = "10" Then
            GamePlay = 23
            Goal = 15
            Assist = 8
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 70
            ShotPct = Goal / Shots
            PlusMinus = 17
            PPG = 2
            GWG = 4
            SHG = 0
            PIM = 8
            picStatfinder.Print "#10 Tom Freeman, Sophomore Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Wocken" Or PlayerName = "wocken" Or PlayerName = "11" Then
            GamePlay = 21
            Goal = 1
            Assist = 2
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 19
            ShotPct = Goal / Shots
            PlusMinus = 13
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 8
            picStatfinder.Print "#11 Matt Wocken, Junior Defenseman"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Paul" Or PlayerName = "paul" Or PlayerName = "12" Then
            GamePlay = 2
            Goal = 0
            Assist = 1
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 2
            ShotPct = Goal / Shots
            PlusMinus = 1
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 2
            picStatfinder.Print "#12 Scott Paul, Freshman Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Wheeler" Or PlayerName = "wheeler" Or PlayerName = "14" Then
            GamePlay = 5
            Goal = 0
            Assist = 0
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 0
            ShotPct = 0
            PlusMinus = 0
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 4
            picStatfinder.Print "#14 Lance Wheeler, Freshman Defenseman"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Williams" Or PlayerName = "williams" Or PlayerName = "15" Then
            GamePlay = 25
            Goal = 8
            Assist = 13
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 63
            ShotPct = Goal / Shots
            PlusMinus = 12
            PPG = 2
            GWG = 2
            SHG = 0
            PIM = 14
            picStatfinder.Print "#15 Blake Williams, Junior Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Levar" Or PlayerName = "levar" Or PlayerName = "17" Then
            GamePlay = 18
            Goal = 3
            Assist = 0
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 36
            ShotPct = Goal / Shots
            PlusMinus = 6
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 20
            picStatfinder.Print "#17 Nick Levar, Senior Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Weigel" Or PlayerName = "Weigel" Or PlayerName = "18" Then
            GamePlay = 18
            Goal = 4
            Assist = 3
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 29
            ShotPct = Goal / Shots
            PlusMinus = 5
            PPG = 1
            GWG = 0
            SHG = 0
            PIM = 18
            picStatfinder.Print "#18 Jason Weigel, Sophomore Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Hartman" Or PlayerName = "hartman" Or PlayerName = "19" Then
            GamePlay = 18
            Goal = 1
            Assist = 4
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 12
            ShotPct = Goal / Shots
            PlusMinus = 8
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 6
            picStatfinder.Print "#19 Tom Hartman, Junior Defenseman"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Zemple" Or PlayerName = "zemple" Or PlayerName = "21" Then
            GamePlay = 22
            Goal = 1
            Assist = 5
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 10
            ShotPct = Goal / Shots
            PlusMinus = 10
            PPG = 0
            GWG = 1
            SHG = 0
            PIM = 26
            picStatfinder.Print "#21 Greg Zemple, Senior Defenseman"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Luger" Or PlayerName = "luger" Or PlayerName = "22" Then
            GamePlay = 26
            Goal = 7
            Assist = 8
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 59
            ShotPct = Goal / Shots
            PlusMinus = 6
            PPG = 2
            GWG = 0
            SHG = 0
            PIM = 77
            picStatfinder.Print "#22 Bille Luger, Senior Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Hipp" Or PlayerName = "hipp" Or PlayerName = "23" Then
            GamePlay = 25
            Goal = 9
            Assist = 10
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 59
            ShotPct = Goal / Shots
            PlusMinus = 10
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 38
            picStatfinder.Print "#23 Jake Hipp, Freshman Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Wild" Or PlayerName = "wild" Or PlayerName = "24" Then
            GamePlay = 26
            Goal = 3
            Assist = 8
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 41
            ShotPct = Goal / Shots
            PlusMinus = 8
            PPG = 0
            GWG = 0
            SHG = 1
            PIM = 8
            picStatfinder.Print "#24 Justin Wild, Junior Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Baker" Or PlayerName = "baker" Or PlayerName = "25" Then
            GamePlay = 6
            Goal = 1
            Assist = 1
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 6
            ShotPct = Goal / Shots
            PlusMinus = -1
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 6
            picStatfinder.Print "#25 Brian Baker, Freshman Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Smoleroff" Or PlayerName = "smoleroff" Or PlayerName = "26" Then
            GamePlay = 26
            Goal = 7
            Assist = 16
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 58
            ShotPct = Goal / Shots
            PlusMinus = 12
            PPG = 3
            GWG = 2
            SHG = 1
            PIM = 22
            picStatfinder.Print "#26 Captain Darryl Smoleroff, Senior Defenseman"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Ross" Or PlayerName = "ross" Or PlayerName = "27" Then
            GamePlay = 20
            Goal = 6
            Assist = 5
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 36
            ShotPct = Goal / Shots
            PlusMinus = 12
            PPG = 1
            GWG = 3
            SHG = 0
            PIM = 27
            picStatfinder.Print "#27 Ian Ross, Junior Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Langenbrunner" Or PlayerName = "langenbrunner" Or PlayerName = "28" Then
            GamePlay = 25
            Goal = 1
            Assist = 13
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 57
            ShotPct = Goal / Shots
            PlusMinus = 1
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 12
            picStatfinder.Print "#28 Ryan Langenbrunner, Senior Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Gardeski" Or PlayerName = "gardeski" Or PlayerName = "29" Then
            picStatfinder.Print "#29 Chris Gardeski, Sophomore Goaltender"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print "Chris Gardeski has not participated in any games this season"
        ElseIf PlayerName = "Hanna" Or PlayerName = "hanna" Or PlayerName = "30" Then
            GamePlay = 26
            Wins = 16
            Losses = 7
            Ties = 3
            Mins = 1575
            ShotAgainst = 665
            Saves = 612
            GoalAgainst = 53
            Savepct = (ShotAgainst - GoalAgainst) / ShotAgainst
            GAA = (GoalAgainst * 60) / Mins
            Shutout = 3
            picStatfinder.Print "#30 Adam Hanna, Senior Goaltender"
            picStatfinder.Print "GP", "W", "L", "T", "Minutes", "SA", "Saves", "GA", "Save %", "GAA", "SO"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Wins, Losses, Ties, Mins, ShotAgainst, Saves, GoalAgainst, FormatPercent(Savepct, 1), FormatNumber(GAA, 2), Shutout
        Else
            'if user enters incorrect roster number or name, display error message
            MsgBox "Sorry, you must enter a vaild roster name or number. Please try again. If necessary, check the 'Player Identifer' to find a player's proper name.", , "Error"
        End If
    End If
    
    'Conducts same stat search as previous if-then statement, except with conference stats
    If cmdSwitchConf.Enabled = False Then
        If PlayerName = "Speidel" Or PlayerName = "speidel" Or PlayerName = "1" Then
            picStatfinder.Print "#1 Nate Speidel, Freshman Goaltender"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print "Nate Speidel has not participated in any games this season"
        ElseIf PlayerName = "Swan" Or PlayerName = "swan" Or PlayerName = "2" Then
            'assign values to variables
            GamePlay = 16
            Goal = 5
            Assist = 8
            'calculate points and shot percentage
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 36
            ShotPct = Goal / Shots
            PlusMinus = 9
            PPG = 1
            GWG = 1
            SHG = 0
            PIM = 8
            'output information
            picStatfinder.Print "#2 Jordan Swan, Sophomore Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Meinz" Or PlayerName = "meinz" Or PlayerName = "3" Then
            GamePlay = 9
            Goal = 0
            Assist = 1
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 9
            ShotPct = Goal / Shots
            PlusMinus = 5
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 8
            picStatfinder.Print "#3 Nate Meinz, Sophomore Defenseman"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Czech" Or PlayerName = "czech" Or PlayerName = "4" Then
            GamePlay = 15
            Goal = 0
            Assist = 6
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 10
            ShotPct = Goal / Shots
            PlusMinus = 11
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 12
            picStatfinder.Print "#4 Matt Czech, Senior Defenseman"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Dorr" Or PlayerName = "dorr" Or PlayerName = "5" Then
            picStatfinder.Print "#5 Sam Dorr, Sophomore Defenseman"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print "Sam Dorr has not participated in any conference games this season"
        ElseIf PlayerName = "Mercado" Or PlayerName = "mercado" Or PlayerName = "6" Then
            GamePlay = 14
            Goal = 1
            Assist = 6
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 24
            ShotPct = Goal / Shots
            PlusMinus = 14
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 14
            picStatfinder.Print "#6 Dustin Mercado, Junior Defenseman"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Bjorklund" Or PlayerName = "bjorklund" Or PlayerName = "7" Then
            GamePlay = 16
            Goal = 10
            Assist = 14
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 42
            ShotPct = Goal / Shots
            PlusMinus = 12
            PPG = 3
            GWG = 0
            SHG = 2
            PIM = 18
            picStatfinder.Print "#7 Alternate Captain Scott Bjorklund, Senior Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Eagles" Or PlayerName = "Eagles" Or PlayerName = "8" Then
            GamePlay = 16
            Goal = 4
            Assist = 10
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 16
            ShotPct = Goal / Shots
            PlusMinus = 11
            PPG = 1
            GWG = 1
            SHG = 0
            PIM = 10
            picStatfinder.Print "#8 Pat Eagles, Sophomore Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Getchell" Or PlayerName = "getchell" Or PlayerName = "9" Then
            GamePlay = 15
            Goal = 4
            Assist = 10
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 32
            ShotPct = Goal / Shots
            PlusMinus = 4
            PPG = 2
            GWG = 0
            SHG = 0
            PIM = 6
            picStatfinder.Print "#9 Alternate Captain Aaron Getchell, Senior Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Freeman" Or PlayerName = "freeman" Or PlayerName = "10" Then
            GamePlay = 16
            Goal = 11
            Assist = 6
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 47
            ShotPct = Goal / Shots
            PlusMinus = 12
            PPG = 2
            GWG = 3
            SHG = 0
            PIM = 8
            picStatfinder.Print "#10 Tom Freeman, Sophomore Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Wocken" Or PlayerName = "wocken" Or PlayerName = "11" Then
            GamePlay = 12
            Goal = 1
            Assist = 2
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 14
            ShotPct = Goal / Shots
            PlusMinus = 13
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 6
            picStatfinder.Print "#11 Matt Wocken, Junior Defenseman"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Paul" Or PlayerName = "paul" Or PlayerName = "12" Then
            GamePlay = 1
            Goal = 0
            Assist = 1
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 1
            ShotPct = Goal / Shots
            PlusMinus = 1
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 2
            picStatfinder.Print "#12 Scott Paul, Freshman Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Wheeler" Or PlayerName = "wheeler" Or PlayerName = "14" Then
            GamePlay = 4
            Goal = 0
            Assist = 0
            Points = Goal + Assist
            PtsAGame = 0
            Shots = 0
            ShotPct = 0
            PlusMinus = 0
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 4
            picStatfinder.Print "#14 Lance Wheeler, Freshman Defenseman"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Williams" Or PlayerName = "williams" Or PlayerName = "15" Then
            GamePlay = 16
            Goal = 6
            Assist = 10
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 38
            ShotPct = Goal / Shots
            PlusMinus = 10
            PPG = 2
            GWG = 1
            SHG = 0
            PIM = 12
            picStatfinder.Print "#15 Blake Williams, Junior Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Levar" Or PlayerName = "levar" Or PlayerName = "17" Then
            GamePlay = 12
            Goal = 2
            Assist = 0
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 22
            ShotPct = Goal / Shots
            PlusMinus = 4
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 18
            picStatfinder.Print "#17 Nick Levar, Senior Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Weigel" Or PlayerName = "Weigel" Or PlayerName = "18" Then
            GamePlay = 10
            Goal = 3
            Assist = 3
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 17
            ShotPct = Goal / Shots
            PlusMinus = 7
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 10
            picStatfinder.Print "#18 Jason Weigel, Sophomore Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Hartman" Or PlayerName = "hartman" Or PlayerName = "19" Then
            GamePlay = 12
            Goal = 1
            Assist = 4
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 6
            ShotPct = Goal / Shots
            PlusMinus = 8
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 6
            picStatfinder.Print "#19 Tom Hartman, Junior Defenseman"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Zemple" Or PlayerName = "zemple" Or PlayerName = "21" Then
            GamePlay = 14
            Goal = 1
            Assist = 4
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 9
            ShotPct = Goal / Shots
            PlusMinus = 11
            PPG = 0
            GWG = 1
            SHG = 0
            PIM = 18
            picStatfinder.Print "#21 Greg Zemple, Senior Defenseman"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Luger" Or PlayerName = "luger" Or PlayerName = "22" Then
            GamePlay = 16
            Goal = 4
            Assist = 5
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 41
            ShotPct = Goal / Shots
            PlusMinus = 4
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 61
            picStatfinder.Print "#22 Bille Luger, Senior Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Hipp" Or PlayerName = "hipp" Or PlayerName = "23" Then
            GamePlay = 16
            Goal = 8
            Assist = 7
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 41
            ShotPct = Goal / Shots
            PlusMinus = 10
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 26
            picStatfinder.Print "#23 Jake Hipp, Freshman Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Wild" Or PlayerName = "wild" Or PlayerName = "24" Then
            GamePlay = 16
            Goal = 3
            Assist = 6
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 26
            ShotPct = Goal / Shots
            PlusMinus = 6
            PPG = 0
            GWG = 0
            SHG = 1
            PIM = 4
            picStatfinder.Print "#24 Justin Wild, Junior Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Baker" Or PlayerName = "baker" Or PlayerName = "25" Then
            GamePlay = 4
            Goal = 1
            Assist = 1
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 3
            ShotPct = Goal / Shots
            PlusMinus = -1
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 0
            picStatfinder.Print "#25 Brian Baker, Freshman Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Smoleroff" Or PlayerName = "smoleroff" Or PlayerName = "26" Then
            GamePlay = 16
            Goal = 6
            Assist = 11
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 29
            ShotPct = Goal / Shots
            PlusMinus = 12
            PPG = 3
            GWG = 2
            SHG = 1
            PIM = 10
            picStatfinder.Print "#26 Captain Darryl Smoleroff, Senior Defenseman"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Ross" Or PlayerName = "ross" Or PlayerName = "27" Then
            GamePlay = 13
            Goal = 5
            Assist = 4
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 25
            ShotPct = Goal / Shots
            PlusMinus = 10
            PPG = 1
            GWG = 2
            SHG = 0
            PIM = 23
            picStatfinder.Print "#27 Ian Ross, Junior Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Langenbrunner" Or PlayerName = "langenbrunner" Or PlayerName = "28" Then
            GamePlay = 15
            Goal = 1
            Assist = 11
            Points = Goal + Assist
            PtsAGame = Points / GamePlay
            Shots = 38
            ShotPct = Goal / Shots
            PlusMinus = 6
            PPG = 0
            GWG = 0
            SHG = 0
            PIM = 10
            picStatfinder.Print "#28 Ryan Langenbrunner, Senior Forward"
            picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        ElseIf PlayerName = "Gardeski" Or PlayerName = "gardeski" Or PlayerName = "29" Then
            picStatfinder.Print "#29 Chris Gardeski, Sophomore Goaltender"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print "Chris Gardeski has not participated in any games this season"
        ElseIf PlayerName = "Hanna" Or PlayerName = "hanna" Or PlayerName = "30" Then
            GamePlay = 16
            Wins = 11
            Losses = 3
            Ties = 2
            Mins = 970
            ShotAgainst = 383
            Saves = 348
            GoalAgainst = 35
            Savepct = (ShotAgainst - GoalAgainst) / ShotAgainst
            GAA = (GoalAgainst * 60) / Mins
            Shutout = 3
            picStatfinder.Print "#30 Adam Hanna, Senior Goaltender"
            picStatfinder.Print "GP", "W", "L", "T", "Minutes", "SA", "Saves", "GA", "Save %", "GAA", "SO"
            picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            picStatfinder.Print GamePlay, Wins, Losses, Ties, Mins, ShotAgainst, Saves, GoalAgainst, FormatPercent(Savepct, 1), FormatNumber(GAA, 2), Shutout
        Else
            MsgBox "Sorry, you must enter a vaild roster name or number. Please try again. If necessary, check the 'Player Identifer' to find a player's proper name.", , "Error"
        End If
    End If
End Sub

Private Sub cmdFindname_Click()
    'declare variable
    Dim PlayerNumb As String
    
    'assign variable to textbox
    PlayerNumb = txtPlayerNumb.Text
    
    'Use select case to identify a player by his number and then output that number
    Select Case PlayerNumb
        Case Is = 1
            MsgBox "#1 Nate Speidel, Goaltender", , "Player Name"
        Case Is = 2
            MsgBox "#2 Jordan Swan, Forward", , "Player Name"
        Case Is = 3
            MsgBox "#3 Nate Meinz, Defenseman", , "Player Name"
        Case Is = 4
            MsgBox "#4 Matt Czech, Defenseman", , "Player Name"
        Case Is = 5
            MsgBox "#5 Sam Dorr, Defenseman", , "Player Name"
        Case Is = 6
            MsgBox "#6 Dustin Mercado, Defenseman", , "Player Name"
        Case Is = 7
            MsgBox "#7 Scott Bjorklund, Forward", , "Player Name"
        Case Is = 8
            MsgBox "#8 Pat Eagles, Forward", , "Player Name"
        Case Is = 9
            MsgBox "#9 Aaron Getchell, Forward", , "Player Name"
        Case Is = 10
            MsgBox "#10 Tom Freeman, Forward", , "Player Name"
        Case Is = 11
            MsgBox "#11 Matt Wocken, Defenseman", , "Player Name"
        Case Is = 12
            MsgBox "#12 Scott Paul, Forward", , "Player Name"
        Case Is = 14
            MsgBox "#14 Lance Wheeler, Defenseman", , "Player Name"
        Case Is = 15
            MsgBox "#15 Blake Williams, Forward", , "Player Name"
        Case Is = 17
            MsgBox "#17 Nick Levar, Forward", , "Player Name"
        Case Is = 18
            MsgBox "#18 Jason Weigel, Forward", , "Player Name"
        Case Is = 19
            MsgBox "#19 Tom Hartman, Defenseman", , "Player Name"
        Case Is = 21
            MsgBox "#21 Greg Zemple, Defenseman", , "Player Name"
        Case Is = 22
            MsgBox "#22 Bille Luger, Forward", , "Player Name"
        Case Is = 23
            MsgBox "#23 Jake Hipp, Forward", , "Player Name"
        Case Is = 24
            MsgBox "#24 Justin Wild, Forward", , "Player Name"
        Case Is = 25
            MsgBox "#25 Brian Baker, Forward", , "Player Name"
        Case Is = 26
            MsgBox "#26 Darryl Smoleroff, Defenseman", , "Player Name"
        Case Is = 27
            MsgBox "#27 Ian Ross, Forward", , "Player Name"
        Case Is = 28
            MsgBox "#28 Ryan Langenbrunner, Forward", , "Player Name"
        Case Is = 29
            MsgBox "#29 Chris Gardeski, Goaltender", , "Player Name"
        Case Is = 30
            MsgBox "#30 Adam Hanna, Goaltender", , "Player Name"
        Case Else
            MsgBox "Sorry, you must enter a vaild roster number. All valid roster numbers are between 1 and 30. Please try again.", , "Error"

    End Select
    
End Sub
Private Sub cmdSwitchOverall_Click()
    'this button informs user that he/she is working with overall statistics
    'clear picture box
    picStatSwitch.Cls
    
    StatSwitch = "Overall Statistics"
    
    picStatSwitch.Print StatSwitch
    
    'disables overall switch and enables conference button
    cmdSwitchOverall.Enabled = False
    cmdSwitchConf.Enabled = True
    
End Sub

Private Sub cmdSwitchConf_Click()
    'this button informs user that he/she is working with conference statistics
    'clear picture box
    picStatSwitch.Cls
    
    StatSwitch = "MIAC Conference Statistics ONLY"
    
    picStatSwitch.Print StatSwitch
    
    'disables conference button and enables overall switch
    cmdSwitchOverall.Enabled = True
    cmdSwitchConf.Enabled = False
     
End Sub

Private Sub cmdPlayerAvg_Click()
    'this button allows user to search for a player's career totals and season averages
    'declare variables
    Dim GamePlay, Goal, Assist, Points, PlusMinus, PIM, GWG, PPG, SHG, Shots, Shutout, ShotAgainst, Saves, Wins, Losses, Mins, Ties, GoalAgainst As Integer
    Dim GAA, Savepct, ShotPct, PtsAGame As Single
    Dim CarGamePlay, CarGoal, CarAssist, CarPoints, CarPtsAGame, CarShotPct, CarPlusMinus, CarPIM, CarGWG, CarPPG, CarSHG, CarShots As Single
    Dim CarWins, CarLosses, CarTies, CarGoalAgainst, CarSaves, CarShotAgainst, CarShutout, CarMins, CarGAA, CarSavepct As Single
    Dim PlayerName As String
    
    cmdSwitchOverall.Enabled = True
    cmdSwitchConf.Enabled = True
    picStatSwitch.Cls
    
    StatSwitch = "Career Totals and Season Averages"
    picStatSwitch.Print StatSwitch
    
    'clear picture box
    picStatfinder.Cls
    
    'Get player last name from user with input box
    PlayerName = InputBox("Enter a Johnnie Hockey Player Last Name or Roster Number:", "Enter Player Name")
    
    'Use if-then statements to determine which player name user wants and output that name
    If PlayerName = "Speidel" Or PlayerName = "speidel" Or PlayerName = "1" Then
        picStatfinder.Print "#1 Nate Speidel, Freshman Goaltender"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print "Nate Speidel has not participated in any games"
    ElseIf PlayerName = "Swan" Or PlayerName = "swan" Or PlayerName = "2" Then
        'assign values to variables
        GamePlay = 31
        Goal = 5
        Assist = 9
        'calculate points and shot percentage
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 51
        ShotPct = Goal / Shots
        PlusMinus = 12
        PPG = 1
        GWG = 1
        SHG = 0
        PIM = 16
        CarGamePlay = GamePlay / 2
        CarGoal = Goal / 2
        CarAssist = Assist / 2
        CarPoints = Points / 2
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM / 2
        CarPlusMinus = PlusMinus / 2
        CarPPG = PPG / 2
        CarSHG = 0
        CarGWG = GWG / 2
        CarShots = Shots / 2
        'output information
        picStatfinder.Print "#2 Jordan Swan, Sophomore Forward"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Meinz" Or PlayerName = "meinz" Or PlayerName = "3" Then
        GamePlay = 42
        Goal = 2
        Assist = 8
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 40
        ShotPct = Goal / Shots
        PlusMinus = 22
        PPG = 0
        GWG = 0
        SHG = 0
        PIM = 26
        CarGamePlay = GamePlay / 2
        CarGoal = Goal / 2
        CarAssist = Assist / 2
        CarPoints = Points / 2
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM / 2
        CarPlusMinus = PlusMinus / 2
        CarPPG = 0
        CarSHG = 0
        CarGWG = 0
        CarShots = Shots / 2
        'output information
        picStatfinder.Print "#3 Nate Meinz, Sophomore Defenseman"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Czech" Or PlayerName = "czech" Or PlayerName = "4" Then
        GamePlay = 67
        Goal = 1
        Assist = 15
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 51
        ShotPct = Goal / Shots
        PlusMinus = 48
        PPG = 0
        GWG = 0
        SHG = 0
        PIM = 90
        CarGamePlay = GamePlay / 2
        CarGoal = Goal / 2
        CarAssist = Assist / 2
        CarPoints = Points / 2
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM / 2
        CarPlusMinus = PlusMinus / 2
        CarPPG = 0
        CarSHG = 0
        CarGWG = 0
        CarShots = Shots / 2
        'output information
        picStatfinder.Print "#4 Matt Czech, Senior Defenseman"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Dorr" Or PlayerName = "dorr" Or PlayerName = "5" Then
        GamePlay = 21
        Goal = 0
        Assist = 3
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 16
        ShotPct = 0
        PlusMinus = 13
        PPG = 0
        GWG = 0
        SHG = 0
        PIM = 25
        CarGamePlay = GamePlay / 2
        CarGoal = Goal / 2
        CarAssist = Assist / 2
        CarPoints = Points / 2
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM / 2
        CarPlusMinus = PlusMinus / 2
        CarPPG = 0
        CarSHG = 0
        CarGWG = 0
        CarShots = Shots / 2
        'output information
        picStatfinder.Print "#5 Sam Dorr, Sophomore Defenseman"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Mercado" Or PlayerName = "mercado" Or PlayerName = "6" Then
        GamePlay = 46
        Goal = 7
        Assist = 13
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 68
        ShotPct = Goal / Shots
        PlusMinus = 33
        PPG = 0
        GWG = 1
        SHG = 0
        PIM = 44
        CarGamePlay = GamePlay / 3
        CarGoal = Goal / 3
        CarAssist = Assist / 3
        CarPoints = Points / 3
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM / 3
        CarPlusMinus = PlusMinus / 3
        CarPPG = PPG / 3
        CarSHG = SHG / 3
        CarGWG = GWG / 3
        CarShots = Shots / 3
        'output information
        picStatfinder.Print "#6 Dustin Mercado, Junior Defenseman"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Bjorklund" Or PlayerName = "bjorklund" Or PlayerName = "7" Then
        GamePlay = 109
        Goal = 52
        Assist = 80
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 337
        ShotPct = Goal / Shots
        PlusMinus = 61
        PPG = 14
        GWG = 6
        SHG = 6
        PIM = 100
        CarGamePlay = GamePlay / 4
        CarGoal = Goal / 4
        CarAssist = Assist / 4
        CarPoints = Points / 4
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM / 4
        CarPlusMinus = PlusMinus / 4
        CarPPG = PPG / 4
        CarSHG = SHG / 4
        CarGWG = GWG / 4
        CarShots = Shots / 4
        'output information
        picStatfinder.Print "#7 Alternate Captain Scott Bjorklund, Senior Forward"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Eagles" Or PlayerName = "Eagles" Or PlayerName = "8" Then
        GamePlay = 54
        Goal = 17
        Assist = 31
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 81
        ShotPct = Goal / Shots
        PlusMinus = 32
        PPG = 6
        GWG = 3
        SHG = 1
        PIM = 16
        CarGamePlay = GamePlay / 2
        CarGoal = Goal / 2
        CarAssist = Assist / 2
        CarPoints = Points / 2
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM / 2
        CarPlusMinus = PlusMinus / 2
        CarPPG = PPG / 2
        CarSHG = SHG / 2
        CarGWG = GWG / 2
        CarShots = Shots / 2
        'output information
        picStatfinder.Print "#8 Pat Eagles, Sophomore Forward"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Getchell" Or PlayerName = "getchell" Or PlayerName = "9" Then
        GamePlay = 102
        Goal = 33
        Assist = 45
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 233
        ShotPct = Goal / Shots
        PlusMinus = 32
        PPG = 6
        GWG = 3
        SHG = 0
        PIM = 73
        CarGamePlay = GamePlay / 4
        CarGoal = Goal / 4
        CarAssist = Assist / 4
        CarPoints = Points / 4
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM / 4
        CarPlusMinus = PlusMinus / 4
        CarPPG = PPG / 4
        CarSHG = SHG / 4
        CarGWG = GWG / 4
        CarShots = Shots / 4
        'output information
        picStatfinder.Print "#9 Alternate Captain Aaron Getchell, Senior Forward"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Freeman" Or PlayerName = "freeman" Or PlayerName = "10" Then
        GamePlay = 28
        Goal = 16
        Assist = 8
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 73
        ShotPct = Goal / Shots
        PlusMinus = 17
        PPG = 2
        GWG = 4
        SHG = 0
        PIM = 8
        CarGamePlay = GamePlay / 2
        CarGoal = Goal / 2
        CarAssist = Assist / 2
        CarPoints = Points / 2
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM / 2
        CarPlusMinus = PlusMinus / 2
        CarPPG = PPG / 2
        CarSHG = SHG / 2
        CarGWG = GWG / 2
        CarShots = Shots / 2
        'output information
        picStatfinder.Print "#10 Tom Freeman, Sophomore Forward"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Wocken" Or PlayerName = "wocken" Or PlayerName = "11" Then
        GamePlay = 34
        Goal = 1
        Assist = 4
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 27
        ShotPct = Goal / Shots
        PlusMinus = 19
        PPG = 0
        GWG = 0
        SHG = 0
        PIM = 18
        CarGamePlay = GamePlay / 2
        CarGoal = Goal / 2
        CarAssist = Assist / 2
        CarPoints = Points / 2
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM / 2
        CarPlusMinus = PlusMinus / 2
        CarPPG = PPG / 2
        CarSHG = SHG / 2
        CarGWG = GWG / 2
        CarShots = Shots / 2
        'output information
        picStatfinder.Print "#11 Matt Wocken, Junior Defenseman"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Paul" Or PlayerName = "paul" Or PlayerName = "12" Then
        GamePlay = 2
        Goal = 0
        Assist = 1
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 2
        ShotPct = Goal / Shots
        PlusMinus = 1
        PPG = 0
        GWG = 0
        SHG = 0
        PIM = 2
        CarGamePlay = GamePlay
        CarGoal = Goal
        CarAssist = Assist
        CarPoints = Points
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM
        CarPlusMinus = PlusMinus
        CarPPG = PPG
        CarSHG = SHG
        CarGWG = GWG
        CarShots = Shots
        'output information
        picStatfinder.Print "#12 Scott Paul, Freshman Forward"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Wheeler" Or PlayerName = "wheeler" Or PlayerName = "14" Then
        GamePlay = 5
        Goal = 0
        Assist = 0
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 0
        ShotPct = 0
        PlusMinus = 0
        PPG = 0
        GWG = 0
        SHG = 0
        PIM = 4
        CarGamePlay = GamePlay
        CarGoal = Goal
        CarAssist = Assist
        CarPoints = Points
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM
        CarPlusMinus = PlusMinus
        CarPPG = PPG
        CarSHG = SHG
        CarGWG = GWG
        CarShots = Shots
        'output information
        picStatfinder.Print "#14 Lance Wheeler, Freshman Defenseman"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Williams" Or PlayerName = "williams" Or PlayerName = "15" Then
        GamePlay = 75
        Goal = 19
        Assist = 32
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 178
        ShotPct = Goal / Shots
        PlusMinus = 35
        PPG = 3
        GWG = 4
        SHG = 0
        PIM = 30
        CarGamePlay = GamePlay / 3
        CarGoal = Goal / 3
        CarAssist = Assist / 3
        CarPoints = Points / 3
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM / 3
        CarPlusMinus = PlusMinus / 3
        CarPPG = PPG / 3
        CarSHG = SHG / 3
        CarGWG = GWG / 3
        CarShots = Shots / 3
        'output information
        picStatfinder.Print "#15 Blake Williams, Junior Forward"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Levar" Or PlayerName = "levar" Or PlayerName = "17" Then
        GamePlay = 78
        Goal = 13
        Assist = 16
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 132
        ShotPct = Goal / Shots
        PlusMinus = 22
        PPG = 0
        GWG = 2
        SHG = 0
        PIM = 68
        CarGamePlay = GamePlay / 4
        CarGoal = Goal / 4
        CarAssist = Assist / 4
        CarPoints = Points / 4
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM / 4
        CarPlusMinus = PlusMinus / 4
        CarPPG = PPG / 4
        CarSHG = SHG / 4
        CarGWG = GWG / 4
        CarShots = Shots / 4
        'output information
        picStatfinder.Print "#17 Nick Levar, Senior Forward"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Weigel" Or PlayerName = "Weigel" Or PlayerName = "18" Then
        GamePlay = 14
        Goal = 4
        Assist = 3
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 29
        ShotPct = Goal / Shots
        PlusMinus = 5
        PPG = 1
        GWG = 0
        SHG = 0
        PIM = 18
        CarGamePlay = GamePlay
        CarGoal = Goal
        CarAssist = Assist
        CarPoints = Points
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM
        CarPlusMinus = PlusMinus
        CarPPG = PPG
        CarSHG = SHG
        CarGWG = GWG
        CarShots = Shots
        'output information
        picStatfinder.Print "#18 Jaso Weigel, Sophomore Forward"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Hartman" Or PlayerName = "hartman" Or PlayerName = "19" Then
        GamePlay = 36
        Goal = 1
        Assist = 7
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 27
        ShotPct = Goal / Shots
        PlusMinus = 18
        PPG = 0
        GWG = 0
        SHG = 0
        PIM = 10
        CarGamePlay = GamePlay / 3
        CarGoal = Goal / 3
        CarAssist = Assist / 3
        CarPoints = Points / 3
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM / 3
        CarPlusMinus = PlusMinus / 3
        CarPPG = PPG / 3
        CarSHG = SHG / 3
        CarGWG = GWG / 3
        CarShots = Shots / 3
        'output information
        picStatfinder.Print "#19 Tom Hartman, Junior Defenseman"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Zemple" Or PlayerName = "zemple" Or PlayerName = "21" Then
        GamePlay = 40
        Goal = 3
        Assist = 9
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 24
        ShotPct = Goal / Shots
        PlusMinus = 20
        PPG = 2
        GWG = 1
        SHG = 0
        PIM = 64
        CarGamePlay = GamePlay / 2
        CarGoal = Goal / 2
        CarAssist = Assist / 2
        CarPoints = Points / 2
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM / 2
        CarPlusMinus = PlusMinus / 2
        CarPPG = PPG / 2
        CarSHG = SHG / 2
        CarGWG = GWG / 2
        CarShots = Shots / 2
        'output information
        picStatfinder.Print "#21 Greg Zemple, Senior Defenseman"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Luger" Or PlayerName = "luger" Or PlayerName = "22" Then
        GamePlay = 53
        Goal = 16
        Assist = 18
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 107
        ShotPct = Goal / Shots
        PlusMinus = 14
        PPG = 2
        GWG = 4
        SHG = 1
        PIM = 97
        CarGamePlay = GamePlay / 2
        CarGoal = Goal / 2
        CarAssist = Assist / 2
        CarPoints = Points / 2
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM / 2
        CarPlusMinus = PlusMinus / 2
        CarPPG = PPG / 2
        CarSHG = SHG / 2
        CarGWG = GWG / 2
        CarShots = Shots / 2
        'output information
        picStatfinder.Print "#22 Bille Luger, Senior Forward"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Hipp" Or PlayerName = "hipp" Or PlayerName = "23" Then
        GamePlay = 25
        Goal = 9
        Assist = 10
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 59
        ShotPct = Goal / Shots
        PlusMinus = 10
        PPG = 0
        GWG = 0
        SHG = 0
        PIM = 38
        CarGamePlay = GamePlay
        CarGoal = Goal
        CarAssist = Assist
        CarPoints = Points
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM
        CarPlusMinus = PlusMinus
        CarPPG = PPG
        CarSHG = SHG
        CarGWG = GWG
        CarShots = Shots
        'output information
        picStatfinder.Print "#23 Jake Hipp, Freshman Forward"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Wild" Or PlayerName = "wild" Or PlayerName = "24" Then
        GamePlay = 65
        Goal = 7
        Assist = 19
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 90
        ShotPct = Goal / Shots
        PlusMinus = 23
        PPG = 0
        GWG = 0
        SHG = 1
        PIM = 20
        CarGamePlay = GamePlay / 3
        CarGoal = Goal / 3
        CarAssist = Assist / 3
        CarPoints = Points / 3
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM / 3
        CarPlusMinus = PlusMinus / 3
        CarPPG = PPG / 3
        CarSHG = SHG / 3
        CarGWG = GWG / 3
        CarShots = Shots / 3
        'output information
        picStatfinder.Print "#24 Justin Wild, Junior Forward"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Baker" Or PlayerName = "baker" Or PlayerName = "25" Then
        GamePlay = 6
        Goal = 1
        Assist = 1
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 6
        ShotPct = Goal / Shots
        PlusMinus = -1
        PPG = 0
        GWG = 0
        SHG = 0
        PIM = 6
        CarGamePlay = GamePlay
        CarGoal = Goal
        CarAssist = Assist
        CarPoints = Points
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM
        CarPlusMinus = PlusMinus
        CarPPG = PPG
        CarSHG = SHG
        CarGWG = GWG
        CarShots = Shots
        'output information
        picStatfinder.Print "#25 Brian Baker, Freshman Forward"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Smoleroff" Or PlayerName = "smoleroff" Or PlayerName = "26" Then
        GamePlay = 107
        Goal = 35
        Assist = 64
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 253
        ShotPct = Goal / Shots
        PlusMinus = 84
        PPG = 14
        GWG = 13
        SHG = 1
        PIM = 97
        CarGamePlay = GamePlay / 4
        CarGoal = Goal / 4
        CarAssist = Assist / 4
        CarPoints = Points / 4
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM / 4
        CarPlusMinus = PlusMinus / 4
        CarPPG = PPG / 4
        CarSHG = SHG / 4
        CarGWG = GWG / 4
        CarShots = Shots / 4
        'output information
        picStatfinder.Print "#26 Captain Darryl Smoleroff, Senior Defenseman"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Ross" Or PlayerName = "ross" Or PlayerName = "27" Then
        GamePlay = 36
        Goal = 9
        Assist = 10
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 64
        ShotPct = Goal / Shots
        PlusMinus = 17
        PPG = 3
        GWG = 4
        SHG = 0
        PIM = 43
        CarGamePlay = GamePlay / 2
        CarGoal = Goal / 2
        CarAssist = Assist / 2
        CarPoints = Points / 2
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM / 2
        CarPlusMinus = PlusMinus / 2
        CarPPG = PPG / 2
        CarSHG = SHG / 2
        CarGWG = GWG / 2
        CarShots = Shots / 2
        'output information
        picStatfinder.Print "#27 Ian Ross, Junior Forward"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Langenbrunner" Or PlayerName = "langenbrunner" Or PlayerName = "28" Then
        GamePlay = 107
        Goal = 41
        Assist = 81
        Points = Goal + Assist
        PtsAGame = Points / GamePlay
        Shots = 323
        ShotPct = Goal / Shots
        PlusMinus = 40
        PPG = 17
        GWG = 5
        SHG = 5
        PIM = 60
        CarGamePlay = GamePlay / 4
        CarGoal = Goal / 4
        CarAssist = Assist / 4
        CarPoints = Points / 4
        CarPtsAGame = CarPoints / CarGamePlay
        CarPIM = PIM / 4
        CarPlusMinus = PlusMinus / 4
        CarPPG = PPG / 4
        CarSHG = SHG / 4
        CarGWG = GWG / 4
        CarShots = Shots / 4
        'output information
        picStatfinder.Print "#28 Ryan Langenbrunner, Senior Forward"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Goal, Assist, Points, FormatNumber(PtsAGame, 2), Shots, FormatPercent(ShotPct, 1), PlusMinus, PIM, PPG, SHG, GWG
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "SHG", "GWG"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarGoal, 1), FormatNumber(CarAssist, 1), FormatNumber(CarPoints, 1), FormatNumber(PtsAGame, 2), FormatNumber(CarShots, 1), FormatPercent(ShotPct, 1), FormatNumber(CarPlusMinus, 1), FormatNumber(CarPIM, 1), FormatNumber(CarPPG, 1), FormatNumber(CarSHG, 1), FormatNumber(CarGWG, 1)
    ElseIf PlayerName = "Gardeski" Or PlayerName = "gardeski" Or PlayerName = "29" Then
        picStatfinder.Print "#29 Chris Gardeski, Sophomore Goaltender"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print "Chris Gardeski has not participated in any games"
    ElseIf PlayerName = "Hanna" Or PlayerName = "hanna" Or PlayerName = "30" Then
        GamePlay = 69
        Wins = 51
        Losses = 13
        Ties = 5
        Mins = 4176
        ShotAgainst = 1719
        Saves = 1596
        GoalAgainst = 123
        Savepct = (ShotAgainst - GoalAgainst) / ShotAgainst
        GAA = (GoalAgainst * 60) / Mins
        Shutout = 14
        CarGamePlay = GamePlay / 3
        CarWins = Wins / 3
        CarLosses = Losses / 3
        CarTies = Ties / 3
        CarMins = Mins / 3
        CarShotAgainst = ShotAgainst / 3
        CarSaves = Saves / 3
        CarGoalAgainst = GoalAgainst / 3
        CarSavepct = Savepct
        CarGAA = GAA
        CarShutout = Shutout / 3
        picStatfinder.Print "#30 Adam Hanna, Senior Goaltender"
        picStatfinder.Print "Career Totals"
        picStatfinder.Print "GP", "W", "L", "T", "Minutes", "SA", "Saves", "GA", "Save %", "GAA", "SO"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print GamePlay, Wins, Losses, Ties, Mins, ShotAgainst, Saves, GoalAgainst, FormatPercent(Savepct, 1), FormatNumber(GAA, 2), Shutout
        picStatfinder.Print "Season Averages"
        picStatfinder.Print "GP", "W", "L", "T", "Minutes", "SA", "Saves", "GA", "Save %", "GAA", "SO"
        picStatfinder.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatfinder.Print FormatNumber(CarGamePlay, 1), FormatNumber(CarWins, 1), FormatNumber(CarLosses, 1), FormatNumber(CarTies, 1), FormatNumber(CarMins, 1), FormatNumber(CarShotAgainst, 1), FormatNumber(CarSaves, 1), FormatNumber(CarGoalAgainst, 1), FormatPercent(CarSavepct, 1), FormatNumber(CarGAA, 2), FormatNumber(CarShutout, 1)
    Else
        MsgBox "Sorry, you must enter a vaild roster name or number. Please try again. If necessary, check the 'Player Identifer' to find a player's proper name.", , "Error"
    End If
End Sub

Private Sub cmdStatSorter_Click()
    'switches to sorting form
    frmStatfinder.Hide
    frmStatsorter.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

