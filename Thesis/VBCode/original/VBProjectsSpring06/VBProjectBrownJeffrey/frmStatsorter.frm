VERSION 5.00
Begin VB.Form frmStatsorter 
   BackColor       =   &H000000FF&
   Caption         =   "Player Stats Sorter"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   Picture         =   "frmStatsorter.frx":0000
   ScaleHeight     =   9855
   ScaleWidth      =   13185
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTeamLeader 
      BackColor       =   &H000000FF&
      Height          =   3015
      Left            =   13320
      ScaleHeight     =   2955
      ScaleWidth      =   1995
      TabIndex        =   32
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdSortGoalies 
      Caption         =   "Sort Goalie Stats"
      Height          =   732
      Left            =   2400
      TabIndex        =   26
      Top             =   9000
      Width           =   1692
   End
   Begin VB.PictureBox picSwitchStats 
      BackColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   4800
      ScaleHeight     =   195
      ScaleWidth      =   4155
      TabIndex        =   25
      Top             =   2520
      Width           =   4212
   End
   Begin VB.PictureBox picPlayerSwitch 
      BackColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   4800
      ScaleHeight     =   195
      ScaleWidth      =   4155
      TabIndex        =   24
      Top             =   1800
      Width           =   4212
   End
   Begin VB.CommandButton cmdsortstats 
      Caption         =   "Sort Skater Stats"
      Height          =   732
      Left            =   360
      TabIndex        =   23
      Top             =   9000
      Width           =   1692
   End
   Begin VB.PictureBox picLeaderName 
      BackColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   10680
      ScaleHeight     =   435
      ScaleWidth      =   2355
      TabIndex        =   22
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton cmdSwitchOverall 
      Caption         =   "Work with Overall Stats"
      Height          =   732
      Left            =   9480
      TabIndex        =   17
      Top             =   2280
      Width           =   1692
   End
   Begin VB.CommandButton cmdSwitchConf 
      Caption         =   "Work with MIAC Conference Stats"
      Height          =   732
      Left            =   11520
      TabIndex        =   16
      Top             =   2280
      Width           =   1692
   End
   Begin VB.CommandButton cmdSwitchGoalies 
      Caption         =   "Work with Goalie Stats"
      Height          =   732
      Left            =   2640
      TabIndex        =   15
      Top             =   2280
      Width           =   1692
   End
   Begin VB.CommandButton cmdSwitchSkaters 
      Caption         =   "Work with Skater Stats"
      Height          =   732
      Left            =   600
      TabIndex        =   14
      Top             =   2280
      Width           =   1692
   End
   Begin VB.OptionButton optGWG 
      BackColor       =   &H000000FF&
      Caption         =   "GWG"
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   9240
      TabIndex        =   12
      Top             =   10080
      Width           =   1212
   End
   Begin VB.OptionButton optPPG 
      BackColor       =   &H00FF0000&
      Caption         =   "PPG"
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   9240
      TabIndex        =   11
      Top             =   9600
      Width           =   1212
   End
   Begin VB.OptionButton optPIM 
      BackColor       =   &H00FF0000&
      Caption         =   "PIM"
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   8040
      TabIndex        =   10
      Top             =   10080
      Width           =   1212
   End
   Begin VB.OptionButton optPlusMinus 
      BackColor       =   &H000000FF&
      Caption         =   "+/-"
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   8040
      TabIndex        =   9
      Top             =   9600
      Width           =   1212
   End
   Begin VB.OptionButton optShotpct 
      BackColor       =   &H000000FF&
      Caption         =   "Shot %"
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   6840
      TabIndex        =   8
      Top             =   10080
      Width           =   1212
   End
   Begin VB.OptionButton optShots 
      BackColor       =   &H00FF0000&
      Caption         =   "Shots"
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   6840
      TabIndex        =   7
      Top             =   9600
      Width           =   1212
   End
   Begin VB.OptionButton optPtsAGame 
      BackColor       =   &H000000FF&
      Caption         =   "Points per Game"
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   5640
      TabIndex        =   6
      Top             =   9600
      Width           =   1212
   End
   Begin VB.OptionButton optPoints 
      BackColor       =   &H00FF0000&
      Caption         =   "Points"
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   5640
      TabIndex        =   5
      Top             =   10080
      Width           =   1212
   End
   Begin VB.OptionButton optAssists 
      BackColor       =   &H000000FF&
      Caption         =   "Assists"
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   4440
      TabIndex        =   4
      Top             =   10080
      Width           =   1212
   End
   Begin VB.OptionButton optGoals 
      BackColor       =   &H00FF0000&
      Caption         =   "Goals"
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   4440
      TabIndex        =   3
      Top             =   9600
      Width           =   1212
   End
   Begin VB.PictureBox picStatsorter 
      BackColor       =   &H00FFFFFF&
      Height          =   5415
      Left            =   240
      ScaleHeight     =   5355
      ScaleWidth      =   13515
      TabIndex        =   2
      Top             =   3240
      Width           =   13575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   732
      Left            =   11400
      TabIndex        =   1
      Top             =   9960
      Width           =   1692
   End
   Begin VB.CommandButton cmdStatfinder 
      Caption         =   "Go to Player Finder"
      Height          =   732
      Left            =   11400
      TabIndex        =   0
      Top             =   9000
      Width           =   1692
   End
   Begin VB.Label lblDirections4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4.)Click proper sorting button below"
      Height          =   252
      Left            =   120
      TabIndex        =   31
      Top             =   1080
      Width           =   2532
   End
   Begin VB.Label lblDirections3part2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "below to sort"
      Height          =   252
      Left            =   120
      TabIndex        =   30
      Top             =   840
      Width           =   2532
   End
   Begin VB.Label lblDirections3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3.)If using skaters, select a stat "
      Height          =   252
      Left            =   120
      TabIndex        =   29
      Top             =   600
      Width           =   2532
   End
   Begin VB.Label lblDirections2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2.)Select overall or MIAC stats"
      Height          =   252
      Left            =   120
      TabIndex        =   28
      Top             =   360
      Width           =   2532
   End
   Begin VB.Label lblDirections1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1.)Select goalies or skaters"
      Height          =   252
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   2532
   End
   Begin VB.Label lblTeamLeader 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The team leader in the category is:"
      Height          =   372
      Left            =   11520
      TabIndex        =   21
      Top             =   120
      Width           =   1572
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "2005-06 SJU Hockey Statistical Sorting Program"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2760
      TabIndex        =   20
      Top             =   120
      Width           =   8292
   End
   Begin VB.Label lblstatswitch 
      Caption         =   "You are currently working with:"
      Height          =   252
      Left            =   4800
      TabIndex        =   19
      Top             =   2280
      Width           =   4212
   End
   Begin VB.Label lblplayerposition 
      Caption         =   "You are currently working with:"
      Height          =   252
      Left            =   4800
      TabIndex        =   18
      Top             =   1560
      Width           =   4212
   End
   Begin VB.Label lblChooseStat 
      BackColor       =   &H00000000&
      Caption         =   $"frmStatsorter.frx":F29D6
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4680
      TabIndex        =   13
      Top             =   9000
      Width           =   5535
   End
End
Attribute VB_Name = "frmStatsorter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Project Name: SJU Hockey Statistical Searcher and Sorter
'Form Name: frmStatsorter
'Author: Jeff Brown
'Date Written: 3/10/2006
'Objective: This is the statistic sorting section of the program. The user can choose
'which stats and which positions to sort and the program will output them

Dim StatSwitch As String
'declare overall skater statistic arrays
Dim OSkaterName(1 To 24) As String
Dim OPosition(1 To 24) As String
Dim OGP(1 To 24) As Single
Dim OGoal(1 To 24) As Single
Dim OAssist(1 To 24) As Single
Dim OPoints(1 To 24) As Single
Dim OPtsAGame(1 To 24) As Single
Dim OShots(1 To 24) As Single
Dim OShotPct(1 To 24) As Single
Dim OPlusMinus(1 To 24) As Single
Dim OPIM(1 To 24) As Single
Dim OPPG(1 To 24) As Single
Dim OGWG(1 To 24) As Single

'declare MIAC skater statistic arrays
Dim MSkaterName(1 To 23) As String
Dim MPosition(1 To 23) As String
Dim MGP(1 To 23) As Single
Dim MGoal(1 To 23) As Single
Dim MAssist(1 To 23) As Single
Dim MPoints(1 To 23) As Single
Dim MPtsAGame(1 To 23) As Single
Dim MShots(1 To 23) As Single
Dim MShotPct(1 To 23) As Single
Dim MPlusMinus(1 To 23) As Single
Dim MPIM(1 To 23) As Single
Dim MPPG(1 To 23) As Single
Dim MGWG(1 To 23) As Single


Private Sub cmdsortstats_Click()
    'Declare variables
    Dim Pos, Pass, SizeMIAC, SizeOverall As Integer
    Dim TemOSkaterName, TemOPosition, TemMSkaterName, TemMPosition As String
    Dim TemOGP, TemOGoal, TemOAssist, TemOPoints, TemOPtsAGame, TemOShots, TemOShotPct, TemOPlusMinus, TemOPIM, TemOPPG, TemOGWG As Single
    Dim TemMGP, TemMGoal, TemMAssist, TemMPoints, TemMPtsAGame, TemMShots, TemMShotPct, TemMPlusMinus, TemMPIM, TemMPPG, TemMGWG As Single
    
    picStatsorter.Cls
    
    'open first text file
    Open App.Path & "\MIACSkaterStats.txt" For Input As #1
    
    Pos = 0
    
    'fill arrays with text file info
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, MSkaterName(Pos), MPosition(Pos), MGP(Pos), MGoal(Pos), MAssist(Pos), MPoints(Pos), MPtsAGame(Pos), MShots(Pos), MShotPct(Pos), MPlusMinus(Pos), MPIM(Pos), MPPG(Pos), MGWG(Pos)
    Loop
    
    SizeMIAC = Pos
    'close first text file
    Close #1

    
    'open second text file
    Open App.Path & "\OverallSkaterStats.txt" For Input As #2
    
    Pos = 0
    
    'fill arrays with text file info
    Do Until EOF(2)
        Pos = Pos + 1
        Input #2, OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), OPtsAGame(Pos), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OGWG(Pos)
    Loop
    
    SizeOverall = Pos
    
    'close second text file
    Close #2
    
    'clear team photo box
    picTeamLeader.Picture = LoadPicture("")
                        
    'display error message if no stat option selected
    If optGoals.Value = False And optAssists.Value = False And optGWG.Value = False And optPIM.Value = False And optPlusMinus.Value = False And optPoints.Value = False And optPPG.Value = False And optShots.Value = False And optPtsAGame.Value = False And optShotpct.Value = False Then
        MsgBox "Please select a skater stat to sort", , "Selection Error"
    End If
                        
   'If skaters and overall stats are selected and certain option buttons are selected, this code will determine which way to sort stats and output it
   If cmdSwitchOverall.Enabled = False And cmdSwitchSkaters.Enabled = False Then
        If optGoals.Value = True Then
            'sorting code for the overall statistical arrays
            'sorts in descending order for each stat, based on the option the user chooses
            For Pass = 1 To (SizeOverall - 1)
                For Pos = 1 To (SizeOverall - Pass)
                    If OGoal(Pos) < OGoal(Pos + 1) Then
                        TemOGoal = OGoal(Pos)
                        OGoal(Pos) = OGoal(Pos + 1)
                        OGoal(Pos + 1) = TemOGoal
                    
                        TemOSkaterName = OSkaterName(Pos)
                        OSkaterName(Pos) = OSkaterName(Pos + 1)
                        OSkaterName(Pos + 1) = TemOSkaterName
                        
                        TemOPosition = OPosition(Pos)
                        OPosition(Pos) = OPosition(Pos + 1)
                        OPosition(Pos + 1) = TemOPosition
                        
                        TemOGP = OGP(Pos)
                        OGP(Pos) = OGP(Pos + 1)
                        OGP(Pos + 1) = TemOGP
                        
                        TemOAssist = OAssist(Pos)
                        OAssist(Pos) = OAssist(Pos + 1)
                        OAssist(Pos + 1) = TemOAssist
                        
                        TemOPoints = OPoints(Pos)
                        OPoints(Pos) = OPoints(Pos + 1)
                        OPoints(Pos + 1) = TemOPoints
                        
                        TemOPtsAGame = OPtsAGame(Pos)
                        OPtsAGame(Pos) = OPtsAGame(Pos + 1)
                        OPtsAGame(Pos + 1) = TemOPtsAGame
                        
                        TemOShots = OShots(Pos)
                        OShots(Pos) = OShots(Pos + 1)
                        OShots(Pos + 1) = TemOShots
                        
                        TemOShotPct = OShotPct(Pos)
                        OShotPct(Pos) = OShotPct(Pos + 1)
                        OShotPct(Pos + 1) = TemOShotPct
                        
                        TemOPlusMinus = OPlusMinus(Pos)
                        OPlusMinus(Pos) = OPlusMinus(Pos + 1)
                        OPlusMinus(Pos + 1) = TemOPlusMinus
                        
                        TemOPIM = OPIM(Pos)
                        OPIM(Pos) = OPIM(Pos + 1)
                        OPIM(Pos + 1) = TemOPIM
                        
                        TemOPPG = OPPG(Pos)
                        OPPG(Pos) = OPPG(Pos + 1)
                        OPPG(Pos + 1) = TemOPPG
                        
                        TemOGWG = OGWG(Pos)
                        OGWG(Pos) = OGWG(Pos + 1)
                        OGWG(Pos + 1) = TemOGWG
                        'picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), OPtsAGame(Pos), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OSHG(Pos), OGWG(Pos)
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "Overall Skater Statstics have been sorted by goals scored"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            'print sorted results
            For Pos = 1 To SizeOverall
                picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), FormatNumber(OPtsAGame(Pos), 2), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OGWG(Pos)
            Next Pos
            picLeaderName.Cls
            'outputs team leader's picture
            picLeaderName.Print "#10 Tom Freeman, Forward"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\Freeman3.JPG")
        ElseIf optAssists.Value = True Then
            For Pass = 1 To (SizeOverall - 1)
                For Pos = 1 To (SizeOverall - Pass)
                    If OAssist(Pos) < OAssist(Pos + 1) Then
                        TemOAssist = OAssist(Pos)
                        OAssist(Pos) = OAssist(Pos + 1)
                        OAssist(Pos + 1) = TemOAssist
                    
                        TemOSkaterName = OSkaterName(Pos)
                        OSkaterName(Pos) = OSkaterName(Pos + 1)
                        OSkaterName(Pos + 1) = TemOSkaterName
                        
                        TemOPosition = OPosition(Pos)
                        OPosition(Pos) = OPosition(Pos + 1)
                        OPosition(Pos + 1) = TemOPosition
                        
                        TemOGP = OGP(Pos)
                        OGP(Pos) = OGP(Pos + 1)
                        OGP(Pos + 1) = TemOGP
                        
                        TemOGoal = OGoal(Pos)
                        OGoal(Pos) = OGoal(Pos + 1)
                        OGoal(Pos + 1) = TemOGoal
                        
                        TemOPoints = OPoints(Pos)
                        OPoints(Pos) = OPoints(Pos + 1)
                        OPoints(Pos + 1) = TemOPoints
                        
                        TemOPtsAGame = OPtsAGame(Pos)
                        OPtsAGame(Pos) = OPtsAGame(Pos + 1)
                        OPtsAGame(Pos + 1) = TemOPtsAGame
                        
                        TemOShots = OShots(Pos)
                        OShots(Pos) = OShots(Pos + 1)
                        OShots(Pos + 1) = TemOShots
                        
                        TemOShotPct = OShotPct(Pos)
                        OShotPct(Pos) = OShotPct(Pos + 1)
                        OShotPct(Pos + 1) = TemOShotPct
                        
                        TemOPlusMinus = OPlusMinus(Pos)
                        OPlusMinus(Pos) = OPlusMinus(Pos + 1)
                        OPlusMinus(Pos + 1) = TemOPlusMinus
                        
                        TemOPIM = OPIM(Pos)
                        OPIM(Pos) = OPIM(Pos + 1)
                        OPIM(Pos + 1) = TemOPIM
                        
                        TemOPPG = OPPG(Pos)
                        OPPG(Pos) = OPPG(Pos + 1)
                        OPPG(Pos + 1) = TemOPPG
                        
                        TemOGWG = OGWG(Pos)
                        OGWG(Pos) = OGWG(Pos + 1)
                        OGWG(Pos + 1) = TemOGWG
                        'picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), OPtsAGame(Pos), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OSHG(Pos), OGWG(Pos)
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "Overall Skater Statstics have been sorted by assists"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            For Pos = 1 To SizeOverall
                picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), FormatNumber(OPtsAGame(Pos), 2), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OGWG(Pos)
            Next Pos
            picLeaderName.Cls
            picLeaderName.Print "#7 Scott Bjorklund, Forward"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\bjorklund3.JPG")
        ElseIf optPoints.Value = True Then
            For Pass = 1 To (SizeOverall - 1)
                For Pos = 1 To (SizeOverall - Pass)
                    If OPoints(Pos) < OPoints(Pos + 1) Then
                        TemOPoints = OPoints(Pos)
                        OPoints(Pos) = OPoints(Pos + 1)
                        OPoints(Pos + 1) = TemOPoints
                    
                        TemOSkaterName = OSkaterName(Pos)
                        OSkaterName(Pos) = OSkaterName(Pos + 1)
                        OSkaterName(Pos + 1) = TemOSkaterName
                        
                        TemOPosition = OPosition(Pos)
                        OPosition(Pos) = OPosition(Pos + 1)
                        OPosition(Pos + 1) = TemOPosition
                        
                        TemOGP = OGP(Pos)
                        OGP(Pos) = OGP(Pos + 1)
                        OGP(Pos + 1) = TemOGP
                        
                        TemOGoal = OGoal(Pos)
                        OGoal(Pos) = OGoal(Pos + 1)
                        OGoal(Pos + 1) = TemOGoal
                        
                        TemOAssist = OAssist(Pos)
                        OAssist(Pos) = OAssist(Pos + 1)
                        OAssist(Pos + 1) = TemOAssist
                        
                        TemOPtsAGame = OPtsAGame(Pos)
                        OPtsAGame(Pos) = OPtsAGame(Pos + 1)
                        OPtsAGame(Pos + 1) = TemOPtsAGame
                        
                        TemOShots = OShots(Pos)
                        OShots(Pos) = OShots(Pos + 1)
                        OShots(Pos + 1) = TemOShots
                        
                        TemOShotPct = OShotPct(Pos)
                        OShotPct(Pos) = OShotPct(Pos + 1)
                        OShotPct(Pos + 1) = TemOShotPct
                        
                        TemOPlusMinus = OPlusMinus(Pos)
                        OPlusMinus(Pos) = OPlusMinus(Pos + 1)
                        OPlusMinus(Pos + 1) = TemOPlusMinus
                        
                        TemOPIM = OPIM(Pos)
                        OPIM(Pos) = OPIM(Pos + 1)
                        OPIM(Pos + 1) = TemOPIM
                        
                        TemOPPG = OPPG(Pos)
                        OPPG(Pos) = OPPG(Pos + 1)
                        OPPG(Pos + 1) = TemOPPG
                        
                        TemOGWG = OGWG(Pos)
                        OGWG(Pos) = OGWG(Pos + 1)
                        OGWG(Pos + 1) = TemOGWG
                        'picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), OPtsAGame(Pos), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OSHG(Pos), OGWG(Pos)
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "Overall Skater Statstics have been sorted by total points"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            For Pos = 1 To SizeOverall
                picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), FormatNumber(OPtsAGame(Pos), 2), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OGWG(Pos)
            Next Pos
            picLeaderName.Cls
            picLeaderName.Print "#7 Scott Bjorklund, Forward"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\bjorklund3.JPG")
        ElseIf optPtsAGame.Value = True Then
            For Pass = 1 To (SizeOverall - 1)
                For Pos = 1 To (SizeOverall - Pass)
                    If OPtsAGame(Pos) < OPtsAGame(Pos + 1) Then
                        TemOPtsAGame = OPtsAGame(Pos)
                        OPtsAGame(Pos) = OPtsAGame(Pos + 1)
                        OPtsAGame(Pos + 1) = TemOPtsAGame
                    
                        TemOSkaterName = OSkaterName(Pos)
                        OSkaterName(Pos) = OSkaterName(Pos + 1)
                        OSkaterName(Pos + 1) = TemOSkaterName
                        
                        TemOPosition = OPosition(Pos)
                        OPosition(Pos) = OPosition(Pos + 1)
                        OPosition(Pos + 1) = TemOPosition
                        
                        TemOGP = OGP(Pos)
                        OGP(Pos) = OGP(Pos + 1)
                        OGP(Pos + 1) = TemOGP
                        
                        TemOGoal = OGoal(Pos)
                        OGoal(Pos) = OGoal(Pos + 1)
                        OGoal(Pos + 1) = TemOGoal
                        
                        TemOAssist = OAssist(Pos)
                        OAssist(Pos) = OAssist(Pos + 1)
                        OAssist(Pos + 1) = TemOAssist
                        
                        TemOPoints = OPoints(Pos)
                        OPoints(Pos) = OPoints(Pos + 1)
                        OPoints(Pos + 1) = TemOPoints
                        
                        TemOShots = OShots(Pos)
                        OShots(Pos) = OShots(Pos + 1)
                        OShots(Pos + 1) = TemOShots
                        
                        TemOShotPct = OShotPct(Pos)
                        OShotPct(Pos) = OShotPct(Pos + 1)
                        OShotPct(Pos + 1) = TemOShotPct
                        
                        TemOPlusMinus = OPlusMinus(Pos)
                        OPlusMinus(Pos) = OPlusMinus(Pos + 1)
                        OPlusMinus(Pos + 1) = TemOPlusMinus
                        
                        TemOPIM = OPIM(Pos)
                        OPIM(Pos) = OPIM(Pos + 1)
                        OPIM(Pos + 1) = TemOPIM
                        
                        TemOPPG = OPPG(Pos)
                        OPPG(Pos) = OPPG(Pos + 1)
                        OPPG(Pos + 1) = TemOPPG
                        
                        TemOGWG = OGWG(Pos)
                        OGWG(Pos) = OGWG(Pos + 1)
                        OGWG(Pos + 1) = TemOGWG
                        'picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), OPtsAGame(Pos), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OSHG(Pos), OGWG(Pos)
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "Overall Skater Statstics have been sorted by points per game"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            For Pos = 1 To SizeOverall
                picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), FormatNumber(OPtsAGame(Pos), 2), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OGWG(Pos)
            Next Pos
            picLeaderName.Cls
            picLeaderName.Print "#7 Scott Bjorklund, Forward"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\bjorklund3.JPG")
        ElseIf optShots.Value = True Then
            For Pass = 1 To (SizeOverall - 1)
                For Pos = 1 To (SizeOverall - Pass)
                    If OShots(Pos) < OShots(Pos + 1) Then
                        TemOShots = OShots(Pos)
                        OShots(Pos) = OShots(Pos + 1)
                        OShots(Pos + 1) = TemOShots
                    
                        TemOSkaterName = OSkaterName(Pos)
                        OSkaterName(Pos) = OSkaterName(Pos + 1)
                        OSkaterName(Pos + 1) = TemOSkaterName
                        
                        TemOPosition = OPosition(Pos)
                        OPosition(Pos) = OPosition(Pos + 1)
                        OPosition(Pos + 1) = TemOPosition
                        
                        TemOGP = OGP(Pos)
                        OGP(Pos) = OGP(Pos + 1)
                        OGP(Pos + 1) = TemOGP
                        
                        TemOGoal = OGoal(Pos)
                        OGoal(Pos) = OGoal(Pos + 1)
                        OGoal(Pos + 1) = TemOGoal
                        
                        TemOAssist = OAssist(Pos)
                        OAssist(Pos) = OAssist(Pos + 1)
                        OAssist(Pos + 1) = TemOAssist
                        
                        TemOPoints = OPoints(Pos)
                        OPoints(Pos) = OPoints(Pos + 1)
                        OPoints(Pos + 1) = TemOPoints
                        
                        TemOPtsAGame = OPtsAGame(Pos)
                        OPtsAGame(Pos) = OPtsAGame(Pos + 1)
                        OPtsAGame(Pos + 1) = TemOPtsAGame
                        
                        TemOShotPct = OShotPct(Pos)
                        OShotPct(Pos) = OShotPct(Pos + 1)
                        OShotPct(Pos + 1) = TemOShotPct
                        
                        TemOPlusMinus = OPlusMinus(Pos)
                        OPlusMinus(Pos) = OPlusMinus(Pos + 1)
                        OPlusMinus(Pos + 1) = TemOPlusMinus
                        
                        TemOPIM = OPIM(Pos)
                        OPIM(Pos) = OPIM(Pos + 1)
                        OPIM(Pos + 1) = TemOPIM
                        
                        TemOPPG = OPPG(Pos)
                        OPPG(Pos) = OPPG(Pos + 1)
                        OPPG(Pos + 1) = TemOPPG
                        
                        TemOGWG = OGWG(Pos)
                        OGWG(Pos) = OGWG(Pos + 1)
                        OGWG(Pos + 1) = TemOGWG
                        'picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), OPtsAGame(Pos), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OSHG(Pos), OGWG(Pos)
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "Overall Skater Statstics have been sorted by shots taken"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            For Pos = 1 To SizeOverall
                picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), FormatNumber(OPtsAGame(Pos), 2), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OGWG(Pos)
            Next Pos
            picLeaderName.Cls
            picLeaderName.Print "#10 Tom Freeman, Forward"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\Freeman3.JPG")
        ElseIf optShotpct.Value = True Then
            For Pass = 1 To (SizeOverall - 1)
                For Pos = 1 To (SizeOverall - Pass)
                    If OShotPct(Pos) < OShotPct(Pos + 1) Then
                        TemOShotPct = OShotPct(Pos)
                        OShotPct(Pos) = OShotPct(Pos + 1)
                        OShotPct(Pos + 1) = TemOShotPct
                    
                        TemOSkaterName = OSkaterName(Pos)
                        OSkaterName(Pos) = OSkaterName(Pos + 1)
                        OSkaterName(Pos + 1) = TemOSkaterName
                        
                        TemOPosition = OPosition(Pos)
                        OPosition(Pos) = OPosition(Pos + 1)
                        OPosition(Pos + 1) = TemOPosition
                        
                        TemOGP = OGP(Pos)
                        OGP(Pos) = OGP(Pos + 1)
                        OGP(Pos + 1) = TemOGP
                        
                        TemOGoal = OGoal(Pos)
                        OGoal(Pos) = OGoal(Pos + 1)
                        OGoal(Pos + 1) = TemOGoal
                        
                        TemOAssist = OAssist(Pos)
                        OAssist(Pos) = OAssist(Pos + 1)
                        OAssist(Pos + 1) = TemOAssist
                        
                        TemOPoints = OPoints(Pos)
                        OPoints(Pos) = OPoints(Pos + 1)
                        OPoints(Pos + 1) = TemOPoints
                        
                        TemOPtsAGame = OPtsAGame(Pos)
                        OPtsAGame(Pos) = OPtsAGame(Pos + 1)
                        OPtsAGame(Pos + 1) = TemOPtsAGame
                        
                        TemOShots = OShots(Pos)
                        OShots(Pos) = OShots(Pos + 1)
                        OShots(Pos + 1) = TemOShots
                        
                        TemOPlusMinus = OPlusMinus(Pos)
                        OPlusMinus(Pos) = OPlusMinus(Pos + 1)
                        OPlusMinus(Pos + 1) = TemOPlusMinus
                        
                        TemOPIM = OPIM(Pos)
                        OPIM(Pos) = OPIM(Pos + 1)
                        OPIM(Pos + 1) = TemOPIM
                        
                        TemOPPG = OPPG(Pos)
                        OPPG(Pos) = OPPG(Pos + 1)
                        OPPG(Pos + 1) = TemOPPG
                        
                        TemOGWG = OGWG(Pos)
                        OGWG(Pos) = OGWG(Pos + 1)
                        OGWG(Pos + 1) = TemOGWG
                        'picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), OPtsAGame(Pos), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OSHG(Pos), OGWG(Pos)
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "Overall Skater Statstics have been sorted by shooting percentage"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            For Pos = 1 To SizeOverall
                picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), FormatNumber(OPtsAGame(Pos), 2), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OGWG(Pos)
            Next Pos
            picLeaderName.Cls
            picLeaderName.Print "#10 Tom Freeman, Forward"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\Freeman3.JPG")
        ElseIf optPlusMinus.Value = True Then
            For Pass = 1 To (SizeOverall - 1)
                For Pos = 1 To (SizeOverall - Pass)
                    If OPlusMinus(Pos) < OPlusMinus(Pos + 1) Then
                        TemOPlusMinus = OPlusMinus(Pos)
                        OPlusMinus(Pos) = OPlusMinus(Pos + 1)
                        OPlusMinus(Pos + 1) = TemOPlusMinus
                    
                        TemOSkaterName = OSkaterName(Pos)
                        OSkaterName(Pos) = OSkaterName(Pos + 1)
                        OSkaterName(Pos + 1) = TemOSkaterName
                        
                        TemOPosition = OPosition(Pos)
                        OPosition(Pos) = OPosition(Pos + 1)
                        OPosition(Pos + 1) = TemOPosition
                        
                        TemOGP = OGP(Pos)
                        OGP(Pos) = OGP(Pos + 1)
                        OGP(Pos + 1) = TemOGP
                        
                        TemOGoal = OGoal(Pos)
                        OGoal(Pos) = OGoal(Pos + 1)
                        OGoal(Pos + 1) = TemOGoal
                        
                        TemOAssist = OAssist(Pos)
                        OAssist(Pos) = OAssist(Pos + 1)
                        OAssist(Pos + 1) = TemOAssist
                        
                        TemOPoints = OPoints(Pos)
                        OPoints(Pos) = OPoints(Pos + 1)
                        OPoints(Pos + 1) = TemOPoints
                        
                        TemOPtsAGame = OPtsAGame(Pos)
                        OPtsAGame(Pos) = OPtsAGame(Pos + 1)
                        OPtsAGame(Pos + 1) = TemOPtsAGame
                        
                        TemOShots = OShots(Pos)
                        OShots(Pos) = OShots(Pos + 1)
                        OShots(Pos + 1) = TemOShots
                        
                        TemOShotPct = OShotPct(Pos)
                        OShotPct(Pos) = OShotPct(Pos + 1)
                        OShotPct(Pos + 1) = TemOShotPct
                        
                        TemOPIM = OPIM(Pos)
                        OPIM(Pos) = OPIM(Pos + 1)
                        OPIM(Pos + 1) = TemOPIM
                        
                        TemOPPG = OPPG(Pos)
                        OPPG(Pos) = OPPG(Pos + 1)
                        OPPG(Pos + 1) = TemOPPG
                        
                        TemOGWG = OGWG(Pos)
                        OGWG(Pos) = OGWG(Pos + 1)
                        OGWG(Pos + 1) = TemOGWG
                        'picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), OPtsAGame(Pos), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OSHG(Pos), OGWG(Pos)
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "Overall Skater Statstics have been sorted by plus/minus"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            For Pos = 1 To SizeOverall
                picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), FormatNumber(OPtsAGame(Pos), 2), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OGWG(Pos)
            Next Pos
            picLeaderName.Cls
            picLeaderName.Print "#6 Dustin Mercado, Defenseman"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\mercado3.JPG")
        ElseIf optPIM.Value = True Then
            For Pass = 1 To (SizeOverall - 1)
                For Pos = 1 To (SizeOverall - Pass)
                    If OPIM(Pos) < OPIM(Pos + 1) Then
                        TemOPIM = OPIM(Pos)
                        OPIM(Pos) = OPIM(Pos + 1)
                        OPIM(Pos + 1) = TemOPIM
                    
                        TemOSkaterName = OSkaterName(Pos)
                        OSkaterName(Pos) = OSkaterName(Pos + 1)
                        OSkaterName(Pos + 1) = TemOSkaterName
                        
                        TemOPosition = OPosition(Pos)
                        OPosition(Pos) = OPosition(Pos + 1)
                        OPosition(Pos + 1) = TemOPosition
                        
                        TemOGP = OGP(Pos)
                        OGP(Pos) = OGP(Pos + 1)
                        OGP(Pos + 1) = TemOGP
                        
                        TemOGoal = OGoal(Pos)
                        OGoal(Pos) = OGoal(Pos + 1)
                        OGoal(Pos + 1) = TemOGoal
                        
                        TemOAssist = OAssist(Pos)
                        OAssist(Pos) = OAssist(Pos + 1)
                        OAssist(Pos + 1) = TemOAssist
                        
                        TemOPoints = OPoints(Pos)
                        OPoints(Pos) = OPoints(Pos + 1)
                        OPoints(Pos + 1) = TemOPoints
                        
                        TemOPtsAGame = OPtsAGame(Pos)
                        OPtsAGame(Pos) = OPtsAGame(Pos + 1)
                        OPtsAGame(Pos + 1) = TemOPtsAGame
                        
                        TemOShots = OShots(Pos)
                        OShots(Pos) = OShots(Pos + 1)
                        OShots(Pos + 1) = TemOShots
                        
                        TemOShotPct = OShotPct(Pos)
                        OShotPct(Pos) = OShotPct(Pos + 1)
                        OShotPct(Pos + 1) = TemOShotPct
                        
                        TemOPlusMinus = OPlusMinus(Pos)
                        OPlusMinus(Pos) = OPlusMinus(Pos + 1)
                        OPlusMinus(Pos + 1) = TemOPlusMinus
                        
                        TemOPPG = OPPG(Pos)
                        OPPG(Pos) = OPPG(Pos + 1)
                        OPPG(Pos + 1) = TemOPPG
                        
                        TemOGWG = OGWG(Pos)
                        OGWG(Pos) = OGWG(Pos + 1)
                        OGWG(Pos + 1) = TemOGWG
                        'picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), OPtsAGame(Pos), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OSHG(Pos), OGWG(Pos)
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "Overall Skater Statstics have been sorted by most penalty minutes"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            For Pos = 1 To SizeOverall
                picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), FormatNumber(OPtsAGame(Pos), 2), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OGWG(Pos)
            Next Pos
            picLeaderName.Cls
            picLeaderName.Print "#22 Bille Luger, Forward"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\luger3.JPG")
        ElseIf optPPG.Value = True Then
            For Pass = 1 To (SizeOverall - 1)
                For Pos = 1 To (SizeOverall - Pass)
                    If OPPG(Pos) < OPPG(Pos + 1) Then
                        TemOPPG = OPPG(Pos)
                        OPPG(Pos) = OPPG(Pos + 1)
                        OPPG(Pos + 1) = TemOPPG
                    
                        TemOSkaterName = OSkaterName(Pos)
                        OSkaterName(Pos) = OSkaterName(Pos + 1)
                        OSkaterName(Pos + 1) = TemOSkaterName
                        
                        TemOPosition = OPosition(Pos)
                        OPosition(Pos) = OPosition(Pos + 1)
                        OPosition(Pos + 1) = TemOPosition
                        
                        TemOGP = OGP(Pos)
                        OGP(Pos) = OGP(Pos + 1)
                        OGP(Pos + 1) = TemOGP
                        
                        TemOGoal = OGoal(Pos)
                        OGoal(Pos) = OGoal(Pos + 1)
                        OGoal(Pos + 1) = TemOGoal
                        
                        TemOAssist = OAssist(Pos)
                        OAssist(Pos) = OAssist(Pos + 1)
                        OAssist(Pos + 1) = TemOAssist
                        
                        TemOPoints = OPoints(Pos)
                        OPoints(Pos) = OPoints(Pos + 1)
                        OPoints(Pos + 1) = TemOPoints
                        
                        TemOPtsAGame = OPtsAGame(Pos)
                        OPtsAGame(Pos) = OPtsAGame(Pos + 1)
                        OPtsAGame(Pos + 1) = TemOPtsAGame
                        
                        TemOShots = OShots(Pos)
                        OShots(Pos) = OShots(Pos + 1)
                        OShots(Pos + 1) = TemOShots
                        
                        TemOShotPct = OShotPct(Pos)
                        OShotPct(Pos) = OShotPct(Pos + 1)
                        OShotPct(Pos + 1) = TemOShotPct
                        
                        TemOPlusMinus = OPlusMinus(Pos)
                        OPlusMinus(Pos) = OPlusMinus(Pos + 1)
                        OPlusMinus(Pos + 1) = TemOPlusMinus
                        
                        TemOPIM = OPIM(Pos)
                        OPIM(Pos) = OPIM(Pos + 1)
                        OPIM(Pos + 1) = TemOPIM
                        
                        TemOGWG = OGWG(Pos)
                        OGWG(Pos) = OGWG(Pos + 1)
                        OGWG(Pos + 1) = TemOGWG
                        'picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), OPtsAGame(Pos), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OSHG(Pos), OGWG(Pos)
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "Overall Skater Statstics have been sorted by power play goals"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            For Pos = 1 To SizeOverall
                picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), FormatNumber(OPtsAGame(Pos), 2), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OGWG(Pos)
            Next Pos
            picLeaderName.Cls
            picLeaderName.Print "#7 Scott Bjorklund, Forward"
            picLeaderName.Print "#9 Aaron Getchell, Forward"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\bjorklund3.JPG")
        ElseIf optGWG.Value = True Then
            For Pass = 1 To (SizeOverall - 1)
                For Pos = 1 To (SizeOverall - Pass)
                    If OGWG(Pos) < OGWG(Pos + 1) Then
                        TemOGWG = OGWG(Pos)
                        OGWG(Pos) = OGWG(Pos + 1)
                        OGWG(Pos + 1) = TemOGWG
                    
                        TemOSkaterName = OSkaterName(Pos)
                        OSkaterName(Pos) = OSkaterName(Pos + 1)
                        OSkaterName(Pos + 1) = TemOSkaterName
                        
                        TemOPosition = OPosition(Pos)
                        OPosition(Pos) = OPosition(Pos + 1)
                        OPosition(Pos + 1) = TemOPosition
                        
                        TemOGP = OGP(Pos)
                        OGP(Pos) = OGP(Pos + 1)
                        OGP(Pos + 1) = TemOGP
                        
                        TemOGoal = OGoal(Pos)
                        OGoal(Pos) = OGoal(Pos + 1)
                        OGoal(Pos + 1) = TemOGoal
                        
                        TemOAssist = OAssist(Pos)
                        OAssist(Pos) = OAssist(Pos + 1)
                        OAssist(Pos + 1) = TemOAssist
                        
                        TemOPoints = OPoints(Pos)
                        OPoints(Pos) = OPoints(Pos + 1)
                        OPoints(Pos + 1) = TemOPoints
                        
                        TemOPtsAGame = OPtsAGame(Pos)
                        OPtsAGame(Pos) = OPtsAGame(Pos + 1)
                        OPtsAGame(Pos + 1) = TemOPtsAGame
                        
                        TemOShots = OShots(Pos)
                        OShots(Pos) = OShots(Pos + 1)
                        OShots(Pos + 1) = TemOShots
                        
                        TemOShotPct = OShotPct(Pos)
                        OShotPct(Pos) = OShotPct(Pos + 1)
                        OShotPct(Pos + 1) = TemOShotPct
                        
                        TemOPlusMinus = OPlusMinus(Pos)
                        OPlusMinus(Pos) = OPlusMinus(Pos + 1)
                        OPlusMinus(Pos + 1) = TemOPlusMinus
                        
                        TemOPIM = OPIM(Pos)
                        OPIM(Pos) = OPIM(Pos + 1)
                        OPIM(Pos + 1) = TemOPIM
                        
                        TemOPPG = OPPG(Pos)
                        OPPG(Pos) = OPPG(Pos + 1)
                        OPPG(Pos + 1) = TemOPPG
                        'picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), OPtsAGame(Pos), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OSHG(Pos), OGWG(Pos)
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "Overall Skater Statstics have been sorted by game-winning goals"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            For Pos = 1 To SizeOverall
                picStatsorter.Print OSkaterName(Pos), OPosition(Pos), OGP(Pos), OGoal(Pos), OAssist(Pos), OPoints(Pos), FormatNumber(OPtsAGame(Pos), 2), OShots(Pos), OShotPct(Pos), OPlusMinus(Pos), OPIM(Pos), OPPG(Pos), OGWG(Pos)
            Next Pos
            picLeaderName.Cls
            picLeaderName.Print "#10 Tom Freeman, Forward"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\Freeman3.JPG")
        End If
    'If skaters and conference stats are selected, this code determines which way to sort MIAC stats and output them
    ElseIf cmdSwitchConf.Enabled = False And cmdSwitchSkaters.Enabled = False Then
        If optGoals.Value = True Then
            For Pass = 1 To (SizeMIAC - 1)
                For Pos = 1 To (SizeMIAC - Pass)
                    If MGoal(Pos) < MGoal(Pos + 1) Then
                        TemMGoal = MGoal(Pos)
                        MGoal(Pos) = MGoal(Pos + 1)
                        MGoal(Pos + 1) = TemMGoal
                    
                        TemMSkaterName = MSkaterName(Pos)
                        MSkaterName(Pos) = MSkaterName(Pos + 1)
                        MSkaterName(Pos + 1) = TemMSkaterName
                        
                        TemMPosition = MPosition(Pos)
                        MPosition(Pos) = MPosition(Pos + 1)
                        MPosition(Pos + 1) = TemMPosition
                        
                        TemMGP = MGP(Pos)
                        MGP(Pos) = MGP(Pos + 1)
                        MGP(Pos + 1) = TemMGP
                        
                        TemMAssist = MAssist(Pos)
                        MAssist(Pos) = MAssist(Pos + 1)
                        MAssist(Pos + 1) = TemMAssist
                        
                        TemMPoints = MPoints(Pos)
                        MPoints(Pos) = MPoints(Pos + 1)
                        MPoints(Pos + 1) = TemMPoints
                        
                        TemMPtsAGame = MPtsAGame(Pos)
                        MPtsAGame(Pos) = MPtsAGame(Pos + 1)
                        MPtsAGame(Pos + 1) = TemMPtsAGame
                        
                        TemMShots = MShots(Pos)
                        MShots(Pos) = MShots(Pos + 1)
                        MShots(Pos + 1) = TemMShots
                        
                        TemMShotPct = MShotPct(Pos)
                        MShotPct(Pos) = MShotPct(Pos + 1)
                        MShotPct(Pos + 1) = TemMShotPct
                        
                        TemMPlusMinus = MPlusMinus(Pos)
                        MPlusMinus(Pos) = MPlusMinus(Pos + 1)
                        MPlusMinus(Pos + 1) = TemMPlusMinus
                        
                        TemMPIM = MPIM(Pos)
                        MPIM(Pos) = MPIM(Pos + 1)
                        MPIM(Pos + 1) = TemMPIM
                        
                        TemMPPG = MPPG(Pos)
                        MPPG(Pos) = MPPG(Pos + 1)
                        MPPG(Pos + 1) = TemMPPG
                        
                        TemMGWG = MGWG(Pos)
                        MGWG(Pos) = MGWG(Pos + 1)
                        MGWG(Pos + 1) = TemMGWG
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "MIAC Skater Statstics have been sorted by goals scored"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            For Pos = 1 To SizeMIAC
                picStatsorter.Print MSkaterName(Pos), MPosition(Pos), MGP(Pos), MGoal(Pos), MAssist(Pos), MPoints(Pos), FormatNumber(MPtsAGame(Pos), 2), MShots(Pos), MShotPct(Pos), MPlusMinus(Pos), MPIM(Pos), MPPG(Pos), MGWG(Pos)
            Next Pos
            picLeaderName.Cls
            picLeaderName.Print "#10 Tom Freeman, Forward"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\Freeman3.JPG")
        ElseIf optAssists.Value = True Then
            For Pass = 1 To (SizeMIAC - 1)
                For Pos = 1 To (SizeMIAC - Pass)
                    If MAssist(Pos) < MAssist(Pos + 1) Then
                        TemMAssist = MAssist(Pos)
                        MAssist(Pos) = MAssist(Pos + 1)
                        MAssist(Pos + 1) = TemMAssist
                    
                        TemMSkaterName = MSkaterName(Pos)
                        MSkaterName(Pos) = MSkaterName(Pos + 1)
                        MSkaterName(Pos + 1) = TemMSkaterName
                        
                        TemMPosition = MPosition(Pos)
                        MPosition(Pos) = MPosition(Pos + 1)
                        MPosition(Pos + 1) = TemMPosition
                        
                        TemMGP = MGP(Pos)
                        MGP(Pos) = MGP(Pos + 1)
                        MGP(Pos + 1) = TemMGP
                        
                        TemMGoal = MGoal(Pos)
                        MGoal(Pos) = MGoal(Pos + 1)
                        MGoal(Pos + 1) = TemMGoal
                        
                        TemMPoints = MPoints(Pos)
                        MPoints(Pos) = MPoints(Pos + 1)
                        MPoints(Pos + 1) = TemMPoints
                        
                        TemMPtsAGame = MPtsAGame(Pos)
                        MPtsAGame(Pos) = MPtsAGame(Pos + 1)
                        MPtsAGame(Pos + 1) = TemMPtsAGame
                        
                        TemMShots = MShots(Pos)
                        MShots(Pos) = MShots(Pos + 1)
                        MShots(Pos + 1) = TemMShots
                        
                        TemMShotPct = MShotPct(Pos)
                        MShotPct(Pos) = MShotPct(Pos + 1)
                        MShotPct(Pos + 1) = TemMShotPct
                        
                        TemMPlusMinus = MPlusMinus(Pos)
                        MPlusMinus(Pos) = MPlusMinus(Pos + 1)
                        MPlusMinus(Pos + 1) = TemMPlusMinus
                        
                        TemMPIM = MPIM(Pos)
                        MPIM(Pos) = MPIM(Pos + 1)
                        MPIM(Pos + 1) = TemMPIM
                        
                        TemMPPG = MPPG(Pos)
                        MPPG(Pos) = MPPG(Pos + 1)
                        MPPG(Pos + 1) = TemMPPG
                        
                        TemMGWG = MGWG(Pos)
                        MGWG(Pos) = MGWG(Pos + 1)
                        MGWG(Pos + 1) = TemMGWG
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "MIAC Skater Statstics have been sorted by assists"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            For Pos = 1 To SizeMIAC
                picStatsorter.Print MSkaterName(Pos), MPosition(Pos), MGP(Pos), MGoal(Pos), MAssist(Pos), MPoints(Pos), FormatNumber(MPtsAGame(Pos), 2), MShots(Pos), MShotPct(Pos), MPlusMinus(Pos), MPIM(Pos), MPPG(Pos), MGWG(Pos)
            Next Pos
            picLeaderName.Cls
            picLeaderName.Print "#7 Scott Bjorklund, Forward"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\bjorklund3.JPG")
        ElseIf optPoints.Value = True Then
            For Pass = 1 To (SizeMIAC - 1)
                For Pos = 1 To (SizeMIAC - Pass)
                    If MPoints(Pos) < MPoints(Pos + 1) Then
                        TemMPoints = MPoints(Pos)
                        MPoints(Pos) = MPoints(Pos + 1)
                        MPoints(Pos + 1) = TemMPoints
                    
                        TemMSkaterName = MSkaterName(Pos)
                        MSkaterName(Pos) = MSkaterName(Pos + 1)
                        MSkaterName(Pos + 1) = TemMSkaterName
                        
                        TemMPosition = MPosition(Pos)
                        MPosition(Pos) = MPosition(Pos + 1)
                        MPosition(Pos + 1) = TemMPosition
                        
                        TemMGP = MGP(Pos)
                        MGP(Pos) = MGP(Pos + 1)
                        MGP(Pos + 1) = TemMGP
                        
                        TemMAssist = MAssist(Pos)
                        MAssist(Pos) = MAssist(Pos + 1)
                        MAssist(Pos + 1) = TemMAssist
                        
                        TemMGoal = MGoal(Pos)
                        MGoal(Pos) = MGoal(Pos + 1)
                        MGoal(Pos + 1) = TemMGoal
                        
                        TemMPtsAGame = MPtsAGame(Pos)
                        MPtsAGame(Pos) = MPtsAGame(Pos + 1)
                        MPtsAGame(Pos + 1) = TemMPtsAGame
                        
                        TemMShots = MShots(Pos)
                        MShots(Pos) = MShots(Pos + 1)
                        MShots(Pos + 1) = TemMShots
                        
                        TemMShotPct = MShotPct(Pos)
                        MShotPct(Pos) = MShotPct(Pos + 1)
                        MShotPct(Pos + 1) = TemMShotPct
                        
                        TemMPlusMinus = MPlusMinus(Pos)
                        MPlusMinus(Pos) = MPlusMinus(Pos + 1)
                        MPlusMinus(Pos + 1) = TemMPlusMinus
                        
                        TemMPIM = MPIM(Pos)
                        MPIM(Pos) = MPIM(Pos + 1)
                        MPIM(Pos + 1) = TemMPIM
                        
                        TemMPPG = MPPG(Pos)
                        MPPG(Pos) = MPPG(Pos + 1)
                        MPPG(Pos + 1) = TemMPPG
                        
                        TemMGWG = MGWG(Pos)
                        MGWG(Pos) = MGWG(Pos + 1)
                        MGWG(Pos + 1) = TemMGWG
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "MIAC Skater Statstics have been sorted by points"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            For Pos = 1 To SizeMIAC
                picStatsorter.Print MSkaterName(Pos), MPosition(Pos), MGP(Pos), MGoal(Pos), MAssist(Pos), MPoints(Pos), FormatNumber(MPtsAGame(Pos), 2), MShots(Pos), MShotPct(Pos), MPlusMinus(Pos), MPIM(Pos), MPPG(Pos), MGWG(Pos)
            Next Pos
            picLeaderName.Cls
            picLeaderName.Print "#7 Scott Bjorklund, Forward"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\bjorklund3.JPG")
        ElseIf optPtsAGame.Value = True Then
            For Pass = 1 To (SizeMIAC - 1)
                For Pos = 1 To (SizeMIAC - Pass)
                    If MPtsAGame(Pos) < MPtsAGame(Pos + 1) Then
                        TemMPtsAGame = MPtsAGame(Pos)
                        MPtsAGame(Pos) = MPtsAGame(Pos + 1)
                        MPtsAGame(Pos + 1) = TemMPtsAGame
                    
                        TemMSkaterName = MSkaterName(Pos)
                        MSkaterName(Pos) = MSkaterName(Pos + 1)
                        MSkaterName(Pos + 1) = TemMSkaterName
                        
                        TemMPosition = MPosition(Pos)
                        MPosition(Pos) = MPosition(Pos + 1)
                        MPosition(Pos + 1) = TemMPosition
                        
                        TemMGP = MGP(Pos)
                        MGP(Pos) = MGP(Pos + 1)
                        MGP(Pos + 1) = TemMGP
                        
                        TemMAssist = MAssist(Pos)
                        MAssist(Pos) = MAssist(Pos + 1)
                        MAssist(Pos + 1) = TemMAssist
                        
                        TemMGoal = MGoal(Pos)
                        MGoal(Pos) = MGoal(Pos + 1)
                        MGoal(Pos + 1) = TemMGoal
                        
                        TemMPoints = MPoints(Pos)
                        MPoints(Pos) = MPoints(Pos + 1)
                        MPoints(Pos + 1) = TemMPoints
                        
                        TemMShots = MShots(Pos)
                        MShots(Pos) = MShots(Pos + 1)
                        MShots(Pos + 1) = TemMShots
                        
                        TemMShotPct = MShotPct(Pos)
                        MShotPct(Pos) = MShotPct(Pos + 1)
                        MShotPct(Pos + 1) = TemMShotPct
                        
                        TemMPlusMinus = MPlusMinus(Pos)
                        MPlusMinus(Pos) = MPlusMinus(Pos + 1)
                        MPlusMinus(Pos + 1) = TemMPlusMinus
                        
                        TemMPIM = MPIM(Pos)
                        MPIM(Pos) = MPIM(Pos + 1)
                        MPIM(Pos + 1) = TemMPIM
                        
                        TemMPPG = MPPG(Pos)
                        MPPG(Pos) = MPPG(Pos + 1)
                        MPPG(Pos + 1) = TemMPPG
                        
                        TemMGWG = MGWG(Pos)
                        MGWG(Pos) = MGWG(Pos + 1)
                        MGWG(Pos + 1) = TemMGWG
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "MIAC Skater Statstics have been sorted by points a game"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            For Pos = 1 To SizeMIAC
                picStatsorter.Print MSkaterName(Pos), MPosition(Pos), MGP(Pos), MGoal(Pos), MAssist(Pos), MPoints(Pos), FormatNumber(MPtsAGame(Pos), 2), MShots(Pos), MShotPct(Pos), MPlusMinus(Pos), MPIM(Pos), MPPG(Pos), MGWG(Pos)
            Next Pos
            picLeaderName.Cls
            picLeaderName.Print "#Scott Bjorklund, Forward"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\bjorklund3.JPG")
        ElseIf optShots.Value = True Then
            For Pass = 1 To (SizeMIAC - 1)
                For Pos = 1 To (SizeMIAC - Pass)
                    If MShots(Pos) < MShots(Pos + 1) Then
                        TemMShots = MShots(Pos)
                        MShots(Pos) = MShots(Pos + 1)
                        MShots(Pos + 1) = TemMShots
                    
                        TemMSkaterName = MSkaterName(Pos)
                        MSkaterName(Pos) = MSkaterName(Pos + 1)
                        MSkaterName(Pos + 1) = TemMSkaterName
                        
                        TemMPosition = MPosition(Pos)
                        MPosition(Pos) = MPosition(Pos + 1)
                        MPosition(Pos + 1) = TemMPosition
                        
                        TemMGP = MGP(Pos)
                        MGP(Pos) = MGP(Pos + 1)
                        MGP(Pos + 1) = TemMGP
                        
                        TemMAssist = MAssist(Pos)
                        MAssist(Pos) = MAssist(Pos + 1)
                        MAssist(Pos + 1) = TemMAssist
                        
                        TemMGoal = MGoal(Pos)
                        MGoal(Pos) = MGoal(Pos + 1)
                        MGoal(Pos + 1) = TemMGoal
                        
                        TemMPoints = MPoints(Pos)
                        MPoints(Pos) = MPoints(Pos + 1)
                        MPoints(Pos + 1) = TemMPoints
                        
                        TemMPtsAGame = MPtsAGame(Pos)
                        MPtsAGame(Pos) = MPtsAGame(Pos + 1)
                        MPtsAGame(Pos + 1) = TemMPtsAGame
                        
                        TemMShotPct = MShotPct(Pos)
                        MShotPct(Pos) = MShotPct(Pos + 1)
                        MShotPct(Pos + 1) = TemMShotPct
                        
                        TemMPlusMinus = MPlusMinus(Pos)
                        MPlusMinus(Pos) = MPlusMinus(Pos + 1)
                        MPlusMinus(Pos + 1) = TemMPlusMinus
                        
                        TemMPIM = MPIM(Pos)
                        MPIM(Pos) = MPIM(Pos + 1)
                        MPIM(Pos + 1) = TemMPIM
                        
                        TemMPPG = MPPG(Pos)
                        MPPG(Pos) = MPPG(Pos + 1)
                        MPPG(Pos + 1) = TemMPPG
                        
                        TemMGWG = MGWG(Pos)
                        MGWG(Pos) = MGWG(Pos + 1)
                        MGWG(Pos + 1) = TemMGWG
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "MIAC Skater Statstics have been sorted by shots taken"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            For Pos = 1 To SizeMIAC
                picStatsorter.Print MSkaterName(Pos), MPosition(Pos), MGP(Pos), MGoal(Pos), MAssist(Pos), MPoints(Pos), FormatNumber(MPtsAGame(Pos), 2), MShots(Pos), MShotPct(Pos), MPlusMinus(Pos), MPIM(Pos), MPPG(Pos), MGWG(Pos)
            Next Pos
            picLeaderName.Cls
            picLeaderName.Print "#10 Tom Freeman, Forward"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\Freeman3.JPG")
        ElseIf optShotpct.Value = True Then
            For Pass = 1 To (SizeMIAC - 1)
                For Pos = 1 To (SizeMIAC - Pass)
                    If MShotPct(Pos) < MShotPct(Pos + 1) Then
                        TemMShotPct = MShotPct(Pos)
                        MShotPct(Pos) = MShotPct(Pos + 1)
                        MShotPct(Pos + 1) = TemMShotPct
                    
                        TemMSkaterName = MSkaterName(Pos)
                        MSkaterName(Pos) = MSkaterName(Pos + 1)
                        MSkaterName(Pos + 1) = TemMSkaterName
                        
                        TemMPosition = MPosition(Pos)
                        MPosition(Pos) = MPosition(Pos + 1)
                        MPosition(Pos + 1) = TemMPosition
                        
                        TemMGP = MGP(Pos)
                        MGP(Pos) = MGP(Pos + 1)
                        MGP(Pos + 1) = TemMGP
                        
                        TemMAssist = MAssist(Pos)
                        MAssist(Pos) = MAssist(Pos + 1)
                        MAssist(Pos + 1) = TemMAssist
                        
                        TemMGoal = MGoal(Pos)
                        MGoal(Pos) = MGoal(Pos + 1)
                        MGoal(Pos + 1) = TemMGoal
                        
                        TemMPoints = MPoints(Pos)
                        MPoints(Pos) = MPoints(Pos + 1)
                        MPoints(Pos + 1) = TemMPoints
                        
                        TemMPtsAGame = MPtsAGame(Pos)
                        MPtsAGame(Pos) = MPtsAGame(Pos + 1)
                        MPtsAGame(Pos + 1) = TemMPtsAGame
                        
                        TemMShots = MShots(Pos)
                        MShots(Pos) = MShots(Pos + 1)
                        MShots(Pos + 1) = TemMShots
                        
                        TemMPlusMinus = MPlusMinus(Pos)
                        MPlusMinus(Pos) = MPlusMinus(Pos + 1)
                        MPlusMinus(Pos + 1) = TemMPlusMinus
                        
                        TemMPIM = MPIM(Pos)
                        MPIM(Pos) = MPIM(Pos + 1)
                        MPIM(Pos + 1) = TemMPIM
                        
                        TemMPPG = MPPG(Pos)
                        MPPG(Pos) = MPPG(Pos + 1)
                        MPPG(Pos + 1) = TemMPPG
                        
                        TemMGWG = MGWG(Pos)
                        MGWG(Pos) = MGWG(Pos + 1)
                        MGWG(Pos + 1) = TemMGWG
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "MIAC Skater Statstics have been sorted by shooting percentage"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            For Pos = 1 To SizeMIAC
                picStatsorter.Print MSkaterName(Pos), MPosition(Pos), MGP(Pos), MGoal(Pos), MAssist(Pos), MPoints(Pos), FormatNumber(MPtsAGame(Pos), 2), MShots(Pos), MShotPct(Pos), MPlusMinus(Pos), MPIM(Pos), MPPG(Pos), MGWG(Pos)
            Next Pos
            picLeaderName.Cls
            picLeaderName.Print "#8 Pat Eagles, Forward"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\eagles3.JPG")
        ElseIf optPlusMinus.Value = True Then
            For Pass = 1 To (SizeMIAC - 1)
                For Pos = 1 To (SizeMIAC - Pass)
                    If MPlusMinus(Pos) < MPlusMinus(Pos + 1) Then
                        TemMPlusMinus = MPlusMinus(Pos)
                        MPlusMinus(Pos) = MPlusMinus(Pos + 1)
                        MPlusMinus(Pos + 1) = TemMPlusMinus
                    
                        TemMSkaterName = MSkaterName(Pos)
                        MSkaterName(Pos) = MSkaterName(Pos + 1)
                        MSkaterName(Pos + 1) = TemMSkaterName
                        
                        TemMPosition = MPosition(Pos)
                        MPosition(Pos) = MPosition(Pos + 1)
                        MPosition(Pos + 1) = TemMPosition
                        
                        TemMGP = MGP(Pos)
                        MGP(Pos) = MGP(Pos + 1)
                        MGP(Pos + 1) = TemMGP
                        
                        TemMAssist = MAssist(Pos)
                        MAssist(Pos) = MAssist(Pos + 1)
                        MAssist(Pos + 1) = TemMAssist
                        
                        TemMGoal = MGoal(Pos)
                        MGoal(Pos) = MGoal(Pos + 1)
                        MGoal(Pos + 1) = TemMGoal
                        
                        TemMPoints = MPoints(Pos)
                        MPoints(Pos) = MPoints(Pos + 1)
                        MPoints(Pos + 1) = TemMPoints
                        
                        TemMPtsAGame = MPtsAGame(Pos)
                        MPtsAGame(Pos) = MPtsAGame(Pos + 1)
                        MPtsAGame(Pos + 1) = TemMPtsAGame
                        
                        TemMShots = MShots(Pos)
                        MShots(Pos) = MShots(Pos + 1)
                        MShots(Pos + 1) = TemMShots
                        
                        TemMShotPct = MShotPct(Pos)
                        MShotPct(Pos) = MShotPct(Pos + 1)
                        MShotPct(Pos + 1) = TemMShotPct
                        
                        TemMPIM = MPIM(Pos)
                        MPIM(Pos) = MPIM(Pos + 1)
                        MPIM(Pos + 1) = TemMPIM
                        
                        TemMPPG = MPPG(Pos)
                        MPPG(Pos) = MPPG(Pos + 1)
                        MPPG(Pos + 1) = TemMPPG
                        
                        TemMGWG = MGWG(Pos)
                        MGWG(Pos) = MGWG(Pos + 1)
                        MGWG(Pos + 1) = TemMGWG
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "MIAC Skater Statstics have been sorted by plus/minus"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            For Pos = 1 To SizeMIAC
                picStatsorter.Print MSkaterName(Pos), MPosition(Pos), MGP(Pos), MGoal(Pos), MAssist(Pos), MPoints(Pos), FormatNumber(MPtsAGame(Pos), 2), MShots(Pos), MShotPct(Pos), MPlusMinus(Pos), MPIM(Pos), MPPG(Pos), MGWG(Pos)
            Next Pos
            picLeaderName.Cls
            picLeaderName.Print "#6 Dustin Mercado, Defenseman"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\Mercado3.JPG")
        ElseIf optPIM.Value = True Then
            For Pass = 1 To (SizeMIAC - 1)
                For Pos = 1 To (SizeMIAC - Pass)
                    If MPIM(Pos) < MPIM(Pos + 1) Then
                        TemMPIM = MPIM(Pos)
                        MPIM(Pos) = MPIM(Pos + 1)
                        MPIM(Pos + 1) = TemMPIM
                    
                        TemMSkaterName = MSkaterName(Pos)
                        MSkaterName(Pos) = MSkaterName(Pos + 1)
                        MSkaterName(Pos + 1) = TemMSkaterName
                        
                        TemMPosition = MPosition(Pos)
                        MPosition(Pos) = MPosition(Pos + 1)
                        MPosition(Pos + 1) = TemMPosition
                        
                        TemMGP = MGP(Pos)
                        MGP(Pos) = MGP(Pos + 1)
                        MGP(Pos + 1) = TemMGP
                        
                        TemMAssist = MAssist(Pos)
                        MAssist(Pos) = MAssist(Pos + 1)
                        MAssist(Pos + 1) = TemMAssist
                        
                        TemMGoal = MGoal(Pos)
                        MGoal(Pos) = MGoal(Pos + 1)
                        MGoal(Pos + 1) = TemMGoal
                        
                        TemMPoints = MPoints(Pos)
                        MPoints(Pos) = MPoints(Pos + 1)
                        MPoints(Pos + 1) = TemMPoints
                        
                        TemMPtsAGame = MPtsAGame(Pos)
                        MPtsAGame(Pos) = MPtsAGame(Pos + 1)
                        MPtsAGame(Pos + 1) = TemMPtsAGame
                        
                        TemMShots = MShots(Pos)
                        MShots(Pos) = MShots(Pos + 1)
                        MShots(Pos + 1) = TemMShots
                        
                        TemMShotPct = MShotPct(Pos)
                        MShotPct(Pos) = MShotPct(Pos + 1)
                        MShotPct(Pos + 1) = TemMShotPct
                        
                        TemMPlusMinus = MPlusMinus(Pos)
                        MPlusMinus(Pos) = MPlusMinus(Pos + 1)
                        MPlusMinus(Pos + 1) = TemMPlusMinus
                        
                        TemMPPG = MPPG(Pos)
                        MPPG(Pos) = MPPG(Pos + 1)
                        MPPG(Pos + 1) = TemMPPG
                        
                        TemMGWG = MGWG(Pos)
                        MGWG(Pos) = MGWG(Pos + 1)
                        MGWG(Pos + 1) = TemMGWG
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "MIAC Skater Statstics have been sorted by penalty minues"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            For Pos = 1 To SizeMIAC
                picStatsorter.Print MSkaterName(Pos), MPosition(Pos), MGP(Pos), MGoal(Pos), MAssist(Pos), MPoints(Pos), FormatNumber(MPtsAGame(Pos), 2), MShots(Pos), MShotPct(Pos), MPlusMinus(Pos), MPIM(Pos), MPPG(Pos), MGWG(Pos)
            Next Pos
            picLeaderName.Cls
            picLeaderName.Print "#22 Bille Luger, Forward"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\Luger3.JPG")
        ElseIf optPPG.Value = True Then
            For Pass = 1 To (SizeMIAC - 1)
                For Pos = 1 To (SizeMIAC - Pass)
                    If MPPG(Pos) < MPPG(Pos + 1) Then
                        TemMPPG = MPPG(Pos)
                        MPPG(Pos) = MPPG(Pos + 1)
                        MPPG(Pos + 1) = TemMPPG
                    
                        TemMSkaterName = MSkaterName(Pos)
                        MSkaterName(Pos) = MSkaterName(Pos + 1)
                        MSkaterName(Pos + 1) = TemMSkaterName
                        
                        TemMPosition = MPosition(Pos)
                        MPosition(Pos) = MPosition(Pos + 1)
                        MPosition(Pos + 1) = TemMPosition
                        
                        TemMGP = MGP(Pos)
                        MGP(Pos) = MGP(Pos + 1)
                        MGP(Pos + 1) = TemMGP
                        
                        TemMAssist = MAssist(Pos)
                        MAssist(Pos) = MAssist(Pos + 1)
                        MAssist(Pos + 1) = TemMAssist
                        
                        TemMGoal = MGoal(Pos)
                        MGoal(Pos) = MGoal(Pos + 1)
                        MGoal(Pos + 1) = TemMGoal
                        
                        TemMPoints = MPoints(Pos)
                        MPoints(Pos) = MPoints(Pos + 1)
                        MPoints(Pos + 1) = TemMPoints
                        
                        TemMPtsAGame = MPtsAGame(Pos)
                        MPtsAGame(Pos) = MPtsAGame(Pos + 1)
                        MPtsAGame(Pos + 1) = TemMPtsAGame
                        
                        TemMShots = MShots(Pos)
                        MShots(Pos) = MShots(Pos + 1)
                        MShots(Pos + 1) = TemMShots
                        
                        TemMShotPct = MShotPct(Pos)
                        MShotPct(Pos) = MShotPct(Pos + 1)
                        MShotPct(Pos + 1) = TemMShotPct
                        
                        TemMPlusMinus = MPlusMinus(Pos)
                        MPlusMinus(Pos) = MPlusMinus(Pos + 1)
                        MPlusMinus(Pos + 1) = TemMPlusMinus
                        
                        TemMPIM = MPIM(Pos)
                        MPIM(Pos) = MPIM(Pos + 1)
                        MPIM(Pos + 1) = TemMPIM
                        
                        TemMGWG = MGWG(Pos)
                        MGWG(Pos) = MGWG(Pos + 1)
                        MGWG(Pos + 1) = TemMGWG
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "MIAC Skater Statstics have been sorted by power play goals"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            For Pos = 1 To SizeMIAC
                picStatsorter.Print MSkaterName(Pos), MPosition(Pos), MGP(Pos), MGoal(Pos), MAssist(Pos), MPoints(Pos), FormatNumber(MPtsAGame(Pos), 2), MShots(Pos), MShotPct(Pos), MPlusMinus(Pos), MPIM(Pos), MPPG(Pos), MGWG(Pos)
            Next Pos
            picLeaderName.Cls
            picLeaderName.Print "#7 Scott Bjorklund, Forward"
            picLeaderName.Print "#9 Darryl Smoleroff, Defenseman"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\Bjorklund3.JPG")
        ElseIf optGWG.Value = True Then
            For Pass = 1 To (SizeMIAC - 1)
                For Pos = 1 To (SizeMIAC - Pass)
                    If MGWG(Pos) < MGWG(Pos + 1) Then
                        TemMGWG = MGWG(Pos)
                        MGWG(Pos) = MGWG(Pos + 1)
                        MGWG(Pos + 1) = TemMGWG
                        TemMPPG = MPPG(Pos)
                        MPPG(Pos) = MPPG(Pos + 1)
                        MPPG(Pos + 1) = TemMPPG
                    
                        TemMSkaterName = MSkaterName(Pos)
                        MSkaterName(Pos) = MSkaterName(Pos + 1)
                        MSkaterName(Pos + 1) = TemMSkaterName
                        
                        TemMPosition = MPosition(Pos)
                        MPosition(Pos) = MPosition(Pos + 1)
                        MPosition(Pos + 1) = TemMPosition
                        
                        TemMGP = MGP(Pos)
                        MGP(Pos) = MGP(Pos + 1)
                        MGP(Pos + 1) = TemMGP
                        
                        TemMAssist = MAssist(Pos)
                        MAssist(Pos) = MAssist(Pos + 1)
                        MAssist(Pos + 1) = TemMAssist
                        
                        TemMGoal = MGoal(Pos)
                        MGoal(Pos) = MGoal(Pos + 1)
                        MGoal(Pos + 1) = TemMGoal
                        
                        TemMPoints = MPoints(Pos)
                        MPoints(Pos) = MPoints(Pos + 1)
                        MPoints(Pos + 1) = TemMPoints
                        
                        TemMPtsAGame = MPtsAGame(Pos)
                        MPtsAGame(Pos) = MPtsAGame(Pos + 1)
                        MPtsAGame(Pos + 1) = TemMPtsAGame
                        
                        TemMShots = MShots(Pos)
                        MShots(Pos) = MShots(Pos + 1)
                        MShots(Pos + 1) = TemMShots
                        
                        TemMShotPct = MShotPct(Pos)
                        MShotPct(Pos) = MShotPct(Pos + 1)
                        MShotPct(Pos + 1) = TemMShotPct
                        
                        TemMPlusMinus = MPlusMinus(Pos)
                        MPlusMinus(Pos) = MPlusMinus(Pos + 1)
                        MPlusMinus(Pos + 1) = TemMPlusMinus
                        
                        TemMPIM = MPIM(Pos)
                        MPIM(Pos) = MPIM(Pos + 1)
                        MPIM(Pos + 1) = TemMPIM
                        
                        TemMPPG = MPPG(Pos)
                        MPPG(Pos) = MPPG(Pos + 1)
                        MPPG(Pos + 1) = TemMPPG
                    End If
                Next Pos
            Next Pass
            picStatsorter.Print "MIAC Skater Statstics have been sorted by game-winning goals"
            picStatsorter.Print "# and Name", "Pos", "GP", "G", "A", "Pts", "Pts/game", "Shots", "Shot %", "+/-", "PIM", "PPG", "GWG"
            picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
            For Pos = 1 To SizeMIAC
                picStatsorter.Print MSkaterName(Pos), MPosition(Pos), MGP(Pos), MGoal(Pos), MAssist(Pos), MPoints(Pos), FormatNumber(MPtsAGame(Pos), 2), MShots(Pos), MShotPct(Pos), MPlusMinus(Pos), MPIM(Pos), MPPG(Pos), MGWG(Pos)
            Next Pos
            picLeaderName.Cls
            picLeaderName.Print "#10 Tom Freeman, Forward"
            Set picTeamLeader.Picture = LoadPicture(App.Path & "\Freeman3.JPG")
        End If
    Else
        MsgBox "Plese select a stat to sort, select skaters and/or goalies and select MIAC or overall statistics", , "Error"
    End If

                 
End Sub

Private Sub cmdSortGoalies_Click()
    'declare variables
    Dim GamePlay, Shutout, ShotAgainst, Saves, Wins, Losses, Mins, Ties, GoalAgainst As Integer
    Dim GAA, Savepct, ShotPct, PtsAGame As Single
    
    'clear stat sorting picture box
    picStatsorter.Cls
    
    'clear team leader picture box
    picTeamLeader.Picture = LoadPicture("")
    
    
    'If goalies are selected, then produce certain output depending on the selected statistical set
    If cmdSwitchOverall.Enabled = False And cmdSwitchConf.Enabled = True And cmdSwitchGoalies.Enabled = False Then
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
        picStatsorter.Cls
        picStatsorter.Print "#30 Adam Hanna, Senior Goaltender"
        picStatsorter.Print "GP", "W", "L", "T", "Minutes", "SA", "Saves", "GA", "Save %", "GAA", "SO"
        picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatsorter.Print GamePlay, Wins, Losses, Ties, Mins, ShotAgainst, Saves, GoalAgainst, FormatPercent(Savepct, 1), FormatNumber(GAA, 2), Shutout
        picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatsorter.Print "Adam Hanna was the only goaltender to play for St. John's in 2005-2006"
        picLeaderName.Cls
        picLeaderName.Print "#30 Adam Hanna, Goalie"
        Set picTeamLeader.Picture = LoadPicture(App.Path & "\Hanna3.JPG")
    ElseIf cmdSwitchConf.Enabled = False And cmdSwitchOverall.Enabled = True And cmdSwitchGoalies.Enabled = False Then
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
        picStatsorter.Cls
        picStatsorter.Print "#30 Adam Hanna, Senior Goaltender"
        picStatsorter.Print "GP", "W", "L", "T", "Minutes", "SA", "Saves", "GA", "Save %", "GAA", "SO"
        picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatsorter.Print GamePlay, Wins, Losses, Ties, Mins, ShotAgainst, Saves, GoalAgainst, FormatPercent(Savepct, 1), FormatNumber(GAA, 2), Shutout
        picStatsorter.Print "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        picStatsorter.Print "Adam Hanna was the only goaltender to play for St. John's in 2005-2006"
        picLeaderName.Cls
        picLeaderName.Print "#30 Adam Hanna, Goalie"
        Set picTeamLeader.Picture = LoadPicture(App.Path & "\Hanna3.JPG")
    'display error if user does not select either goalies or skaters
    ElseIf cmdSwitchGoalies.Enabled = True And cmdSwitchSkaters.Enabled = True Then
        MsgBox "Please select either 'Goalies' or 'Skaters'", , "Selection Error"
    ElseIf cmdSwitchGoalies.Enabled = True And cmdSwitchSkaters.Enabled = False Then
        MsgBox "Please select 'Goalies'", , "Selection Error"
    Else
        'produce error message that says user needs to select a statistical set
        MsgBox "Please select either overall or MIAC conference statistics", , "Select Stats"
    End If
    
End Sub
Private Sub cmdSwitchGoalies_Click()
    'changes button settings so user knows which team section he/she is working with
    cmdSwitchGoalies.Enabled = False
    cmdSwitchSkaters.Enabled = True
    
    'resets sorting form when user switches player type
    cmdSwitchConf.Enabled = True
    cmdSwitchOverall.Enabled = True
    picSwitchStats.Cls
    
    'clear stat picture box
    picStatsorter.Cls
    'alter picture box height to make goalie stat proportions look better
    picStatsorter.Height = 2055

    'clear option buttons
    optGoals.Value = False
    optAssists.Value = False
    optGWG.Value = False
    optPPG.Value = False
    optPIM.Value = False
    optPlusMinus.Value = False
    optPoints.Value = False
    optShots.Value = False
    optShotpct.Value = False
    optPtsAGame.Value = False
    
    'clears team leader picture box
    picPlayerSwitch.Cls
    picTeamLeader.Picture = LoadPicture("")
    
    'print in picturebox what section of team the user is working with
    picPlayerSwitch.Print "Goalies"
End Sub

Private Sub cmdSwitchSkaters_Click()
    'changes button settings so user knows which team section he/she is working with
    cmdSwitchGoalies.Enabled = True
    cmdSwitchSkaters.Enabled = False
    
    'resets sorting form when user switches player type
    cmdSwitchConf.Enabled = True
    cmdSwitchOverall.Enabled = True
    picSwitchStats.Cls
    
    'clear stat picture box
    picStatsorter.Cls
    'return sorting picture box to original height
    picStatsorter.Height = 5415
    
    'clears team leader picture box
    picPlayerSwitch.Cls
    picTeamLeader.Picture = LoadPicture("")
    
    'print in picturebox what section of team the user is working with
    picPlayerSwitch.Print "Skaters"

End Sub
Private Sub cmdSwitchOverall_Click()
    'changes button settings so user knows which stat set he/she is working with
    cmdSwitchConf.Enabled = True
    cmdSwitchOverall.Enabled = False
    
    'clear team leader picturebox and which stat the user has chosen
    picSwitchStats.Cls
    picTeamLeader.Picture = LoadPicture("")
    
    'output new user stat selection
    StatSwitch = "Overall Statistics"
    
    'print in picturebox what stat set the user is working with
    picSwitchStats.Print StatSwitch
End Sub

Private Sub cmdSwitchConf_Click()
    'changes button settings so user knows which stat set he/she is working with
    cmdSwitchConf.Enabled = False
    cmdSwitchOverall.Enabled = True
    
    'clear team leader picturebox and which stat the user has chosen
    picSwitchStats.Cls
    picTeamLeader.Picture = LoadPicture("")
    
    'output new user stat selection
    StatSwitch = "MIAC Conference Statistics ONLY"
    
    'print in picturebox what stat set the user is working with
    picSwitchStats.Print StatSwitch
End Sub

Private Sub cmdStatfinder_Click()
    'switches between forms of program
    frmStatsorter.Hide
    frmStatfinder.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub


