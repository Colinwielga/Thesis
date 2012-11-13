VERSION 5.00
Begin VB.Form frmPlayerMatch 
   BackColor       =   &H00808000&
   Caption         =   "Player Matchup"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox fileTeam2 
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   930
      Left            =   7080
      Pattern         =   "*Team.txt*"
      TabIndex        =   24
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton cmdDisplayTeam 
      Caption         =   "Display Team Averages"
      Height          =   495
      Left            =   4680
      TabIndex        =   23
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmdSaveOTeamAvg 
      Caption         =   "Save Team #2 Averages"
      Height          =   495
      Left            =   8040
      TabIndex        =   21
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmdSaveTeamAvg 
      Caption         =   "Save Team #1 Averages"
      Height          =   495
      Left            =   1080
      TabIndex        =   20
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmdSortPoints 
      Caption         =   "Points"
      Height          =   495
      Left            =   4680
      TabIndex        =   17
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdSortMinutes 
      Caption         =   "Minutes"
      Height          =   495
      Left            =   4680
      TabIndex        =   16
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton cmdSortRebounds 
      Caption         =   "Rebounds"
      Height          =   495
      Left            =   4680
      TabIndex        =   15
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton cmdSortTurnovers 
      Caption         =   "Turnovers"
      Height          =   495
      Left            =   4680
      TabIndex        =   14
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton cmdOpen2 
      Caption         =   "Open Selected Team #2"
      Height          =   615
      Left            =   9840
      TabIndex        =   11
      Top             =   2280
      Width           =   1455
   End
   Begin VB.FileListBox fileTeam 
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   930
      Left            =   0
      Pattern         =   "*Team.txt*"
      TabIndex        =   10
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton cmdOpen1 
      Caption         =   "Open Selected Team #1"
      Height          =   615
      Left            =   2760
      TabIndex        =   9
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdNavigateIndivdualAvg 
      Caption         =   "Individual Menu"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   7680
      Width           =   1455
   End
   Begin VB.PictureBox picOtherPlayers 
      Height          =   2655
      Left            =   7080
      ScaleHeight     =   2595
      ScaleWidth      =   4155
      TabIndex        =   4
      Top             =   3480
      Width           =   4215
   End
   Begin VB.PictureBox picHomePlayers 
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2595
      ScaleWidth      =   4035
      TabIndex        =   3
      Top             =   3480
      Width           =   4095
   End
   Begin VB.CommandButton cmdNavigateTeamMatch 
      Caption         =   "Team Menu"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdNavigateMainMenu 
      Caption         =   "Main Menu"
      Height          =   495
      Left            =   8280
      TabIndex        =   1
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   9600
      TabIndex        =   0
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label lblDirection2 
      BackColor       =   &H00808000&
      Caption         =   "**Be Sure Appropriate Team is                  Highlighted**"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   7800
      TabIndex        =   26
      Top             =   6240
      Width           =   2775
   End
   Begin VB.Label lblInstruction 
      BackColor       =   &H00808000&
      Caption         =   "*Must Highlight a Team File*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6840
      TabIndex        =   25
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label lblDirection 
      BackColor       =   &H00808000&
      Caption         =   "**Be Sure Appropriate Team is                  Highlighted**"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   840
      TabIndex        =   22
      Top             =   6240
      Width           =   2775
   End
   Begin VB.Label lblTeamFile 
      BackColor       =   &H00808000&
      Caption         =   "*Must Highlight a Team File*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label lblDesigner 
      BackColor       =   &H00808000&
      Caption         =   "By: Erik Gamradt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Label lblSorting 
      BackColor       =   &H00808000&
      Caption         =   "Rank by Leading Average in :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Team Manager Pro"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label lblOtherPlayers 
      BackColor       =   &H00808000&
      Caption         =   "Team #2 Player Statistics:  (average per game)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7080
      TabIndex        =   7
      Top             =   3240
      Width           =   4095
   End
   Begin VB.Label lblHomePlayers 
      BackColor       =   &H00808000&
      Caption         =   "Team #1 Player Statistics: (average per game)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   3975
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPlayerMatchup 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player Matchup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   555
      Left            =   4320
      TabIndex        =   5
      Top             =   960
      Width           =   2775
   End
End
Attribute VB_Name = "frmPlayerMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Team Manager Pro (ErikGamradtVBProject.vbp)
'frmMainMenu (frmMainMenu.frm)
'Designed By: Erik Gamradt
'22 March 2006
'Users are able to select a Team File, display the full team roster including the saved averaged statistics, sort by specific statistic greatest to least, find and display the team average as a whole, and save this team average.
Option Explicit
Dim OtherNames(1 To 100) As String
Dim OAvgPoints(1 To 100), OAvgMinutes(1 To 100), OAvgRebounds(1 To 100), OAvgTurnovers(1 To 100) As Single
Dim HomeNames(1 To 100) As String
Dim AvgPoints(1 To 100), AvgMinutes(1 To 100), AvgRebounds(1 To 100), AvgTurnovers(1 To 100) As Single
Dim TeamAvgPoints, TeamAvgMinutes, TeamAvgRebounds, TeamAvgTurnovers As Single
Dim OTeamAvgPoints, OTeamAvgMinutes, OTeamAvgRebounds, OTeamAvgTurnovers As Single
Dim Size As Integer

Private Sub cmdDisplayTeam_Click()
    Dim Pos As Integer
    Dim Sum As Single
    Pos = 0
    Sum = 0
    For Pos = 1 To Size                  'Finds whole team averages with known player averages
        Sum = Sum + AvgPoints(Pos)
    Next Pos
    TeamAvgPoints = Sum / Size
    Pos = 0
    Sum = 0
    For Pos = 1 To Size
        Sum = Sum + AvgMinutes(Pos)
    Next Pos
    TeamAvgMinutes = Sum / Size
    Pos = 0
    Sum = 0
    For Pos = 1 To Size
        Sum = Sum + AvgRebounds(Pos)
    Next Pos
    TeamAvgRebounds = Sum / Size
    Pos = 0
    Sum = 0
    For Pos = 1 To Size
        Sum = Sum + AvgTurnovers(Pos)
    Next Pos
    TeamAvgTurnovers = Sum / Size
    picHomePlayers.Print "-------------------------------------------------------------------------------------------"
    picHomePlayers.Print "Team Avg:"; Tab(20); FormatNumber(TeamAvgMinutes); Tab(30); FormatNumber(TeamAvgPoints); Tab(40); FormatNumber(TeamAvgRebounds); Tab(50); FormatNumber(TeamAvgTurnovers)
    'this is the start of the opposing sides team averages
    Pos = 0
    Sum = 0
    For Pos = 1 To Size
        Sum = Sum + OAvgPoints(Pos)
    Next Pos
    OTeamAvgPoints = Sum / Size
    Pos = 0
    Sum = 0
    For Pos = 1 To Size
        Sum = Sum + OAvgMinutes(Pos)
    Next Pos
    OTeamAvgMinutes = Sum / Size
    Pos = 0
    Sum = 0
    For Pos = 1 To Size
        Sum = Sum + OAvgRebounds(Pos)
    Next Pos
    OTeamAvgRebounds = Sum / Size
    Pos = 0
    Sum = 0
    For Pos = 1 To Size
        Sum = Sum + OAvgTurnovers(Pos)
    Next Pos
    OTeamAvgTurnovers = Sum / Size
    picOtherPlayers.Print "-------------------------------------------------------------------------------------------"
    picOtherPlayers.Print "Team Avg:"; Tab(20); FormatNumber(OTeamAvgMinutes); Tab(30); FormatNumber(OTeamAvgPoints); Tab(40); FormatNumber(OTeamAvgRebounds); Tab(50); FormatNumber(OTeamAvgTurnovers)
End Sub

Private Sub cmdNavigateIndivdualAvg_Click()
    frmPlayerMatch.Hide
    frmIndividualAvg.Show
End Sub

Private Sub cmdNavigateMainMenu_Click()
    frmPlayerMatch.Hide
    frmMainMenu.Show
End Sub

Private Sub cmdNavigateTeamMatch_Click()
    frmPlayerMatch.Hide
    frmTeamMatch.Show
End Sub

Private Sub cmdOpen1_Click()
    Dim Pos As Integer
    picHomePlayers.Cls
    picHomePlayers.Print "Player"; Tab(20); "Min"; Tab(30); "Pts"; Tab(40); "Reb"; Tab(50); "TO"
    picHomePlayers.Print "-------------------------------------------------------------------------------------------"
    Pos = 0
    Open fileTeam.Path & "\" & fileTeam.FileName For Input As #1 'opens selected file from file list
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, HomeNames(Pos), AvgMinutes(Pos), AvgPoints(Pos), AvgRebounds(Pos), AvgTurnovers(Pos)
        picHomePlayers.Print HomeNames(Pos); Tab(20); AvgMinutes(Pos); Tab(30); AvgPoints(Pos); Tab(40); AvgRebounds(Pos); Tab(50); AvgTurnovers(Pos)
    Loop
    Size = Pos
    Close #1
End Sub

Private Sub cmdOpen2_Click()
    picOtherPlayers.Cls
    picOtherPlayers.Print "Player"; Tab(20); "Min"; Tab(30); "Pts"; Tab(40); "Reb"; Tab(50); "TO"
    picOtherPlayers.Print "-------------------------------------------------------------------------------------------"
    Dim Pos As Integer
    Dim Size As Integer
    
    Pos = 0
    Open fileTeam2.Path & "\" & fileTeam2.FileName For Input As #1  'opens selected file from file list
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, OtherNames(Pos), OAvgMinutes(Pos), OAvgPoints(Pos), OAvgRebounds(Pos), OAvgTurnovers(Pos)
        picOtherPlayers.Print OtherNames(Pos); Tab(20); OAvgMinutes(Pos); Tab(30); OAvgPoints(Pos); Tab(40); OAvgRebounds(Pos); Tab(50); OAvgTurnovers(Pos)
    Loop
    Close #1
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSaveOTeamAvg_Click()
    Dim OutFile As String
    OutFile = Replace(fileTeam2.FileName, "Team", "")  'changes file selected to new file for data to be saved in this new file name
    Open fileTeam2.Path & "\" & OutFile For Append As #1
    Write #1, OTeamAvgPoints, OTeamAvgMinutes, OTeamAvgRebounds, OTeamAvgTurnovers
    Close #1
End Sub

Private Sub cmdSaveTeamAvg_Click()
    Dim OutFile As String
    OutFile = Replace(fileTeam.FileName, "Team", "")
    Open fileTeam.Path & "\" & OutFile For Append As #1
    Write #1, TeamAvgPoints, TeamAvgMinutes, TeamAvgRebounds, TeamAvgTurnovers
    Close #1
End Sub

Private Sub cmdSortMinutes_Click()
    Dim Pos, Pass As Integer
    Dim Temp, Temp1, Temp2, Temp3, OTemp, OTemp1, OTemp2, OTemp3  As Single
    Dim Temp4, OTemp4 As String
    For Pass = 1 To (Size - 1)
        For Pos = 1 To (Size - Pass)
            If AvgMinutes(Pos) < AvgMinutes(Pos + 1) Then   'sorts stats by highest to lowest
                Temp = AvgTurnovers(Pos)
                AvgTurnovers(Pos) = AvgTurnovers(Pos + 1)
                AvgTurnovers(Pos + 1) = Temp
                Temp1 = AvgPoints(Pos)
                AvgPoints(Pos) = AvgPoints(Pos + 1)
                AvgPoints(Pos + 1) = Temp1
                Temp2 = AvgMinutes(Pos)
                AvgMinutes(Pos) = AvgMinutes(Pos + 1)
                AvgMinutes(Pos + 1) = Temp2
                Temp3 = AvgRebounds(Pos)
                AvgRebounds(Pos) = AvgRebounds(Pos + 1)
                AvgRebounds(Pos + 1) = Temp3
                Temp4 = HomeNames(Pos)
                HomeNames(Pos) = HomeNames(Pos + 1)
                HomeNames(Pos + 1) = Temp4
            End If
            If OAvgMinutes(Pos) < OAvgMinutes(Pos + 1) Then  'sorts opposing teams stats highest to lowest as well, this is repeated for all other stats in their respected command buttons
                OTemp = OAvgTurnovers(Pos)
                OAvgTurnovers(Pos) = OAvgTurnovers(Pos + 1)
                OAvgTurnovers(Pos + 1) = OTemp
                OTemp1 = OAvgPoints(Pos)
                OAvgPoints(Pos) = OAvgPoints(Pos + 1)
                OAvgPoints(Pos + 1) = OTemp1
                OTemp2 = OAvgMinutes(Pos)
                OAvgMinutes(Pos) = OAvgMinutes(Pos + 1)
                OAvgMinutes(Pos + 1) = OTemp2
                OTemp3 = OAvgRebounds(Pos)
                OAvgRebounds(Pos) = OAvgRebounds(Pos + 1)
                OAvgRebounds(Pos + 1) = OTemp3
                OTemp4 = OtherNames(Pos)
                OtherNames(Pos) = OtherNames(Pos + 1)
                OtherNames(Pos + 1) = OTemp4
            End If
        Next Pos
    Next Pass
    Pos = 0
    picHomePlayers.Cls
    picHomePlayers.Print "Player"; Tab(20); "Min"; Tab(30); "Pts"; Tab(40); "Reb"; Tab(50); "TO"
    picHomePlayers.Print "-------------------------------------------------------------------------------------------"
    For Pos = 1 To Size
        picHomePlayers.Print HomeNames(Pos); Tab(20); AvgMinutes(Pos); Tab(30); AvgPoints(Pos); Tab(40); AvgRebounds(Pos); Tab(50); AvgTurnovers(Pos)
    Next Pos
    Pos = 0
    picOtherPlayers.Cls
    picOtherPlayers.Print "Player"; Tab(20); "Min"; Tab(30); "Pts"; Tab(40); "Reb"; Tab(50); "TO"
    picOtherPlayers.Print "-------------------------------------------------------------------------------------------"
    For Pos = 1 To Size
        picOtherPlayers.Print OtherNames(Pos); Tab(20); OAvgMinutes(Pos); Tab(30); OAvgPoints(Pos); Tab(40); OAvgRebounds(Pos); Tab(50); OAvgTurnovers(Pos)
    Next Pos
End Sub

Private Sub cmdSortPoints_Click()
    Dim Pos, Pass As Integer
    Dim Temp, Temp1, Temp2, Temp3, OTemp, OTemp1, OTemp2, OTemp3  As Single
    Dim Temp4, OTemp4 As String
    For Pass = 1 To (Size - 1)
        For Pos = 1 To (Size - Pass)
            If AvgPoints(Pos) < AvgPoints(Pos + 1) Then
                Temp = AvgTurnovers(Pos)
                AvgTurnovers(Pos) = AvgTurnovers(Pos + 1)
                AvgTurnovers(Pos + 1) = Temp
                Temp1 = AvgPoints(Pos)
                AvgPoints(Pos) = AvgPoints(Pos + 1)
                AvgPoints(Pos + 1) = Temp1
                Temp2 = AvgMinutes(Pos)
                AvgMinutes(Pos) = AvgMinutes(Pos + 1)
                AvgMinutes(Pos + 1) = Temp2
                Temp3 = AvgRebounds(Pos)
                AvgRebounds(Pos) = AvgRebounds(Pos + 1)
                AvgRebounds(Pos + 1) = Temp3
                Temp4 = HomeNames(Pos)
                HomeNames(Pos) = HomeNames(Pos + 1)
                HomeNames(Pos + 1) = Temp4
            End If
            If OAvgPoints(Pos) < OAvgPoints(Pos + 1) Then
                OTemp = OAvgTurnovers(Pos)
                OAvgTurnovers(Pos) = OAvgTurnovers(Pos + 1)
                OAvgTurnovers(Pos + 1) = OTemp
                OTemp1 = OAvgPoints(Pos)
                OAvgPoints(Pos) = OAvgPoints(Pos + 1)
                OAvgPoints(Pos + 1) = OTemp1
                OTemp2 = OAvgMinutes(Pos)
                OAvgMinutes(Pos) = OAvgMinutes(Pos + 1)
                OAvgMinutes(Pos + 1) = OTemp2
                OTemp3 = OAvgRebounds(Pos)
                OAvgRebounds(Pos) = OAvgRebounds(Pos + 1)
                OAvgRebounds(Pos + 1) = OTemp3
                OTemp4 = OtherNames(Pos)
                OtherNames(Pos) = OtherNames(Pos + 1)
                OtherNames(Pos + 1) = OTemp4
            End If
        Next Pos
    Next Pass
    Pos = 0
    picHomePlayers.Cls
    picHomePlayers.Print "Player"; Tab(20); "Min"; Tab(30); "Pts"; Tab(40); "Reb"; Tab(50); "TO"
    picHomePlayers.Print "-------------------------------------------------------------------------------------------"
    For Pos = 1 To Size
        picHomePlayers.Print HomeNames(Pos); Tab(20); AvgMinutes(Pos); Tab(30); AvgPoints(Pos); Tab(40); AvgRebounds(Pos); Tab(50); AvgTurnovers(Pos)
    Next Pos
    Pos = 0
    picOtherPlayers.Cls
    picOtherPlayers.Print "Player"; Tab(20); "Min"; Tab(30); "Pts"; Tab(40); "Reb"; Tab(50); "TO"
    picOtherPlayers.Print "-------------------------------------------------------------------------------------------"
    For Pos = 1 To Size
        picOtherPlayers.Print OtherNames(Pos); Tab(20); OAvgMinutes(Pos); Tab(30); OAvgPoints(Pos); Tab(40); OAvgRebounds(Pos); Tab(50); OAvgTurnovers(Pos)
    Next Pos
End Sub

Private Sub cmdSortRebounds_Click()
    Dim Pos, Pass As Integer
    Dim Temp, Temp1, Temp2, Temp3, OTemp, OTemp1, OTemp2, OTemp3  As Single
    Dim Temp4, OTemp4 As String
    For Pass = 1 To (Size - 1)
        For Pos = 1 To (Size - Pass)
            If AvgRebounds(Pos) < AvgRebounds(Pos + 1) Then
                Temp = AvgTurnovers(Pos)
                AvgTurnovers(Pos) = AvgTurnovers(Pos + 1)
                AvgTurnovers(Pos + 1) = Temp
                Temp1 = AvgPoints(Pos)
                AvgPoints(Pos) = AvgPoints(Pos + 1)
                AvgPoints(Pos + 1) = Temp1
                Temp2 = AvgMinutes(Pos)
                AvgMinutes(Pos) = AvgMinutes(Pos + 1)
                AvgMinutes(Pos + 1) = Temp2
                Temp3 = AvgRebounds(Pos)
                AvgRebounds(Pos) = AvgRebounds(Pos + 1)
                AvgRebounds(Pos + 1) = Temp3
                Temp4 = HomeNames(Pos)
                HomeNames(Pos) = HomeNames(Pos + 1)
                HomeNames(Pos + 1) = Temp4
            End If
            If OAvgRebounds(Pos) < OAvgRebounds(Pos + 1) Then
                OTemp = OAvgTurnovers(Pos)
                OAvgTurnovers(Pos) = OAvgTurnovers(Pos + 1)
                OAvgTurnovers(Pos + 1) = OTemp
                OTemp1 = OAvgPoints(Pos)
                OAvgPoints(Pos) = OAvgPoints(Pos + 1)
                OAvgPoints(Pos + 1) = OTemp1
                OTemp2 = OAvgMinutes(Pos)
                OAvgMinutes(Pos) = OAvgMinutes(Pos + 1)
                OAvgMinutes(Pos + 1) = OTemp2
                OTemp3 = OAvgRebounds(Pos)
                OAvgRebounds(Pos) = OAvgRebounds(Pos + 1)
                OAvgRebounds(Pos + 1) = OTemp3
                OTemp4 = OtherNames(Pos)
                OtherNames(Pos) = OtherNames(Pos + 1)
                OtherNames(Pos + 1) = OTemp4
            End If
        Next Pos
    Next Pass
    Pos = 0
    picHomePlayers.Cls
    picHomePlayers.Print "Player"; Tab(20); "Min"; Tab(30); "Pts"; Tab(40); "Reb"; Tab(50); "TO"
    picHomePlayers.Print "-------------------------------------------------------------------------------------------"
    For Pos = 1 To Size
        picHomePlayers.Print HomeNames(Pos); Tab(20); AvgMinutes(Pos); Tab(30); AvgPoints(Pos); Tab(40); AvgRebounds(Pos); Tab(50); AvgTurnovers(Pos)
    Next Pos
    Pos = 0
    picOtherPlayers.Cls
    picOtherPlayers.Print "Player"; Tab(20); "Min"; Tab(30); "Pts"; Tab(40); "Reb"; Tab(50); "TO"
    picOtherPlayers.Print "-------------------------------------------------------------------------------------------"
    For Pos = 1 To Size
        picOtherPlayers.Print OtherNames(Pos); Tab(20); OAvgMinutes(Pos); Tab(30); OAvgPoints(Pos); Tab(40); OAvgRebounds(Pos); Tab(50); OAvgTurnovers(Pos)
    Next Pos
End Sub

Private Sub cmdSortTurnovers_Click()
    Dim Pos, Pass As Integer
    Dim Temp, Temp1, Temp2, Temp3, OTemp, OTemp1, OTemp2, OTemp3  As Single
    Dim Temp4, OTemp4 As String
    For Pass = 1 To (Size - 1)
        For Pos = 1 To (Size - Pass)
            If AvgTurnovers(Pos) < AvgTurnovers(Pos + 1) Then
                Temp = AvgTurnovers(Pos)
                AvgTurnovers(Pos) = AvgTurnovers(Pos + 1)
                AvgTurnovers(Pos + 1) = Temp
                Temp1 = AvgPoints(Pos)
                AvgPoints(Pos) = AvgPoints(Pos + 1)
                AvgPoints(Pos + 1) = Temp1
                Temp2 = AvgMinutes(Pos)
                AvgMinutes(Pos) = AvgMinutes(Pos + 1)
                AvgMinutes(Pos + 1) = Temp2
                Temp3 = AvgRebounds(Pos)
                AvgRebounds(Pos) = AvgRebounds(Pos + 1)
                AvgRebounds(Pos + 1) = Temp3
                Temp4 = HomeNames(Pos)
                HomeNames(Pos) = HomeNames(Pos + 1)
                HomeNames(Pos + 1) = Temp4
            End If
            If OAvgTurnovers(Pos) < OAvgTurnovers(Pos + 1) Then
                OTemp = OAvgTurnovers(Pos)
                OAvgTurnovers(Pos) = OAvgTurnovers(Pos + 1)
                OAvgTurnovers(Pos + 1) = OTemp
                OTemp1 = OAvgPoints(Pos)
                OAvgPoints(Pos) = OAvgPoints(Pos + 1)
                OAvgPoints(Pos + 1) = OTemp1
                OTemp2 = OAvgMinutes(Pos)
                OAvgMinutes(Pos) = OAvgMinutes(Pos + 1)
                OAvgMinutes(Pos + 1) = OTemp2
                OTemp3 = OAvgRebounds(Pos)
                OAvgRebounds(Pos) = OAvgRebounds(Pos + 1)
                OAvgRebounds(Pos + 1) = OTemp3
                OTemp4 = OtherNames(Pos)
                OtherNames(Pos) = OtherNames(Pos + 1)
                OtherNames(Pos + 1) = OTemp4
            End If
        Next Pos
    Next Pass
    Pos = 0
    picHomePlayers.Cls
    picHomePlayers.Print "Player"; Tab(20); "Min"; Tab(30); "Pts"; Tab(40); "Reb"; Tab(50); "TO"
    picHomePlayers.Print "-------------------------------------------------------------------------------------------"
    For Pos = 1 To Size
        picHomePlayers.Print HomeNames(Pos); Tab(20); AvgMinutes(Pos); Tab(30); AvgPoints(Pos); Tab(40); AvgRebounds(Pos); Tab(50); AvgTurnovers(Pos)
    Next Pos
    Pos = 0
    picOtherPlayers.Cls
    picOtherPlayers.Print "Player"; Tab(20); "Min"; Tab(30); "Pts"; Tab(40); "Reb"; Tab(50); "TO"
    picOtherPlayers.Print "-------------------------------------------------------------------------------------------"
    For Pos = 1 To Size
        picOtherPlayers.Print OtherNames(Pos); Tab(20); OAvgMinutes(Pos); Tab(30); OAvgPoints(Pos); Tab(40); OAvgRebounds(Pos); Tab(50); OAvgTurnovers(Pos)
    Next Pos
End Sub

