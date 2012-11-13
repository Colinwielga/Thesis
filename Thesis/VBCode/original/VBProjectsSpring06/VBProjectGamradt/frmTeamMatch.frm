VERSION 5.00
Begin VB.Form frmTeamMatch 
   BackColor       =   &H00808000&
   Caption         =   "Team Matchup"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   Picture         =   "frmTeamMatch.frx":0000
   ScaleHeight     =   7515
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox fileTeam2 
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   390
      Left            =   7200
      Pattern         =   "*_.txt*"
      TabIndex        =   11
      Top             =   3240
      Width           =   2775
   End
   Begin VB.CommandButton cmdWinner 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Projected Winner"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   10
      Top             =   4920
      Width           =   2535
   End
   Begin VB.CommandButton cmdOpenOtherTeam 
      Caption         =   "Open Selected Team #2"
      Height          =   255
      Left            =   7680
      TabIndex        =   9
      Top             =   4320
      Width           =   1935
   End
   Begin VB.PictureBox picOtherTeam 
      Height          =   495
      Left            =   6480
      ScaleHeight     =   435
      ScaleWidth      =   4035
      TabIndex        =   8
      Top             =   3720
      Width           =   4095
   End
   Begin VB.PictureBox picHomeTeam 
      Height          =   495
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   4155
      TabIndex        =   7
      Top             =   3720
      Width           =   4215
   End
   Begin VB.CommandButton cmdNavigatePlayerMatch 
      Caption         =   "Player Matchup"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdNavigateMainMenu 
      Caption         =   "Main Menu"
      Height          =   495
      Left            =   8040
      TabIndex        =   3
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpenHome 
      Caption         =   "Open Selected Team #1"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   4320
      Width           =   1935
   End
   Begin VB.FileListBox fileTeam 
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   390
      Left            =   1200
      Pattern         =   "*_.txt*"
      TabIndex        =   1
      Top             =   3240
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   9600
      TabIndex        =   0
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label lblDirection2 
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
      Left            =   7200
      TabIndex        =   15
      Top             =   3000
      Width           =   3015
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
      Left            =   1200
      TabIndex        =   14
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label lblPlayerMatchup 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Team Matchup"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   555
      Left            =   3840
      TabIndex        =   13
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label lblVS 
      BackColor       =   &H00808000&
      Caption         =   "VS"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   915
      Left            =   4920
      TabIndex        =   12
      Top             =   3600
      Width           =   1335
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
      Left            =   240
      TabIndex        =   6
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Team Manager Pro"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmTeamMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Team Manager Pro (ErikGamradtVBProject.vbp)
'frmMainMenu (frmMainMenu.frm)
'Designed By: Erik Gamradt
'22 March 2006
'In this menu, the user is able to access the saved team average statistics by selected the desired team and opening the file to be displayed.  There are two teams to be selected so the user can compare the desired matchup.  Additionally, the user can find out who the projected winner is by selecting Projected Winner after teams have been selected from file lists.  A message box will appear depicting who the winner is based on team statistics.
Option Explicit
Dim OTeamAvgPoints, OTeamAvgMinutes, OTeamAvgRebounds, OTeamAvgTurnovers As Single
Dim TeamAvgPoints, TeamAvgMinutes, TeamAvgRebounds, TeamAvgTurnovers As Single
Private Sub cmdNavigateMainMenu_Click()
    frmMainMenu.Show
    frmTeamMatch.Hide
End Sub

Private Sub cmdNavigatePlayerMatch_Click()
    frmTeamMatch.Hide
    frmPlayerMatch.Show
End Sub


Private Sub cmdOpenHome_Click()
    picHomeTeam.Cls
    picHomeTeam.Print "PPG"; Tab(15); "MPG"; Tab(30); "RPG"; Tab(45); "TPG"  'prints headings in display
    Open fileTeam.Path & "\" & fileTeam.FileName For Input As #1
    Do Until EOF(1)
        Input #1, TeamAvgPoints, TeamAvgMinutes, TeamAvgRebounds, TeamAvgTurnovers
        picHomeTeam.Print FormatNumber(TeamAvgPoints); Tab(15); FormatNumber(TeamAvgMinutes); Tab(30); FormatNumber(TeamAvgRebounds); Tab(45); FormatNumber(TeamAvgTurnovers)
    Loop
    Close #1
End Sub

Private Sub cmdOpenOtherTeam_Click()
    picOtherTeam.Cls
    picOtherTeam.Print "PPG"; Tab(15); "MPG"; Tab(30); "RPG"; Tab(45); "TPG"
    Open fileTeam2.Path & "\" & fileTeam2.FileName For Input As #1
    Do Until EOF(1)
        Input #1, OTeamAvgPoints, OTeamAvgMinutes, OTeamAvgRebounds, OTeamAvgTurnovers
        picOtherTeam.Print FormatNumber(OTeamAvgPoints); Tab(15); FormatNumber(OTeamAvgMinutes); Tab(30); FormatNumber(OTeamAvgRebounds); Tab(45); FormatNumber(OTeamAvgTurnovers)
    Loop
    Close #1
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdWinner_Click()
    Dim HomeSum As Single
    Dim OtherSum As Single
    HomeSum = TeamAvgPoints + TeamAvgRebounds - TeamAvgTurnovers  'the magic test to see which team is predicted to win
    OtherSum = OTeamAvgPoints + OTeamAvgRebounds - OTeamAvgTurnovers
    If HomeSum > OtherSum Then
        MsgBox "The Predicted Winner is Team #1", vbInformation, "Predicted Winner!"
    ElseIf OtherSum > HomeSum Then
        MsgBox "The Predicted Winner is Team #2", vbInformation, "Predicted Winner!"
    Else
        MsgBox "This matchup is too close to call!", vbInformation, "Predicted Winner!"
    End If
End Sub
