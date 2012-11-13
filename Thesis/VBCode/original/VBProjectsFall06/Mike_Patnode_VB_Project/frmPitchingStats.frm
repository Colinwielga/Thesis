VERSION 5.00
Begin VB.Form frmPitchingStats 
   BackColor       =   &H80000007&
   Caption         =   "Twins Pitching Stats"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10140
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   5
      Top             =   5520
      Width           =   975
   End
   Begin VB.PictureBox picDisplay 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   1080
      Picture         =   "frmPitchingStats.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   7755
      TabIndex        =   4
      Top             =   240
      Width           =   7815
   End
   Begin VB.CommandButton cmdGoToCalc 
      Caption         =   "Calculate Your Own Statistics"
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      TabIndex        =   3
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdGoToSort 
      Caption         =   "Go To Sort Option"
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   2
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdGoToHitting 
      Caption         =   "Go To Hitting Statistics"
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   1
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdDisplayPitching 
      Caption         =   "Display Pitching Statistics"
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   5280
      Width           =   1455
   End
End
Attribute VB_Name = "frmPitchingStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Size As Integer
'Twins Statistics, frmPitchingStats, By Mike Patnode, Written Nov.2, 2006, Objective to diplay Twins Stats




Private Sub cmdDisplayPitching_Click()
    'Input array
    Open App.Path & "\pitcherstats.txt" For Input As #2
    Counter = 0
    Do Until EOF(2)
        Input #2, PNames, PWins, PLoss, PERA, PSaves, PHBP, PBB, PSO
        Counter2 = Counter2 + 1
        PitchNames(Counter2) = PNames
        PitchWins(Counter2) = PWins
        PitchLoss(Counter2) = PLoss
        PitchERA(Counter2) = PERA
        PitchSaves(Counter2) = PSaves
        PitchHBP(Counter2) = PHBP
        PitchBB(Counter2) = PBB
        PitchSO(Counter2) = PSO
    Loop
    Close #2
    picDisplay.Cls
    picDisplay.Print "Names"; Tab(23); "Wins"; Tab(30); "Losses"; Tab(40); "ERA"; Tab(50); "Saves"; Tab(60); "HBP"; Tab(70); "BB"; Tab(80); "SO"; Tab(90)
    picDisplay.Print "---------------------------------------------------------------------------------------------------------------------------------------"
    For Size = 1 To Counter2
        picDisplay.Print PitchNames(Size); Tab(23); PitchWins(Size); Tab(30); PitchLoss(Size); Tab(40); FormatNumber(PitchERA(Size)); Tab(50); PitchSaves(Size); Tab(60); PitchHBP(Size); Tab(70); PitchBB(Size); Tab(80); PitchSO(Size); Tab(90)
    Next Size 'Display all players names and statistics
End Sub

Private Sub cmdGoToCalc_Click()
    'switching forms
    frmHittingStats.Hide
    frmPitchingStats.Hide
    frmCalcOwnStats.Show
    frmSortTwins.Hide
End Sub

Private Sub cmdGoToHitting_Click()
    'switching forms
    frmHittingStats.Show
    frmPitchingStats.Hide
    frmCalcOwnStats.Hide
    frmSortTwins.Hide
End Sub

Private Sub cmdGoToSort_Click()
    'switching forms
    frmHittingStats.Hide
    frmPitchingStats.Hide
    frmCalcOwnStats.Hide
    frmSortTwins.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub
