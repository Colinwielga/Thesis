VERSION 5.00
Begin VB.Form frmHittingStats 
   BackColor       =   &H00800000&
   Caption         =   "Twins Hitting Stats"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10545
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLogo 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8760
      Picture         =   "frmHittingStats.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1635
      TabIndex        =   6
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   5
      Top             =   6960
      Width           =   1095
   End
   Begin VB.PictureBox picDisplayResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   240
      Picture         =   "frmHittingStats.frx":0BCB
      ScaleHeight     =   5595
      ScaleWidth      =   10035
      TabIndex        =   4
      Top             =   600
      Width           =   10095
   End
   Begin VB.CommandButton cmdGoToCalc 
      Caption         =   "Calculate Your Own Statistics"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5160
      TabIndex        =   3
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton cmdGoToPitching 
      Caption         =   "Go To Pitching Statistics"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   2
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton cmdGoToSort 
      Caption         =   "Go To Sort Option"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      TabIndex        =   1
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton cmdDisplayHitting 
      BackColor       =   &H80000000&
      Caption         =   "Display Hitting Statitstics"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmHittingStats.frx":196FD
      TabIndex        =   0
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2006 Minnesota Twins Statistics, By: Mike Patnode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   10095
   End
End
Attribute VB_Name = "frmHittingStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 'Twins Statistics, frmHittingStats, By Mike Patnode, Written Nov.2, 2006, Objective to diplay Twins Stats
Private Sub cmdDisplayHitting_Click()
    'Load hitting statistics array
    Counter = 0
    Open App.Path & "\PlayerStats.txt" For Input As #1
    Do Until EOF(1)
        Counter = Counter + 1
        Input #1, TName, TPos, THits, TRuns, THR, TRBI, TBB, TSO, TOBP, TAVG
        TwinsName(Counter) = TName
        TwinsPos(Counter) = TPos
        TwinsHits(Counter) = THits
        TwinsRuns(Counter) = TRuns
        TwinsHR(Counter) = THR
        TwinsRBI(Counter) = TRBI
        TwinsBB(Counter) = TBB
        TwinsSO(Counter) = TSO
        TwinsOBP(Counter) = TOBP
        TwinsAVG(Counter) = TAVG
    Loop
    Close #1
    picDisplayResults.Cls
    picDisplayResults.Print "Name"; Tab(20); "Position"; Tab(35); "Hits"; Tab(45); "Runs"; Tab(55); "HR"; Tab(65); "RBI"; Tab(75); "BB"; Tab(85); "SO"; Tab(95); "OBP"; Tab(105); "AVG"; Tab(115)
    picDisplayResults.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Dim Size As Integer
    For Size = 1 To Counter
        picDisplayResults.Print TwinsName(Size); Tab(20); TwinsPos(Size); Tab(35); TwinsHits(Size); Tab(45); TwinsRuns(Size); Tab(55); TwinsHR(Size); Tab(65); TwinsRBI(Size); Tab(75); TwinsBB(Size); Tab(85); TwinsSO(Size); Tab(95); FormatNumber(TwinsOBP(Size), 3); Tab(105); FormatNumber(TwinsAVG(Size), 3); Tab(115)
    Next Size 'Print out entire hitting array
End Sub

Private Sub cmdGoToCalc_Click()
    'switching forms
    frmHittingStats.Hide
    frmPitchingStats.Hide
    frmCalcOwnStats.Show
    frmSortTwins.Hide
End Sub

Private Sub cmdGoToPitching_Click()
    'switching forms
    frmHittingStats.Hide
    frmPitchingStats.Show
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
