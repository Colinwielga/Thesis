VERSION 5.00
Begin VB.Form frmSortTwins 
   Caption         =   "Sort Twins Hitters/Pitchers"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   Picture         =   "frmSortTwins.frx":0000
   ScaleHeight     =   6975
   ScaleWidth      =   10080
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   27
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdKs 
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   24
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdEarnedRunAv 
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   23
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdHitts 
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   22
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdHomeRuns 
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   21
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdAverage 
      BackColor       =   &H80000000&
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   20
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdSortSO 
      Caption         =   "Sort by Strike Outs"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   18
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdSortERA 
      Caption         =   "Sort by ERA"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      TabIndex        =   17
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdHits 
      Caption         =   "Sort by Hits"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   16
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdSortHR 
      Caption         =   "Sort by Home Runs"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   15
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdSortAvg 
      Caption         =   "Sort by Batting Average"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      TabIndex        =   14
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtSO 
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtERA 
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtHits 
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtHR 
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtAvg 
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.PictureBox picResults2 
      Height          =   5175
      Left            =   5280
      Picture         =   "frmSortTwins.frx":118B3
      ScaleHeight     =   5115
      ScaleWidth      =   4275
      TabIndex        =   3
      Top             =   480
      Width           =   4335
   End
   Begin VB.CommandButton cmdCalcOwn 
      Caption         =   "Calculate Your Own Statistics"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   2
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdGoToPitching 
      Caption         =   "Go To Pitching Statistics"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      TabIndex        =   1
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdGoToHitting 
      Caption         =   "Go To Hitting Statistics"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label lblPitching 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pitching:"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   26
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label lblBatting 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hitting:"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   25
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Show all players with above:"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label lblSO 
      BackStyle       =   0  'Transparent
      Caption         =   "Strike Outs"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblERA 
      BackStyle       =   0  'Transparent
      Caption         =   "ERA"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblHits 
      BackStyle       =   0  'Transparent
      Caption         =   "Hits"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblHR 
      BackStyle       =   0  'Transparent
      Caption         =   "Home Runs"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblAvg 
      BackStyle       =   0  'Transparent
      Caption         =   "Batting AVG"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "frmSortTwins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pos As Integer
    'Twins Statistics, frmSortTwins, By Mike Patnode, Written Nov.2, 2006, Objective to diplay Twins Stats
Private Sub cmdAverage_Click()
    'Input the arrays
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
    picResults2.Cls
    'Clear the Picture box
    picResults2.Print "Name"; Tab(20); "AVG"; Tab(30)
    picResults2.Print "-----------------------------------------------"
    Pos = 0
    Dim Average As Single
    Average = txtAvg.Text 'Input Avg you wish to sort by
    For Pos = 1 To Counter
        If TwinsAVG(Pos) >= Average Then
            picResults2.Print TwinsName(Pos); Tab(20); FormatNumber(TwinsAVG(Pos), 3); Tab(30)
        End If 'if the Avg is greater than or equal to input then display player name and avg
    Next Pos
End Sub

Private Sub cmdCalcOwn_Click()
    'Go from frmSortTwins to frmCalcOwnStats
    frmHittingStats.Hide
    frmPitchingStats.Hide
    frmCalcOwnStats.Show
    frmSortTwins.Hide
End Sub



Private Sub cmdEarnedRunAv_Click()
Open App.Path & "\pitcherstats.txt" For Input As #2
    Counter2 = 0
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
    picResults2.Cls
    picResults2.Print "Name"; Tab(20); "ERA"; Tab(30)
    picResults2.Print "-----------------------------------------------"
    Pos = 0
    Dim ERA As Single
    ERA = txtERA.Text
    For Pos = 1 To Counter2
        If PitchERA(Pos) <= ERA Then
            picResults2.Print PitchNames(Pos); Tab(20); PitchERA(Pos); Tab(30)
        End If 'If ERA is less then specified amount, then display name and ERA
    Next Pos
End Sub

Private Sub cmdGoToHitting_Click()
    'go from frmSortTwins to frmHitting
    frmHittingStats.Show
    frmPitchingStats.Hide
    frmCalcOwnStats.Hide
    frmSortTwins.Hide
End Sub

Private Sub cmdGoToPitching_Click()
    'go from frmSortTwins to frmPitching
    frmHittingStats.Hide
    frmPitchingStats.Show
    frmCalcOwnStats.Hide
    frmSortTwins.Hide
End Sub

Private Sub cmdHits_Click()
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
    Dim Pass As Integer
    Dim Comp As Integer
    Dim TempHits As Single
    Dim TempName As String
    Dim Pos As Integer
    Pass = 0
    For Pass = 1 To (Counter - 1) 'Sorts players by number of total hits
        For Comp = 1 To (Counter - Pass)
            If TwinsHits(Comp) < TwinsHits(Comp + 1) Then
            TempHits = TwinsHits(Comp)
            TwinsHits(Comp) = TwinsHits(Comp + 1)
            TwinsHits(Comp + 1) = TempHits
            
            TempName = TwinsName(Comp)
            TwinsName(Comp) = TwinsName(Comp + 1)
            TwinsName(Comp + 1) = TempName
            End If
        Next Comp
    Next Pass
    picResults2.Cls
    picResults2.Print "Name"; Tab(20); "Hits"; Tab(30)
    picResults2.Print "------------------------------------------------"
    For Comp = 1 To Counter
        picResults2.Print TwinsName(Comp); Tab(20); TwinsHits(Comp); Tab(30); ""
    Next Comp 'displays the results after sorting
End Sub

Private Sub cmdHitts_Click()
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
    picResults2.Cls
    picResults2.Print "Name"; Tab(20); "Hits"; Tab(30)
    picResults2.Print "-----------------------------------------------"
    Pos = 0
    Dim Hits As Integer
    Hits = txtHits.Text
    For Pos = 1 To Counter
        If TwinsHits(Pos) >= Hits Then
            picResults2.Print TwinsName(Pos); Tab(20); TwinsHits(Pos); Tab(30)
        End If
    Next Pos
End Sub

Private Sub cmdHomeRuns_Click()
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
    picResults2.Cls
    picResults2.Print "Name"; Tab(20); "Home Runs"; Tab(30)
    picResults2.Print "-----------------------------------------------"
    Pos = 0
    Dim HomeRuns As Integer
    HomeRuns = txtHR.Text
    For Pos = 1 To Counter
        If TwinsHR(Pos) >= HomeRuns Then
            picResults2.Print TwinsName(Pos); Tab(20); TwinsHR(Pos); Tab(30)
        End If
    Next Pos
    
End Sub

Private Sub cmdKs_Click()
     Open App.Path & "\pitcherstats.txt" For Input As #2
    Counter2 = 0
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
    picResults2.Cls
    picResults2.Print "Name"; Tab(20); "Strike Outs"; Tab(30)
    picResults2.Print "-----------------------------------------------"
    Pos = 0
    Dim Strike As Integer
    Strike = txtSO.Text
    For Pos = 1 To Counter2
        If PitchSO(Pos) >= Strike Then
            picResults2.Print PitchNames(Pos); Tab(20); PitchSO(Pos); Tab(30)
        End If
    Next Pos
End Sub



Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSortAvg_Click()
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
    Dim Pass As Integer
    Dim Comp As Integer
    Dim TempAVG As Single
    Dim TempName As String
    Pass = 0
    For Pass = 1 To (Counter - 1)
        For Comp = 1 To (Counter - Pass)
            If TwinsAVG(Comp) < TwinsAVG(Comp + 1) Then
            TempAVG = TwinsAVG(Comp)
            TwinsAVG(Comp) = TwinsAVG(Comp + 1)
            TwinsAVG(Comp + 1) = TempAVG
            
            TempName = TwinsName(Comp)
            TwinsName(Comp) = TwinsName(Comp + 1)
            TwinsName(Comp + 1) = TempName
            End If
        Next Comp
    Next Pass
    picResults2.Cls
    picResults2.Print "Name"; Tab(20); "AVG"; Tab(30)
    picResults2.Print "------------------------------------------------"
    For Comp = 1 To Counter
        picResults2.Print TwinsName(Comp); Tab(20); TwinsAVG(Comp); Tab(30); ""
    Next Comp
    
End Sub

Private Sub cmdSortERA_Click()
    Open App.Path & "\pitcherstats.txt" For Input As #2
    Counter2 = 0
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
    Dim Pass As Integer
    Dim Comp As Integer
    Dim TempERA As Single
    Dim TempName As String
    Dim Pos As Integer
    Pass = 0
    For Pass = 1 To (Counter2 - 1)
        For Comp = 1 To (Counter2 - Pass)
            If PitchERA(Comp) > PitchERA(Comp + 1) Then
            TempERA = PitchERA(Comp)
            PitchERA(Comp) = PitchERA(Comp + 1)
            PitchERA(Comp + 1) = TempERA
            
            TempName = PitchNames(Comp)
            PitchNames(Comp) = PitchNames(Comp + 1)
            PitchNames(Comp + 1) = TempName
            End If
        Next Comp
    Next Pass
    picResults2.Cls
    picResults2.Print "Name"; Tab(20); "ERA"; Tab(30)
    picResults2.Print "------------------------------------------------"
    For Comp = 1 To Counter2
        picResults2.Print PitchNames(Comp); Tab(20); PitchERA(Comp); Tab(30); ""
    Next Comp
End Sub

Private Sub cmdSortHR_Click()
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
    Dim Pass As Integer
    Dim Comp As Integer
    Dim TempHR As Single
    Dim TempName As String
    Dim Pos As Integer
    Pass = 0
    For Pass = 1 To (Counter - 1)
        For Comp = 1 To (Counter - Pass)
            If TwinsHR(Comp) < TwinsHR(Comp + 1) Then
            TempHR = TwinsHR(Comp)
            TwinsHR(Comp) = TwinsHR(Comp + 1)
            TwinsHR(Comp + 1) = TempHR
            
            TempName = TwinsName(Comp)
            TwinsName(Comp) = TwinsName(Comp + 1)
            TwinsName(Comp + 1) = TempName
            End If
        Next Comp
    Next Pass
    picResults2.Cls
    picResults2.Print "Name"; Tab(20); "Home Runs"; Tab(30)
    picResults2.Print "------------------------------------------------"
    For Comp = 1 To Counter
        picResults2.Print TwinsName(Comp); Tab(20); TwinsHR(Comp); Tab(30); ""
    Next Comp
         
End Sub


Private Sub cmdSortSO_Click()
    Open App.Path & "\pitcherstats.txt" For Input As #2
    Counter2 = 0
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
    Dim Pass As Integer
    Dim Comp As Integer
    Dim TempSO As Single
    Dim TempName As String
    Dim Pos As Integer
    Pass = 0
    For Pass = 1 To (Counter2 - 1)
        For Comp = 1 To (Counter2 - Pass)
            If PitchSO(Comp) < PitchSO(Comp + 1) Then
            TempSO = PitchSO(Comp)
            PitchSO(Comp) = PitchSO(Comp + 1)
            PitchSO(Comp + 1) = TempSO
            
            TempName = PitchNames(Comp)
            PitchNames(Comp) = PitchNames(Comp + 1)
            PitchNames(Comp + 1) = TempName
            End If
            picResults2.Print PitchNames(Comp); Tab(20); PitchSO(Comp); Tab(30); ""
        Next Comp
    Next Pass
    picResults2.Cls
    picResults2.Print "Name"; Tab(20); "Strike Outs"; Tab(30)
    picResults2.Print "------------------------------------------------"
    For Comp = 1 To Counter2
        picResults2.Print PitchNames(Comp); Tab(20); PitchSO(Comp); Tab(30); ""
    Next Comp
End Sub
