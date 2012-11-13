VERSION 5.00
Begin VB.Form frmRankings 
   Caption         =   "Rankings"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picGlove 
      Height          =   1455
      Left            =   5760
      Picture         =   "frmRankings.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View Rankings"
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtPosition 
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdReturnRank 
      Caption         =   "Return to Main Page"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.PictureBox picRankings 
      Height          =   5535
      Left            =   720
      ScaleHeight     =   5475
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   2280
      Width           =   6615
   End
   Begin VB.Label lblInstruct 
      Alignment       =   2  'Center
      Caption         =   $"frmRankings.frx":0DE5
      Height          =   855
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label lblChart 
      Caption         =   $"frmRankings.frx":0E77
      Height          =   735
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmRankings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form allows the user to see an ordered list of players at a position

Private Sub cmdReturnRank_Click()
    'Makes home page appear
    frmHome.Visible = True
    frmExplanation.Visible = False
    frmSearch.Visible = False
    frmRankings.Visible = False
    
End Sub

Private Sub cmdView_Click()
    'Declaring variables
    Dim userpos As Integer, ctr As Integer, size As Integer
    Dim pass As Integer, comp As Integer, n As Integer
    Dim tempname As String, tempscore As Integer
    size = 0
    'Receives position number indicator from text box
    userpos = txtPosition.Text
    'Adds player and info to corresponding position array
    If userpos = 0 Then
        For ctr = 1 To pcounter
            If PitcherPos(ctr) = 2 Then
                size = size + 1
                Relievers(size) = Pitchers(ctr)
                RelieverScores(size) = PitcherTotals(ctr)
            End If
        Next ctr
    ElseIf userpos = 1 Then
        For ctr = 1 To pcounter
            If PitcherPos(ctr) = 1 Then
                size = size + 1
                Starters(size) = Pitchers(ctr)
                StarterScores(size) = PitcherTotals(ctr)
            End If
        Next ctr
    ElseIf userpos = 2 Then
        For ctr = 1 To bcounter
            If BatterPos(ctr) = 2 Then
                size = size + 1
                Catchers(size) = Batters(ctr)
                CatcherScores(size) = BatterTotals(ctr)
            End If
        Next ctr
    ElseIf userpos = 3 Then
        For ctr = 1 To bcounter
            If BatterPos(ctr) = 3 Then
                size = size + 1
                FirstBases(size) = Batters(ctr)
                FirstBaseScores(size) = BatterTotals(ctr)
            End If
        Next ctr
    ElseIf userpos = 4 Then
        For ctr = 1 To bcounter
            If BatterPos(ctr) = 4 Then
                size = size + 1
                SecondBases(size) = Batters(ctr)
                SecondBaseScores(size) = BatterTotals(ctr)
            End If
        Next ctr
    ElseIf userpos = 5 Then
        For ctr = 1 To bcounter
            If BatterPos(ctr) = 5 Then
                size = size + 1
                ThirdBases(size) = Batters(ctr)
                ThirdBaseScores(size) = BatterTotals(ctr)
            End If
        Next ctr
    ElseIf userpos = 6 Then
        For ctr = 1 To bcounter
            If BatterPos(ctr) = 6 Then
                size = size + 1
                Shortstops(size) = Batters(ctr)
                ShortstopScores(size) = BatterTotals(ctr)
            End If
        Next ctr
    ElseIf userpos = 7 Then
        For ctr = 1 To bcounter
            If BatterPos(ctr) = 7 Then
                size = size + 1
                Corners(size) = Batters(ctr)
                CornerScores(size) = BatterTotals(ctr)
            End If
        Next ctr
    ElseIf userpos = 8 Then
        For ctr = 1 To bcounter
            If BatterPos(ctr) = 8 Then
                size = size + 1
                Centerfielders(size) = Batters(ctr)
                CenterScores(size) = BatterTotals(ctr)
            End If
        Next ctr
    Else
        MsgBox "You entered an invalid number. Please enter a new number", , "Error"
    End If
    
    picRankings.Cls
    'Orders the array for the position and prints the list
    If userpos = 0 Then
        For pass = 1 To size - 1
            For comp = 1 To size - pass
                If RelieverScores(comp) < RelieverScores(comp + 1) Then
                    tempscore = RelieverScores(comp)
                    RelieverScores(comp) = RelieverScores(comp + 1)
                    RelieverScores(comp + 1) = tempscore
                    tempname = Relievers(comp)
                    Relievers(comp) = Relievers(comp + 1)
                    Relievers(comp + 1) = tempname
                End If
            Next comp
        Next pass
        picRankings.Print "A reliever needs a score of 1450 to be deserving of the Hall of Fame "
        For n = 1 To size
            picRankings.Print Relievers(n); Tab(25); RelieverScores(n)
        Next n
    ElseIf userpos = 1 Then
         For pass = 1 To size - 1
            For comp = 1 To size - pass
                If StarterScores(comp) < StarterScores(comp + 1) Then
                    tempscore = StarterScores(comp)
                    StarterScores(comp) = StarterScores(comp + 1)
                    StarterScores(comp + 1) = tempscore
                    tempname = Starters(comp)
                    Starters(comp) = Starters(comp + 1)
                    Starters(comp + 1) = tempname
                End If
            Next comp
        Next pass
        picRankings.Print "A starter needs a score of 2300 to be deserving of the Hall of Fame"
        For n = 1 To size
            picRankings.Print Starters(n); Tab(25); StarterScores(n)
        Next n
    ElseIf userpos = 2 Then
         For pass = 1 To size - 1
            For comp = 1 To size - pass
                If CatcherScores(comp) < CatcherScores(comp + 1) Then
                    tempscore = CatcherScores(comp)
                    CatcherScores(comp) = CatcherScores(comp + 1)
                    CatcherScores(comp + 1) = tempscore
                    tempname = Catchers(comp)
                    Catchers(comp) = Catchers(comp + 1)
                    Catchers(comp + 1) = tempname
                End If
            Next comp
        Next pass
        picRankings.Print "A catcher needs a score of 2200 to be deserving of the Hall of Fame"
        For n = 1 To size
            picRankings.Print Catchers(n); Tab(25); CatcherScores(n)
        Next n
    ElseIf userpos = 3 Then
         For pass = 1 To size - 1
            For comp = 1 To size - pass
                If FirstBaseScores(comp) < FirstBaseScores(comp + 1) Then
                    tempscore = FirstBaseScores(comp)
                    FirstBaseScores(comp) = FirstBaseScores(comp + 1)
                    FirstBaseScores(comp + 1) = tempscore
                    tempname = FirstBases(comp)
                    FirstBases(comp) = FirstBases(comp + 1)
                    FirstBases(comp + 1) = tempname
                End If
            Next comp
        Next pass
        picRankings.Print "A first baseman needs a score of 2500 to be deserving of the Hall of Fame"
        For n = 1 To size
            picRankings.Print FirstBases(n); Tab(25); FirstBaseScores(n)
        Next n
    ElseIf userpos = 4 Then
         For pass = 1 To size - 1
            For comp = 1 To size - pass
                If SecondBaseScores(comp) < SecondBaseScores(comp + 1) Then
                    tempscore = SecondBaseScores(comp)
                    SecondBaseScores(comp) = SecondBaseScores(comp + 1)
                    SecondBaseScores(comp + 1) = tempscore
                    tempname = SecondBases(comp)
                    SecondBases(comp) = SecondBases(comp + 1)
                    SecondBases(comp + 1) = tempname
                End If
            Next comp
        Next pass
        picRankings.Print "A second baseman needs a score of 2300 to be deserving of the Hall of Fame"
        For n = 1 To size
            picRankings.Print SecondBases(n); Tab(25); SecondBaseScores(n)
        Next n
    ElseIf userpos = 5 Then
         For pass = 1 To size - 1
            For comp = 1 To size - pass
                If ThirdBaseScores(comp) < ThirdBaseScores(comp + 1) Then
                    tempscore = ThirdBaseScores(comp)
                    ThirdBaseScores(comp) = ThirdBaseScores(comp + 1)
                    ThirdBaseScores(comp + 1) = tempscore
                    tempname = ThirdBases(comp)
                    ThirdBases(comp) = ThirdBases(comp + 1)
                    ThirdBases(comp + 1) = tempname
                End If
            Next comp
        Next pass
        picRankings.Print "A third basemen needs a score of 2350 to be deserving of the Hall of Fame"
        For n = 1 To size
            picRankings.Print ThirdBases(n); Tab(25); ThirdBaseScores(n)
        Next n
    ElseIf userpos = 6 Then
         For pass = 1 To size - 1
            For comp = 1 To size - pass
                If ShortstopScores(comp) < ShortstopScores(comp + 1) Then
                    tempscore = ShortstopScores(comp)
                    ShortstopScores(comp) = ShortstopScores(comp + 1)
                    ShortstopScores(comp + 1) = tempscore
                    tempname = Shortstops(comp)
                    Shortstops(comp) = Shortstops(comp + 1)
                    Shortstops(comp + 1) = tempname
                End If
            Next comp
        Next pass
        picRankings.Print "A shortstop needs a score of 2200 to be deserving of the Hall of Fame"
        For n = 1 To size
            picRankings.Print Shortstops(n); Tab(25); ShortstopScores(n)
        Next n
    ElseIf userpos = 7 Then
         For pass = 1 To size - 1
            For comp = 1 To size - pass
                If CornerScores(comp) < CornerScores(comp + 1) Then
                    tempscore = CornerScores(comp)
                    CornerScores(comp) = CornerScores(comp + 1)
                    CornerScores(comp + 1) = tempscore
                    tempname = Corners(comp)
                    Corners(comp) = Corners(comp + 1)
                    Corners(comp + 1) = tempname
                End If
            Next comp
        Next pass
        picRankings.Print "A corner outfielder needs a score of 2450 to be deserving of the Hall of Fame"
        For n = 1 To size
            picRankings.Print Corners(n); Tab(25); CornerScores(n)
        Next n
    ElseIf userpos = 8 Then
         For pass = 1 To size - 1
            For comp = 1 To size - pass
                If CenterScores(comp) < CenterScores(comp + 1) Then
                    tempscore = CenterScores(comp)
                    CenterScores(comp) = CenterScores(comp + 1)
                    CenterScores(comp + 1) = tempscore
                    tempname = Centerfielders(comp)
                    Centerfielders(comp) = Centerfielders(comp + 1)
                    Centerfielders(comp + 1) = tempname
                End If
            Next comp
        Next pass
        picRankings.Print "A centerfielder needs a score of 2300 to be deserving of the Hall of Fame"
        For n = 1 To size
            picRankings.Print Centerfielders(n); Tab(25); CenterScores(n)
        Next n
    End If
End Sub

