VERSION 5.00
Begin VB.Form MNWild 
   BackColor       =   &H80000014&
   Caption         =   "2002/03 MN Wild Regular Season Stats"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9735
   FillColor       =   &H000000FF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton GoalieBox 
      BackColor       =   &H80000013&
      Caption         =   "Goalies (Save Percentage)"
      Height          =   1695
      Left            =   120
      Picture         =   "MNWILD~1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton QuitBox 
      BackColor       =   &H80000013&
      Caption         =   "Quit"
      Height          =   1575
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   2295
   End
   Begin VB.PictureBox PictureBox 
      Height          =   1215
      Left            =   6000
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton PlusMinusBox 
      BackColor       =   &H80000013&
      Caption         =   " Find Out How Valuable You Are"
      Height          =   1815
      Left            =   120
      Picture         =   "MNWILD~1.frx":1483
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton GoalBox 
      BackColor       =   &H80000013&
      Caption         =   "Goals/Other Stats"
      Height          =   1695
      Left            =   120
      Picture         =   "MNWILD~1.frx":2C6A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox StatsBox 
      BackColor       =   &H8000000D&
      Height          =   6015
      Left            =   3840
      ScaleHeight     =   5955
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "MNWild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                             ' Project1 (MNWildStats.vbp)
Dim pass As Single                          ' MNWild (MN Wild Statistics.frm)
Dim I As Single                             ' Jeff Gray
Dim Player(1 To 29) As String               ' 3/23/06
Dim Goals(1 To 29) As Single                ' The last time the Wild were actually good
Dim Penalty_Minutes(1 To 29) As Single
Dim Games(1 To 29) As Single
Dim temp_player As String
Dim temp_goals As Single
Dim temp_penaltyminutes As Single
Dim temp_games As Single
Dim PM As Integer
Dim Goalie(1 To 3) As String
Dim ShotsAgainst(1 To 3) As Integer
Dim Saves(1 To 3) As Integer
Dim SavePercentage(1 To 3) As Double
Dim temp_savepercentage As Double
Dim temp_goalie As String
Dim GP(1 To 3) As Integer



Private Sub GoalBox_Click()                ' Shows player statistics, sorted by most to fewest goals
StatsBox.Cls
PictureBox.Cls
Open App.Path & "\PlayerStats.txt" For Input As #1
For I = 1 To 29
    Input #1, Player(I), Goals(I), Penalty_Minutes(I), Games(I)
Next I
Close #1
StatsBox.Print "Player "; Tab(25); "Goals"; Tab(40); "Penalty Minutes"; Tab(60); "Games Played"
For pass = 1 To 29 - 1
    For I = 1 To 29 - pass
        If Goals(I) < Goals(I + 1) Then
            'swap Player and Goals and Penalty_Minutes and Plus_Minus
            temp_player = Player(I)
            Player(I) = Player(I + 1)
            Player(I + 1) = temp_player
            temp_goals = Goals(I)
            Goals(I) = Goals(I + 1)
            Goals(I + 1) = temp_goals
            temp_penaltyminutes = Penalty_Minutes(I)
            Penalty_Minutes(I) = Penalty_Minutes(I + 1)
            Penalty_Minutes(I + 1) = temp_penaltyminutes
            temp_games = Games(I)
            Games(I) = Games(I + 1)
            Games(I + 1) = temp_games
        End If
    Next I
Next pass
For I = 1 To 29
    StatsBox.Print Player(I); Tab(25); Goals(I); Tab(40); Penalty_Minutes(I); Tab(60); Games(I)
Next I
PictureBox.Picture = LoadPicture(App.Path & "\Gaborik.jpg")
End Sub

Private Sub GoalieBox_Click()       ' Shows goalie percentages in descending order, as well as games played
StatsBox.Cls
PictureBox.Cls
Open App.Path & "\GoalieStats.txt" For Input As #1
StatsBox.Print "Goalie ", "Shots Against", Tab(35), "Games Played"
For I = 1 To 3
    Input #1, Goalie(I), ShotsAgainst(I), Saves(I), GP(I)
Next I
Close #1
For I = 1 To 3
    SavePercentage(I) = Saves(I) / ShotsAgainst(I)
Next I
For pass = 1 To 3 - 1
    For I = 1 To 3 - pass
        If SavePercentage(I) < SavePercentage(I + 1) Then
        'swap SavePercentage and Goalie and GP
            temp_savepercentage = SavePercentage(I)
            SavePercentage(I) = SavePercentage(I + 1)
            SavePercentage(I + 1) = temp_savepercentage
            temp_goalie = Goalie(I)
            Goalie(I) = Goalie(I + 1)
            Goalie(I + 1) = temp_goalie
            temp_games = GP(I)
            GP(I) = GP(I + 1)
            GP(I + 1) = temp_games
        End If
    Next I
Next pass
For I = 1 To 3
    StatsBox.Print Goalie(I); Tab(20); Left(SavePercentage(I), 5); Tab(45); GP(I)
Next I
PictureBox.Picture = LoadPicture(App.Path & "\Roloson.jpg")
End Sub



Private Sub PlusMinusBox_Click()        ' Allows user to see what a good and what a bad plus/minus is
StatsBox.Cls
PictureBox.Cls
PictureBox.Picture = LoadPicture(App.Path & "\Wild Logo.jpg")
PM = (InputBox("What is your Plus/Minus rating this/last season?", "Information"))
If PM >= 50 Then
    MsgBox "You could be the next Wayne Gretzky", , "response"
    ElseIf 30 <= PM Then
    MsgBox "I have to hand it to you, you put points on the board", , "response"
    ElseIf 10 <= PM Then
    MsgBox "You either have talented line mates or you have some skill", , "response"
    ElseIf -10 <= PM Then
    MsgBox "You can hold your own", , "response"
    ElseIf -25 <= PM Then
    MsgBox "You better pick it up or you may be cut", , "response"
    Else
    MsgBox "Thats horrible, you must be in the wrong league", , "response"
End If
End Sub

Private Sub QuitBox_Click()
    End
End Sub
