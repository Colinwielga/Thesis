VERSION 5.00
Begin VB.Form frmDefense 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19995
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   19995
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BackColor       =   &H000000C0&
      Height          =   6015
      Left            =   240
      ScaleHeight     =   5955
      ScaleWidth      =   19635
      TabIndex        =   13
      Top             =   1440
      Width           =   19695
      Begin VB.PictureBox picResults2 
         BackColor       =   &H00FFFF00&
         Height          =   3495
         Left            =   360
         ScaleHeight     =   3435
         ScaleWidth      =   3195
         TabIndex        =   14
         Top             =   2400
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back to Main Page"
      Height          =   495
      Left            =   8640
      TabIndex        =   12
      Top             =   7800
      Width           =   2535
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "Learn More About the Players"
      Height          =   615
      Left            =   840
      TabIndex        =   11
      Top             =   7800
      Width           =   3015
   End
   Begin VB.CommandButton cmdPenaltyMins 
      Caption         =   "Organize by Penalty Minutes"
      Height          =   855
      Left            =   15960
      TabIndex        =   10
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdPlusMinus 
      Caption         =   "Organize by Plus/Minus"
      Height          =   855
      Left            =   14400
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdPlayingTime 
      Caption         =   "Organize by Playing Time"
      Height          =   855
      Left            =   12840
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdShootingPercentage 
      Caption         =   "Organize by Shooting Percentage"
      Height          =   855
      Left            =   11280
      TabIndex        =   7
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdShots 
      Caption         =   "Organize by Shots"
      Height          =   855
      Left            =   9720
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdTeams 
      Caption         =   "Organize by Teams"
      Height          =   855
      Left            =   8160
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdPoints 
      Caption         =   "Organize by Points"
      Height          =   855
      Left            =   6600
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdAssists 
      Caption         =   "Organize by Assists"
      Height          =   855
      Left            =   5040
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdGoals 
      Caption         =   "Organize by Goals"
      Height          =   855
      Left            =   3480
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdAlaphabet 
      Caption         =   "View Players Alaphabetically"
      Height          =   855
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View Top 10 Defense"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Please Type Name of Player Exactly as it Appears Above"
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   7920
      Width           =   2175
   End
End
Attribute VB_Name = "frmDefense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Hockey Statistics
'Form Name: frmhockeystatistics
'Autor: Weston Lake
'Date Written: October 19, 2009
'Objective: To see who are the best players in the NHL so far this season based on their stats
Option Explicit
Dim Player(1 To 10) As String
Dim Teams(1 To 10) As String
Dim Position(1 To 10) As String
Dim GamesPlayed(1 To 10) As Integer, Goals(1 To 10) As Integer, Assists(1 To 10) As Integer, Points(1 To 10) As Integer, PlusMinus(1 To 10) As Integer, PenaltyMins(1 To 10) As Integer, PowerPlay(1 To 10) As Integer, Shots(1 To 10) As Integer, ShootingPercentage(1 To 10) As Single, TimeOnIce(1 To 10) As Single
Dim I As Integer
Dim Pass As Integer
Dim Temp As String, Temp2 As String, Temp3 As String, Temp4 As String, Temp5 As String, Temp6 As String
Dim Temp7 As String, Temp8 As String, Temp9 As String, Temp10 As String, Temp11 As String, Temp12 As String, Temp13 As String





Private Sub cmdAlaphabet_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If Player(I) > Player(I + 1) Then
        Temp = Goals(I)
        Goals(I) = Goals(I + 1)
        Goals(I + 1) = Temp
        Temp2 = Assists(I)
        Assists(I) = Assists(I + 1)
        Assists(I + 1) = Temp2
        Temp3 = GamesPlayed(I)
        GamesPlayed(I) = GamesPlayed(1 + I)
        GamesPlayed(I + 1) = Temp3
        Temp4 = Points(I)
        Points(I) = Points(I + 1)
        Points(I + 1) = Temp4
        Temp5 = PlusMinus(I)
        PlusMinus(I) = PlusMinus(I + 1)
        PlusMinus(I + 1) = Temp5
        Temp6 = PenaltyMins(I)
        PenaltyMins(I) = PenaltyMins(I + 1)
        PenaltyMins(I + 1) = Temp6
        Temp7 = PowerPlay(I)
        PowerPlay(I) = PowerPlay(I + 1)
        PowerPlay(I + 1) = Temp7
        Temp8 = Shots(I)
        Shots(I) = Shots(I + 1)
        Shots(I + 1) = Temp8
        Temp9 = ShootingPercentage(I)
        ShootingPercentage(I) = ShootingPercentage(I + 1)
        ShootingPercentage(I + 1) = Temp9
        Temp10 = TimeOnIce(I)
        TimeOnIce(I) = TimeOnIce(I + 1)
        TimeOnIce(I + 1) = Temp10
        Temp11 = Player(I)
        Player(I) = Player(I + 1)
        Player(I + 1) = Temp11
        Temp12 = Teams(I)
        Teams(I) = Teams(I + 1)
        Teams(I + 1) = Temp12
        Temp13 = Position(I)
        Position(I) = Position(I + 1)
        Position(I + 1) = Temp13
        End If
        Next I
        Next Pass
        
        picResults.Cls
        picResults.Print "Player", Tab(20); "Team", Tab(44); "Position", Tab(60); "Games Played", "Goals", "Assists", "Points", "Plus/Minus", "Penalty Minutes", "Power Play Goals", "Shots", "Shooting Percentage", "Time On Ice Per Game"
        picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"
        
        For I = 1 To 10
           picResults.Print Player(I); Tab(20); Teams(I); Tab; Position(I); Tab; Tab; GamesPlayed(I); Tab; Goals(I); Tab; Assists(I); Tab; Points(I); Tab; PlusMinus(I); Tab; PenaltyMins(I); Tab; Tab; PowerPlay(I); Tab; Tab; Shots(I); Tab; Tab; ShootingPercentage(I); Tab; TimeOnIce(I)
        Next I
End Sub

Private Sub cmdAssists_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If Assists(I) < Assists(I + 1) Then
        Temp = Goals(I)
        Goals(I) = Goals(I + 1)
        Goals(I + 1) = Temp
        Temp2 = Assists(I)
        Assists(I) = Assists(I + 1)
        Assists(I + 1) = Temp2
        Temp3 = GamesPlayed(I)
        GamesPlayed(I) = GamesPlayed(1 + I)
        GamesPlayed(I + 1) = Temp3
        Temp4 = Points(I)
        Points(I) = Points(I + 1)
        Points(I + 1) = Temp4
        Temp5 = PlusMinus(I)
        PlusMinus(I) = PlusMinus(I + 1)
        PlusMinus(I + 1) = Temp5
        Temp6 = PenaltyMins(I)
        PenaltyMins(I) = PenaltyMins(I + 1)
        PenaltyMins(I + 1) = Temp6
        Temp7 = PowerPlay(I)
        PowerPlay(I) = PowerPlay(I + 1)
        PowerPlay(I + 1) = Temp7
        Temp8 = Shots(I)
        Shots(I) = Shots(I + 1)
        Shots(I + 1) = Temp8
        Temp9 = ShootingPercentage(I)
        ShootingPercentage(I) = ShootingPercentage(I + 1)
        ShootingPercentage(I + 1) = Temp9
        Temp10 = TimeOnIce(I)
        TimeOnIce(I) = TimeOnIce(I + 1)
        TimeOnIce(I + 1) = Temp10
        Temp11 = Player(I)
        Player(I) = Player(I + 1)
        Player(I + 1) = Temp11
        Temp12 = Teams(I)
        Teams(I) = Teams(I + 1)
        Teams(I + 1) = Temp12
        Temp13 = Position(I)
        Position(I) = Position(I + 1)
        Position(I + 1) = Temp13
        End If
        Next I
        Next Pass
        
        picResults.Cls
        picResults.Print "Player", Tab(20); "Team", Tab(44); "Position", Tab(60); "Games Played", "Goals", "Assists", "Points", "Plus/Minus", "Penalty Minutes", "Power Play Goals", "Shots", "Shooting Percentage", "Time On Ice Per Game"
        picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"
        
        For I = 1 To 10
           picResults.Print Player(I); Tab(20); Teams(I); Tab; Position(I); Tab; Tab; GamesPlayed(I); Tab; Goals(I); Tab; Assists(I); Tab; Points(I); Tab; PlusMinus(I); Tab; PenaltyMins(I); Tab; Tab; PowerPlay(I); Tab; Tab; Shots(I); Tab; Tab; ShootingPercentage(I); Tab; TimeOnIce(I)
        Next I
End Sub

Private Sub cmdBack_Click()
 frmDefense.Hide
    frmHockeyStatistics.Show

End Sub

Private Sub cmdGoals_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If Goals(I) < Goals(I + 1) Then
        Temp = Goals(I)
        Goals(I) = Goals(I + 1)
        Goals(I + 1) = Temp
        Temp2 = Assists(I)
        Assists(I) = Assists(I + 1)
        Assists(I + 1) = Temp2
        Temp3 = GamesPlayed(I)
        GamesPlayed(I) = GamesPlayed(1 + I)
        GamesPlayed(I + 1) = Temp3
        Temp4 = Points(I)
        Points(I) = Points(I + 1)
        Points(I + 1) = Temp4
        Temp5 = PlusMinus(I)
        PlusMinus(I) = PlusMinus(I + 1)
        PlusMinus(I + 1) = Temp5
        Temp6 = PenaltyMins(I)
        PenaltyMins(I) = PenaltyMins(I + 1)
        PenaltyMins(I + 1) = Temp6
        Temp7 = PowerPlay(I)
        PowerPlay(I) = PowerPlay(I + 1)
        PowerPlay(I + 1) = Temp7
        Temp8 = Shots(I)
        Shots(I) = Shots(I + 1)
        Shots(I + 1) = Temp8
        Temp9 = ShootingPercentage(I)
        ShootingPercentage(I) = ShootingPercentage(I + 1)
        ShootingPercentage(I + 1) = Temp9
        Temp10 = TimeOnIce(I)
        TimeOnIce(I) = TimeOnIce(I + 1)
        TimeOnIce(I + 1) = Temp10
        Temp11 = Player(I)
        Player(I) = Player(I + 1)
        Player(I + 1) = Temp11
        Temp12 = Teams(I)
        Teams(I) = Teams(I + 1)
        Teams(I + 1) = Temp12
        Temp13 = Position(I)
        Position(I) = Position(I + 1)
        Position(I + 1) = Temp13
        End If
        Next I
        Next Pass
        
        picResults.Cls
        picResults.Print "Player", Tab(20); "Team", Tab(44); "Position", Tab(60); "Games Played", "Goals", "Assists", "Points", "Plus/Minus", "Penalty Minutes", "Power Play Goals", "Shots", "Shooting Percentage", "Time On Ice Per Game"
        picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"
        
        For I = 1 To 10
           picResults.Print Player(I); Tab(20); Teams(I); Tab; Position(I); Tab; Tab; GamesPlayed(I); Tab; Goals(I); Tab; Assists(I); Tab; Points(I); Tab; PlusMinus(I); Tab; PenaltyMins(I); Tab; Tab; PowerPlay(I); Tab; Tab; Shots(I); Tab; Tab; ShootingPercentage(I); Tab; TimeOnIce(I)
        Next I
End Sub

Private Sub cmdMore_Click()
Dim J As String
   Dim I As Integer
   
   J = InputBox("Enter a players name you want to see more information about", "Enter Player")
  For I = 1 To 10
   If J = "Michael Del Zotto" Then
    picResults.Cls
    picResults.Print "NUMBER: 4"
    picResults.Print "HEIGHT: 6' 1"""
    picResults.Print "WEIGHT: 195"
    picResults.Print "SHOOTS: Left"
    picResults.Print "BIRTHDATE: Jun 24, 1990  (AGE 19)"
    picResults.Print "BIRTHPLACE: Stouffville, ON, Canada"
    picResults.Print "DRAFTED: NYR / 2008 NHL Entry Draft"
    picResults.Print "ROUND: 1st  (20th overall)"
    picResults2.Picture = LoadPicture(App.Path & "\delzotto.jpg")
ElseIf J = "Matt Carle" Then
    picResults.Cls
    picResults.Print "NUMBER: 25"
    picResults.Print "HEIGHT: 6' 0"""
    picResults.Print "WEIGHT: 205"
    picResults.Print "SHOOTS: Left"
    picResults.Print "BIRTHDATE: Sep 25, 1984  (AGE 25)"
    picResults.Print "BIRTHPLACE: Anchorage, AK, United States"
    picResults.Print "DRAFTED: SJS / 2003 NHL Entry Draft"
    picResults.Print "ROUND: 2nd  (47th overall)"
    picResults2.Picture = LoadPicture(App.Path & "\carle.jpg")
ElseIf J = "Dion Phaneuf" Then
    picResults.Cls
    picResults.Print "NUMBER: 3"
    picResults.Print "HEIGHT: 6' 3"""
    picResults.Print "WEIGHT: 214"
    picResults.Print "SHOOTS: Left"
    picResults.Print "BIRTHDATE: Apr 10, 1985  (AGE 24)"
    picResults.Print "BIRTHPLACE: Edmonton, AB, Canada"
    picResults.Print "DRAFTED: CGY / 2003 NHL Entry Draft"
    picResults.Print "ROUND: 1st  (9th overall)"
    picResults2.Picture = LoadPicture(App.Path & "\phaneuf.jpg")
ElseIf J = "Sergei Gonchar" Then
    picResults.Cls
    picResults.Print "NUMBER: 55"
    picResults.Print "HEIGHT: 6' 2"""
    picResults.Print "WEIGHT: 211"
    picResults.Print "SHOOTS: Left"
    picResults.Print "BIRTHDATE: Apr 13, 1974  (AGE 35)"
    picResults.Print "BIRTHPLACE: Chelyabinsk, Russia"
    picResults.Print "DRAFTED: WSH / 1992 NHL Entry Draft"
    picResults.Print "ROUND: 1st  (14th overall)"
    picResults2.Picture = LoadPicture(App.Path & "\gonchar.jpg")
ElseIf J = "Brain Campbell" Then
    picResults.Cls
    picResults.Print "NUMBER: 51"
    picResults.Print "HEIGHT: 6' 0"""
    picResults.Print "WEIGHT: 189"
    picResults.Print "SHOOTS: Left"
    picResults.Print "BIRTHDATE: May 23, 1979  (AGE 30)"
    picResults.Print "BIRTHPLACE: Strathroy, ON, Canada"
    picResults.Print "DRAFTED: BUF / 1997 NHL Entry Draft"
    picResults.Print "ROUND: 6th  (156th overall)"
    picResults2.Picture = LoadPicture(App.Path & "\campbell.jpg")
ElseIf J = "Brent Seabrook" Then
    picResults.Cls
    picResults.Print "NUMBER: 7"
    picResults.Print "HEIGHT: 6' 3"""
    picResults.Print "WEIGHT: 218"
    picResults.Print "SHOOTS: Right"
    picResults.Print "BIRTHDATE: Apr 20, 1985  (AGE 24)"
    picResults.Print "BIRTHPLACE: Richmond, BC, Canada"
    picResults.Print "DRAFTED: CHI / 2003 NHL Entry Draft"
    picResults.Print "ROUND: 1st  (14th overall)"
    picResults2.Picture = LoadPicture(App.Path & "\seabrook.jpg")
ElseIf J = "Denis Grebeshkov" Then
    picResults.Cls
    picResults.Print "NUMBER: 37"
    picResults.Print "HEIGHT: 6' 0"""
    picResults.Print "WEIGHT: 209"
    picResults.Print "SHOOTS: Left"
    picResults.Print "BIRTHDATE: Oct 11, 1983  (AGE 26)"
    picResults.Print "BIRTHPLACE: Yaroslavl, Russia"
    picResults.Print "DRAFTED: LAK / 2002 NHL Entry Draft"
    picResults.Print "ROUND: 1st  (18th overall)"
    picResults2.Picture = LoadPicture(App.Path & "\grebeshkov.jpg")
ElseIf J = "Drew Doughty" Then
    picResults.Cls
    picResults.Print "NUMBER: 8"
    picResults.Print "HEIGHT: 6' 0"""
    picResults.Print "WEIGHT: 211"
    picResults.Print "SHOOTS: Right"
    picResults.Print "BIRTHDATE: Dec 8, 1989  (AGE 19)"
    picResults.Print "BIRTHPLACE: London, ON, Canada"
    picResults.Print "DRAFTED: LAK / 2008 NHL Entry Draft"
    picResults.Print "ROUND: 1st  (2nd overall)"
    picResults2.Picture = LoadPicture(App.Path & "\doughty.jpg")
ElseIf J = "Christian Ehrhoff" Then
    picResults.Cls
    picResults.Print "NUMBER: 5"
    picResults.Print "HEIGHT: 6' 2"""
    picResults.Print "WEIGHT: 200"
    picResults.Print "SHOOTS: Left"
    picResults.Print "BIRTHDATE: Jul 6, 1982  (AGE 27)"
    picResults.Print "BIRTHPLACE: Moers, Germany"
    picResults.Print "DRAFTED: SJS / 2001 NHL Entry Draft"
    picResults.Print "ROUND: 4th  (106th overall)"
    picResults2.Picture = LoadPicture(App.Path & "\ehrhoff.jpg")
ElseIf J = "Kyle Quincey" Then
    picResults.Cls
    picResults.Print "NUMBER: 27"
    picResults.Print "HEIGHT: 6' 2"""
    picResults.Print "WEIGHT: 207"
    picResults.Print "SHOOTS: Left"
    picResults.Print "BIRTHDATE: Aug 12, 1985  (AGE 24)"
    picResults.Print "BIRTHPLACE: Kitchener, ON, Canada"
    picResults.Print "DRAFTED: DET / 2003 NHL Entry Draft"
    picResults.Print "ROUND: 4th  (132nd overall)"
    picResults2.Picture = LoadPicture(App.Path & "\quincey.jpg")
  Else
     MsgBox "Please Enter a Correct Name", , "Error"
    End If
   Next I
End Sub

Private Sub cmdPenaltyMins_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If PenaltyMins(I) < PenaltyMins(I + 1) Then
        Temp = Goals(I)
        Goals(I) = Goals(I + 1)
        Goals(I + 1) = Temp
        Temp2 = Assists(I)
        Assists(I) = Assists(I + 1)
        Assists(I + 1) = Temp2
        Temp3 = GamesPlayed(I)
        GamesPlayed(I) = GamesPlayed(1 + I)
        GamesPlayed(I + 1) = Temp3
        Temp4 = Points(I)
        Points(I) = Points(I + 1)
        Points(I + 1) = Temp4
        Temp5 = PlusMinus(I)
        PlusMinus(I) = PlusMinus(I + 1)
        PlusMinus(I + 1) = Temp5
        Temp6 = PenaltyMins(I)
        PenaltyMins(I) = PenaltyMins(I + 1)
        PenaltyMins(I + 1) = Temp6
        Temp7 = PowerPlay(I)
        PowerPlay(I) = PowerPlay(I + 1)
        PowerPlay(I + 1) = Temp7
        Temp8 = Shots(I)
        Shots(I) = Shots(I + 1)
        Shots(I + 1) = Temp8
        Temp9 = ShootingPercentage(I)
        ShootingPercentage(I) = ShootingPercentage(I + 1)
        ShootingPercentage(I + 1) = Temp9
        Temp10 = TimeOnIce(I)
        TimeOnIce(I) = TimeOnIce(I + 1)
        TimeOnIce(I + 1) = Temp10
        Temp11 = Player(I)
        Player(I) = Player(I + 1)
        Player(I + 1) = Temp11
        Temp12 = Teams(I)
        Teams(I) = Teams(I + 1)
        Teams(I + 1) = Temp12
        Temp13 = Position(I)
        Position(I) = Position(I + 1)
        Position(I + 1) = Temp13
        End If
        Next I
        Next Pass
        
        picResults.Cls
        picResults.Print "Player", Tab(20); "Team", Tab(44); "Position", Tab(60); "Games Played", "Goals", "Assists", "Points", "Plus/Minus", "Penalty Minutes", "Power Play Goals", "Shots", "Shooting Percentage", "Time On Ice Per Game"
        picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"
        
        For I = 1 To 10
           picResults.Print Player(I); Tab(20); Teams(I); Tab; Position(I); Tab; Tab; GamesPlayed(I); Tab; Goals(I); Tab; Assists(I); Tab; Points(I); Tab; PlusMinus(I); Tab; PenaltyMins(I); Tab; Tab; PowerPlay(I); Tab; Tab; Shots(I); Tab; Tab; ShootingPercentage(I); Tab; TimeOnIce(I)
        Next I
End Sub

Private Sub cmdPlayingTime_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If TimeOnIce(I) < TimeOnIce(I + 1) Then
        Temp = Goals(I)
        Goals(I) = Goals(I + 1)
        Goals(I + 1) = Temp
        Temp2 = Assists(I)
        Assists(I) = Assists(I + 1)
        Assists(I + 1) = Temp2
        Temp3 = GamesPlayed(I)
        GamesPlayed(I) = GamesPlayed(1 + I)
        GamesPlayed(I + 1) = Temp3
        Temp4 = Points(I)
        Points(I) = Points(I + 1)
        Points(I + 1) = Temp4
        Temp5 = PlusMinus(I)
        PlusMinus(I) = PlusMinus(I + 1)
        PlusMinus(I + 1) = Temp5
        Temp6 = PenaltyMins(I)
        PenaltyMins(I) = PenaltyMins(I + 1)
        PenaltyMins(I + 1) = Temp6
        Temp7 = PowerPlay(I)
        PowerPlay(I) = PowerPlay(I + 1)
        PowerPlay(I + 1) = Temp7
        Temp8 = Shots(I)
        Shots(I) = Shots(I + 1)
        Shots(I + 1) = Temp8
        Temp9 = ShootingPercentage(I)
        ShootingPercentage(I) = ShootingPercentage(I + 1)
        ShootingPercentage(I + 1) = Temp9
        Temp10 = TimeOnIce(I)
        TimeOnIce(I) = TimeOnIce(I + 1)
        TimeOnIce(I + 1) = Temp10
        Temp11 = Player(I)
        Player(I) = Player(I + 1)
        Player(I + 1) = Temp11
        Temp12 = Teams(I)
        Teams(I) = Teams(I + 1)
        Teams(I + 1) = Temp12
        Temp13 = Position(I)
        Position(I) = Position(I + 1)
        Position(I + 1) = Temp13
        End If
        Next I
        Next Pass
        
        picResults.Cls
        picResults.Print "Player", Tab(20); "Team", Tab(44); "Position", Tab(60); "Games Played", "Goals", "Assists", "Points", "Plus/Minus", "Penalty Minutes", "Power Play Goals", "Shots", "Shooting Percentage", "Time On Ice Per Game"
        picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"
        
        For I = 1 To 10
           picResults.Print Player(I); Tab(20); Teams(I); Tab; Position(I); Tab; Tab; GamesPlayed(I); Tab; Goals(I); Tab; Assists(I); Tab; Points(I); Tab; PlusMinus(I); Tab; PenaltyMins(I); Tab; Tab; PowerPlay(I); Tab; Tab; Shots(I); Tab; Tab; ShootingPercentage(I); Tab; TimeOnIce(I)
        Next I
End Sub

Private Sub cmdPlusMinus_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If PlusMinus(I) < PlusMinus(I + 1) Then
        Temp = Goals(I)
        Goals(I) = Goals(I + 1)
        Goals(I + 1) = Temp
        Temp2 = Assists(I)
        Assists(I) = Assists(I + 1)
        Assists(I + 1) = Temp2
        Temp3 = GamesPlayed(I)
        GamesPlayed(I) = GamesPlayed(1 + I)
        GamesPlayed(I + 1) = Temp3
        Temp4 = Points(I)
        Points(I) = Points(I + 1)
        Points(I + 1) = Temp4
        Temp5 = PlusMinus(I)
        PlusMinus(I) = PlusMinus(I + 1)
        PlusMinus(I + 1) = Temp5
        Temp6 = PenaltyMins(I)
        PenaltyMins(I) = PenaltyMins(I + 1)
        PenaltyMins(I + 1) = Temp6
        Temp7 = PowerPlay(I)
        PowerPlay(I) = PowerPlay(I + 1)
        PowerPlay(I + 1) = Temp7
        Temp8 = Shots(I)
        Shots(I) = Shots(I + 1)
        Shots(I + 1) = Temp8
        Temp9 = ShootingPercentage(I)
        ShootingPercentage(I) = ShootingPercentage(I + 1)
        ShootingPercentage(I + 1) = Temp9
        Temp10 = TimeOnIce(I)
        TimeOnIce(I) = TimeOnIce(I + 1)
        TimeOnIce(I + 1) = Temp10
        Temp11 = Player(I)
        Player(I) = Player(I + 1)
        Player(I + 1) = Temp11
        Temp12 = Teams(I)
        Teams(I) = Teams(I + 1)
        Teams(I + 1) = Temp12
        Temp13 = Position(I)
        Position(I) = Position(I + 1)
        Position(I + 1) = Temp13
        End If
        Next I
        Next Pass
        
        picResults.Cls
        picResults.Print "Player", Tab(20); "Team", Tab(44); "Position", Tab(60); "Games Played", "Goals", "Assists", "Points", "Plus/Minus", "Penalty Minutes", "Power Play Goals", "Shots", "Shooting Percentage", "Time On Ice Per Game"
        picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"
        
        For I = 1 To 10
           picResults.Print Player(I); Tab(20); Teams(I); Tab; Position(I); Tab; Tab; GamesPlayed(I); Tab; Goals(I); Tab; Assists(I); Tab; Points(I); Tab; PlusMinus(I); Tab; PenaltyMins(I); Tab; Tab; PowerPlay(I); Tab; Tab; Shots(I); Tab; Tab; ShootingPercentage(I); Tab; TimeOnIce(I)
        Next I
End Sub

Private Sub cmdPoints_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If Points(I) < Points(I + 1) Then
        Temp = Goals(I)
        Goals(I) = Goals(I + 1)
        Goals(I + 1) = Temp
        Temp2 = Assists(I)
        Assists(I) = Assists(I + 1)
        Assists(I + 1) = Temp2
        Temp3 = GamesPlayed(I)
        GamesPlayed(I) = GamesPlayed(1 + I)
        GamesPlayed(I + 1) = Temp3
        Temp4 = Points(I)
        Points(I) = Points(I + 1)
        Points(I + 1) = Temp4
        Temp5 = PlusMinus(I)
        PlusMinus(I) = PlusMinus(I + 1)
        PlusMinus(I + 1) = Temp5
        Temp6 = PenaltyMins(I)
        PenaltyMins(I) = PenaltyMins(I + 1)
        PenaltyMins(I + 1) = Temp6
        Temp7 = PowerPlay(I)
        PowerPlay(I) = PowerPlay(I + 1)
        PowerPlay(I + 1) = Temp7
        Temp8 = Shots(I)
        Shots(I) = Shots(I + 1)
        Shots(I + 1) = Temp8
        Temp9 = ShootingPercentage(I)
        ShootingPercentage(I) = ShootingPercentage(I + 1)
        ShootingPercentage(I + 1) = Temp9
        Temp10 = TimeOnIce(I)
        TimeOnIce(I) = TimeOnIce(I + 1)
        TimeOnIce(I + 1) = Temp10
        Temp11 = Player(I)
        Player(I) = Player(I + 1)
        Player(I + 1) = Temp11
        Temp12 = Teams(I)
        Teams(I) = Teams(I + 1)
        Teams(I + 1) = Temp12
        Temp13 = Position(I)
        Position(I) = Position(I + 1)
        Position(I + 1) = Temp13
        End If
        Next I
        Next Pass
        
        picResults.Cls
        picResults.Print "Player", Tab(20); "Team", Tab(44); "Position", Tab(60); "Games Played", "Goals", "Assists", "Points", "Plus/Minus", "Penalty Minutes", "Power Play Goals", "Shots", "Shooting Percentage", "Time On Ice Per Game"
        picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"
        
        For I = 1 To 10
           picResults.Print Player(I); Tab(20); Teams(I); Tab; Position(I); Tab; Tab; GamesPlayed(I); Tab; Goals(I); Tab; Assists(I); Tab; Points(I); Tab; PlusMinus(I); Tab; PenaltyMins(I); Tab; Tab; PowerPlay(I); Tab; Tab; Shots(I); Tab; Tab; ShootingPercentage(I); Tab; TimeOnIce(I)
        Next I
End Sub

Private Sub cmdShootingPercentage_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If ShootingPercentage(I) < ShootingPercentage(I + 1) Then
        Temp = Goals(I)
        Goals(I) = Goals(I + 1)
        Goals(I + 1) = Temp
        Temp2 = Assists(I)
        Assists(I) = Assists(I + 1)
        Assists(I + 1) = Temp2
        Temp3 = GamesPlayed(I)
        GamesPlayed(I) = GamesPlayed(1 + I)
        GamesPlayed(I + 1) = Temp3
        Temp4 = Points(I)
        Points(I) = Points(I + 1)
        Points(I + 1) = Temp4
        Temp5 = PlusMinus(I)
        PlusMinus(I) = PlusMinus(I + 1)
        PlusMinus(I + 1) = Temp5
        Temp6 = PenaltyMins(I)
        PenaltyMins(I) = PenaltyMins(I + 1)
        PenaltyMins(I + 1) = Temp6
        Temp7 = PowerPlay(I)
        PowerPlay(I) = PowerPlay(I + 1)
        PowerPlay(I + 1) = Temp7
        Temp8 = Shots(I)
        Shots(I) = Shots(I + 1)
        Shots(I + 1) = Temp8
        Temp9 = ShootingPercentage(I)
        ShootingPercentage(I) = ShootingPercentage(I + 1)
        ShootingPercentage(I + 1) = Temp9
        Temp10 = TimeOnIce(I)
        TimeOnIce(I) = TimeOnIce(I + 1)
        TimeOnIce(I + 1) = Temp10
        Temp11 = Player(I)
        Player(I) = Player(I + 1)
        Player(I + 1) = Temp11
        Temp12 = Teams(I)
        Teams(I) = Teams(I + 1)
        Teams(I + 1) = Temp12
        Temp13 = Position(I)
        Position(I) = Position(I + 1)
        Position(I + 1) = Temp13
        End If
        Next I
        Next Pass
        
        picResults.Cls
        picResults.Print "Player", Tab(20); "Team", Tab(44); "Position", Tab(60); "Games Played", "Goals", "Assists", "Points", "Plus/Minus", "Penalty Minutes", "Power Play Goals", "Shots", "Shooting Percentage", "Time On Ice Per Game"
        picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"
        
        For I = 1 To 10
           picResults.Print Player(I); Tab(20); Teams(I); Tab; Position(I); Tab; Tab; GamesPlayed(I); Tab; Goals(I); Tab; Assists(I); Tab; Points(I); Tab; PlusMinus(I); Tab; PenaltyMins(I); Tab; Tab; PowerPlay(I); Tab; Tab; Shots(I); Tab; Tab; ShootingPercentage(I); Tab; TimeOnIce(I)
        Next I
End Sub

Private Sub cmdShots_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If Shots(I) < Shots(I + 1) Then
        Temp = Goals(I)
        Goals(I) = Goals(I + 1)
        Goals(I + 1) = Temp
        Temp2 = Assists(I)
        Assists(I) = Assists(I + 1)
        Assists(I + 1) = Temp2
        Temp3 = GamesPlayed(I)
        GamesPlayed(I) = GamesPlayed(1 + I)
        GamesPlayed(I + 1) = Temp3
        Temp4 = Points(I)
        Points(I) = Points(I + 1)
        Points(I + 1) = Temp4
        Temp5 = PlusMinus(I)
        PlusMinus(I) = PlusMinus(I + 1)
        PlusMinus(I + 1) = Temp5
        Temp6 = PenaltyMins(I)
        PenaltyMins(I) = PenaltyMins(I + 1)
        PenaltyMins(I + 1) = Temp6
        Temp7 = PowerPlay(I)
        PowerPlay(I) = PowerPlay(I + 1)
        PowerPlay(I + 1) = Temp7
        Temp8 = Shots(I)
        Shots(I) = Shots(I + 1)
        Shots(I + 1) = Temp8
        Temp9 = ShootingPercentage(I)
        ShootingPercentage(I) = ShootingPercentage(I + 1)
        ShootingPercentage(I + 1) = Temp9
        Temp10 = TimeOnIce(I)
        TimeOnIce(I) = TimeOnIce(I + 1)
        TimeOnIce(I + 1) = Temp10
        Temp11 = Player(I)
        Player(I) = Player(I + 1)
        Player(I + 1) = Temp11
        Temp12 = Teams(I)
        Teams(I) = Teams(I + 1)
        Teams(I + 1) = Temp12
        Temp13 = Position(I)
        Position(I) = Position(I + 1)
        Position(I + 1) = Temp13
        End If
        Next I
        Next Pass
        
        picResults.Cls
        picResults.Print "Player", Tab(20); "Team", Tab(44); "Position", Tab(60); "Games Played", "Goals", "Assists", "Points", "Plus/Minus", "Penalty Minutes", "Power Play Goals", "Shots", "Shooting Percentage", "Time On Ice Per Game"
        picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"
        
        For I = 1 To 10
           picResults.Print Player(I); Tab(20); Teams(I); Tab; Position(I); Tab; Tab; GamesPlayed(I); Tab; Goals(I); Tab; Assists(I); Tab; Points(I); Tab; PlusMinus(I); Tab; PenaltyMins(I); Tab; Tab; PowerPlay(I); Tab; Tab; Shots(I); Tab; Tab; ShootingPercentage(I); Tab; TimeOnIce(I)
        Next I
End Sub

Private Sub cmdTeams_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If Teams(I) > Teams(I + 1) Then
        Temp = Goals(I)
        Goals(I) = Goals(I + 1)
        Goals(I + 1) = Temp
        Temp2 = Assists(I)
        Assists(I) = Assists(I + 1)
        Assists(I + 1) = Temp2
        Temp3 = GamesPlayed(I)
        GamesPlayed(I) = GamesPlayed(1 + I)
        GamesPlayed(I + 1) = Temp3
        Temp4 = Points(I)
        Points(I) = Points(I + 1)
        Points(I + 1) = Temp4
        Temp5 = PlusMinus(I)
        PlusMinus(I) = PlusMinus(I + 1)
        PlusMinus(I + 1) = Temp5
        Temp6 = PenaltyMins(I)
        PenaltyMins(I) = PenaltyMins(I + 1)
        PenaltyMins(I + 1) = Temp6
        Temp7 = PowerPlay(I)
        PowerPlay(I) = PowerPlay(I + 1)
        PowerPlay(I + 1) = Temp7
        Temp8 = Shots(I)
        Shots(I) = Shots(I + 1)
        Shots(I + 1) = Temp8
        Temp9 = ShootingPercentage(I)
        ShootingPercentage(I) = ShootingPercentage(I + 1)
        ShootingPercentage(I + 1) = Temp9
        Temp10 = TimeOnIce(I)
        TimeOnIce(I) = TimeOnIce(I + 1)
        TimeOnIce(I + 1) = Temp10
        Temp11 = Player(I)
        Player(I) = Player(I + 1)
        Player(I + 1) = Temp11
        Temp12 = Teams(I)
        Teams(I) = Teams(I + 1)
        Teams(I + 1) = Temp12
        Temp13 = Position(I)
        Position(I) = Position(I + 1)
        Position(I + 1) = Temp13
        End If
        Next I
        Next Pass
        
        picResults.Cls
        picResults.Print "Player", Tab(20); "Team", Tab(44); "Position", Tab(60); "Games Played", "Goals", "Assists", "Points", "Plus/Minus", "Penalty Minutes", "Power Play Goals", "Shots", "Shooting Percentage", "Time On Ice Per Game"
        picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"
        
        For I = 1 To 10
           picResults.Print Player(I); Tab(20); Teams(I); Tab; Position(I); Tab; Tab; GamesPlayed(I); Tab; Goals(I); Tab; Assists(I); Tab; Points(I); Tab; PlusMinus(I); Tab; PenaltyMins(I); Tab; Tab; PowerPlay(I); Tab; Tab; Shots(I); Tab; Tab; ShootingPercentage(I); Tab; TimeOnIce(I)
        Next I
End Sub

Private Sub cmdView_Click()
Open App.Path & "\Defense.txt" For Input As #1
picResults.Print "Player", Tab(20); "Team", Tab(44); "Position", Tab(60); "Games Played", "Goals", "Assists", "Points", "Plus/Minus", "Penalty Minutes", "Power Play Goals", "Shots", "Shooting Percentage", "Time On Ice Per Game"
picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"
For I = 1 To 10
    Input #1, Player(I), Teams(I), Position(I), GamesPlayed(I), Goals(I), Assists(I), Points(I), PlusMinus(I), PenaltyMins(I), PowerPlay(I), Shots(I), ShootingPercentage(I), TimeOnIce(I)
 Next I
 For I = 1 To 10
    picResults.Print Player(I); Tab(20); Teams(I); Tab; Position(I); Tab; Tab; GamesPlayed(I); Tab; Goals(I); Tab; Assists(I); Tab; Points(I); Tab; PlusMinus(I); Tab; PenaltyMins(I); Tab; Tab; PowerPlay(I); Tab; Tab; Shots(I); Tab; Tab; ShootingPercentage(I); Tab; TimeOnIce(I)
 Next I

End Sub

