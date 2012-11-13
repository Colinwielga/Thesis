VERSION 5.00
Begin VB.Form frmForwards 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20070
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   20070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back to Main Page"
      Height          =   495
      Left            =   10920
      TabIndex        =   15
      Top             =   7800
      Width           =   4455
   End
   Begin VB.CommandButton cmdMore 
      BackColor       =   &H80000010&
      Caption         =   "Learn More About the Players"
      Height          =   615
      Left            =   600
      TabIndex        =   12
      Top             =   7680
      Width           =   2655
   End
   Begin VB.CommandButton cmdPenaltyMin 
      Caption         =   "Organize by Penalty Minutes"
      Height          =   975
      Left            =   15840
      TabIndex        =   11
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdPlusMinus 
      Caption         =   "Organize by Plus/Minus"
      Height          =   975
      Left            =   14280
      TabIndex        =   10
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdTimeOnIce 
      Caption         =   "Organize by Playing Time"
      Height          =   975
      Left            =   12720
      TabIndex        =   9
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdShootingPercentage 
      Caption         =   "Organize by Shooting Percentage"
      Height          =   975
      Left            =   11160
      TabIndex        =   8
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdShot 
      Caption         =   "Organize by Shots"
      Height          =   975
      Left            =   9600
      TabIndex        =   7
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdTeams 
      Caption         =   "Organize by Teams"
      Height          =   975
      Left            =   8040
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdAssists 
      Caption         =   "Organize by Assists"
      Height          =   975
      Left            =   4920
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdPlayer 
      BackColor       =   &H80000012&
      Caption         =   "View Players Alaphabetically"
      Height          =   975
      Left            =   1800
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdPoints 
      Caption         =   "Organize by Points"
      Height          =   975
      Left            =   6480
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdGoals 
      Caption         =   "Organize by Goals"
      Height          =   975
      Left            =   3360
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdAllForwards 
      Caption         =   "View Top 10 Forwards"
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H000000C0&
      Height          =   6135
      Left            =   0
      ScaleHeight     =   6075
      ScaleWidth      =   19755
      TabIndex        =   0
      Top             =   1440
      Width           =   19815
      Begin VB.PictureBox picResults2 
         BackColor       =   &H00FFFF00&
         Height          =   3255
         Left            =   480
         ScaleHeight     =   3195
         ScaleWidth      =   2835
         TabIndex        =   14
         Top             =   2640
         Width           =   2895
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Please Type Name of Player Exactly as is Appears Above"
      Height          =   495
      Left            =   3360
      TabIndex        =   13
      Top             =   7800
      Width           =   2295
   End
End
Attribute VB_Name = "frmForwards"
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

Private Sub cmdAllForwards_Click()
   
    
Open App.Path & "\Forwards.txt" For Input As #1
picResults.Print "Player", Tab(20); "Team", Tab(44); "Position", Tab(60); "Games Played", "Goals", "Assists", "Points", "Plus/Minus", "Penalty Minutes", "Power Play Goals", "Shots", "Shooting Percentage", "Time On Ice Per Game"
picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"
For I = 1 To 10
    Input #1, Player(I), Teams(I), Position(I), GamesPlayed(I), Goals(I), Assists(I), Points(I), PlusMinus(I), PenaltyMins(I), PowerPlay(I), Shots(I), ShootingPercentage(I), TimeOnIce(I)
 Next I
 For I = 1 To 10
    picResults.Print Player(I); Tab(20); Teams(I); Tab; Position(I); Tab; Tab; GamesPlayed(I); Tab; Goals(I); Tab; Assists(I); Tab; Points(I); Tab; PlusMinus(I); Tab; PenaltyMins(I); Tab; Tab; PowerPlay(I); Tab; Tab; Shots(I); Tab; Tab; ShootingPercentage(I); Tab; TimeOnIce(I)
 Next I
 




End Sub

Private Sub cmdCenter_Click()

End Sub

Private Sub cmdAssists_Click()
Dim Pass As Integer
Dim Temp As String, Temp2 As String, Temp3 As String, Temp4 As String, Temp5 As String, Temp6 As String
Dim Temp7 As String, Temp8 As String, Temp9 As String, Temp10 As String, Temp11 As String, Temp12 As String, Temp13 As String

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
    frmForwards.Hide
    frmHockeyStatistics.Show

End Sub

Private Sub cmdGoals_Click()
Dim Pass As Integer
Dim Temp As String, Temp2 As String, Temp3 As String, Temp4 As String, Temp5 As String, Temp6 As String
Dim Temp7 As String, Temp8 As String, Temp9 As String, Temp10 As String, Temp11 As String, Temp12 As String, Temp13 As String

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
   If J = "Joe Thornton" Then
    picResults.Cls
    picResults.Print "NUMBER: 19"
    picResults.Print "HEIGHT: 6' 4"""
    picResults.Print "WEIGHT: 235"
    picResults.Print "SHOOTS: Left"
    picResults.Print "BIRTHDATE: Jul 2, 1979  (AGE 30)"
    picResults.Print "BIRTHPLACE: London, ON, Canada"
    picResults.Print "DRAFTED: BOS / 1997 NHL Entry Draft"
    picResults.Print "ROUND: 1st  (1st overall)"
    picResults2.Picture = LoadPicture(App.Path & "\Thornton.jpg")
ElseIf J = "Alex Ovechkin" Then
    picResults.Cls
    picResults.Print "NUMBER: 8"
    picResults.Print "HEIGHT: 6' 2"""
    picResults.Print "WEIGHT: 233"
    picResults.Print "Shoots: Right"
    picResults.Print "BIRTHDATE: Sep 17, 1985  (AGE 24)"
    picResults.Print "BIRTHPLACE: Moscow, Russia"
    picResults.Print "DRAFTED: WSH / 2004 NHL Entry Draft"
    picResults.Print "ROUND: 1st  (1st overall)"
    picResults2.Picture = LoadPicture(App.Path & "\ovech.jpg")
 ElseIf J = "Marian Gaborik" Then
    picResults.Cls
    picResults.Print "NUMBER: 10"
    picResults.Print "HEIGHT: 6' 1"""
    picResults.Print "WEIGHT: 200"
    picResults.Print "Shoots: Left"
    picResults.Print "BIRTHDATE: Feb 14, 1982  (AGE 27)"
    picResults.Print "BIRTHPLACE: Trencin, Slovakia"
    picResults.Print "DRAFTED: MIN / 2000 NHL Entry Draft"
    picResults.Print "ROUND: 1st  (3rd overall)"
    picResults2.Picture = LoadPicture(App.Path & "\Gaborik.jpg")
ElseIf J = "Anze Kopitar" Then
    picResults.Cls
    picResults.Print "NUMBER: 11"
    picResults.Print "HEIGHT: 6' 3"""
    picResults.Print "WEIGHT: 222"
    picResults.Print "Shoots: Left"
    picResults.Print "BIRTHDATE: Aug 24, 1987  (AGE 22)"
    picResults.Print "BIRTHPLACE: Jesenice, Slovenia"
    picResults.Print "DRAFTED: LAK / 2005 NHL Entry Draft"
    picResults.Print "ROUND: 1st  (11th overall)"
    picResults2.Picture = LoadPicture(App.Path & "\Kopitar.jpg")
  ElseIf J = "Martin St Louis" Then
    picResults.Cls
    picResults.Print "NUMBER: 26"
    picResults.Print "HEIGHT: 5' 9"""
    picResults.Print "WEIGHT: 177"
    picResults.Print "Shoots: Left"
    picResults.Print "BIRTHDATE: Jun 18, 1975  (AGE 34)"
    picResults.Print "BIRTHPLACE: Laval, QC, Canada"
    picResults2.Picture = LoadPicture(App.Path & "\st louis.jpg")
ElseIf J = "Henrik Sedin" Then
    picResults.Cls
    picResults.Print "NUMBER: 33"
    picResults.Print "HEIGHT: 6' 2"""
    picResults.Print "WEIGHT: 183"
    picResults.Print "Shoots: Left"
    picResults.Print "BIRTHDATE: Sep 26, 1980  (AGE 29)"
    picResults.Print "BIRTHPLACE: Ornskoldsvik, Sweden"
    picResults.Print "DRAFTED: VAN / 1999 NHL Entry Draft"
    picResults.Print "ROUND: 1st  (3rd overall)"
    picResults2.Picture = LoadPicture(App.Path & "\Sedin.jpg")
ElseIf J = "Vinny Prospal" Then
    picResults.Cls
    picResults.Print "NUMBER: 20"
    picResults.Print "HEIGHT: 6' 2"""
    picResults.Print "WEIGHT: 198"
    picResults.Print "Shoots: Left"
    picResults.Print "BIRTHDATE: Feb 17, 1975  (AGE 34)"
    picResults.Print "BIRTHPLACE: Ceske Budejovice, Czech Republic"
    picResults.Print "DRAFTED: PHI / 1993 NHL Entry Draft"
    picResults.Print "ROUND: 3rd  (71st overall)"
    picResults2.Picture = LoadPicture(App.Path & "\prospal.jpg")
ElseIf J = "Nicklas Backstrom" Then
    picResults.Cls
    picResults.Print "NUMBER: 19"
    picResults.Print "HEIGHT: 6' 1"""
    picResults.Print "WEIGHT: 210"
    picResults.Print "Shoots: Left"
    picResults.Print "BIRTHDATE: Nov 23, 1987  (AGE 21)"
    picResults.Print "BIRTHPLACE: Gavle, Sweden"
    picResults.Print "DRAFTED: WSH / 2006 NHL Entry Draft"
    picResults.Print "ROUND: 1st  (4th overall)"
    picResults2.Picture = LoadPicture(App.Path & "\backstrom.jpg")
ElseIf J = "Alexander Semin" Then
    picResults.Cls
    picResults.Print "NUMBER: 28"
    picResults.Print "HEIGHT: 6' 2"""
    picResults.Print "WEIGHT: 208"
    picResults.Print "Shoots: Right"
    picResults.Print "BIRTHDATE: Mar 3, 1984  (AGE 25)"
    picResults.Print "BIRTHPLACE: Krasnojarsk, Russia"
    picResults.Print "DRAFTED: WSH / 2002 NHL Entry Draft"
    picResults.Print "ROUND: 1st  (13th overall)"
    picResults2.Picture = LoadPicture(App.Path & "\semin.jpg")
ElseIf J = "Dany Heatley" Then
    picResults.Cls
    picResults.Print "NUMBER: 15"
    picResults.Print "HEIGHT: 6' 4"""
    picResults.Print "WEIGHT: 221"
    picResults.Print "Shoots: Left"
    picResults.Print "BIRTHDATE: Jan 21, 1981  (AGE 28)"
    picResults.Print "BIRTHPLACE: Freiburg, Germany"
    picResults.Print "DRAFTED: ATL / 2000 NHL Entry Draft"
    picResults.Print "ROUND: 1st  (2nd overall)"
    picResults2.Picture = LoadPicture(App.Path & "\heatley.jpg")
 Else
     MsgBox "Please Enter a Correct Name", , "Error"
    End If
   Next I
End Sub

Private Sub cmdPenaltyMin_Click()
Dim Pass As Integer
Dim Temp As String, Temp2 As String, Temp3 As String, Temp4 As String, Temp5 As String, Temp6 As String
Dim Temp7 As String, Temp8 As String, Temp9 As String, Temp10 As String, Temp11 As String, Temp12 As String, Temp13 As String

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

Private Sub cmdPlayer_Click()
Dim Pass As Integer
Dim Temp As String, Temp2 As String, Temp3 As String, Temp4 As String, Temp5 As String, Temp6 As String
Dim Temp7 As String, Temp8 As String, Temp9 As String, Temp10 As String, Temp11 As String, Temp12 As String, Temp13 As String

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

Private Sub cmdPlusMinus_Click()
Dim Pass As Integer
Dim Temp As String, Temp2 As String, Temp3 As String, Temp4 As String, Temp5 As String, Temp6 As String
Dim Temp7 As String, Temp8 As String, Temp9 As String, Temp10 As String, Temp11 As String, Temp12 As String, Temp13 As String

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
    Dim Pass As Integer
Dim Temp As String, Temp2 As String, Temp3 As String, Temp4 As String, Temp5 As String, Temp6 As String
Dim Temp7 As String, Temp8 As String, Temp9 As String, Temp10 As String, Temp11 As String, Temp12 As String, Temp13 As String

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
Dim Pass As Integer
Dim Temp As String, Temp2 As String, Temp3 As String, Temp4 As String, Temp5 As String, Temp6 As String
Dim Temp7 As String, Temp8 As String, Temp9 As String, Temp10 As String, Temp11 As String, Temp12 As String, Temp13 As String

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

Private Sub cmdShot_Click()
Dim Pass As Integer
Dim Temp As String, Temp2 As String, Temp3 As String, Temp4 As String, Temp5 As String, Temp6 As String
Dim Temp7 As String, Temp8 As String, Temp9 As String, Temp10 As String, Temp11 As String, Temp12 As String, Temp13 As String

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
Dim Pass As Integer
Dim Temp As String, Temp2 As String, Temp3 As String, Temp4 As String, Temp5 As String, Temp6 As String
Dim Temp7 As String, Temp8 As String, Temp9 As String, Temp10 As String, Temp11 As String, Temp12 As String, Temp13 As String

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

Private Sub cmdTimeOnIce_Click()
Dim Pass As Integer
Dim Temp As String, Temp2 As String, Temp3 As String, Temp4 As String, Temp5 As String, Temp6 As String
Dim Temp7 As String, Temp8 As String, Temp9 As String, Temp10 As String, Temp11 As String, Temp12 As String, Temp13 As String

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

