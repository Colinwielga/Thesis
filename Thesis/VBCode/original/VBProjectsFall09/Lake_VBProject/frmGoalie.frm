VERSION 5.00
Begin VB.Form frmGoalie 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   11340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19845
   LinkTopic       =   "Form1"
   ScaleHeight     =   11340
   ScaleWidth      =   19845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back to Main Page"
      Height          =   735
      Left            =   13920
      TabIndex        =   11
      Top             =   8280
      Width           =   4335
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "Learn More About the Players"
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   8520
      Width           =   3135
   End
   Begin VB.CommandButton cmdTimeOnIce 
      Caption         =   "Organize by Time on Ice"
      Height          =   1215
      Left            =   16080
      TabIndex        =   9
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdShutouts 
      Caption         =   "Organize by Shut Outs"
      Height          =   1215
      Left            =   13920
      TabIndex        =   8
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdSavePercentage 
      Caption         =   "Organize by Save Percentage"
      Height          =   1215
      Left            =   11880
      TabIndex        =   7
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdSaves 
      Caption         =   "Organize by Saves"
      Height          =   1215
      Left            =   9840
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdGGA 
      Caption         =   "Organize by Goals Against Avg"
      Height          =   1215
      Left            =   7920
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdGoals 
      Caption         =   "Organize by Goals Against"
      Height          =   1215
      Left            =   6000
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdShots 
      Caption         =   "Organize by Shots Against"
      Height          =   1215
      Left            =   4080
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdAlaphabetically 
      Caption         =   "View Players Alaphabetically"
      Height          =   1215
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "View Top 10 Goalies"
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H000000C0&
      Height          =   6255
      Left            =   720
      ScaleHeight     =   6195
      ScaleWidth      =   18555
      TabIndex        =   0
      Top             =   1920
      Width           =   18615
      Begin VB.PictureBox picResults2 
         BackColor       =   &H000000C0&
         Height          =   2655
         Left            =   600
         ScaleHeight     =   2595
         ScaleWidth      =   3195
         TabIndex        =   12
         Top             =   3360
         Width           =   3255
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Please Type Name of Player Exactly as it Appears Above"
      Height          =   495
      Left            =   5160
      TabIndex        =   13
      Top             =   8520
      Width           =   2295
   End
End
Attribute VB_Name = "frmGoalie"
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
Dim GamesPlayed(1 To 10) As Integer, Shots(1 To 10) As Single, GoalsAgainst(1 To 10) As Single, GoalsAA(1 To 10) As Single, Saves(1 To 10) As Single, SavePercentage(1 To 10) As Single, ShutOuts(1 To 10) As Single, TimeOnIce(1 To 10) As Single
Dim I As Integer
Dim Pass As Integer
Dim Temp As String, Temp2 As String, Temp3 As String, Temp4 As String, Temp5 As String, Temp6 As String
Dim Temp7 As String, Temp8 As String, Temp9 As String, Temp10 As String, Temp11 As String, Temp12 As String, Temp13 As String




Private Sub cmdAlaphabetically_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If Player(I) > Player(I + 1) Then
        Temp = GoalsAgainst(I)
        GoalsAgainst(I) = GoalsAgainst(I + 1)
        GoalsAgainst(I + 1) = Temp
        Temp2 = Player(I)
        Player(I) = Player(I + 1)
        Player(I + 1) = Temp2
        Temp3 = Teams(I)
        Teams(I) = Teams(I + 1)
        Teams(I + 1) = Temp3
        Temp4 = GamesPlayed(I)
        GamesPlayed(I) = GamesPlayed(I + 1)
        GamesPlayed(I + 1) = Temp4
        Temp5 = Shots(I)
        Shots(I) = Shots(I + 1)
        Shots(I + 1) = Temp5
        Temp6 = GoalsAgainst(I)
        GoalsAgainst(I) = GoalsAgainst(I + 1)
        GoalsAgainst(I + 1) = Temp6
        Temp7 = GoalsAA(I)
        GoalsAA(I) = GoalsAA(I + 1)
        GoalsAA(I + 1) = Temp7
        Temp8 = Saves(I)
        Saves(I) = Saves(I + 1)
        Saves(I + 1) = Temp8
        Temp9 = SavePercentage(I)
        SavePercentage(I) = SavePercentage(I + 1)
        SavePercentage(I + 1) = Temp9
        Temp10 = ShutOuts(I)
        ShutOuts(I) = ShutOuts(I + 1)
        ShutOuts(I + 1) = Temp10
        Temp11 = TimeOnIce(I)
        TimeOnIce(I) = TimeOnIce(I + 1)
        TimeOnIce(I + 1) = Temp11
        
        End If
        Next I
        Next Pass
        
        picResults.Cls
       picResults.Print "Player", Tab(20); "Team", Tab(40); "Games Played", "Shots", "Goals Against", "Goals Against Avg", "Saves", "SavePercentage", "Shutouts", "Time On Ice Per Game"
        picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"

        For I = 1 To 10
            picResults.Print Player(I); Tab(20); Teams(I); Tab; GamesPlayed(I); Tab; Shots(I); Tab; GoalsAgainst(I); Tab; GoalsAA(I); Tab; Tab; Saves(I); Tab; SavePercentage(I); Tab; Tab; ShutOuts(I); Tab; TimeOnIce(I)
        Next I
End Sub

Private Sub cmdAll_Click()
Open App.Path & "\Goalies.txt" For Input As #1
picResults.Print "Player", Tab(20); "Team", Tab(44); "Games Played", "Shots", "Goals Against", "Goals Against Avg", "Saves", "SavePercentage", "Shutouts", "Time On Ice Per Game"
picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"
For I = 1 To 10
    Input #1, Player(I), Teams(I), GamesPlayed(I), Shots(I), GoalsAgainst(I), GoalsAA(I), Saves(I), SavePercentage(I), ShutOuts(I), TimeOnIce(I)
 Next I
 For I = 1 To 10
    picResults.Print Player(I); Tab(20); Teams(I); Tab; GamesPlayed(I); Tab; Tab; Shots(I); Tab; GoalsAgainst(I); Tab; GoalsAA(I); Tab; Tab; Saves(I); Tab; SavePercentage(I); Tab; Tab; ShutOuts(I); Tab; TimeOnIce(I)

 Next I
End Sub

Private Sub cmdBack_Click()
    frmGoalie.Hide
    frmHockeyStatistics.Show
End Sub

Private Sub cmdGGA_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If GoalsAA(I) < GoalsAA(I + 1) Then
        Temp = GoalsAgainst(I)
        GoalsAgainst(I) = GoalsAgainst(I + 1)
        GoalsAgainst(I + 1) = Temp
        Temp2 = Player(I)
        Player(I) = Player(I + 1)
        Player(I + 1) = Temp2
        Temp3 = Teams(I)
        Teams(I) = Teams(I + 1)
        Teams(I + 1) = Temp3
        Temp4 = GamesPlayed(I)
        GamesPlayed(I) = GamesPlayed(I + 1)
        GamesPlayed(I + 1) = Temp4
        Temp5 = Shots(I)
        Shots(I) = Shots(I + 1)
        Shots(I + 1) = Temp5
        Temp6 = GoalsAgainst(I)
        GoalsAgainst(I) = GoalsAgainst(I + 1)
        GoalsAgainst(I + 1) = Temp6
        Temp7 = GoalsAA(I)
        GoalsAA(I) = GoalsAA(I + 1)
        GoalsAA(I + 1) = Temp7
        Temp8 = Saves(I)
        Saves(I) = Saves(I + 1)
        Saves(I + 1) = Temp8
        Temp9 = SavePercentage(I)
        SavePercentage(I) = SavePercentage(I + 1)
        SavePercentage(I + 1) = Temp9
        Temp10 = ShutOuts(I)
        ShutOuts(I) = ShutOuts(I + 1)
        ShutOuts(I + 1) = Temp10
        Temp11 = TimeOnIce(I)
        TimeOnIce(I) = TimeOnIce(I + 1)
        TimeOnIce(I + 1) = Temp11
        
        End If
        Next I
        Next Pass
        
        picResults.Cls
       picResults.Print "Player", Tab(20); "Team", Tab(44); "Games Played", "Shots", "Goals Against", "Goals Against Avg", "Saves", "SavePercentage", "Shutouts", "Time On Ice Per Game"
        picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"

        For I = 1 To 10
           picResults.Print Player(I); Tab(20); Teams(I); Tab; GamesPlayed(I); Tab; Tab; Shots(I); Tab; GoalsAgainst(I); Tab; GoalsAA(I); Tab; Tab; Saves(I); Tab; SavePercentage(I); Tab; Tab; ShutOuts(I); Tab; TimeOnIce(I)
        Next I
End Sub

Private Sub cmdGoals_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If GoalsAgainst(I) < GoalsAgainst(I + 1) Then
        Temp = GoalsAgainst(I)
        GoalsAgainst(I) = GoalsAgainst(I + 1)
        GoalsAgainst(I + 1) = Temp
        Temp2 = Player(I)
        Player(I) = Player(I + 1)
        Player(I + 1) = Temp2
        Temp3 = Teams(I)
        Teams(I) = Teams(I + 1)
        Teams(I + 1) = Temp3
        Temp4 = GamesPlayed(I)
        GamesPlayed(I) = GamesPlayed(I + 1)
        GamesPlayed(I + 1) = Temp4
        Temp5 = Shots(I)
        Shots(I) = Shots(I + 1)
        Shots(I + 1) = Temp5
        Temp6 = GoalsAgainst(I)
        GoalsAgainst(I) = GoalsAgainst(I + 1)
        GoalsAgainst(I + 1) = Temp6
        Temp7 = GoalsAA(I)
        GoalsAA(I) = GoalsAA(I + 1)
        GoalsAA(I + 1) = Temp7
        Temp8 = Saves(I)
        Saves(I) = Saves(I + 1)
        Saves(I + 1) = Temp8
        Temp9 = SavePercentage(I)
        SavePercentage(I) = SavePercentage(I + 1)
        SavePercentage(I + 1) = Temp9
        Temp10 = ShutOuts(I)
        ShutOuts(I) = ShutOuts(I + 1)
        ShutOuts(I + 1) = Temp10
        Temp11 = TimeOnIce(I)
        TimeOnIce(I) = TimeOnIce(I + 1)
        TimeOnIce(I + 1) = Temp11
        
        End If
        Next I
        Next Pass
        
        picResults.Cls
       picResults.Print "Player", Tab(20); "Team", Tab(44); "Games Played", "Shots", "Goals Against", "Goals Against Avg", "Saves", "SavePercentage", "Shutouts", "Time On Ice Per Game"
        picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"

        For I = 1 To 10
           picResults.Print Player(I); Tab(20); Teams(I); Tab; GamesPlayed(I); Tab; Tab; Shots(I); Tab; GoalsAgainst(I); Tab; GoalsAA(I); Tab; Tab; Saves(I); Tab; SavePercentage(I); Tab; Tab; ShutOuts(I); Tab; TimeOnIce(I)
        Next I
End Sub

Private Sub cmdMore_Click()
 Dim J As String
   Dim I As Integer
   
   J = InputBox("Enter a players name you want to see more information about", "Enter Player")
  For I = 1 To 10
   If J = "Marc-Andre Fleury" Then
    picResults.Cls
    picResults.Print "NUMBER: 29"
    picResults.Print "HEIGHT: 6' 2"""
    picResults.Print "WEIGHT: 280"
    picResults.Print "Catches Left"
    picResults.Print "BIRTHDATE: Nov 28, 1984  (AGE 24)"
    picResults.Print "BIRTHPLACE: Sorel, QC, Canada"
    picResults.Print "DRAFTED: PIT / 2003 NHL Entry Draft"
    picResults.Print "ROUND: 1st  (1st overall)"
    picResults2.Picture = LoadPicture(App.Path & "\fleury.jpg")
   ElseIf J = "Craig Anderson" Then
    picResults.Cls
    picResults.Print "NUMBER: 41"
    picResults.Print "HEIGHT: 6' 2"""
    picResults.Print "WEIGHT: 280"
    picResults.Print "Catches Left"
    picResults.Print "BIRTHDATE: May 21, 1981  (AGE 28)"
    picResults.Print "BIRTHPLACE: Park Ridge, IL, United States"
    picResults.Print "DRAFTED: CHI / 2001 NHL Entry Draft"
    picResults.Print "ROUND: 3rd  (73rd overall)"
    picResults2.Picture = LoadPicture(App.Path & "\anderson.jpg")
ElseIf J = "Henrik Lundqvist" Then
    picResults.Cls
    picResults.Print "NUMBER: 30"
    picResults.Print "HEIGHT: 6' 1"""
    picResults.Print "WEIGHT: 198"
    picResults.Print "Catches Left"
    picResults.Print "BIRTHDATE: Mar 2, 1982  (AGE 27)"
    picResults.Print "BIRTHPLACE: Are, Sweden"
    picResults.Print "DRAFTED: NYR / 2000 NHL Entry Draft "
    picResults.Print "ROUND: 7th  (205th overall)"
    picResults2.Picture = LoadPicture(App.Path & "\lundqvist.jpg")
ElseIf J = "Ilya Bryzgalov" Then
    picResults.Cls
    picResults.Print "NUMBER: 30"
    picResults.Print "HEIGHT: 6' 3"""
    picResults.Print "WEIGHT: 210"
    picResults.Print "Catches Left"
    picResults.Print "BIRTHDATE: Jun 22, 1980  (AGE 29)"
    picResults.Print "BIRTHPLACE: Togliatti, Russia"
    picResults.Print "DRAFTED: ANA / 2000 NHL Entry Draft"
    picResults.Print "ROUND: 2nd  (44th overall)"
    picResults2.Picture = LoadPicture(App.Path & "\bryzgalov.jpg")
ElseIf J = "Miikka Kiprusoff" Then
    picResults.Cls
    picResults.Print "NUMBER: 34"
    picResults.Print "HEIGHT: 6' 1"""
    picResults.Print "WEIGHT: 184"
    picResults.Print "Catches Left"
    picResults.Print "BIRTHDATE: Oct 26, 1976  (AGE 32)"
    picResults.Print "BIRTHPLACE: Turku, Finland"
    picResults.Print "DRAFTED: SJS / 1995 NHL Entry Draft"
    picResults.Print "ROUND: 5th  (116th overall)"
    picResults2.Picture = LoadPicture(App.Path & "\kiprusoff.jpg")
ElseIf J = "Ryan Miller" Then
    picResults.Cls
    picResults.Print "NUMBER: 30"
    picResults.Print "HEIGHT: 6' 2"""
    picResults.Print "WEIGHT: 175"
    picResults.Print "Catches Left"
    picResults.Print "BIRTHDATE: Jul 17, 1980  (AGE 29)"
    picResults.Print "BIRTHPLACE: East Lansing, MI, United States"
    picResults.Print "DRAFTED: BUF / 1999 NHL Entry Draft"
    picResults.Print "ROUND: 5th (138th overall)"
    picResults2.Picture = LoadPicture(App.Path & "\miller.jpg")
ElseIf J = "Pascal Leclaire" Then
    picResults.Cls
    picResults.Print "NUMBER: 33"
    picResults.Print "HEIGHT: 6' 2"""
    picResults.Print "WEIGHT: 202"
    picResults.Print "Catches Left"
    picResults.Print "BIRTHDATE: Nov 7, 1982  (AGE 26)"
    picResults.Print "BIRTHPLACE: Repentigny, QC, Canada"
    picResults.Print "DRAFTED: CBJ / 2001 NHL Entry Draft"
    picResults.Print "ROUND: 1st  (8th overall) "
    picResults2.Picture = LoadPicture(App.Path & "\leclaire.jpg")
ElseIf J = "Steve Mason" Then
    picResults.Cls
    picResults.Print "NUMBER: 1"
    picResults.Print "HEIGHT: 6' 4"""
    picResults.Print "WEIGHT: 220"
    picResults.Print "Catches Right"
    picResults.Print "BIRTHDATE: May 29, 1988  (AGE 21)"
    picResults.Print "BIRTHPLACE: Oakville, ON, Canada"
    picResults.Print "DRAFTED: CBJ / 2006 NHL Entry Draft"
    picResults.Print "ROUND: 3rd  (69th overall)  "
    picResults2.Picture = LoadPicture(App.Path & "\mason.jpg")
ElseIf J = "Evgeni Nabokov" Then
    picResults.Cls
    picResults.Print "NUMBER: 20"
    picResults.Print "HEIGHT: 6' 0"""
    picResults.Print "WEIGHT: 200"
    picResults.Print "Catches Left"
    picResults.Print "BIRTHDATE: Jul 25, 1975  (AGE 34)"
    picResults.Print "BIRTHPLACE: Kamenogorsk, Kazakhstan"
    picResults.Print "DRAFTED: SJS / 1994 NHL Entry Draft"
    picResults.Print "ROUND: 9th  (219th overall) "
    picResults2.Picture = LoadPicture(App.Path & "\nobokov.jpg")
ElseIf J = "Martin Brodeur" Then
    picResults.Cls
    picResults.Print "NUMBER: 30"
    picResults.Print "HEIGHT: 6' 2"""
    picResults.Print "WEIGHT: 215"
    picResults.Print "Catches Left"
    picResults.Print "BIRTHDATE: May 6, 1972  (AGE 37)"
    picResults.Print "BIRTHPLACE: Montreal, QC, Canada"
    picResults.Print "DRAFTED: NJD / 1990 NHL Entry Draft"
    picResults.Print "ROUND: 1st  (20th overall) "
    picResults2.Picture = LoadPicture(App.Path & "\brodeur.jpg")
Else
MsgBox "Please Enter a Correct Name", , "Error"
    End If
   Next I

End Sub

Private Sub cmdSavePercentage_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If SavePercentage(I) < SavePercentage(I + 1) Then
        Temp = GoalsAgainst(I)
        GoalsAgainst(I) = GoalsAgainst(I + 1)
        GoalsAgainst(I + 1) = Temp
        Temp2 = Player(I)
        Player(I) = Player(I + 1)
        Player(I + 1) = Temp2
        Temp3 = Teams(I)
        Teams(I) = Teams(I + 1)
        Teams(I + 1) = Temp3
        Temp4 = GamesPlayed(I)
        GamesPlayed(I) = GamesPlayed(I + 1)
        GamesPlayed(I + 1) = Temp4
        Temp5 = Shots(I)
        Shots(I) = Shots(I + 1)
        Shots(I + 1) = Temp5
        Temp6 = GoalsAgainst(I)
        GoalsAgainst(I) = GoalsAgainst(I + 1)
        GoalsAgainst(I + 1) = Temp6
        Temp7 = GoalsAA(I)
        GoalsAA(I) = GoalsAA(I + 1)
        GoalsAA(I + 1) = Temp7
        Temp8 = Saves(I)
        Saves(I) = Saves(I + 1)
        Saves(I + 1) = Temp8
        Temp9 = SavePercentage(I)
        SavePercentage(I) = SavePercentage(I + 1)
        SavePercentage(I + 1) = Temp9
        Temp10 = ShutOuts(I)
        ShutOuts(I) = ShutOuts(I + 1)
        ShutOuts(I + 1) = Temp10
        Temp11 = TimeOnIce(I)
        TimeOnIce(I) = TimeOnIce(I + 1)
        TimeOnIce(I + 1) = Temp11
        
        End If
        Next I
        Next Pass
        
        picResults.Cls
       picResults.Print "Player", Tab(20); "Team", Tab(44); "Games Played", "Shots", "Goals Against", "Goals Against Avg", "Saves", "SavePercentage", "Shutouts", "Time On Ice Per Game"
        picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"

        For I = 1 To 10
           picResults.Print Player(I); Tab(20); Teams(I); Tab; GamesPlayed(I); Tab; Tab; Shots(I); Tab; GoalsAgainst(I); Tab; GoalsAA(I); Tab; Tab; Saves(I); Tab; SavePercentage(I); Tab; Tab; ShutOuts(I); Tab; TimeOnIce(I)
        Next I
End Sub

Private Sub cmdSaves_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If Saves(I) < Saves(I + 1) Then
        Temp = GoalsAgainst(I)
        GoalsAgainst(I) = GoalsAgainst(I + 1)
        GoalsAgainst(I + 1) = Temp
        Temp2 = Player(I)
        Player(I) = Player(I + 1)
        Player(I + 1) = Temp2
        Temp3 = Teams(I)
        Teams(I) = Teams(I + 1)
        Teams(I + 1) = Temp3
        Temp4 = GamesPlayed(I)
        GamesPlayed(I) = GamesPlayed(I + 1)
        GamesPlayed(I + 1) = Temp4
        Temp5 = Shots(I)
        Shots(I) = Shots(I + 1)
        Shots(I + 1) = Temp5
        Temp6 = GoalsAgainst(I)
        GoalsAgainst(I) = GoalsAgainst(I + 1)
        GoalsAgainst(I + 1) = Temp6
        Temp7 = GoalsAA(I)
        GoalsAA(I) = GoalsAA(I + 1)
        GoalsAA(I + 1) = Temp7
        Temp8 = Saves(I)
        Saves(I) = Saves(I + 1)
        Saves(I + 1) = Temp8
        Temp9 = SavePercentage(I)
        SavePercentage(I) = SavePercentage(I + 1)
        SavePercentage(I + 1) = Temp9
        Temp10 = ShutOuts(I)
        ShutOuts(I) = ShutOuts(I + 1)
        ShutOuts(I + 1) = Temp10
        Temp11 = TimeOnIce(I)
        TimeOnIce(I) = TimeOnIce(I + 1)
        TimeOnIce(I + 1) = Temp11
        
        End If
        Next I
        Next Pass
        
        picResults.Cls
       picResults.Print "Player", Tab(20); "Team", Tab(44); "Games Played", "Shots", "Goals Against", "Goals Against Avg", "Saves", "SavePercentage", "Shutouts", "Time On Ice Per Game"
        picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"

        For I = 1 To 10
           picResults.Print Player(I); Tab(20); Teams(I); Tab; GamesPlayed(I); Tab; Tab; Shots(I); Tab; GoalsAgainst(I); Tab; GoalsAA(I); Tab; Tab; Saves(I); Tab; SavePercentage(I); Tab; Tab; ShutOuts(I); Tab; TimeOnIce(I)
        Next I
End Sub

Private Sub cmdShots_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If Shots(I) < Shots(I + 1) Then
        Temp = GoalsAgainst(I)
        GoalsAgainst(I) = GoalsAgainst(I + 1)
        GoalsAgainst(I + 1) = Temp
        Temp2 = Player(I)
        Player(I) = Player(I + 1)
        Player(I + 1) = Temp2
        Temp3 = Teams(I)
        Teams(I) = Teams(I + 1)
        Teams(I + 1) = Temp3
        Temp4 = GamesPlayed(I)
        GamesPlayed(I) = GamesPlayed(I + 1)
        GamesPlayed(I + 1) = Temp4
        Temp5 = Shots(I)
        Shots(I) = Shots(I + 1)
        Shots(I + 1) = Temp5
        Temp6 = GoalsAgainst(I)
        GoalsAgainst(I) = GoalsAgainst(I + 1)
        GoalsAgainst(I + 1) = Temp6
        Temp7 = GoalsAA(I)
        GoalsAA(I) = GoalsAA(I + 1)
        GoalsAA(I + 1) = Temp7
        Temp8 = Saves(I)
        Saves(I) = Saves(I + 1)
        Saves(I + 1) = Temp8
        Temp9 = SavePercentage(I)
        SavePercentage(I) = SavePercentage(I + 1)
        SavePercentage(I + 1) = Temp9
        Temp10 = ShutOuts(I)
        ShutOuts(I) = ShutOuts(I + 1)
        ShutOuts(I + 1) = Temp10
        Temp11 = TimeOnIce(I)
        TimeOnIce(I) = TimeOnIce(I + 1)
        TimeOnIce(I + 1) = Temp11
        
        End If
        Next I
        Next Pass
        
        picResults.Cls
       picResults.Print "Player", Tab(20); "Team", Tab(44); "Games Played", "Shots", "Goals Against", "Goals Against Avg", "Saves", "SavePercentage", "Shutouts", "Time On Ice Per Game"
        picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"

        For I = 1 To 10
           picResults.Print Player(I); Tab(20); Teams(I); Tab; GamesPlayed(I); Tab; Tab; Shots(I); Tab; GoalsAgainst(I); Tab; GoalsAA(I); Tab; Tab; Saves(I); Tab; SavePercentage(I); Tab; Tab; ShutOuts(I); Tab; TimeOnIce(I)
        Next I
End Sub

Private Sub cmdShutouts_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If ShutOuts(I) < ShutOuts(I + 1) Then
        Temp = GoalsAgainst(I)
        GoalsAgainst(I) = GoalsAgainst(I + 1)
        GoalsAgainst(I + 1) = Temp
        Temp2 = Player(I)
        Player(I) = Player(I + 1)
        Player(I + 1) = Temp2
        Temp3 = Teams(I)
        Teams(I) = Teams(I + 1)
        Teams(I + 1) = Temp3
        Temp4 = GamesPlayed(I)
        GamesPlayed(I) = GamesPlayed(I + 1)
        GamesPlayed(I + 1) = Temp4
        Temp5 = Shots(I)
        Shots(I) = Shots(I + 1)
        Shots(I + 1) = Temp5
        Temp6 = GoalsAgainst(I)
        GoalsAgainst(I) = GoalsAgainst(I + 1)
        GoalsAgainst(I + 1) = Temp6
        Temp7 = GoalsAA(I)
        GoalsAA(I) = GoalsAA(I + 1)
        GoalsAA(I + 1) = Temp7
        Temp8 = Saves(I)
        Saves(I) = Saves(I + 1)
        Saves(I + 1) = Temp8
        Temp9 = SavePercentage(I)
        SavePercentage(I) = SavePercentage(I + 1)
        SavePercentage(I + 1) = Temp9
        Temp10 = ShutOuts(I)
        ShutOuts(I) = ShutOuts(I + 1)
        ShutOuts(I + 1) = Temp10
        Temp11 = TimeOnIce(I)
        TimeOnIce(I) = TimeOnIce(I + 1)
        TimeOnIce(I + 1) = Temp11
        
        End If
        Next I
        Next Pass
        
        picResults.Cls
       picResults.Print "Player", Tab(20); "Team", Tab(44); "Games Played", "Shots", "Goals Against", "Goals Against Avg", "Saves", "SavePercentage", "Shutouts", "Time On Ice Per Game"
        picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"

        For I = 1 To 10
           picResults.Print Player(I); Tab(20); Teams(I); Tab; GamesPlayed(I); Tab; Tab; Shots(I); Tab; GoalsAgainst(I); Tab; GoalsAA(I); Tab; Tab; Saves(I); Tab; SavePercentage(I); Tab; Tab; ShutOuts(I); Tab; TimeOnIce(I)
        Next I
End Sub

Private Sub cmdTimeOnIce_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If TimeOnIce(I) < TimeOnIce(I + 1) Then
        Temp = GoalsAgainst(I)
        GoalsAgainst(I) = GoalsAgainst(I + 1)
        GoalsAgainst(I + 1) = Temp
        Temp2 = Player(I)
        Player(I) = Player(I + 1)
        Player(I + 1) = Temp2
        Temp3 = Teams(I)
        Teams(I) = Teams(I + 1)
        Teams(I + 1) = Temp3
        Temp4 = GamesPlayed(I)
        GamesPlayed(I) = GamesPlayed(I + 1)
        GamesPlayed(I + 1) = Temp4
        Temp5 = Shots(I)
        Shots(I) = Shots(I + 1)
        Shots(I + 1) = Temp5
        Temp6 = GoalsAgainst(I)
        GoalsAgainst(I) = GoalsAgainst(I + 1)
        GoalsAgainst(I + 1) = Temp6
        Temp7 = GoalsAA(I)
        GoalsAA(I) = GoalsAA(I + 1)
        GoalsAA(I + 1) = Temp7
        Temp8 = Saves(I)
        Saves(I) = Saves(I + 1)
        Saves(I + 1) = Temp8
        Temp9 = SavePercentage(I)
        SavePercentage(I) = SavePercentage(I + 1)
        SavePercentage(I + 1) = Temp9
        Temp10 = ShutOuts(I)
        ShutOuts(I) = ShutOuts(I + 1)
        ShutOuts(I + 1) = Temp10
        Temp11 = TimeOnIce(I)
        TimeOnIce(I) = TimeOnIce(I + 1)
        TimeOnIce(I + 1) = Temp11
        
        End If
        Next I
        Next Pass
        
        picResults.Cls
       picResults.Print "Player", Tab(20); "Team", Tab(44); "Games Played", "Shots", "Goals Against", "Goals Against Avg", "Saves", "SavePercentage", "Shutouts", "Time On Ice Per Game"
        picResults.Print "********************************************************************************************************************************************************************************************************************************************************************************************************************************"

        For I = 1 To 10
           picResults.Print Player(I); Tab(20); Teams(I); Tab; GamesPlayed(I); Tab; Tab; Shots(I); Tab; GoalsAgainst(I); Tab; GoalsAA(I); Tab; Tab; Saves(I); Tab; SavePercentage(I); Tab; Tab; ShutOuts(I); Tab; TimeOnIce(I)
        Next I
End Sub
