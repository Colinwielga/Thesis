VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H000000FF&
   Caption         =   "Form3"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "BACK TO THE START"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6480
      TabIndex        =   10
      Top             =   6960
      Width           =   3495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "QUIT"
      Height          =   615
      Left            =   1440
      TabIndex        =   9
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "CLEAR"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   6720
      Width           =   975
   End
   Begin VB.PictureBox picbox 
      Height          =   5175
      Left            =   3000
      ScaleHeight     =   5115
      ScaleWidth      =   9795
      TabIndex        =   7
      Top             =   1320
      Width           =   9855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "PREDICTIONS"
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Projected Homerun Totals at Age 42"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Average homeruns hit per season"
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sort players by homeruns"
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sort players by present age"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Hank Aaron's last year of playing baseball was at the age of 42."
      Height          =   495
      Left            =   10680
      TabIndex        =   11
      Top             =   9240
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   $"Homerun3.frx":0000
      Height          =   615
      Left            =   5280
      TabIndex        =   1
      Top             =   600
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "Who really has a chance at the record?"
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
      Left            =   5520
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:  Project1 (Homerun.vbp)
'Form Name:  Form3 (Homerun3.frm)
'Author:  Garret Flood
'Date Written:  Oct 28th, 2003
'Purpose of form:  Sorts players by age and by homeruns.
'Projects how many homeruns they will hit by a certain age.
'Give predictions on what they chances are of breaking the homerun record.

'Option Explict makes the programmer declare all varables on the form.
Option Explicit
Dim Players(1 To 12) As String
Dim YearsPlayed(1 To 12), AtBats(1 To 12), Hits(1 To 12), Homeruns(1 To 12), BattingAvg(1 To 12), Age(1 To 12) As Single
Dim X As Integer
Private Sub Command1_Click()
picbox.Cls
'This command sorts players by age.
Dim Pass As Integer
Dim TempAge, TempYearsPlayed, TempAtBats, TempHits, TempBattingAvg, TempHomeruns As Single
Dim TempPlayers As String
Dim Y As Integer

Open PATH & "stats.txt" For Input As #1
    picbox.Print "*************************************************************************************************************************************************************"
    picbox.Print "Players"; Tab(20); "Years Played"; Tab(40); "At Bats"; Tab(60); "Hits"; Tab(80); "Homeruns"; Tab(100); "Batting Avg"; Tab(120); "Age"
    picbox.Print "*************************************************************************************************************************************************************"

For X = 1 To 12
    Input #1, Players(X), YearsPlayed(X), AtBats(X), Hits(X), Homeruns(X), BattingAvg(X), Age(X)
Next X

'Sorts players by age and changes information to match right player
For Pass = 1 To 12
    For X = 1 To 12 - Pass
        If Age(X) > Age(X + 1) Then
            TempAge = Age(X)
            Age(X) = Age(X + 1)
            Age(X + 1) = TempAge
            TempPlayers = Players(X)
            Players(X) = Players(X + 1)
            Players(X + 1) = TempPlayers
            TempYearsPlayed = YearsPlayed(X)
            YearsPlayed(X) = YearsPlayed(X + 1)
            YearsPlayed(X + 1) = TempYearsPlayed
            TempAtBats = AtBats(X)
            AtBats(X) = AtBats(X + 1)
            AtBats(X + 1) = TempAtBats
            TempHits = Hits(X)
            Hits(X) = Hits(X + 1)
            Hits(X + 1) = TempHits
            TempHomeruns = Homeruns(X)
            Homeruns(X) = Homeruns(X + 1)
            Homeruns(X + 1) = TempHomeruns
            TempBattingAvg = BattingAvg(X)
            BattingAvg(X) = BattingAvg(X + 1)
            BattingAvg(X + 1) = TempBattingAvg
        End If
Next X
Next Pass

'Prints players in the appropriate order
For Y = 1 To 12
    picbox.Print Players(Y); Tab(20); YearsPlayed(Y); Tab(40); AtBats(Y); Tab(60); Hits(Y); Tab(80); Homeruns(Y); Tab(100); BattingAvg(Y); Tab(120); Age(Y)
Next Y

Close #1
End Sub

Private Sub Command2_Click()
picbox.Cls
'This command sorts players by homeruns
Dim Pass As Integer
Dim TempAge, TempYearsPlayed, TempAtBats, TempHits, TempBattingAvg, TempHomeruns As Single
Dim TempPlayers As String
Dim Y As Integer

Open PATH & "stats.txt" For Input As #1
    picbox.Print "*************************************************************************************************************************************************************"
    picbox.Print "Players"; Tab(20); "Years Played"; Tab(40); "At Bats"; Tab(60); "Hits"; Tab(80); "Homeruns"; Tab(100); "Batting Avg"; Tab(120); "Age"
    picbox.Print "*************************************************************************************************************************************************************"

For X = 1 To 12
    Input #1, Players(X), YearsPlayed(X), AtBats(X), Hits(X), Homeruns(X), BattingAvg(X), Age(X)
Next X

'This sorts the players by homeruns and matches the right information
For Pass = 1 To 12
    For X = 1 To 12 - Pass
        If Homeruns(X) < Homeruns(X + 1) Then
            TempHomeruns = Homeruns(X)
            Homeruns(X) = Homeruns(X + 1)
            Homeruns(X + 1) = TempHomeruns
            TempPlayers = Players(X)
            Players(X) = Players(X + 1)
            Players(X + 1) = TempPlayers
            TempYearsPlayed = YearsPlayed(X)
            YearsPlayed(X) = YearsPlayed(X + 1)
            YearsPlayed(X + 1) = TempYearsPlayed
            TempAtBats = AtBats(X)
            AtBats(X) = AtBats(X + 1)
            AtBats(X + 1) = TempAtBats
            TempHits = Hits(X)
            Hits(X) = Hits(X + 1)
            Hits(X + 1) = TempHits
            TempAge = Age(X)
            Age(X) = Age(X + 1)
            Age(X + 1) = TempAge
            TempBattingAvg = BattingAvg(X)
            BattingAvg(X) = BattingAvg(X + 1)
            BattingAvg(X + 1) = TempBattingAvg
        End If
Next X
Next Pass
        
'Prints the players in order of homeruns
For Y = 1 To 12
    picbox.Print Players(Y); Tab(20); YearsPlayed(Y); Tab(40); AtBats(Y); Tab(60); Hits(Y); Tab(80); Homeruns(Y); Tab(100); BattingAvg(Y); Tab(120); Age(Y)
Next Y

Close #1
End Sub

Private Sub Command3_Click()
picbox.Cls
'Calculates how many homeruns a player hits per year
Dim Avg As Integer
Open PATH & "stats.txt" For Input As #1
    picbox.Print "******************************************************"
    picbox.Print "Players"; Tab(20); "Avg Homeruns Per Year"
    picbox.Print "******************************************************"

'Calculates and prints player and homeruns per year
For X = 1 To 12
    Input #1, Players(X), YearsPlayed(X), AtBats(X), Hits(X), Homeruns(X), BattingAvg(X), Age(X)
    Avg = Homeruns(X) / YearsPlayed(X)
    picbox.Print Players(X); Tab(20), Avg
Next X

Close #1
End Sub

Private Sub Command4_Click()
picbox.Cls
'Calculates how many homeruns a player will hit by the age of 42.
Dim Total As Integer

Open PATH & "stats.txt" For Input As #1
    picbox.Print "*******************************************************************************************"
    picbox.Print "Players"; Tab(20); "Projected Homeruns At Age 42"; Tab(60); "Present Age"
    picbox.Print "*******************************************************************************************"

'Calculates and prints player and his homerun total at age 42.
For X = 1 To 12
    Input #1, Players(X), YearsPlayed(X), AtBats(X), Hits(X), Homeruns(X), BattingAvg(X), Age(X)
    Total = (Homeruns(X) / YearsPlayed(X)) * (42 - Age(X)) + Homeruns(X)
    picbox.Print Players(X); Tab(20); Total; Tab(60); Age(X)
Next X

Close #1
End Sub

Private Sub Command5_Click()
picbox.Cls
'I give my opinion for each player on their chances of breaking the homerun record.
    picbox.Print "*****************************************************************PREDICTIONS************************************************************************"
    picbox.Print "Jeff Bagwell--He is a good hitter, but the record is out of his reach."
    picbox.Print "Rafael Palmeiro--Always hit a lot of homeruns, but he is too old."
    picbox.Print "Juan Gonzalez--Injuries will keep him at a respectable homerun total, but not enough."
    picbox.Print "Frank Thomas--Frank tries to keep his batting average up.  This lowers his homerun totals."
    picbox.Print "Vladimir Guerrero--Vlad is capable of any kind of hitting record, but he swings at too many bad pitches."
    picbox.Print "Manny Ramirez--Does not have enough homeruns per year to break the record."
    picbox.Print "Ken Griffey--He has a chance because of his projected homerun total, but he is injured too often."
    picbox.Print "Albert Puljos--Will shatter the record if he keeps up his homerun pace.  He is very young though, too soon to tell."
    picbox.Print "Alex Rodriguez--Alex is a very consistent homerun hitter.  He will most likely pass Hank Aaron for career homeruns."
    picbox.Print "Sammy Sosa--Sammy is one of the great homerun hitters of all time.  If he stays healthy and plays for a while, he has a good shot."
    picbox.Print "Barry Bonds--Possibly the greatest homerun hitter of all time.  He will break the record first if he does not get injured."
    picbox.Print "************************************************************************************************************************************************************"
    picbox.Print "Best Chance:  Barry Bonds"
End Sub

Private Sub Command6_Click()
'Clears the picture box
picbox.Cls
End Sub

Private Sub Command7_Click()
'Quits program
End
End Sub

Private Sub Command8_Click()
'Goes back to the start of the program
Form3.Hide
Form1.Show
End Sub

