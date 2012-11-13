VERSION 5.00
Begin VB.Form frmStatsEnter 
   BackColor       =   &H000000FF&
   Caption         =   "Stats Enter"
   ClientHeight    =   5370
   ClientLeft      =   3480
   ClientTop       =   3150
   ClientWidth     =   6660
   FillColor       =   &H000000FF&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   6660
   Begin VB.CommandButton cmdMain 
      Caption         =   "Return to Main Menu"
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   4560
      Width           =   3975
   End
   Begin VB.CommandButton cmdDefense 
      Caption         =   "Enter Team Defensive Stats"
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter Offensive Stats"
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   4095
      Left            =   2160
      ScaleHeight     =   4035
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmStatsEnter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDefense_Click()
'Declares Variables
Dim Games As Integer
Dim TotalPoints As Integer
Dim PointsPerGame As Double
Dim Yardage As Integer
Dim YardsPerGame As Double

'Takes in Raw Data to be processed
TotalPoints = InputBox("Enter Points Against")
Yardage = InputBox("Enter Total Yardage Against")
Games = InputBox("Enter Number of Games Played")

'Process data into new statistics
PointsPerGame = TotalPoints / Games
YardsPerGame = Yardage / Games

'Prints
picResults.Print "The defense gives up "; PointsPerGame; Chr(10); "points per game and"; Chr(10); YardsPerGame; "yards per game."


End Sub

Private Sub cmdEnter_Click()
'Declares all variables needed for array and search function
Dim ctr As Integer
Dim name(1 To 100) As String
Dim RushAtt(1 To 100) As Integer
Dim RushYards(1 To 100) As Integer
Dim RushTD(1 To 100) As Integer
Dim PassATT(1 To 100) As Integer
Dim PassComp(1 To 100) As Integer
Dim Passyds(1 To 100) As Integer
Dim PassTD(1 To 100) As Integer
Dim PassInt(1 To 100) As Integer
Dim Rec(1 To 100) As Integer
Dim RecYds(1 To 100) As Integer
Dim RecTD(1 To 100) As Integer
Dim counter As Integer
Dim nameTemp As String
Dim RushAttTemp As Integer
Dim RushYardsTemp As Integer
Dim RushTDTemp As Integer
Dim PassATTTemp As Integer
Dim PassCompTemp As Integer
Dim PassydsTemp As Integer
Dim PassTDTemp As Integer
Dim PassIntTemp As Integer
Dim RecTemp As Integer
Dim RecYdsTemp As Integer
Dim RecTDTemp As Integer
Dim Found As Boolean
Found = False
Dim SearchName As String

'Fills the array with stats
Open App.Path & "\OffensiveStats.txt" For Input As #1
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, name(ctr), RushAtt(ctr), RushYards(ctr), RushTD(ctr), PassATT(ctr), PassComp(ctr), Passyds(ctr), PassTD(ctr), PassInt(ctr), Rec(ctr), RecYds(ctr), RecTD(ctr)
Loop
Close #1

'Find Player to enter stats for
MsgBox ("Enter a name to enter stats for, formated FIRSTNAME LASTNAME.")
SearchName = InputBox("Player Name")

'Facilitates the input of new stats
Do While counter < ctr And Not Found
    counter = counter + 1
    If name(counter) = SearchName Then
    Found = True
    End If
Loop
    RushAttTemp = InputBox("Rushing Attempts")
    RushAtt(counter) = RushAtt(counter) + RushAttTemp
    RushTDTemp = InputBox("Rushing Touchdowns")
    RushTD(counter) = RushTD(counter) + RushTDTemp
    RushYardsTemp = InputBox("Rushing Yards")
    RushYards(counter) = RushYards(counter) + RushYardsTemp
    PassATTTemp = InputBox("Passing Attempts")
    PassATT(counter) = PassATT(counter) + PassATTTemp
    PassCompTemp = InputBox("Completions")
    PassComp(counter) = PassComp(counter) + PassCompTemp
    PassydsTemp = InputBox("Passing Yards")
    Passyds(counter) = Passyds(counter) + PassydsTemp
    PassTDTemp = InputBox("Passing Touchdowns")
    PassTD(counter) = PassTD(counter) + PassTDTemp
    PassIntTemp = InputBox("Passing Interceptions")
    PassInt(counter) = PassInt(counter) + PassIntTemp
    RecTemp = InputBox("Receptions")
    Rec(counter) = Rec(counter) + RecTemp
    RecYdsTemp = InputBox("Reception Yards")
    RecYds(counter) = RecYds(counter) + RecYdsTemp
    RecTDTemp = InputBox("Reception Touchdowns")
    RecTD(counter) = RecTD(counter) + RecTDTemp
    
    'Prints Players New Stat Totals
    picResults.Print "Name", name(counter); Chr(10); "Rush Att", RushAtt(counter); Chr(10); "Rushing Yrds", RushYards(counter); Chr(10); "Rushing TDs", RushTD(counter); Chr(10); "Pass Att", PassATT(counter); Chr(10); "Pass Comp", PassComp(counter); Chr(10); "Pass Yds", Passyds(counter); Chr(10); "Pass TDs", PassTD(counter); Chr(10); "Pass Int", PassInt(counter); Chr(10); "Rec", Rec(counter); Chr(10); "RecYds", RecYds(counter); Chr(10); "RecTDs", RecTD(counter); Chr(10)
    
    'Export the new data to text file
Open App.Path & "\OffensiveStats.txt" For Output As #1
Print #1, name(counter), ",", RushAtt(counter), ",", RushYards(counter), ",", RushTD(counter), ",", PassATT(counter), ",", PassComp(counter), ",", Passyds(counter), ",", PassTD(counter), ",", PassInt(counter), ",", Rec(counter), ",", RecYds(counter), ",", RecTD(counter)
Close #1

End Sub
'Takes user back to home page
Private Sub cmdMain_Click()
    frmStatsEnter.Hide
    frmHome.Show
End Sub
'Quits the program
Private Sub cmdQuit_Click()
    End
End Sub
