VERSION 5.00
Begin VB.Form frmSchedulesScores 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   5160
   ClientLeft      =   4425
   ClientTop       =   2775
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   6510
   Begin VB.CommandButton cmdMenu 
      BackColor       =   &H000000FF&
      Caption         =   "Menu"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdSortOpponent 
      Caption         =   "Opponent"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSortWins 
      Caption         =   "Wins"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdSortDate 
      Caption         =   "Date (most recent)"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   960
      Width           =   6255
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display Schedule and Recent Scores"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "...Sort by..."
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmSchedulesScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'St. John's Rugby
'Sam Herrmann
'March 2009

Option Explicit

'sort by wins/losses
'sort by opponents
'sort by date

Dim time(1 To 20) As Date, team(1 To 20) As String, location(1 To 20) As String, outcome(1 To 20) As String
Dim tempTime As Date, tempTeam As String, tempLocation As String, tempOutcome As String
Dim pass As Integer, pos As Integer, j As Integer

Private Sub cmdDisplay_Click()

CTR = 0
picResults.Cls
Open App.Path & "\schedscore.txt" For Input As #1

Do While Not EOF(1)
    CTR = CTR + 1
        Input #1, time(CTR), team(CTR), location(CTR), outcome(CTR)
        picResults.Print time(CTR); Tab(20); team(CTR); Tab(50); location(CTR); Tab(69); outcome(CTR)
Loop

Close #1

End Sub

Private Sub cmdMenu_Click()
frmMenu.Show
frmSchedulesScores.Hide
End Sub

Private Sub cmdSortDate_Click()

picResults.Cls

For pass = 1 To CTR - 1
    For pos = 1 To CTR - pass
        If time(pos) < time(pos + 1) Then
            tempTime = time(pos)
            time(pos) = time(pos + 1)
            time(pos + 1) = tempTime
            tempTeam = team(pos)
            team(pos) = team(pos + 1)
            team(pos + 1) = tempTeam
            tempLocation = location(pos)
            location(pos) = location(pos + 1)
            location(pos + 1) = tempLocation
            tempOutcome = outcome(pos)
            outcome(pos) = outcome(pos + 1)
            outcome(pos + 1) = tempOutcome
        End If
    Next pos
Next pass

    For j = 1 To CTR
             picResults.Print time(j); Tab(20); team(j); Tab(50); location(j); Tab(69); outcome(j)
    Next j

End Sub

Private Sub cmdSortOpponent_Click()

picResults.Cls

For pass = 1 To CTR - 1
    For pos = 1 To CTR - pass
        If team(pos) < team(pos + 1) Then
            tempTime = time(pos)
            time(pos) = time(pos + 1)
            time(pos + 1) = tempTime
            tempTeam = team(pos)
            team(pos) = team(pos + 1)
            team(pos + 1) = tempTeam
            tempLocation = location(pos)
            location(pos) = location(pos + 1)
            location(pos + 1) = tempLocation
            tempOutcome = outcome(pos)
            outcome(pos) = outcome(pos + 1)
            outcome(pos + 1) = tempOutcome
        End If
    Next pos
Next pass

    For j = 1 To CTR
             picResults.Print time(j); Tab(20); team(j); Tab(50); location(j); Tab(69); outcome(j)
    Next j


End Sub

Private Sub cmdSortWins_Click()

picResults.Cls

For pass = 1 To CTR - 1
    For pos = 1 To CTR - pass
        If outcome(pos) < outcome(pos + 1) Then
            tempTime = time(pos)
            time(pos) = time(pos + 1)
            time(pos + 1) = tempTime
            tempTeam = team(pos)
            team(pos) = team(pos + 1)
            team(pos + 1) = tempTeam
            tempLocation = location(pos)
            location(pos) = location(pos + 1)
            location(pos + 1) = tempLocation
            tempOutcome = outcome(pos)
            outcome(pos) = outcome(pos + 1)
            outcome(pos + 1) = tempOutcome
        End If
    Next pos
Next pass

    For j = 1 To CTR
             picResults.Print time(j); Tab(20); team(j); Tab(50); location(j); Tab(69); outcome(j)
    Next j

End Sub

