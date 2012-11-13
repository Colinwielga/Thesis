VERSION 5.00
Begin VB.Form frmstats 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   11640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   ScaleHeight     =   11640
   ScaleWidth      =   15210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return To Main Menu"
      Height          =   1215
      Left            =   360
      TabIndex        =   7
      Top             =   9240
      Width           =   3015
   End
   Begin VB.CommandButton cmdsearchscoringpos 
      Caption         =   "Sort By Best with Runners in Scoring Position"
      Height          =   1215
      Left            =   360
      TabIndex        =   6
      Top             =   7920
      Width           =   3015
   End
   Begin VB.CommandButton cmdwinning 
      Caption         =   " Sort by Best Winning Percentage"
      Height          =   1215
      Left            =   360
      TabIndex        =   5
      Top             =   6600
      Width           =   3015
   End
   Begin VB.CommandButton cmdbyavg 
      Caption         =   "Sort by Best Batting Average"
      Height          =   1215
      Left            =   360
      TabIndex        =   4
      Top             =   5160
      Width           =   3015
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search by Name of Player"
      Height          =   1215
      Left            =   360
      TabIndex        =   3
      Top             =   3840
      Width           =   3015
   End
   Begin VB.CommandButton cmddisplay 
      Caption         =   "Display all players Stats"
      Height          =   1215
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load Player Data"
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   3015
   End
   Begin VB.PictureBox picresults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10095
      Left            =   3600
      ScaleHeight     =   10035
      ScaleWidth      =   11235
      TabIndex        =   0
      Top             =   720
      Width           =   11295
   End
End
Attribute VB_Name = "frmstats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form shows the current players names and some stats about those players
'stats that are displayed are their name, batting average, avg with runners in scoring position, and win percentage
'this form uses arrays, loops, sorting, and searching
Dim player(1 To 20) As String
Dim battingavg(1 To 20) As Single
Dim scoringpos(1 To 20) As Single
Dim winning(1 To 20) As Single
Dim CTR As Integer
Private Sub cmdbyavg_Click()
'this subroutine sorts the batting averages in descending order by using the bubble sort method

Dim pos As Integer
Dim Temp1 As String
Dim temp2 As Single
Dim pass As Integer
Dim temp3 As Integer
Dim x As Integer



picresults.Cls



pos = 0
For pass = 1 To CTR - 1
    For pos = 1 To (CTR - pass)
        If battingavg(pos) < battingavg(pos + 1) Then
            Temp1 = battingavg(pos)
            battingavg(pos) = battingavg(pos + 1)
            battingavg(pos + 1) = Temp1
            temp2 = scoringpos(pos)
            scoringpos(pos) = scoringpos(pos + 1)
            scoringpos(pos + 1) = temp2
            temp3 = winning(pos)
            winning(pos) = winning(pos + 1)
            winning(pos + 1) = temp3
        End If
    Next pos
Next pass

picresults.Print "Player Name"; Tab(35); "Batting Average"; Tab(55); "Winning Percentage"; Tab(65)
picresults.Print "*****************************************************************************************************"
For x = 1 To CTR
    picresults.Print player(x); Tab(35); FormatNumber(battingavg(x), 3); Tab(55); FormatNumber(winning(x), 3); Tab(65)
    picresults.Print "-----------------------------------------------------------------------------------------------"
Next x
End Sub

Private Sub cmddisplay_Click()
'this subroutine displays all of the players information in a picturebox
Dim x As Integer
Dim pos As Integer

picresults.Cls

pos = 0
CTR = 10
picresults.Print "Player Name"; Tab(15); "Batting Average"; Tab(40); "With Runners in Scoring Position"; Tab(55)
    picresults.Print "************************************************************************************************************"
For pos = 1 To CTR
    picresults.Print player(pos); Tab(15); FormatNumber(battingavg(pos), 3); Tab(40); FormatNumber(scoringpos(pos), 3); Tab(65)
    picresults.Print "------------------------------------------------------------------------------------------------"
Next pos
End Sub

Private Sub cmdload_Click()
'this subroutine loads all of the inforamtion from a text file into arrays within this form
Open App.Path & "\Stats.txt" For Input As #1

CTR = 0
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, player(CTR), battingavg(CTR), scoringpos(CTR), winning(CTR)
Loop


Close #1



End Sub

Private Sub cmdreturn_Click()
'returns the user back to the main menu
frmmain.Show
frmstats.Hide
End Sub

Private Sub cmdsearch_Click()
'this subroutine asks the user for a name and searches the arrays until it finds
Dim Sname As String
Dim pos As Integer
Dim Found As Boolean
Dim players(1 To 20) As String
Dim x As Integer
Dim batting(1 To 20) As Single
Dim scoring(1 To 20) As Single
Dim percent(1 To 20) As Single
Open App.Path & "\stats.txt" For Input As #1
x = 0
Do Until EOF(1)
        x = x + 1
        Input #1, players(x), batting(x), scoring(x), percent(x)
Loop
Close #1
picresults.Cls
Sname = InputBox("Input Player Name")
pos = 0
Found = False
Do While (Found = False And pos < x)
    pos = pos + 1
    If LCase(player(pos)) = LCase(Sname) Then
        Found = True
    End If
Loop
    If Found = True Then
        picresults.Print player(pos); " has a batting average of "; FormatNumber(battingavg(pos)) & " and an batting average of "; FormatNumber(scoringpos(pos)); " with runners in scoring position. "
        picresults.Print player(pos); " also has a winning percentage of "; FormatNumber(winning(pos))
    Else
        MsgBox "No Matches"
End If
End Sub

Private Sub cmdsearchscoringpos_Click()
'sorts the people in the array by avg with runners in scoring position in descending order
Dim pos As Integer
Dim Temp1 As String
Dim temp2 As Single
Dim pass As Integer
Dim temp3 As Integer
Dim x As Integer



picresults.Cls



pos = 0
For pass = 1 To CTR - 1
    For pos = 1 To (CTR - pass)
        If scoringpos(pos) < scoringpos(pos + 1) Then
            Temp1 = scoringpos(pos)
            scoringpos(pos) = scoringpos(pos + 1)
            scoringpos(pos + 1) = Temp1
        End If
    Next pos
Next pass

picresults.Print "Player Name"; Tab(35); "Winning Percentage"; Tab(55)
picresults.Print "*********************************************************************************************************"
For x = 1 To CTR
    picresults.Print player(x); Tab(35); FormatNumber(scoringpos(x), 3)
    picresults.Print "-----------------------------------------------------------------------------------------------"
Next x

End Sub

Private Sub cmdwinning_Click()
'sorts the users in the array by winning percentage in descending order
Dim pos As Integer
Dim Temp1 As String
Dim temp2 As Single
Dim pass As Integer
Dim temp3 As Integer
Dim x As Integer



picresults.Cls



pos = 0
For pass = 1 To CTR - 1
    For pos = 1 To (CTR - pass)
        If winning(pos) < winning(pos + 1) Then
            Temp1 = winning(pos)
            winning(pos) = winning(pos + 1)
            winning(pos + 1) = Temp1
        End If
    Next pos
Next pass

picresults.Print "Player Name"; Tab(35); "Winning Percentage"; Tab(55)
picresults.Print "********************************************************************************************************"
For x = 1 To CTR
    picresults.Print player(x); Tab(35); FormatNumber(winning(x)) & " %. "
    picresults.Print "-----------------------------------------------------------------------------------------------"
Next x

End Sub
