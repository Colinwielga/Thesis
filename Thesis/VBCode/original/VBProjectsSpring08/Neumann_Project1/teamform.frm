VERSION 5.00
Begin VB.Form teamform 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13890
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   13890
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picone 
      AutoSize        =   -1  'True
      FontTransparent =   0   'False
      Height          =   3060
      Left            =   10920
      Picture         =   "teamform.frx":0000
      ScaleHeight     =   3000
      ScaleWidth      =   3000
      TabIndex        =   9
      Top             =   3720
      Width           =   3060
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   -120
      Picture         =   "teamform.frx":245B
      ScaleHeight     =   3000
      ScaleWidth      =   3000
      TabIndex        =   8
      Top             =   3720
      Width           =   3060
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Go back to home page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1800
      TabIndex        =   7
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9840
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
   End
   Begin VB.PictureBox picresults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   3000
      ScaleHeight     =   3435
      ScaleWidth      =   7755
      TabIndex        =   5
      Top             =   3480
      Width           =   7815
   End
   Begin VB.CommandButton cmddata 
      Caption         =   "Get team data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5400
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton cmdgame 
      Caption         =   "Find the  highest game of the night"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10080
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdseries 
      Caption         =   "Sort by highest series"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6960
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdfill 
      Caption         =   "Calculate fill percentages"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3840
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdhandicap 
      Caption         =   "Calculate team handicap"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "teamform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Bowling prodject
'team form
'Zach Neumann
'3/30/2008
'This form is meant to look at different aspects of team data that might be important to a team
Option Explicit
Dim names(1 To 5) As String, gameone(1 To 5) As Integer, gametwo(1 To 5) As Integer, gamethree(1 To 5) As Integer, avg(1 To 5) As Single, fills(1 To 5) As Integer, ctr As Integer

Private Sub cmdback_Click()
    teamform.Hide
    startform.Show
End Sub
'loads the data from a document that has the bowlers 1st, 2nd, and 3rd games, their current average
'and the number of frames they filled(got either a strike or spare)
Private Sub cmddata_Click()
cmdhandicap.Enabled = True
cmdfill.Enabled = True
cmdseries.Enabled = True
cmdgame.Enabled = True

ctr = 0

Open App.Path & "\data.txt" For Input As #1
Do Until EOF(1)
    ctr = ctr + 1
    Input #1, names(ctr), gameone(ctr), gametwo(ctr), gamethree(ctr), avg(ctr), fills(ctr)
Loop
Close #1
End Sub

Private Sub cmdfill_Click()
Dim pass As Integer, pos As Integer, temp As Integer, fillpercentage As Single, N As Integer, tempnames As String
picresults.Cls
'bubble sort to look at and sort each bowler's fill%
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If fills(pos) < fills(pos + 1) Then
            temp = fills(pos)
            fills(pos) = fills(pos + 1)
            fills(pos + 1) = temp
            tempnames = names(pos)
            names(pos) = names(pos + 1)
            names(pos + 1) = tempnames
        End If
    Next pos
Next pass

For N = 1 To ctr
    fillpercentage = fills(N) / 36
    picresults.Print names(N); Tab(20); FormatPercent(fillpercentage, 1)
Next N
End Sub

Private Sub cmdgame_Click()
Dim N As Integer, highgameone As Integer, highgametwo As Integer, highgamethree As Integer, highscore As Integer
'this finds the high game bowled that night
ctr = ctr
picresults.Cls
highgameone = 0
For N = 1 To ctr
    If highgameone < gameone(N) Then
        highgameone = gameone(N)
    End If
Next N
highgametwo = 0
For N = 1 To ctr
    If highgametwo < gametwo(N) Then
        highgametwo = gametwo(N)
    End If
Next N
highgamethree = 0
For N = 1 To ctr
    If highgamethree < gamethree(N) Then
        highgamethree = gamethree(N)
    End If
Next N


If highgameone >= highgametwo And highgameone >= highgamethree Then
        picresults.Print "The highest game was a score of: "; highgameone
ElseIf highgametwo > highgameone And highgametwo > highgamethree Then
        picresults.Print "The highest game was a score of: "; highgametwo
ElseIf highgamethree > highgameone And highgamethree > highgametwo Then
        picresults.Print "The highest game with a score of: "; highgamethree
End If


End Sub

Private Sub cmdhandicap_Click()
picresults.Cls
Dim N As Integer, handicap As Integer
For N = 1 To ctr
    handicap = 220 - avg(N) + handicap
Next N
picresults.Print "The team handicap is "; handicap
End Sub

Private Sub cmdquit_Click()
End

End Sub

Private Sub cmdseries_Click()
picresults.Cls
Dim pass As Integer, pos As Integer, temp As Integer, tempnames As String, N As Integer, series(1 To 5) As Integer, I As Integer
'this searches through the players series(total of all three games) that each bowler bowled
'and then sorts them from high to low with the name of the bowler who bowled it

For I = 1 To 5
    series(I) = gameone(I) + gametwo(I) + gamethree(I)
Next I

For pass = 1 To 5 - 1
    For pos = 1 To 5 - pass
        If series(pos) < series(pos + 1) Then
            temp = series(pos)
            series(pos) = series(pos + 1)
            series(pos + 1) = temp
            tempnames = names(pos)
            names(pos) = names(pos + 1)
            names(pos + 1) = tempnames
        End If
    Next pos
Next pass

For N = 1 To 5
    picresults.Print names(N); Tab(20); series(N)
Next N
End Sub

Private Sub picone_Click()
picone.Image = pin

End Sub
