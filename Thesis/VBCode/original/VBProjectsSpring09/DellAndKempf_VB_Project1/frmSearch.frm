VERSION 5.00
Begin VB.Form frmSearch 
   BackColor       =   &H00808080&
   Caption         =   "Search for your Player"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   5895
      Left            =   10680
      Picture         =   "frmSearch.frx":0000
      ScaleHeight     =   5835
      ScaleWidth      =   4275
      TabIndex        =   4
      Top             =   1560
      Width           =   4335
   End
   Begin VB.TextBox txtPlayer 
      Height          =   735
      Left            =   720
      TabIndex        =   3
      Text            =   "Enter Last Name"
      Top             =   1080
      Width           =   3975
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Back to Main Menu"
      Height          =   495
      Left            =   8280
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000009&
      Height          =   3975
      Left            =   360
      ScaleHeight     =   3915
      ScaleWidth      =   9675
      TabIndex        =   1
      Top             =   2760
      Width           =   9735
   End
   Begin VB.CommandButton cmdFind1 
      BackColor       =   &H00000000&
      Caption         =   "Find Player"
      Height          =   495
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project-NFL Stats
'Form-frmSearch
'Written by Ryan Kempf and Ryan Dell
'3-22-09
'This form searches through the database and finds a player for you.
Private Sub cmdFind1_Click()
    Dim Player As String
    Dim Search As String
    Dim j As Integer
    Player = txtPlayer.Text
    Search = InputBox("Enter Position of Player--QB, RB, WR.", "Position")
    If Search = "QB" Then
        For j = 1 To ctrQ
            If Player = LastNameQB(j) Then
            picResults.Print "Name"; Tab(28); "Team"; Tab(50); "Yards", "QB Rating", "Comp %", "Yds/Att", "Pass TD"
            picResults.Print FirstNameQB(j); " "; LastNameQB(j); Tab(28); TeamQB(j), Tab(50); YardsQB(j), FormatNumber(QBRating(j), 3), FormatPercent(CompPct(j)), FormatNumber(YdsAtt(j), 2), PassTD(j)
            End If
        Next j
    ElseIf Search = "RB" Then
        For j = 1 To ctrR
            If Player = LastNameRB(j) Then
            picResults.Print "Name"; Tab(28); "Team"; Tab(50); "Yards", "Yards/Carry", "TD"
            picResults.Print FirstNameRB(j); " "; LastNameRB(j); Tab(28); TeamRB(j); Tab(50); YardsRB(j), FormatNumber(YPCRB(j), 2), TDRB(j)
            End If
        Next j
    ElseIf Search = "WR" Then
        For j = 1 To ctrW
            If Player = LastNameWR(j) Then
            picResults.Print "Name"; Tab(28); "Team"; Tab(50); "Yards", "Yds/Rec", "TD"
            picResults.Print FirstNameWR(j); " "; LastNameWR(j); Tab(28); TeamWR(j); Tab(50); YardsWR(j), FormatNumber(YPRWR(j), 2), TDWR(j)
            End If
        Next j
    Else: MsgBox "We could not find that player at the position, check your spelling", , "Error"
    End If
End Sub

Private Sub cmdMain_Click()
    frmSearch.Hide
    frmStartup.Show
End Sub
