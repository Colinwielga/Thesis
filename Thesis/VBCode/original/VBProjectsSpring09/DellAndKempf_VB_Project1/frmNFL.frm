VERSION 5.00
Begin VB.Form frmNFL 
   BackColor       =   &H0000C000&
   Caption         =   "NFL"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   5655
      Left            =   10080
      Picture         =   "frmNFL.frx":0000
      ScaleHeight     =   5595
      ScaleWidth      =   4395
      TabIndex        =   7
      Top             =   2880
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   11400
      Picture         =   "frmNFL.frx":7630
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   6
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Return to Main Menu"
      Height          =   1335
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   6960
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      Height          =   9255
      Left            =   3600
      ScaleHeight     =   9195
      ScaleWidth      =   5715
      TabIndex        =   4
      Top             =   120
      Width           =   5775
   End
   Begin VB.CommandButton cmdRB 
      Caption         =   "Rushing"
      Height          =   1335
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   2655
   End
   Begin VB.CommandButton cmdWR 
      Caption         =   "Receiving"
      Height          =   1335
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   4920
      Width           =   2655
   End
   Begin VB.CommandButton cmdQB 
      Caption         =   "Passing"
      Height          =   1335
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label lblStats 
      Alignment       =   2  'Center
      Caption         =   "Choose what Stats to sort by"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmNFL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project-NFL Stats
'Form-frmNFL
'Written by Ryan Kempf and Ryan Dell
'3-22-09
'This form sorts the entire NFL by a variety of statistical categories.
Private Sub cmdMain_Click(Index As Integer)
    frmNFL.Hide
    frmStartup.Show
End Sub

Private Sub cmdQB_Click(Index As Integer)
    Dim Search As String
    Dim Pass As Integer
    Dim pos As Integer
    Dim j As Integer
    Search = InputBox("Enter Yards, QB Rating, Completion Percentage, Yards/Attempt or TD to search statistics", "Sort By")
    If Search = "Yards" Then
        For Pass = 1 To ctrQ
            For pos = 1 To (ctrQ - Pass)
                If YardsQB(pos) < YardsQB(pos + 1) Then
                    temp1 = FirstNameQB(pos)
                    FirstNameQB(pos) = FirstNameQB(pos + 1)
                    FirstNameQB(pos + 1) = temp1
                    temp1 = LastNameQB(pos)
                    LastNameQB(pos) = LastNameQB(pos + 1)
                    LastNameQB(pos + 1) = temp1
                    temp1 = TeamQB(pos)
                    TeamQB(pos) = TeamQB(pos + 1)
                    TeamQB(pos + 1) = temp1
                    temp2 = YardsQB(pos)
                    YardsQB(pos) = YardsQB(pos + 1)
                    YardsQB(pos + 1) = temp2
                    temp3 = QBRating(pos)
                    QBRating(pos) = QBRating(pos + 1)
                    QBRating(pos + 1) = temp3
                    temp3 = CompPct(pos)
                    CompPct(pos) = CompPct(pos + 1)
                    CompPct(pos + 1) = temp3
                    temp3 = YdsAtt(pos)
                    YdsAtt(pos) = YdsAtt(pos + 1)
                    YdsAtt(pos + 1) = temp3
                    temp3 = PassTD(pos)
                    PassTD(pos) = PassTD(pos + 1)
                    PassTD(pos + 1) = temp3
                End If
            Next pos
        Next Pass
        picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "Yards"
        picResults.Print "************************************************************************************"
        MsgBox FirstNameQB(1) & " " & LastNameQB(1) & " has the most Passing Yards"
        For j = 1 To ctrQ
            picResults.Print FirstNameQB(j); " "; LastNameQB(j); Tab(28); TeamQB(j); Tab(50); YardsQB(j)
        Next j
    ElseIf Search = "QB Rating" Then
        For Pass = 1 To ctrQ
            For pos = 1 To (ctrQ - Pass)
                If QBRating(pos) < QBRating(pos + 1) Then
                    temp1 = FirstNameQB(pos)
                    FirstNameQB(pos) = FirstNameQB(pos + 1)
                    FirstNameQB(pos + 1) = temp1
                    temp1 = LastNameQB(pos)
                    LastNameQB(pos) = LastNameQB(pos + 1)
                    LastNameQB(pos + 1) = temp1
                    temp1 = TeamQB(pos)
                    TeamQB(pos) = TeamQB(pos + 1)
                    TeamQB(pos + 1) = temp1
                    temp3 = QBRating(pos)
                    QBRating(pos) = QBRating(pos + 1)
                    QBRating(pos + 1) = temp3
                    temp2 = YardsQB(pos)
                    YardsQB(pos) = YardsQB(pos + 1)
                    YardsQB(pos + 1) = temp2
                    temp3 = CompPct(pos)
                    CompPct(pos) = CompPct(pos + 1)
                    CompPct(pos + 1) = temp3
                    temp3 = YdsAtt(pos)
                    YdsAtt(pos) = YdsAtt(pos + 1)
                    YdsAtt(pos + 1) = temp3
                    temp3 = PassTD(pos)
                    PassTD(pos) = PassTD(pos + 1)
                    PassTD(pos + 1) = temp3
                End If
            Next pos
        Next Pass
        picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "QB Rating"
        picResults.Print "************************************************************************************"
        MsgBox FirstNameQB(1) & " " & LastNameQB(1) & " has the highest QB Rating"
        For j = 1 To ctrQ
            picResults.Print FirstNameQB(j); " "; LastNameQB(j); Tab(28); TeamQB(j); Tab(50); FormatNumber(QBRating(j), 3)
        Next j
    ElseIf Search = "Completion Percentage" Then
        For Pass = 1 To ctrQ
            For pos = 1 To (ctrQ - Pass)
                If CompPct(pos) < CompPct(pos + 1) Then
                    temp1 = FirstNameQB(pos)
                    FirstNameQB(pos) = FirstNameQB(pos + 1)
                    FirstNameQB(pos + 1) = temp1
                    temp1 = LastNameQB(pos)
                    LastNameQB(pos) = LastNameQB(pos + 1)
                    LastNameQB(pos + 1) = temp1
                    temp1 = TeamQB(pos)
                    TeamQB(pos) = TeamQB(pos + 1)
                    TeamQB(pos + 1) = temp1
                    temp3 = CompPct(pos)
                    CompPct(pos) = CompPct(pos + 1)
                    CompPct(pos + 1) = temp3
                    temp2 = YardsQB(pos)
                    YardsQB(pos) = YardsQB(pos + 1)
                    YardsQB(pos + 1) = temp2
                    temp3 = QBRating(pos)
                    QBRating(pos) = QBRating(pos + 1)
                    QBRating(pos + 1) = temp3
                    temp3 = YdsAtt(pos)
                    YdsAtt(pos) = YdsAtt(pos + 1)
                    YdsAtt(pos + 1) = temp3
                    temp3 = PassTD(pos)
                    PassTD(pos) = PassTD(pos + 1)
                    PassTD(pos + 1) = temp3
                End If
            Next pos
        Next Pass
         picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "Completion Percentage"
        picResults.Print "************************************************************************************"
        MsgBox FirstNameQB(1) & " " & LastNameQB(1) & " has the highest Completion Percentage"
        For j = 1 To ctrQ
            picResults.Print FirstNameQB(j); " "; LastNameQB(j); Tab(28); TeamQB(j); Tab(50); FormatPercent(CompPct(j))
        Next j
    ElseIf Search = "Yards/Attempt" Then
        For Pass = 1 To ctrQ
            For pos = 1 To (ctrQ - Pass)
                If YdsAtt(pos) < YdsAtt(pos + 1) Then
                    temp1 = FirstNameQB(pos)
                    FirstNameQB(pos) = FirstNameQB(pos + 1)
                    FirstNameQB(pos + 1) = temp1
                    temp1 = LastNameQB(pos)
                    LastNameQB(pos) = LastNameQB(pos + 1)
                    LastNameQB(pos + 1) = temp1
                    temp1 = TeamQB(pos)
                    TeamQB(pos) = TeamQB(pos + 1)
                    TeamQB(pos + 1) = temp1
                    temp3 = YdsAtt(pos)
                    YdsAtt(pos) = YdsAtt(pos + 1)
                    YdsAtt(pos + 1) = temp3
                    temp2 = YardsQB(pos)
                    YardsQB(pos) = YardsQB(pos + 1)
                    YardsQB(pos + 1) = temp2
                    temp3 = QBRating(pos)
                    QBRating(pos) = QBRating(pos + 1)
                    QBRating(pos + 1) = temp3
                    temp3 = CompPct(pos)
                    CompPct(pos) = CompPct(pos + 1)
                    CompPct(pos + 1) = temp3
                    temp3 = PassTD(pos)
                    PassTD(pos) = PassTD(pos + 1)
                    PassTD(pos + 1) = temp3
                End If
            Next pos
        Next Pass
         picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "Yards/Attempt"
        picResults.Print "************************************************************************************"
        MsgBox FirstNameQB(1) & " " & LastNameQB(1) & " has the most Passing Yards/Attempt"
        For j = 1 To ctrQ
            picResults.Print FirstNameQB(j); " "; LastNameQB(j); Tab(28); TeamQB(j); Tab(50); FormatNumber(YdsAtt(j), 2)
        Next j
    ElseIf Search = "TD" Then
        For Pass = 1 To ctrQ
            For pos = 1 To (ctrQ - Pass)
                If PassTD(pos) < PassTD(pos + 1) Then
                    temp1 = FirstNameQB(pos)
                    FirstNameQB(pos) = FirstNameQB(pos + 1)
                    FirstNameQB(pos + 1) = temp1
                    temp1 = LastNameQB(pos)
                    LastNameQB(pos) = LastNameQB(pos + 1)
                    LastNameQB(pos + 1) = temp1
                    temp1 = TeamQB(pos)
                    TeamQB(pos) = TeamQB(pos + 1)
                    TeamQB(pos + 1) = temp1
                    temp3 = PassTD(pos)
                    PassTD(pos) = PassTD(pos + 1)
                    PassTD(pos + 1) = temp3
                    temp2 = YardsQB(pos)
                    YardsQB(pos) = YardsQB(pos + 1)
                    YardsQB(pos + 1) = temp2
                    temp3 = QBRating(pos)
                    QBRating(pos) = QBRating(pos + 1)
                    QBRating(pos + 1) = temp3
                    temp3 = CompPct(pos)
                    CompPct(pos) = CompPct(pos + 1)
                    CompPct(pos + 1) = temp3
                    temp3 = YdsAtt(pos)
                    YdsAtt(pos) = YdsAtt(pos + 1)
                    YdsAtt(pos + 1) = temp3
                End If
            Next pos
        Next Pass
         picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "TD's"
        picResults.Print "************************************************************************************"
        MsgBox FirstNameQB(1) & " " & LastNameQB(1) & " has the most Passing Touchdowns"
        For j = 1 To ctrQ
            picResults.Print FirstNameQB(j); " "; LastNameQB(j); Tab(28); TeamQB(j); Tab(50); PassTD(j)
        Next j
    Else: MsgBox "Error-Check your Spelling", , "Error"
    End If
End Sub

Private Sub cmdRB_Click(Index As Integer)
    Dim Search As String
    Dim Pass As Integer
    Dim pos As Integer
    Dim j As Integer
    Search = InputBox("Enter Yards, Yards/Carry, or TD to search statistics", "Sort By")
    If Search = "Yards" Then
        For Pass = 1 To ctrR
            For pos = 1 To (ctrR - Pass)
                If YardsRB(pos) < YardsRB(pos + 1) Then
                    temp1 = FirstNameRB(pos)
                    FirstNameRB(pos) = FirstNameRB(pos + 1)
                    FirstNameRB(pos + 1) = temp1
                    temp1 = LastNameRB(pos)
                    LastNameRB(pos) = LastNameRB(pos + 1)
                    LastNameRB(pos + 1) = temp1
                    temp1 = TeamRB(pos)
                    TeamRB(pos) = TeamRB(pos + 1)
                    TeamRB(pos + 1) = temp1
                    temp2 = YardsRB(pos)
                    YardsRB(pos) = YardsRB(pos + 1)
                    YardsRB(pos + 1) = temp2
                    temp3 = YPCRB(pos)
                    YPCRB(pos) = YPCRB(pos + 1)
                    YPCRB(pos + 1) = temp3
                    temp2 = TDRB(pos)
                    TDRB(pos) = TDRB(pos + 1)
                    TDRB(pos + 1) = temp2
                End If
            Next pos
        Next Pass
        picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "Yards"
        picResults.Print "************************************************************************************"
        MsgBox FirstNameRB(1) & " " & LastNameRB(1) & " has the most Rushing Yards"
        For j = 1 To ctrR
            picResults.Print FirstNameRB(j); " "; LastNameRB(j); Tab(28); TeamRB(j); Tab(50); YardsRB(j)
        Next j
    ElseIf Search = "TD" Then
        For Pass = 1 To ctrR
            For pos = 1 To (ctrR - Pass)
                If TDRB(pos) < TDRB(pos + 1) Then
                    temp1 = FirstNameRB(pos)
                    FirstNameRB(pos) = FirstNameRB(pos + 1)
                    FirstNameRB(pos + 1) = temp1
                    temp1 = LastNameRB(pos)
                    LastNameRB(pos) = LastNameRB(pos + 1)
                    LastNameRB(pos + 1) = temp1
                    temp1 = TeamRB(pos)
                    TeamRB(pos) = TeamRB(pos + 1)
                    TeamRB(pos + 1) = temp1
                    temp2 = YardsRB(pos)
                    YardsRB(pos) = YardsRB(pos + 1)
                    YardsRB(pos + 1) = temp2
                    temp3 = YPCRB(pos)
                    YPCRB(pos) = YPCRB(pos + 1)
                    YPCRB(pos + 1) = temp3
                    temp2 = TDRB(pos)
                    TDRB(pos) = TDRB(pos + 1)
                    TDRB(pos + 1) = temp2
                End If
            Next pos
        Next Pass
        picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "TD"
        picResults.Print "************************************************************************************"
        MsgBox FirstNameRB(1) & " " & LastNameRB(1) & " has the most Rushing Touchdowns"
        For j = 1 To ctrR
            picResults.Print FirstNameRB(j); " "; LastNameRB(j); Tab(28); TeamRB(j); Tab(50); TDRB(j)
        Next j
    ElseIf Search = "Yards/Carry" Then
        For Pass = 1 To ctrR
            For pos = 1 To (ctrR - Pass)
                If YPCRB(pos) < YPCRB(pos + 1) Then
                    temp1 = FirstNameRB(pos)
                    FirstNameRB(pos) = FirstNameRB(pos + 1)
                    FirstNameRB(pos + 1) = temp1
                    temp1 = LastNameRB(pos)
                    LastNameRB(pos) = LastNameRB(pos + 1)
                    LastNameRB(pos + 1) = temp1
                    temp1 = TeamRB(pos)
                    TeamRB(pos) = TeamRB(pos + 1)
                    TeamRB(pos + 1) = temp1
                    temp2 = YardsRB(pos)
                    YardsRB(pos) = YardsRB(pos + 1)
                    YardsRB(pos + 1) = temp2
                    temp3 = YPCRB(pos)
                    YPCRB(pos) = YPCRB(pos + 1)
                    YPCRB(pos + 1) = temp3
                    temp2 = TDRB(pos)
                    TDRB(pos) = TDRB(pos + 1)
                    TDRB(pos + 1) = temp2
                End If
            Next pos
        Next Pass
        picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "Yards/Carry"
        picResults.Print "************************************************************************************"
        MsgBox FirstNameRB(1) & " " & LastNameRB(1) & " has the most Rushing Yards/Carry"
        For j = 1 To ctrR
            picResults.Print FirstNameRB(j); " "; LastNameRB(j); Tab(28); TeamRB(j); Tab(50); FormatNumber(YPCRB(j), 2)
        Next j
    Else: MsgBox "Error-Check your Spelling", , "Error"
    End If
End Sub

Private Sub cmdWR_Click(Index As Integer)
    Dim Search As String
    Dim Pass As Integer
    Dim pos As Integer
    Dim j As Integer
    Search = InputBox("Enter Yards, Yards/Reception, or TD to search statistics", "Sort By")
    If Search = "Yards" Then
        For Pass = 1 To ctrW
            For pos = 1 To (ctrW - Pass)
                If YardsWR(pos) < YardsWR(pos + 1) Then
                    temp1 = FirstNameWR(pos)
                    FirstNameWR(pos) = FirstNameWR(pos + 1)
                    FirstNameWR(pos + 1) = temp1
                    temp1 = LastNameWR(pos)
                    LastNameWR(pos) = LastNameWR(pos + 1)
                    LastNameWR(pos + 1) = temp1
                    temp1 = TeamWR(pos)
                    TeamWR(pos) = TeamWR(pos + 1)
                    TeamWR(pos + 1) = temp1
                    temp2 = YardsWR(pos)
                    YardsWR(pos) = YardsWR(pos + 1)
                    YardsWR(pos + 1) = temp2
                    temp3 = YPRWR(pos)
                    YPRWR(pos) = YPRWR(pos + 1)
                    YPRWR(pos + 1) = temp3
                    temp2 = TDWR(pos)
                    TDWR(pos) = TDWR(pos + 1)
                    TDWR(pos + 1) = temp2
                End If
            Next pos
        Next Pass
        picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "Yards"
        picResults.Print "************************************************************************************"
        MsgBox FirstNameWR(1) & " " & LastNameWR(1) & " has the most Receiving Yards"
        For j = 1 To ctrW
            picResults.Print FirstNameWR(j); " "; LastNameWR(j); Tab(28); TeamWR(j); Tab(50); YardsWR(j)
        Next j
    ElseIf Search = "Yards/Reception" Then
        For Pass = 1 To ctrW
            For pos = 1 To (ctrW - Pass)
                If YPRWR(pos) < YPRWR(pos + 1) Then
                    temp1 = FirstNameWR(pos)
                    FirstNameWR(pos) = FirstNameWR(pos + 1)
                    FirstNameWR(pos + 1) = temp1
                    temp1 = LastNameWR(pos)
                    LastNameWR(pos) = LastNameWR(pos + 1)
                    LastNameWR(pos + 1) = temp1
                    temp1 = TeamWR(pos)
                    TeamWR(pos) = TeamWR(pos + 1)
                    TeamWR(pos + 1) = temp1
                    temp2 = YardsWR(pos)
                    YardsWR(pos) = YardsWR(pos + 1)
                    YardsWR(pos + 1) = temp2
                    temp3 = YPRWR(pos)
                    YPRWR(pos) = YPRWR(pos + 1)
                    YPRWR(pos + 1) = temp3
                    temp2 = TDWR(pos)
                    TDWR(pos) = TDWR(pos + 1)
                    TDWR(pos + 1) = temp2
                End If
            Next pos
        Next Pass
        picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "Yards/Reception"
        picResults.Print "************************************************************************************"
        MsgBox FirstNameWR(1) & " " & LastNameWR(1) & " has the most Receiving Yards/Reception"
        For j = 1 To ctrW
            picResults.Print FirstNameWR(j); " "; LastNameWR(j); Tab(28); TeamWR(j); Tab(50); FormatNumber(YPRWR(j), 2)
        Next j
    ElseIf Search = "TD" Then
        For Pass = 1 To ctrW
            For pos = 1 To (ctrW - Pass)
                If TDWR(pos) < TDWR(pos + 1) Then
                    temp1 = FirstNameWR(pos)
                    FirstNameWR(pos) = FirstNameWR(pos + 1)
                    FirstNameWR(pos + 1) = temp1
                    temp1 = LastNameWR(pos)
                    LastNameWR(pos) = LastNameWR(pos + 1)
                    LastNameWR(pos + 1) = temp1
                    temp1 = TeamWR(pos)
                    TeamWR(pos) = TeamWR(pos + 1)
                    TeamWR(pos + 1) = temp1
                    temp2 = YardsWR(pos)
                    YardsWR(pos) = YardsWR(pos + 1)
                    YardsWR(pos + 1) = temp2
                    temp3 = YPRWR(pos)
                    YPRWR(pos) = YPRWR(pos + 1)
                    YPRWR(pos + 1) = temp3
                    temp2 = TDWR(pos)
                    TDWR(pos) = TDWR(pos + 1)
                    TDWR(pos + 1) = temp2
                End If
            Next pos
        Next Pass
        picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "TD"
        picResults.Print "************************************************************************************"
        MsgBox FirstNameWR(1) & " " & LastNameWR(1) & " has the most Receiving Touchdowns"
        For j = 1 To ctrW
            picResults.Print FirstNameWR(j); " "; LastNameWR(j); Tab(28); TeamWR(j); Tab(50); TDWR(j)
        Next j
    Else: MsgBox "Error-Check your Spelling", , "Error"
    End If
End Sub
