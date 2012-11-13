VERSION 5.00
Begin VB.Form frmAFC 
   BackColor       =   &H000000FF&
   Caption         =   "AFC"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   6375
      Left            =   9720
      Picture         =   "frmAFC.frx":0000
      ScaleHeight     =   6315
      ScaleWidth      =   5355
      TabIndex        =   7
      Top             =   2160
      Width           =   5415
   End
   Begin VB.PictureBox Picture3 
      Height          =   975
      Left            =   11400
      Picture         =   "frmAFC.frx":14315
      ScaleHeight     =   915
      ScaleWidth      =   1875
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdQB 
      Caption         =   "Passing"
      Height          =   1335
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton cmdWR 
      Caption         =   "Receiving"
      Height          =   1335
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton cmdRB 
      Caption         =   "Rushing"
      Height          =   1335
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      Height          =   9255
      Left            =   3360
      ScaleHeight     =   9195
      ScaleWidth      =   5715
      TabIndex        =   1
      Top             =   240
      Width           =   5775
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Return to Main Menu"
      Height          =   1335
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Label lblStats 
      Alignment       =   2  'Center
      Caption         =   "Choose what Stats to sort by"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmAFC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project-NFL Stats
'Form-frmAFC
'Written by Ryan Kempf and Ryan Dell
'3-22-09
'This form sorts a specific conference (AFC) by a variety of statistical categories.
Private Sub cmdMain_Click(Index As Integer)
    frmAFC.Hide
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
                    temp1 = DivisionQB(pos)
                    DivisionQB(pos) = DivisionQB(pos + 1)
                    DivisionQB(pos + 1) = temp1
                End If
            Next pos
        Next Pass
        picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "Yards"
        picResults.Print "************************************************************************************"
        For j = 1 To ctrQ
            If DivisionQB(j) = "A" Then
                picResults.Print FirstNameQB(j); " "; LastNameQB(j); Tab(28); TeamQB(j); Tab(50); YardsQB(j)
            End If
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
                     temp1 = DivisionQB(pos)
                    DivisionQB(pos) = DivisionQB(pos + 1)
                    DivisionQB(pos + 1) = temp1
                End If
            Next pos
        Next Pass
        picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "QB Rating"
        picResults.Print "************************************************************************************"
        For j = 1 To ctrQ
            If DivisionQB(j) = "A" Then
            picResults.Print FirstNameQB(j); " "; LastNameQB(j); Tab(28); TeamQB(j); Tab(50); FormatNumber(QBRating(j), 3)
            End If
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
                     temp1 = DivisionQB(pos)
                    DivisionQB(pos) = DivisionQB(pos + 1)
                    DivisionQB(pos + 1) = temp1
                End If
            Next pos
        Next Pass
         picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "Completion Percentage"
        picResults.Print "************************************************************************************"
        For j = 1 To ctrQ
            If DivisionQB(j) = "A" Then
                picResults.Print FirstNameQB(j); " "; LastNameQB(j); Tab(28); TeamQB(j); Tab(50); FormatPercent(CompPct(j))
            End If
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
                     temp1 = DivisionQB(pos)
                    DivisionQB(pos) = DivisionQB(pos + 1)
                    DivisionQB(pos + 1) = temp1
                End If
            Next pos
        Next Pass
         picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "Yards/Attempt"
        picResults.Print "************************************************************************************"
        For j = 1 To ctrQ
            If DivisionQB(j) = "A" Then
                picResults.Print FirstNameQB(j); " "; LastNameQB(j); Tab(28); TeamQB(j); Tab(50); FormatNumber(YdsAtt(j), 2)
            End If
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
                     temp1 = DivisionQB(pos)
                    DivisionQB(pos) = DivisionQB(pos + 1)
                    DivisionQB(pos + 1) = temp1
                End If
            Next pos
        Next Pass
         picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "TD's"
        picResults.Print "************************************************************************************"
        For j = 1 To ctrQ
            If DivisionQB(j) = "A" Then
                picResults.Print FirstNameQB(j); " "; LastNameQB(j); Tab(28); TeamQB(j); Tab(50); PassTD(j)
            End If
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
                     temp1 = DivisionRB(pos)
                    DivisionRB(pos) = DivisionRB(pos + 1)
                    DivisionRB(pos + 1) = temp1
                End If
            Next pos
        Next Pass
        picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "Yards"
        picResults.Print "************************************************************************************"
        For j = 1 To ctrR
            If DivisionRB(j) = "A" Then
                picResults.Print FirstNameRB(j); " "; LastNameRB(j); Tab(28); TeamRB(j); Tab(50); YardsRB(j)
            End If
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
                     temp1 = DivisionRB(pos)
                    DivisionRB(pos) = DivisionRB(pos + 1)
                    DivisionRB(pos + 1) = temp1
                End If
            Next pos
        Next Pass
        picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "TD"
        picResults.Print "************************************************************************************"
        For j = 1 To ctrR
            If DivisionRB(j) = "A" Then
                picResults.Print FirstNameRB(j); " "; LastNameRB(j); Tab(28); TeamRB(j); Tab(50); TDRB(j)
            End If
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
                     temp1 = DivisionRB(pos)
                    DivisionRB(pos) = DivisionRB(pos + 1)
                    DivisionRB(pos + 1) = temp1
                End If
            Next pos
        Next Pass
        picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "Yards/Carry"
        picResults.Print "************************************************************************************"
        For j = 1 To ctrR
            If DivisionRB(j) = "A" Then
                picResults.Print FirstNameRB(j); " "; LastNameRB(j); Tab(28); TeamRB(j); Tab(50); FormatNumber(YPCRB(j), 2)
            End If
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
                     temp1 = DivisionWR(pos)
                    DivisionWR(pos) = DivisionWR(pos + 1)
                    DivisionWR(pos + 1) = temp1
                End If
            Next pos
        Next Pass
        picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "Yards"
        picResults.Print "************************************************************************************"
        For j = 1 To ctrW
            If DivisionWR(j) = "A" Then
                picResults.Print FirstNameWR(j); " "; LastNameWR(j); Tab(28); TeamWR(j); Tab(50); YardsWR(j)
            End If
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
                      temp1 = DivisionWR(pos)
                    DivisionWR(pos) = DivisionWR(pos + 1)
                    DivisionWR(pos + 1) = temp1
                End If
            Next pos
        Next Pass
        picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "Yards/Reception"
        picResults.Print "************************************************************************************"
        For j = 1 To ctrW
            If DivisionWR(j) = "A" Then
                picResults.Print FirstNameWR(j); " "; LastNameWR(j); Tab(28); TeamWR(j); Tab(50); FormatNumber(YPRWR(j), 2)
            End If
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
                    temp1 = DivisionWR(pos)
                    DivisionWR(pos) = DivisionWR(pos + 1)
                    DivisionWR(pos + 1) = temp1
                End If
            Next pos
        Next Pass
        picResults.Cls
        picResults.Print "Name"; Tab(28); "Team"; Tab(50); "TD"
        picResults.Print "************************************************************************************"
        For j = 1 To ctrW
            If DivisionWR(j) = "A" Then
                picResults.Print FirstNameWR(j); " "; LastNameWR(j); Tab(28); TeamWR(j); Tab(50); TDWR(j)
            End If
        Next j
    Else: MsgBox "Error-Check your Spelling", , "Error"
    End If
End Sub
