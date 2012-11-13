VERSION 5.00
Begin VB.Form FrmFielders 
   BackColor       =   &H000000FF&
   Caption         =   "Minnesota Twins Fielders"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   3480
      Picture         =   "FrmFielders.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   5955
      TabIndex        =   13
      Top             =   3120
      Width           =   6015
   End
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H00FF0000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8040
      Width           =   2055
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to stats page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton CmdQAvg 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Am I right again?"
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7440
      Width           =   2535
   End
   Begin VB.CommandButton CmdQRBI 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check my answer"
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6960
      Width           =   2895
   End
   Begin VB.TextBox TxtBattingAvg 
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   7440
      Width           =   2895
   End
   Begin VB.TextBox TxtRBI 
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   6960
      Width           =   2895
   End
   Begin VB.CommandButton CmdAvg 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Average Batting Average of the  Minesota Twins"
      DisabledPicture =   "FrmFielders.frx":219A
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   1935
   End
   Begin VB.PictureBox PicResults 
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   2160
      ScaleHeight     =   2475
      ScaleWidth      =   8595
      TabIndex        =   3
      Top             =   480
      Width           =   8655
   End
   Begin VB.CommandButton CmdHomeruns 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rank Players By Home Runs"
      DisabledPicture =   "FrmFielders.frx":35C3
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton CmdBestBattingAvg 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rank Players by Batting Average"
      DisabledPicture =   "FrmFielders.frx":49EC
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton CmdLoad 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "What is considered to be a good Batting Average?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   7440
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "What does RBI stand for?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   6960
      Width           =   2775
   End
End
Attribute VB_Name = "FrmFielders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Minnesota Twins
'FrmFielders
'Sarah Mayer and Jake Klis
'Written on 03/22/09
Option Explicit
Dim Fielders(1 To 100) As String, Positions(1 To 100) As String
Dim BattingAvg(1 To 100) As Single, HomeRuns(1 To 100) As Integer
Dim RBI(1 To 100) As Integer, Ctr As Integer
' This button finds the average batting average of the Minnesota Twins
Private Sub CmdAvg_Click()
Dim K As Integer, Avg As Single, Sum As Single
Sum = 0
For K = 1 To Ctr
    Sum = Sum + BattingAvg(K)
Next K
    Avg = Sum / Ctr
MsgBox ("The average batting average for the Minnesota Twins is " & FormatNumber(Avg, 3))
End Sub
'This button sorts the fielders based on their batting averages
Private Sub CmdBestBattingAvg_Click()
Dim Pass As Integer, Pos As Integer, TempName As String, TempAvg As Single
Dim J As Integer, TempRBI As Integer, TempHR As Integer
PicResults.Cls
    For Pass = 1 To Ctr - 1
        For Pos = 1 To Ctr - Pass
            If BattingAvg(Pos) < BattingAvg(Pos + 1) Then
            TempAvg = BattingAvg(Pos)
            BattingAvg(Pos) = BattingAvg(Pos + 1)
            BattingAvg(Pos + 1) = TempAvg
            TempHR = HomeRuns(Pos)
            HomeRuns(Pos) = HomeRuns(Pos + 1)
            HomeRuns(Pos + 1) = TempHR
            TempName = Fielders(Pos)
            Fielders(Pos) = Fielders(Pos + 1)
            Fielders(Pos + 1) = TempName
            TempRBI = RBI(Pos)
            RBI(Pos) = RBI(Pos + 1)
            RBI(Pos + 1) = TempRBI
            End If
        Next Pos
    Next Pass
PicResults.Print "Names", Tab(30); "Batting Average"
PicResults.Print "-------------------------------------------------------------------------"
For J = 1 To Ctr
    PicResults.Print Fielders(J), Tab(30); FormatNumber(BattingAvg(J), 3)
Next J
End Sub
' This button sorts the fielders by the number of homeruns that they have hit
Private Sub CmdHomeruns_Click()
Dim Pass As Integer, Pos As Integer, TempHR As Integer, TempName As String, TempRBI As Integer
Dim K As Integer, TempAvg As Single
PicResults.Cls
For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If HomeRuns(Pos + 1) > HomeRuns(Pos) Then
            TempHR = HomeRuns(Pos)
            HomeRuns(Pos) = HomeRuns(Pos + 1)
            HomeRuns(Pos + 1) = TempHR
            TempName = Fielders(Pos)
            Fielders(Pos) = Fielders(Pos + 1)
            Fielders(Pos + 1) = TempName
            TempRBI = RBI(Pos)
            RBI(Pos) = RBI(Pos + 1)
            RBI(Pos + 1) = TempRBI
            TempAvg = BattingAvg(Pos)
            BattingAvg(Pos) = BattingAvg(Pos + 1)
            BattingAvg(Pos + 1) = TempAvg
        End If
    Next Pos
Next Pass
PicResults.Print "Name", Tab(30); "Home Runs"
PicResults.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
For K = 1 To Ctr
    PicResults.Print Fielders(K), Tab(30); HomeRuns(K)
Next K

End Sub
'This button loads the twinsfielders.txt file into parallel arrays
Private Sub CmdLoad_Click()
Dim K As Integer
Ctr = 0
Open App.Path & "\twinsfielders.txt" For Input As #2
Do While Not EOF(2)
    Ctr = Ctr + 1
    Input #2, Fielders(Ctr), Positions(Ctr), BattingAvg(Ctr), HomeRuns(Ctr), RBI(Ctr)
Loop
PicResults.Print "Fielders Name", Tab(30); "Position Played", Tab(60); "Batting Average", Tab(85); "Home Runs", Tab(100); "RBI"
PicResults.Print "--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
For K = 1 To Ctr
PicResults.Print Fielders(K), Tab(30); Positions(K), Tab(60); FormatNumber(BattingAvg(K), 3); Tab(85); HomeRuns(K); Tab(100); RBI(K)
Next K
CmdBestBattingAvg.Enabled = True
CmdHomeruns.Enabled = True
CmdAvg.Enabled = True
CmdLoad.Enabled = False

Close #2

End Sub
'This button sees if the the input from the user matches the correct answer. It does this using
'select case statements to see how close the user can get to the correct answer
'If it is correct the Trivia counter is incremented by one.
Private Sub CmdQAvg_Click()
Dim Guess As Single
Guess = TxtBattingAvg
Select Case Guess
    Case Is >= 1
        MsgBox ("That Number isn't even a possibility. It should be between 0 and 1")
    Case 0.4 To 0.99
         MsgBox ("Wow you have extremely high standards! The correct answer is .300")
    Case 0.301 To 0.39
        MsgBox ("So close, the answer we were looking for is .300")
    Case 0.3
        MsgBox ("You are correct! You are so knowledgeable in the field of baseball!")
        TriviaCtr = TriviaCtr + 1
    Case 0.25 To 0.29
        MsgBox ("That average is OKAY, but a GOOD average is .300!")
    Case 0 To 0.25
        MsgBox ("C'mon this is a PROFESSIONAL baseball league! The correct answer is .300")
    End Select
    
End Sub

'This button sees if the the input from the user matches the correct answer. If it is correct
'the Trivia counter is incremented by one.
Private Sub CmdQRBI_Click()
Dim TextBox As String
TextBox = TxtRBI
If TextBox = "Runs Batted In" Then
    MsgBox ("You are correct! Keep up the good work")
    TriviaCtr = TriviaCtr + 1
    Else
    MsgBox ("Come on that was supposed to be a gimme! The correct answer is Runs Batted In!")
End If


End Sub

Private Sub Cmdback_Click()
frmStats.Show
FrmFielders.Hide
End Sub

Private Sub CmdQuit_Click()
MsgBox "You got " & TriviaCtr & " answers correct out of 5 possible", , "Good Job!"
End
End Sub
