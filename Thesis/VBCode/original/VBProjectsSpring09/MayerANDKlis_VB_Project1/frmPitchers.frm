VERSION 5.00
Begin VB.Form frmPitchers 
   BackColor       =   &H00C00000&
   Caption         =   "Minnesota Twins Pitchers"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   11490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6120
      Width           =   2655
   End
   Begin VB.CommandButton CmdCheck 
      BackColor       =   &H000000FF&
      Caption         =   "Check this answer!"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7200
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2760
      TabIndex        =   8
      Top             =   7200
      Width           =   2415
   End
   Begin VB.CommandButton CmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go back to the Statistics Page"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7560
      Width           =   2655
   End
   Begin VB.CommandButton CmdAnswer 
      BackColor       =   &H000000FF&
      Caption         =   "Check My Answer Here!"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   2295
   End
   Begin VB.TextBox TxtAnswer 
      Height          =   615
      Left            =   2760
      TabIndex        =   4
      Top             =   6360
      Width           =   2415
   End
   Begin VB.CommandButton CmdMostWins 
      BackColor       =   &H000000FF&
      Caption         =   "Sort by Wins"
      DisabledPicture =   "frmPitchers.frx":0000
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton CmdBestERA 
      BackColor       =   &H000000FF&
      Caption         =   "Sort the Twins Pitchers by Their Best ERA"
      DisabledPicture =   "frmPitchers.frx":1429
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
   End
   Begin VB.PictureBox PicResults 
      Height          =   3255
      Left            =   2760
      ScaleHeight     =   3195
      ScaleWidth      =   8235
      TabIndex        =   1
      Top             =   240
      Width           =   8295
   End
   Begin VB.CommandButton CmdLoad 
      BackColor       =   &H000000FF&
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lbl1 
      BackColor       =   &H000000FF&
      Caption         =   "Why does Joe Nathan have so few wins compared to everyone else? He is the _______?"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   8.25
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   7200
      Width           =   2655
   End
   Begin VB.Label Caption 
      BackColor       =   &H000000FF&
      Caption         =   "What does ERA stand for?"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   6360
      Width           =   2535
   End
End
Attribute VB_Name = "frmPitchers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Minnesota Twins
'FrmPitchers
'Sarah Mayer and Jake Klis
'Written on 03/22/09
'This form is used to load the statistical data on the Twins pitchers into arrays,
'and perform various sorting functions dealing with ranking the pitchers as well as
'asking two more trivia questions

Option Explicit
Dim Pitchers(1 To 100) As String, ERA(1 To 100) As Single, Wins(1 To 100) As Integer
Dim Strikeouts(1 To 100) As Integer
'This button sees if the the input from the user matches the correct answer. If it is correct
'the Trivia counter is incremented by one.
Private Sub CmdAnswer_Click()
Dim TextBox As String
TextBox = TxtAnswer
If TextBox = "Earned Run Average" Then
    MsgBox ("You know your baseball facts!")
    TriviaCtr = TriviaCtr + 1
    Else
    MsgBox ("Brush up on your baseball knowledge! It means Earned Run Average")
    End If


End Sub
' This button sorts the array by the lowest ERA, because you want a low ERA to be successful
Private Sub CmdBestERA_Click()
Dim Pass As Integer, Pos As Integer, TempERA As Single, TempName As String, I As Integer, TempStrikeouts As Integer, TempWins As Integer
PicResults.Cls

For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If ERA(Pos + 1) < ERA(Pos) Then
            TempERA = ERA(Pos)
            ERA(Pos) = ERA(Pos + 1)
            ERA(Pos + 1) = TempERA
            
            TempName = Pitchers(Pos)
            Pitchers(Pos) = Pitchers(Pos + 1)
            Pitchers(Pos + 1) = TempName
            
            TempStrikeouts = Strikeouts(Pos)
            Strikeouts(Pos) = Strikeouts(Pos + 1)
            Strikeouts(Pos + 1) = TempStrikeouts
            
            TempWins = Wins(Pos)
            Wins(Pos) = Wins(Pos + 1)
            Wins(Pos + 1) = TempWins
        End If
    Next Pos
Next Pass
PicResults.Print "Pitchers Names", Tab(30); "ERA"
PicResults.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
For I = 1 To Ctr
PicResults.Print Pitchers(I), Tab(30); ERA(I)
Next I
End Sub
'This button sees if the the input from the user matches the correct answer. If it is correct
'the Trivia counter is incremented by one.
Private Sub CmdCheck_Click()
Dim Answer As String
Answer = Text1
If Answer = "Closer" Then
MsgBox ("You are correct!")
TriviaCtr = TriviaCtr + 1
Else
MsgBox ("I'm sorry but the answer is closer")
End If
End Sub
'This button loads the data into parallel arrays
Private Sub CmdLoad_Click()
Dim K As Integer
Open App.Path & "\twinspitchers.txt" For Input As #1
Ctr = 0
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Pitchers(Ctr), ERA(Ctr), Wins(Ctr), Strikeouts(Ctr)
Loop
PicResults.Print "Pitchers Names", Tab(30); "ERA", Tab(45); "Wins", Tab(60); "Strikeouts"
PicResults.Print "--------------------------------------------------------------------------------------------------------------------"
For K = 1 To Ctr
PicResults.Print Pitchers(K), Tab(30); ERA(K), Tab(45); Wins(K), Tab(60); Strikeouts(K)

Next K
CmdBestERA.Enabled = True
CmdMostWins.Enabled = True
CmdLoad.Enabled = False

Close #1

End Sub
' This button sorts the pitchers by most wins
Private Sub CmdMostWins_Click()
Dim Pass As Integer, Pos As Integer, TempWins As Integer, TempName As String, J As Integer, TempERA As Single
Dim TempStrikeouts As Integer
PicResults.Cls

For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Wins(Pos + 1) > Wins(Pos) Then
            TempERA = ERA(Pos)
            ERA(Pos) = ERA(Pos + 1)
            ERA(Pos + 1) = TempERA
            
            TempName = Pitchers(Pos)
            Pitchers(Pos) = Pitchers(Pos + 1)
            Pitchers(Pos + 1) = TempName
            
            TempStrikeouts = Strikeouts(Pos)
            Strikeouts(Pos) = Strikeouts(Pos + 1)
            Strikeouts(Pos + 1) = TempStrikeouts
            
            TempWins = Wins(Pos)
            Wins(Pos) = Wins(Pos + 1)
            Wins(Pos + 1) = TempWins
        End If
    Next Pos
Next Pass
PicResults.Print "Pitchers Names", Tab(30); "Games Won"
PicResults.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
For J = 1 To Ctr
PicResults.Print Pitchers(J), Tab(30); Wins(J)
Next J
End Sub

Private Sub CmdQuit_Click()
MsgBox "You got " & TriviaCtr & " answers correct out of 5 possible", , "Good Job!"
End
End Sub

Private Sub Cmdreturn_Click()
frmPitchers.Hide
frmStats.Show
End Sub
