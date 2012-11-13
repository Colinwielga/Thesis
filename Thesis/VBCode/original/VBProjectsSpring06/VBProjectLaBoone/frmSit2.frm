VERSION 5.00
Begin VB.Form frmSit2 
   BackColor       =   &H000000FF&
   Caption         =   "Situation 2"
   ClientHeight    =   5370
   ClientLeft      =   3540
   ClientTop       =   1080
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "Myriad Pro"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmSit2.frx":0000
   ScaleHeight     =   5370
   ScaleWidth      =   8205
   Begin VB.CommandButton cmdShow 
      Caption         =   "Click First to Show Scores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   3
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton cmdResults 
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to main page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   5160
      ScaleHeight     =   3195
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblExplination 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   $"frmSit2.frx":BE96
      BeginProperty Font 
         Name            =   "Franklin Gothic Demi"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      TabIndex        =   4
      Top             =   4320
      Width           =   6615
   End
End
Attribute VB_Name = "frmSit2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Private Sub cmdResults_Click()
    Dim fifth, forth, third, second, first As String
    I = 5
    Do While fifth <> Names(I)
        fifth = InputBox("What is the name of the gymnasts in 5th place?", "Fifth")
        Select Case fifth
            Case Names(I)
                MsgBox "Good Job.", , "Correct"
            Case Else
                MsgBox "Try again.", , "Error"
        End Select
    Loop
    I = I - 1
    Do While forth <> Names(I)
        forth = InputBox("What is the name of the gymnasts in 4th place?", "Forth")
        Select Case forth
            Case Names(I)
                MsgBox "Good Job.", , "Correct"
            Case Else
                MsgBox "Try again.", , "Error"
        End Select
    Loop
    I = I - 1
    Do While third <> Names(I)
        third = InputBox("What is the name of the gymnasts in 3rd place?", "Third")
        Select Case third
            Case Names(I)
                MsgBox "Good Job.", , "Correct"
            Case Else
                MsgBox "Try again.", , "Error"
        End Select
    Loop
    I = I - 1
    Do While second <> Names(I)
        second = InputBox("What is the name of the gymnasts in 2nd place?", "Second")
        Select Case second
            Case Names(I)
                MsgBox "Good Job.", , "Correct"
            Case Else
                MsgBox "Try again.", , "Error"
        End Select
    Loop
    I = I - 1
    Do While first <> Names(I)
        first = InputBox("What is the name of the gymnasts in 1st place?", "First")
        Select Case first
            Case Names(I)
                MsgBox "Good Job.", , "Correct"
            Case Else
                MsgBox "Try again.", , "Error"
        End Select
    Loop
    MsgBox "Nice job!", , "Finished"
    picResults.Cls
    picResults.Print "This is the final scores"
    picResults.Print "Rank", "Name", "Score"
    For I = 1 To 10
        picResults.Print I, Names(I), Scores(I)
    Next I
End Sub

Private Sub cmdReturn_Click()
    frmSit2.Hide
    frmIntro.Show
End Sub

Private Sub cmdShow_Click()
    Dim I, Size, Pass As Integer
    I = 0
    Open App.Path & "\scores.txt" For Input As #1
    picResults.Print "Name", "Score", "SJ Score"
    Do Until EOF(1)
        I = I + 1
        Input #1, Names(I), Scores(I), SJ(I)
        picResults.Print Names(I), FormatNumber(Scores(I), 1), FormatNumber(SJ(I), 1)
    Loop
    Close #1
    Size = I
    For Pass = 1 To (Size - 1)
        For I = 1 To (Size - Pass)
            If Scores(I) < Scores(I + 1) Then
                TempScores = Scores(I)
                Scores(I) = Scores(I + 1)
                Scores(I + 1) = TempScores
                TempNames = Names(I)
                Names(I) = Names(I + 1)
                Names(I + 1) = TempNames
                TempSJ = SJ(I)
                SJ(I) = SJ(I + 1)
                SJ(I + 1) = TempSJ
            ElseIf Scores(I) = Scores(I + 1) Then
                If SJ(I) < SJ(I + 1) Then
                    TempScores = Scores(I)
                    Scores(I) = Scores(I + 1)
                    Scores(I + 1) = TempScores
                    TempNames = Names(I)
                    Names(I) = Names(I + 1)
                    Names(I + 1) = TempNames
                    TempSJ = SJ(I)
                    SJ(I) = SJ(I + 1)
                    SJ(I + 1) = TempSJ
                End If
            End If
        Next I
    Next Pass
End Sub

