VERSION 5.00
Begin VB.Form frmGames 
   BackColor       =   &H00FFC0C0&
   Caption         =   "GAMES GALORE"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9975
   FillColor       =   &H00FFFF80&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   Picture         =   "game.frx":0000
   ScaleHeight     =   7185
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdshow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "START HERE, CLICK ME!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   0
      Width           =   2415
   End
   Begin VB.PictureBox picWelcome 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      Height          =   615
      Left            =   840
      ScaleHeight     =   555
      ScaleWidth      =   4875
      TabIndex        =   2
      Top             =   840
      Width           =   4935
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdplay 
      BackColor       =   &H000000FF&
      Caption         =   "NOW PLAY THE GAME!!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6960
      MaskColor       =   &H0000C000&
      TabIndex        =   0
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Label lblnames 
      BackStyle       =   0  'Transparent
      Caption         =   "by Lisa Hammer and Kate Bancks"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   6720
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   7155
      Left            =   0
      Picture         =   "game.frx":263BD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9960
   End
End
Attribute VB_Name = "frmGames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdplay_Click()                     'the purpose of frmGames is to ask the user 10 simple questions from two parallel arrays, while keeping a running total or correct answers.
    Dim QuestionsArray(1 To 10) As String       'this is level one of the game.
    Dim AnswersArray(1 To 10) As String         'this button reads a file into the parallel arrays and compares the users' input with the actual answer part of the array.
    Dim pos, count As Integer                   'this button also allows the user to switch forms and receive feedback
    Dim x As String
    pos = 0
    count = 0

Open App.Path & "\Game.txt" For Input As #1
    Do Until EOF(1)
        pos = pos + 1
        Input #1, QuestionsArray(pos), AnswersArray(pos)
    Loop
Close #1
    pos = pos + 1
    MsgBox "Please Use Lower-Case One Word Answers", , "Rules"
        For pos = 1 To 10
            x = InputBox(QuestionsArray(pos), "Question")
            If x = AnswersArray(pos) Then
                count = count + 1
            End If
        Next pos
    MsgBox "The Number of Questions Correct is:" & count, , "GREAT JOB!"
    MsgBox "You Have Made it to Level 2", , "Level 2"
    C = count
    frmGames.Hide
    frmGames2.Show
    
End Sub

Private Sub cmdquit_Click()                     'the user is allowed to end the program
    End
End Sub

Private Sub cmdShow_Click()                     'this button prints out a welcome to the user
   picWelcome.Print "Welcome To Circus Fun, " & N & "!"
End Sub
