VERSION 5.00
Begin VB.Form frmQuiz 
   Caption         =   "Project Quiz"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6405
   LinkTopic       =   "Form2"
   ScaleHeight     =   5460
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "defeat"
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   5040
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Victory!"
      Height          =   495
      Left            =   2520
      TabIndex        =   15
      Top             =   4920
      Width           =   1215
   End
   Begin VB.PictureBox pbxScore 
      Height          =   495
      Left            =   5160
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   12
      Top             =   3720
      Width           =   1095
   End
   Begin VB.PictureBox pbxResult 
      Height          =   1215
      Left            =   3960
      ScaleHeight     =   1155
      ScaleWidth      =   1035
      TabIndex        =   11
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Input Guess"
      Height          =   975
      Left            =   2520
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdHint 
      Caption         =   "Hint: Limit of three and all cost points."
      Height          =   615
      Left            =   2520
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
   End
   Begin VB.PictureBox pbxQuiz 
      Height          =   1215
      Left            =   3960
      ScaleHeight     =   1155
      ScaleWidth      =   2235
      TabIndex        =   8
      Top             =   2400
      Width           =   2295
   End
   Begin VB.OptionButton Option3 
      Caption         =   "3"
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   1920
      Width           =   375
   End
   Begin VB.OptionButton Option2 
      Caption         =   "2"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   1920
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1"
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox pbxC3 
      Height          =   1575
      Left            =   4440
      ScaleHeight     =   1515
      ScaleWidth      =   1755
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox pbxC2 
      Height          =   1575
      Left            =   2280
      ScaleHeight     =   1515
      ScaleWidth      =   1755
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox pbxC1 
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1515
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start the Quiz/Next Question"
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   120
      Picture         =   "Quiz.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Designed by Chris Davin"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label scorebox 
      Caption         =   "Current Score       Is Above"
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   4440
      Width           =   1095
   End
End
Attribute VB_Name = "frmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : MemoryGamesEtc (Chris Davin's VB Project.vbp)
'Form Name : frmQuiz (Quiz.frm)
'Author: Chris Davin
'Date Written: October 29, 2003
'Purpose of Form: To Quiz the player/user on all they did in the program
                'See how many details the player spoted.  Keeps
                'track of the score and has hints.

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Dim Pichu As String, Bulbasaur As String, Vulpix As String, SevenBeasts As String
Dim SixBeasts As String, EightBeasts As String, Hitmonlee As String
Dim Red As String, Green As String, Yellow As String, FourBeasts As String
Dim FiveBeasts As String, Hamtaro As String, Fish As String, Ein As String
Dim Question As Integer, Hint As Integer
Dim Score As Double
'Depending on the option button selected will respond
'saying wheather the choice was correct and updating
'the score.  Also allows player to move on to next question
'by enabling the next question button
Private Sub cmdChoice_Click()
    Option1.Visible = False
    Option2.Visible = False
    Option3.Visible = False
    cmdChoice.Enabled = False
    cmdStart.Enabled = True
    pbxResult.Cls
    Select Case Question
        Case Is = 1
            If Option1 = True Then
                    pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option1 = False
                ElseIf Option2 = True Then
                    pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option2 = False
                ElseIf Option3 = True Then
                    pbxResult.Print "Correct"
                    Score = Score + 1
                    pbxScore.Cls
                    pbxScore.Print Score
                    Option3 = False
            End If
        Case Is = 2
            If Option1 = True Then
                    pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option1 = False
                ElseIf Option2 = True Then
                    pbxResult.Print "Correct"
                    Score = Score + 1
                    pbxScore.Cls
                    Option2 = False
                    pbxScore.Print Score
                ElseIf Option3 = True Then
                    pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option3 = False
            End If
        Case Is = 3
            If Option1 = True Then
                    pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option1 = False
                ElseIf Option2 = True Then
                    pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option2 = False
                ElseIf Option3 = True Then
                    pbxResult.Print "Correct"
                    Score = Score + 1
                    pbxScore.Cls
                    Option3 = False
                    pbxScore.Print Score
                End If
        Case Is = 4
            If Option1 = True Then
                    pbxResult.Print "Correct"
                    Score = Score + 1
                    pbxScore.Cls
                    Option1 = False
                    pbxScore.Print Score
                ElseIf Option2 = True Then
                    pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option2 = False
                ElseIf Option3 = True Then
                    pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option3 = False
                End If
        Case Is = 5
            If Option1 = True Then
                    pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option1 = False
                ElseIf Option2 = True Then
                    pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option2 = False
                ElseIf Option3 = True Then
                    pbxResult.Print "Correct"
                    Score = Score + 1
                    pbxScore.Cls
                    pbxScore.Print Score
                    Option3 = False
            End If
        Case Is = 6
            If Option1 = True Then
                    pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option1 = False
                ElseIf Option2 = True Then
                    pbxResult.Print "Correct"
                    pbxScore.Cls
                    Score = Score + 1
                    pbxScore.Print Score
                    Option2 = False
                ElseIf Option3 = True Then
                    pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option3 = False
            End If
        Case Is = 7
            If Option1 = True Then
                    pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option1 = False
                ElseIf Option2 = True Then
                    pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option2 = False
                ElseIf Option3 = True Then
                    pbxResult.Print "Correct"
                    Score = Score + 1
                    pbxScore.Cls
                    pbxScore.Print Score
                    Option3 = False
            End If
        Case Is = 8
           If Option1 = True Then
                    pbxResult.Print "Correct"
                    Score = Score + 1
                    pbxScore.Cls
                    Option1 = False
                    pbxScore.Print Score
                ElseIf Option2 = True Then
                    pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option2 = False
                ElseIf Option3 = True Then
                    pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option3 = False
            End If
        Case Is = 9
            If Option1 = True Then
                    pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option1 = False
                ElseIf Option2 = True Then
                    pbxResult.Print "Correct"
                    Score = Score + 1
                    pbxScore.Cls
                    pbxScore.Print Score
                    Option2 = False
                ElseIf Option3 = True Then
                    pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option3 = False
            End If
        Case Is = 10
           If Option1 = True Then
                pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option1 = False
                ElseIf Option2 = True Then
                    pbxResult.Print "That is"
                    pbxResult.Print "Incorrect"
                    Option2 = False
                ElseIf Option3 = True Then
                    pbxResult.Print "Correct"
                    Score = Score + 1
                    pbxScore.Cls
                    pbxScore.Print Score
                    Option3 = False
            End If
    End Select
End Sub
'Will print out hint for current question and subtract .5
'from the current score.  After three uses will only print
'out an an appology.
Private Sub cmdHint_Click()
    If Hint > 0 Then
            Hint = Hint - 1
            Select Case Question
                Case Is = 1
                    pbxQuiz.Print "Vulpix is a fire type."
                    Score = Score - 0.5
                    pbxScore.Cls
                    pbxScore.Print Score
                Case Is = 2
                    pbxQuiz.Print "Not all the pictures"
                    pbxQuiz.Print "were different."
                    Score = Score - 0.5
                    pbxScore.Cls
                    pbxScore.Print Score
                Case Is = 3
                    pbxQuiz.Print "It was on the top of her picture."
                    Score = Score - 0.5
                    pbxScore.Cls
                    pbxScore.Print Score
                Case Is = 4
                    pbxQuiz.Print "Think of his name HITmonlee."
                    Score = Score - 0.5
                    pbxScore.Cls
                    pbxScore.Print Score
                Case Is = 5
                    pbxQuiz.Print "Pikachu is a lightning type."
                    Score = Score - 0.5
                    pbxScore.Cls
                    pbxScore.Print Score
                Case Is = 6
                    pbxQuiz.Print "The character was"
                    pbxQuiz.Print "very honorable."
                    Score = Score - 0.5
                    pbxScore.Cls
                    pbxScore.Print Score
                Case Is = 7
                    pbxQuiz.Print "You can probably cheat"
                    pbxQuiz.Print "by looking to the right."
                    Score = Score - 0.5
                    pbxScore.Cls
                    pbxScore.Print Score
                Case Is = 8
                    pbxQuiz.Print "The number was a"
                    pbxQuiz.Print "Pokedex Number."
                    Score = Score - 0.5
                    pbxScore.Cls
                    pbxScore.Print Score
                Case Is = 9
                    pbxQuiz.Print "It was grey."
                    Score = Score - 0.5
                    pbxScore.Cls
                    pbxScore.Print Score
                Case Is = 10
                    pbxQuiz.Print "Ein is a dog."
                    Score = Score - 0.5
                    pbxScore.Cls
                    pbxScore.Print Score
            End Select
        Else
            pbxQuiz.Print "Sorry, you're out of hints."
    End If
End Sub
'Completes the Program
Private Sub cmdQuit_Click()
    End
End Sub
'Returns to the Main Menu for Review
'Reinitializes variables
Private Sub cmdReturn_Click()
    Atempt = True
    Option1.Visible = False
    Option2.Visible = False
    Option3.Visible = False
    Question = 0
    Score = 0
    Hint = 3
    frmQuiz.Hide
    frmMainMenu.Show
End Sub

'Loads the various questions and pictures
Private Sub cmdStart_Click()
    Option1.Visible = True
    Option2.Visible = True
    Option3.Visible = True
    Question = Question + 1
    pbxQuiz.Cls
        Select Case Question
            Case Is = 1
                pbxC1.Picture = LoadPicture(Pichu)
                pbxC2.Picture = LoadPicture(Bulbasaur)
                pbxC3.Picture = LoadPicture(Vulpix)
                pbxQuiz.Print "Which one is Vulpix?"
                cmdStart.Enabled = False
                cmdHint.Enabled = True
            Case Is = 2
                pbxC1.Picture = LoadPicture(SevenBeasts)
                pbxC2.Picture = LoadPicture(SixBeasts)
                pbxC3.Picture = LoadPicture(EightBeasts)
                pbxQuiz.Print "How many different kind of"
                pbxQuiz.Print "goblin pictures were there?"
                pbxQuiz.Print "1. Seven"
                pbxQuiz.Print "2. Six"
                pbxQuiz.Print "3. Eight"
                cmdStart.Enabled = False
                cmdHint.Enabled = True
            Case Is = 3
                pbxQuiz.Print "What is Jasmine's full name?"
                pbxQuiz.Print "1. Jasmine Bree"
                pbxQuiz.Print "2. Jasmine Bradford"
                pbxQuiz.Print "3. Jasmine Boreal"
                cmdStart.Enabled = False
                cmdHint.Enabled = True
            Case Is = 4
                pbxC1.Picture = LoadPicture(Hitmonlee)
                pbxC2.Picture = LoadPicture(Hitmonlee)
                pbxC3.Picture = LoadPicture(Hitmonlee)
                pbxQuiz.Print "Hitmonlee is a"
                pbxQuiz.Print "(1.Fighting 2.Flying 3.Grass)"
                pbxQuiz.Print "type pokemon."
                cmdStart.Enabled = False
                cmdHint.Enabled = True
            Case Is = 5
                pbxC1.Picture = LoadPicture(Red)
                pbxC2.Picture = LoadPicture(Green)
                pbxC3.Picture = LoadPicture(Yellow)
                pbxQuiz.Print "What color is Pikachu primarily?"
                pbxQuiz.Print "1.Red"
                pbxQuiz.Print "2.Green"
                pbxQuiz.Print "3.Yellow"
                cmdStart.Enabled = False
                cmdHint.Enabled = True
            Case Is = 6
                pbxQuiz.Print "What did you click to reach"
                pbxQuiz.Print "the Information section?"
                pbxQuiz.Print "1. A Ninja"
                pbxQuiz.Print "2. A Samurai"
                pbxQuiz.Print "3. A Warrior"
                cmdStart.Enabled = False
                cmdHint.Enabled = True
            Case Is = 7
                pbxC1.Picture = LoadPicture(FourBeasts)
                pbxC2.Picture = LoadPicture(FiveBeasts)
                pbxC3.Picture = LoadPicture(SixBeasts)
                pbxQuiz.Print "How many sections are there"
                pbxQuiz.Print "in this project?"
                pbxQuiz.Print "1. Four"
                pbxQuiz.Print "2. Five"
                pbxQuiz.Print "3. Six"
                cmdStart.Enabled = False
                cmdHint.Enabled = True
            Case Is = 8
                pbxQuiz.Print "What aspect couldn't you"
                pbxQuiz.Print "sort Pokemon by?"
                pbxQuiz.Print "1. Amount"
                pbxQuiz.Print "2. Pokedex"
                pbxQuiz.Print "3. Name"
                cmdStart.Enabled = False
                cmdHint.Enabled = True
            Case Is = 9
                pbxQuiz.Print "What guards the Quiz?"
                pbxQuiz.Print "1. A Red Fox"
                pbxQuiz.Print "2. A Grey Wolf"
                pbxQuiz.Print "3. A Siberian Husky"
                cmdStart.Enabled = False
                cmdHint.Enabled = True
            Case Is = 10
                pbxC1.Picture = LoadPicture(Hamtaro)
                pbxC2.Picture = LoadPicture(Fish)
                pbxC3.Picture = LoadPicture(Ein)
                pbxQuiz.Print "Which one is Ein?"
                cmdStart.Enabled = False
                cmdHint.Enabled = True
            Case Is > 10
                If Score > 6 Then
                        MsgBox "Good job, you passed.", , "Quiz Master"
                        cmdQuit.Visible = True
                    Else
                        MsgBox "You still have much to learn.  I will be waiting", , "Quiz Master"
                        cmdReturn.Visible = True
                End If
        End Select
End Sub
'initializes variables
Private Sub Form_Load()
    Option1.Visible = False
    Option2.Visible = False
    Option3.Visible = False
    cmdQuit.Visible = False
    cmdReturn.Visible = False
    cmdChoice.Enabled = False
    cmdHint.Enabled = False
    Hint = 3
    Question = 0
    Score = 0
    Path = "N:\CS130\handin\Chris Davin's VBProject\"
    Pichu = Path & "Images Used\Pichu.jpg"
    Bulbasaur = Path & "Images Used\01card.gif"
    Vulpix = Path & "Images Used\Vulpix.jpg"
    SevenBeasts = Path & "Images Used\SevenBattleBeasts.jpg"
    SixBeasts = Path & "Images Used\SixNumberBeasts.jpg"
    EightBeasts = Path & "Images Used\EightNumberBeasts.jpg"
    Hitmonlee = Path & "Images Used\hitmonlee.jpg"
    Red = Path & "Images Used\red.gif"
    Green = Path & "Images Used\green.gif"
    Yellow = Path & "Images Used\yellow.gif"
    FourBeasts = Path & "Images Used\fourNumberBeasts.jpg"
    FiveBeasts = Path & "Images Used\Beast1.jpg"
    Hamtaro = Path & "Images Used\Hamtaro.jpg"
    Fish = Path & "Images Used\Fish1.jpg"
    Ein = Path & "Images Used\Ein.jpg"
End Sub
'allows player to input the choice
'By enabling the Input Guess button
Private Sub Option1_Click()
    cmdChoice.Enabled = True
End Sub
'allows player to input the choice
'By enabling the Input Guess button
Private Sub Option2_Click()
    cmdChoice.Enabled = True
End Sub
'allows player to input the choice
'By enabling the Input Guess button
Private Sub Option3_Click()
    cmdChoice.Enabled = True
End Sub
'Data about the picture clicked
Private Sub Picture1_Click()
    MsgBox "Welcome.  I am the Quiz Master.", , "Quiz Master"
End Sub
