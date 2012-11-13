VERSION 5.00
Begin VB.Form FrmPlayGame 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hang-Man"
   ClientHeight    =   8685
   ClientLeft      =   3030
   ClientTop       =   1695
   ClientWidth     =   10065
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "HEGHangManGame.frx":0000
   ScaleHeight     =   8685
   ScaleWidth      =   10065
   Begin VB.CommandButton CmdHint 
      BackColor       =   &H00008000&
      Caption         =   "Hint..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8160
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton CmdPlayOptions 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go-to Option Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton CmdEnterLetter 
      BackColor       =   &H0000FFFF&
      Caption         =   "Enter New Letter"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2400
      Width           =   1815
   End
   Begin VB.PictureBox PicMan 
      BorderStyle     =   0  'None
      Height          =   1335
      Index           =   5
      Left            =   1920
      Picture         =   "HEGHangManGame.frx":129FBA
      ScaleHeight     =   1335
      ScaleWidth      =   1095
      TabIndex        =   21
      Top             =   5160
      Width           =   1095
      Begin VB.PictureBox Picture8 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   735
         TabIndex        =   22
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.PictureBox PicMan 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   4
      Left            =   840
      Picture         =   "HEGHangManGame.frx":12FF5C
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   19
      Top             =   5160
      Width           =   1095
      Begin VB.PictureBox Picture7 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   735
         TabIndex        =   20
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.PictureBox PicMan 
      BorderStyle     =   0  'None
      Height          =   1095
      Index           =   3
      Left            =   2040
      Picture         =   "HEGHangManGame.frx":135EFE
      ScaleHeight     =   1095
      ScaleWidth      =   975
      TabIndex        =   17
      Top             =   3840
      Width           =   975
      Begin VB.PictureBox Picture6 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   735
         TabIndex        =   18
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.PictureBox PicMan 
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   2
      Left            =   960
      Picture         =   "HEGHangManGame.frx":13BEA0
      ScaleHeight     =   1215
      ScaleWidth      =   975
      TabIndex        =   15
      Top             =   3720
      Width           =   975
      Begin VB.PictureBox Picture5 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   735
         TabIndex        =   16
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.PictureBox PicMan 
      BorderStyle     =   0  'None
      Height          =   1815
      Index           =   1
      Left            =   1560
      Picture         =   "HEGHangManGame.frx":141E42
      ScaleHeight     =   1815
      ScaleWidth      =   975
      TabIndex        =   13
      Top             =   3480
      Width           =   975
      Begin VB.PictureBox Picture2 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   735
         TabIndex        =   14
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.PictureBox PicMan 
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   0
      Left            =   1560
      Picture         =   "HEGHangManGame.frx":149B50
      ScaleHeight     =   1215
      ScaleWidth      =   975
      TabIndex        =   11
      Top             =   2400
      Width           =   975
      Begin VB.PictureBox Picture3 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   735
         TabIndex        =   12
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.PictureBox PicOutputGuess 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   8280
      ScaleHeight     =   4695
      ScaleWidth      =   1575
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   7
      Left            =   6840
      Picture         =   "HEGHangManGame.frx":14D376
      ScaleHeight     =   975
      ScaleWidth      =   855
      TabIndex        =   9
      Top             =   7560
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   2
      Left            =   2640
      Picture         =   "HEGHangManGame.frx":151114
      ScaleHeight     =   975
      ScaleWidth      =   855
      TabIndex        =   8
      Top             =   7560
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   6
      Left            =   6000
      Picture         =   "HEGHangManGame.frx":154EB2
      ScaleHeight     =   975
      ScaleWidth      =   855
      TabIndex        =   7
      Top             =   7560
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   5
      Left            =   5160
      Picture         =   "HEGHangManGame.frx":158C50
      ScaleHeight     =   975
      ScaleWidth      =   855
      TabIndex        =   6
      Top             =   7560
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   4
      Left            =   4320
      Picture         =   "HEGHangManGame.frx":15C9EE
      ScaleHeight     =   975
      ScaleWidth      =   855
      TabIndex        =   5
      Top             =   7560
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   3
      Left            =   3480
      Picture         =   "HEGHangManGame.frx":16078C
      ScaleHeight     =   975
      ScaleWidth      =   855
      TabIndex        =   4
      Top             =   7560
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   1
      Left            =   1800
      Picture         =   "HEGHangManGame.frx":16452A
      ScaleHeight     =   975
      ScaleWidth      =   855
      TabIndex        =   3
      Top             =   7560
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   0
      Left            =   960
      Picture         =   "HEGHangManGame.frx":1682C8
      ScaleHeight     =   975
      ScaleWidth      =   855
      TabIndex        =   2
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton CmdNew 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton CmdEnd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Leave Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   27
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Incorrect Guesses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8280
      TabIndex        =   26
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hang-Man!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1680
      TabIndex        =   25
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "FrmPlayGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Hang Man
'Form Name: FrmHowTo
'Authors: Breanna Riskey and Heidi Golliet
'Date Completed: Monday, March 31st
'Objective: This form is where the user actually plays the game. Here the user t

Option Explicit
Dim WordBank(1 To 100) As String
Dim WordCount As Integer
Dim Word As String
Dim WordLen As Integer
Dim GuessCount As Single
Dim LettersGuessed(1 To 20) As String
Dim WrongGuessCount As Integer
Dim CorrectCount As Integer
Dim IncorrectGuesses(1 To 6) As String
Dim Hint(1 To 100) As String
Dim ComputeInt As Integer



Private Sub CmdEnd_Click()
    FrmHome.Visible = True
    FrmPlayGame.Visible = False
    FrmOptions.Visible = False
End Sub

Private Sub CmdEnterLetter_Click()
'If Letter is equal to any part of the word, then the letter should be printed
'in the appropriate picture box. That means that each picture box should
'be assigned a letter. If the letter is not equal to any part of the word string
'then it should be printed in the box reserved for non-matching guesses
'This process should be repeated until there are either six wrong guesses
'or the word is completed.

Dim characters As String
Dim Length As Integer
Dim CTR As Integer
Dim foundstr As Boolean
Dim Letter As String
Dim NewGame As Integer
Letter = UCase(Trim(InputBox("Enter your guess - any one letter a through z:", "Enter Letter")))
NewGame = 1
Length = Len(Letter)

PicOutputGuess.FontSize = 25

'Add the letter to LettersGuessed

foundstr = False

'PicOutputGuess.Print Letter (this is a test)

For CTR = 1 To WordLen
    characters = UCase(Mid(Word, CTR, 1))
    'PicOutputGuess.Print characters (this is a test)
    If Letter = UCase(Mid(Word, CTR, 1)) Then
        foundstr = True
        CorrectCount = CorrectCount + 1
        Picture1(CTR - 1).Print Letter
    End If
Next CTR

'PicOutputGuess is the name of the picture box that holds the non-matching guesses
'Picture1 is the name of the controlled array for the letter boxes

If Length = 1 Then
    GuessCount = GuessCount + 1
'    PicOutputGuess.Print foundstr (this is a test)
    If Not foundstr Then
        'here we need to add letter to a string we'll call IncorrectGuesses
        PicOutputGuess.Print Letter
        WrongGuessCount = WrongGuessCount + 1
    End If
Else
    MsgBox "Sorry, please try entering just one letter", , "Error"
End If



'have an exhaustive search for finding match



'this is where the visibility for the hangman figure comes from
If WrongGuessCount > 0 And WrongGuessCount <= 6 Then
    PicMan(WrongGuessCount - 1).Visible = True
End If

'this tells the user when they have won
If CorrectCount = WordLen Then
    MsgBox "Congratulations- You Win!!", , "Notice:"
    CorrectCount = -21
End If



If CorrectCount > -21 And CorrectCount <= -1 Then
    MsgBox "You have already won... Start a new game!", , "Notice:"
End If

If WrongGuessCount = 6 Then
    MsgBox "Sorry- You Lost!!", , "Notice:"
    WrongGuessCount = -21
End If

If WrongGuessCount > -21 And WrongGuessCount < 0 Then
    MsgBox "You already lost. Try playing a new game", , "Notice:"
End If


End Sub

Private Sub CmdHint_Click()
'this displays the hint that goes along with the word
MsgBox "Here is your hint: " & Hint(ComputeInt) & ".", , "Hint"
End Sub

Private Sub CmdNew_Click()
Dim DashReset As Integer
Dim RandomNum As Single
Dim CTR As Integer
'Dim ComputeInt As Integer

GuessCount = 0
WrongGuessCount = 0
CorrectCount = 0

PicOutputGuess.Cls

'This is the randomizer code
RandomNum = Rnd(3.1415926)
'(test)PicOutputGuess.Print (RandomNum)

If RandomNum = 0 Then RandomNum = 3
'(test)PicOutputGuess.Print (RandomNum)

ComputeInt = Int(RandomNum * WordCount + 1)
'(test)PicOutputGuess.Print (ComputeInt)

If ComputeInt = 0 Then ComputeInt = 3

'The following trims the word to ensure no spaces and determines length of word
Word = Trim(WordBank(ComputeInt))
WordLen = Len(Word)


'(below is a test to check how random the randomizer is)
'PicOutputGuess.Print (Word)

DashReset = 0
Do While DashReset < 8
    Picture1(DashReset).Visible = True
    DashReset = DashReset + 1
Loop

For CTR = WordLen To 7
    Picture1(CTR).Visible = False
Next CTR

'This clears all the spaces of their letters for the new game
Picture1(0).Cls
Picture1(1).Cls
Picture1(2).Cls
Picture1(3).Cls
Picture1(4).Cls
Picture1(5).Cls
Picture1(6).Cls
Picture1(7).Cls

'this resets the visibility of the stick figure
For CTR = 0 To 5
    PicMan(CTR).Visible = False
Next CTR

'this clear the LettersGuessed array
For CTR = 1 To 20
    LettersGuessed(CTR) = ""
Next CTR

End Sub

Private Sub CmdPlayOptions_Click()
    FrmHome.Visible = False
    FrmOptions.Visible = True
    FrmPlayGame.Visible = False
End Sub

Private Sub Form_Load()

'When the form is loaded, the wordbank is loaded so that it is immediately ready for the user
'because of this, it is this form where the counter is set

Dim CTR As Integer
CTR = 0

Open App.Path & "/wordbank.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, WordBank(CTR), Hint(CTR)
Loop

Close #1

CorrectCount = 0
WordCount = CTR
GuessCount = 0

'This formats the spaces
For CTR = 0 To 6
    Picture1(CTR).FontSize = 35
    Picture1(CTR).Cls
Next CTR

'This sets the stick figure's visibility as false
For CTR = 0 To 5
    PicMan(CTR).Visible = False
Next CTR

Call CmdNew_Click

End Sub

