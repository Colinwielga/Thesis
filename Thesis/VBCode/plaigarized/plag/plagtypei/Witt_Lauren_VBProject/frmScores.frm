VERSION 5.00
Begin VB.Form frmScores
   BackColor       =   &H00000000&
   Caption         =   "Scores"
   ClientHeight    =   12675
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   ScaleHeight     =   12675
   ScaleWidth      =   11460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgoback
      BackColor       =   &H00FFC0FF&
      Caption         =   "Go back to the Roommate Challenge"
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7560
      Width           =   2175
   End
   Begin VB.PictureBox picResults
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10455
      Left            =   3000
      ScaleHeight     =   10395
      ScaleWidth      =   8115
      TabIndex        =   4
      Top             =   1560
      Width           =   8175
   End
   Begin VB.CommandButton cmdCompare
      BackColor       =   &H00FFC0FF&
      Caption         =   "Compare your score team's high score with the average"
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton cmdcertainscore
      BackColor       =   &H00FFC0FF&
      Caption         =   "Search for all teams with a certain score"
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdDescending
      BackColor       =   &H00FFC0FF&
      Caption         =   "Show scores in descending order"
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton cmdShowScores
      BackColor       =   &H00FFC0FF&
      Caption         =   "Show scores"
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lblScores
      BackColor       =   &H00FF80FF&
      Caption         =   "             Scores!"
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   36
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   11535
   End
End
Attribute VB_Name = "frmScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

Dim CTR As Integer
Dim total As Integer
Dim Player(1 To 50) As String
Dim Score(1 To 50) As Integer
Dim tempScore As Integer

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

Dim tempPlayer As String
Dim pos As Integer

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

Dim pass As Integer

Private Sub cmdgoback_Click()
frmScores.Hide

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

frmRoommateChallenge3.Show

End Sub

Private Sub cmdShowScores_Click()
Dim pos As Integer

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

picResults.Print "Players", , "Score"
picResults.Print "*****************************************"
Open App.Path & "\Scores.txt" For Input As #1
CTR = 0
Do While Not EOF(1)

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

CTR = CTR + 1
Input #1, Player(CTR), Score(CTR)
picResults.Print Player(CTR), Score(CTR)

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

total = total + Score(CTR)
Loop
Close #1


' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

End Sub
Private Sub cmdDescending_Click()

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

picResults.Print "Players", , "Score"
picResults.Print "*****************************************"
For pass = 1 To CTR
For pos = 1 To CTR - 1
If Score(pos) < Score(pos + 1) Then

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

tempScore = Score(pos)
Score(pos) = Score(pos + 1)

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

Score(pos + 1) = tempScore

tempPlayer = Player(pos)

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

Player(pos) = Player(pos + 1)
Player(pos + 1) = tempPlayer

End If
Next pos

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

Next pass

picResults.Cls
For pos = 1 To CTR
picResults.Print Player(pos), Score(pos)

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

Next pos

End Sub

Private Sub cmdcertainscore_Click()
Dim found As Boolean

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

Dim pos As Integer, InputScore As Integer

InputScore = InputBox("Enter the score you are looking for.", "Input Score")

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

found = False
picResults.Cls
For pos = 1 To CTR

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

If InputScore = Score(pos) Then
found = True
picResults.Print Player(pos), , Score(pos)
End If
Next pos

If found = False Then

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

MsgBox "There were no teams with that score.", , "None"
End If



' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

End Sub

Private Sub cmdCompare_Click()
Dim found As Boolean
Dim pos As Integer

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

Dim InputName As String

InputName = InputBox("Enter your team's name in the format: 'Player 1 & Player 2.' Eg: Lauren & Caitlin")

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

pos = 0
found = False
Do While found = False And pos < CTR
pos = pos + 1

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

If InputName = Player(pos) Then
found = True
picResults.Cls

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

picResults.Print "The average score is "; FormatNumber(total / CTR); ". And your highest score is "; Score(pos); "."
If FormatNumber(total / CTR) > Score(pos) Then
picResults.Print "You suck! Spend more time getting to know your roommate!"

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

ElseIf FormatNumber(total / CTR) = Score(pos) Then
picResults.Print "You're just average."
Else
picResults.Print "Good for you! You clearly know your roommie!"

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

End If
End If
Loop

If found = False Then

' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words

MsgBox "There is no data for a team with those names.  Try switching the order you entered your names.", , "Error"
End If


End Sub


' block comment here
' and here
' and some words here
' blah stuff things meh
' more different distinct words


