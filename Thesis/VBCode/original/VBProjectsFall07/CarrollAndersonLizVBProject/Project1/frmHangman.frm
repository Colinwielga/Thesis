VERSION 5.00
Begin VB.Form frmHangman 
   BackColor       =   &H80000009&
   Caption         =   "Hangman"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picNoLetter 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   7560
      ScaleHeight     =   2715
      ScaleWidth      =   915
      TabIndex        =   25
      Top             =   840
      Width           =   975
   End
   Begin VB.PictureBox picE3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   23
      Top             =   5040
      Width           =   375
   End
   Begin VB.PictureBox picC3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   22
      Top             =   5040
      Width           =   375
   End
   Begin VB.PictureBox picE2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   21
      Top             =   5040
      Width           =   375
   End
   Begin VB.PictureBox picI1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   20
      Top             =   5040
      Width           =   375
   End
   Begin VB.PictureBox picC2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   19
      Top             =   5040
      Width           =   375
   End
   Begin VB.PictureBox picS1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   18
      Top             =   5040
      Width           =   375
   End
   Begin VB.PictureBox picR1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   17
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox picE1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   16
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox picT1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   15
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox picU1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   14
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox picP1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   13
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox picM1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   12
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox picO1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   11
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox picC1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   720
      TabIndex        =   9
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdBackToEnglish 
      Caption         =   "Go Back to English Screen"
      Height          =   615
      Left            =   720
      TabIndex        =   8
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdSolvePuzzle 
      Caption         =   "Solve Puzzle"
      Height          =   615
      Left            =   720
      TabIndex        =   7
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdEnterLetter 
      Caption         =   "Enter Letter"
      Height          =   615
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   6
      Top             =   600
      Width           =   2175
   End
   Begin VB.PictureBox picFive 
      Height          =   3615
      Left            =   3600
      Picture         =   "frmHangman.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   3075
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.PictureBox picFour 
      Height          =   3615
      Left            =   3600
      Picture         =   "frmHangman.frx":24582
      ScaleHeight     =   3555
      ScaleWidth      =   3075
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.PictureBox picThree 
      Height          =   3615
      Left            =   3600
      Picture         =   "frmHangman.frx":48B04
      ScaleHeight     =   3555
      ScaleWidth      =   3075
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.PictureBox picTwo 
      Height          =   3615
      Left            =   3600
      Picture         =   "frmHangman.frx":6D086
      ScaleHeight     =   3555
      ScaleWidth      =   3075
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.PictureBox picOne 
      Height          =   3615
      Left            =   3600
      Picture         =   "frmHangman.frx":91608
      ScaleHeight     =   3555
      ScaleWidth      =   3075
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.PictureBox picZero 
      Height          =   3615
      Left            =   3600
      Picture         =   "frmHangman.frx":B5B8A
      ScaleHeight     =   3555
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label lblAlreadyUsed 
      BackStyle       =   0  'Transparent
      Caption         =   "Incorrect Letters:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   24
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "frmHangman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBackToEnglish_Click()
frmHangman.Hide
frmEnglish.Show
End Sub

Private Sub cmdEnterLetter_Click()
CTR = 0
Total = 0
Do While CTR < 5 And Total < 15
Found = False
Letter = InputBox("Please enter one Upper case letter")
    If "C" = Letter Then
        picC1.Print "C"
        picC2.Print "C"
        picC3.Print "C"
        Found = True
        Total = Total + 3
    ElseIf "O" = Letter Then
        picO1.Print "O"
        Found = True
        Total = Total + 1
    ElseIf "M" = Letter Then
        picM1.Print "M"
        Found = True
        Total = Total + 1
    ElseIf "P" = Letter Then
        picP1.Print "P"
        Found = True
        Total = Total + 1
    ElseIf "U" = Letter Then
        picU1.Print "U"
        Found = True
        Total = Total + 1
    ElseIf "T" = Letter Then
        picT1.Print "T"
        Found = True
        Total = Total + 1
    ElseIf "E" = Letter Then
        picE1.Print "E"
        picE2.Print "E"
        picE3.Print "E"
        Found = True
        Total = Total + 3
    ElseIf "R" = Letter Then
        picR1.Print "R"
        Found = True
        Total = Total + 1
    ElseIf "S" = Letter Then
        picS1.Print "S"
        Found = True
        Total = Total + 1
    ElseIf "I" = Letter Then
        picI1.Print "I"
        Found = True
        Total = Total + 1
    ElseIf "N" = Letter Then
        picN1.Print "N"
        Found = True
        Total = Total + 1
    End If
    If Found = False Then
        CTR = CTR + 1
        picNoLetter.Print Letter
    End If
     If CTR = 0 Then
        picZero.Visible = True
        picOne.Visible = False
        picTwo.Visible = False
        picThree.Visible = False
        picFour.Visible = False
        picFive.Visible = False
    ElseIf CTR = 1 Then
        picZero.Visible = False
        picOne.Visible = True
        picTwo.Visible = False
        picThree.Visible = False
        picFour.Visible = False
        picFive.Visible = False
    ElseIf CTR = 2 Then
        picZero.Visible = False
        picOne.Visible = False
        picTwo.Visible = True
        picThree.Visible = False
        picFour.Visible = False
        picFive.Visible = False
    ElseIf CTR = 3 Then
        picZero.Visible = False
        picOne.Visible = False
        picTwo.Visible = False
        picThree.Visible = True
        picFour.Visible = False
        picFive.Visible = False
    ElseIf CTR = 4 Then
        picZero.Visible = False
        picOne.Visible = False
        picTwo.Visible = False
        picThree.Visible = False
        picFour.Visible = True
        picFive.Visible = False
    ElseIf CTR = 5 Then
        picZero.Visible = False
        picOne.Visible = False
        picTwo.Visible = False
        picThree.Visible = False
        picFour.Visible = False
        picFive.Visible = True
    End If
Loop
    If CTR = 0 Then
        picZero.Visible = True
        picOne.Visible = False
        picTwo.Visible = False
        picThree.Visible = False
        picFour.Visible = False
        picFive.Visible = False
    ElseIf CTR = 1 Then
        picZero.Visible = False
        picOne.Visible = True
        picTwo.Visible = False
        picThree.Visible = False
        picFour.Visible = False
        picFive.Visible = False
    ElseIf CTR = 2 Then
        picZero.Visible = False
        picOne.Visible = False
        picTwo.Visible = True
        picThree.Visible = False
        picFour.Visible = False
        picFive.Visible = False
    ElseIf CTR = 3 Then
        picZero.Visible = False
        picOne.Visible = False
        picTwo.Visible = False
        picThree.Visible = True
        picFour.Visible = False
        picFive.Visible = False
    ElseIf CTR = 4 Then
        picZero.Visible = False
        picOne.Visible = False
        picTwo.Visible = False
        picThree.Visible = False
        picFour.Visible = True
        picFive.Visible = False
    ElseIf CTR = 5 Then
        picZero.Visible = False
        picOne.Visible = False
        picTwo.Visible = False
        picThree.Visible = False
        picFour.Visible = False
        picFive.Visible = True
    End If
    

If Total = 15 Then
    MsgBox ("Congradulations, " & Name1 & " you solved the puzzle!")
End If
If CTR = 5 Then
    MsgBox ("Sorry you ran out of guesses.")
End If
End Sub

Private Sub cmdQuit_Click()
MsgBox ("Good luck with your " & Homework & " hours of homework!")
End
End Sub

Private Sub cmdSolvePuzzle_Click()
Dim Answer As String
Answer = InputBox("Please enter guess in all lower case, with a space between words.")
If Answer = "computer science" Then
    MsgBox ("Congradulations your answer is correct!")
Else
    MsgBox ("This answer is not correct, try guessing a couple more letters and try again.")
End If
End Sub
