VERSION 5.00
Begin VB.Form frmUnscramble 
   Caption         =   "Unscramble Words"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   Picture         =   "frmUnscramble.frx":0000
   ScaleHeight     =   5280
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   4560
      TabIndex        =   12
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdBackToEnglish 
      Caption         =   "Back to English Screen"
      Height          =   615
      Left            =   4560
      TabIndex        =   11
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdCheckAnswers 
      Caption         =   "Check Answers!"
      Height          =   615
      Left            =   4560
      TabIndex        =   10
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtPuzzleFive 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txtPuzzleFour 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox txtPuzzleThree 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtPuzzleTwo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtPuzzleOne 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblUnscramble 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "UnScramble the Following Words:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   1440
      TabIndex        =   13
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label lblPuzzleFive 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "aliiroaacn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label lblPuzzleFour 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "noziaar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblPuzzleThree 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "oorlcaod"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblPuzzleTwo 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "hooi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblPuzzleOne 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "intaosnme"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "frmUnscramble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBackToEnglish_Click()
frmUnscramble.Hide
frmEnglish.Show
End Sub

Private Sub cmdCheckAnswers_Click()
Dim POne As String, PTwo As String, PThree As String
Dim PFour As String, PFive As String
Dim CTR As Integer

POne = txtPuzzleOne.Text
PTwo = txtPuzzleTwo.Text
PThree = txtPuzzleThree.Text
PFour = txtPuzzleFour.Text
PFive = txtPuzzleFive.Text
CTR = 0
If POne = "" Or (POne <> "Minnesota" And POne <> "minnesota") Then
    CTR = CTR + 1
End If
If PTwo = "" Or (PTwo <> "Ohio" And PTwo <> "ohio") Then
    CTR = CTR + 1
End If
If PThree = "" Or (PThree <> "Colorado" And PThree <> "colorado") Then
    CTR = CTR + 1
End If
If PFour = "" Or (PFour <> "Arizona" And PFour <> "arizona") Then
    CTR = CTR + 1
End If
If PFive = "" Or (PFive <> "California" And PFive <> "california") Then
    CTR = CTR + 1
End If
If CTR = 0 Then
    MsgBox ("Congradulations you got them all correct!")
Else
    MsgBox ("You got " & CTR & " incorrect out of five.")
End If
End Sub

Private Sub cmdQuit_Click()
MsgBox ("Good luck with your " & Homework & " hours of homework!")
End
End Sub
