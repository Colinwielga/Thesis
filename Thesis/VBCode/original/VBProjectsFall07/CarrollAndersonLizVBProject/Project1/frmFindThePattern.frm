VERSION 5.00
Begin VB.Form frmFindThePattern 
   BackColor       =   &H80000013&
   Caption         =   "Find The Pattern"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   2640
      TabIndex        =   13
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdMainScreen 
      Caption         =   "Back to Main Screen"
      Height          =   495
      Left            =   480
      TabIndex        =   12
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton cmdBackToMath 
      Caption         =   "Back to Math    Main Screen"
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdCheckAnswers 
      Caption         =   "Check Answers"
      Height          =   495
      Left            =   480
      TabIndex        =   10
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdHintThree 
      BackColor       =   &H8000000D&
      Caption         =   "Hint!"
      Height          =   495
      Left            =   4200
      TabIndex        =   9
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtPatternThree 
      BackColor       =   &H8000000B&
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdHintTwo 
      BackColor       =   &H8000000D&
      Caption         =   "Hint!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdHintOne 
      BackColor       =   &H8000000D&
      Caption         =   "Hint!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtPatternTwo 
      BackColor       =   &H8000000B&
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtPatternOne 
      BackColor       =   &H8000000A&
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblPatternThree 
      BackColor       =   &H8000000D&
      Caption         =   "1 , 11 , 21 , 1211"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblPatternTwo 
      BackColor       =   &H8000000D&
      Caption         =   "1 , 4 , 9 , 16 , 25"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblPatternOne 
      BackColor       =   &H8000000D&
      Caption         =   "3 , 4 , 7 , 11 , 18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblFindAPattern 
      BackColor       =   &H8000000D&
      Caption         =   "Find a Pattern"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmFindThePattern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBackToMath_Click()

'This takes the user back to the math screen

frmFindThePattern.Hide
frmMath.Show
End Sub

Private Sub cmdCheckAnswers_Click()
Dim CTR As Integer
Dim Score As Integer
Dim PatternOne As String, PatternTwo As String
Dim PatternThree As String

'This button checks the answers using if statements
'then tells the user how many they got wrong

PatternOne = txtPatternOne.Text
PatternTwo = txtPatternTwo.Text
PatternThree = txtPatternThree.Text

If PatternOne <> "29" Or PatternOne = "" Then
    CTR = CTR + 1
End If
If PatternTwo <> "36" Or PatternTwo = "" Then
    CTR = CTR + 1
End If
If PatternThree <> "111221" Or PatternThree = "" Then
    CTR = CTR + 1
End If

Score = 3 - CTR

MsgBox (Name1 & ", your got " & Score & " out of three correct.")

End Sub

'Hint one two and three simply message box the user a
'hint on how to solve the patterns

Private Sub cmdHintOne_Click()
MsgBox ("Try to find out how you can get the third term from the first two.")
End Sub

Private Sub cmdHintThree_Click()
MsgBox ("Try saying the numbers out loud, one number at a time")
End Sub

Private Sub cmdHintTwo_Click()
MsgBox ("Try to find out how you can go from '1' to the first term, from '2' to the second term, ect.")
End Sub

Private Sub cmdMainScreen_Click()

'This takes the user back to the main screen

frmFindThePattern.Hide
frmMainScreen.Show
End Sub

Private Sub cmdQuit_Click()

'This tells the user good luck with their homework, then
'ends the program

MsgBox ("Good luck with your " & Homework & " hours of homework!")
End
End Sub
