VERSION 5.00
Begin VB.Form frmQuiz 
   BackColor       =   &H00FF8080&
   Caption         =   "Practice Quiz"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDesigner 
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   6120
      TabIndex        =   10
      Text            =   "Designed by Amanda Aamodt"
      Top             =   5640
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuestion9 
      Caption         =   "Question 9"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6240
      TabIndex        =   9
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuestion8 
      Caption         =   "Question 8"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   8
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuestion7 
      Caption         =   "Question 7"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuestion6 
      Caption         =   "Question 6"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6240
      TabIndex        =   6
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuestion5 
      Caption         =   "Question 5"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   5
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuestion4 
      Caption         =   "Question 4"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuestion3 
      Caption         =   "Question 3"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6240
      TabIndex        =   3
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuestion2 
      Caption         =   "Question 2"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuestion1 
      Caption         =   "Question 1"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   $"frmQuiz.frx":0000
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8535
   End
End
Attribute VB_Name = "frmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this part of the program is for the students to take a practice quiz.
'when the button for each question is pressed, a question appears in an input box so the student can input their answer
'using If/Then statements, the program determines whether or not the inputed answer is correct
'based on the input, the program outputs a message box with a message of either "correct" or "incorrect" with the correct answer

Private Sub cmdQuestion1_Click()
    Dim X As Single     'declare variable
    X = InputBox("Evaluate cos 0°", "Question 1")   'causes input box to appear for the student to read and answer the question
    If X = 1 Then   '
        MsgBox "Good Job! cos 0° = 1", , "Correct"  'message displayed in message box if answer inputed in the input box is correct
        Else
        MsgBox "Incorrect. cos 0° = 1", , "Incorrect"   'message displayed in message box if answer inputed in the input box is incorrect
    End If
End Sub

Private Sub cmdQuestion2_Click()
    Dim X As Single 'declare variable
    X = InputBox("Complete this proportion,  4 : 9 = 24 : ?", "Question 2")
    If X = 54 Then
        MsgBox "Good Job! 4 : 9 = 24 : 54", , "Correct" 'message displayed in message box if answer inputed in the input box is correct
        Else
        MsgBox "Incorrect. 4 : 9 = 24 : 54", , "Incorrect"  'message displayed in message box if answer inputed in the input box is incorrect
    End If
End Sub

Private Sub cmdQuestion3_Click()
    Dim X As Single 'declare variable
    X = InputBox("The square root of 64 is ?", "Question 3")    'causes input box to appear for the student to read and answer the question
    If X = 8 Then
        MsgBox "Good Job! The square root of 64 is 8", , "Correct"  'message displayed in message box if answer inputed in the input box is correct
        Else
        MsgBox "Incorrect. The square root of 64 is 8", , "Incorrect"   'message displayed in message box if answer inputed in the input box is incorrect
    End If
End Sub

Private Sub cmdQuestion4_Click()
    Dim X As Single 'declare variable
    X = InputBox("The lengths of the two sides of a right triange are 5 and 12. What is the length of the hypotenuse?", "Question 4")   'causes input box to appear for the student to read and answer the question
    If X = 13 Then
        MsgBox "Good Job! The length of the hypotenuse is 13", , "Correct"  'message displayed in message box if answer inputed in the input box is correct
        Else
        MsgBox "Incorrect. The length of the hypotenuse is 13", , "Incorrect"   'message displayed in message box if answer inputed in the input box is incorrect
    End If
End Sub

Private Sub cmdQuestion5_Click()
    Dim X As Single 'declare variable
    X = InputBox("At x = 3(pi)/2, sin(x) = ?", "Question 5")    'causes input box to appear for the student to read and answer the question
    If X = -1 Then
        MsgBox "Good Job! sin(x) = -1", , "Correct" 'message displayed in message box if answer inputed in the input box is correct
        Else
        MsgBox "Incorrect. sin(x) = -1", , "Incorrect"  'message displayed in message box if answer inputed in the input box is incorrect
    End If
End Sub

Private Sub cmdQuestion6_Click()
    Dim X As Single 'declare variable
    X = InputBox("The length of one side of a right triange is 8 and the length of the hypotenuse is 17. What is the length of the other side?", "Question 6")  'causes input box to appear for the student to read and answer the question
    If X = 15 Then
        MsgBox "Good Job! The length of the other side is 15", , "Correct"  'message displayed in message box if answer inputed in the input box is correct
        Else
        MsgBox "Incorrect. The length of the other side is 15", , "Incorrect"   'message displayed in message box if answer inputed in the input box is incorrect
    End If
End Sub

Private Sub cmdQuestion7_Click()
    Dim X As Single 'declare variable
    X = InputBox("3 is the tenth part of what number?", "Question 7")   'causes input box to appear for the student to read and answer the question
    If X = 30 Then
        MsgBox "Good Job! 3 is the tenth part of 30", , "Correct"   'message displayed in message box if answer inputed in the input box is correct
        Else
        MsgBox "Incorrect. 3 is the tenth part of 30", , "Incorrect"    'message displayed in message box if answer inputed in the input box is incorrect
    End If
End Sub

Private Sub cmdQuestion8_Click()
    Dim X As String 'declare variable
    X = InputBox("Which of the following numbers are rational? 1, -6, 3½, 45, -13, 5, 0, 7.38609", "Question 8")    'causes input box to appear for the student to read and answer the question
    If X = "All" Then
        MsgBox "Good Job! They are all rational", , "Correct"   'message displayed in message box if answer inputed in the input box is correct
        Else
        MsgBox "Incorrect. They are all rational", , "Incorrect"    'message displayed in message box if answer inputed in the input box is incorrect
    End If
End Sub

Private Sub cmdQuestion9_Click()
    Dim X As Single 'declare variable
    X = InputBox("Solve this proportion:  8 : 12 = 2 : ?", "Question 9")    'causes input box to appear for the student to read and answer the question
    If X = 3 Then
        MsgBox "Good Job! 8 : 12 = 2 : 3", , "Correct"  'message displayed in message box if answer inputed in the input box is correct
        Else
        MsgBox "Incorrect. 8 : 12 = 2 : 3", , "Incorrect" 'message displayed in message box if answer inputed in the input box is incorrect
    End If
End Sub
