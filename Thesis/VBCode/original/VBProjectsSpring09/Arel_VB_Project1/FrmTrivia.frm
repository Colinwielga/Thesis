VERSION 5.00
Begin VB.Form FrmTrivia 
   BackColor       =   &H8000000E&
   Caption         =   "Form1"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   Picture         =   "FrmTrivia.frx":0000
   ScaleHeight     =   8250
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   8295
      Left            =   0
      Picture         =   "FrmTrivia.frx":630A2
      ScaleHeight     =   8235
      ScaleWidth      =   10875
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.CommandButton Command4 
         Caption         =   "Return to Main"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6720
         Picture         =   "FrmTrivia.frx":1A0574
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   7080
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8040
         Picture         =   "FrmTrivia.frx":1A1C42
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   7080
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Caption         =   "How to Play?"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6720
         Picture         =   "FrmTrivia.frx":1A3310
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6240
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000D&
         Caption         =   "Click To Begin!!!"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   3360
         MaskColor       =   &H008080FF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   6240
         UseMaskColor    =   -1  'True
         Width           =   3015
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "Test Your Twin's Trivia Knowledge!"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3735
         Left            =   1560
         TabIndex        =   1
         Top             =   1560
         Width           =   7455
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         Height          =   4455
         Left            =   1200
         TabIndex        =   2
         Top             =   1200
         Width           =   8175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00400000&
         Height          =   5175
         Left            =   840
         TabIndex        =   3
         Top             =   840
         Width           =   8895
      End
   End
End
Attribute VB_Name = "FrmTrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Twins Fan
'Form Name: FrmMain
'Project By: Stephanie Arel
'Date Written: 3/16/2009
'The purpose of this form is to take the user through a series of trivia questions.
'First the user clicks the start button, and is then taken to the questions.
'A running total is kept of their correct answers.
'When finished, the user is taken to a form that tells them their score.
Option Explicit


Private Sub Command1_Click()
'Brings user to a how to play form
FrmTrivia.Hide
FrmHow.Show
End Sub

Private Sub Command2_Click()
Dim mascot As String
Dim manager As String
Dim won As String
Dim Where As String
Dim seats As String



Dim Correct As Integer
'Message box to prepare the user!
MsgBox "Are you ready?!?", , "Let's Go!"

'Correct = 0 is the number which the user has correct. Each time the user gets another correct answer, "correct" increases by 1.
Correct = 0

'Question 1
mascot = InputBox("What is the Twins Mascot's Name?", "Question #1")
If mascot = "TC" Then
    MsgBox "That is Correct! Way to Go!", , "Yay!"
    Correct = Correct + 1
Else
    MsgBox "I'm sorry! TC is the twin's mascot.", , ":("
End If

'Question 2
manager = InputBox("Who is the Twin's manager?", "Question #2")
If manager = "Ron Gardenhire" Then
    MsgBox "Congrats! You are right!", , "Yay!"
    Correct = Correct + 1
ElseIf manager = "Gardenhire" Then
    MsgBox "Congrats! You are right!", , "Yay!"
    Correct = Correct + 1
Else
    MsgBox "I'm sorry! Ron Gardenhire is the Twin's Manager.", , ":("
End If

'Question 3
won = InputBox("How many world series titles have the twins won?", "Question #3")
If won = "2" Then
    MsgBox "Yay! You are good!", , "Yay!"
    Correct = Correct + 1
Else
    MsgBox "I'm sorry! The twins have won 2 titles! 1987 and 1991!", , ":("
End If

'Question 4
Where = InputBox("Where do the Twins currently play their home games?", "Question #4")
If Where = "Metrodome" Then
    MsgBox "You are right! You're on a roll!", , "Yay!"
    Correct = Correct + 1
ElseIf Where = "Dome" Then
    MsgBox "You are right! You're on a roll!", , "Yay!"
    Correct = Correct + 1
ElseIf Where = "HHH Metrodome" Then
    MsgBox "You are right! You're on a roll!", , "Yay!"
    Correct = Correct + 1
ElseIf Where = "Hubert H. Humphrey Metrodome" Then
    MsgBox "You are right! You're on a roll!", , "Yay!"
    Correct = Correct + 1
Else
    MsgBox "Sorry! The twins play in the Metrodome!", , ":("
End If

'Question 5
seats = InputBox("What color are the seats in the dome?", "Question #5")
    If seats = "blue" Then
    MsgBox "I am impressed! You are good!", , "Yay!"
    Correct = Correct + 1
ElseIf seats = "Blue" Then
    MsgBox "I am impressed! You are good!", , "Yay!"
    Correct = Correct + 1
Else
    MsgBox "I'm sorry! The seats are blue!", , ":("
End If

MsgBox "You've completed the quiz! Click ok to go to results!"


'Here, "correct" is totaled. For each different value of "correct" (0-5) the player is taken to a new form.
If Correct = 5 Then
    FrmTrivia.Hide
    Frm5.Show
ElseIf Correct = 4 Then
    FrmTrivia.Hide
    Frm4.Show
ElseIf Correct = 3 Then
    FrmTrivia.Hide
    Frm3.Show
ElseIf Correct = 2 Then
    FrmTrivia.Hide
    Frm2.Show
ElseIf Correct = 1 Then
    FrmTrivia.Hide
    Frm1.Show
Else: FrmTrivia.Hide
    Frm0.Show
    
End If

End Sub

Private Sub Command3_Click()
'Ends the program
End
End Sub

Private Sub Command4_Click()
'Takes User Back To Main Menu
FrmTrivia.Hide
FrmMain.Show
End Sub

Private Sub Picture1_Click()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
