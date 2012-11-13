VERSION 5.00
Begin VB.Form frmQuiz 
   BackColor       =   &H000000FF&
   Caption         =   "Coach Quiz"
   ClientHeight    =   10740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10740
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHint 
      Caption         =   "Hint (read before starting the quiz)"
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Top             =   7440
      Width           =   4575
   End
   Begin VB.CommandButton cmdBack2 
      Caption         =   "Back to Main Screen"
      Height          =   735
      Left            =   5400
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "START THE QUIZ"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   11055
   End
End
Attribute VB_Name = "frmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form: Quiz
    'This page allows the user to take a quiz base on some of the information from the coaches page, and also some general knowledge and nonesense questions thrown in for good measure.
    'The quiz is given in a series of input boxes and 1 point is earned for each question answered correctly.
    'There is a button that allows the user to view a hint that will help them accurately answer the quiz questions.



Option Explicit

Private Sub cmdBack2_Click()               'Navigates to main screen (frmFirst)

frmFirst.Show
frmQuiz.Hide
frmPlayers.Hide
frmCoaches.Hide
End Sub

Private Sub cmdHint_Click()                 'Shows user a hint to help him or her answer the questions accurately.
MsgBox "The answers to the questions, unless specified, is either Mark or Thomas.", , "You're Welcome"
End Sub

Private Sub cmdStart_Click()                'This button begins the quiz, which is a series of input boxes, and keeps track of the number of correct responses.
Dim answer1 As String, answer2 As String, answer3 As String, answer4 As String, answer5 As String, answer6 As String, answer7 As String



answer1 = InputBox("Which coach is originally from New Jersey?", "Question 1")
If answer1 = "Mark" Then
    MsgBox "Correct.", , "GOOD!"
    points = points + 1
Else
    MsgBox "Sorry, you got it wrong.", , "WAH WAH"
End If
answer2 = InputBox("Which coach's playing career was cut short be a knee injury?", "Question 2")
If answer2 = "Thomas" Then
  MsgBox "Correct.", , "GOOD!"
  points = points + 1
Else
    MsgBox "Sorry, you got it wrong.", , "WAH WAH"
End If
answer3 = InputBox("Do you like the coaches' apparel? (yes or no)", "Question 3")
If answer3 = "yes" Then
    MsgBox "Correct.", , "GOOD!"
    points = points + 1
Else
    MsgBox "Sorry, you got it wrong.", , "WAH WAH"
End If
answer4 = InputBox("Which coach's facial hair is far superiour to the others'?", "Question 4")
If answer4 = "Thomas" Then
    MsgBox "Correct.", , "GOOD!"
    points = points + 1
Else
    MsgBox "Sorry, you got it wrong.", , "WAH WAH"
End If
answer5 = InputBox("Is St. John's Lacrosse going to win a National Title this year (yes or no)?", "Question 5")
If answer5 = "yes" Then
    MsgBox "Correct.", , "GOOD!"
    points = points + 1
Else
    MsgBox "Sorry, you got it wrong.", , "WAH WAH"
End If
answer6 = InputBox("Which coach would win in a fight?", "Question 6")
If answer6 = "Thomas" Then
    MsgBox "Correct.", , "GOOD!"
    points = points + 1
Else
    MsgBox "Sorry, you got it wrong.", , "WAH WAH"
End If
answer7 = InputBox("Which coach has more experience both playing and coaching?", "Question 7")
If answer7 = "Mark" Then
    MsgBox "Correct.", , "GOOD!"
    points = points + 1
Else
    MsgBox "Sorry, you got it wrong.", , "WAH WAH"
End If

Select Case points                  'This portion shows the quiz score and encourages the user to do better if their quiz score was not satisfactory to the program creator.
    Case 7
    MsgBox "Great job, " & user_name & ", you aced the quiz!! " & points & " of 7 questions answered correctly.", , "NICE WORK!"
    Case 6
    MsgBox "Not Bad, " & user_name & ", you almost had all questions right." & points & " of 7 questions answered correctly.", , "Not Too Shabby!"
    Case 5
    MsgBox "Pretty good, " & user_name & ", but you could use some work!!" & points & " of 7 questions answered correctly.", , "Gettin' There"
    Case 4
    MsgBox "Really, " & user_name & ", that's the best you could do??!!" & points & " of 7 questions answered correctly.", , "Meh, **yawn**, I've seen better."
    Case Is < 4
    MsgBox "You really need to learn more, " & user_name & ", it would do you good." & points & " of 7 questions answered correctly.", , "Sad Day..... :-("
End Select
End Sub

