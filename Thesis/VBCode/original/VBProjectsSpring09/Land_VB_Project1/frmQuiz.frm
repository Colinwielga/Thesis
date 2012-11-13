VERSION 5.00
Begin VB.Form frmQuiz 
   BackColor       =   &H00000080&
   Caption         =   "Quiz"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Brush Script MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdBreakingDawn 
      BackColor       =   &H00000080&
      Caption         =   "Click to quiz yourself on Breaking Dawn"
      BeginProperty Font 
         Name            =   "Brush Script MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdEclipse 
      BackColor       =   &H00000080&
      Caption         =   "Click to quiz yourself on Eclipse"
      BeginProperty Font 
         Name            =   "Brush Script MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdNewMoon 
      BackColor       =   &H00000080&
      Caption         =   "Click to quiz yourself on New Moon"
      BeginProperty Font 
         Name            =   "Brush Script MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdTwilight 
      BackColor       =   &H00000080&
      Caption         =   "Click to quiz yourself on Twilight"
      BeginProperty Font 
         Name            =   "Brush Script MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Image Image5 
      Height          =   2430
      Left            =   6120
      Picture         =   "frmQuiz.frx":0000
      Top             =   360
      Width           =   1635
   End
   Begin VB.Image Image3 
      Height          =   2445
      Left            =   4200
      Picture         =   "frmQuiz.frx":0F7F
      Top             =   360
      Width           =   1665
   End
   Begin VB.Image Image2 
      Height          =   2430
      Left            =   2280
      Picture         =   "frmQuiz.frx":4A86
      Top             =   360
      Width           =   1710
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   360
      Picture         =   "frmQuiz.frx":5C0F
      Top             =   360
      Width           =   1620
   End
End
Attribute VB_Name = "frmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Twilight
'Form Name: frmQuiz
'Author: Mollie Land
'Date Written: 3/22/09
'Objective: This code is designed to quiz the user on the books
'with this quiz, the user has the option of which book they would like to be quizzed on

'Dim global variables
Dim QuestionOne As String, QuestionTwo As String, QuestionThree As String

'This button quizzes on Breaking Dawn
Private Sub cmdBreakingDawn_Click()
    'Initialize the points value to reset it back to 0
    Points = 0
    
    'Ask a question and get an answer from the user using an Input Box
    QuestionOne = InputBox("Does Bella become a vampire?", "Breaking Dawn")
    
    'Use If/Then/Else to determine if the answer is correct or not
    'let the user know what the correct answer is if they got it incorrect
    If QuestionOne = "Yes" Then
        MsgBox "You are correct!", , "Breaking Dawn"
        Points = Points + 1
    ElseIf QuestionOne = "yes" Then
        MsgBox "You are correct!", , "Breaking Dawn"
        Points = Points + 1
    Else
        MsgBox "Sorry, you are incorrect. The answer is yes.", , "Breaking Dawn"
    End If
    
    'Use If/Then/Else to determine if the answer is correct or not
    'let the user know what the correct answer is if they got it incorrect
    QuestionTwo = InputBox("Do Bella and Edward get married?", "Breaking Dawn")
    
    If QuestionTwo = "Yes" Then
        MsgBox "You are correct!", , "Breaking Dawn"
        Points = Points + 1
    ElseIf QuestionTwo = "yes" Then
        MsgBox "You are correct!", , "Breaking Dawn"
        Points = Points + 1
    Else
        MsgBox "Sorry, you are incorrect. The correct answer is yes.", , "Breaking Dawn"
    End If
    
    'Use If/Then/Else to determine if the answer is correct or not
    'let the user know what the correct answer is if they got it incorrect
    QuestionThree = InputBox("Who is coming to see if Bella is a vampire?", "Breaking Dawn")
    
    If QuestionThree = "The Volturi" Then
        MsgBox "You are correct!", , "Breaking Dawn"
        Points = Points + 1
    ElseIf QuestionThree = "Volturi" Then
        MsgBox "You are correct!", , "Breaking Dawn"
        Points = Points + 1
    Else
        MsgBox "Sorry, you are incorrect. The correct answer is the Volturi.", , "Breaking Dawn"
    End If
    
     'print the user's total points using a message box
    If Points = 3 Then
        MsgBox "You know so much about Breaking Dawn! You received " & Points & " points!", , "Breaking Dawn"
    ElseIf Points = 2 Then
        MsgBox "Close, but not quite perfect! You received " & Points & " points!", , "Breakin Dawn"
    ElseIf Points = 1 Then
        MsgBox "You need to learn more about Breaking Dawn. You received " & Points & " point!", , "Breaking Dawn"
    Else
        MsgBox "You don't know anything about Breaking Dawn! You received " & Points & " points!", , "Breaking Dawn"
    End If
    
    
    
End Sub

'This button quizzes on Eclipse
Private Sub cmdEclipse_Click()
    'Initialize the points value
    Points = 0
    
    'Ask a question and get an answer from the user using an Input Box
    QuestionOne = InputBox("What is Jacob?", "Eclipse")
    
    'Use If/Then/Else to determine if the answer is correct or not
    'let the user know what the correct answer is if they got it incorrect
    If QuestionOne = "Werewolf" Then
        MsgBox "You are correct!", , "Eclipse"
        Points = Points + 1
    ElseIf QuestionOne = "wereworlf" Then
        MsgBox "You are correct!", , "Eclipse"
        Points = Points + 1
    Else
        MsgBox "Sorry, you are incorrect. The correct answer is Werewolf.", , "Eclipse"
    End If
    
    QuestionTwo = InputBox("Who are the new born vampires coming after?", "Eclipse")
    
    'Use If/Then/Else to determine if the answer is correct or not
    'let the user know what the correct answer is if they got it incorrect
    If QuestionTwo = "Bella" Then
        MsgBox "You are correct!", , "Eclipse"
        Points = Points + 1
    Else
        MsgBox "Sorry, you are incorrect. The correct answer is Bella.", , "Eclipse"
    End If
    
    QuestionThree = InputBox("Do the vampires and the wolves have to work together even though they are enemies?", "Eclipse")
    
    'Use If/Then/Else to determine if the answer is correct or not
    'let the user know what the correct answer is if they got it incorrect
    If QuestionThree = "Yes" Then
        MsgBox "You are correct!", , "Eclipse"
        Points = Points + 1
    ElseIf QuestionThree = "yes" Then
        MsgBox "You are correct!", , "Eclipse"
        Points = Points + 1
    Else
        MsgBox "Sorry, you are incorrect. The correct answer is yes.", , "Eclipse"
    End If
    
     'print the user's total points using a message box
    If Points = 3 Then
        MsgBox "You know so much about Eclipse! You received " & Points & " points!", , "Eclipse"
    ElseIf Points = 2 Then
        MsgBox "Close, but not quite perfect! You received " & Points & " points!", , "Eclipse"
    ElseIf Points = 1 Then
        MsgBox "You need to learn more about Eclipse. You received " & Points & " point!", , "Eclipse"
    Else
        MsgBox "You don't know anything about Eclipse! You received " & Points & " points!", , "Eclipse"
    End If
    
    
    
    
End Sub

'This button quizzes on New Moon
Private Sub cmdNewMoon_Click()
    'Initialize the points value
    Points = 0
    
    'Ask a question and get an answer from the user using an Input Box
    QuestionOne = InputBox("Who is Bella's new best friend?", "New Moon")
    
    'Use If/Then/Else to determine if the answer is correct or not
    'let the user know what the correct answer is if they got it incorrect
    If QuestionOne = "Jacob Black" Then
        MsgBox "You are correct!", , "New Moon"
        Points = Points + 1
    ElseIf QuestionOne = "Jacob" Then
        MsgBox "You are correct!", , "New Moon"
        Points = Points + 1
    Else
        MsgBox "Sorry, you are incorrect! The correct answer is Jacob Black.", , "New Moon"
    End If
    
    QuestionTwo = InputBox("Does Bella kill herself?", "New Moon")
    
    'Use If/Then/Else to determine if the answer is correct or not
    'let the user know what the correct answer is if they got it incorrect
    If QuestionTwo = "No" Then
        MsgBox "You are correct!", , "New Moon"
        Points = Points + 1
    ElseIf QuestionTwo = "no" Then
        MsgBox "You are correct!", , "New Moon"
        Points = Points + 1
    Else
        MsgBox "Sorry, you are incorrect! The correct answer is no.", , "New Moon"
    End If
    
    QuestionThree = InputBox("Which country do Bella and Alice have to go to in order to save Edward?", "New Moon")
    
    'Use If/Then/Else to determine if the answer is correct or not
    'let the user know what the correct answer is if they got it incorrect
    If QuestionThree = "Italy" Then
        MsgBox "You are correct!", , "New Moon"
        Points = Points + 1
    Else
        MsgBox "Sorry, you are incorrect! The correct answer is Italy.", , "New Moon"
    End If
    
    'print the user's total points using a message box
    If Points = 3 Then
        MsgBox "You know so much about New Moon! You received " & Points & " points!", , "New Moon"
    ElseIf Points = 2 Then
        MsgBox "Close, but not quite perfect! You received " & Points & " points!", , "New Moon"
    ElseIf Points = 1 Then
        MsgBox "You need to learn more about New Moon. You received " & Points & " point!", , "New Moon"
    Else
        MsgBox "You don't know anything about New Moon! You received " & Points & " points!", , "New Moon"
    End If
    
    
End Sub

Private Sub cmdReturn_Click()
    'return to start form
    frmStart.Show
    frmQuiz.Hide
End Sub
'This button quizzes on Twilight
Private Sub cmdTwilight_Click()
    'Initialize the Points
    Points = 0
    
    'Ask a question and get an answer from the user using an Input Box
    QuestionOne = InputBox("Does Edward turn Bella into a vampire?", "Twilight")
    
    'Use If/Then/Else to determine if the answer is correct or not
    'let the user know what the correct answer is if they got it incorrect
    If QuestionOne = "No" Then
        MsgBox "You are correct!", , "Twilight"
        Points = Points + 1
    ElseIf QuestionOne = "no" Then
        MsgBox "You are correct!", , "Twilight"
        Points = Points + 1
    Else
        MsgBox "Sorry, you are incorrect! The correct answer is no.", , "Twilight"
    End If
    
    QuestionTwo = InputBox("What city does Twilight take place in?", "Twilight")
    
    'Use If/Then/Else to determine if the answer is correct or not
    'let the user know what the correct answer is if they got it incorrect
    If QuestionTwo = "Forks" Then
        MsgBox "You are correct!", , "Twilight"
        Points = Points + 1
    Else
        MsgBox "Sorry, you are incorrect! The correct answer is Forks.", , "Twilight"
    End If
    
    QuestionThree = InputBox("Which two characters fall in love?", "Twilight")
    
    'Use If/Then/Else to determine if the answer is correct or not
    'let the user know what the correct answer is if they got it incorrect
    If QuestionThree = "Edward and Bella" Then
        MsgBox "You are correct!", , "Twilight"
        Points = Points + 1
    ElseIf QuestionThree = "Bella and Edward" Then
        MsgBox "You are correct!", , "Twilight"
        Points = Points + 1
    Else
        MsgBox "Sorry, you are incorrect! The correct answer is Bella and Edward", , "Twilight"
    End If
    
    'print the user's total points using a message box
    If Points = 3 Then
        MsgBox "You know so much about Twilight! You received " & Points & " points!", , "Twilight"
    ElseIf Points = 2 Then
        MsgBox "Close, but not quite perfect! You received " & Points & " points!", , "Twilight"
    ElseIf Points = 1 Then
        MsgBox "You need to learn more about Twilight. You received " & Points & " point!", , "Twilight"
    Else
        MsgBox "You don't know anything about Twilight! You received " & Points & " points!", , "Twilight"
    End If
    
End Sub
