VERSION 5.00
Begin VB.Form Triviaform 
   Caption         =   "Form1"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   6570
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Height          =   1215
      Left            =   4560
      Picture         =   "Form3.frx":DBE3A
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   8
      Top             =   3360
      Width           =   1695
   End
   Begin VB.PictureBox Picture3 
      Height          =   1215
      Left            =   2400
      Picture         =   "Form3.frx":E422C
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   7
      Top             =   3360
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Height          =   1215
      Left            =   4560
      Picture         =   "Form3.frx":EC536
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   2400
      Picture         =   "Form3.frx":F4848
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   5
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H00FF80FF&
      Caption         =   "Go to Main Page"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF80FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Hard Questions!"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Easy Questions!!"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Trivia"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Triviaform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'allows user to answer either hard or easy questions
'in a Trivia game!


Option Explicit
Dim ctr As Single

Private Sub cmdMain_Click()
Triviaform.Hide
Welcomeform2.Show
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub Command1_Click()
Dim question1 As String
Dim question2 As String
Dim question3 As String
Dim question4 As String


question1 = InputBox("Which pet must be kept in water?", "Question 1")
    If question1 = "fish" Then
        ctr = ctr + 1
        MsgBox "Fish is a correct answer!!", , "Answer 1"
    Else
        MsgBox "Sorry,that answer is incorrect!!", , "Answer 1"
    End If

question2 = InputBox("Which pet loves to go for a walk?", "Question 2")
    If question2 = "dog" Then
        ctr = ctr + 1
        MsgBox "Dog is a correct answer", , "Answer 2"
    Else
        MsgBox "Sorry, that answer is incorrect.", , "Answer 2"
End If

question3 = InputBox("Which pet can live on earth and in the water?", "Question 3")
    If question3 = "turtle" Then
        ctr = ctr + 1
        MsgBox "Turtle, is a correct answer!!", , "Answer 3"
    Else
        MsgBox "Sorry, that answer is incorrect", , "Answer 3"
    End If

question4 = InputBox("Which pet is fluffy?", "Question 4")
    If question4 = "cat" Then
        ctr = ctr + 1
        MsgBox "Cat is a correct answer!!", , "Answer 4"
    Else
        MsgBox "Sorry, that answer is incorrect", , "Answer 4"
    End If



MsgBox "You have " & ctr & " answers correct!!"

End Sub

Private Sub Command2_Click()

Dim question5 As String
Dim question6 As String
Dim question7 As String
Dim question8 As String
ctr = 0

question5 = InputBox("Which pet barks?", "Question 1")
    If question5 = "dog" Then
        ctr = ctr + 1
        MsgBox "Dog, is a correct answer!!", , "Answer 1"
    Else
        MsgBox "Sorry, that answer is incorrect", , "Answer 1"
    End If

question6 = InputBox("Which pet loves to sleep?", "Question 2")
    If question6 = "cat" Then
        ctr = ctr + 1
        MsgBox "Cat, is a correct answer!!", , "Answer 2"
    Else
        MsgBox "Sorry, that answer is incorrect", , "Answer 2"
    End If

question7 = InputBox("Which pet eats on its own pleasure?", "Question 3")
    If question7 = "cat" Then
        ctr = ctr + 1
        MsgBox "Cat is a correct answer!!", , "Answer 3"
    Else
        MsgBox "Sorry, that answer is incorrect", , "Answer 3"
    End If

question8 = InputBox("Which pet loves to swim?", "Question 4")
    If question8 = "fish" Then
        ctr = ctr + 1
        MsgBox "Fish is a correct asnwer", , "Answer 4"
    Else
        MsgBox "Sorry, that answer is incorrect", , "Answer 4"
    End If




MsgBox "Congratulations!!You have  " & ctr & "  correct answers!!"



End Sub

Private Sub Form_Load()

End Sub
