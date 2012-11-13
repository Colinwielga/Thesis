VERSION 5.00
Begin VB.Form SafetyQuiz 
   BackColor       =   &H0000C000&
   Caption         =   "Safety Quiz"
   ClientHeight    =   3090
   ClientLeft      =   345
   ClientTop       =   1320
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   Picture         =   "SafetyQuiz.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "See how you did on the quiz"
      Height          =   1215
      Left            =   3360
      TabIndex        =   17
      Top             =   8400
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Click to do Question 10"
      Height          =   975
      Left            =   7920
      TabIndex        =   16
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Click to do Question 9"
      Height          =   1095
      Left            =   7920
      TabIndex        =   15
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Click to do Question 8"
      Height          =   975
      Left            =   7920
      TabIndex        =   14
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clilck to do Question 7"
      Height          =   1095
      Left            =   7920
      TabIndex        =   13
      Top             =   2040
      Width           =   2415
   End
   Begin VB.PictureBox picresults 
      Height          =   855
      Left            =   1800
      ScaleHeight     =   795
      ScaleWidth      =   7515
      TabIndex        =   12
      Top             =   7440
      Width           =   7575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click to do Question 6"
      Height          =   1095
      Left            =   7920
      TabIndex        =   11
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   1095
      Left            =   5520
      TabIndex        =   5
      Top             =   6000
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   975
      Left            =   5520
      TabIndex        =   4
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   975
      Left            =   5520
      TabIndex        =   3
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   5520
      TabIndex        =   2
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   5520
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Back 
      Caption         =   "Go back to main menu."
      Height          =   1215
      Left            =   6240
      TabIndex        =   0
      Top             =   8400
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   $"SafetyQuiz.frx":84E82
      Height          =   615
      Left            =   600
      TabIndex        =   18
      Top             =   0
      Width           =   9135
   End
   Begin VB.Label Label5 
      Caption         =   "Q5.  If there is a fire in your house should you try to get your favorite toy before getting out?"
      Height          =   1095
      Left            =   240
      TabIndex        =   10
      Top             =   6000
      Width           =   4935
   End
   Begin VB.Label Label4 
      Caption         =   "Q2. Is it a good idea to prepair for fire's before it happens?"
      Height          =   975
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Label Label3 
      Caption         =   "Q3. Should you throw items out the window while riding on the bus?"
      Height          =   975
      Left            =   240
      TabIndex        =   8
      Top             =   3360
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "Q4. If there is a tornado while you are in your car should you get into a ditch or ravine?"
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Q1. While taking shelter from a tornado is it a good idea to stay away from windows?"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   4815
   End
End
Attribute VB_Name = "SafetyQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim question1 As Boolean
Dim question2 As Boolean
Dim Question3 As Boolean
Dim question4 As Boolean
Dim question5 As Boolean
Dim question6 As Boolean
Dim question7 As Boolean
Dim Question8 As Boolean
Dim question9 As Boolean
Dim question10 As Boolean

Private Sub Back_Click()
SafetyQuiz.Hide
Form1.Show
End Sub

Private Sub Command1_Click()
Dim answer As String
question6 = False
answer = InputBox("Q.6 Should you ever go behind the bus to cross the street?")
 If answer = "no" Then
    question6 = True
    Else
    question6 = False
    End If
End Sub

Private Sub Command2_Click()
Dim answer As String
question7 = False
answer = InputBox("Q.7 After a tornado is over is there sometimes broken glass that you should be careful to avoid?")
If answer = "yes" Then
    question7 = True
    Else
    question7 = False
    End If
End Sub

Private Sub Command3_Click()
Question8 = False
Dim answer As String
answer = InputBox("Q.8 If you drop an item under the bus should you try to get it?")
If answer = "no" Then
    Question8 = True
    Else
    Question8 = False
    End If
End Sub

Private Sub Command4_Click()
question9 = False
Dim answer As String
answer = InputBox("Q.9 After you are out of a house that was on fire should you go back in for any reason?")
If answer = "no" Then
    question9 = True
    Else
    question9 = False
    End If
End Sub

Private Sub Command5_Click()
question10 = False
Dim answer As String
answer = InputBox("Q.10  Is a tornado watch when a tornado has been sighted?")
If answer = "no" Then
    question10 = True
    Else
    question10 = False
    End If
End Sub

Private Sub Command6_Click()
question1 = False
Dim score As Integer
Dim answer1 As String
Dim answer2 As String
Dim answer3 As String
Dim answer4 As String
Dim answer5 As String
answer1 = Text1
If Text1 = "yes" Then
    question1 = True
    Else
    question1 = False
    End If
answer2 = Text2
If Text2 = "yes" Then
    question2 = True
    Else
    question2 = False
    End If
answer3 = Text3
If Text3 = "no" Then
    Question3 = True
    Else
    Question3 = False
    End If
answer4 = Text4
If Text4 = "yes" Then
    question4 = True
    Else
    question4 = False
    End If
answer5 = Text5
If Text5 = "no" Then
    question5 = True
    Else
    question5 = False
    End If
score = 0
If question1 = True Then
    score = score + 1
    End If
If question2 = True Then
    score = score + 1
    End If
If Question3 = True Then
    score = score + 1
    End If
If question4 = True Then
    score = score + 1
    End If
If question5 = True Then
    score = score + 1
    End If
If question6 = True Then
    score = score + 1
    End If
If question7 = True Then
    score = score + 1
    End If
If Question8 = True Then
    score = score + 1
End If
If question9 = True Then
    score = score + 1
    End If
If question10 = True Then
    score = score + 1
End If
picresults.Cls
If score > 9 Then
picresults.Print "great job you got a perfect Score of "; score
    ElseIf score > 7 Then
    picresults.Print "your doing pretty good with a score of "; score
    ElseIf score > 4 Then
    picresults.Print "You should probobly get some more practice beacause your score was "; score
    ElseIf score > -1 Then
    picresults.Print "make sure to go through the information again for your safety. your score was "; score
    End If
End Sub

