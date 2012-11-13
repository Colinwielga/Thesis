VERSION 5.00
Begin VB.Form frmIreland2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Ireland 2"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form2"
   ScaleHeight     =   7980
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtScore 
      Height          =   855
      Left            =   8040
      TabIndex        =   12
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmd9 
      Caption         =   "Q9) A single good day of weather followed by days of bad weather is called?  1) Spot 2) Lucky Day 3) Patrick's Day 4) Pet Day"
      Height          =   1095
      Left            =   8400
      TabIndex        =   11
      Top             =   4800
      Width           =   2775
   End
   Begin VB.CommandButton cmd8 
      Caption         =   "Q8) What is the only indigenous reptile in Ireland? 1) Snake 2) Turtle 3) Frog      4) Newt"
      Height          =   615
      Left            =   4920
      TabIndex        =   10
      Top             =   7200
      Width           =   3015
   End
   Begin VB.CommandButton cmd10 
      Caption         =   $"frmIreland2.frx":0000
      Height          =   975
      Left            =   7800
      TabIndex        =   9
      Top             =   6120
      Width           =   3255
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "Q7) What is the largest county (by land) in Ireland?  1) Dublin 2) Limerick 3) Galway 4) Cork"
      Height          =   615
      Left            =   4080
      TabIndex        =   8
      Top             =   6240
      Width           =   3015
   End
   Begin VB.CommandButton cmd6 
      Caption         =   "Q6) Who was the first Saint to preach to Ireland?  1) Patrick 2) Augustine 3) Jesus 4) Abban"
      Height          =   735
      Left            =   1560
      TabIndex        =   7
      Top             =   7200
      Width           =   2895
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "Q5) What is the Irish National Symbol? 1) Clover  2) Dove 3) Olive 4) Harp"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   6360
      Width           =   2895
   End
   Begin VB.CommandButton cmd4 
      Caption         =   $"frmIreland2.frx":0098
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   5160
      Width           =   3135
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "Q3)  According to some historians, what percent of American Presidents have irish Ancesty in them?  1) 10%  2) 74% 3) 40% 4) 60%"
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   3960
      Width           =   2895
   End
   Begin VB.CommandButton cmdResults 
      Caption         =   "See How You Did!"
      Height          =   615
      Left            =   9120
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Q2) Which one is a real place in Ireland?  1) Walsau   2) Devon  3) Andover  4) Muckanaghederdauhaulia"
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   2640
      Width           =   3015
   End
   Begin VB.CommandButton cmdGoHome 
      Caption         =   "Go Home"
      Height          =   375
      Left            =   10080
      TabIndex        =   1
      Top             =   7440
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Q1)    What are the colors of the Irish Flag? 1) White, Green   2) Green, Yellow, Black 3) Black, Green  4) Green, White, Oragne"
      Height          =   1215
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblScore 
      Caption         =   "Score Box ==> (Keep changing as you get right answers)"
      Height          =   855
      Left            =   6720
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   8025
      Left            =   0
      Picture         =   "frmIreland2.frx":012B
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmIreland2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project name:Information About Ireland
'Form Name: Ireland2
'Author: Rachel Lietzke
'Date Written: March 27, 2008
'Objective: To give a Multiple Chioce Test on You Fun Fact Knowledge of
'Ireland and to examin your score

Private Sub cmd1_Click()
Dim Answer As Integer

Answer = InputBox("Enter your Answer (Will only Accept 1, 2, 3, 4)", "Question 1")

If Answer = 4 Then
    MsgBox "You are Right add one to the Score Box!", , "You are Right!"
Else
    MsgBox "You are Wrong, Do not add anything to the Score box.", , "You are Wrong!"
End If

End Sub

Private Sub cmd10_Click()
Dim Ansewer As Integer

Answer = InputBox("Enter your Answer (Will only Accept 1, 2, 3, 4)", "Question 10")

Select Case Answer
    Case Is = 4
        MsgBox "You are Right add one to the Score Box!", , "You are Right!"
    Case Is = 3
        MsgBox "You are Wrong, Do not add anything to the Score box.", , "You are Wrong!"
    Case Is = 2
        MsgBox "You are Wrong, Do not add anything to the Score box.", , "You are Wrong!"
    Case Is = 1
        MsgBox "You are Wrong, Do not add anything to the Score box.", , "You are Wrong!"
End Select
End Sub

Private Sub cmd2_Click()
Dim Answer As Integer

Answer = InputBox("Enter your Answer (Will only Accept 1, 2, 3, 4)", "Question 2")

If Answer = 4 Then
    MsgBox "You are Right add one to the Score Box!", , "You are Right!"
Else
    MsgBox "You are Wrong, Do not add anything to the Score box.", , "You are Wrong!"
End If

End Sub

Private Sub cmd3_Click()
Dim Answer As Integer

Answer = InputBox("Enter your Answer (Will only Accept 1, 2, 3, 4)", "Question 3")

If Answer = 4 Then
    MsgBox "You are Right add one to the Score Box!", , "You are Right!"
Else
    MsgBox "You are Wrong, Do not add anything to the Score box.", , "You are Wrong!"
End If

End Sub

Private Sub cmd4_Click()
Dim Answer As Integer

Answer = InputBox("Enter your Answer (Will only Accept 1, 2, 3, 4)", "Question 4")

If Answer = 4 Then
    MsgBox "You are Right add one to the Score Box!", , "You are Right!"
Else
    MsgBox "You are Wrong, Do not add anything to the Score box.", , "You are Wrong!"
End If
End Sub

Private Sub cmd5_Click()
Dim Answer As Integer

Answer = InputBox("Enter your Answer (Will only Accept 1, 2, 3, 4)", "Question 5")

If Answer = 4 Then
    MsgBox "You are Right add one to the Score Box!", , "You are Right!"
Else
    MsgBox "You are Wrong, Do not add anything to the Score box.", , "You are Wrong!"
End If
End Sub

Private Sub cmd6_Click()
Dim Answer As Integer

Answer = InputBox("Enter your Answer (Will only Accept 1, 2, 3, 4)", "Question 6")

If Answer = 4 Then
    MsgBox "You are Right add one to the Score Box!", , "You are Right!"
Else
    MsgBox "You are Wrong, Do not add anything to the Score box.", , "You are Wrong!"
End If
End Sub

Private Sub cmd7_Click()
Dim Ansewer As Integer

Answer = InputBox("Enter your Answer (Will only Accept 1, 2, 3, 4)", "Question 7")

Select Case Answer
    Case Is = 4
        MsgBox "You are Right add one to the Score Box!", , "You are Right!"
    Case Is = 3
        MsgBox "You are Wrong, Do not add anything to the Score box.", , "You are Wrong!"
    Case Is = 2
        MsgBox "You are Wrong, Do not add anything to the Score box.", , "You are Wrong!"
    Case Is = 1
        MsgBox "You are Wrong, Do not add anything to the Score box.", , "You are Wrong!"
End Select

End Sub

Private Sub cmd8_Click()
Dim Ansewer As Integer

Answer = InputBox("Enter your Answer (Will only Accept 1, 2, 3, 4)", "Question 8")

Select Case Answer
    Case Is = 4
        MsgBox "You are Right add one to the Score Box!", , "You are Right!"
    Case Is = 3
        MsgBox "You are Wrong, Do not add anything to the Score box.", , "You are Wrong!"
    Case Is = 2
        MsgBox "You are Wrong, Do not add anything to the Score box.", , "You are Wrong!"
    Case Is = 1
        MsgBox "You are Wrong, Do not add anything to the Score box.", , "You are Wrong!"
End Select
End Sub

Private Sub cmd9_Click()
Dim Ansewer As Integer

Answer = InputBox("Enter your Answer (Will only Accept 1, 2, 3, 4)", "Question 9")

Select Case Answer
    Case Is = 4
        MsgBox "You are Right add one to the Score Box!", , "You are Right!"
    Case Is = 3
        MsgBox "You are Wrong, Do not add anything to the Score box.", , "You are Wrong!"
    Case Is = 2
        MsgBox "You are Wrong, Do not add anything to the Score box.", , "You are Wrong!"
    Case Is = 1
        MsgBox "You are Wrong, Do not add anything to the Score box.", , "You are Wrong!"
End Select
End Sub

Private Sub cmdGoHome_Click()
frmIreland2.Hide
frmIreland1.Show
End Sub

Private Sub cmdResults_Click()
Dim Score As Integer

Score = txtScore.Text

Select Case Score
    Case Is = 10
        MsgBox "Your knowledge of Ireland is amazing!", , "Perfect"
    Case Is = 9
        MsgBox "Your knowledge of Ireland is amazing!", , "Almost Perfect"
    Case Is = 8
        MsgBox "You know a lot about Ireland, Good Job!", , "Good Job"
    Case Is = 7
        MsgBox "You know a lot about Ireland, Good Job!", , "Good Job"
    Case Is = 6
        MsgBox "Not bad but you shoudl brush up on your knowledge.", , "Good"
    Case Is = 5
        MsgBox "Not bad but you shoudl brush up on your knowledge.", , "Good"
    Case Is = 4
        MsgBox "You should brush up on you knowledge.", , "OK"
    Case Is < 4
        MsgBox "You should go back to school.", , "Go To School"
End Select

End Sub
