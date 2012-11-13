VERSION 5.00
Begin VB.Form frmLevel1 
   BackColor       =   &H00008000&
   Caption         =   "Level 1"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetQuestion 
      Caption         =   "Get Question"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   2520
      Width           =   4215
   End
   Begin VB.TextBox txtAnswer 
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "Are you right?"
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   4080
      Width           =   2055
   End
   Begin VB.PictureBox picResult 
      BackColor       =   &H0080FF80&
      Height          =   1215
      Left            =   360
      ScaleHeight     =   1155
      ScaleWidth      =   6915
      TabIndex        =   0
      Top             =   1080
      Width           =   6975
   End
   Begin VB.Label lblDirections 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   $"frmLevel1.frx":0000
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   840
      Picture         =   "frmLevel1.frx":00E9
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   2400
   End
   Begin VB.Image imgQuit 
      Height          =   705
      Left            =   4560
      Picture         =   "frmLevel1.frx":53FE
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label lblAnswer 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Enter your answer here:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   3240
      Width           =   2055
   End
End
Attribute VB_Name = "frmLevel1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Counter As Integer, Questions As String, A As String, B As String, C As String
    Dim D As String, Answer As String, Q(1 To 100) As String, AnsA(1 To 100) As String
    Dim AnsB(1 To 100) As String, AnsC(1 To 100) As String, AnsD(1 To 100) As String, Ans(1 To 100) As String
    Dim QuestNum As Integer
    

Private Sub cmdA_Click()
    
    Dim response As String
    response = txtAnswer.Text                       'Answer to question
    
    If response = Ans(QuestNum) Then                'If answer is correct, then this will be printed
        MsgBox YourName & " You are correct!"
    Else
        MsgBox YourName & " Sorry, Try Again!"      'If answer is incorrect, then this will be printed
    End If
    
    txtAnswer.Text = ""
    
End Sub

Private Sub cmdGetQuestion_Click()
    If QuestNum < Counter Then                      'If Question Number is less then Counter then program continues
        picResult.Cls                               'Allows the picture box to erase previous question
        QuestNum = QuestNum + 1                     'QuestNum is same as Counter, it states which question is next
        picResult.Print Q(QuestNum)                 'Question is printes
        picResult.Print AnsA(QuestNum), AnsB(QuestNum)      '2 Answers to question is printed
        picResult.Print AnsC(QuestNum), AnsD(QuestNum)      'Rest of Answers to question is printed
    Else
        QuestNum = 0                                'When QuestNum reaches 0, this indicates that the questions have been answered
            MsgBox "You've finished the Quiz!"      'Appears to let player know they've finished
        frmLevel1.Visible = False                   'Level 1 disappears
        frmLevel2.Visible = True                    'Level 2 appears
    End If

End Sub


Private Sub Form_Load()
Dim Pos As Integer
    Open App.Path & "\Questions.txt" For Input As #1        'opens file
    Do While Not EOF(1)                 'completes program until the end of file
        Input #1, Questions, A, B, C, D, Answer     'inputs the question along with its answers and correct answer
        Counter = Counter + 1           'Counts and adds each time a new question is brought into the game
        Q(Counter) = Questions          'Puts number brought in from the file into the array
        AnsA(Counter) = A           'Puts number brought in from the file into the array
        AnsB(Counter) = B           'Puts number brought in from the file into the array
        AnsC(Counter) = C           'Puts number brought in from the file into the array
        AnsD(Counter) = D           'Puts number brought in from the file into the array
        Ans(Counter) = Answer       'Puts number brought in from the file into the array
    Loop                        'Continues process until end of file
    Close #1
End Sub

Private Sub imgQuit_Click()
    End             'ends program
End Sub
