VERSION 5.00
Begin VB.Form frmHerMattsonEleven 
   Caption         =   "Answers"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "Return to Main Menu"
      Height          =   855
      Left            =   480
      TabIndex        =   3
      Top             =   5160
      Width           =   3615
   End
   Begin VB.CommandButton cmdSpecific 
      Caption         =   "Show Correct Answer for Specific Question"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   4320
      Width           =   3615
   End
   Begin VB.CommandButton cmdShowAnswers 
      Caption         =   "Show Correct Answers to All Questions"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   3240
      Width           =   3615
   End
   Begin VB.PictureBox picResults 
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2595
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmHerMattsonEleven"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Globalize these variables because used in multiple command buttons, but the same arrays
Dim Answer(1 To 100) As String
Dim pos As Integer

Private Sub cmdMainMenu_Click()
    frmHerMattsonEleven.Hide
    frmHerMattsonOne.Show
End Sub

Private Sub cmdShowAnswers_Click()
    'Display all results from an array
        pos = 0
    'Open Array
    Open App.Path & "\Answers.txt" For Input As #1
    Do Until EOF(1)
        pos = pos + 1
        Input #1, Answer(pos)
        'To show each Answer with its position:
        picResults.Print pos, Answer(pos)
    Loop
    Close #1
End Sub

Private Sub cmdSpecific_Click()
    Dim SearchAnswer As Single
    Dim InvalidResponse As Integer
    'Clear picture box; in case asked for all answers first
    picResults.Cls
    'Obtain a SearchAnswer by using an input box
    SearchAnswer = InputBox("What question would you like the answer to?", "Find Answer")
    'pos = size at this point since it has already been uploaded by the user
    If SearchAnswer <= pos Then
        picResults.Print SearchAnswer, Answer(SearchAnswer)
    Else
       InvalidResponse = MsgBox("Incorrect Number", , "Please input a valid Number")
    End If
End Sub

Private Sub Form_Load()
'In order to load the numbers, in case the user decides to search for an answer before looking at all the answers
'This form will automatically load the Answers.txt file
   Open App.Path & "\Answers.txt" For Input As #1
    Do Until EOF(1)
        pos = pos + 1
        Input #1, Answer(pos)
    Loop
    Close #1
End Sub
