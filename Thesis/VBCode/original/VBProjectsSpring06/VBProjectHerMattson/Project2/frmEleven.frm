VERSION 5.00
Begin VB.Form frmHerMattson11 
   BackColor       =   &H000000FF&
   Caption         =   "Answers"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdwordsearch 
      BackColor       =   &H80000009&
      Caption         =   "Search for Question by Entering a Key Word"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   10935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3075
      ScaleWidth      =   10875
      TabIndex        =   3
      Top             =   120
      Width           =   10935
   End
   Begin VB.CommandButton cmdShowAnswers 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Correct Answers to All Questions"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   10935
   End
   Begin VB.CommandButton cmdSpecific 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Correct Answer for Specific Question"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   10935
   End
   Begin VB.CommandButton cmdMainMenu 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   10935
   End
End
Attribute VB_Name = "frmHerMattson11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Globalize these variables because used in multiple command buttons, but the same arrays
'AmazingQuiz
'frmHerMattson11
'Ee Her and Jennifer Mattson
'Written on 3/22/06
'This menu allows the user to access answers from the quiz.

Dim Answer(1 To 100) As String
Dim Pos As Integer
Dim Questions(1 To 100) As String

Private Sub cmdMainMenu_Click()
    frmHerMattson11.Hide
    frmHerMattson1.Show
End Sub

Private Sub cmdShowAnswers_Click()
    'Display all results from an array
        Pos = 0
    'Open Array
    Open App.Path & "\Answers.txt" For Input As #1
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, Answer(Pos), Questions(Pos)
        'To show each Answer with its position:
         picResults.Print Pos; Questions(Pos); ""; Answer(Pos)
    Loop
    Close #1
End Sub

Private Sub cmdSpecific_Click()
    Dim SearchAnswer As Integer
    Dim InvalidResponse As Integer
    'Clear picture box; in case asked for all answers first
    picResults.Cls
    'Obtain a SearchAnswer by using an input box
    SearchAnswer = InputBox("What question would you like the answer to?", "Find Answer")
    'pos = size at this point since it has already been uploaded by the user
    If SearchAnswer <= Pos Then
        picResults.Print Questions(SearchAnswer); Answer(SearchAnswer)
    Else
       InvalidResponse = MsgBox("Incorrect Number", , "Please input a valid Number")
    End If
End Sub

Private Sub cmdwordsearch_Click()
'This command allows the user to input a key word and it will find the question and answer.
    Dim SearchName As String
    Dim X As Single
    Dim Pos As Integer
    Dim Found As Boolean
    SearchName = InputBox("Enter a key word from the question", "Key Word Search")
    For Pos = 1 To Size
        X = InStr(Questions(Pos), SearchName)
        If X <> 0 Then
            picResults.Print Questions(Pos), Answer(Pos)
            Found = True
        End If
    Next Pos
    If Found = False Then
        picResults.Print "No Match Found"
    End If
End Sub

Private Sub Form_Load()
'In order to load the numbers, in case the user decides to search for an answer before looking at all the answers
'This form will automatically load the Answers.txt file
   Open App.Path & "\Answers.txt" For Input As #1
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, Answer(Pos), Questions(Pos)
    Loop
    Close #1
    Size = Pos
End Sub
