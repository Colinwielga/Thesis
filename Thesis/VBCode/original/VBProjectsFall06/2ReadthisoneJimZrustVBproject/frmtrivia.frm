VERSION 5.00
Begin VB.Form frmtrivia 
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdhome 
      BackColor       =   &H00400040&
      Caption         =   "Return to the Frontpage"
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   2535
   End
   Begin VB.CommandButton cmdstart 
      BackColor       =   &H0000FFFF&
      Caption         =   "Play The Trivia Challenge"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   5895
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   6000
      Left            =   0
      Picture         =   "frmtrivia.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6450
   End
End
Attribute VB_Name = "frmtrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Vikings Fan Page

'Form Name: Trivia

'Written by Jim Zrust

'Date: November 1, 2006

'Form objective: when i started making this program this was the form that i was most excited to make. all sports fans
'enjoy testing their knowledge when it comes to their favorite team and that is what this form does.
'i created a data file that this form accesses and it will ask a question and the user needs to attempt
'to answer this question. the only limitation is that the answer must be inputted by the user exactly how
'it was in the data file (proper spacing, spelling, capitalization)
Option Explicit

Private Sub cmdhome_Click() 'button to return home
frmtrivia.Hide
frmhome.Show
End Sub

Private Sub cmdstart_Click()
    Dim Questions(1 To 10) As String 'declare as array
    Dim Answers(1 To 10) As String
    Dim correct As Integer
    Dim user As String
    correct = 0
Open App.Path & "\trivia.txt" For Input As #1
    Do Until EOF(1)
        I = I + 1
        Input #1, Questions(I), Answers(I) 'input trivia file
    Loop
Close #1
    MsgBox "Let's Begin!", , "Start"
        For I = 1 To 10
            user = InputBox(Questions(I), "Question") 'inputbox will appear with the question displayed and the user will be able to enter his answer in the same box
            If user = Answers(I) Then 'compares "user" which is what the user inputted in the file to the answer in the array that corresponds to the question asked
                correct = correct + 1 'counter to count how many of the 10 answers the user got right
                MsgBox "You are Correct!", , "Way to go"
            Else
                correct = correct
                MsgBox "Sorry, Wrong Answer", , "Too Bad"
            End If
        Next I
    MsgBox "The Number of Questions Correct is:" & correct, , "Well Done!"
End Sub

