VERSION 5.00
Begin VB.Form frmTrivia 
   BackColor       =   &H80000007&
   Caption         =   "Trivia"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   Picture         =   "frmTrivia.frx":0000
   ScaleHeight     =   8580
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   4680
      TabIndex        =   2
      Top             =   6840
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   1920
      ScaleHeight     =   1035
      ScaleWidth      =   7035
      TabIndex        =   1
      Top             =   240
      Width           =   7095
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Start"
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Top             =   7920
      Width           =   2175
   End
End
Attribute VB_Name = "frmTrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmTrivia
'Author: Sean Johnson and Nick Lane
'Date Written: Monday March 17th, 2007
'Objective of form: this form uses arrays to create a trivia game which a user can use to
'                   gain extra discounts from their purchase.


Private Sub cmdBegin_Click()
Dim Questions(1 To 20) As String
Dim answers(1 To 20) As String
Dim ctr As Integer
Dim found As Boolean
Dim ans As String, name As String


name = InputBox("Please Enter your Name", "Welcome")    'allows user to enter his/her name
MsgBox "Type all answers in lower caps!"    'gives specific instructions on how to answe the trivia questions
found = False

ctr = 0

'opens the array of question to be used for the trivia game
Open App.Path & "\TriviaQArray.txt" For Input As #1

'reads the array into a parallel array
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Questions(ctr), answers(ctr)
Loop
Close #1


For j = 1 To ctr
    ans = InputBox(Questions(j), "Trivia Question") 'allows the questions in the array to be seen
        If ans = answers(j) Then    'this if loop lets the user know if his/her answer was correct
            found = True
            MsgBox "Correct"
        End If
        If ans <> answers(j) Then   'this if loop lets the user know if he/she answer was wrong
            found = False
            MsgBox "Incorrect Answer"
        End If
        If found = True Then    'this if loop keeps a count of how much correct answers the user got
            discount = discount + 1
        End If
Next j
        picResults.Print name; "you got "; discount; "correct"  'prints a messgae that lets the user know how many answer they got correct
        discount = discount * 5
        'calculates the amount of the discount in proportion to the amount of questions the user got correct
        picResults.Print " You have earned "; discount; " percent off"  'diplays the discount amount
        discount = discount / 100
        
cmdBegin.Enabled = False
End Sub
Private Sub cmdExit_Click()
'this button will hide this form and show the previous form
 frmTrivia.Hide
 frmSecondPage.Show
End Sub

Private Sub Form_Load()
MsgBox "You will be asked 12 questions. for each question correct you will recieve and additional 5% off your purchase. Do you want to continue?", vbYesNo
End Sub
