VERSION 5.00
Begin VB.Form frmQuiz 
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "Johnson_frmQuiz.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton CmdBack1 
      Caption         =   "Back to Home"
      Height          =   975
      Left            =   2160
      TabIndex        =   3
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back"
      Height          =   1095
      Left            =   2160
      TabIndex        =   2
      Top             =   12120
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000E&
      Height          =   4335
      Left            =   480
      ScaleHeight     =   4275
      ScaleWidth      =   5955
      TabIndex        =   1
      Top             =   3360
      Width           =   6015
   End
   Begin VB.CommandButton cmdQuiz 
      Caption         =   "TAKE THE BREW CREW QUIZ, NOW!!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   6015
   End
End
Attribute VB_Name = "frmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Milwaukee Brewers Fan Club Program 2008

'Form Name: Brewers Trivia

'Author: Matthew Johnson

'Date Written: 11/4/2008

'Objective: In this program, I construct a trivia test that's kind of like a game for the
'user to test their knowledge of the Brewers.  It demonstrates my ability in using inputs
'from inputboxes, and also demonstrates my proficiency in using select case statements.

Private Sub cmdQuiz_Click()
'Here I declare the needed variables to make this part of the program to work.
Dim q1 As String, q2 As String, q3 As String, q4 As String, q5 As String, q6 As String, q7 As String, q8 As Integer, q9 As Integer, q10 As Integer, correctTotal As Integer

picResults.Cls

'Here I assign the variables q1-q10 to Input Boxes.  The user will enter an answer
'into the input box and it will be assigned to one of the variables.

q1 = InputBox("What brewer won the national league rookie of the year in 2007?", "Which brewer was he?")
q2 = InputBox("What brewer pitcher tore his ACL at the beginning of the the 2008 season versus the Cubs?", "Which brewer was he?")
q3 = InputBox("What brewer manager lost his job during the 2008 season?", "Who was the manager?")
q4 = InputBox("What Cy Young Award Winning Pitcher joined the brewers during the mid-season of 2008?", "Who was that pitcher?")
q5 = InputBox("Who was the youngest baseball player to ever hit over 50 homeruns in a single season?", "Hint: It was a brewer!")
q6 = InputBox("Who won both the American League MVP and the Cy Young Award in 1981?", "Hint: It was a brewer!")
q7 = InputBox("Which baseball team beat the Brewers in the 1981 World Series?", "Hint: the team is now in the same division as the Brewers.")
q8 = InputBox("In what year did Paul Molitor leave the Brewers?", "Enter a year")
q9 = InputBox("In what year did Miller Park Open?", "What year was it?")
q10 = InputBox("When did the Brewers begin their franchise?", "What year was it?")
correctTotal = 0

'If the above variables equal the correct answer, it adds a score of one to the variable
'called, "correctTotal".

If q1 = "Ryan Braun" Then
    correctTotal = correctTotal + 1
End If

If q2 = "Yovani Gallardo" Then
    correctTotal = correctTotal + 1
End If

If q3 = "Ned Yost" Then
    correctTotal = correctTotal + 1
End If

If InStr(q4, "Sabathia") Then           'I'm using string functions here :)
    correctTotal = correctTotal + 1
End If

If q5 = "Prince Fielder" Then
    correctTotal = correctTotal + 1
End If

If q6 = "Rollie Fingers" Then
    correctTotal = correctTotal + 1
End If

If InStr(q7, "Cardinals") Then          'I'm using string functions here :)
    correctTotal = correctTotal + 1
End If

If q8 = 1992 Then
    correctTotal = correctTotal + 1
End If

If q9 = 2000 Then
    correctTotal = correctTotal + 1
End If

If q10 = 1970 Then
    correctTotal = correctTotal + 1
End If

'Here I use select case statements to differentiate the trivia masters and the "no-nothings" about the brewers. :)
'It uses the "correctTotal", which is how many questions the user got right that will determine
'their achievement as a brewers trivia person.

Select Case correctTotal
    Case Is >= 9
        picResults.Print "You're a whiz, when it comes to Brewer trivia."
    Case 7 To 8
        picResults.Print "You're above average in Brewer trivia, by my standards."
    Case 5 To 6
        picResults.Print "You're average."
    Case 3 To 4
        picResults.Print "You failed!"
    Case 1 To 2
        picResults.Print "You barely know anything about the Brewers."
    Case Is = 0
        picResults.Print "You're beyond a failure! What are you?... A cubs fan!"
End Select

        
End Sub

'This allows the user to go back to the initial form
Private Sub cmdBack1_Click()
    frmIntro.Show
    frmQuiz.Hide
End Sub
