VERSION 5.00
Begin VB.Form frmTrivia 
   BackColor       =   &H00004000&
   Caption         =   "Form1"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080FF80&
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5760
      Width           =   3495
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7560
      ScaleHeight     =   1635
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   3600
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   1680
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H0080FF80&
      Caption         =   "Start Scrubs Trivia!"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   3855
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   240
      ScaleHeight     =   5955
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   1080
      Width           =   7095
   End
End
Attribute VB_Name = "frmTrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Scrubs Project
'Scrubs Trivia Quiz (frmTrivia)
'Ann Boeckmann
'October 28, 2008
'The purpose of this form is to provide a fun quiz where fans can test their knowledge of the show



Private Sub cmdBack_Click()

frmTrivia.Hide
frmOptions.Show

End Sub

Private Sub cmdStart_Click()
Dim Answer1 As String, Answer2 As String, Answer3 As String, Answer4 As String, Answer5 As String, Answer6 As String
Dim Jack As String, Rowdy As String, Brain As String, Laverne As String, Keith As String, Turkleton As String
Dim CTR As Integer, Score As Single

CTR = 0 'set counter to keep track of correct answers

Jack = App.Path & "\coxfam.jpg"
Rowdy = App.Path & "\rowdy_steven.jpg"
Brain = App.Path & "\braintrust.jpg"
Laverne = App.Path & "\Laverne.jpg"
Keith = App.Path & "\Keith.jpg"
Turkleton = App.Path & "\turk_kelso.jpg"

'Each question is asked via input box.  A picture related to the content of the question appears
'in the picture box and changes with each question
 
 picResults.Cls
 picResults.Picture = LoadPicture(Jack)
Answer1 = InputBox("What is the name of Dr. Cox's son?", "Question 1")
 picResults.Cls
 picResults.Picture = LoadPicture(Rowdy)
Answer2 = InputBox("What is the name of the dog that Carla gave Turk and JD when she accidentally lost Rowdy?", "Question 2")
 picResults.Cls
 picResults.Picture = LoadPicture(Brain)
Answer3 = InputBox("The permanent members of the Brain Trust include the Janitor, Doug, Ted and who else?", "Question 3")
 picResults.Cls
 picResults.Picture = LoadPicture(Laverne)
Answer4 = InputBox("What is the name of the nurse who started working at Sacred Heart after Laverne's death who looked exactly like Laverne?", "Question 4")
 picResults.Cls
 picResults.Picture = LoadPicture(Keith)
Answer5 = InputBox("What is the first and last name of Elliot's ex-fiance?", "Question 5")
 picResults.Cls
 picResults.Picture = LoadPicture(Turkleton)
Answer6 = InputBox("What does Dr. Kelso think that Turk's last name is?", "Question 6")

'computes number of correct answers
If Answer1 = "Jack" Or Answer1 = "jack" Then
CTR = CTR + 1

If Answer2 = "Steven" Or Answer2 = "steven" Then
CTR = CTR + 1

If Answer3 = "Todd" Or Answer3 = "todd" Then
CTR = CTR + 1

If Answer4 = "Shirley" Or Answer4 = "shirley" Then
CTR = CTR + 1

If Answer5 = "Keith Dudemeister" Or Answer5 = "keith dudemeister" Then
CTR = CTR + 1

If Answer6 = "Turkleton" Or Answer6 = "turkleton" Then
CTR = CTR + 1

End If
    End If
        End If
            End If
                End If
                    End If
                    
                    
Score = CTR / 6 'computes score as a percentage

picOutput.Cls
picOutput.Print "Your score: "; CTR
picOutput.Print "                                  "
picOutput.Print "Your percentage correct: "; FormatPercent(Score)
picOutput.Print "                                  "

Select Case CTR 'assigns an appropriate medical postition or "rank" based on the user's score
Case 0 To 2
    picOutput.Print "Your Rank: Intern"
Case 3
    picOutput.Print "Your Rank: Resident"
Case 4
    picOutput.Print "Your Rank: Chief Resident"
Case 5
    picOutput.Print "Your Rank: Attending Physician"
Case 6
    picOutput.Print "Your Rank: Chief of Medicine"
End Select







 


End Sub

