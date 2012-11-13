VERSION 5.00
Begin VB.Form frmQuiz 
   Caption         =   "frmQuiz"
   ClientHeight    =   10845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   Picture         =   "frmBoat.frx":0000
   ScaleHeight     =   10845
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9840
      Width           =   1335
   End
   Begin VB.CommandButton cmdAnswers 
      BackColor       =   &H00000080&
      Caption         =   "See the Answers!!"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      Height          =   3855
      Left            =   360
      ScaleHeight     =   3795
      ScaleWidth      =   3315
      TabIndex        =   3
      Top             =   4320
      Width           =   3375
   End
   Begin VB.CommandButton cmdCrewQuiz 
      BackColor       =   &H00000080&
      Caption         =   "Take the Quiz"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to the Main Screen"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9000
      Width           =   2415
   End
   Begin VB.Label lblQuiz 
      BackStyle       =   0  'Transparent
      Caption         =   "Test Your Knowledge"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1455
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "frmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: CSB/SJU Crew
'Form name: frmMeettheMembers
'Authors: Lauren Nephew and Rachel Stalley
'Date: October 18th, 2009
'Objective:To give to user a quiz on the information given in other forms using input boxes.
Option Explicit

Private Sub cmdAnswers_Click()
picResults.Print 'this prints all of the answers letting the user know if they were correct
picResults.Print "********************ANSWERS**************************"
picResults.Print "#1 = Cox, Coxwain"
picResults.Print "#2 = Lake Sag"
picResults.Print "#3 = Stern"
picResults.Print "#4 = Backward"
picResults.Print "#5 = To stop"
picResults.Print "#6 = CoxBox"
picResults.Print "#7 = Morning"
picResults.Print "#8 = Forward"
picResults.Print "#9 = Star side"
picResults.Print "#10 = YES!!!!!!!!!!!"
End Sub

Private Sub cmdCrewQuiz_Click()
Dim Answer As String, Answer1 As String, Answer2 As String, Answer3 As String, Answer4 As String
Dim Answer5 As String, Answer6 As String, Answer7 As String, Answer8 As String, Answer9 As String
'these are input boxes for each of our questions
    Answer = InputBox("Who does not row, but is in the boat?", "Question 1")
    Answer1 = InputBox("What lake does the crew team practice on?", "Question 2")
    Answer2 = InputBox("What is the front of the boat called?", "Question 3")
    Answer3 = InputBox("What direction does the rower face?", "Question 4")
    Answer4 = InputBox("What does wain off mean?", "Question 5")
    Answer5 = InputBox("What does a cox use so that all of the rowers can hear him or her?", "Question 6")
    Answer6 = InputBox("What time of day does the crew team practice at?", "Question 7")
    Answer7 = InputBox("What direction does the cox face?", "Question 8")
    Answer8 = InputBox("What is the name of the rower's left side?", "Question 9")
    Answer9 = InputBox("Did you enjoy this quiz?", "Question 10")
    
End Sub

Private Sub cmdQuit_Click() 'ends the program
End
End Sub

Private Sub cmdReturn_Click() 'This allows the user to go back to the main menu screen
frmCSBSJUCrewMain.Show
frmQuiz.Hide
End Sub

