VERSION 5.00
Begin VB.Form frmHerMattson1 
   BackColor       =   &H0080FF80&
   Caption         =   "Main Menu"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmtTest 
      Left            =   840
      Top             =   3120
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   1800
      ScaleHeight     =   2595
      ScaleWidth      =   5355
      TabIndex        =   6
      Top             =   360
      Width           =   5415
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF8080&
      Caption         =   "Quit"
      Height          =   615
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdPersonalInFo 
      BackColor       =   &H00FFFF00&
      Caption         =   "Personal Information"
      Height          =   615
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H000080FF&
      Caption         =   "Search Results"
      Height          =   615
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdAnswers 
      BackColor       =   &H000000FF&
      Caption         =   "Answers"
      Height          =   615
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCompute 
      BackColor       =   &H0000FFFF&
      Caption         =   "Compute Results"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuiz 
      BackColor       =   &H00FF00FF&
      Caption         =   "Quiz"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmHerMattson1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AmazingQuiz
'frmHerMattson1
'Ee Her and Jennifer Mattson
'Written Sunday 3/12/06
'This is our main menu. It provides links to the quiz, personal information, search engine, answers, and quit.
'The purpose of the project is to have users take a timed quiz and stores their information. It allows the user to search for correct answers and it provides their time and scores.


Private Sub cmdAnswers_Click()
'This will move to another form to show answers either all of them or specific ones
    frmHerMattson1.Hide
    frmHerMattson11.Show
End Sub

Private Sub cmdCompute_Click()
      'To compute the user's score
    picResults.Print "You answered"; Counter; "right"; " out of 8"
    picResults.Print "You finished the test in "; elapsedTime / 1000; " second"
End Sub

Private Sub cmdPersonalInFo_Click()
    frmHerMattson1.Hide
    frmHerMattson2.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdQuiz_Click()
    frmHerMattson1.Hide
    frmHerMattson3.Show
End Sub

Private Sub cmdSearch_Click()
    frmHerMattson1.Hide
    frmHerMattson12.Show
End Sub

Private Sub Command1_Click()
    picResults.Print elapsedTime
End Sub

Private Sub Form_Load()
'Initialize timer only starts when frm two is open.
    tmtTest.Enabled = False
    run = False
    elapsedTime = 0
End Sub

Private Sub tmtTest_Timer()
'The timer is set to run only during the quiz.
    If run = True Then
    elapsedTime = elapsedTime + tmtTest.Interval
    End If
End Sub
