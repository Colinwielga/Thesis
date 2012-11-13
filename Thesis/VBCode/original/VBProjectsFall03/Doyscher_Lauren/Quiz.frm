VERSION 5.00
Begin VB.Form QuizForm 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   5715
      TabIndex        =   16
      Top             =   4200
      Width           =   5775
   End
   Begin VB.CommandButton cmdQuiz 
      Caption         =   "Start Quiz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1560
      TabIndex        =   15
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton cmdMainForm 
      BackColor       =   &H000000FF&
      Caption         =   "Main Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6000
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Lauren Doyscher"
      Height          =   255
      Left            =   8880
      TabIndex        =   17
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"Quiz.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   480
      TabIndex        =   14
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kari Bruns"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   7080
      TabIndex        =   11
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Linnea Calderon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   10
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lauren Doyscher"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   7080
      TabIndex        =   9
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Heather Fischer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   7080
      TabIndex        =   8
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Elizabeth Gatschet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   7080
      TabIndex        =   7
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Heather Hampton"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   7080
      TabIndex        =   6
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sarah Henning"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   7080
      TabIndex        =   5
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bridget Javorski"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   7080
      TabIndex        =   4
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Jennifer Kruse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   7080
      TabIndex        =   3
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kathryn Ness"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   7080
      TabIndex        =   2
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Elaina Reinke"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   7080
      TabIndex        =   1
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kathleen Swart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   7080
      TabIndex        =   0
      Top             =   5640
      Width           =   2055
   End
End
Attribute VB_Name = "QuizForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: SophmoreDancers (VBProject.vbp)
'Form Name: QuizForm (Quiz.frm)
'Author: Lauren Doyscher
'Date Written: 10/27/03
'This form's purpose is to ask the user to answer five questions and rate the
'user according to their score.
Option Explicit

Private Sub cmdMainForm_Click()
'Brings you back to main page
QuizForm.Hide
MainForm.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdQuiz_Click()
Dim Answer1 As String
Dim Answer2 As String
Dim Answer3 As String
Dim Answer4 As String
Dim Answer5 As String
Dim Score As Integer
Dim Correct As Boolean
'Intitializes the value of Score to zero and Clears the screen
Score = 0
picResults.Cls
'This part of the form asks the user Five questions about the dancers and adds one to the score
'if the user was correct.  It shows what answers the user got wrong.
Answer1 = InputBox("Name a dancer whoes favorite color is red", "Question 1 of 5")
picResults.Print "Question 1",
    If Answer1 = "Sarah Henning" Or Answer1 = "Elizabeth Gatschet" Then
        Score = Score + 1
        picResults.Print "Correct"
        Else
        picResults.Print "Incorrect"
        End If
Answer2 = InputBox("Name one of the dancers who has danced the longest", "Question 2 of 5")
picResults.Print "Question 2",
    If Answer2 = "Kari Bruns" Or Answer2 = "Kathleen Swart" Then
        Score = Score + 1
        picResults.Print "Correct"
        Else
        picResults.Print "Incorrect"
    End If
Answer3 = InputBox("Name one girl who is on the Regional Team", "Question 3 of 5")
picResults.Print "Question 3",
    If Answer3 = "Bridget Javorski" Or Answer3 = "Linnea Calderon" Or Answer3 = "Elizabeth Gatschet" Or Answer3 = "Kathleen Swart" Or Answer3 = "Heather Hampton" Or Answer3 = "Heather Fischer" Then
        Score = Score + 1
        picResults.Print "Correct"
        Else
        picResults.Print "Incorrect"
    End If
Answer4 = InputBox("Name one girl who is on the National Team", "Question 4 of 5")
picResults.Print "Question 4",
    If Answer4 = "Kari Bruns" Or Answer4 = "Lauren Doyscher" Or Answer4 = "Sarah Henning" Or Answer4 = "Kathryn Ness" Or Answer4 = "Elaina Reinke" Or Answer4 = "Jennifer Kruse" Then
        Score = Score + 1
        picResults.Print "Correct"
        Else
        picResults.Print "Incorrect"
    End If
Answer5 = InputBox("What is the average number of years danced by the sophomores, rounded to the nearest whole number?", "Question 5 of 5")
picResults.Print "Question 5",
    If Answer5 = "14" Then
        Score = Score + 1
        picResults.Print "Correct"
        Else
        picResults.Print "Incorrect"
    End If
'This part shows the final score of the user and gives a statemtent about how well
'the user did according to their score.
picResults.Print "Your score is"; Score; "Out of 5"
Select Case Score
    Case Is = 5
        picResults.Print "Excellent!  You know the know the dancers really well!"
    Case Is = 4
        picResults.Print "Good Job!  You know the dancers pretty well!"
    Case Is = 3
        picResults.Print "Not bad!  But there is room for improvement!"
    Case Is = 2
        picResults.Print "Poor!  You should know more about the dancers!"
    Case Is = 1
        picResults.Print "Very Poor!  You don't know much about the dancers!"
    Case Is = 0
        picResults.Print "Awfull!  You need to go back and read more about the dancers!"
End Select
End Sub

