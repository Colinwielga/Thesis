VERSION 5.00
Begin VB.Form frmQuiz 
   BackColor       =   &H0080FFFF&
   Caption         =   "An alien quiz..."
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFinish 
      Caption         =   "Finish Quiz?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7200
      TabIndex        =   14
      Top             =   6000
      Width           =   1815
   End
   Begin VB.TextBox txtShips 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   3960
      TabIndex        =   13
      Top             =   6000
      Width           =   2895
   End
   Begin VB.CommandButton cmdNext1 
      Caption         =   "Next question..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7200
      TabIndex        =   10
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox txtPopulation 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   3840
      TabIndex        =   7
      Top             =   3960
      Width           =   3015
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next question..."
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7200
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtCapital 
      Height          =   975
      Left            =   3840
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label lblQ3a 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a number 1 - 600"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   12
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label lblQ3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Guess how many alien ships invaded today?  Correct answer is within 100 and there were no more than 600..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   720
      TabIndex        =   11
      Top             =   5400
      Width           =   3015
   End
   Begin VB.Label lblQ2b 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Example: for 1.5 billion, enter 1.5"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label lblQ2a 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the exact number.  Don't include 'billion'"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label lblQ2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "If earth had 6.5 billion people before the attack, and we wiped out 4/5 of the population, how many people are still alive?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   720
      TabIndex        =   6
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "(Please type exactly how it is above:)"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Holalula, Honaluna, or Honolulu"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Earlier today, us aliens destroyed the state of Hawaii.  What was it's capital?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "A Quiz For The Humans:"
      BeginProperty Font 
         Name            =   "Minion Pro Med"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Capital As String
Dim Population As Single
Dim Ships As Single

Private Sub cmdFinish_Click()   'click to answer the last question and finish the quiz.  if correct, you dont fight, if wrong, you do
    Ships = txtShips.Text
    Select Case Ships
        Case Is > 600
            MsgBox ("WRONG!  Now you will fight me!"), , ("Wrong!")
            frmFight.Show
            frmQuiz.Hide
        Case Is >= 500
            MsgBox ("Correct!  Now I will let you go free.  You are a wise human."), , ("Correct!")
            frmtunnel.Show
            frmQuiz.Hide
            MsgBox ("you see a tunnel and enter."), , ("Continuing...")
        Case 0 To 500
            MsgBox ("WRONG!  Now you will fight me!"), , ("Wrong!")
            frmFight.Show
            frmQuiz.Hide
        Case Else
            MsgBox ("So you think a negative number of ships attacked.  I will kill you now because of your stupidity!"), , ("Wrong!")
             frmFight.Show
            frmQuiz.Hide
    End Select
        
        
End Sub

Private Sub cmdNext_Click()     'answer the first question.  if correct, go to next, if wrong, fight him
    Capital = txtCapital.Text
    If Capital = "Honolulu" Then
        MsgBox ("Good, I see you know geography..."), , ("Correct!")
        cmdNext1.Enabled = True
        txtPopulation.Enabled = True
        lblQ2.Enabled = True
        lblQ2a.Enabled = True
        lblQ2b.Enabled = True
    Else
        MsgBox ("WRONG! Now you will fight me!"), , ("Wrong!")
        frmFight.Show
        frmQuiz.Hide
    End If
    
End Sub

Private Sub cmdNext1_Click()    'answer question 2, if right, go to 3, if wrong, fight him
    Population = txtPopulation.Text
    If Population = 1.3 Then
        MsgBox ("Correct!  You are a smart human, now answer the last question."), , ("Corect!")
        cmdFinish.Enabled = True
        txtShips.Enabled = True
        lblQ3.Enabled = True
        lblQ3a.Enabled = True
            
    Else
        MsgBox ("WRONG!  Now you will fight me!"), , ("Wrong!")
        frmFight.Show
        frmQuiz.Hide
    End If

End Sub

