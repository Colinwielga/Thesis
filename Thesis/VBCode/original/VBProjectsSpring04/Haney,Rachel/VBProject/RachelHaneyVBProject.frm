VERSION 5.00
Begin VB.Form RachelHaney1 
   BackColor       =   &H0000FFFF&
   Caption         =   "RachelHaney1"
   ClientHeight    =   3255
   ClientLeft      =   3540
   ClientTop       =   3180
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   5145
   Begin VB.CommandButton cmdPeople 
      Caption         =   "How many people are going on your vacation?"
      Height          =   975
      Left            =   3120
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdSpend 
      BackColor       =   &H00FFFF00&
      Caption         =   "How much do you want to spend?"
      Height          =   975
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H00800080&
      Caption         =   "Continue"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FF00FF&
      Caption         =   "Plan your very own five day vacation!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "RachelHaney1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'RachelHaney1 (RachelHaneyVBProject.frm)
'Rachel Haney 3/11/04
'This program will allow people to plan their vacation
'They can decide where they want to go, what transportation
'they will take to get there, where they will stay,
'what they would like to do while on the vacation,
'how much the trip will cost, and display the complete
'results of their chosen vacation.

Private Sub cmdContinue_Click()
'This button will allow the person to view the next form
'on this project by displaying the next form and hiding
'the current form
    RachelHaney1.Visible = False
    RachelHaney2.Visible = True
    cmdContinue.Visible = False
End Sub

Private Sub cmdPeople_Click()
'this code gets the number of people taking the vacation from the user
    People = InputBox("Enter the number of people who are going on the vacation.", "Number of People")
    cmdContinue.Visible = True
    cmdPeople.Visible = False
    MsgBox "Make sure you have entered information under both buttons before clicking continue", , "Pause"
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSpend_Click()
'this code lets the user determine their budget
    Spend = InputBox("Enter the amount of money you would like to spend on your vacation.", "Budget")
    cmdSpend.Visible = False
End Sub

Private Sub Form_Load()
    cmdContinue.Visible = False
End Sub
