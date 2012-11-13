VERSION 5.00
Begin VB.Form frmQuestions 
   BackColor       =   &H00FFFF80&
   Caption         =   "Form1"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAnswer 
      Height          =   1335
      Left            =   6000
      TabIndex        =   13
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CommandButton cmdQuestion2 
      Caption         =   "Click here to answer"
      Height          =   855
      Left            =   2760
      TabIndex        =   9
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuestion1 
      Caption         =   "Click here to answer"
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label lblQuest12 
      BackColor       =   &H00FFFF80&
      Caption         =   "3. Bethel"
      Height          =   375
      Left            =   6000
      TabIndex        =   17
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label bllQuest12 
      BackColor       =   &H00FFFF80&
      Caption         =   "2. Concordia"
      Height          =   495
      Left            =   6000
      TabIndex        =   16
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label lblQuest11 
      BackColor       =   &H00FFFF80&
      Caption         =   "1. St. Thomas"
      Height          =   375
      Left            =   6000
      TabIndex        =   15
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label lblQuest10 
      BackColor       =   &H00FFFF80&
      Caption         =   "Who just lost to the Johnnies in O/T?"
      Height          =   615
      Left            =   6120
      TabIndex        =   14
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label lblQuest9 
      BackColor       =   &H00FFFF80&
      Caption         =   "3. Hamline"
      Height          =   495
      Left            =   2880
      TabIndex        =   12
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lblQuest8 
      BackColor       =   &H00FFFF80&
      Caption         =   "2. St. Mary's"
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblQuest7 
      BackColor       =   &H00FFFF80&
      Caption         =   "1. St John's"
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblQuest6 
      BackColor       =   &H00FFFF80&
      Caption         =   "Who are the Pipers?"
      Height          =   735
      Left            =   2880
      TabIndex        =   8
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblQuest5 
      BackColor       =   &H00FFFF80&
      Caption         =   "Enter corresponding number"
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblQuest4 
      BackColor       =   &H00FFFF80&
      Caption         =   "3. Hamline"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lblQuest3 
      BackColor       =   &H00FFFF80&
      Caption         =   "2. St. Catherine"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lblQuest2 
      BackColor       =   &H00FFFF80&
      Caption         =   "1. St. Thomas"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblQuest1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Who has both men and women students but only womens sports?"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmQuestions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: quick Facts about the MIAC'
    'Form name:Questions
    'Author:Alec Beatty'
    'Written 10/18/2009'
    'Objective: to give trivia questions testing the knowledge they jsut learned.
Option Explicit

Private Sub cmdQuestion1_Click()

Dim Number As Single

Number = InputBox("Enter corresponding number", "Answer") 'Creating an inputbox'

If Number = 2 Then 'paramaters'
        MsgBox "Congrats, you payed attention", , "Notice"
    ElseIf Number <> 2 Then
        MsgBox "Nope Guess again", , "Notice"
    End If
End Sub

Private Sub cmdQuestion2_Click()
Dim Number As Single

Number = InputBox("Enter corresponding number", "Answer")

If Number = 3 Then
        MsgBox "Congrats, you payed attention", , "Notice"
    ElseIf Number <> 3 Then
        MsgBox "Nope Guess again", , "Notice"
    End If
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
    frmQuestions.Hide
    frmMIAC.Show
    
End Sub

Private Sub txtQuest_Change()


    
    
End If

End Sub

Private Sub txtAnswer_Change()
Dim Answer As Single
Answer = txtAnswer.Text

    If Answer = 1 Then
        MsgBox "You are smart", , "Notice"
    ElseIf Answer <> 1 Then
        MsgBox "Don't you know anything?", , "Notice"
    End If
    
End Sub
