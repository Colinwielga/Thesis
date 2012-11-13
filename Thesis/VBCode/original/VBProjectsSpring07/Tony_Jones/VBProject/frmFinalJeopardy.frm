VERSION 5.00
Begin VB.Form frmFinalJeopardy 
   BackColor       =   &H00FF0000&
   Caption         =   "Final Jeopardy"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuestion 
      BackColor       =   &H0000FF00&
      Caption         =   "Write in your Question"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   3495
   End
   Begin VB.Label lblcited 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Taken from http://www.users.csbsju.edu/~irahal/"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7800
      TabIndex        =   1
      Top             =   6840
      Width           =   3015
   End
   Begin VB.Label lblHis1000 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   $"frmFinalJeopardy.frx":0000
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmFinalJeopardy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQuestion_Click()
    
    'Declaring local variables
    Dim QuestionFinal As String
    Dim CorrectFinal As String
    
    'Asking user for question and declaring correct question
    QuestionFinal = InputBox("Enter your question", "Question to Computer Science Professors for" & " " & Wager)
    CorrectFinal = CorrectQuestions(31)
    
    'Comparing user's question with correct question
    If LCase(QuestionFinal) = LCase(CorrectFinal) Then
        Winnings = Winnings + Wager
        MsgBox "That is the correct question", , "Correct Question"
    Else
        Winnings = Winnings - Wager
        MsgBox "That is incorrect. The correct question is" & " " & CorrectQuestions(31), , "Incorrect Question"
    End If
    
    'Shows and hides the forms
    frmFinalJeopardy.Hide
    frmCongrats.Show
    
    'Displays the final winnings on Congrats form
    If Winnings >= 0 Then
        frmCongrats.picTotal.Print FormatCurrency(Winnings, 0)
    Else
        frmCongrats.picTotal.Print FormatCurrency(0, 0)
    End If
    
End Sub
