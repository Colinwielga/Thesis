VERSION 5.00
Begin VB.Form frmStarting 
   BackColor       =   &H00FF0000&
   Caption         =   "Jeopardy"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8400
   FillColor       =   &H00FF0000&
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4200
      TabIndex        =   1
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1440
      TabIndex        =   0
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   3765
      Left            =   1560
      Picture         =   "frmStarting.frx":0000
      Top             =   240
      Width           =   4980
   End
End
Attribute VB_Name = "frmStarting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
    
    'Declaring local variables
    Dim Ctr As Integer
    
    'Setting initial value of counter
    Ctr = 0
    
    'Opening the data file into an array
    Open App.Path & "/CorrectQuestions.txt" For Input As #1
    
    'Reading the data file into an array
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, CorrectQuestions(Ctr)
    Loop
    
    'Closing the data file
    Close #1
    
    'Showing character form but hiding starting form
    frmCharacter.Show
    frmStarting.Hide
    
End Sub

Private Sub cmdQuit_Click()
    
    'Stopping the program
    End
    
End Sub

