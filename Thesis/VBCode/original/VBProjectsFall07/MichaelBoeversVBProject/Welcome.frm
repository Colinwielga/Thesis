VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H8000000D&
   Caption         =   "Welcome to Family Feud"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtnames 
      Height          =   615
      Left            =   2040
      TabIndex        =   5
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   5160
      MaskColor       =   &H8000000F&
      TabIndex        =   4
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton CmdQuestion3 
      Caption         =   "Load Question 3"
      Height          =   615
      Left            =   5160
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuestion2 
      Caption         =   "Load Question 2"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton cmdquestion1 
      Caption         =   "Load Question 1"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Cmdplay 
      Caption         =   "PLAY!!!"
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label lbldirections 
      BackColor       =   &H00FFC0C0&
      Caption         =   "    Input your name in the box         below load a question and                press play to begin."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1455
      Left            =   1080
      TabIndex        =   7
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Label lblwelcome 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Welcome to Family Feud!!!"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmdplay_Click()
Names = txtnames    'Store user's name
frmGame.Show        'Switches forms
frmWelcome.Hide
MsgBox (Names & " get ready to play Family Feud!") 'Preparing User to start game
End Sub

Private Sub cmdquestion1_Click()
Question = 1 'tells the other form which question will be asked.
CTR = 0
Open App.Path & "\question1.txt" For Input As #1    'Opening notepad text file

Do Until EOF(1)     'Loading array with notepad info
    CTR = CTR + 1
    Input #1, Answer(CTR), Answernum(CTR)
Loop

Close #1    'closing notepad text file
End Sub

Private Sub cmdQuestion2_Click()
Question = 2 'tells the other form which question will be asked.
CTR = 0
Open App.Path & "\question2.txt" For Input As #2    'Opening notepad text file

Do Until EOF(2)     'Loading array with notepad info
    CTR = CTR + 1
    Input #2, Answer(CTR), Answernum(CTR)
Loop

Close #2    'closing notepad text file
End Sub

Private Sub CmdQuestion3_Click()
Question = 3 'tells the other form which question will be asked.
CTR = 0
Open App.Path & "\question3.txt" For Input As #3    'Opening notepad text file

Do Until EOF(3)     'Loading array with notepad info
    CTR = CTR + 1
    Input #3, Answer(CTR), Answernum(CTR)
Loop

Close #3    'closing notepad text file
End Sub

Private Sub CmdQuit_Click()
End     'Exiting Program
End Sub
