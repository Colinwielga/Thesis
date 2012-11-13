VERSION 5.00
Begin VB.Form frmHerMattson7 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Question #5"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option3 
      Caption         =   "St. John's Abbey"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      Picture         =   "frmSeven.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      Caption         =   "The Great Hall"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   2400
      Picture         =   "frmSeven.frx":0A59
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Alquin Library"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   4680
      Picture         =   "frmSeven.frx":1C9A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Text            =   "Where is the SJU dungeon located?"
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "frmHerMattson7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AmazingQuiz
'frmHerMattson7
'Ee Her and Jennifer Mattson
'This is the fifth quiz question.
'Written on 3/19/06

Dim RightAnswer As Boolean
Dim Answer As String

Private Sub Option1_Click()
    RightAnswer = True
    Counter = Counter + 1
    Answer = MsgBox("Congratulations! You got the answer right!", , "Right Answer")
    frmHerMattson7.Hide
    frmHerMattson8.Show
End Sub

Private Sub Option2_Click()
    RightAnswer = False
    Answer = MsgBox("Wrong Answer, the correct Answer was Alquin Library", , "Wrong")
    frmHerMattson7.Hide
    frmHerMattson8.Show
End Sub

Private Sub Option3_Click()
    RightAnswer = False
    Answer = MsgBox("Wrong Answer, the correct Answer was Alquin Library", , "Wrong")
    frmHerMattson7.Hide
    frmHerMattson8.Show
End Sub
