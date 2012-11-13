VERSION 5.00
Begin VB.Form frmHerMattson9 
   BackColor       =   &H00FFFF80&
   Caption         =   "Question #7"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "The DaVinci Code"
      Height          =   2655
      Left            =   5400
      Picture         =   "frmNine.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "BrokeBack Mountain"
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   2880
      Picture         =   "frmNine.frx":0F73
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0000FF00&
      Caption         =   "Memoirs of Geisha"
      Height          =   2655
      Left            =   120
      Picture         =   "frmNine.frx":18BC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Text            =   "Which book will be made into a movie?"
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "frmHerMattson9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AmazingQuiz
'frmHerMattson9
'Ee Her and Jennifer Mattson
'This is the seventh quiz question.
'Written on 3/19/06

Dim RightAnswer As Boolean
Dim Answer As String

Private Sub Option1_Click()
    RightAnswer = False
    Answer = MsgBox("Wrong Answer, the correct Answer was The DaVinci Code", , "Wrong")
    frmHerMattson9.Hide
    frmHerMattson10.Show
End Sub

Private Sub Option2_Click()
    RightAnswer = False
    Answer = MsgBox("Wrong Answer, the correct Answer was The DaVinci Code", , "Wrong")
    frmHerMattson9.Hide
    frmHerMattson10.Show
End Sub

Private Sub Option3_Click()
    RightAnswer = True
    Counter = Counter + 1
    Answer = MsgBox("Congratulations! You got the answer right!", , "Right Answer")
    frmHerMattson9.Hide
    frmHerMattson10.Show
End Sub

