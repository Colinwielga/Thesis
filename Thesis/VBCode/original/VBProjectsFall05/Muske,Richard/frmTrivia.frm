VERSION 5.00
Begin VB.Form frmTrivia 
   BackColor       =   &H80000009&
   Caption         =   "Question 1"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Question"
      Height          =   615
      Left            =   3120
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H8000000E&
      Height          =   2655
      Left            =   3360
      ScaleHeight     =   2595
      ScaleWidth      =   4035
      TabIndex        =   6
      Top             =   1080
      Width           =   4095
   End
   Begin VB.CommandButton cmdD 
      Caption         =   "D. Roger Staubach    "
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "C. Joe Montana          "
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdB 
      Caption         =   "B. John Elway            "
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "A. Terry Bradshaw      "
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton cmdB2MM2 
      Caption         =   "Back To Main Menu"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label lblQ1 
      BackColor       =   &H8000000E&
      Caption         =   "1. Who has the most touchdown passes in Super Bowl history?"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmTrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: SportsWinners (Rich Muske's SportsWinners.vbp)
'From Name:  frmTrivia (frmTrivia.frm)
'Author:  Rich Muske
'Date Written: 10/30
'Purpose:  To get someone to answer my first question by clicking the right command button.





Private Sub cmdA_Click()
picOutput.Cls
picOutput.Print "Terry Bradshaw is Incorrect"
End Sub

Private Sub cmdB_Click()
picOutput.Cls
picOutput.Print "John Elway is Incorrect"
End Sub

Private Sub cmdB2MM2_Click()
    frmTrivia.Hide
    frmMainMenu.Show
End Sub

Private Sub cmdC_Click()
picOutput.Cls
picOutput.Print "Joe Montana is correct."
picOutput.Print "He threw 11 touchdown passes in 4 games."
End Sub

Private Sub cmdD_Click()
picOutput.Cls
picOutput.Print "Roger Staubach is Incorrect"
End Sub

Private Sub cmdNext_Click()
frmTrivia.Hide
frmQuestion2.Show
End Sub

