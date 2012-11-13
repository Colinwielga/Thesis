VERSION 5.00
Begin VB.Form frmQuestion2 
   BackColor       =   &H8000000E&
   Caption         =   "Question 2"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOutput 
      BackColor       =   &H8000000E&
      Height          =   2655
      Left            =   2760
      ScaleHeight     =   2595
      ScaleWidth      =   3795
      TabIndex        =   7
      Top             =   1080
      Width           =   3855
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Main menu"
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Next Question"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdD 
      Caption         =   "D.  Don Larson"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "C.  Whitey Ford"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdB 
      BackColor       =   &H8000000E&
      Caption         =   "B.  Cy Young  "
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "A.  Nolan Ryan"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblQ2 
      BackColor       =   &H8000000E&
      Caption         =   "2. Who threw the only no-hitter in World Series history?"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmQuestion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: SportsWinners (Rich Muske's SportsWinners.vbp
'Form Name: frmQuestion2 (frmQuestion2.frm)
'Author: Rich Muske
'Date Written: 10/30
'Purpose: The purpose of this form is to ask my second question and get someone to click on the right answer



Private Sub cmdA_Click()
picOutput.Cls
picOutput.Print "Nolan Ryan is incorrect."
End Sub

Private Sub cmdB_Click()
picOutput.Cls
picOutput.Print "Cy Young is incorrect."
End Sub

Private Sub cmdBack_Click()
frmQuestion2.Hide
frmMainMenu.Show
End Sub

Private Sub cmdC_Click()
picOutput.Cls
picOutput.Print "Whitey Ford is incorrect."
End Sub

Private Sub cmdD_Click()
picOutput.Cls
picOutput.Print "Don Larson is correct."
picOutput.Print "He threw a perfect game in the 1956 World Series"
End Sub

Private Sub cmdNext_Click()
frmQuestion2.Hide
frmQuestion3.Show
End Sub

