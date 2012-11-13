VERSION 5.00
Begin VB.Form frmQuestion4 
   BackColor       =   &H8000000E&
   Caption         =   "Question 4"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOutput 
      BackColor       =   &H8000000E&
      Height          =   2535
      Left            =   2880
      ScaleHeight     =   2475
      ScaleWidth      =   4275
      TabIndex        =   6
      Top             =   960
      Width           =   4335
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Main Menu"
      Height          =   735
      Left            =   720
      TabIndex        =   5
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdD 
      Caption         =   "D.  Ray Bourque"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "C.  Patrick Roy   "
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdB 
      Caption         =   "B.  Wayne Gretzky"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "A.  Sergei Federov"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblQ4 
      BackColor       =   &H8000000E&
      Caption         =   "4.  Who has played in the most playoff games in NHL history?"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmQuestion4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: SportsWinners (Rich Muske's SportsWinners.vbp)
'Form Name: frmQuestion4 (frmQuestion4.frm)
'Author: Rich Muske
'Date Written: 10/30
'Purpose: The purpose of this form is to get someone to answer my third question by clicking on the right command button.




Private Sub cmdA_Click()
picOutput.Cls
picOutput.Print "Sergei Federov is incorrect."
End Sub

Private Sub cmdB_Click()
picOutput.Cls
picOutput.Print "Wayne Gretzky is incorrect."
End Sub

Private Sub cmdBack_Click()
frmQuestion4.Hide
frmMainMenu.Show
End Sub

Private Sub cmdC_Click()
picOutput.Cls
picOutput.Print "Patrick Roy is correct."
picOutput.Print "He has played in 247 playoff games"
End Sub

Private Sub cmdD_Click()
picOutput.Cls
picOutput.Print "Ray Bourque is incorrect."
End Sub

