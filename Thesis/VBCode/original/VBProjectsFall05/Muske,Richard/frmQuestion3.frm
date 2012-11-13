VERSION 5.00
Begin VB.Form frmQuestion3 
   BackColor       =   &H8000000E&
   Caption         =   "Question 3"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Question"
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back To Main Menu"
      Height          =   735
      Left            =   600
      TabIndex        =   6
      Top             =   4200
      Width           =   1695
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H8000000E&
      Height          =   2415
      Left            =   2880
      ScaleHeight     =   2355
      ScaleWidth      =   3915
      TabIndex        =   5
      Top             =   1200
      Width           =   3975
   End
   Begin VB.CommandButton cmdD 
      Caption         =   "D.  Larry Bird       "
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "C.  Wilt Chamberlin"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdB 
      Caption         =   "B.  Michael Jordan"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "A.  Elgin Baylor     "
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblQ3 
      BackColor       =   &H8000000E&
      Caption         =   "Who Scored the most points in an NBA finals game?"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmQuestion3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: SportsWinners (Rich Muske's SportsWinners.vbp)
'Form Name:  frmQuestion3 (frmQuestion3.frm)
'Author: Rich Muske
'Date Written 10/30
'Purpose:  The purpose of this form is to get someone to click on the right answer to my third question.



Private Sub cmdA_Click()
picOutput.Cls
picOutput.Print "Elgin Baylor is correct."
picOutput.Print "He scored 61 points against Boston in 1962"
End Sub

Private Sub cmdB_Click()
picOutput.Cls
picOutput.Print "Michael Jordan is incorrect."
End Sub

Private Sub cmdBack_Click()
frmQuestion3.Hide
frmMainMenu.Show
End Sub

Private Sub cmdC_Click()
picOutput.Cls
picOutput.Print "Wilt Chamberlin is incorrect."
End Sub

Private Sub cmdD_Click()
picOutput.Cls
picOutput.Print "Larry Bird is incorrect."
End Sub

Private Sub cmdNext_Click()
frmQuestion3.Hide
frmQuestion4.Show
End Sub

