VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   Picture         =   "frmMainMenu.frx":0000
   ScaleHeight     =   4785
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTrivia 
      Caption         =   "Go To Trivia"
      Height          =   1095
      Left            =   2640
      TabIndex        =   2
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   975
      Left            =   4680
      TabIndex        =   1
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton cmdWinners 
      Caption         =   "Find The Winners Of The Four Major Sports"
      Height          =   1095
      Left            =   6960
      TabIndex        =   0
      Top             =   5880
      Width           =   1935
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: SportsWinners (Rich Muske's SportsWinners.vbp)
'Form Name: frmMainMenu (frmMainMenu.frm)
'Author: Rich Muske
'Date Written: 10/28
'Purpose: To be able to search and sort the different winners of the main 4 sporting events. Baseball, basketball, football, and hockey.



Option Explicit
Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdTrivia_Click()
    frmMainMenu.Hide
    frmTrivia.Show
End Sub

Private Sub cmdWinners_Click()
    frmMainMenu.Hide
    frmWinners.Show
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Image1_Click()

End Sub

