VERSION 5.00
Begin VB.Form frmWinnings 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Winnings"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   Picture         =   "frmWinnings.frx":0000
   ScaleHeight     =   3045
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picWinnings 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   960
      ScaleHeight     =   555
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmWinnings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Shows money won
'Ends the program

Private Sub cmdEnd_Click()
End
End Sub

