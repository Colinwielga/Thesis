VERSION 5.00
Begin VB.Form frmquestion1 
   Caption         =   "Question1"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form2"
   Picture         =   "frmquestion1.frx":0000
   ScaleHeight     =   3000
   ScaleMode       =   0  'User
   ScaleWidth      =   8187.134
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd2 
      Caption         =   "B. CharmmyKitty"
      Height          =   855
      Left            =   2160
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "C. Chococat"
      Height          =   855
      Left            =   4080
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "D. Chi Chai Monchan"
      Height          =   855
      Left            =   6000
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "A. Badtz-Maru"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblq1 
      Caption         =   "Which character is marketed to both males and females?"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7215
   End
End
Attribute VB_Name = "frmquestion1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd1_Click()
Score = Score + 1
frmquestion1.Visible = False
frmquestion2.Visible = True
End Sub

Private Sub cmd2_Click()
frmquestion1.Visible = False
frmquestion2.Visible = True

End Sub

Private Sub cmd3_Click()
frmquestion1.Visible = False
frmquestion2.Visible = True

End Sub

Private Sub cmd4_Click()
frmquestion1.Visible = False
frmquestion2.Visible = True

End Sub
