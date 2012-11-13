VERSION 5.00
Begin VB.Form frmquestion6 
   Caption         =   "Question6"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form2"
   Picture         =   "frmquestion6.frx":0000
   ScaleHeight     =   3000
   ScaleMode       =   0  'User
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd2 
      Caption         =   "B. Monkichi"
      Height          =   1065
      Left            =   2280
      TabIndex        =   3
      Top             =   1800
      Width           =   1500
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "C. Chibimaru"
      Height          =   1065
      Left            =   4200
      TabIndex        =   2
      Top             =   1800
      Width           =   1500
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "D. Pochacco"
      Height          =   1065
      Left            =   6120
      TabIndex        =   1
      Top             =   1800
      Width           =   1500
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "A. Pandapple"
      Height          =   1065
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Label lblq6 
      Caption         =   "who is a boy panda who absolutely LOVES apples?"
      Height          =   1095
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   7215
   End
End
Attribute VB_Name = "frmquestion6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
Score = Score + 1
frmquestion6.Visible = False
frmquestion7.Visible = True
End Sub

Private Sub cmd2_Click()

frmquestion6.Visible = False
frmquestion7.Visible = True

End Sub

Private Sub cmd3_Click()

frmquestion6.Visible = False
frmquestion7.Visible = True

End Sub

Private Sub cmd4_Click()

frmquestion6.Visible = False
frmquestion7.Visible = True

End Sub

