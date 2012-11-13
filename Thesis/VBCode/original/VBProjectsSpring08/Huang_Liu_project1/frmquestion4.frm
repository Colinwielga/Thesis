VERSION 5.00
Begin VB.Form frmquestion4 
   Caption         =   "Question4"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form2"
   Picture         =   "frmquestion4.frx":0000
   ScaleHeight     =   3000
   ScaleMode       =   0  'User
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd2 
      Caption         =   "B. Pandapple"
      Height          =   1065
      Left            =   2160
      TabIndex        =   3
      Top             =   1800
      Width           =   1500
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "C. My Melody"
      Height          =   1065
      Left            =   4200
      TabIndex        =   2
      Top             =   1800
      Width           =   1500
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "D. Minna No Tabo"
      Height          =   1065
      Left            =   6000
      TabIndex        =   1
      Top             =   1800
      Width           =   1500
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "A. Pekkle"
      Height          =   1065
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Label lblq4 
      Caption         =   "Whose straightforward and can-do approach to life wins him many friends?"
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   7215
   End
End
Attribute VB_Name = "frmquestion4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()

frmquestion4.Visible = False
frmquestion5.Visible = True
End Sub

Private Sub cmd2_Click()
frmquestion4.Visible = False
frmquestion5.Visible = True

End Sub

Private Sub cmd3_Click()

frmquestion4.Visible = False
frmquestion5.Visible = True

End Sub

Private Sub cmd4_Click()
Score = Score + 1
frmquestion4.Visible = False
frmquestion5.Visible = True

End Sub

