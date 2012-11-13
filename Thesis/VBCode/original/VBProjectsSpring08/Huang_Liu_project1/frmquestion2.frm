VERSION 5.00
Begin VB.Form frmquestion2 
   Caption         =   "Question2"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form2"
   Picture         =   "frmquestion2.frx":0000
   ScaleHeight     =   3000
   ScaleMode       =   0  'User
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd2 
      Caption         =   "B. Deery-lou"
      Height          =   1035
      Left            =   2160
      TabIndex        =   4
      Top             =   1680
      Width           =   1500
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "C. Cinnamoroll"
      Height          =   1035
      Left            =   4080
      TabIndex        =   3
      Top             =   1680
      Width           =   1500
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "D. Keroppi"
      Height          =   1035
      Left            =   6000
      TabIndex        =   2
      Top             =   1680
      Width           =   1500
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "A. Chibimaru"
      Height          =   1035
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   1500
   End
   Begin VB.Label lblq2 
      Caption         =   "Whose tail was plump and curled up like a cinnamon roll?"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7455
   End
End
Attribute VB_Name = "frmquestion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()

frmquestion2.Visible = False
frmquestion3.Visible = True
End Sub

Private Sub cmd2_Click()
frmquestion2.Visible = False
frmquestion3.Visible = True

End Sub

Private Sub cmd3_Click()
Score = Score + 1
frmquestion2.Visible = False
frmquestion3.Visible = True

End Sub

Private Sub cmd4_Click()
frmquestion2.Visible = False
frmquestion3.Visible = True

End Sub
