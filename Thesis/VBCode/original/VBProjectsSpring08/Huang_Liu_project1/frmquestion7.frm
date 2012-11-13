VERSION 5.00
Begin VB.Form frmquestion7 
   Caption         =   "Question7"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form2"
   Picture         =   "frmquestion7.frx":0000
   ScaleHeight     =   3000
   ScaleMode       =   0  'User
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd4 
      Caption         =   "D. Hello Kitty"
      Height          =   1065
      Left            =   6000
      TabIndex        =   3
      Top             =   1800
      Width           =   1500
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "C. Keroppi"
      Height          =   1065
      Left            =   4080
      TabIndex        =   2
      Top             =   1800
      Width           =   1500
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "B. Little Twin Stars"
      Height          =   1065
      Left            =   2160
      TabIndex        =   1
      Top             =   1800
      Width           =   1500
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "A. Purin"
      Height          =   1065
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Label lblq7 
      Caption         =   "Kiki And Lala."
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   7215
   End
End
Attribute VB_Name = "frmquestion7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()

frmquestion7.Visible = False
frmquestion8.Visible = True
End Sub

Private Sub cmd2_Click()
Score = Score + 1
frmquestion7.Visible = False
frmquestion8.Visible = True

End Sub

Private Sub cmd3_Click()

frmquestion7.Visible = False
frmquestion8.Visible = True

End Sub

Private Sub cmd4_Click()

frmquestion7.Visible = False
frmquestion8.Visible = True

End Sub
