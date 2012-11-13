VERSION 5.00
Begin VB.Form frmquestion3 
   Caption         =   "Question3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form2"
   Picture         =   "frmquestion3.frx":0000
   ScaleHeight     =   3000
   ScaleMode       =   0  'User
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd2 
      Caption         =   "B. Little Twin Stars"
      Height          =   1065
      Left            =   2160
      TabIndex        =   4
      Top             =   1680
      Width           =   1500
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "C. Chibimaru"
      Height          =   1065
      Left            =   4080
      TabIndex        =   3
      Top             =   1680
      Width           =   1500
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "D. Chi Chai Monchan"
      Height          =   1065
      Left            =   6000
      TabIndex        =   2
      Top             =   1680
      Width           =   1500
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "A. Minna No Tabo"
      Height          =   1065
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   1500
   End
   Begin VB.Label lblq3 
      Caption         =   "Who lives in a house with a red roof with his favorite toys-his stuffed animals?"
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "frmquestion3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()

frmquestion3.Visible = False
frmquestion4.Visible = True
End Sub

Private Sub cmd2_Click()
frmquestion3.Visible = False
frmquestion4.Visible = True

End Sub

Private Sub cmd3_Click()
Score = Score + 1
frmquestion3.Visible = False
frmquestion4.Visible = True

End Sub

Private Sub cmd4_Click()
frmquestion3.Visible = False
frmquestion4.Visible = True

End Sub

