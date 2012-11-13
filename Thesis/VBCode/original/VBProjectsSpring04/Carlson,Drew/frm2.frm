VERSION 5.00
Begin VB.Form frm2 
   Caption         =   "Form2"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11115
   LinkTopic       =   "Form2"
   ScaleHeight     =   6765
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   8640
      TabIndex        =   2
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Main Stats Program"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   2055
   End
   Begin VB.PictureBox picresults 
      Height          =   5175
      Left            =   0
      Picture         =   "frm2.frx":0000
      ScaleHeight     =   5115
      ScaleWidth      =   10755
      TabIndex        =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Click()

End Sub

Private Sub cmdreturn_Click()
Form1.Show
frm2.Hide

End Sub

Private Sub Quit_Click()
End
End Sub
