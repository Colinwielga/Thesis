VERSION 5.00
Begin VB.Form frmBonus 
   Caption         =   "Form1"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCanvas 
      Height          =   3015
      Left            =   360
      ScaleHeight     =   2955
      ScaleWidth      =   4515
      TabIndex        =   1
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label lblCongrats 
      Caption         =   "Congrats on being done!"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   3735
      Left            =   240
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmBonus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    picCanvas.AutoRedraw = True
    picCanvas.DrawWidth = 2
    picCanvas.ForeColor = vbBlack
    picCanvas.BackColor = vbWhite
End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picCanvas.Line (X, Y)-(X, Y)
    End If
End Sub
Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picCanvas.Line -(X, Y)
    End If
End Sub

