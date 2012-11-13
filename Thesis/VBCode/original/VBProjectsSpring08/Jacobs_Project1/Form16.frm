VERSION 5.00
Begin VB.Form Form16 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form16"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form16"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWin 
      Caption         =   "Take his hand."
      Height          =   1095
      Left            =   600
      TabIndex        =   3
      Top             =   3840
      Width           =   2535
   End
   Begin VB.CommandButton cmdKill 
      Caption         =   "Attack the King"
      Height          =   1215
      Left            =   600
      TabIndex        =   2
      Top             =   2400
      Width           =   2535
   End
   Begin VB.PictureBox picKing 
      BackColor       =   &H0000FFFF&
      Height          =   4815
      Left            =   6000
      ScaleHeight     =   4755
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   $"Form16.frx":0000
      Height          =   1815
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdKill_Click()
'One more quick game over possiblity'
MsgBox "As you draw your weapon, one of the kings guard stabs you through the throat."
MsgBox "You die. Why attack the king when you could win...idiot. GAME OVER!"
End
End Sub

Private Sub cmdWin_Click()
'Moves to the results screen'
MsgBox "People cheer as you shake the kings hand. You win."
Form16.Hide
Form17.Show
End Sub

Private Sub Form_Load()
'loads the picture of the king'
picKing.Picture = LoadPicture("King.jpg")
End Sub
