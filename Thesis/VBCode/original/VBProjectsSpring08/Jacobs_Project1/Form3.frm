VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00808080&
   Caption         =   "Form3"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form3"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEast 
      Caption         =   "East"
      Height          =   1335
      Left            =   2880
      TabIndex        =   1
      Top             =   5640
      Width           =   2295
   End
   Begin VB.PictureBox picdungeon 
      BackColor       =   &H00808080&
      Height          =   3015
      Left            =   960
      ScaleHeight     =   2955
      ScaleWidth      =   5475
      TabIndex        =   0
      Top             =   2280
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "You go west and eventually come to a dead end. You can only travel East."
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   3495
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEast_Click()
'To show movement back'
Form3.Hide
Form2.Show
End Sub

Private Sub Form_Load()
'Dead end picture'
picDungeon.Picture = LoadPicture("Deadend.JPG")
End Sub
