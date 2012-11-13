VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00808080&
   Caption         =   "Form7"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form7"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWest 
      Caption         =   "West"
      Height          =   1215
      Left            =   360
      TabIndex        =   3
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdSouth 
      Caption         =   "South"
      Height          =   1215
      Left            =   1920
      TabIndex        =   2
      Top             =   6120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   5175
      Left            =   5040
      ScaleHeight     =   5115
      ScaleWidth      =   5115
      TabIndex        =   1
      Top             =   480
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "You come to another corner in the dungeon. Your exits are West and South."
      Height          =   1575
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSouth_Click()
'Movement
Form7.Hide
Form5.Show
End Sub

Private Sub cmdWest_Click()
'movement'
Form7.Hide
Form8.Show
End Sub

Private Sub Form_Load()
'Pictures'
picDungeon.Picture = LoadPicture("Dungeon4d.jpg")
End Sub
