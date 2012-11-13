VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00C000C0&
   Caption         =   "Form8"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H00C000C0&
   LinkTopic       =   "Form8"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTrap 
      Caption         =   "Lever"
      Height          =   975
      Left            =   960
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdEast 
      Caption         =   "East"
      Height          =   1095
      Left            =   3000
      TabIndex        =   2
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdWest 
      Caption         =   "West"
      Height          =   1095
      Left            =   840
      TabIndex        =   1
      Top             =   3840
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C000C0&
      Height          =   3015
      Left            =   5040
      ScaleHeight     =   2955
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label lblDescribe 
      BackColor       =   &H00C000C0&
      Caption         =   "You are going down a hall when you see a lever. You can pull the lever, go West, or East."
      Height          =   855
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEast_Click()
'movement'
Form8.Hide
Form7.Show
End Sub

Private Sub cmdTrap_Click()
'I decided to put a trap in the game, a little obvious, but it ends the game nevertheless'
MsgBox ("The wall opens up and fire comes out. ITS A TRAP! You die, Game over.")
End
End Sub

Private Sub cmdWest_Click()
'movement'
Form8.Hide
Form10.Show
End Sub

Private Sub Form_Load()
'Picture'
picDungeon.Picture = LoadPicture("Dungeon6d.jpg")
End Sub
