VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00004000&
   Caption         =   "Form5"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H00004000&
   LinkTopic       =   "Form5"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDungeon 
      Height          =   3135
      Left            =   4320
      ScaleHeight     =   3075
      ScaleWidth      =   3315
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.CommandButton cmdWest 
      Caption         =   "West"
      Height          =   1335
      Left            =   720
      TabIndex        =   1
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton CmdNorth 
      Caption         =   "North"
      Height          =   1335
      Left            =   2040
      TabIndex        =   0
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00004000&
      Caption         =   "You find yourself at a corner in the dungeon. You can go North, or turn back the way you came."
      ForeColor       =   &H00C0C000&
      Height          =   975
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNorth_Click()
'Dungeon Movement'
Form5.Hide
Form6.Show
End Sub

Private Sub cmdWest_Click()
'Dungeon Movement'
Form5.Hide
Form2.Show
End Sub

Private Sub Form_Load()
'Shows the picture'
picDungeon.Picture = LoadPicture("Dungeon9d.jpg")
End Sub
