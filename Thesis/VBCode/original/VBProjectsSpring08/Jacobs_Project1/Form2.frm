VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00808080&
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDungeon 
      Height          =   3255
      Left            =   360
      ScaleHeight     =   3195
      ScaleWidth      =   5955
      TabIndex        =   4
      Top             =   1200
      Width           =   6015
   End
   Begin VB.CommandButton cmdLeave 
      Caption         =   "Leave Dungeon"
      Height          =   1215
      Left            =   2760
      TabIndex        =   3
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdEast 
      Caption         =   "East"
      Height          =   1335
      Left            =   4560
      TabIndex        =   2
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdWest 
      Caption         =   "West"
      Height          =   1335
      Left            =   960
      TabIndex        =   1
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdNorth 
      Caption         =   "North"
      Height          =   1215
      Left            =   2760
      TabIndex        =   0
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   $"Form2.frx":0000
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PicDescription_Click()
End Sub

Private Sub cmdEast_Click()
'Used to Show movement'
Form2.Hide
Form4.Show
End Sub

Private Sub cmdLeave_Click()
'Essentially there for looks, and because I liked the comment'
MsgBox ("You can't quit already!")
End Sub

Private Sub cmdNorth_Click()
'used for Movement'
Form2.Hide
Form9.Show
End Sub

Private Sub cmdWest_Click()
'Used for Movement
Form2.Hide
Form3.Show
End Sub

Private Sub Form_Load()
'Show a picture of the dungeon'
picDungeon.Picture = LoadPicture("Dungeon3d.jpg")
End Sub
