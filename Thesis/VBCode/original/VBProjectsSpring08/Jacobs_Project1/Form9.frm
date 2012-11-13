VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00808080&
   Caption         =   "Form9"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form9"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSouth 
      Caption         =   "South"
      Height          =   1095
      Left            =   1200
      TabIndex        =   3
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton cmdNorth 
      Caption         =   "North"
      Height          =   1095
      Left            =   1200
      TabIndex        =   2
      Top             =   3360
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   2775
      Left            =   4680
      ScaleHeight     =   2715
      ScaleWidth      =   3915
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "You Head north to another area of the dungeon. It appears nondescipt. You can  continue North, or head South."
      Height          =   1815
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNorth_Click()
'movement'
Form9.Hide
Form10.Show
End Sub

Private Sub cmdSouth_Click()
'movement'
Form9.Hide
Form2.Show
End Sub

Private Sub Form_Load()
'show picture'
picDungeon.Picture = LoadPicture("Dungeon5d.jpg")
End Sub
