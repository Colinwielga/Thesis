VERSION 5.00
Begin VB.Form frmStairs 
   BackColor       =   &H80000007&
   Caption         =   "Stairs"
   ClientHeight    =   9000
   ClientLeft      =   2370
   ClientTop       =   1935
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   7185
   Begin VB.PictureBox picTure 
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   960
      Picture         =   "frmStairs.frx":0000
      ScaleHeight     =   6615
      ScaleWidth      =   5295
      TabIndex        =   2
      Top             =   120
      Width           =   5295
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Escape"
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   8160
      Width           =   1335
   End
   Begin VB.PictureBox picStairstxt 
      Height          =   975
      Left            =   240
      ScaleHeight     =   915
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   7080
      Width           =   6735
   End
End
Attribute VB_Name = "frmStairs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNext_Click()
frmStairs.Hide
frmOutside.Show
End Sub

Private Sub Form_activate()
picStairstxt.Print "The door behind you shuts and locks. The place has completely locked down."
picStairstxt.Print "The only way out is up."
End Sub

