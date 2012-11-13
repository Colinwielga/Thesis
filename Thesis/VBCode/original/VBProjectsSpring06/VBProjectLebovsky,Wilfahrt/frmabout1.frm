VERSION 5.00
Begin VB.Form frmabout 
   BackColor       =   &H00000000&
   Caption         =   "frmabout"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00808080&
      Caption         =   "exit"
      Height          =   615
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton cmdmain 
      BackColor       =   &H00808080&
      Caption         =   "Return to main"
      Height          =   615
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox picabout 
      Height          =   6975
      Left            =   0
      Picture         =   "frmabout.frx":0000
      ScaleHeight     =   6915
      ScaleWidth      =   10755
      TabIndex        =   1
      Top             =   600
      Width           =   10815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Click anywhere to learn about our project"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdmain_Click()
frmabout.Hide
frmmain.Show
End Sub

Private Sub Picabout_Click()
MsgBox "This program was intended solely for the purpose of entertainment for those who choose to use it.  It was created by Andy Lebovsky and Clay Wilfahrt for a computer science class."
End Sub


