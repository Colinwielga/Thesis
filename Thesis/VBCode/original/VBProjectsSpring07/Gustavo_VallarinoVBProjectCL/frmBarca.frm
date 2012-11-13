VERSION 5.00
Begin VB.Form frmBarca 
   BackColor       =   &H8000000D&
   Caption         =   "Champions 2006"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdwin 
      Caption         =   "Back to Champions"
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   5280
      Width           =   2775
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Back to Main Menu"
      Height          =   735
      Left            =   5040
      TabIndex        =   1
      Top             =   5280
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   5055
      Left            =   0
      Picture         =   "frmBarca.frx":0000
      ScaleHeight     =   4995
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmBarca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'used this form to ilustrate this tears champion and print some picutres of it

Private Sub cmdMenu_Click()
frmChampions.Show
frmBarca.Hide
End Sub

Private Sub cmdwin_Click()
frmWinner.Show
frmBarca.Hide
End Sub
