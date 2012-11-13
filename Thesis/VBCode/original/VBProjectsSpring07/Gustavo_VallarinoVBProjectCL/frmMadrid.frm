VERSION 5.00
Begin VB.Form frmMadrid 
   BackColor       =   &H8000000D&
   Caption         =   "The ETERNAL CHAMPION"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMain 
      Caption         =   "Back to Main Menu"
      Height          =   855
      Left            =   7800
      TabIndex        =   5
      Top             =   7440
      Width           =   2775
   End
   Begin VB.CommandButton cmdChampions 
      Caption         =   "Back to Champions"
      Height          =   975
      Left            =   720
      TabIndex        =   4
      Top             =   4320
      Width           =   2655
   End
   Begin VB.PictureBox picresults 
      Height          =   3735
      Left            =   4200
      Picture         =   "frmMadrid.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   4200
      Width           =   2775
   End
   Begin VB.PictureBox Picture3 
      Height          =   6735
      Left            =   5640
      Picture         =   "frmMadrid.frx":54B3
      ScaleHeight     =   6675
      ScaleWidth      =   5115
      TabIndex        =   3
      Top             =   0
      Width           =   5175
   End
   Begin VB.PictureBox Picture2 
      Height          =   2775
      Left            =   0
      Picture         =   "frmMadrid.frx":EB82
      ScaleHeight     =   2715
      ScaleWidth      =   4515
      TabIndex        =   2
      Top             =   6000
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000D&
      Height          =   3375
      Left            =   240
      Picture         =   "frmMadrid.frx":13A0A
      ScaleHeight     =   3315
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmMadrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The purpose of this form is to show the team that has been champion most times, and my favorite team in the tournament.
'they are only pictures added in picture boxes and used certain picture edditing features.

Private Sub cmdChampions_Click()
frmWinner.Show
frmMadrid.Hide
End Sub

Private Sub cmdMain_Click()
frmChampions.Show
frmMadrid.Hide
End Sub
