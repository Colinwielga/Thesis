VERSION 5.00
Begin VB.Form frmGame 
   Caption         =   "Game"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTrivia 
      Caption         =   "Trivia"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1515
      ScaleWidth      =   7995
      TabIndex        =   1
      Top             =   3240
      Width           =   8055
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   8835
      Left            =   -3840
      Picture         =   "frmGame.frx":0000
      Top             =   -2160
      Width           =   13710
   End
   Begin VB.Image Image1 
      Height          =   4935
      Left            =   0
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdHome_Click()
    frmGame.Hide
    frmTitle.Show
End Sub

