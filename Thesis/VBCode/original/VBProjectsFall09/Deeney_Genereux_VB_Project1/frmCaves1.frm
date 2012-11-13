VERSION 5.00
Begin VB.Form frmCaves1 
   BackColor       =   &H00000000&
   Caption         =   "Pick a Cave!"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   975
   End
   Begin VB.Image img6 
      Height          =   2085
      Left            =   360
      Picture         =   "frmCaves1.frx":0000
      Top             =   240
      Width           =   2220
   End
   Begin VB.Image img7 
      Height          =   2085
      Left            =   2520
      Picture         =   "frmCaves1.frx":1286
      Top             =   240
      Width           =   2220
   End
   Begin VB.Image img8 
      Height          =   2085
      Left            =   4680
      Picture         =   "frmCaves1.frx":255C
      Top             =   240
      Width           =   2220
   End
   Begin VB.Image img9 
      Height          =   2085
      Left            =   6840
      Picture         =   "frmCaves1.frx":3811
      Top             =   240
      Width           =   2220
   End
End
Attribute VB_Name = "frmCaves1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'Caves1
'The user has to pick from 4 caves by clicking
'on the one they want

Private Sub CmdQuit_Click()
    End
End Sub

Private Sub img6_Click()
frmCaves1.Hide
frmCave1.Show
End Sub

Private Sub img7_Click()
frmCaves1.Hide
frmCave2.Show
End Sub

Private Sub img8_Click()
frmCaves1.Hide
frmCave3.Show
End Sub

Private Sub img9_Click()
frmCaves1.Hide
frmCave4.Show
End Sub
