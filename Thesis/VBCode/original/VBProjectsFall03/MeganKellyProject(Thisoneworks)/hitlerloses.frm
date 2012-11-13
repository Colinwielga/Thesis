VERSION 5.00
Begin VB.Form hitlerloses 
   BackColor       =   &H0082A952&
   Caption         =   "Oh, how the mighty have fallen..."
   ClientHeight    =   7260
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   7560
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton QuitHitlerL 
      BackColor       =   &H00C000C0&
      Caption         =   "Quit"
      Height          =   855
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   3015
   End
   Begin VB.CommandButton Quit 
      Caption         =   "Click here to get away from the scathing image of that haircut..."
      Height          =   1095
      Left            =   2400
      TabIndex        =   1
      Top             =   8760
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   5325
      Left            =   720
      Picture         =   "hitlerloses.frx":0000
      Top             =   120
      Width           =   6195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0082A952&
      Caption         =   "You've just won WWII!  (Well, sort of....) Luckily the Allies got yo' back. "
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   5520
      Width           =   6855
   End
End
Attribute VB_Name = "hitlerloses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub QuitHitlerL_Click()
End
End Sub
