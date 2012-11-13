VERSION 5.00
Begin VB.Form mansonwins 
   BackColor       =   &H00000080&
   Caption         =   "Nothin but love for ya..."
   ClientHeight    =   5190
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6840
   LinkTopic       =   "Form3"
   ScaleHeight     =   5190
   ScaleWidth      =   6840
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton QuitMansonW 
      BackColor       =   &H000040C0&
      Caption         =   "He's still creepy to me... I want my mommy."
      Height          =   975
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   2490
      Left            =   360
      Picture         =   "mansonwins.frx":0000
      Top             =   240
      Width           =   6120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   $"mansonwins.frx":3732
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   2880
      Width           =   6135
   End
End
Attribute VB_Name = "mansonwins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub QuitMansonW_Click()
End
End Sub
