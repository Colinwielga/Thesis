VERSION 5.00
Begin VB.Form Ashcroftwins 
   BackColor       =   &H00FF8080&
   Caption         =   "Can't fight the mojo..."
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   6480
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton QuitAshcroft 
      Caption         =   "Quit"
      Height          =   975
      Left            =   1800
      TabIndex        =   1
      Top             =   6240
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Programmed by Megan Kelly"
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   7440
      Width           =   2655
   End
   Begin VB.Label Ashcroftwins 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   $"Ashcroftwins.frx":0000
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   1320
      TabIndex        =   0
      Top             =   5040
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   4770
      Left            =   1320
      Picture         =   "Ashcroftwins.frx":00DC
      Top             =   240
      Width           =   3450
   End
End
Attribute VB_Name = "Ashcroftwins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub QuitAshcroft_Click()
End
End Sub
