VERSION 5.00
Begin VB.Form beatbush 
   BackColor       =   &H000000FF&
   Caption         =   "He's so endearingly pathetic..."
   ClientHeight    =   6570
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   2280
      TabIndex        =   1
      Top             =   8520
      Width           =   2295
   End
   Begin VB.CommandButton GoAgain1 
      Caption         =   "Uh oh!  You made Georgie lose!  "
      Height          =   1575
      Left            =   1440
      TabIndex        =   0
      Top             =   4560
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Programmed by Megan Kelly"
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   4290
      Left            =   720
      Picture         =   "beatbush.frx":0000
      Top             =   120
      Width           =   6150
   End
End
Attribute VB_Name = "beatbush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Quit_Click()
End
End Sub
