VERSION 5.00
Begin VB.Form Results 
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   7815
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Quit8 
      Caption         =   "Quit"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton GoAgain 
      Caption         =   "Click here for another round."
      Height          =   1095
      Left            =   1560
      TabIndex        =   1
      Top             =   4440
      Width           =   3855
   End
   Begin VB.PictureBox FinalPicture 
      Height          =   4095
      Left            =   960
      ScaleHeight     =   4035
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "Results"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
