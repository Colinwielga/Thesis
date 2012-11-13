VERSION 5.00
Begin VB.Form Form17 
   BackColor       =   &H00FF0000&
   Caption         =   "Form17"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form17"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit Game"
      Height          =   1095
      Left            =   7320
      TabIndex        =   2
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdResults 
      Caption         =   "Show my Results"
      Height          =   1095
      Left            =   840
      TabIndex        =   1
      Top             =   5400
      Width           =   6255
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FF0000&
      Height          =   4695
      Left            =   840
      ScaleHeight     =   4635
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   480
      Width           =   8175
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQuit_Click()
'Ends the game'
End
End Sub

Private Sub cmdResults_Click()
'Shows the Results'
picResults.Print "Treasure: "; Treasure;
picResults.Print "Level: "; Level;
picResults.Print "Congratulations, play again anytime.'"
End Sub

