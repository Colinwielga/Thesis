VERSION 5.00
Begin VB.Form George_W_Bush 
   BackColor       =   &H000000C0&
   Caption         =   "The leader of the free world..."
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4725
   LinkTopic       =   "Form8"
   Picture         =   "bushwins.frx":0000
   ScaleHeight     =   6195
   ScaleWidth      =   4725
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton QuitBushW 
      BackColor       =   &H00FF8080&
      Caption         =   "Quit"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Sorry, you've lost to our highly intellectual fearless leader.  Don't worry, it could happen to anyone...  Thanks for playing!"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   4455
   End
End
Attribute VB_Name = "George_W_Bush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub QuitBushW_Click()
End
End Sub
