VERSION 5.00
Begin VB.Form frmMoneyRoom 
   Caption         =   "Casino Storage Room"
   ClientHeight    =   5085
   ClientLeft      =   3975
   ClientTop       =   2250
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   Picture         =   "frmMoneyRoom.frx":0000
   ScaleHeight     =   5085
   ScaleWidth      =   7065
   Begin VB.PictureBox Picture1 
      Height          =   5175
      Left            =   0
      Picture         =   "frmMoneyRoom.frx":CC16
      ScaleHeight     =   5115
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton cmdLeave 
         BackColor       =   &H008080FF&
         Caption         =   "Take a mental picture and leave"
         Height          =   975
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3960
         Width           =   2295
      End
      Begin VB.CommandButton cmdJump 
         BackColor       =   &H0080FFFF&
         Caption         =   "Jump in the money"
         Height          =   975
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3960
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmMoneyRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form is where the Casino keeps its money
'You need to obtain a key to enable the button

Private Sub cmdJump_Click()
    'When clicked, the user gets caught by the cops and returned to the lobby
    MsgBox "There he is! Get 'em!", , "Security caughting illegal business"
    MsgBox "After a wild chase through the money room, you were caught and sentenced to 50 years in Casino Jail to restack bills.", , "Busted"
    MsgBox "But after further thought you gave the key to the cops in exchange for your freedom and were released.", , "Redemption"
    frmMoneyRoom.Hide
    frmLobby.Show
    frmLobby.cmdEntrance.Enabled = False

End Sub

Private Sub cmdLeave_Click()
    'Go back to Lobby
    frmLobby.Show
    frmMoneyRoom.Hide
    
End Sub

