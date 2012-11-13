VERSION 5.00
Begin VB.Form frmWallet 
   Caption         =   "Wallet"
   ClientHeight    =   4125
   ClientLeft      =   4410
   ClientTop       =   2460
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   6975
   Begin VB.CommandButton cmdLobby 
      BackColor       =   &H000000FF&
      Caption         =   "Go back to Lobby"
      Height          =   735
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00000080&
      Height          =   615
      Left            =   1680
      ScaleHeight     =   555
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   2040
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   0
      Picture         =   "frmWallet.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   6915
      TabIndex        =   2
      Top             =   0
      Width           =   6975
      Begin VB.CommandButton cmdMoney 
         BackColor       =   &H0000C000&
         Caption         =   "Count your money"
         Height          =   735
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3120
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmWallet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form gives the user a place to check their balance
Private Sub cmdLobby_Click()
    'Go back to Lobby
    frmWallet.Hide
    frmLobby.Show
End Sub

Private Sub cmdMoney_Click()
    'Displays current balance in picturebox
    picResults.Cls
    picResults.Print "Your wallet holds " & FormatCurrency(balanceglobal) & "."
End Sub


