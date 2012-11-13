VERSION 5.00
Begin VB.Form frmLobby 
   Caption         =   "Mystake Lake Lobby"
   ClientHeight    =   9240
   ClientLeft      =   795
   ClientTop       =   870
   ClientWidth     =   13680
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   13680
   Begin VB.PictureBox Picture1 
      Height          =   11055
      Left            =   -120
      Picture         =   "frmLobby.frx":0000
      ScaleHeight     =   10995
      ScaleWidth      =   15195
      TabIndex        =   0
      ToolTipText     =   "You could win big at Mystake Lake Casino!"
      Top             =   -120
      Width           =   15255
      Begin VB.CommandButton cmdLeaveCasino 
         BackColor       =   &H000000C0&
         Caption         =   "Leave Casino"
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5880
         Width           =   2055
      End
      Begin VB.CommandButton cmdEntrance 
         BackColor       =   &H00800080&
         Caption         =   "Casino Personnel Only"
         Enabled         =   0   'False
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   7080
         Width           =   2055
      End
      Begin VB.CommandButton cmdFood 
         BackColor       =   &H0080FF80&
         Caption         =   "Food Court"
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Working hard or hardly working, why don't you take a break and visit our delicious food court"
         Top             =   4680
         Width           =   2055
      End
      Begin VB.CommandButton cmdStats 
         BackColor       =   &H0080C0FF&
         Caption         =   "Check Today's Stats"
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton cmdWallet 
         BackColor       =   &H0000C000&
         Caption         =   "Check your Wallet"
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton cmdRoulette 
         BackColor       =   &H0080FFFF&
         Caption         =   "Go to Roulette Table"
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to Mystake Lake Casino"
         BeginProperty Font 
            Name            =   "Magneto"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   855
         Left            =   2160
         TabIndex        =   7
         Top             =   240
         Width           =   8175
      End
   End
End
Attribute VB_Name = "frmLobby"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Mystake Lake Casino
'Authors: David Johnson And Jeremy Iverson
'Date: Monday, November 5, 2007

Option Explicit
'This form is a place for the user to choose different actions in the casino
'From this form, the user can go play roulette, check Stats, get food,
'check Wallet (current balance), leave casino, or enter a locked door provided he or she has the key

Private Sub cmdEntrance_Click()
    'Go to Money Room
    frmLobby.Hide
    frmMoneyRoom.Show
    
End Sub

Private Sub cmdFood_Click()
    'Go to Food Menu
    frmLobby.Hide
    frmMenu.Show
End Sub

Private Sub cmdLeaveCasino_Click()
    'Leave Casino, if the user borrowed money from the user, they deal with Loan Sharks
    'Otherwise the program ends
    If clicked = True Then
       frmLobby.Hide
       MsgBox "Uh Oh", , "Shark Attack?"
       frmGangster.Show
    Else
        MsgBox "Thanks for visiting Mystake Lake Casino", , "Goodbye"
        End
    End If
    
End Sub

Private Sub cmdRoulette_Click()
    'Go to Roulette table
    frmLobby.Hide
    frmRoulette.Show
End Sub

Private Sub cmdStats_Click()
    'Go see Stats of other players for the day
    frmLobby.Hide
    frmStats.Show
    
End Sub

Private Sub cmdWallet_Click()
    'Check your current balance
    frmWallet.Show
End Sub

